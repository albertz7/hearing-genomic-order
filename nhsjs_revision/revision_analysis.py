from __future__ import annotations

import csv
import hashlib
import json
import math
import pathlib
import random
import re
import zlib
from collections import Counter, defaultdict

import numpy as np


SOURCE_ROOT = pathlib.Path(__file__).resolve().parents[1]
OUT_ROOT = pathlib.Path(__file__).resolve().parent
RESULTS = OUT_ROOT / "results"
REPORT = SOURCE_ROOT / "data" / "ncbi_xylella_dataset_report.json"
FASTA_DIR = SOURCE_ROOT / "data" / "fastas"

WINDOW_SIZE = 3_000
WINDOWS_PER_GENOME = 60
ANOMALIES_PER_GENOME = 4
BACKGROUND_RATE = 0.001
SEED = 20260712
PERMUTATIONS = 1_000
RANDOM_MAPS = 10
BASES = "ACGT"
TRIPLETS = [a + b + c for a in BASES for b in BASES for c in BASES]
TRIPLET_INDEX = {triplet: i for i, triplet in enumerate(TRIPLETS)}
LEVEL_RANK = {"Complete Genome": 0, "Chromosome": 1, "Scaffold": 2}
METHODS = ["composition", "triplet_distribution", "pitch_order", "triplet_transition", "sixmer", "compression"]


def attributes(report: dict) -> dict[str, str]:
    values = {}
    for item in report.get("assembly_info", {}).get("biosample", {}).get("attributes", []):
        if item.get("name") and item.get("value") and item.get("value") != "missing":
            values[item["name"]] = item["value"]
    return values


def strain_name(report: dict) -> str:
    return report.get("organism", {}).get("infraspecific_names", {}).get("strain") or attributes(report).get("strain") or ""


def normalized_strain(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", value.lower())


def host_source(report: dict) -> str:
    attrs = attributes(report)
    return attrs.get("host") or attrs.get("isolation_source") or ""


def host_group(host: str) -> str | None:
    value = host.lower()
    if "vitis" in value or "grape" in value:
        return "grapevine"
    if "prunus dulcis" in value or "almond" in value:
        return "almond"
    if "olea" in value or "olive" in value:
        return "olive"
    if "coffea" in value or "coffee" in value:
        return "coffee"
    if any(token in value for token in ["prunus", "plum", "cherry", "cerasifera", "salicina", "avium"]):
        return "other_prunus"
    return None


def passes_quality(report: dict) -> bool:
    assembly = report.get("assembly_info", {})
    if report.get("source_database") != "SOURCE_DATABASE_REFSEQ":
        return False
    if assembly.get("assembly_level") not in LEVEL_RANK:
        return False
    checkm = report.get("checkm_info", {})
    completeness = float(checkm.get("completeness") or 0)
    contamination = float(checkm.get("contamination") or 0)
    return (not completeness or completeness >= 92.5) and (not contamination or contamination <= 3)


def quality_key(report: dict) -> tuple:
    assembly = report.get("assembly_info", {})
    stats = report.get("assembly_stats", {})
    return (
        LEVEL_RANK.get(assembly.get("assembly_level", ""), 99),
        int(stats.get("number_of_contigs") or 10_000),
        -int(stats.get("contig_n50") or 0),
        report.get("accession", ""),
    )


def select_unique_genomes() -> tuple[list[dict], list[dict]]:
    reports = json.loads(REPORT.read_text(encoding="utf-8"))["reports"]
    local_accessions = {p.stem for p in FASTA_DIR.glob("*.fna")}
    candidates: dict[str, list[dict]] = defaultdict(list)
    for report in reports:
        group = host_group(host_source(report))
        if group and passes_quality(report) and report.get("accession") in local_accessions:
            candidates[group].append(report)
    selected = []
    rejected = []
    for group in ["grapevine", "almond", "olive", "coffee", "other_prunus"]:
        seen = set()
        # Preserve the original rule (the ten highest-quality eligible assemblies),
        # then remove repeated strain identifiers instead of replacing them post hoc.
        for report in sorted(candidates[group], key=quality_key)[:10]:
            key = normalized_strain(strain_name(report)) or report["accession"].lower()
            if key in seen:
                rejected.append({"accession": report["accession"], "strain": strain_name(report), "group": group, "reason": "duplicate normalized strain name"})
                continue
            seen.add(key)
            selected.append(report)
    return selected, rejected


def read_fasta_segments(path: pathlib.Path) -> tuple[list[str], dict]:
    records: list[str] = []
    current: list[str] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.startswith(">"):
            if current:
                records.append("".join(current).upper())
            current = []
        else:
            current.append(line.strip())
    if current:
        records.append("".join(current).upper())
    valid_segments = []
    ambiguous = 0
    for record in records:
        ambiguous += sum(ch not in BASES for ch in record)
        valid_segments.extend(segment for segment in re.split(r"[^ACGT]+", record) if segment)
    stats = {
        "contigs": len(records),
        "ambiguous_bases": ambiguous,
        "valid_segments": len(valid_segments),
        "valid_bases": sum(map(len, valid_segments)),
    }
    return valid_segments, stats


def windows_from_segments(segments: list[str]) -> list[str]:
    windows = []
    for segment in segments:
        for start in range(0, len(segment) - WINDOW_SIZE + 1, WINDOW_SIZE):
            windows.append(segment[start : start + WINDOW_SIZE])
            if len(windows) == WINDOWS_PER_GENOME:
                return windows
    return windows


def triplet_ids(seq: str) -> np.ndarray:
    return np.fromiter((TRIPLET_INDEX[seq[i : i + 3]] for i in range(0, len(seq) - 2, 3)), dtype=np.int16)


def base_triplet_composition(seq: str) -> np.ndarray:
    n = len(seq)
    base = np.array([seq.count(base) / n for base in BASES], dtype=float)
    ids = triplet_ids(seq)
    triplets = np.bincount(ids, minlength=64).astype(float) / max(len(ids), 1)
    gc = (seq.count("G") + seq.count("C")) / n
    return np.concatenate([base, [gc], triplets])


def triplet_distribution(seq: str) -> np.ndarray:
    ids = triplet_ids(seq)
    return np.bincount(ids, minlength=64).astype(float) / max(len(ids), 1)


def pitch_order(seq: str, pitch_map: np.ndarray | None = None) -> np.ndarray:
    pitch_map = np.arange(36, 100, dtype=float) if pitch_map is None else pitch_map.astype(float)
    pitches = pitch_map[triplet_ids(seq)]
    intervals = np.diff(pitches)
    interval_values = np.arange(-63, 64)
    histogram = np.array([(intervals == value).sum() for value in interval_values], dtype=float)
    histogram /= max(histogram.sum(), 1)
    second = np.diff(intervals)
    scalars = np.array([
        np.mean(np.abs(intervals)),
        np.std(intervals),
        np.mean(np.abs(second)) if len(second) else 0.0,
    ])
    return np.concatenate([scalars, histogram])


def triplet_transition(seq: str) -> np.ndarray:
    ids = triplet_ids(seq)
    pairs = ids[:-1].astype(np.int32) * 64 + ids[1:].astype(np.int32)
    values = np.bincount(pairs, minlength=4096).astype(float)
    return values / max(values.sum(), 1)


def sixmer(seq: str) -> np.ndarray:
    base_index = {base: i for i, base in enumerate(BASES)}
    values = np.zeros(4096, dtype=float)
    if len(seq) < 6:
        return values
    code = 0
    for ch in seq[:6]:
        code = code * 4 + base_index[ch]
    values[code] += 1
    mask = 4 ** 5
    for i in range(6, len(seq)):
        code = (code % mask) * 4 + base_index[seq[i]]
        values[code] += 1
    return values / values.sum()


def compression_distance(a: str, b: str) -> float:
    ba, bb = a.encode("ascii"), b.encode("ascii")
    ca, cb = len(zlib.compress(ba, 9)), len(zlib.compress(bb, 9))
    cab = len(zlib.compress(ba + bb, 9))
    cba = len(zlib.compress(bb + ba, 9))
    denominator = max(ca, cb)
    return 0.5 * ((cab - min(ca, cb)) / denominator + (cba - min(ca, cb)) / denominator)


FEATURES = {
    "composition": base_triplet_composition,
    "triplet_distribution": triplet_distribution,
    "pitch_order": pitch_order,
    "triplet_transition": triplet_transition,
    "sixmer": sixmer,
}


def mutate(seq: str, rate: float, rng: random.Random) -> str:
    chars = list(seq)
    for index in rng.sample(range(len(chars)), max(1, int(len(chars) * rate))):
        chars[index] = rng.choice([base for base in BASES if base != chars[index]])
    return "".join(chars)


def anomaly_variant(window: str, anomaly: str, rng: random.Random, donor: str | None = None) -> str:
    if anomaly == "clustered_substitution":
        return mutate(window, 0.015, rng)
    if anomaly == "triplet_order_shuffle":
        start = (len(window) // 5) // 3 * 3
        end = (4 * len(window) // 5) // 3 * 3
        units = [window[i : i + 3] for i in range(start, end, 3)]
        rng.shuffle(units)
        return window[:start] + "".join(units) + window[end:]
    if anomaly == "foreign_segment":
        assert donor is not None
        size = len(window) // 5
        start = (len(window) - size) // 2
        return window[:start] + donor[:size] + window[start + size :]
    raise ValueError(anomaly)


def auc(labels: list[int], scores: list[float]) -> float:
    pos = [score for score, label in zip(scores, labels) if label]
    neg = [score for score, label in zip(scores, labels) if not label]
    wins = sum(1 if p > n else 0.5 if p == n else 0 for p in pos for n in neg)
    return wins / (len(pos) * len(neg))


def precision_at_four(labels: list[int], scores: list[float]) -> float:
    order = np.argsort(np.asarray(scores))[::-1][:4]
    return sum(labels[i] for i in order) / 4


def standardized_distance(before: np.ndarray, after: np.ndarray, scale: np.ndarray) -> float:
    return float(np.linalg.norm((after - before) / scale))


def genome_feature(windows: list[str], method: str) -> np.ndarray:
    matrix = np.vstack([FEATURES[method](window) for window in windows])
    return np.concatenate([matrix.mean(axis=0), matrix.std(axis=0)])


def standardize_fold(train: np.ndarray, test: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
    mean = train.mean(axis=0)
    scale = train.std(axis=0)
    scale[scale == 0] = 1
    return (train - mean) / scale, (test - mean) / scale


def loo_accuracy(features: np.ndarray, labels: list[str]) -> float:
    labels_array = np.asarray(labels)
    correct = 0
    for held in range(len(labels)):
        train_index = np.array([i for i in range(len(labels)) if i != held])
        train, test = standardize_fold(features[train_index], features[[held]])
        centroids = {label: train[labels_array[train_index] == label].mean(axis=0) for label in sorted(set(labels))}
        predicted = min(centroids, key=lambda label: np.linalg.norm(test[0] - centroids[label]))
        correct += predicted == labels[held]
    return correct / len(labels)


def permutation_p(features: np.ndarray, labels: list[str], observed: float, rng: random.Random) -> float:
    hits = 0
    labels_copy = labels[:]
    for _ in range(PERMUTATIONS):
        rng.shuffle(labels_copy)
        hits += loo_accuracy(features, labels_copy) >= observed
    return (hits + 1) / (PERMUTATIONS + 1)


def minhash_sketch(segments: list[str], k: int = 21, size: int = 2_000) -> set[int]:
    values = []
    for segment in segments:
        for i in range(0, max(0, len(segment) - k + 1), 5):
            digest = hashlib.blake2b(segment[i : i + k].encode("ascii"), digest_size=8).digest()
            values.append(int.from_bytes(digest, "big"))
    return set(sorted(values)[:size])


def write_csv(path: pathlib.Path, rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        return
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    RESULTS.mkdir(parents=True, exist_ok=True)
    rng = random.Random(SEED)
    selected, rejected = select_unique_genomes()
    selection_rows = []
    windows_by_accession = {}
    preprocessing_rows = []
    sketches = {}
    for report in selected:
        accession = report["accession"]
        group = host_group(host_source(report))
        segments, stats = read_fasta_segments(FASTA_DIR / f"{accession}.fna")
        windows = windows_from_segments(segments)
        if len(windows) < WINDOWS_PER_GENOME:
            raise RuntimeError(f"{accession} yielded only {len(windows)} valid windows")
        windows_by_accession[accession] = windows
        sketches[accession] = minhash_sketch(segments)
        selection_rows.append({
            "accession": accession,
            "strain": strain_name(report),
            "host_group": group,
            "host_or_source": host_source(report),
            "assembly_level": report["assembly_info"].get("assembly_level", ""),
            "selection_rule": "best quality locally available RefSeq assembly after normalized strain-name deduplication",
        })
        preprocessing_rows.append({"accession": accession, **stats, "windows_used": len(windows)})

    redundancy_rows = []
    accessions = [row["accession"] for row in selection_rows]
    for i, a in enumerate(accessions):
        for b in accessions[i + 1 :]:
            union = sketches[a] | sketches[b]
            similarity = len(sketches[a] & sketches[b]) / max(len(union), 1)
            redundancy_rows.append({"accession_a": a, "accession_b": b, "minhash_jaccard_21mer": similarity})
    redundancy_rows.sort(key=lambda row: row["minhash_jaccard_21mer"], reverse=True)

    labels = [row["host_group"] for row in selection_rows]
    classification = []
    for method in ["composition", "pitch_order"]:
        matrix = np.vstack([genome_feature(windows_by_accession[row["accession"]], method) for row in selection_rows])
        observed = loo_accuracy(matrix, labels)
        p_value = permutation_p(matrix, labels, observed, random.Random(SEED + len(classification)))
        classification.append({"method": method, "accuracy": observed, "permutations": PERMUTATIONS, "p_value": p_value})

    reference_features = {}
    scales = {}
    all_windows = [window for accession in accessions for window in windows_by_accession[accession]]
    for method, function in FEATURES.items():
        matrix = np.vstack([function(window) for window in all_windows])
        reference_features[method] = matrix
        scale = matrix.std(axis=0)
        scale[scale == 0] = 1
        scales[method] = scale

    by_group = defaultdict(list)
    for row in selection_rows:
        by_group[row["host_group"]].append(row["accession"])
    anomaly_detail = []
    example_rows = []
    for accession in accessions:
        windows = windows_by_accession[accession]
        candidates = list(range(5, len(windows) - 5))
        anomaly_indices = sorted(rng.sample(candidates, ANOMALIES_PER_GENOME))
        source_group = next(row["host_group"] for row in selection_rows if row["accession"] == accession)
        donor_group = rng.choice([group for group in by_group if group != source_group])
        donor_accession = rng.choice(by_group[donor_group])
        donor_pool = windows_by_accession[donor_accession]
        for anomaly_type in ["clustered_substitution", "triplet_order_shuffle", "foreign_segment"]:
            modified = [mutate(window, BACKGROUND_RATE, rng) for window in windows]
            donors_used = {}
            for idx in anomaly_indices:
                donor_idx = rng.randrange(len(donor_pool))
                donors_used[idx] = donor_idx
                modified[idx] = anomaly_variant(modified[idx], anomaly_type, rng, donor_pool[donor_idx])
            labels_binary = [int(i in anomaly_indices) for i in range(len(windows))]
            for method in METHODS:
                scores = []
                for before, after in zip(windows, modified):
                    if method == "compression":
                        score = compression_distance(before, after)
                    else:
                        score = standardized_distance(FEATURES[method](before), FEATURES[method](after), scales[method])
                    scores.append(score)
                anomaly_detail.append({
                    "accession": accession,
                    "host_group": source_group,
                    "anomaly_type": anomaly_type,
                    "method": method,
                    "auc": auc(labels_binary, scores),
                    "precision_at_4": precision_at_four(labels_binary, scores),
                    "anomaly_indices": ";".join(map(str, anomaly_indices)),
                    "donor_group": donor_group if anomaly_type == "foreign_segment" else "",
                    "donor_accession": donor_accession if anomaly_type == "foreign_segment" else "",
                    "donor_window_indices": ";".join(f"{idx}:{donors_used[idx]}" for idx in anomaly_indices) if anomaly_type == "foreign_segment" else "",
                })
                if accession == accessions[0] and anomaly_type == "foreign_segment" and method == "pitch_order":
                    for index, (label, score) in enumerate(zip(labels_binary, scores)):
                        example_rows.append({"accession": accession, "window_index": index, "anomaly": label, "score": score, "method": method, "anomaly_type": anomaly_type})

    grouped = defaultdict(list)
    for row in anomaly_detail:
        grouped[(row["method"], row["anomaly_type"])].append(row)
    summary_rows = []
    boot_rng = random.Random(SEED + 400)
    for (method, anomaly_type), rows in sorted(grouped.items()):
        auc_values = [float(row["auc"]) for row in rows]
        bootstrap = []
        for _ in range(2_000):
            bootstrap.append(np.mean([boot_rng.choice(auc_values) for _ in auc_values]))
        summary_rows.append({
            "method": method,
            "anomaly_type": anomaly_type,
            "mean_auc": np.mean(auc_values),
            "ci_low": np.percentile(bootstrap, 2.5),
            "ci_high": np.percentile(bootstrap, 97.5),
            "mean_precision_at_4": np.mean([float(row["precision_at_4"]) for row in rows]),
        })

    random_map_rows = []
    shuffle_records = [row for row in anomaly_detail if row["anomaly_type"] == "triplet_order_shuffle" and row["method"] == "pitch_order"]
    # Recreate the same anomaly realization deterministically for a mapping-sensitivity test.
    for map_index in range(RANDOM_MAPS):
        map_rng = random.Random(SEED + 10_000 + map_index)
        permutation = list(range(36, 100))
        map_rng.shuffle(permutation)
        mapping = np.asarray(permutation, dtype=float)
        realization_rng = random.Random(SEED)
        aucs = []
        for accession in accessions:
            windows = windows_by_accession[accession]
            indices = sorted(realization_rng.sample(list(range(5, len(windows) - 5)), ANOMALIES_PER_GENOME))
            # Advance donor choices consistently with the primary experiment.
            source_group = next(row["host_group"] for row in selection_rows if row["accession"] == accession)
            donor_group = realization_rng.choice([group for group in by_group if group != source_group])
            donor_accession = realization_rng.choice(by_group[donor_group])
            donor_pool = windows_by_accession[donor_accession]
            modified = [mutate(window, BACKGROUND_RATE, realization_rng) for window in windows]
            for idx in indices:
                donor = donor_pool[realization_rng.randrange(len(donor_pool))]
                modified[idx] = anomaly_variant(modified[idx], "triplet_order_shuffle", realization_rng, donor)
            binary = [int(i in indices) for i in range(len(windows))]
            original_matrix = np.vstack([pitch_order(window, mapping) for window in windows])
            scale = original_matrix.std(axis=0)
            scale[scale == 0] = 1
            scores = [standardized_distance(pitch_order(a, mapping), pitch_order(b, mapping), scale) for a, b in zip(windows, modified)]
            aucs.append(auc(binary, scores))
        random_map_rows.append({"mapping": map_index + 1, "mean_auc_triplet_order_shuffle": np.mean(aucs), "seed": SEED + 10_000 + map_index})

    write_csv(RESULTS / "selected_unique_genomes.csv", selection_rows)
    write_csv(RESULTS / "excluded_duplicate_strain_accessions.csv", rejected)
    write_csv(RESULTS / "preprocessing_audit.csv", preprocessing_rows)
    write_csv(RESULTS / "redundancy_minhash_pairs.csv", redundancy_rows)
    write_csv(RESULTS / "classification_revised.csv", classification)
    write_csv(RESULTS / "anomaly_detection_revised_by_genome.csv", anomaly_detail)
    write_csv(RESULTS / "anomaly_detection_revised_summary.csv", summary_rows)
    write_csv(RESULTS / "mapping_sensitivity.csv", random_map_rows)
    write_csv(RESULTS / "example_track_revised.csv", example_rows)
    report = {
        "parameters": {
            "window_size": WINDOW_SIZE,
            "windows_per_genome": WINDOWS_PER_GENOME,
            "anomalies_per_genome": ANOMALIES_PER_GENOME,
            "background_substitution_rate": BACKGROUND_RATE,
            "random_seed": SEED,
            "permutations": PERMUTATIONS,
            "bootstrap_resamples": 2_000,
            "random_pitch_mappings": RANDOM_MAPS,
        },
        "host_groups": dict(Counter(labels)),
        "classification": classification,
        "top_redundancy_pairs": redundancy_rows[:10],
        "anomaly_summary": summary_rows,
        "mapping_sensitivity": {
            "mean": float(np.mean([row["mean_auc_triplet_order_shuffle"] for row in random_map_rows])),
            "min": float(np.min([row["mean_auc_triplet_order_shuffle"] for row in random_map_rows])),
            "max": float(np.max([row["mean_auc_triplet_order_shuffle"] for row in random_map_rows])),
        },
        "preprocessing_totals": {
            "contigs": sum(row["contigs"] for row in preprocessing_rows),
            "ambiguous_bases": sum(row["ambiguous_bases"] for row in preprocessing_rows),
            "valid_segments": sum(row["valid_segments"] for row in preprocessing_rows),
        },
    }
    (RESULTS / "revision_summary.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(json.dumps(report, indent=2))


if __name__ == "__main__":
    main()
