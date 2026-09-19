from __future__ import annotations

import csv
import hashlib
import json
import pathlib
import random
from collections import Counter, defaultdict

import numpy as np

from sonification_core import (
    audio_features,
    auc_from_scores,
    bootstrap_ci,
    composition_features,
    euclidean,
    iter_windows,
    mutate_substitutions,
    precision_at_k,
    read_fasta,
    replace_middle,
    reverse_complement,
    write_midi_from_scores,
)


ROOT = pathlib.Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
RESULTS = ROOT / "results"
FIGURES = RESULTS / "figures"
AUDIO = ROOT / "output" / "audio"
SELECTED = DATA / "study_genomes.csv"

WINDOW_SIZE = 3_000
MAX_WINDOWS_PER_GENOME = 60
ANOMALY_WINDOWS_PER_GENOME = 4
BACKGROUND_SUBSTITUTION_RATE = 0.001
RANDOM_SEED = 20260712
MAPS = ["codon64", "amino_class"]


def standardize(train: np.ndarray, test: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
    mean = train.mean(axis=0)
    std = train.std(axis=0)
    std[std == 0] = 1.0
    return (train - mean) / std, (test - mean) / std


def leave_one_out_nearest_centroid(features: np.ndarray, labels: list[str]) -> tuple[float, list[dict]]:
    predictions = []
    correct = 0
    labels_arr = np.array(labels)
    for idx in range(len(labels)):
        train_idx = np.array([i for i in range(len(labels)) if i != idx])
        test_idx = np.array([idx])
        train, test = standardize(features[train_idx], features[test_idx])
        centroids = {}
        for label in sorted(set(labels)):
            centroids[label] = train[labels_arr[train_idx] == label].mean(axis=0)
        distances = {label: euclidean(test[0], centroid) for label, centroid in centroids.items()}
        pred = min(distances, key=distances.get)
        is_correct = pred == labels[idx]
        correct += int(is_correct)
        predictions.append(
            {
                "accession": "",
                "true_label": labels[idx],
                "predicted_label": pred,
                "correct": int(is_correct),
                "nearest_distance": distances[pred],
            }
        )
    return correct / len(labels), predictions


def permutation_p_value(features: np.ndarray, labels: list[str], observed: float, rng: random.Random, rounds: int = 50) -> float:
    hits = 0
    labels_copy = labels[:]
    for _ in range(rounds):
        rng.shuffle(labels_copy)
        acc, _ = leave_one_out_nearest_centroid(features, labels_copy)
        if acc >= observed:
            hits += 1
    return (hits + 1) / (rounds + 1)


def genome_level_features(windows: list[str], kind: str) -> np.ndarray:
    if kind == "composition":
        matrix = np.vstack([composition_features(window) for window in windows])
    else:
        matrix = np.vstack([audio_features(window, kind) for window in windows])
    return np.concatenate([matrix.mean(axis=0), matrix.std(axis=0)])


def anomaly_variant(window: str, anomaly_type: str, rng: random.Random, donor_window: str | None = None) -> str:
    if anomaly_type == "clustered_substitution":
        return mutate_substitutions(window, 0.015, rng)
    if anomaly_type == "codon_order_shuffle":
        start = len(window) // 5
        end = 4 * len(window) // 5
        start -= start % 3
        end -= end % 3
        codons = [window[i : i + 3] for i in range(start, end, 3)]
        rng.shuffle(codons)
        return window[:start] + "".join(codons) + window[end:]
    if anomaly_type == "foreign_segment":
        if donor_window is None:
            raise ValueError("foreign_segment requires a donor window")
        return replace_middle(window, donor_window, size=len(window) // 5)
    raise ValueError(f"Unknown anomaly type: {anomaly_type}")


def evaluate_anomalies(rows: list[dict], windows_by_accession: dict[str, list[str]], rng: random.Random) -> tuple[list[dict], list[dict]]:
    row_by_accession = {row["accession"]: row for row in rows}
    accessions = list(windows_by_accession)
    by_group = defaultdict(list)
    for row in rows:
        by_group[row["host_group"]].append(row["accession"])

    detail_rows = []
    summary_rows = []
    for accession in accessions:
        windows = windows_by_accession[accession]
        if len(windows) < 20:
            continue
        candidate_indexes = list(range(5, len(windows) - 5))
        anomaly_indexes = sorted(rng.sample(candidate_indexes, ANOMALY_WINDOWS_PER_GENOME))
        source_group = row_by_accession[accession]["host_group"]
        donor_group = rng.choice([group for group in by_group if group != source_group])
        donor_accession = rng.choice(by_group[donor_group])
        donor_windows = windows_by_accession[donor_accession]
        donor_pool = donor_windows[5 : min(len(donor_windows), len(windows) - 5)] or donor_windows

        for anomaly_type in ["clustered_substitution", "codon_order_shuffle", "foreign_segment"]:
            modified = [mutate_substitutions(window, BACKGROUND_SUBSTITUTION_RATE, rng) for window in windows]
            for idx in anomaly_indexes:
                donor = rng.choice(donor_pool)
                modified[idx] = anomaly_variant(modified[idx], anomaly_type, rng, donor)
            labels = [1 if idx in anomaly_indexes else 0 for idx in range(len(windows))]

            feature_kinds = ["composition"] + MAPS
            for kind in feature_kinds:
                scores = []
                for before, after in zip(windows, modified):
                    if kind == "composition":
                        score = euclidean(composition_features(before), composition_features(after))
                    else:
                        score = euclidean(audio_features(before, kind), audio_features(after, kind))
                    scores.append(score)
                auc = auc_from_scores(labels, scores)
                p_at_k = precision_at_k(labels, scores, ANOMALY_WINDOWS_PER_GENOME)
                summary_rows.append(
                    {
                        "accession": accession,
                        "host_group": source_group,
                        "anomaly_type": anomaly_type,
                        "feature_kind": kind,
                        "auc": auc,
                        "precision_at_k": p_at_k,
                        "anomaly_windows": ANOMALY_WINDOWS_PER_GENOME,
                        "total_windows": len(windows),
                    }
                )
                if accession == accessions[0] and anomaly_type == "foreign_segment" and kind == "codon64":
                    AUDIO.mkdir(parents=True, exist_ok=True)
                    write_midi_from_scores(scores, AUDIO / f"{accession}_foreign_segment_anomaly_track.mid")
                    for idx, score in enumerate(scores):
                        detail_rows.append(
                            {
                                "accession": accession,
                                "window_index": idx,
                                "label": labels[idx],
                                "score": score,
                                "feature_kind": kind,
                                "anomaly_type": anomaly_type,
                            }
                        )
    return summary_rows, detail_rows


def write_csv(path: pathlib.Path, rows: list[dict]) -> None:
    if not rows:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)


def write_bar_svg(path: pathlib.Path, rows: list[dict], metric: str, title: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    labels = [row["label"] for row in rows]
    values = [float(row[metric]) for row in rows]
    width = 760
    height = 420
    left = 130
    bottom = 350
    bar_h = 38
    max_v = max(values + [1.0])
    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">',
        '<rect width="100%" height="100%" fill="white"/>',
        f'<text x="{width/2}" y="32" text-anchor="middle" font-family="Arial" font-size="18" font-weight="700">{title}</text>',
    ]
    for i, (label, value) in enumerate(zip(labels, values)):
        y = 70 + i * 58
        w = int((width - left - 90) * value / max_v)
        parts.append(f'<text x="{left-8}" y="{y+24}" text-anchor="end" font-family="Arial" font-size="13">{label}</text>')
        parts.append(f'<rect x="{left}" y="{y}" width="{w}" height="{bar_h}" fill="#4C78A8"/>')
        parts.append(f'<text x="{left+w+8}" y="{y+24}" font-family="Arial" font-size="13">{value:.3f}</text>')
    parts.append(f'<line x1="{left}" y1="{bottom}" x2="{width-60}" y2="{bottom}" stroke="#222"/>')
    parts.append("</svg>")
    path.write_text("\n".join(parts), encoding="utf-8")


def write_line_svg(path: pathlib.Path, rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        return
    scores = [float(row["score"]) for row in rows]
    labels = [int(row["label"]) for row in rows]
    width = 900
    height = 360
    left = 55
    top = 45
    plot_w = width - 95
    plot_h = height - 95
    hi = max(scores) or 1.0
    points = []
    for i, score in enumerate(scores):
        x = left + (plot_w * i / max(len(scores) - 1, 1))
        y = top + plot_h - (plot_h * score / hi)
        points.append(f"{x:.1f},{y:.1f}")
    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">',
        '<rect width="100%" height="100%" fill="white"/>',
        '<text x="450" y="26" text-anchor="middle" font-family="Arial" font-size="18" font-weight="700">Example sonified anomaly score track</text>',
        f'<line x1="{left}" y1="{top+plot_h}" x2="{left+plot_w}" y2="{top+plot_h}" stroke="#333"/>',
        f'<line x1="{left}" y1="{top}" x2="{left}" y2="{top+plot_h}" stroke="#333"/>',
        f'<polyline points="{" ".join(points)}" fill="none" stroke="#4C78A8" stroke-width="2"/>',
    ]
    for i, label in enumerate(labels):
        if label:
            x = left + (plot_w * i / max(len(scores) - 1, 1))
            parts.append(f'<rect x="{x-2}" y="{top}" width="4" height="{plot_h}" fill="#F58518" opacity="0.25"/>')
    parts.append('<text x="450" y="340" text-anchor="middle" font-family="Arial" font-size="13">Genome window index; orange bands are inserted anomaly windows</text>')
    parts.append("</svg>")
    path.write_text("\n".join(parts), encoding="utf-8")


def main() -> None:
    rng = random.Random(RANDOM_SEED)
    RESULTS.mkdir(exist_ok=True)
    FIGURES.mkdir(exist_ok=True)
    AUDIO.mkdir(parents=True, exist_ok=True)

    rows = list(csv.DictReader(SELECTED.open(encoding="utf-8")))
    windows_by_accession = {}
    for row in rows:
        seq = read_fasta(ROOT / row["fasta"])
        windows_by_accession[row["accession"]] = iter_windows(seq, WINDOW_SIZE)[:MAX_WINDOWS_PER_GENOME]

    labels = [row["host_group"] for row in rows]
    classification_rows = []
    prediction_rows = []
    feature_matrices = {}
    for kind in ["composition"] + MAPS:
        matrix = np.vstack([genome_level_features(windows_by_accession[row["accession"]], kind) for row in rows])
        feature_matrices[kind] = matrix
        accuracy, predictions = leave_one_out_nearest_centroid(matrix, labels)
        p_value = permutation_p_value(matrix, labels, accuracy, rng)
        classification_rows.append({"feature_kind": kind, "accuracy": accuracy, "permutation_p_value": p_value, "label": kind})
        for row, pred in zip(rows, predictions):
            pred["accession"] = row["accession"]
            pred["feature_kind"] = kind
            prediction_rows.append(pred)

    anomaly_rows, detail_rows = evaluate_anomalies(rows, windows_by_accession, rng)
    grouped = defaultdict(list)
    for row in anomaly_rows:
        grouped[(row["feature_kind"], row["anomaly_type"])].append(float(row["auc"]))
    rng_ci = random.Random(RANDOM_SEED + 1)
    anomaly_summary = []
    for (feature_kind, anomaly_type), values in sorted(grouped.items()):
        lo, hi = bootstrap_ci(values, rng_ci)
        pvals = [float(row["precision_at_k"]) for row in anomaly_rows if row["feature_kind"] == feature_kind and row["anomaly_type"] == anomaly_type]
        anomaly_summary.append(
            {
                "feature_kind": feature_kind,
                "anomaly_type": anomaly_type,
                "mean_auc": float(np.mean(values)),
                "auc_ci_low": lo,
                "auc_ci_high": hi,
                "mean_precision_at_k": float(np.mean(pvals)),
                "label": f"{feature_kind}: {anomaly_type}",
            }
        )

    write_csv(RESULTS / "classification_results.csv", classification_rows)
    write_csv(RESULTS / "classification_predictions.csv", prediction_rows)
    write_csv(RESULTS / "anomaly_detection_by_genome.csv", anomaly_rows)
    write_csv(RESULTS / "anomaly_detection_summary.csv", anomaly_summary)
    write_csv(RESULTS / "example_anomaly_window_scores.csv", detail_rows)

    write_bar_svg(FIGURES / "classification_accuracy.svg", classification_rows, "accuracy", "Leave-one-genome-out host-group classification")
    write_bar_svg(FIGURES / "anomaly_auc.svg", anomaly_summary, "mean_auc", "Mean AUC for local anomaly detection")
    write_line_svg(FIGURES / "example_anomaly_track.svg", detail_rows)

    best_audio_auc = max(
        (row for row in anomaly_summary if row["feature_kind"] in MAPS and row["anomaly_type"] == "foreign_segment"),
        key=lambda r: r["mean_auc"],
    )
    codon_order_rows = [row for row in anomaly_summary if row["anomaly_type"] == "codon_order_shuffle"]
    summary = {
        "study_genomes": len(rows),
        "host_groups": dict(Counter(labels)),
        "window_size_bp": WINDOW_SIZE,
        "max_windows_per_genome": MAX_WINDOWS_PER_GENOME,
        "anomaly_windows_per_genome": ANOMALY_WINDOWS_PER_GENOME,
        "classification": classification_rows,
        "best_foreign_segment_audio_auc": best_audio_auc,
        "codon_order_shuffle_results": codon_order_rows,
        "random_seed": RANDOM_SEED,
        "results_hash": hashlib.sha256((RESULTS / "anomaly_detection_summary.csv").read_bytes()).hexdigest(),
    }
    (RESULTS / "summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
