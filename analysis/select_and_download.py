from __future__ import annotations

import csv
import gzip
import json
import pathlib
import re
import urllib.request


ROOT = pathlib.Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
FASTA_DIR = DATA / "fastas"
REPORT = DATA / "ncbi_xylella_dataset_report.json"
SELECTED = DATA / "study_genomes.csv"

TARGET_GROUPS = {
    "grapevine": {
        "n": 10,
        "description": "Vitis/grape-associated assemblies; Pierce's disease context",
    },
    "almond": {
        "n": 10,
        "description": "Prunus dulcis/almond-associated assemblies; almond leaf scorch context",
    },
    "olive": {
        "n": 10,
        "description": "Olea/olive-associated assemblies; olive quick decline context",
    },
    "coffee": {
        "n": 10,
        "description": "Coffea/coffee-associated assemblies; coffee leaf scorch context",
    },
    "other_prunus": {
        "n": 10,
        "description": "Non-almond Prunus/plum/cherry-associated assemblies",
    },
}

LEVEL_RANK = {"Complete Genome": 0, "Chromosome": 1, "Scaffold": 2}


def attrs(report: dict) -> dict[str, str]:
    biosample = report.get("assembly_info", {}).get("biosample", {})
    out = {}
    for item in biosample.get("attributes", []):
        name = item.get("name")
        value = item.get("value")
        if name and value and value != "missing":
            out[name] = value
    return out


def strain(report: dict) -> str:
    names = report.get("organism", {}).get("infraspecific_names", {})
    return names.get("strain") or attrs(report).get("strain") or ""


def host_or_source(report: dict) -> str:
    at = attrs(report)
    return at.get("host") or at.get("isolation_source") or ""


def host_group(host: str) -> str | None:
    s = (host or "").lower()
    if "vitis" in s or "grape" in s:
        return "grapevine"
    if "prunus dulcis" in s or "almond" in s:
        return "almond"
    if "olea" in s or "olive" in s:
        return "olive"
    if "coffea" in s or "coffee" in s:
        return "coffee"
    if any(token in s for token in ["prunus", "plum", "cherry", "cerasifera", "salicina", "avium"]):
        return "other_prunus"
    return None


def inferred_subspecies(report: dict) -> str:
    ani = report.get("average_nucleotide_identity", {})
    for key in ("submitted_ani_match", "best_ani_match"):
        name = ani.get(key, {}).get("organism_name", "")
        match = re.search(r"subsp\.\s+([A-Za-z0-9_-]+)", name)
        if match:
            return match.group(1).lower()
    return "not_resolved"


def ftp_url(accession: str, assembly_name: str) -> str:
    match = re.match(r"^(GC[AF])_(\d{3})(\d{3})(\d{3})\.(\d+)$", accession)
    if not match:
        raise ValueError(f"Unexpected accession format: {accession}")
    prefix, a, b, c, _version = match.groups()
    base = f"{accession}_{assembly_name}"
    return f"https://ftp.ncbi.nlm.nih.gov/genomes/all/{prefix}/{a}/{b}/{c}/{base}/{base}_genomic.fna.gz"


def read_fasta(path: pathlib.Path) -> str:
    text = path.read_text(encoding="utf-8")
    return "".join(line.strip().upper() for line in text.splitlines() if not line.startswith(">"))


def quality_score(report: dict) -> tuple[int, int, int, str]:
    ai = report.get("assembly_info", {})
    stats = report.get("assembly_stats", {})
    level = ai.get("assembly_level", "")
    contigs = int(stats.get("number_of_contigs") or 10_000)
    n50 = int(stats.get("contig_n50") or 0)
    return (LEVEL_RANK.get(level, 99), contigs, -n50, report.get("accession", ""))


def passes_quality(report: dict) -> bool:
    ai = report.get("assembly_info", {})
    if report.get("source_database") != "SOURCE_DATABASE_REFSEQ":
        return False
    if ai.get("assembly_level") not in LEVEL_RANK:
        return False
    checkm = report.get("checkm_info", {})
    completeness = float(checkm.get("completeness") or 0)
    contamination = float(checkm.get("contamination") or 0)
    if completeness and completeness < 92.5:
        return False
    if contamination and contamination > 3:
        return False
    return True


def main() -> None:
    DATA.mkdir(exist_ok=True)
    FASTA_DIR.mkdir(exist_ok=True)
    reports = json.loads(REPORT.read_text(encoding="utf-8"))["reports"]
    candidates: dict[str, list[dict]] = {group: [] for group in TARGET_GROUPS}
    for report in reports:
        if not passes_quality(report):
            continue
        group = host_group(host_or_source(report))
        if group in candidates:
            candidates[group].append(report)

    selected: list[dict] = []
    for group, settings in TARGET_GROUPS.items():
        group_reports = sorted(candidates[group], key=quality_score)[: settings["n"]]
        if len(group_reports) < settings["n"]:
            raise RuntimeError(f"Only found {len(group_reports)} usable assemblies for {group}")
        selected.extend(group_reports)

    rows = []
    for report in selected:
        ai = report["assembly_info"]
        at = attrs(report)
        accession = report["accession"]
        assembly_name = ai["assembly_name"]
        out = FASTA_DIR / f"{accession}.fna"
        url = ftp_url(accession, assembly_name)
        if not out.exists():
            raw = urllib.request.urlopen(url, timeout=120).read()
            out.write_bytes(gzip.decompress(raw))
        seq_len = len(read_fasta(out))
        host = host_or_source(report)
        group = host_group(host)
        rows.append(
            {
                "accession": accession,
                "assembly_name": assembly_name,
                "assembly_level": ai.get("assembly_level", ""),
                "strain": strain(report),
                "host_group": group,
                "host_or_source": host,
                "host_group_description": TARGET_GROUPS[group]["description"],
                "inferred_subspecies_from_ani": inferred_subspecies(report),
                "geo_loc_name": at.get("geo_loc_name", ""),
                "release_date": ai.get("release_date", ""),
                "sequence_length": seq_len,
                "gc_percent_ncbi": report.get("assembly_stats", {}).get("gc_percent", ""),
                "checkm_completeness": report.get("checkm_info", {}).get("completeness", ""),
                "checkm_contamination": report.get("checkm_info", {}).get("contamination", ""),
                "fasta": str(out.relative_to(ROOT)),
                "ncbi_ftp_url": url,
                "selection_note": "RefSeq assembly; quality-filtered; complete/chromosome preferred before scaffold within host group",
            }
        )

    with SELECTED.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)
    print(f"Wrote {SELECTED} with {len(rows)} assemblies across {len(TARGET_GROUPS)} host groups")
    for group in TARGET_GROUPS:
        print(group, sum(1 for row in rows if row["host_group"] == group))


if __name__ == "__main__":
    main()
