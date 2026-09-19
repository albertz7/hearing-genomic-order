# Hearing Genomic Order — project export

This archive collects the manuscripts, code, public NCBI genome data, derived
results, figures, audio examples, reviewer materials, and the current NHSJS
follow-up analysis for the project commonly titled:

**Hearing Genomic Order: Windowed Sonification for Representing Mutation and
Anomaly Patterns in *Xylella fastidiosa* Assemblies**

## Folder map

- `analysis/` — established deterministic analysis and document-building code.
- `data/` — accession metadata and downloaded NCBI RefSeq FASTA assemblies.
- `results/` — outputs of the established benchmark, including figures.
- `audio/` — MIDI previews and an example anomaly track.
- `manuscripts/` — Word/PDF manuscript versions and response letters prepared
  for NHSJS, Critical Debates, and Curieux Review.
- `reviewer_materials/` — the downloaded NHSJS decision, response-letter
  example, and submission guidelines.
- `source_materials/` — text extractions used while revising the paper.
- `nhsjs_revision/` — the reviewer-driven follow-up analysis and its current
  outputs. These results are newer than the manuscript and have not yet been
  incorporated into a final author-written revision.

## Established pipeline

From the export root, install the dependencies in `requirements.txt`, then run:

```text
python analysis/select_and_download.py
python analysis/run_benchmark.py
```

The scripts use deterministic random seeds. `select_and_download.py` accesses
public NCBI resources; the included `data/fastas/` directory lets the benchmark
be inspected without downloading the assemblies again.

## NHSJS follow-up analysis

Run:

```text
python nhsjs_revision/revision_analysis.py
```

That script adds strain deduplication, contig-aware preprocessing, 1,000-label
permutation tests, sequence-order baselines, redundancy checks, and pitch-mapping
sensitivity checks. Its results should be independently reviewed before being
used in a manuscript.

## Scope and provenance

- Genome assemblies and metadata are public NCBI RefSeq records identified by
  accession in the included tables.
- The analysis evaluates representation and detection of synthetic sequence
  anomalies. It does not claim clinical or biological diagnosis.
- The archive excludes Git internals, Python caches, temporary page renders,
  and redundant visual-QA images.
- Some manuscript-building scripts retain placeholders such as
  `[School Name or District]` and `[Corresponding Email]`; verify all author and
  submission metadata before using generated documents.

## Archive variants

The full archive contains the downloaded FASTA files. The lightweight archive
contains the same code, manuscripts, results, reviewer materials, and metadata,
but omits `data/fastas/`; rerun `analysis/select_and_download.py` to restore them.
