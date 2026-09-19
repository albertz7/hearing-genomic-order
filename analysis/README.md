# Genomic sonification analysis pipeline

This folder contains the reproducible analysis used for the revised manuscript.
The current study design treats sonification as a windowed genomic anomaly
display method rather than as a replacement for alignment or variant calling.

The pipeline uses NCBI RefSeq Xylella fastidiosa assemblies from five balanced
host/pathosystem groups. It creates deterministic codon-to-pitch and
amino-acid-class sonifications, evaluates whether audio-derived features retain
host-group signal, and tests whether local synthetic anomalies can be detected
from window-level sonification features.

Run order:

1. `python analysis/select_and_download.py`
2. `python analysis/run_benchmark.py`

Key outputs:

- `data/study_genomes.csv`: selected assemblies and metadata.
- `results/classification_results.csv`: host-group representation test.
- `results/anomaly_detection_summary.csv`: local anomaly detection AUCs.
- `results/figures/*.svg`: manuscript-ready figures.
- `output/audio/*_anomaly_track.mid`: playable MIDI anomaly display.
