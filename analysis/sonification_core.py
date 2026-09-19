from __future__ import annotations

import math
import pathlib
import random
import struct
from collections import Counter

import numpy as np


BASES = "ACGT"
CODONS = [a + b + c for a in BASES for b in BASES for c in BASES]
CODON_INDEX = {codon: i for i, codon in enumerate(CODONS)}
COMPLEMENT = str.maketrans("ACGT", "TGCA")

CODON_TO_AA = {
    "TTT": "F", "TTC": "F", "TTA": "L", "TTG": "L",
    "TCT": "S", "TCC": "S", "TCA": "S", "TCG": "S",
    "TAT": "Y", "TAC": "Y", "TAA": "*", "TAG": "*",
    "TGT": "C", "TGC": "C", "TGA": "*", "TGG": "W",
    "CTT": "L", "CTC": "L", "CTA": "L", "CTG": "L",
    "CCT": "P", "CCC": "P", "CCA": "P", "CCG": "P",
    "CAT": "H", "CAC": "H", "CAA": "Q", "CAG": "Q",
    "CGT": "R", "CGC": "R", "CGA": "R", "CGG": "R",
    "ATT": "I", "ATC": "I", "ATA": "I", "ATG": "M",
    "ACT": "T", "ACC": "T", "ACA": "T", "ACG": "T",
    "AAT": "N", "AAC": "N", "AAA": "K", "AAG": "K",
    "AGT": "S", "AGC": "S", "AGA": "R", "AGG": "R",
    "GTT": "V", "GTC": "V", "GTA": "V", "GTG": "V",
    "GCT": "A", "GCC": "A", "GCA": "A", "GCG": "A",
    "GAT": "D", "GAC": "D", "GAA": "E", "GAG": "E",
    "GGT": "G", "GGC": "G", "GGA": "G", "GGG": "G",
}

AA_CLASS = {
    "A": "hydrophobic", "V": "hydrophobic", "I": "hydrophobic", "L": "hydrophobic", "M": "hydrophobic",
    "F": "aromatic", "Y": "aromatic", "W": "aromatic",
    "S": "polar", "T": "polar", "N": "polar", "Q": "polar", "C": "polar",
    "D": "acidic", "E": "acidic",
    "K": "basic", "R": "basic", "H": "basic",
    "G": "special", "P": "special",
    "*": "stop",
}
CLASS_ORDER = ["hydrophobic", "polar", "acidic", "basic", "aromatic", "special", "stop"]
CLASS_PITCH = {
    "hydrophobic": 48,
    "polar": 52,
    "acidic": 55,
    "basic": 59,
    "aromatic": 62,
    "special": 65,
    "stop": 36,
}


def read_fasta(path: pathlib.Path) -> str:
    text = path.read_text(encoding="utf-8")
    return "".join(ch for line in text.splitlines() if not line.startswith(">") for ch in line.upper() if ch in BASES)


def reverse_complement(seq: str) -> str:
    return seq.translate(COMPLEMENT)[::-1]


def iter_windows(seq: str, window_size: int) -> list[str]:
    usable = len(seq) - (len(seq) % window_size)
    return [seq[i : i + window_size] for i in range(0, usable, window_size)]


def shannon(values: list[int]) -> float:
    if not values:
        return 0.0
    counts = Counter(values)
    total = len(values)
    return -sum((count / total) * math.log2(count / total) for count in counts.values())


def codon_pitches(seq: str, map_name: str = "codon64") -> list[int]:
    pitches: list[int] = []
    for i in range(0, len(seq) - 2, 3):
        codon = seq[i : i + 3]
        if codon not in CODON_INDEX:
            continue
        if map_name == "codon64":
            pitches.append(36 + CODON_INDEX[codon])
        elif map_name == "amino_class":
            aa = CODON_TO_AA[codon]
            pitches.append(CLASS_PITCH[AA_CLASS[aa]])
        else:
            raise ValueError(f"Unknown sonification map: {map_name}")
    return pitches


def composition_features(seq: str) -> np.ndarray:
    n = max(len(seq), 1)
    base = [seq.count(b) / n for b in BASES]
    codons = np.zeros(64, dtype=float)
    total = 0
    for i in range(0, len(seq) - 2, 3):
        codon = seq[i : i + 3]
        if codon in CODON_INDEX:
            codons[CODON_INDEX[codon]] += 1
            total += 1
    codons = codons / max(total, 1)
    gc = (seq.count("G") + seq.count("C")) / n
    return np.array(base + [gc] + codons.tolist(), dtype=float)


def audio_features(seq: str, map_name: str = "codon64") -> np.ndarray:
    pitches = codon_pitches(seq, map_name)
    if not pitches:
        bins = 64 if map_name == "codon64" else len(CLASS_ORDER)
        return np.zeros(8 + bins, dtype=float)
    pitch_arr = np.array(pitches, dtype=float)
    intervals = np.diff(pitch_arr) if len(pitch_arr) > 1 else np.zeros(1)
    if map_name == "codon64":
        hist = np.bincount(np.array(pitches) - 36, minlength=64).astype(float)
    else:
        class_to_idx = {name: i for i, name in enumerate(CLASS_ORDER)}
        idxs = []
        for i in range(0, len(seq) - 2, 3):
            codon = seq[i : i + 3]
            if codon in CODON_TO_AA:
                idxs.append(class_to_idx[AA_CLASS[CODON_TO_AA[codon]]])
        hist = np.bincount(np.array(idxs), minlength=len(CLASS_ORDER)).astype(float)
    hist = hist / max(hist.sum(), 1)
    scalar = [
        float(pitch_arr.mean()),
        float(pitch_arr.std()),
        float(np.percentile(pitch_arr, 10)),
        float(np.percentile(pitch_arr, 90)),
        shannon(pitches),
        float(np.mean(np.abs(intervals))),
        float(np.std(intervals)),
        float(np.mean(np.abs(np.diff(intervals)))) if len(intervals) > 1 else 0.0,
    ]
    return np.array(scalar + hist.tolist(), dtype=float)


def mutate_substitutions(seq: str, rate: float, rng: random.Random) -> str:
    chars = list(seq)
    k = max(1, int(len(chars) * rate))
    for idx in rng.sample(range(len(chars)), k):
        choices = [base for base in BASES if base != chars[idx]]
        chars[idx] = rng.choice(choices)
    return "".join(chars)


def replace_middle(seq: str, replacement: str, size: int | None = None) -> str:
    size = min(size or len(seq) // 2, len(seq), len(replacement))
    start = (len(seq) - size) // 2
    return seq[:start] + replacement[:size] + seq[start + size :]


def euclidean(a: np.ndarray, b: np.ndarray) -> float:
    return float(np.linalg.norm(a - b))


def auc_from_scores(labels: list[int], scores: list[float]) -> float:
    positives = [(score, label) for score, label in zip(scores, labels) if label == 1]
    negatives = [(score, label) for score, label in zip(scores, labels) if label == 0]
    if not positives or not negatives:
        return 0.0
    wins = 0.0
    for ps, _ in positives:
        for ns, _ in negatives:
            if ps > ns:
                wins += 1
            elif ps == ns:
                wins += 0.5
    return wins / (len(positives) * len(negatives))


def precision_at_k(labels: list[int], scores: list[float], k: int) -> float:
    order = np.argsort(np.array(scores))[::-1][:k]
    return float(sum(labels[i] for i in order) / max(k, 1))


def bootstrap_ci(values: list[float], rng: random.Random, rounds: int = 500) -> tuple[float, float]:
    if not values:
        return (0.0, 0.0)
    samples = []
    for _ in range(rounds):
        draw = [values[rng.randrange(len(values))] for _ in values]
        samples.append(float(np.mean(draw)))
    return (float(np.percentile(samples, 2.5)), float(np.percentile(samples, 97.5)))


def write_var_len(value: int) -> bytes:
    buffer = value & 0x7F
    out = []
    while value := value >> 7:
        buffer <<= 8
        buffer |= ((value & 0x7F) | 0x80)
    while True:
        out.append(buffer & 0xFF)
        if buffer & 0x80:
            buffer >>= 8
        else:
            break
    return bytes(out)


def write_midi_from_scores(scores: list[float], path: pathlib.Path, title: str = "genome anomaly track") -> None:
    if not scores:
        return
    lo = min(scores)
    hi = max(scores)
    spread = hi - lo if hi > lo else 1.0
    track = bytearray()
    track.extend(b"\x00\xff\x03" + bytes([len(title)]) + title.encode("ascii", "ignore")[:120])
    track.extend(b"\x00\xff\x51\x03\x07\xa1\x20")
    for score in scores[:256]:
        norm = (score - lo) / spread
        pitch = int(48 + round(norm * 24))
        velocity = int(40 + round(norm * 55))
        duration = 90 if norm < 0.75 else 180
        track.extend(write_var_len(0))
        track.extend(bytes([0x90, pitch, velocity]))
        track.extend(write_var_len(duration))
        track.extend(bytes([0x80, pitch, 0]))
    track.extend(b"\x00\xff\x2f\x00")
    data = b"MThd" + struct.pack(">IHHH", 6, 0, 1, 480)
    data += b"MTrk" + struct.pack(">I", len(track)) + bytes(track)
    path.write_bytes(data)
