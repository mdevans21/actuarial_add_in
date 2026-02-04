"""Reconcile Python formulas with the eltr worked example outputs.

This script uses the raw ELT input from tests/data/eltr_worked_example.csv,
computes create_elt outputs, and validates them against the worked example
values provided in the vignette text. It also reproduces the annual loss
summary and OEP curve using the worked example YLT output.
"""

from __future__ import annotations

from dataclasses import dataclass
import csv
import math
from pathlib import Path
from typing import Iterable

import numpy as np

DATA_PATH = Path(__file__).resolve().parent / "data" / "eltr_worked_example.csv"


@dataclass(frozen=True)
class RawEvent:
    event_id: int
    rate: float
    mean: float
    sdevi: float
    sdevc: float
    exp: float


@dataclass(frozen=True)
class DerivedEvent:
    event_id: int
    rate: float
    mean: float
    sdevi: float
    sdevc: float
    exp: float
    mdr: float
    sdev: float
    cov: float
    alpha: float
    beta: float
    random_num: float


EXPECTED_ELT = [
    # id, rate, mean, sdevi, sdevc, exp, mdr, random_num
    (1, 0.10, 500, 500, 200, 100000, 0.0050000, 0.081967213),
    (2, 0.10, 200, 400, 100, 5000, 0.0400000, 0.081967213),
    (3, 0.20, 300, 200, 400, 40000, 0.0075000, 0.163934426),
    (4, 0.10, 100, 300, 500, 4000, 0.0250000, 0.081967213),
    (5, 0.20, 500, 100, 200, 2000, 0.2500000, 0.163934426),
    (6, 0.25, 200, 200, 500, 50000, 0.0040000, 0.204918033),
    (7, 0.01, 1000, 500, 600, 100000, 0.0100000, 0.008196721),
    (8, 0.12, 250, 300, 100, 5000, 0.0500000, 0.098360656),
    (9, 0.14, 1000, 500, 200, 6000, 0.1666667, 0.114754098),
    (10, 0.00, 10000, 1000, 500, 1000000, 0.0100000, 0.000000000),
]

EXPECTED_YLT = [
    (1, 0.000000e00, None),
    (2, 8.574328e01, 5),
    (3, 9.137924e-01, 2),
    (4, 2.611786e02, 1),
    (5, 2.686697e00, 8),
    (6, 2.529234e02, 1),
    (6, 9.173005e00, 3),
    (7, 0.000000e00, None),
    (8, 3.633260e-07, 6),
    (9, 1.286863e02, 3),
    (9, 2.296461e02, 6),
    (10, 0.000000e00, None),
]

EXPECTED_OEP = [
    (10000, 358.245795),
    (5000, 358.159183),
    (1000, 357.466284),
    (500, 356.600160),
    (250, 354.867913),
    (200, 354.001789),
    (100, 349.671171),
    (50, 341.009935),
    (25, 323.687463),
    (10, 271.720046),
    (5, 261.362178),
    (2, 1.800245),
]


def load_raw_events(path: Path) -> list[RawEvent]:
    events: list[RawEvent] = []
    with path.open(newline="") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            if not row:
                continue
            events.append(
                RawEvent(
                    event_id=int(row["id"]),
                    rate=float(row["rate"]),
                    mean=float(row["mean"]),
                    sdevi=float(row["sdevi"]),
                    sdevc=float(row["sdevc"]),
                    exp=float(row["exp"]),
                )
            )
    return events


def create_elt(events: Iterable[RawEvent]) -> list[DerivedEvent]:
    events = list(events)
    total_rate = sum(event.rate for event in events)
    derived: list[DerivedEvent] = []
    for event in events:
        sdev = math.sqrt(event.sdevi ** 2 + event.sdevc ** 2)
        mdr = event.mean / event.exp
        cov = sdev / event.mean
        variance = (sdev / event.exp) ** 2
        k = mdr * (1 - mdr) / variance - 1
        alpha = mdr * k
        beta = (1 - mdr) * k
        random_num = event.rate / total_rate if total_rate > 0 else 0.0
        derived.append(
            DerivedEvent(
                event_id=event.event_id,
                rate=event.rate,
                mean=event.mean,
                sdevi=event.sdevi,
                sdevc=event.sdevc,
                exp=event.exp,
                mdr=mdr,
                sdev=sdev,
                cov=cov,
                alpha=alpha,
                beta=beta,
                random_num=random_num,
            )
        )
    return derived


def summarize_ylt(ylt_rows: list[tuple[int, float, int | None]]) -> list[tuple[int, float]]:
    summary: dict[int, float] = {}
    for year, loss, _event in ylt_rows:
        summary[year] = summary.get(year, 0.0) + loss
    return sorted(summary.items())


def oep_curve(losses: list[float], return_periods: list[float]) -> list[tuple[float, float]]:
    values = sorted(losses)
    curve: list[tuple[float, float]] = []
    for rp in return_periods:
        ep = 1.0 / rp
        quantile = float(np.quantile(values, 1 - ep))
        curve.append((rp, quantile))
    return curve


def main() -> None:
    raw_events = load_raw_events(DATA_PATH)
    derived = create_elt(raw_events)

    print("ELT derived columns (tolerance 1e-6):")
    for expected, actual in zip(EXPECTED_ELT, derived):
        _, rate, mean, sdevi, sdevc, exp, mdr, random_num = expected
        assert math.isclose(actual.rate, rate, rel_tol=0, abs_tol=1e-12)
        assert math.isclose(actual.mean, mean, rel_tol=0, abs_tol=1e-12)
        assert math.isclose(actual.sdevi, sdevi, rel_tol=0, abs_tol=1e-12)
        assert math.isclose(actual.sdevc, sdevc, rel_tol=0, abs_tol=1e-12)
        assert math.isclose(actual.exp, exp, rel_tol=0, abs_tol=1e-12)
        assert math.isclose(actual.mdr, mdr, rel_tol=0, abs_tol=1e-6)
        assert math.isclose(actual.random_num, random_num, rel_tol=0, abs_tol=1e-9)
    print("- ELT derived columns match expected output.")

    ann = summarize_ylt(EXPECTED_YLT)
    ann_losses = [loss for _year, loss in ann]
    avg_loss = sum(ann_losses) / len(ann_losses)

    print("\nAnnual loss summary (from worked example YLT):")
    for year, loss in ann:
        print(f"Year {year}: {loss:.6f}")
    print(f"Expected annual loss: {avg_loss:.5f}")

    return_periods = [rp for rp, _ in EXPECTED_OEP]
    aep_curve = oep_curve(ann_losses, return_periods)
    max_losses = [max(loss for year, loss, _event in EXPECTED_YLT if year == target_year)
                  for target_year, _loss in ann]
    oep_curve_values = oep_curve(max_losses, return_periods)

    print("\nAEP curve:")
    for (rp, expected_loss), (_, loss) in zip(EXPECTED_OEP, aep_curve):
        if not math.isclose(loss, expected_loss, rel_tol=0, abs_tol=1e-3):
            raise AssertionError(f"AEP mismatch for RP {rp}: {loss} vs {expected_loss}")
        print(f"RP {rp}: {loss:.6f}")
    print("- AEP curve matches expected output.")

    print("\nOEP curve (from annual max losses):")
    for (rp, _expected_loss), (_, loss) in zip(EXPECTED_OEP, oep_curve_values):
        print(f"RP {rp}: {loss:.6f}")


if __name__ == "__main__":
    main()
