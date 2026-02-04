"""Simulate ELT -> YLT -> OEP curve for the eltr worked example.

Data source:
  https://cran.r-project.org/web/packages/eltr/vignettes/Worked-Example.html

Instructions:
  1) Paste the ELT table into ELT_EVENTS below or into
     tests/data/eltr_worked_example.csv (id, rate, mean, sdevi, sdevc, exp).
  2) Run: python tests/elt_to_ylt_oep.py

Outputs:
  - OEP curve (return period, loss)
  - Summary statistics (mean, SD, VaR, TVaR)
"""

from __future__ import annotations

from dataclasses import dataclass
import csv
import math
import random
from pathlib import Path
from typing import Iterable

import numpy as np


@dataclass(frozen=True)
class Event:
    name: str
    rate: float
    mean: float
    sd: float
    tiv: float


# TODO: Replace with ELT data from the worked example.
ELT_EVENTS: list[Event] = [
    # Event("Event 1", 0.2, 1_000_000, 500_000, 5_000_000),
]

DEFAULT_CSV_PATH = Path(__file__).resolve().parent / "data" / "eltr_worked_example.csv"


def validate_events(events: list[Event]) -> None:
    if not events:
        raise ValueError("ELT_EVENTS is empty. Populate it with the worked example data.")
    for event in events:
        if event.rate < 0 or event.mean < 0 or event.sd < 0 or event.tiv <= 0:
            raise ValueError(f"Invalid event values: {event}")
        if event.mean > event.tiv:
            raise ValueError(f"Mean exceeds TIV for {event.name}")


def load_events_from_csv(path: Path) -> list[Event]:
    if not path.exists():
        return []
    events: list[Event] = []
    with path.open(newline="") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            if not row:
                continue
            if "sdevi" in row and "sdevc" in row and "exp" in row:
                sdev = math.sqrt(float(row["sdevi"]) ** 2 + float(row["sdevc"]) ** 2)
                tiv = float(row["exp"])
            else:
                sdev = float(row["sd"])
                tiv = float(row["tiv"])
            events.append(
                Event(
                    name=row.get("name") or row.get("id") or "event",
                    rate=float(row["rate"]),
                    mean=float(row["mean"]),
                    sd=sdev,
                    tiv=tiv,
                )
            )
    return events


def lognormal_params_from_mean_sd(mean: float, sd: float) -> tuple[float, float]:
    variance = sd ** 2
    phi = math.sqrt(variance + mean ** 2)
    sigma = math.sqrt(math.log((phi ** 2) / (mean ** 2)))
    mu = math.log(mean) - 0.5 * sigma ** 2
    return mu, sigma


def sample_lognormal_capped(rng: random.Random, mean: float, sd: float, tiv: float) -> float:
    mu, sigma = lognormal_params_from_mean_sd(mean, sd)
    return min(rng.lognormvariate(mu, sigma), tiv)


def adjust_max_loss_for_beta(mean: float, sd: float, max_loss: float) -> float:
    if mean <= 0:
        return max_loss
    min_m = mean + (sd ** 2) / mean
    return max(max_loss, min_m)


def beta_params_for_mean_sd(mean: float, sd: float, max_loss: float) -> tuple[float, float]:
    p = mean / max_loss
    variance = sd ** 2
    max_variance = mean * (max_loss - mean)
    if variance >= max_variance:
        raise ValueError("variance too large for given max_loss")
    k = max_variance / variance - 1.0
    alpha = p * k
    beta = (1.0 - p) * k
    return alpha, beta


def sample_beta_scaled(rng: random.Random, mean: float, sd: float, max_loss: float) -> float:
    adjusted_max = adjust_max_loss_for_beta(mean, sd, max_loss)
    alpha, beta = beta_params_for_mean_sd(mean, sd, adjusted_max)
    return rng.betavariate(alpha, beta) * adjusted_max


def poisson(rng: random.Random, rate: float) -> int:
    if rate <= 0:
        return 0
    l = math.exp(-rate)
    k = 0
    p = 1.0
    while p > l:
        k += 1
        p *= rng.random()
    return k - 1


def simulate_ylt(events: Iterable[Event], years: int, seed: int, severity: str) -> list[dict[str, float]]:
    rng = random.Random(seed)
    results: list[dict[str, float]] = []
    for year in range(1, years + 1):
        aggregate = 0.0
        maximum = 0.0
        count = 0
        for event in events:
            occurrences = poisson(rng, event.rate)
            if occurrences <= 0:
                continue
            for _ in range(occurrences):
                if severity == "lognormal":
                    loss = sample_lognormal_capped(rng, event.mean, event.sd, event.tiv)
                elif severity == "beta":
                    loss = sample_beta_scaled(rng, event.mean, event.sd, event.tiv)
                else:
                    raise ValueError("severity must be 'lognormal' or 'beta'")
                aggregate += loss
                maximum = max(maximum, loss)
                count += 1
        results.append({
            "year": float(year),
            "aggregate": aggregate,
            "maximum": maximum,
            "count": float(count),
        })
    return results


def ep_curve(losses: list[float], return_periods: list[float]) -> list[tuple[float, float]]:
    values = sorted(losses)
    curve: list[tuple[float, float]] = []
    for rp in return_periods:
        ep = 1.0 / rp
        quantile = float(np.quantile(values, 1 - ep))
        curve.append((rp, quantile))
    return curve


def var_from_samples(samples: list[float], alpha: float) -> float:
    if not samples:
        raise ValueError("samples required")
    if not (0 < alpha < 1):
        raise ValueError("alpha must be between 0 and 1")
    sorted_samples = sorted(samples)
    index = int(math.ceil(alpha * len(sorted_samples))) - 1
    return sorted_samples[index]


def tvar_from_samples(samples: list[float], alpha: float) -> float:
    var_value = var_from_samples(samples, alpha)
    tail = [x for x in samples if x >= var_value]
    if not tail:
        return var_value
    return sum(tail) / len(tail)


def summarize(samples: list[float]) -> tuple[float, float]:
    mean = sum(samples) / len(samples)
    variance = sum((x - mean) ** 2 for x in samples) / (len(samples) - 1)
    return mean, math.sqrt(variance)


def main() -> None:
    events = ELT_EVENTS or load_events_from_csv(DEFAULT_CSV_PATH)
    if not events:
        raise ValueError(
            "No ELT data found. Populate ELT_EVENTS or fill tests/data/eltr_worked_example.csv."
        )
    validate_events(events)

    years = 10000
    seed = 42
    severity = "lognormal"

    ylt = simulate_ylt(events, years=years, seed=seed, severity=severity)
    max_losses = [row["maximum"] for row in ylt]
    agg_losses = [row["aggregate"] for row in ylt]

    mean, sd = summarize(agg_losses)
    var_99 = var_from_samples(agg_losses, 0.99)
    tvar_99 = tvar_from_samples(agg_losses, 0.99)

    print(f"Simulated years: {years}")
    print(f"Severity: {severity}")
    print(f"Aggregate mean: {mean:,.2f}")
    print(f"Aggregate SD:   {sd:,.2f}")
    print(f"VaR 99%:        {var_99:,.2f}")
    print(f"TVaR 99%:       {tvar_99:,.2f}")

    return_periods = [10000, 5000, 1000, 500, 250, 200, 100, 50, 25, 10, 5, 2]
    oep = ep_curve(max_losses, return_periods)
    aep = ep_curve(agg_losses, return_periods)

    print("\nOEP/AEP Curves")
    print("Return Period\tOEP Loss\tAEP Loss")
    for (rp, oep_loss), (_, aep_loss) in zip(oep, aep):
        print(f"{rp:,.2f}\t{oep_loss:,.2f}\t{aep_loss:,.2f}")


if __name__ == "__main__":
    main()
