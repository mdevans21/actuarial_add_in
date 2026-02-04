"""Compare ELT aggregation approaches using the eltr worked example data.

Data source:
  https://cran.r-project.org/web/packages/eltr/vignettes/Worked-Example.html

Instructions:
  1) Paste the ELT table into ELT_EVENTS below or into
     tests/data/eltr_worked_example.csv (id, rate, mean, sdevi, sdevc, exp).
  2) Run: python tests/eltr_worked_example.py

This script focuses on the variance combination question and provides
lognormal (capped at TIV) and beta (scaled to M) severity simulators.
"""

from __future__ import annotations

from dataclasses import dataclass
import csv
import math
import random
from pathlib import Path
from typing import Iterable


@dataclass(frozen=True)
class Event:
    name: str
    rate: float
    mean: float
    sd: float
    tiv: float


# TODO: Replace with the ELT from the eltr worked example.
# Expected columns: event name/id, annual rate (lambda), mean loss, loss SD, TIV.
# If the CSV provides sdevi/sdevc/exp, SD is computed via SRSS and TIV = exp.
ELT_EVENTS: list[Event] = [
    # Example format:
    # Event("Event 1", 0.2, 1_000_000, 500_000, 5_000_000),
]

DEFAULT_CSV_PATH = Path(__file__).resolve().parent / "data" / "eltr_worked_example.csv"


def validate_events(events: list[Event]) -> None:
    if not events:
        raise ValueError(
            "ELT_EVENTS is empty. Populate it with the data from the eltr worked example."
        )
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


def compound_poisson_stats(events: list[Event]) -> tuple[float, float]:
    """Mean/SD for aggregate annual loss per compound Poisson across events."""
    mean = sum(event.rate * event.mean for event in events)
    variance = sum(event.rate * (event.sd ** 2 + event.mean ** 2) for event in events)
    return mean, math.sqrt(variance)


def scaled_sd_stats(events: list[Event]) -> tuple[float, float]:
    """Alternative: scale the severity SD only (ignores mean^2 term)."""
    mean = sum(event.rate * event.mean for event in events)
    variance = sum(event.rate * (event.sd ** 2) for event in events)
    return mean, math.sqrt(variance)


def lognormal_params_from_mean_sd(mean: float, sd: float) -> tuple[float, float]:
    variance = sd ** 2
    phi = math.sqrt(variance + mean ** 2)
    sigma = math.sqrt(math.log((phi ** 2) / (mean ** 2)))
    mu = math.log(mean) - 0.5 * sigma ** 2
    return mu, sigma


def sample_lognormal_capped(rng: random.Random, mean: float, sd: float, tiv: float) -> float:
    mu, sigma = lognormal_params_from_mean_sd(mean, sd)
    sample = rng.lognormvariate(mu, sigma)
    return min(sample, tiv)


def beta_params_for_mean_sd(mean: float, sd: float, max_loss: float) -> tuple[float, float]:
    if max_loss <= 0:
        raise ValueError("max_loss must be positive")
    p = mean / max_loss
    if p <= 0 or p >= 1:
        raise ValueError("mean must be between 0 and max_loss")
    variance = sd ** 2
    max_variance = mean * (max_loss - mean)
    if variance >= max_variance:
        raise ValueError("variance too large for given max_loss")
    k = max_variance / variance - 1.0
    alpha = p * k
    beta = (1.0 - p) * k
    return alpha, beta


def adjust_max_loss_for_beta(mean: float, sd: float, max_loss: float) -> float:
    """Ensure max_loss is large enough for the requested mean/sd."""
    if mean <= 0:
        return max_loss
    min_m = mean + (sd ** 2) / mean
    return max(max_loss, min_m)


def sample_beta_scaled(rng: random.Random, mean: float, sd: float, max_loss: float) -> float:
    adjusted_max = adjust_max_loss_for_beta(mean, sd, max_loss)
    alpha, beta = beta_params_for_mean_sd(mean, sd, adjusted_max)
    return rng.betavariate(alpha, beta) * adjusted_max


def simulate_ylt(
    events: Iterable[Event],
    years: int,
    seed: int,
    severity: str,
) -> list[dict[str, float]]:
    rng = random.Random(seed)
    results: list[dict[str, float]] = []
    for year in range(1, years + 1):
        aggregate = 0.0
        maximum = 0.0
        count = 0
        for event in events:
            occurrences = rng.poisson(event.rate) if hasattr(rng, "poisson") else poisson(rng, event.rate)
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


def main() -> None:
    events = ELT_EVENTS or load_events_from_csv(DEFAULT_CSV_PATH)
    if not events:
        raise ValueError(
            "No ELT data found. Populate ELT_EVENTS or fill tests/data/eltr_worked_example.csv."
        )
    validate_events(events)

    cp_mean, cp_sd = compound_poisson_stats(events)
    scaled_mean, scaled_sd = scaled_sd_stats(events)

    print("ELT Worked Example Summary")
    print("---------------------------")
    print(f"Events: {len(events)}")
    print(f"Compound Poisson mean: {cp_mean:,.2f}")
    print(f"Compound Poisson SD:   {cp_sd:,.2f}")
    print()
    print("Alternative SD treatments (same mean):")
    print(f"Scaled SD only:        {scaled_sd:,.2f}")
    print()
    print("Notes:")
    print("- Compound Poisson variance uses rate * (sd^2 + mean^2) per event.")
    print("- The alternative omits the mean^2 term and can understate volatility.")


if __name__ == "__main__":
    main()
