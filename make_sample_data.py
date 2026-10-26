"""Generate synthetic endurance-test data (generic demo; no real product data)."""
from pathlib import Path
import numpy as np
import pandas as pd

rng = np.random.default_rng(42)
N = 500
cycles = np.arange(1, N + 1)
t0 = pd.Timestamp("2026-03-02 08:00")
df = pd.DataFrame({
    "cycle": cycles,
    "timestamp": [t0 + pd.Timedelta(seconds=float(s)) for s in np.cumsum(rng.normal(42, 1.5, N))],
    "peak_pressure_bar": rng.normal(4.0, 0.06, N),
    "cycle_time_s": rng.normal(12.0, 0.25, N),
    "current_A": rng.normal(2.1, 0.04, N),
    "temperature_C": 23 + np.cumsum(rng.normal(0.002, 0.03, N)),
})
# Injected faults
df.loc[[87, 233, 412], "peak_pressure_bar"] += [0.55, -0.6, 0.7]   # pressure spikes / drops
df.loc[300:, "current_A"] += np.linspace(0, 0.24, N - 300)          # slow current drift (wear)
df.loc[[150, 151], "cycle_time_s"] += [3.1, 2.8]                    # two slow cycles
df.loc[320, "peak_pressure_bar"] += 0.31                           # spike still inside limits
df = df.drop(index=[198, 199, 200])                                 # missing cycles (logger gap)
df["timestamp"] = pd.to_datetime(df["timestamp"]).dt.floor("s")
df = df.round({"peak_pressure_bar": 3, "cycle_time_s": 2, "current_A": 3, "temperature_C": 2})
out = Path(__file__).parent / "sample_data"
out.mkdir(exist_ok=True)
df.to_csv(out / "endurance_test_log.csv", index=False)
print(f"{len(df)} cycles written to {out / 'endurance_test_log.csv'}")
