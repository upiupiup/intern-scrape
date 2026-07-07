"""
forecast_pertumbuhan_ekonomi.py

Script forecast pertumbuhan ekonomi Indonesia (YoY, %) memakai ARIMA(1,1,1).

Cara jalanin:
    python forecast_pertumbuhan_ekonomi.py
"""

import os
import warnings

warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
from statsmodels.tsa.arima.model import ARIMA

DATA_PATH = (
    "data/raw/groupA/usecase_ekonomi.ekonomi_pertumbuhan_ekonomi_kuartal_2010_2026.csv"
)
OUTPUT_PATH = (
    "data/processed/groupA/usecase_ekonomi.ekonomi_pertumbuhan_ekonomi_kuartal.csv"
)

ARIMA_ORDER = (1, 1, 1)
FORECAST_STEPS = 3
CI_ALPHA = 0.20
OUTPUT_FROM_YEAR = 2024

FORECAST_SOURCE_LABEL = "Forecast internal ARIMA(1,1,1) (statsmodels, Python)"


# Pipeline
def load_data(path: str) -> pd.DataFrame:
    df = pd.read_csv(path)
    df = df.sort_values(["tahun", "kuartal"]).reset_index(drop=True)
    df["period"] = pd.PeriodIndex.from_fields(
        year=df["tahun"], quarter=df["kuartal"], freq="Q"
    )
    return df


def fit_and_forecast(df: pd.DataFrame, order, steps, alpha):
    series = df.set_index("period")["growth_pct"].astype(float)

    model = ARIMA(series, order=order).fit()
    fc_res = model.get_forecast(steps=steps)
    fc_mean = fc_res.predicted_mean
    fc_ci = fc_res.conf_int(alpha=alpha)

    future_periods = pd.period_range(df["period"].iloc[-1] + 1, periods=steps, freq="Q")

    forecast_df = pd.DataFrame(
        {
            "tahun": [p.year for p in future_periods],
            "kuartal": [p.quarter for p in future_periods],
            "forecast": fc_mean.values.round(2),
            "forecast_lower": fc_ci.iloc[:, 0].values.round(2),
            "forecast_upper": fc_ci.iloc[:, 1].values.round(2),
        }
    )
    return forecast_df, model


def build_output(
    df: pd.DataFrame, forecast_df: pd.DataFrame, from_year: int, forecast_source: str
) -> pd.DataFrame:
    hist = df.loc[
        df["tahun"] >= from_year, ["tahun", "kuartal", "growth_pct", "data_source"]
    ].copy()
    hist["forecast"] = np.nan
    hist["forecast_lower"] = np.nan
    hist["forecast_upper"] = np.nan

    fc = forecast_df.copy()
    fc["growth_pct"] = np.nan
    fc["data_source"] = forecast_source

    out = pd.concat([hist, fc], ignore_index=True, sort=False)
    out = out.sort_values(["tahun", "kuartal"]).reset_index(drop=True)
    out.insert(0, "id", range(1, len(out) + 1))
    out = out[
        [
            "id",
            "tahun",
            "kuartal",
            "growth_pct",
            "forecast",
            "forecast_lower",
            "forecast_upper",
            "data_source",
        ]
    ]
    return out


def main():
    df = load_data(DATA_PATH)

    forecast_df, model = fit_and_forecast(
        df, order=ARIMA_ORDER, steps=FORECAST_STEPS, alpha=CI_ALPHA
    )

    out = build_output(df, forecast_df, OUTPUT_FROM_YEAR, FORECAST_SOURCE_LABEL)

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    out.to_csv(OUTPUT_PATH, index=False)

    print(out.to_string(index=False))
    print(f"\nTersimpan di: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
