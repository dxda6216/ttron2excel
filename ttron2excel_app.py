"""Taylortron TRACES -> Excel converter (stand-alone desktop application).

Single-file GUI version of the Colab notebook "ttron_data_file_to_excel_file".
Reads a Taylortron TRACES.nnn file and writes Excel workbooks, a ZIP of
per-channel .dat files (LumiCycle) and a multi-page PDF of plots.

Run with:   python ttron2excel_app.py

Layout of this file:
  1. Processing core  (functions taken unchanged from the notebook, plus run_analysis)
  2. Desktop GUI      (Tkinter; the processing runs in a background thread)
"""

from __future__ import annotations

import contextlib
import datetime as dt
import json
import math
import os
import queue
import re
import subprocess
import sys
import tempfile
import threading
import traceback
import tkinter as tk
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import matplotlib

matplotlib.use("Agg")  # headless backend: figures go straight to the PDF
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
from matplotlib.backends.backend_pdf import PdfPages  # noqa: E402
from matplotlib.ticker import FormatStrFormatter, MultipleLocator  # noqa: E402
from scipy import signal  # noqa: E402
from scipy.optimize import curve_fit  # noqa: E402

APP_TITLE = "Taylortron TRACES \u2192 Excel Converter"


# =========================================================================== #
# 1. Processing core
# =========================================================================== #

# --------------------------------------------------------------------------- #
# Constants
# --------------------------------------------------------------------------- #

N_CHANNELS = 30
CHANNELS = [f"{k:02d}" for k in range(N_CHANNELS)]
COLUMN_NAMES = ["Hours"] + CHANNELS

HOURS_PER_DAY = 24.0
PEAK_MIN_SEPARATION_HOURS = 12.0   # minimum spacing between detected peaks
SMOOTHING_WINDOWS = (5, 9)         # moving-average windows reported in the workbook
FIT_MIN_PERIOD_HOURS = 12.0        # bounds for the damped sine period
FIT_MAX_PERIOD_HOURS = 60.0
FIT_MIN_POINTS = 5                 # one point per fitted parameter
EXCEL_SHEET_NAME_LIMIT = 31        # hard limit imposed by the .xlsx format


# --------------------------------------------------------------------------- #
# Small utilities
# --------------------------------------------------------------------------- #

def sheet_name(name: str) -> str:
    """Return a sheet name Excel will accept (<= 31 chars, no []:*?/\\)."""
    for bad in "[]:*?/\\":
        name = name.replace(bad, "-")
    return name[:EXCEL_SHEET_NAME_LIMIT]


def upper_ylimit(series: pd.Series, headroom: float = 1.10) -> float:
    """Top of the y-axis for a channel, safe for empty or all-NaN channels."""
    if not series.notna().any():
        return 1.0
    return max(float(np.nanmax(series.to_numpy())) * headroom, 1.0)


def utc_now() -> datetime:
    return datetime.now(timezone.utc)


def utc_now_string() -> str:
    return utc_now().strftime("%Y-%m-%d %H:%M:%S")


# --------------------------------------------------------------------------- #
# Reading and describing the data
# --------------------------------------------------------------------------- #

@dataclass
class Recording:
    """A TRACES file and the timing facts derived from it."""

    filename: str
    data: pd.DataFrame

    @property
    def n_points(self) -> int:
        return len(self.data.index)

    @property
    def first_hour(self) -> float:
        return float(self.data["Hours"].iloc[0])

    @property
    def last_hour(self) -> float:
        return float(self.data["Hours"].iloc[-1])

    @property
    def duration_hours(self) -> float:
        return self.last_hour - self.first_hour

    @property
    def time_interval(self) -> float:
        """Average spacing between samples, in hours."""
        return self.duration_hours / (self.n_points - 1)

    def describe(self) -> None:
        print(f"Number of rows: {self.n_points}")
        print(f"First time point: {self.first_hour} h")
        print(f"Last time point: {self.last_hour} h")
        print(
            f"Total time duration: {self.duration_hours} h "
            f"= {self.duration_hours / HOURS_PER_DAY} days"
        )
        print(f"Average time interval: {self.time_interval} h")


def read_traces(path: str) -> Recording:
    """Read a tab-separated TRACES.nnn file (3 header lines, 1 footer line)."""
    data = pd.read_csv(
        path,
        header=None,
        sep="\t",
        skiprows=3,
        skipfooter=1,
        index_col=False,
        names=COLUMN_NAMES,
        engine="python",
    )
    return Recording(filename=path, data=data.loc[:, COLUMN_NAMES])


def rolling_mean(data: pd.DataFrame, window: int) -> pd.DataFrame:
    """Centered moving average of the channel columns; 'Hours' is left alone.

    (Smoothing 'Hours' as well would shift the time axis at the edges, where
    min_periods=1 averages over an incomplete window.)
    """
    smoothed = data[CHANNELS].rolling(window=window, center=True, min_periods=1).mean()
    smoothed.insert(0, "Hours", data["Hours"])
    return smoothed


# --------------------------------------------------------------------------- #
# Detrending
# --------------------------------------------------------------------------- #

@dataclass
class Trend:
    """A trend line plus a human-readable description of how it was made."""

    data: pd.DataFrame
    label: str          # shown in plot legends
    short_label: str    # used in the Excel sheet name (31-character limit)
    summary: str        # shown on the 'Note' sheet


def moving_average_trend(
    data: pd.DataFrame, window_hours: float, time_interval: float
) -> Trend:
    n_points = math.ceil(window_hours / time_interval)
    if n_points % 2 == 0:
        n_points += 1
    window_span = time_interval * (n_points - 1)
    print(
        f"Window size for trend line (Moving Average): "
        f"{window_span:.6f} h ({n_points} points)"
    )

    trend = rolling_mean(data, n_points).bfill().ffill()
    return Trend(
        data=trend,
        label=f"{n_points}-point moving average, centered",
        short_label=f"{n_points}PMA",
        summary=f"{window_hours}h window ({n_points} points)",
    )


def sinc_filter_trend(
    data: pd.DataFrame, cutoff_period_hours: float, order: int, time_interval: float
) -> Trend:
    n_points = len(data.index)
    sampling_rate = 1.0 / time_interval                 # samples per hour
    cutoff_frequency = 1.0 / cutoff_period_hours        # cycles per hour
    nyquist_frequency = 0.5 * sampling_rate
    norm_cutoff = cutoff_frequency / nyquist_frequency

    # filtfilt pads the signal by 3 * len(taps), so the order has to stay well
    # below the record length; it also has to be odd for a zero-phase lowpass.
    max_order = int(n_points / 3) - 1
    if order >= max_order:
        print(
            f"Warning: Sinc Filter Order ({order}) is too high for data length "
            f"({n_points} points)."
        )
        order = max(1, max_order)
        print(f"Adjusting Sinc Filter Order to {order} for stability with filtfilt.")
    if order % 2 == 0:
        order += 1

    print(f"Sampling Rate: {sampling_rate:.2f} samples/hour")
    print(f"Cutoff Period for Sinc Filter: {cutoff_period_hours} hours")
    print(f"Cutoff Frequency: {cutoff_frequency:.4f} cycles/hour")
    print(f"Normalized Cutoff Frequency: {norm_cutoff:.4f}")
    print(f"Sinc Filter Order: {order}")

    taps = signal.firwin(order, norm_cutoff, pass_zero="lowpass")

    trend = data.copy()
    for channel in CHANNELS:
        if not data[channel].notna().any():
            trend[channel] = np.nan
            continue

        # filtfilt cannot handle NaNs; interpolate across gaps first.
        filled = data[channel].interpolate(method="linear", limit_direction="both")
        try:
            trend[channel] = signal.filtfilt(taps, [1.0], filled)
        except ValueError as error:
            print(f"Warning: could not filter channel {channel}: {error}")
            print("Consider reducing the filter order or using a longer recording.")
            trend[channel] = np.nan

    return Trend(
        data=trend,
        label=f"Sinc Filter C={cutoff_period_hours}h O={order}",
        short_label=f"Sinc C={cutoff_period_hours}h O={order}",
        summary=f"{cutoff_period_hours}h cutoff, order {order}",
    )


def build_trend(
    data: pd.DataFrame,
    method: str,
    time_interval: float,
    *,
    window_hours: float,
    cutoff_period_hours: float,
    order: int,
) -> Trend:
    if method == "Moving Average":
        return moving_average_trend(data, window_hours, time_interval)
    if method == "Sinc Filter":
        return sinc_filter_trend(data, cutoff_period_hours, order, time_interval)
    raise ValueError(f"Unknown detrending method: {method!r}")


def detrend(data: pd.DataFrame, trend: pd.DataFrame) -> pd.DataFrame:
    """Subtract the trend from the channels, keeping the original time axis."""
    detrended = data[CHANNELS] - trend[CHANNELS]
    detrended.insert(0, "Hours", data["Hours"])
    return detrended


# --------------------------------------------------------------------------- #
# Peaks and troughs
# --------------------------------------------------------------------------- #

def find_peaks_and_troughs(
    series: pd.Series,
    min_separation_hours: float = PEAK_MIN_SEPARATION_HOURS,
    time_interval: float | None = None,
):
    """Return the series indices of the peaks and of the troughs."""
    valid = series.dropna()
    if valid.empty:
        empty = pd.Index([])
        return empty, empty

    if min_separation_hours and time_interval:
        distance = max(1, int(min_separation_hours / time_interval))
    else:
        distance = 5

    peaks, _ = signal.find_peaks(valid.to_numpy(), distance=distance)
    troughs, _ = signal.find_peaks(-valid.to_numpy(), distance=distance)
    return valid.index[peaks], valid.index[troughs]


PeakSelection = dict  # channel -> (list of peak row labels, list of trough row labels)


def detect_all_peaks(smoothed: pd.DataFrame, time_interval: float) -> PeakSelection:
    """Automatic detection for every channel (row labels of `smoothed`)."""
    result = {}
    for channel in CHANNELS:
        peaks, troughs = find_peaks_and_troughs(
            smoothed[channel], PEAK_MIN_SEPARATION_HOURS, time_interval
        )
        result[channel] = ([int(i) for i in peaks], [int(i) for i in troughs])
    return result


def peaks_and_troughs_table(
    smoothed: pd.DataFrame,
    time_interval: float,
    selection: PeakSelection | None = None,
    auto: PeakSelection | None = None,
) -> pd.DataFrame:
    """One row per peak/trough across all channels.

    `selection` is the (possibly hand-edited) set to report; `auto` is what the
    automatic detection found, so each row can be labelled Auto or Manual.
    """
    if selection is None:
        selection = detect_all_peaks(smoothed, time_interval)
    if auto is None:
        auto = selection
    rows = []
    for channel in CHANNELS:
        peaks, troughs = selection[channel]
        auto_peaks, auto_troughs = auto[channel]
        for kind, indices, automatic in (
            ("Peak", peaks, auto_peaks),
            ("Trough", troughs, auto_troughs),
        ):
            for index in indices:
                rows.append(
                    {
                        "Channel": channel,
                        "Type": kind,
                        "Time (Hours)": smoothed.loc[index, "Hours"],
                        "Source": "Auto" if index in automatic else "Manual",
                    }
                )
    return pd.DataFrame(rows, columns=["Channel", "Type", "Time (Hours)", "Source"])


def selection_differs(selection: PeakSelection, auto: PeakSelection) -> bool:
    return any(
        sorted(selection[ch][0]) != sorted(auto[ch][0])
        or sorted(selection[ch][1]) != sorted(auto[ch][1])
        for ch in CHANNELS
    )


PEAKS_JSON_FORMAT = "ttron-peaks-v1"


def save_peaks_json(path, selection: PeakSelection, hours: pd.Series, traces_file: str) -> None:
    """Save the peak/trough selection (by time, so it survives re-processing)."""
    payload = {
        "format": PEAKS_JSON_FORMAT,
        "traces_file": traces_file,
        "channels": {
            channel: {
                "peaks": [float(hours.loc[i]) for i in peaks],
                "troughs": [float(hours.loc[i]) for i in troughs],
            }
            for channel, (peaks, troughs) in selection.items()
        },
    }
    Path(path).write_text(json.dumps(payload, indent=1), encoding="utf-8")


def load_peaks_json(path, hours: pd.Series) -> tuple[PeakSelection, int]:
    """Load a saved selection. Returns (selection, number of times not matched)."""
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    if not isinstance(payload, dict) or payload.get("format") != PEAKS_JSON_FORMAT:
        raise ValueError("This is not a peaks/troughs file saved by this application.")
    valid = hours.dropna()
    skipped = 0

    def to_rows(times) -> list[int]:
        nonlocal skipped
        rows = set()
        for hour in times:
            nearest = (valid - float(hour)).abs().idxmin()
            if abs(float(valid.loc[nearest]) - float(hour)) <= 1e-3:
                rows.add(int(nearest))
            else:
                skipped += 1
        return sorted(rows)

    result = {}
    for channel, entry in payload.get("channels", {}).items():
        if channel in CHANNELS:
            result[channel] = (to_rows(entry.get("peaks", [])), to_rows(entry.get("troughs", [])))
    return result, skipped


# --------------------------------------------------------------------------- #
# Excel / .dat / zip output
# --------------------------------------------------------------------------- #

def build_note_sheet(rows: list[tuple[str, object]]) -> pd.DataFrame:
    """Labels in column A, values in column E (as in the original workbook)."""
    blank = [""] * len(rows)
    return pd.DataFrame(
        {
            "A": [label for label, _ in rows],
            "B": blank,
            "C": blank,
            "D": blank,
            "E": [value for _, value in rows],
        }
    )


def write_main_workbook(
    path: str,
    note: pd.DataFrame,
    raw: pd.DataFrame,
    smoothed: dict[int, pd.DataFrame],
    trend: Trend,
    detrended: pd.DataFrame,
    detrended_smoothed: dict[int, pd.DataFrame],
    peaks_troughs: pd.DataFrame,
) -> None:
    with pd.ExcelWriter(path) as writer:
        note.to_excel(writer, sheet_name="Note", index=False, header=False)
        raw.to_excel(writer, sheet_name="Raw Data")
        for window, frame in smoothed.items():
            frame.to_excel(
                writer, sheet_name=sheet_name(f"{window}-point moving average ({window}PMA)")
            )
        trend.data.to_excel(writer, sheet_name=sheet_name(f"Trend line ({trend.short_label})"))
        detrended.to_excel(writer, sheet_name="Detrended Data")
        for window, frame in detrended_smoothed.items():
            frame.to_excel(writer, sheet_name=sheet_name(f"Detrended Data {window}PMA"))

        for channel in CHANNELS:
            per_channel = pd.DataFrame(
                {
                    "Hours": raw["Hours"],
                    "Raw_data": raw[channel],
                    **{f"{w}PMA": smoothed[w][channel] for w in smoothed},
                    "trend_line": trend.data[channel],
                    "detrended_data": detrended[channel],
                    **{
                        f"detrended_data_{w}PMA": detrended_smoothed[w][channel]
                        for w in detrended_smoothed
                    },
                }
            )
            per_channel.to_excel(writer, sheet_name=f"Channel {channel}")

        if peaks_troughs.empty:
            print("No peaks or troughs detected for any channel.")
        else:
            peaks_troughs.to_excel(writer, sheet_name="Peaks and Troughs", index=False)


def write_raw_only_workbook(path: str, raw: pd.DataFrame, time_interval: float) -> None:
    """Single-sheet workbook for pyBOAT; the interval is encoded in the sheet name."""
    name = sheet_name(f"INTVL = {time_interval:.15f} h")
    with pd.ExcelWriter(path) as writer:
        raw.to_excel(writer, sheet_name=name, index=False, header=True)


def write_dat_files(raw: pd.DataFrame, directory: Path) -> list[Path]:
    """One LumiCycle-readable .dat file per channel: days<TAB>value."""
    days = raw["Hours"] / HOURS_PER_DAY
    written = []
    for channel in CHANNELS:
        path = directory / f"{channel}.dat"
        pd.DataFrame({"Days": days, channel: raw[channel]}).to_csv(
            path, header=False, index=False, sep="\t"
        )
        written.append(path)
    return written


def zip_files(paths: list[Path], archive_path: Path) -> None:
    with zipfile.ZipFile(archive_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in paths:
            archive.write(path, arcname=path.name)


# --------------------------------------------------------------------------- #
# Plotting
# --------------------------------------------------------------------------- #

OVERVIEW_RC = {
    "font.size": 5,
    "axes.titlesize": 4,
    "axes.labelsize": 4,
    "xtick.labelsize": 4,
    "ytick.labelsize": 4,
    "legend.fontsize": 3,
    "figure.titlesize": 5,
}

CHANNEL_PAGE_RC = {
    "font.size": 10,
    "axes.titlesize": 10,
    "axes.labelsize": 10,
    "xtick.labelsize": 9,
    "ytick.labelsize": 9,
    "legend.fontsize": 6,
    "figure.titlesize": 12,
}

FIT_PAGE_RC = {
    "font.size": 8,
    "axes.titlesize": 8,
    "axes.labelsize": 8,
    "xtick.labelsize": 7,
    "ytick.labelsize": 7,
    "legend.fontsize": 6,
}


@dataclass
class AxisStyle:
    """Shared x-axis settings for every time-series plot."""

    x_min: int
    x_max: int
    ticks: list[int]
    minor_tick: float
    show_minor: bool

    def apply(self, ax, x_min: float | None = None) -> None:
        ax.set_xlim(self.x_min if x_min is None else x_min, self.x_max)
        ax.set_xticks(self.ticks)
        if self.show_minor:
            ax.xaxis.set_minor_locator(MultipleLocator(self.minor_tick))
        ax.grid(True, linewidth=0.5, color="lightgray", linestyle="--")


def make_axis_style(hours: pd.Series, major_tick: int, minor_tick: float) -> AxisStyle:
    x_min = int(math.floor(hours.min() / HOURS_PER_DAY)) * 24
    x_max = int(math.ceil(hours.max() / 12)) * 12 + 12
    return AxisStyle(
        x_min=x_min,
        x_max=x_max,
        ticks=list(range(x_min, x_max, major_tick)),
        minor_tick=minor_tick,
        show_minor=bool(minor_tick) and minor_tick < major_tick,
    )


def plot_overview_grid(
    pdf: PdfPages,
    *,
    title: str,
    channel_order: list[str],
    scatter: pd.DataFrame,
    line: pd.DataFrame,
    line_label: str,
    label_suffix: str,
    y_axis_label: str,
    axis: AxisStyle,
    rows: int,
    cols: int,
    start_y_at_zero: bool,
) -> None:
    """One page with a small panel per channel."""
    with plt.rc_context(OVERVIEW_RC):
        figure = plt.figure(figsize=(11, 8.5))
        figure.subplots_adjust(hspace=0.15)
        figure.suptitle(title, fontsize=12)

        for position, channel in enumerate(channel_order, start=1):
            ax = figure.add_subplot(rows, cols, position)
            label = f"Ch # {channel}{label_suffix}"
            ax.scatter(scatter["Hours"], scatter[channel], s=0.1, c="blue", label=label)
            ax.plot(
                line["Hours"], line[channel], "-r", linewidth=0.5, label=line_label
            )
            axis.apply(ax, x_min=axis.x_min - 6)
            if start_y_at_zero:
                ax.set_ylim(0, upper_ylimit(scatter[channel]))
            ax.legend(loc="upper right", fontsize=4)

        figure.text(0.50, 0.06, "Time (hours)", ha="center", fontsize=10)
        figure.text(
            0.08, 0.50, y_axis_label, ha="center", va="center",
            rotation="vertical", fontsize=10,
        )
        pdf.savefig(figure)
        plt.close(figure)


def _draw_actogram(
    ax,
    smoothed: pd.DataFrame,
    peaks,
    troughs,
    day_length: float,
    last_hour: float,
    show_labels: bool,
) -> None:
    """Double-plotted actogram of peak and trough times."""

    def to_actogram_points(indices):
        points = []
        for hour in smoothed.loc[indices, "Hours"].to_numpy():
            day = hour // day_length
            time_in_day = hour - day * day_length
            points.append((time_in_day, day, hour))
            if day > 0:  # second copy of the day, shifted one cycle right
                points.append((time_in_day + day_length, day - 1, hour))
        return points

    for indices, color, label in (
        (peaks, "red", "Peaks"),
        (troughs, "blue", "Troughs"),
    ):
        points = to_actogram_points(indices)
        if not points:
            continue
        xs, ys, original_hours = zip(*points)
        ax.scatter(xs, ys, marker="o", s=20, color=color, label=label)
        if show_labels:
            for x, y, hour in zip(xs, ys, original_hours):
                ax.text(x + 0.5, y, f"{hour:.2f}", fontsize=5, color=color)

    if show_labels:  # leave room for the text labels
        ax.set_xlim(-0.085 * day_length, 2.085 * day_length)
    else:
        ax.set_xlim(0, 2 * day_length)

    ax.set_xticks([fraction * day_length for fraction in np.arange(0, 2.25, 0.25)])

    # Enough decimal places that the tick labels stay distinguishable.
    scaled = day_length if float(day_length).is_integer() else day_length * 10
    extra = 0 if float(day_length).is_integer() else 1
    if scaled % 4 == 0:
        decimals = 0 + extra
    elif scaled % 2 == 0:
        decimals = 1 + extra
    else:
        decimals = 2 + extra
    ax.xaxis.set_major_formatter(FormatStrFormatter(f"%.{decimals}f"))

    if day_length == HOURS_PER_DAY:
        ax.set_title("Peaks and Troughs", fontsize=10)
        ax.set_ylabel("Days", fontsize=10)
    else:
        ax.set_title(
            f"Peaks and Troughs (scaled x-axis: T = {day_length} hours, T{day_length})",
            fontsize=10,
        )
        ax.set_ylabel(f"Days (T{day_length})", fontsize=10)

    n_days = max(last_hour // day_length, 1)
    ax.set_ylim(-n_days * 0.05, n_days * 1.05)
    ax.yaxis.set_major_locator(MultipleLocator(base=1))
    ax.set_xlabel("Time (Hours)", fontsize=10)
    ax.grid(True, linestyle="--", alpha=0.7)
    if ax.get_legend_handles_labels()[0]:  # nothing detected on a flat channel
        ax.legend(loc="upper right", fontsize=6)
    ax.invert_yaxis()  # day 0 at the top


def plot_channel_page(
    pdf: PdfPages,
    *,
    channel: str,
    experiment_number: str,
    raw: pd.DataFrame,
    trend: Trend,
    detrended: pd.DataFrame,
    detrended_smoothed: pd.DataFrame,
    axis: AxisStyle,
    time_interval: float,
    day_length: float,
    last_hour: float,
    label_peaks: bool,
    label_actogram: bool,
    peaks_troughs: tuple | None = None,
) -> None:
    """One page per channel: raw + trend, detrended + peaks, actogram."""
    with plt.rc_context(CHANNEL_PAGE_RC):
        figure = plt.figure(figsize=(8.5, 11))
        figure.suptitle(f"{experiment_number}   Ch # {channel}", fontsize=12)

        # --- raw data with its trend line ---------------------------------- #
        ax_raw = figure.add_subplot(3, 1, 1)
        ax_raw.scatter(
            raw["Hours"], raw[channel], s=3.0, c="violet", label="Bioluminescence"
        )
        ax_raw.plot(
            trend.data["Hours"], trend.data[channel], "-r", linewidth=1.0,
            label=f"Trend line ({trend.label})",
        )
        axis.apply(ax_raw)
        ax_raw.set_ylim(0, upper_ylimit(raw[channel]))
        ax_raw.set_xlabel("Hours", fontsize=10)
        ax_raw.set_ylabel("Bioluminescence", fontsize=10)
        ax_raw.legend(loc="upper right", fontsize=5)

        # --- detrended data with peaks and troughs -------------------------- #
        ax_detrended = figure.add_subplot(3, 1, 2)
        ax_detrended.scatter(
            detrended["Hours"], detrended[channel], s=3.0, c="violet",
            label="Detrended bioluminescence",
        )
        ax_detrended.plot(
            detrended_smoothed["Hours"], detrended_smoothed[channel], "-b",
            linewidth=1.0, label="Smoothed line (9-point moving average, centered)",
        )

        if peaks_troughs is None:
            peaks, troughs = find_peaks_and_troughs(
                detrended_smoothed[channel], PEAK_MIN_SEPARATION_HOURS, time_interval
            )
        else:  # hand-edited selection from the review window
            peaks, troughs = peaks_troughs
        y_span = None
        for indices, color, label, offset_sign, va in (
            (peaks, "red", "Peaks", 1, "bottom"),
            (troughs, "blue", "Troughs", -1, "top"),
        ):
            if len(indices) == 0:
                continue
            hours = detrended_smoothed.loc[indices, "Hours"]
            values = detrended_smoothed.loc[indices, channel]
            ax_detrended.scatter(hours, values, marker="o", s=30, color=color, label=label)
            if not label_peaks:
                continue
            if y_span is None:  # measured once, after both scatters are drawn
                bottom, top = ax_detrended.get_ylim()
                y_span = top - bottom
            for hour, value in zip(hours, values):
                ax_detrended.text(
                    hour, value + offset_sign * y_span * 0.05, f"{hour:.2f} h",
                    fontsize=5, color=color, ha="center", va=va,
                )

        axis.apply(ax_detrended)
        ax_detrended.set_xlabel("Hours", fontsize=10)
        ax_detrended.set_ylabel("Detrended bioluminescence", fontsize=10)
        ax_detrended.legend(loc="upper right", fontsize=5)

        # --- actogram ------------------------------------------------------- #
        ax_actogram = figure.add_subplot(3, 1, 3)
        _draw_actogram(
            ax_actogram, detrended_smoothed, peaks, troughs,
            day_length, last_hour, label_actogram,
        )

        figure.tight_layout(rect=(0, 0.02, 1, 0.98))
        pdf.savefig(figure)
        plt.close(figure)


# --------------------------------------------------------------------------- #
# Damped sine fitting
# --------------------------------------------------------------------------- #

FIT_COLUMNS = [
    "Channel",
    "Fitted Amplitude",
    "Fitted Period (Hours)",
    "Fitted Phase (radians)",
    "Fitted Decay Rate",
    "Fitted Offset",
]


def damped_sine(t, amplitude, period, phase, decay_rate, offset):
    return amplitude * np.exp(-decay_rate * t) * np.sin(2 * np.pi * t / period + phase) + offset


def _empty_fit(channel: str) -> dict:
    return {"Channel": channel, **{column: np.nan for column in FIT_COLUMNS[1:]}}


def fit_damped_sine(times: pd.Series, values: pd.Series) -> tuple[np.ndarray, float]:
    """Fit the damped sine model. Returns the parameters and the normalized phase."""
    amplitude_guess = (values.max() - values.min()) / 2.0
    if amplitude_guess <= 0:
        amplitude_guess = values.std() * 2
    if not amplitude_guess:
        amplitude_guess = 1.0

    guesses = [amplitude_guess, HOURS_PER_DAY, 0.0, 0.01, values.mean()]
    lower = [0, FIT_MIN_PERIOD_HOURS, -2 * np.pi, 0, -np.inf]
    upper = [np.inf, FIT_MAX_PERIOD_HOURS, 2 * np.pi, 0.5, np.inf]

    params, _ = curve_fit(
        damped_sine, times, values, p0=guesses, bounds=(lower, upper), maxfev=10000
    )
    phase = float(params[2]) % (2 * np.pi)
    return params, phase


def plot_fit_page(
    pdf: PdfPages,
    *,
    channel: str,
    experiment_number: str,
    raw: pd.DataFrame,
    trend: Trend,
    detrended: pd.DataFrame,
    fit_times: pd.Series,
    fit_values: pd.Series,
    model_times: pd.Series,
    params: np.ndarray,
    phase: float,
    axis: AxisStyle,
    start_hour: float,
    end_hour: float,
) -> None:
    amplitude, period, _, decay_rate, offset = params
    fitted = damped_sine(model_times, *params)

    with plt.rc_context(FIT_PAGE_RC):
        figure, (ax_raw, ax_detrended) = plt.subplots(
            2, 1, figsize=(8.5, 11), sharex=True
        )
        figure.suptitle(
            f"{experiment_number} - Damped Sine Fit Channel {channel}", fontsize=12
        )

        # The raw-scale fit is the detrended fit put back on top of the trend.
        raw_fit = fitted + trend.data.loc[fit_times.index, channel]
        ax_raw.plot(
            raw["Hours"], raw[channel], "o", markersize=2, color="gray",
            label="Full Raw Data",
        )
        ax_raw.plot(
            trend.data["Hours"], trend.data[channel], "k--", linewidth=1.0,
            label="Trend Line",
        )
        ax_raw.plot(fit_times, raw_fit, "r-", linewidth=1.5, label="Raw Data Fit")
        ax_raw.set_ylabel("Bioluminescence", fontsize=10)
        ax_raw.set_title(
            f"Raw Data Fit (P={period:.2f}h A={amplitude:.2f} Ph={phase:.2f})",
            fontsize=11,
        )
        ax_raw.set_ylim(bottom=0)
        ax_raw.legend(loc="upper right", fontsize=8)
        ax_raw.grid(True, linestyle="--", alpha=0.7)

        ax_detrended.plot(
            detrended["Hours"], detrended[channel], "o", markersize=2, color="gray",
            label="Full Detrended Data",
        )
        ax_detrended.plot(
            fit_times, fit_values, "o", markersize=2, color="blue",
            label=f"Data for Fitting (from {start_hour}h to {end_hour}h)",
        )
        ax_detrended.plot(
            fit_times, fitted, "r-", linewidth=1.5, label="Damped Sine Fit"
        )
        ax_detrended.set_xlabel("Time (Hours)", fontsize=10)
        ax_detrended.set_ylabel("Detrended Bioluminescence", fontsize=10)
        ax_detrended.set_title(
            f"Detrended Data Fit (Dec={decay_rate:.6f} Off={offset:.6f})", fontsize=11
        )
        ax_detrended.legend(loc="upper right", fontsize=8)
        ax_detrended.grid(True, linestyle="--", alpha=0.7)
        ax_detrended.set_xlim(axis.x_min, axis.x_max)
        ax_detrended.xaxis.set_major_locator(MultipleLocator(base=24))
        ax_detrended.xaxis.set_minor_locator(MultipleLocator(base=12))

        figure.tight_layout(rect=(0, 0.03, 1, 0.95))
        pdf.savefig(figure)
        plt.close(figure)


def fit_all_channels(
    pdf: PdfPages,
    *,
    channel_order: list[str],
    experiment_number: str,
    raw: pd.DataFrame,
    trend: Trend,
    detrended: pd.DataFrame,
    axis: AxisStyle,
    start_hour: float,
    end_hour: float,
    on_channel=None,
) -> pd.DataFrame:
    window = detrended[
        detrended["Hours"].between(start_hour, end_hour)
    ]

    results = []
    for channel in channel_order:
        if on_channel is not None:
            on_channel(channel)
        values = window[channel].dropna()
        times = window.loc[values.index, "Hours"]

        if len(values) < FIT_MIN_POINTS:
            print(
                f"Skipping damped sine fit for Channel {channel}: only {len(values)} "
                f"points in the fitting window (need {FIT_MIN_POINTS})."
            )
            results.append(_empty_fit(channel))
            continue

        # The model's decay term assumes t starts at 0.
        model_times = times - times.min()
        try:
            params, phase = fit_damped_sine(model_times, values)
        except (RuntimeError, ValueError) as error:
            print(f"Could not fit damped sine to Channel {channel}: {error}")
            results.append(_empty_fit(channel))
            continue

        amplitude, period, _, decay_rate, offset = params
        results.append(
            {
                "Channel": channel,
                "Fitted Amplitude": amplitude,
                "Fitted Period (Hours)": period,
                "Fitted Phase (radians)": phase,
                "Fitted Decay Rate": decay_rate,
                "Fitted Offset": offset,
            }
        )

        plot_fit_page(
            pdf,
            channel=channel,
            experiment_number=experiment_number,
            raw=raw,
            trend=trend,
            detrended=detrended,
            fit_times=times,
            fit_values=values,
            model_times=model_times,
            params=params,
            phase=phase,
            axis=axis,
            start_hour=start_hour,
            end_hour=end_hour,
        )

    return pd.DataFrame(results, columns=FIT_COLUMNS)





# --------------------------------------------------------------------------- #
# Running an analysis (replaces "CELL 3" of the notebook)
# --------------------------------------------------------------------------- #

DETRENDING_METHODS = ("Sinc Filter", "Moving Average")
DATA_PLOTTING_OPTIONS = ("Plotting the channel 00 data first", "Plotting the channel 00 data last")
SUBPLOT_GRID = {"3-column by 10-row": (10, 3), "4-column by 8-row": (8, 4)}
MAJOR_TICK_HOURS = {"Every 12 hours": 12, "Every 24 hours": 24, "Every 48 hours": 48}
MINOR_TICK_HOURS = {
    "No minor ticks": 0,
    "Every 2 hours": 2,
    "Every 4 hours": 4,
    "Every 6 hours": 6,
    "Every 12 hours": 12,
    "Every 24 hours": 24,
}


@dataclass
class Settings:
    """All user-adjustable parameters (defaults = the notebook's defaults)."""

    experiment_number: str = "CYxxx"
    experiment_title: str = ""
    date_started: str = "2026-01-01"
    detrending_method: str = "Sinc Filter"
    sinc_cutoff_period_hours: float = 48
    sinc_order: int = 101
    moving_average_window: float = 24
    data_plotting: str = "Plotting the channel 00 data last"
    subplot_matrix: str = "3-column by 10-row"
    major_ticks: str = "Every 24 hours"
    minor_ticks: str = "Every 12 hours"
    label_detrended_plot: bool = True
    label_actogram: bool = True
    actogram_x_scale: float = 24.0
    fit_start_hour: float = 24
    fit_end_hour: float = 120


def safe_filename(text: str) -> str:
    """Make a string usable as part of a file name on Windows/macOS/Linux."""
    return re.sub(r'[\\/:*?"<>|\s]+', "_", text.strip()).strip("._")


def expected_outputs(input_path, output_dir, settings: Settings) -> list[Path]:
    """The files run_analysis() will write (used to warn before overwriting)."""
    experiment = (
        safe_filename(settings.experiment_number)
        or safe_filename(Path(input_path).stem)
        or "experiment"
    )
    d = Path(output_dir)
    return [
        d / f"{experiment}_data.xlsx",
        d / f"{experiment}_data_2.xlsx",
        d / f"{experiment}_damped_sine_fit_results.xlsx",
        d / f"{experiment}_data.zip",
        d / f"{experiment}_data_plots.pdf",
    ]


@dataclass
class Analysis:
    """Everything computed before any output file is written."""

    input_path: Path
    settings: Settings
    recording: Recording
    time_interval: float
    smoothed: dict
    trend: Trend
    detrended: pd.DataFrame
    detrended_smoothed: dict
    auto_peaks: PeakSelection

    @property
    def raw(self) -> pd.DataFrame:
        return self.recording.data

    @property
    def detrended_9pma(self) -> pd.DataFrame:
        return self.detrended_smoothed[9]


def _reporter(progress):
    def report(fraction: float, text: str) -> None:
        if progress is not None:
            progress(min(max(fraction, 0.0), 1.0), text)

    return report


def analyze(input_path, settings: Settings, progress=None) -> Analysis:
    """Read, smooth, detrend and detect peaks/troughs (no files are written).

    Progress messages are written with print(); ``progress(fraction, text)`` is
    called (fraction in 0..1) so a GUI can drive a progress bar.
    """
    report = _reporter(progress)
    s = settings
    input_path = Path(input_path)

    start_time = utc_now()
    print(f"Started at {start_time:%Y-%m-%d %H:%M:%S} (UTC)")

    report(0.0, "Reading the data...")
    print(f"\nReading {input_path.name} ...")
    recording = read_traces(str(input_path))
    raw = recording.data
    if recording.n_points < 10:
        raise ValueError(
            f"Only {recording.n_points} data rows were read. "
            "Is this really a Taylortron TRACES.nnn file?"
        )
    time_interval = recording.time_interval
    recording.describe()

    report(0.3, "Detrending...")
    print(f"\nCalculating trend line and detrended data using {s.detrending_method}...")
    smoothed = {window: rolling_mean(raw, window) for window in SMOOTHING_WINDOWS}
    trend = build_trend(
        raw,
        s.detrending_method,
        time_interval,
        window_hours=s.moving_average_window,
        cutoff_period_hours=s.sinc_cutoff_period_hours,
        order=int(s.sinc_order),
    )
    detrended = detrend(raw, trend.data)
    detrended_smoothed = {w: rolling_mean(detrended, w) for w in SMOOTHING_WINDOWS}

    report(0.8, "Detecting peaks and troughs...")
    auto_peaks = detect_all_peaks(detrended_smoothed[9], time_interval)
    print("\nAnalysis finished.")
    report(1.0, "Analysis finished")

    return Analysis(
        input_path=input_path,
        settings=settings,
        recording=recording,
        time_interval=time_interval,
        smoothed=smoothed,
        trend=trend,
        detrended=detrended,
        detrended_smoothed=detrended_smoothed,
        auto_peaks=auto_peaks,
    )


def write_outputs(
    analysis: Analysis,
    output_dir,
    peaks: PeakSelection | None = None,
    progress=None,
) -> list[Path]:
    """Write the Excel files, ZIP and PDF. Returns the list of files written.

    `peaks` is the peak/trough selection to use (default: automatic detection).
    """
    report = _reporter(progress)
    s = analysis.settings
    recording = analysis.recording
    raw = analysis.raw
    time_interval = analysis.time_interval
    trend = analysis.trend
    detrended = analysis.detrended
    smoothed = analysis.smoothed
    detrended_smoothed = analysis.detrended_smoothed
    detrended_9pma = analysis.detrended_9pma
    input_path = analysis.input_path
    if peaks is None:
        peaks = analysis.auto_peaks
    edited = selection_differs(peaks, analysis.auto_peaks)

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    main_workbook, raw_workbook, fit_workbook, zip_archive, plot_pdf = expected_outputs(
        input_path, output_dir, s
    )
    major_tick = MAJOR_TICK_HOURS[s.major_ticks]
    minor_tick = MINOR_TICK_HOURS[s.minor_ticks]
    grid_rows, grid_cols = SUBPLOT_GRID[s.subplot_matrix]
    # Channel 00 is the reference channel, so it is plotted either first or last.
    channel_order = (
        CHANNELS if s.data_plotting.endswith("first") else CHANNELS[1:] + CHANNELS[:1]
    )
    day_length = float(s.actogram_x_scale)

    start_hour, end_hour = s.fit_start_hour, s.fit_end_hour
    if end_hour - start_hour <= 23:  # too short to see a full cycle; fall back
        start_hour, end_hour = 24, 120
        print(f"Fitting window was shorter than 24 h; using {start_hour}-{end_hour} h instead.")

    start_time = utc_now()
    print(f"\nWriting output files ({'hand-edited' if edited else 'automatic'} peaks/troughs)")

    # ---- Excel / .dat / zip --------------------------------------------- #
    report(0.10, "Writing Excel files...")
    print("\nGenerating an Excel file...")
    note = build_note_sheet(
        [
            ("Experiment Number", s.experiment_number),
            ("Experiment Title", s.experiment_title),
            ("Experiment Start Date", s.date_started),
            ("TRACES File", input_path.name),
            ("", ""),
            ("Number of Time Points", recording.n_points),
            ("Total Time Duration (Hours)", recording.duration_hours),
            ("Average Time Interval (Hours)", time_interval),
            ("Detrending Method", s.detrending_method),
            ("Trend Line Parameter", trend.summary),
            ("Peaks and Troughs", "Edited manually" if edited else "Automatic detection"),
            ("", ""),
            ("Data Processed Date and Time (UTC)", utc_now_string()),
        ]
    )

    write_main_workbook(
        str(main_workbook),
        note=note,
        raw=raw,
        smoothed=smoothed,
        trend=trend,
        detrended=detrended,
        detrended_smoothed=detrended_smoothed,
        peaks_troughs=peaks_and_troughs_table(
            detrended_9pma, time_interval, selection=peaks, auto=analysis.auto_peaks
        ),
    )
    print(f"Excel file: {main_workbook.name}  has been generated.")
    report(0.20, "Writing Excel files...")

    write_raw_only_workbook(str(raw_workbook), raw, time_interval)
    print(f"Excel file: {raw_workbook.name}  has been generated.")

    print("\nGenerating a data file for each channel (.dat files) and zipping them...")
    with tempfile.TemporaryDirectory() as tmp:
        zip_files(write_dat_files(raw, Path(tmp)), zip_archive)
    print(f"ZIP file: {zip_archive.name}  has been generated.")
    report(0.25, "Plotting...")

    # ---- plots + fitting ------------------------------------------------- #
    print("\nPlotting...")
    axis = make_axis_style(raw["Hours"], major_tick, minor_tick)
    plt.rcParams.update({"figure.max_open_warning": 0})

    with PdfPages(str(plot_pdf)) as pdf:
        plot_overview_grid(
            pdf,
            title=s.experiment_number,
            channel_order=channel_order,
            scatter=raw,
            line=trend.data,
            line_label=f"trend line ({s.detrending_method})",
            label_suffix="",
            y_axis_label="Bioluminescence",
            axis=axis,
            rows=grid_rows,
            cols=grid_cols,
            start_y_at_zero=True,
        )
        report(0.30, "Plotting...")
        plot_overview_grid(
            pdf,
            title=f"{s.experiment_number} - detrended data ({s.detrending_method})",
            channel_order=channel_order,
            scatter=detrended,
            line=detrended_smoothed[5],
            line_label="5-point moving average",
            label_suffix=" - detrend",
            y_axis_label="Detrended Bioluminescence",
            axis=axis,
            rows=grid_rows,
            cols=grid_cols,
            start_y_at_zero=False,
        )

        print("\nPlotting individual channels with actograms...")
        for i, channel in enumerate(channel_order):
            report(0.35 + 0.35 * i / N_CHANNELS, f"Plotting channel {channel}...")
            plot_channel_page(
                pdf,
                channel=channel,
                experiment_number=s.experiment_number,
                raw=raw,
                trend=trend,
                detrended=detrended,
                detrended_smoothed=detrended_9pma,
                axis=axis,
                time_interval=time_interval,
                day_length=day_length,
                last_hour=recording.last_hour,
                label_peaks=s.label_detrended_plot,
                label_actogram=s.label_actogram,
                peaks_troughs=peaks[channel],
            )

        print("\nPerforming damped sine curve fitting...")
        counter = {"n": 0}

        def on_channel(channel: str) -> None:
            report(0.70 + 0.28 * counter["n"] / N_CHANNELS, f"Fitting channel {channel}...")
            counter["n"] += 1

        fit_results = fit_all_channels(
            pdf,
            channel_order=channel_order,
            experiment_number=s.experiment_number,
            raw=raw,
            trend=trend,
            detrended=detrended,
            axis=axis,
            start_hour=start_hour,
            end_hour=end_hour,
            on_channel=on_channel,
        )

    with pd.ExcelWriter(str(fit_workbook)) as writer:
        fit_results.to_excel(writer, sheet_name="Damped Sine Fit", index=False)
    print(f"\nDamped sine fitting results written to: {fit_workbook.name}")

    outputs = [main_workbook, raw_workbook, fit_workbook, zip_archive, plot_pdf]
    end_time = utc_now()
    print(f"\nElapsed time: {(end_time - start_time).total_seconds():.0f} seconds")
    print(f"Completed at {end_time:%Y-%m-%d %H:%M:%S} (UTC)")
    report(1.0, "Done")
    return outputs


def run_analysis(
    input_path, output_dir, settings: Settings, progress=None, peaks: PeakSelection | None = None
) -> list[Path]:
    """analyze() + write_outputs() in one call (no review step)."""
    analysis = analyze(input_path, settings, progress)
    return write_outputs(analysis, output_dir, peaks, progress)


# =========================================================================== #
# 2. Desktop GUI
# =========================================================================== #

def _fit_geometry(window, width: int, height: int) -> None:
    """Set the window size, but never larger than the screen (taskbar allowance)."""
    width = min(width, window.winfo_screenwidth() - 20)
    height = min(height, window.winfo_screenheight() - 80)
    window.geometry(f"{width}x{height}")


class _QueueWriter:
    """File-like object that forwards print() output to the GUI thread."""

    def __init__(self, q: queue.Queue) -> None:
        self._q = q

    def write(self, text: str) -> int:
        if text:
            self._q.put(("log", text))
        return len(text)

    def flush(self) -> None:
        pass


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.queue: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.last_output_dir: Path | None = None
        self.analysis: Analysis | None = None
        self.pending_out_dir = ""

        root.title(APP_TITLE)
        _fit_geometry(root, 1120, 900)
        root.minsize(1000, 680)

        self._make_variables()
        self._build_ui()
        self._on_method_change()
        self.root.after(100, self._poll_queue)

    # ------------------------------------------------------------------ #
    # Variables (defaults = the notebook's defaults)
    # ------------------------------------------------------------------ #

    def _make_variables(self) -> None:
        d = Settings()
        self._spin_specs: list = []
        self.in_var = tk.StringVar()
        self.out_var = tk.StringVar()

        self.exp_var = tk.StringVar(value=d.experiment_number)
        self.title_var = tk.StringVar(value=d.experiment_title)
        self.date_var = tk.StringVar(value=dt.date.today().isoformat())

        self.method_var = tk.StringVar(value=d.detrending_method)
        self.cutoff_var = tk.IntVar(value=int(d.sinc_cutoff_period_hours))
        self.order_var = tk.IntVar(value=int(d.sinc_order))
        self.window_var = tk.IntVar(value=int(d.moving_average_window))

        self.plot_order_var = tk.StringVar(value=d.data_plotting)
        self.grid_var = tk.StringVar(value=d.subplot_matrix)
        self.major_var = tk.StringVar(value=d.major_ticks)
        self.minor_var = tk.StringVar(value=d.minor_ticks)
        self.label_det_var = tk.BooleanVar(value=d.label_detrended_plot)
        self.label_act_var = tk.BooleanVar(value=d.label_actogram)
        self.acto_var = tk.DoubleVar(value=d.actogram_x_scale)

        self.fit_start_var = tk.IntVar(value=int(d.fit_start_hour))
        self.fit_end_var = tk.IntVar(value=int(d.fit_end_hour))

        self.review_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Select a TRACES file to begin.")
        self.summary_var = tk.StringVar(value="")

    # ------------------------------------------------------------------ #
    # UI construction
    # ------------------------------------------------------------------ #

    def _build_ui(self) -> None:
        root = self.root
        root.columnconfigure(0, weight=0)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(0, weight=1)

        left = ttk.Frame(root, padding=(12, 10, 6, 10))
        left.grid(row=0, column=0, sticky="ns")
        right = ttk.Frame(root, padding=(6, 10, 12, 10))
        right.grid(row=0, column=1, sticky="nsew")

        self._build_files(left)
        self._build_experiment(left)
        self._build_detrending(left)
        self._build_plots(left)
        self._build_fit(left)
        self._build_right(right)

    @staticmethod
    def _box(parent, text: str) -> ttk.LabelFrame:
        box = ttk.LabelFrame(parent, text=text, padding=8)
        box.pack(fill="x", pady=(0, 8))
        box.columnconfigure(1, weight=1)
        return box

    @staticmethod
    def _label(box, row: int, text: str) -> ttk.Label:
        label = ttk.Label(box, text=text)
        label.grid(row=row, column=0, sticky="w", padx=(0, 10), pady=2)
        return label

    def _combo(self, box, row: int, text: str, var, values) -> ttk.Combobox:
        self._label(box, row, text)
        combo = ttk.Combobox(
            box, textvariable=var, values=list(values), state="readonly", width=34
        )
        combo.grid(row=row, column=1, columnspan=2, sticky="ew", pady=2)
        return combo

    def _slider(self, box, row: int, text: str, var, lo, hi, step):
        """Labelled numeric spinner (name kept from the former slider version)."""
        label = self._label(box, row, text)
        is_float = isinstance(step, float)
        spin = ttk.Spinbox(
            box,
            from_=lo,
            to=hi,
            increment=step,
            textvariable=var,
            width=10,
            justify="right",
            format="%.1f" if is_float else "%.0f",
        )
        spin.grid(row=row, column=1, columnspan=2, sticky="w", pady=2)
        self._spin_specs.append((text, var, lo, hi, spin))
        return label, spin

    def _validate_spinners(self) -> None:
        """Make sure every spinner holds a number inside its range (raises ValueError)."""
        for text, var, lo, hi, spin in self._spin_specs:
            if spin.instate(["disabled"]):
                continue
            try:
                value = float(var.get())
            except (tk.TclError, ValueError):
                raise ValueError(f"'{text}' must be a number.") from None
            if not lo <= value <= hi:
                raise ValueError(f"'{text}' must be between {lo} and {hi}.")

    def _build_files(self, parent) -> None:
        box = self._box(parent, "1. Files")
        self._label(box, 0, "TRACES file")
        ttk.Entry(box, textvariable=self.in_var, width=34).grid(
            row=0, column=1, sticky="ew", pady=2
        )
        ttk.Button(box, text="Browse...", command=self._browse_input).grid(
            row=0, column=2, padx=(6, 0), pady=2
        )
        self._label(box, 1, "Output folder")
        ttk.Entry(box, textvariable=self.out_var, width=34).grid(
            row=1, column=1, sticky="ew", pady=2
        )
        ttk.Button(box, text="Browse...", command=self._browse_output).grid(
            row=1, column=2, padx=(6, 0), pady=2
        )
        ttk.Label(
            box, textvariable=self.summary_var, wraplength=380, foreground="#555555"
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(4, 0))

    def _build_experiment(self, parent) -> None:
        box = self._box(parent, "2. Experiment")
        self._label(box, 0, "Experiment number")
        ttk.Entry(box, textvariable=self.exp_var).grid(
            row=0, column=1, columnspan=2, sticky="ew", pady=2
        )
        self._label(box, 1, "Experiment title")
        ttk.Entry(box, textvariable=self.title_var).grid(
            row=1, column=1, columnspan=2, sticky="ew", pady=2
        )
        self._label(box, 2, "Start date (YYYY-MM-DD)")
        ttk.Entry(box, textvariable=self.date_var).grid(
            row=2, column=1, columnspan=2, sticky="ew", pady=2
        )

    def _build_detrending(self, parent) -> None:
        box = self._box(parent, "3. Detrending")
        combo = self._combo(box, 0, "Method", self.method_var, DETRENDING_METHODS)
        combo.bind("<<ComboboxSelected>>", lambda _e: self._on_method_change())
        # each entry is a (label, spinner) pair, enabled/disabled with the method
        self._sinc_widgets = [
            self._slider(box, 1, "Sinc cutoff period (hours)", self.cutoff_var, 1, 240, 1),
            self._slider(box, 2, "Sinc filter order (odd)", self.order_var, 1, 361, 2),
        ]
        self._ma_widgets = [
            self._slider(box, 3, "Moving-average window (hours)", self.window_var, 1, 120, 1),
        ]

    def _build_plots(self, parent) -> None:
        box = self._box(parent, "4. Plots")
        self._combo(box, 0, "Channel 00", self.plot_order_var, DATA_PLOTTING_OPTIONS)
        self._combo(box, 1, "Subplot matrix", self.grid_var, SUBPLOT_GRID)
        self._combo(box, 2, "Major ticks", self.major_var, MAJOR_TICK_HOURS)
        self._combo(box, 3, "Minor ticks", self.minor_var, MINOR_TICK_HOURS)
        ttk.Checkbutton(
            box,
            text="Label peaks/troughs in detrended-data plot",
            variable=self.label_det_var,
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(4, 0))
        ttk.Checkbutton(
            box, text="Label peaks/troughs in actogram", variable=self.label_act_var
        ).grid(row=5, column=0, columnspan=3, sticky="w")
        self._slider(box, 6, "Actogram x-axis scale (hours)", self.acto_var, 12, 60, 0.1)

    def _build_fit(self, parent) -> None:
        box = self._box(parent, "5. Damped sine fit (time range)")
        row = ttk.Frame(box)
        row.grid(row=0, column=0, columnspan=3, sticky="w")
        for column, (text, var) in enumerate(
            (("Start hour", self.fit_start_var), ("End hour", self.fit_end_var))
        ):
            ttk.Label(row, text=text).grid(
                row=0, column=column * 2, sticky="w", padx=(0 if column == 0 else 16, 6), pady=2
            )
            spin = ttk.Spinbox(
                row, from_=0, to=360, increment=1, textvariable=var,
                width=8, justify="right", format="%.0f",
            )
            spin.grid(row=0, column=column * 2 + 1, sticky="w", pady=2)
            self._spin_specs.append((text, var, 0, 360, spin))

    def _build_right(self, parent) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(3, weight=1)

        buttons = ttk.Frame(parent)
        buttons.grid(row=0, column=0, sticky="ew")
        self.run_button = ttk.Button(buttons, text="Convert", command=self._on_run)
        self.run_button.pack(side="left")
        ttk.Checkbutton(
            buttons, text="Review peaks/troughs before export", variable=self.review_var
        ).pack(side="left", padx=(10, 0))
        self.open_button = ttk.Button(
            buttons, text="Open output folder", command=self._open_output, state="disabled"
        )
        self.open_button.pack(side="left", padx=8)

        self.progress = ttk.Progressbar(parent, mode="determinate", maximum=100)
        self.progress.grid(row=1, column=0, sticky="ew", pady=(10, 2))
        ttk.Label(parent, textvariable=self.status_var).grid(row=2, column=0, sticky="w")

        log_frame = ttk.Frame(parent)
        log_frame.grid(row=3, column=0, sticky="nsew", pady=(8, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log = tk.Text(log_frame, wrap="word", state="disabled", height=20)
        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

    # ------------------------------------------------------------------ #
    # Event handlers
    # ------------------------------------------------------------------ #

    def _on_method_change(self) -> None:
        sinc = self.method_var.get() == "Sinc Filter"
        for group, enabled in ((self._sinc_widgets, sinc), (self._ma_widgets, not sinc)):
            for label, spin in group:
                label.state(["!disabled"] if enabled else ["disabled"])
                spin.state(["!disabled"] if enabled else ["disabled"])

    def _browse_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Select a TRACES.nnn file",
            filetypes=[
                ("All files", "*.*"),
                ("TRACES files", "TRACES.*"),
            ],
        )
        if not path:
            return
        self.in_var.set(path)
        if not self.out_var.get().strip():
            self.out_var.set(str(Path(path).parent))
        self._summarize_input(path)

    def _browse_output(self) -> None:
        path = filedialog.askdirectory(title="Select the output folder")
        if path:
            self.out_var.set(path)

    def _summarize_input(self, path: str) -> None:
        """Quick look at the chosen file so a wrong file is noticed early."""
        try:
            rec = read_traces(path)
            self.summary_var.set(
                f"{rec.n_points} rows, {rec.first_hour:g}\u2013{rec.last_hour:g} h "
                f"({rec.duration_hours / 24:.2f} days), "
                f"average interval {rec.time_interval:.3f} h"
            )
        except Exception as error:  # noqa: BLE001 - any parse problem is reported
            self.summary_var.set(f"Warning: could not read this as a TRACES file ({error})")

    def _collect_settings(self) -> Settings:
        """Validate the form and build a Settings object (raises ValueError)."""
        experiment = self.exp_var.get().strip()
        if not experiment:
            raise ValueError("Please enter an experiment number.")
        date_text = self.date_var.get().strip()
        try:
            dt.datetime.strptime(date_text, "%Y-%m-%d")
        except ValueError:
            raise ValueError("The start date must be in YYYY-MM-DD format.") from None
        self._validate_spinners()

        return Settings(
            experiment_number=experiment,
            experiment_title=self.title_var.get().strip(),
            date_started=date_text,
            detrending_method=self.method_var.get(),
            sinc_cutoff_period_hours=self.cutoff_var.get(),
            sinc_order=self.order_var.get(),
            moving_average_window=self.window_var.get(),
            data_plotting=self.plot_order_var.get(),
            subplot_matrix=self.grid_var.get(),
            major_ticks=self.major_var.get(),
            minor_ticks=self.minor_var.get(),
            label_detrended_plot=bool(self.label_det_var.get()),
            label_actogram=bool(self.label_act_var.get()),
            actogram_x_scale=round(float(self.acto_var.get()), 1),
            fit_start_hour=self.fit_start_var.get(),
            fit_end_hour=self.fit_end_var.get(),
        )

    def _on_run(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            return

        in_path = self.in_var.get().strip()
        if not in_path or not Path(in_path).is_file():
            messagebox.showerror(APP_TITLE, "Please select an existing TRACES file.")
            return
        out_dir = self.out_var.get().strip() or str(Path(in_path).parent)
        self.out_var.set(out_dir)

        try:
            settings = self._collect_settings()
        except ValueError as error:
            messagebox.showerror(APP_TITLE, str(error))
            return

        existing = [
            p.name for p in expected_outputs(in_path, out_dir, settings) if p.exists()
        ]
        if existing and not messagebox.askyesno(
            APP_TITLE,
            "These files already exist in the output folder and will be overwritten:\n\n"
            + "\n".join(existing)
            + "\n\nContinue?",
        ):
            return

        self._clear_log()
        self.progress["value"] = 0
        self.status_var.set("Starting...")
        self.run_button.configure(state="disabled")
        self.open_button.configure(state="disabled")

        self.pending_out_dir = out_dir
        self.analysis = None
        self._run_in_thread(self._analysis_worker, in_path, settings)

    def _open_output(self) -> None:
        folder = self.last_output_dir
        if folder is None:
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(folder))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", str(folder)], check=False)
            else:
                subprocess.run(["xdg-open", str(folder)], check=False)
        except OSError as error:
            messagebox.showerror(APP_TITLE, f"Could not open the folder: {error}")

    # ------------------------------------------------------------------ #
    # Background work
    # ------------------------------------------------------------------ #

    def _run_in_thread(self, target, *args) -> None:
        self.worker = threading.Thread(target=target, args=args, daemon=True)
        self.worker.start()

    def _progress_callback(self):
        return lambda fraction, text: self.queue.put(("progress", fraction, text))

    def _analysis_worker(self, in_path: str, settings: Settings) -> None:
        """Background thread, step 1: read, detrend and detect peaks."""
        writer = _QueueWriter(self.queue)
        try:
            with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
                analysis = analyze(in_path, settings, progress=self._progress_callback())
            self.queue.put(("analyzed", analysis))
        except Exception as error:  # noqa: BLE001 - shown to the user
            self.queue.put(("error", traceback.format_exc(), str(error)))

    def _export_worker(self, analysis: Analysis, out_dir: str, peaks) -> None:
        """Background thread, step 2: write Excel / ZIP / PDF files."""
        writer = _QueueWriter(self.queue)
        try:
            with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
                outputs = write_outputs(
                    analysis, out_dir, peaks, progress=self._progress_callback()
                )
            self.queue.put(("done", outputs))
        except Exception as error:  # noqa: BLE001 - shown to the user
            self.queue.put(("error", traceback.format_exc(), str(error)))

    def _on_analyzed(self, analysis: Analysis) -> None:
        self.analysis = analysis
        if self.review_var.get():
            self.status_var.set("Review the detected peaks and troughs...")
            ReviewWindow(
                self.root, analysis, on_ok=self._start_export, on_cancel=self._on_review_cancel
            )
        else:
            self._start_export(analysis.auto_peaks)

    def _start_export(self, peaks) -> None:
        self.progress["value"] = 0
        self.status_var.set("Writing output files...")
        self._run_in_thread(self._export_worker, self.analysis, self.pending_out_dir, peaks)

    def _on_review_cancel(self) -> None:
        self.run_button.configure(state="normal")
        self.progress["value"] = 0
        self.status_var.set("Cancelled. No files were written.")
        self._append_log("\nCancelled in the review window. No files were written.\n")

    def _poll_queue(self) -> None:
        try:
            while True:
                message = self.queue.get_nowait()
                kind = message[0]
                if kind == "log":
                    self._append_log(message[1])
                elif kind == "progress":
                    self.progress["value"] = message[1] * 100
                    self.status_var.set(message[2])
                elif kind == "analyzed":
                    self._on_analyzed(message[1])
                elif kind == "done":
                    self._on_done(message[1])
                elif kind == "error":
                    self._on_error(message[1], message[2])
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _on_done(self, outputs) -> None:
        self.run_button.configure(state="normal")
        self.progress["value"] = 100
        self.status_var.set("Done.")
        if outputs:
            self.last_output_dir = Path(outputs[0]).parent
            self.open_button.configure(state="normal")
        names = "\n".join(Path(p).name for p in outputs)
        messagebox.showinfo(
            APP_TITLE, f"Conversion finished. Files written to:\n{self.last_output_dir}\n\n{names}"
        )

    def _on_error(self, details: str, short: str) -> None:
        self.run_button.configure(state="normal")
        self.status_var.set("Failed.")
        self._append_log("\n" + details)
        messagebox.showerror(APP_TITLE, f"The conversion failed:\n\n{short}")

    # ------------------------------------------------------------------ #
    # Log helpers
    # ------------------------------------------------------------------ #

    def _append_log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")


class ReviewWindow:
    """Modal window for checking and hand-editing the detected peaks/troughs."""

    HIT_RADIUS_PX = 25  # how close (in pixels) a click must be to grab a marker

    HELP = (
        "Remove: choose 'Remove' and click a marker.   "
        "Add: choose 'Add peak' or 'Add trough' and click near the curve "
        "(snaps to the nearest sample point).   "
        "While a zoom/pan tool of the toolbar is active, clicks do not edit anything.   "
        "Switch channels with the list or the Left/Right keys."
    )

    def __init__(self, parent, analysis: Analysis, on_ok, on_cancel) -> None:
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
        from matplotlib.figure import Figure

        self.analysis = analysis
        self.on_ok = on_ok
        self.on_cancel = on_cancel
        self.data = analysis.detrended
        self.smooth = analysis.detrended_9pma
        self.auto = analysis.auto_peaks
        self.selection = {ch: (list(p), list(t)) for ch, (p, t) in self.auto.items()}
        self.channel = CHANNELS[0]
        self._marker_artists: list = []

        win = self.win = tk.Toplevel(parent)
        win.title("Review peaks and troughs")
        _fit_geometry(win, 1200, 980)
        win.minsize(900, 620)
        win.transient(parent)
        win.protocol("WM_DELETE_WINDOW", self._cancel)
        win.columnconfigure(1, weight=1)
        win.rowconfigure(1, weight=1)

        self.mode_var = tk.StringVar(value="remove")
        self.status_var = tk.StringVar(value="")

        ttk.Label(win, text=self.HELP, wraplength=1120, padding=(10, 8)).grid(
            row=0, column=0, columnspan=2, sticky="ew"
        )

        # --- channel list ------------------------------------------------- #
        left = ttk.Frame(win, padding=(10, 0, 4, 0))
        left.grid(row=1, column=0, sticky="ns")
        ttk.Label(left, text="Channel   (P = peaks, T = troughs, * = edited)").pack(anchor="w")
        list_frame = ttk.Frame(left)
        list_frame.pack(fill="y", expand=True)
        self.listbox = tk.Listbox(list_frame, width=26, height=30, exportselection=False)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="y")
        scrollbar.pack(side="left", fill="y")
        for channel in CHANNELS:
            self.listbox.insert("end", self._list_text(channel))
        self.listbox.selection_set(0)
        self.listbox.bind("<<ListboxSelect>>", self._on_list_select)

        # --- plot area ------------------------------------------------------- #
        centre = ttk.Frame(win, padding=(4, 0, 10, 0))
        centre.grid(row=1, column=1, sticky="nsew")
        controls = ttk.Frame(centre)
        controls.pack(fill="x", pady=(0, 4))
        for text, value in (
            ("Remove", "remove"),
            ("Add peak", "add_peak"),
            ("Add trough", "add_trough"),
        ):
            ttk.Radiobutton(controls, text=text, value=value, variable=self.mode_var).pack(
                side="left", padx=(0, 14)
            )
        ttk.Button(controls, text="Next \u25b6", command=lambda: self._step(1)).pack(side="right")
        ttk.Button(controls, text="\u25c0 Previous", command=lambda: self._step(-1)).pack(
            side="right", padx=4
        )
        ttk.Button(controls, text="Reset this channel", command=self._reset_channel).pack(
            side="right", padx=(0, 12)
        )

        self.fig = Figure(figsize=(8, 7.2), dpi=100)
        self.ax = self.fig.add_subplot(2, 1, 1)
        self.ax_act = self.fig.add_subplot(2, 1, 2)
        self.canvas = FigureCanvasTkAgg(self.fig, master=centre)
        self.toolbar = NavigationToolbar2Tk(self.canvas, centre, pack_toolbar=False)
        self.toolbar.pack(side="bottom", fill="x")
        self.canvas.get_tk_widget().pack(side="top", fill="both", expand=True)
        self.canvas.mpl_connect("button_press_event", self._on_click)

        # --- bottom bar ------------------------------------------------------ #
        bottom = ttk.Frame(win, padding=10)
        bottom.grid(row=2, column=0, columnspan=2, sticky="ew")
        ttk.Label(bottom, textvariable=self.status_var).pack(side="left")
        ttk.Button(bottom, text="OK \u2013 export files", command=self._ok).pack(side="right")
        ttk.Button(bottom, text="Cancel", command=self._cancel).pack(side="right", padx=6)
        ttk.Button(bottom, text="Reset all", command=self._reset_all).pack(side="right", padx=6)
        ttk.Button(bottom, text="Load edits...", command=self._load_edits).pack(
            side="right", padx=6
        )
        ttk.Button(bottom, text="Save edits...", command=self._save_edits).pack(side="right")

        win.bind("<Left>", lambda _e: self._step(-1))
        win.bind("<Right>", lambda _e: self._step(1))

        self._draw_full()
        try:  # make the window modal; harmless if the platform refuses
            win.wait_visibility()
            win.grab_set()
        except tk.TclError:
            pass
        win.focus_set()

    # ------------------------------------------------------------------ #
    # Channel list
    # ------------------------------------------------------------------ #

    def _is_edited(self, channel: str) -> bool:
        peaks, troughs = self.selection[channel]
        auto_peaks, auto_troughs = self.auto[channel]
        return sorted(peaks) != sorted(auto_peaks) or sorted(troughs) != sorted(auto_troughs)

    def _list_text(self, channel: str) -> str:
        peaks, troughs = self.selection[channel]
        mark = "  *" if self._is_edited(channel) else ""
        return f"Ch {channel}    P {len(peaks)}   T {len(troughs)}{mark}"

    def _refresh_list_item(self, channel: str) -> None:
        i = CHANNELS.index(channel)
        self.listbox.delete(i)
        self.listbox.insert(i, self._list_text(channel))
        if channel == self.channel:
            self.listbox.selection_set(i)

    def _on_list_select(self, _event=None) -> None:
        selected = self.listbox.curselection()
        if selected and CHANNELS[selected[0]] != self.channel:
            self.channel = CHANNELS[selected[0]]
            self._draw_full()

    def _step(self, delta: int) -> None:
        i = min(max(CHANNELS.index(self.channel) + delta, 0), len(CHANNELS) - 1)
        self.listbox.selection_clear(0, "end")
        self.listbox.selection_set(i)
        self.listbox.see(i)
        if CHANNELS[i] != self.channel:
            self.channel = CHANNELS[i]
            self._draw_full()

    # ------------------------------------------------------------------ #
    # Drawing
    # ------------------------------------------------------------------ #

    def _draw_full(self) -> None:
        """Redraw the curve, the markers and the actogram for the current channel."""
        channel = self.channel
        ax = self.ax
        ax.clear()
        self._marker_artists = []  # ax.clear() already discarded the old markers
        ax.scatter(
            self.data["Hours"], self.data[channel], s=3.0, c="violet",
            label="Detrended bioluminescence",
        )
        ax.plot(
            self.smooth["Hours"], self.smooth[channel], "-b", linewidth=1.0,
            label="9-point moving average",
        )
        ax.grid(True, linewidth=0.5, color="lightgray", linestyle="--")
        ax.set_xlabel("Hours")
        ax.set_ylabel("Detrended bioluminescence")
        ax.set_title(f"{self.analysis.settings.experiment_number}   Ch # {channel}")
        ax.autoscale_view()
        low, high = ax.get_ylim()
        pad = (high - low) * 0.10
        ax.set_ylim(low - pad, high + pad)
        ax.set_autoscale_on(False)  # keep the view fixed while markers change
        self.fig.tight_layout()
        self.toolbar.update()  # forget the previous channel's zoom history
        self._draw_markers()
        self._draw_actogram_preview()
        self.status_var.set(f"Channel {channel}")

    def _draw_markers(self) -> None:
        for artist in self._marker_artists:
            try:
                artist.remove()
            except (NotImplementedError, ValueError):
                pass  # already gone (e.g. the axes were cleared)
        self._marker_artists = []
        ax = self.ax
        channel = self.channel
        peaks, troughs = self.selection[channel]
        low, high = ax.get_ylim()
        span = high - low
        for indices, color, name, sign, va in (
            (peaks, "red", "Peaks", 1, "bottom"),
            (troughs, "blue", "Troughs", -1, "top"),
        ):
            if not indices:
                continue
            hours = self.smooth.loc[indices, "Hours"]
            values = self.smooth.loc[indices, channel]
            self._marker_artists.append(
                ax.scatter(hours, values, marker="o", s=45, color=color, zorder=5, label=name)
            )
            for hour, value in zip(hours, values):
                self._marker_artists.append(
                    ax.text(
                        hour, value + sign * span * 0.03, f"{hour:.2f} h",
                        fontsize=7, color=color, ha="center", va=va, clip_on=True,
                    )
                )
        ax.legend(loc="upper right", fontsize=7, framealpha=0.6)

    def _draw_actogram_preview(self) -> None:
        settings = self.analysis.settings
        peaks, troughs = self.selection[self.channel]
        self.ax_act.clear()
        _draw_actogram(
            self.ax_act, self.smooth, peaks, troughs,
            float(settings.actogram_x_scale), self.analysis.recording.last_hour,
            settings.label_actogram,
        )
        self.ax_act.set_navigate(False)  # zoom/pan applies to the top plot only
        self.canvas.draw_idle()

    def _after_edit(self, message: str) -> None:
        self._refresh_list_item(self.channel)
        self._draw_markers()
        self._draw_actogram_preview()
        self.status_var.set(message)

    # ------------------------------------------------------------------ #
    # Editing
    # ------------------------------------------------------------------ #

    def _toolbar_active(self) -> bool:
        mode = getattr(self.toolbar, "mode", "")
        return bool(getattr(mode, "value", mode))

    def _on_click(self, event) -> None:
        if event.inaxes is not self.ax or event.button != 1 or event.xdata is None:
            return
        if self._toolbar_active():
            return
        mode = self.mode_var.get()
        if mode == "remove":
            self._remove_nearest(event.x, event.y)
        else:
            self._add_at(event.xdata, "peak" if mode == "add_peak" else "trough")

    def _remove_nearest(self, pixel_x: float, pixel_y: float) -> None:
        channel = self.channel
        peaks, troughs = self.selection[channel]
        candidates = [("peak", i) for i in peaks] + [("trough", i) for i in troughs]
        if not candidates:
            self.status_var.set("There are no markers in this channel.")
            return
        points = np.array(
            [(self.smooth.loc[i, "Hours"], self.smooth.loc[i, channel]) for _, i in candidates],
            dtype=float,
        )
        pixels = self.ax.transData.transform(points)
        distances = np.hypot(pixels[:, 0] - pixel_x, pixels[:, 1] - pixel_y)
        nearest = int(np.argmin(distances))
        if distances[nearest] > self.HIT_RADIUS_PX:
            self.status_var.set("No marker near the click.")
            return
        kind, index = candidates[nearest]
        (peaks if kind == "peak" else troughs).remove(index)
        hour = float(self.smooth.loc[index, "Hours"])
        self._after_edit(f"Removed {kind} at {hour:.2f} h.")

    def _add_at(self, x: float, kind: str) -> None:
        channel = self.channel
        valid = self.smooth.loc[self.smooth[channel].notna(), "Hours"]
        if valid.empty:
            self.status_var.set("This channel has no data.")
            return
        index = int((valid - x).abs().idxmin())
        peaks, troughs = self.selection[channel]
        target, other = (peaks, troughs) if kind == "peak" else (troughs, peaks)
        hour = float(valid.loc[index])
        if index in target:
            self.status_var.set(f"There is already a {kind} at {hour:.2f} h.")
            return
        if index in other:
            self.status_var.set(
                f"{hour:.2f} h is already marked as the opposite type; remove it first."
            )
            return
        target.append(index)
        target.sort()
        self._after_edit(f"Added {kind} at {hour:.2f} h.")

    def _reset_channel(self) -> None:
        peaks, troughs = self.auto[self.channel]
        self.selection[self.channel] = (list(peaks), list(troughs))
        self._after_edit(f"Channel {self.channel} reset to the automatic detection.")

    def _reset_all(self) -> None:
        if not messagebox.askyesno(
            APP_TITLE, "Discard all edits in every channel?", parent=self.win
        ):
            return
        self.selection = {ch: (list(p), list(t)) for ch, (p, t) in self.auto.items()}
        for channel in CHANNELS:
            self._refresh_list_item(channel)
        self._draw_markers()
        self._draw_actogram_preview()
        self.status_var.set("All channels reset to the automatic detection.")

    # ------------------------------------------------------------------ #
    # Save / load / finish
    # ------------------------------------------------------------------ #

    def _save_edits(self) -> None:
        path = filedialog.asksaveasfilename(
            parent=self.win,
            title="Save peak/trough edits",
            defaultextension=".json",
            initialfile=f"{self.analysis.input_path.name}_peaks.json",
            filetypes=[("JSON files", "*.json")],
        )
        if not path:
            return
        try:
            save_peaks_json(
                path, self.selection, self.smooth["Hours"], self.analysis.input_path.name
            )
        except OSError as error:
            messagebox.showerror(APP_TITLE, f"Could not save the file: {error}", parent=self.win)
            return
        self.status_var.set(f"Saved: {Path(path).name}")

    def _load_edits(self) -> None:
        path = filedialog.askopenfilename(
            parent=self.win,
            title="Load peak/trough edits",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            loaded, skipped = load_peaks_json(path, self.smooth["Hours"])
        except (OSError, ValueError) as error:
            messagebox.showerror(APP_TITLE, f"Could not load the file: {error}", parent=self.win)
            return
        self.selection.update(loaded)
        for channel in CHANNELS:
            self._refresh_list_item(channel)
        self._draw_markers()
        self._draw_actogram_preview()
        note = f" ({skipped} time points did not match this recording and were skipped)" if skipped else ""
        self.status_var.set(f"Loaded: {Path(path).name}{note}")

    def _close(self) -> None:
        try:
            self.win.grab_release()
        except tk.TclError:
            pass
        self.win.destroy()

    def _ok(self) -> None:
        selection = {ch: (sorted(p), sorted(t)) for ch, (p, t) in self.selection.items()}
        self._close()
        self.on_ok(selection)

    def _cancel(self) -> None:
        self._close()
        self.on_cancel()


def main() -> None:
    if sys.platform.startswith("win"):  # crisp text on high-DPI displays
        try:
            import ctypes

            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:  # noqa: BLE001
            pass

    root = tk.Tk()
    if sys.platform.startswith("linux"):
        style = ttk.Style(root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
