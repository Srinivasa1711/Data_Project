"""
platform_helpers.py
Small, dependency-light helpers shared by recommendations_engine.py and
report_builder.py.

recommendations_engine.py originally imported these three functions from
analytics.py (Part 3). In the merged project the analytics code lives inside
the main app file, so these helpers are provided here as a standalone module
instead. They only READ the DataFrame they are given.
"""

import warnings

import pandas as pd


def detect_numeric_columns(df):
    """Return the names of columns holding numeric data (int/float, not bool)."""
    numeric = df.select_dtypes(include=["number"]).columns.tolist()
    # Booleans are numeric to pandas but are categories to a human reader.
    return [c for c in numeric if not pd.api.types.is_bool_dtype(df[c])]


def detect_datetime_columns(df):
    """
    Return the names of columns holding dates.

    Includes columns pandas already parsed as datetimes, plus text columns whose
    values are overwhelmingly parseable as dates (for example a "date" column
    read from a CSV as plain text). Plain numbers and ordinary words are never
    treated as dates.
    """
    found = df.select_dtypes(include=["datetime", "datetimetz"]).columns.tolist()
    for column in df.select_dtypes(include=["object"]).columns:
        values = df[column].dropna()
        values = values[values.map(lambda v: isinstance(v, str))]
        if values.empty:
            continue
        sample = values.head(50).str.strip()
        # Real dates contain digits; this avoids parsing ordinary words.
        sample = sample[sample.str.contains(r"\d", regex=True)]
        if sample.empty:
            continue
        # Pure numbers (e.g. "123") are numbers, not dates.
        if pd.to_numeric(sample.str.replace(r"[,$%\s]", "", regex=True),
                         errors="coerce").notna().mean() >= 0.5:
            continue
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            try:
                parsed = pd.to_datetime(sample, errors="coerce", format="mixed")
            except (TypeError, ValueError):
                parsed = pd.to_datetime(sample, errors="coerce")
        if parsed.notna().mean() >= 0.9:
            found.append(column)
    return found


def make_unique_columns(columns):
    """
    Make column names unique by appending " (2)", " (3)", ... to repeats.

    Returns (new_names, renamed) where new_names is a list the same length as
    `columns` and renamed is a list of (original, new) pairs for every column
    that had to be changed. The caller's DataFrame is not touched.
    """
    seen = {}
    new_names, renamed = [], []
    for column in columns:
        name = str(column) if not isinstance(column, str) else column
        if name in seen:
            seen[name] += 1
            candidate = f"{name} ({seen[name]})"
            while candidate in seen:
                seen[name] += 1
                candidate = f"{name} ({seen[name]})"
            seen[candidate] = 1
            new_names.append(candidate)
            renamed.append((name, candidate))
        else:
            seen[name] = 1
            new_names.append(column)
    return new_names, renamed
