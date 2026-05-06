from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

from analytics_store import load_analytics_events, resolve_analytics_storage_dir


st.set_page_config(page_title="Allokering Analytics", layout="wide")


def _events_to_dataframe(storage_dir: Path) -> pd.DataFrame:
    rows: list[dict] = []
    for record in load_analytics_events(storage_dir):
        properties = record.get("properties")
        row = {
            "event": record.get("event", ""),
            "timestamp": record.get("timestamp"),
            "source_file": record.get("_source_file", ""),
        }
        if isinstance(properties, dict):
            row.update(properties)
        rows.append(row)

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    df["timestamp"] = pd.to_datetime(df["timestamp"], errors="coerce", utc=True)
    if "session_seconds" in df.columns:
        df["session_seconds"] = pd.to_numeric(df["session_seconds"], errors="coerce")
    else:
        df["session_seconds"] = pd.Series(pd.NA, index=df.index, dtype="Float64")

    for column in ("event", "feature", "action", "file_type", "install_id", "app_version"):
        if column not in df.columns:
            df[column] = ""
        df[column] = df[column].fillna("").astype(str)
    return df.sort_values("timestamp", ascending=False).reset_index(drop=True)


@st.cache_data(ttl=30)
def _load_dataframe(storage_dir_text: str) -> pd.DataFrame:
    return _events_to_dataframe(Path(storage_dir_text))


def _text_series(df: pd.DataFrame, column: str) -> pd.Series:
    if column in df.columns:
        return df[column].fillna("").astype(str)
    return pd.Series("", index=df.index, dtype="string")


def _safe_minutes_from_seconds(series: pd.Series) -> float:
    values = pd.to_numeric(series, errors="coerce").dropna()
    if values.empty:
        return 0.0
    return round(float(values.mean()) / 60, 1)


def _user_summary_table(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "install_id" not in df.columns:
        return pd.DataFrame()

    uploads = (
        df.loc[df["event"].eq("input_selected")]
        .groupby("install_id")
        .size()
        .rename("uploads")
    )
    feature_runs = (
        df.loc[df["event"].eq("feature_usage") & _text_series(df, "action").eq("run_completed")]
        .groupby("install_id")
        .size()
        .rename("feature_runs")
    )
    closed_df = df.loc[df["event"].eq("app_closed")].copy()
    if "session_seconds" not in closed_df.columns:
        closed_df["session_seconds"] = pd.NA
    sessions = closed_df.groupby("install_id").agg(
        sessions=("event", "size"),
        avg_session_minutes=("session_seconds", lambda s: round((s.dropna().mean() or 0) / 60, 1)),
        total_session_hours=("session_seconds", lambda s: round(s.dropna().sum() / 3600, 2)),
    )
    seen = (
        df.groupby("install_id")
        .agg(
            first_seen=("timestamp", "min"),
            last_seen=("timestamp", "max"),
        )
    )

    summary = pd.concat([uploads, feature_runs, sessions, seen], axis=1).fillna(0)
    if "first_seen" in summary.columns:
        summary["first_seen"] = pd.to_datetime(summary["first_seen"], errors="coerce").dt.strftime("%Y-%m-%d %H:%M")
    if "last_seen" in summary.columns:
        summary["last_seen"] = pd.to_datetime(summary["last_seen"], errors="coerce").dt.strftime("%Y-%m-%d %H:%M")
    return summary.reset_index().rename(columns={"install_id": "user"})


def _render_chart_table(title: str, df: pd.DataFrame, column_name: str) -> None:
    st.subheader(title)
    if df.empty:
        st.info("Ingen data an.")
        return
    chart_df = df.set_index(column_name)
    st.bar_chart(chart_df)
    st.dataframe(df, width="stretch", hide_index=True)


default_storage_dir = resolve_analytics_storage_dir()
st.title("Allokering Analytics")
st.caption(
    "Dashboarden laser lokala analyticsfiler. Om du vill samla flera anvandare utan server kan du senare peka appen mot en delad mapp."
)

with st.sidebar:
    st.header("Installning")
    storage_dir_text = st.text_input("Analytics-mapp", value=str(default_storage_dir))
    if st.button("Ladda om data"):
        st.cache_data.clear()

storage_dir = Path(storage_dir_text).expanduser()
events_df = _load_dataframe(str(storage_dir))

if events_df.empty:
    st.info("Inga analytics-events hittades an.")
    st.code(str(storage_dir))
    st.stop()

min_ts = events_df["timestamp"].dropna().min()
max_ts = events_df["timestamp"].dropna().max()
if pd.notna(min_ts) and pd.notna(max_ts):
    with st.sidebar:
        date_range = st.date_input(
            "Datumintervall",
            value=(min_ts.date(), max_ts.date()),
            min_value=min_ts.date(),
            max_value=max_ts.date(),
        )
else:
    date_range = ()

filtered_df = events_df.copy()
if isinstance(date_range, (tuple, list)) and len(date_range) == 2:
    start_date, end_date = date_range
    mask = filtered_df["timestamp"].dt.date.between(start_date, end_date)
    filtered_df = filtered_df.loc[mask].copy()

unique_users = int(filtered_df["install_id"].replace("", pd.NA).dropna().nunique()) if "install_id" in filtered_df.columns else 0
upload_count = int(filtered_df["event"].eq("input_selected").sum())
feature_runs_count = int(
    (
        filtered_df["event"].eq("feature_usage")
        & _text_series(filtered_df, "action").eq("run_completed")
    ).sum()
)
session_df = filtered_df.loc[filtered_df["event"].eq("app_closed")].copy()
avg_session_minutes = _safe_minutes_from_seconds(session_df["session_seconds"])

metric_columns = st.columns(4)
metric_columns[0].metric("Unika anvandare", unique_users)
metric_columns[1].metric("Filuppladdningar", upload_count)
metric_columns[2].metric("Feature-korningar", feature_runs_count)
metric_columns[3].metric("Snitt oppen tid (min)", avg_session_minutes)

feature_chart_df = (
    filtered_df.loc[
        filtered_df["event"].eq("feature_usage")
        & _text_series(filtered_df, "action").eq("run_completed")
        & _text_series(filtered_df, "feature").ne("")
    ]
    .groupby("feature")
    .size()
    .reset_index(name="runs")
    .sort_values("runs", ascending=False)
)

upload_chart_df = (
    filtered_df.loc[
        filtered_df["event"].eq("input_selected")
        & _text_series(filtered_df, "file_type").ne("")
    ]
    .groupby("file_type")
    .size()
    .reset_index(name="uploads")
    .sort_values("uploads", ascending=False)
)

left_column, right_column = st.columns(2)
with left_column:
    _render_chart_table("Popul araste funktioner", feature_chart_df, "feature")
with right_column:
    _render_chart_table("Mest uppladdade filer", upload_chart_df, "file_type")

st.subheader("Anvandare")
user_summary = _user_summary_table(filtered_df)
if user_summary.empty:
    st.info("Ingen anvandardata an.")
else:
    st.dataframe(user_summary, width="stretch", hide_index=True)

uploads_per_user = (
    filtered_df.loc[
        filtered_df["event"].eq("input_selected")
        & _text_series(filtered_df, "install_id").ne("")
        & _text_series(filtered_df, "file_type").ne("")
    ]
    .pivot_table(index="install_id", columns="file_type", values="event", aggfunc="count", fill_value=0)
    .reset_index()
)

st.subheader("Filuppladdningar per anvandare")
if uploads_per_user.empty:
    st.info("Ingen filuppladdningsdata an.")
else:
    st.dataframe(uploads_per_user, width="stretch", hide_index=True)

st.subheader("Senaste events")
recent_columns = [column for column in ["timestamp", "event", "install_id", "feature", "action", "file_type", "session_seconds"] if column in filtered_df.columns]
st.dataframe(filtered_df[recent_columns].head(200), width="stretch", hide_index=True)
