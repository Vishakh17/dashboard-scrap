from __future__ import annotations

import json
import os
import re
from pathlib import Path

import pandas as pd
import plotly.express as px
from flask import Flask, Response, redirect, render_template, request, url_for
from plotly.offline import get_plotlyjs
from werkzeug.utils import secure_filename


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = BASE_DIR / "uploads"
UPLOAD_FOLDER.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {".xls", ".xlsx"}
PREFERRED_SHEETS = {"26-27", "2026-27", "data", "scrap data"}

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.secret_key = os.environ.get("SCRAP_DASHBOARD_SECRET", "scrap-dashboard-local")


def allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def clean_label(value: object) -> str:
    text = str(value or "").strip()
    text = re.sub(r"\s+", " ", text)
    return text


def normalized_label(value: object) -> str:
    text = clean_label(value).lower()
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def header_score(row: pd.Series) -> int:
    text = " ".join(normalized_label(value) for value in row.dropna().tolist())
    score = 0
    for token in ("date", "customer", "kg", "dept", "department", "flash", "rejection", "remark", "total"):
        if token in text:
            score += 1
    return score


def detect_header_row(path: Path, sheet_name: str) -> int:
    preview = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=12)
    scores = [(index, header_score(row)) for index, row in preview.iterrows()]
    best_index, best_score = max(scores, key=lambda item: item[1])
    return int(best_index if best_score >= 3 else 0)


def choose_sheet(path: Path) -> tuple[str, int]:
    workbook = pd.ExcelFile(path)
    for sheet in workbook.sheet_names:
        if normalized_label(sheet) in PREFERRED_SHEETS:
            return sheet, detect_header_row(path, sheet)

    scored_sheets = []
    for sheet in workbook.sheet_names:
        header_row = detect_header_row(path, sheet)
        preview = pd.read_excel(path, sheet_name=sheet, header=None, nrows=header_row + 1)
        scored_sheets.append((sheet, header_row, header_score(preview.iloc[header_row])))

    sheet, header_row, _ = max(scored_sheets, key=lambda item: item[2])
    return sheet, header_row


def canonical_column_name(column: object) -> str:
    label = normalized_label(column)
    if label == "date" or label.startswith("date "):
        return "Date"
    if "customer" in label or "material" in label:
        return "Customer Name"
    if label in {"kg", "kgs", "weight"} or ("weight" in label and "total" not in label):
        return "KG"
    if "dept" in label or "department" in label:
        return "Dept"
    if "rejection" in label or "flash" in label or label in {"type", "type of scrap"}:
        return "Type"
    if "remark" in label or "note" in label:
        return "Remark"
    if "total" in label and "weight" in label:
        return "Total Weight"
    return clean_label(column)


def load_scrap_data(path: Path) -> tuple[pd.DataFrame, dict[str, object]]:
    sheet_name, header_row = choose_sheet(path)
    df = pd.read_excel(path, sheet_name=sheet_name, header=header_row)
    df = df.dropna(how="all")
    df = df.loc[:, ~pd.Index(df.columns).astype(str).str.match(r"^Unnamed")]
    df.columns = [canonical_column_name(column) for column in df.columns]

    if "Type" not in df.columns and "Customer Name" in df.columns:
        df["Type"] = df["Customer Name"]

    required = {"Date", "Customer Name", "KG", "Dept", "Type"}
    missing = sorted(required.difference(df.columns))
    if missing:
        raise ValueError(f"Could not find these required columns: {', '.join(missing)}")

    keep_columns = ["Date", "Customer Name", "KG", "Dept", "Type", "Remark", "Total Weight"]
    available_columns = [column for column in keep_columns if column in df.columns]
    df = df[available_columns].copy()

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["KG"] = pd.to_numeric(df["KG"], errors="coerce")
    df = df[df["Date"].notna() & df["KG"].notna()]

    for column in ("Customer Name", "Dept", "Type"):
        df[column] = df[column].fillna("Unknown").astype(str).str.strip()
        df.loc[df[column].isin(["", "nan", "NaT"]), column] = "Unknown"

    if "Remark" in df.columns:
        df["Remark"] = df["Remark"].fillna("").astype(str).str.strip()
    else:
        df["Remark"] = ""

    if "Total Weight" not in df.columns:
        df["Total Weight"] = None

    df = df.sort_values("Date")
    df["Day"] = df["Date"].dt.date
    df["Month"] = df["Date"].dt.to_period("M").dt.to_timestamp()
    df["Month Label"] = df["Month"].dt.strftime("%b %Y")

    metadata = {
        "sheet_name": sheet_name,
        "header_row": header_row + 1,
        "source_rows": len(df),
    }
    return df, metadata


def latest_workbook() -> Path | None:
    files = [
        file
        for file in UPLOAD_FOLDER.iterdir()
        if file.is_file() and file.suffix.lower() in ALLOWED_EXTENSIONS
    ]
    if not files:
        return None
    return max(files, key=lambda file: file.stat().st_mtime)


def format_number(value: float, decimals: int = 2) -> str:
    return f"{value:,.{decimals}f}"


def filter_options(df: pd.DataFrame) -> dict[str, list[str]]:
    return {
        "dept": sorted(df["Dept"].dropna().unique().tolist()),
        "type": sorted(df["Type"].dropna().unique().tolist()),
        "month": sorted(df["Month Label"].dropna().unique().tolist(), key=lambda x: pd.to_datetime(x)),
    }


def apply_filters(df: pd.DataFrame, filters: dict[str, str]) -> pd.DataFrame:
    filtered = df.copy()
    if filters["dept"] != "All":
        filtered = filtered[filtered["Dept"] == filters["dept"]]
    if filters["type"] != "All":
        filtered = filtered[filtered["Type"] == filters["type"]]
    if filters["month"] != "All":
        filtered = filtered[filtered["Month Label"] == filters["month"]]
    return filtered


def style_chart(fig):
    fig.update_layout(
        template="plotly_white",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter, Segoe UI, Arial, sans-serif", "color": "#18202f"},
        margin={"l": 36, "r": 20, "t": 54, "b": 42},
        colorway=["#0f9f7a", "#4f46e5", "#f59e0b", "#dc2626", "#64748b", "#0891b2"],
        legend={"orientation": "h", "yanchor": "bottom", "y": 1.02, "xanchor": "right", "x": 1},
    )
    fig.update_xaxes(showgrid=False, linecolor="#d7dde8")
    fig.update_yaxes(gridcolor="#e8edf5", zerolinecolor="#d7dde8")
    return json.loads(fig.to_json())


def empty_chart(title: str) -> dict[str, object]:
    fig = px.scatter(title=title)
    fig.add_annotation(text="No records match the selected filters", showarrow=False, x=0.5, y=0.5)
    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)
    return style_chart(fig)


def build_charts(df: pd.DataFrame) -> dict[str, dict[str, object]]:
    if df.empty:
        return {
            "daily": empty_chart("Daily Scrap Trend"),
            "dept": empty_chart("Department Wise Scrap"),
            "type": empty_chart("Scrap Type Split"),
            "material": empty_chart("Top Scrap Materials"),
            "month": empty_chart("Monthly Scrap by Type"),
        }

    daily = df.groupby("Day", as_index=False)["KG"].sum()
    fig_daily = px.area(daily, x="Day", y="KG", title="Daily Scrap Trend", markers=True)
    fig_daily.update_traces(line={"width": 3}, fillcolor="rgba(15, 159, 122, 0.18)")

    dept = df.groupby("Dept", as_index=False)["KG"].sum().sort_values("KG", ascending=True)
    fig_dept = px.bar(dept, x="KG", y="Dept", orientation="h", title="Department Wise Scrap", text="KG")
    fig_dept.update_traces(marker_color="#4f46e5", texttemplate="%{text:,.1f}", textposition="outside")

    type_df = df.groupby("Type", as_index=False)["KG"].sum().sort_values("KG", ascending=False)
    fig_type = px.pie(type_df, names="Type", values="KG", title="Scrap Type Split", hole=0.52)
    fig_type.update_traces(textposition="inside", textinfo="percent+label")

    material = (
        df.groupby("Customer Name", as_index=False)["KG"]
        .sum()
        .sort_values("KG", ascending=False)
        .head(10)
        .sort_values("KG", ascending=True)
    )
    fig_material = px.bar(
        material,
        x="KG",
        y="Customer Name",
        orientation="h",
        title="Top Scrap Materials",
        text="KG",
    )
    fig_material.update_traces(marker_color="#0f9f7a", texttemplate="%{text:,.1f}", textposition="outside")

    month_type = df.groupby(["Month", "Month Label", "Type"], as_index=False)["KG"].sum()
    month_type = month_type.sort_values("Month")
    fig_month = px.bar(
        month_type,
        x="Month Label",
        y="KG",
        color="Type",
        title="Monthly Scrap by Type",
        barmode="stack",
    )

    return {
        "daily": style_chart(fig_daily),
        "dept": style_chart(fig_dept),
        "type": style_chart(fig_type),
        "material": style_chart(fig_material),
        "month": style_chart(fig_month),
    }


def build_kpis(df: pd.DataFrame) -> list[dict[str, str]]:
    if df.empty:
        return [
            {"label": "Total Scrap", "value": "0.00 KG"},
            {"label": "Flash", "value": "0.00 KG"},
            {"label": "Rejection", "value": "0.00 KG"},
            {"label": "Entries", "value": "0"},
        ]

    flash = df[df["Type"].str.contains("flash", case=False, na=False)]["KG"].sum()
    rejection = df[df["Type"].str.contains("rejection", case=False, na=False)]["KG"].sum()
    avg_daily = df.groupby("Day")["KG"].sum().mean()

    return [
        {"label": "Total Scrap", "value": f"{format_number(df['KG'].sum())} KG"},
        {"label": "Flash", "value": f"{format_number(flash)} KG"},
        {"label": "Rejection", "value": f"{format_number(rejection)} KG"},
        {"label": "Avg / Day", "value": f"{format_number(avg_daily)} KG"},
        {"label": "Entries", "value": f"{len(df):,}"},
    ]


def recent_records(df: pd.DataFrame) -> list[dict[str, object]]:
    table = df.sort_values("Date", ascending=False).head(50).copy()
    table["Date"] = table["Date"].dt.strftime("%d %b %Y")
    table["KG"] = table["KG"].map(lambda value: format_number(value))
    table["Total Weight"] = table["Total Weight"].map(
        lambda value: "" if pd.isna(value) else format_number(float(value))
    )
    return table[["Date", "Customer Name", "KG", "Dept", "Type", "Remark", "Total Weight"]].to_dict("records")


@app.route("/plotly.js")
def plotly_js():
    return Response(get_plotlyjs(), mimetype="application/javascript")


@app.route("/", methods=["GET", "POST"])
def dashboard():
    error = None

    if request.method == "POST":
        uploaded = request.files.get("file")
        if not uploaded or uploaded.filename == "":
            return redirect(url_for("dashboard"))
        if not allowed_file(uploaded.filename):
            return render_template("index.html", uploaded=False, error="Please upload an Excel workbook.")

        filename = secure_filename(uploaded.filename)
        path = UPLOAD_FOLDER / filename
        uploaded.save(path)
        return redirect(url_for("dashboard", file=filename))

    selected_name = request.args.get("file")
    selected_file = UPLOAD_FOLDER / selected_name if selected_name else latest_workbook()

    if not selected_file or not selected_file.exists():
        return render_template("index.html", uploaded=False, error=error)

    try:
        raw_df, metadata = load_scrap_data(selected_file)
        options = filter_options(raw_df)
        filters = {
            "dept": request.args.get("dept", "All"),
            "type": request.args.get("type", "All"),
            "month": request.args.get("month", "All"),
        }
        filtered_df = apply_filters(raw_df, filters)
        charts = build_charts(filtered_df)
        kpis = build_kpis(filtered_df)

        date_range = "No dates"
        if not raw_df.empty:
            date_range = f"{raw_df['Date'].min():%d %b %Y} to {raw_df['Date'].max():%d %b %Y}"

        return render_template(
            "index.html",
            uploaded=True,
            error=None,
            file_name=selected_file.name,
            metadata=metadata,
            date_range=date_range,
            options=options,
            filters=filters,
            kpis=kpis,
            charts=charts,
            row_count=len(raw_df),
            filtered_count=len(filtered_df),
            table_data=recent_records(filtered_df),
        )
    except Exception as exc:
        error = f"Could not build the dashboard: {exc}"
        return render_template("index.html", uploaded=False, error=error)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8502, debug=False, use_reloader=False)
