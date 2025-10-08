# TimeScheduleGenerator.py
import pandas as pd
import plotly.graph_objects as go
import re
from collections import defaultdict

# Replace this with your OneDrive direct download link
EXCEL_PATH = "https://quattrongmbh.sharepoint.com/sites/GB1-ProjekteIL/_layouts/15/download.aspx?SourceUrl=https://quattrongmbh.sharepoint.com/sites/GB1-ProjekteIL/Freigegebene%20Dokumente/226-141_Increase_to_13tph/Bearbeitung/Rostering/TrainScheduleApp/Tankenliste.xlsm?d=we952f7a148d04a348509142b04ffc1a3"

SHEET_LAUF = "Laufwege"
SHEET_VERKN = "IN_Verknüpfungen"
SHEET_ZUG = "IN_Zugliste"

START_ROW = 2
COL_FAHRTNAME = 0
COL_START = 6
COL_DISTANCE = 2

def parse_duration(val):
    if pd.isna(val) or str(val).strip() == "":
        return pd.Timedelta(0)
    try:
        return pd.to_timedelta(str(val))
    except:
        return pd.Timedelta(0)

def extract_number(s):
    m = re.search(r'\d+', str(s))
    return int(m.group()) if m else float('inf')

def clean_column_name(name):
    return str(name).strip().replace('\u2011', '-')

_station_re = re.compile(r'^[A-Z]{2,}$')
def extract_stations(train_name):
    parts = str(train_name).split("_")
    stations = [p for p in parts if _station_re.match(p)]
    if len(stations) >= 2: return stations[0], stations[-1]
    if len(stations) == 1: return stations[0], stations[0]
    fallback = [re.sub(r'[^A-Z]', '', p).strip() for p in parts if re.search(r'[A-Z]', p)]
    fallback = [p for p in fallback if len(p) >= 2]
    if len(fallback) >= 2: return fallback[0], fallback[-1]
    if len(fallback) == 1: return fallback[0], fallback[0]
    return "", ""

def color_for_zugklasse(z):
    if not z: return "#7f7f7f"
    z = str(z).strip()
    if z == "NRz": return "#2171b5"
    if z == "Lz": return "#238b45"
    return "#525252"

def generate_schedule_html():
    """Generates the Gantt chart and composition list as an HTML string"""
    lauf = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_LAUF, header=START_ROW, engine="openpyxl")
    verkn = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_VERKN, header=2, engine="openpyxl")
    zugliste = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_ZUG, header=0, engine="openpyxl")

    verkn.columns = [clean_column_name(c) for c in verkn.columns]
    zugliste.columns = [clean_column_name(c) for c in zugliste.columns]

    NAME_COL = "Erste Zugfahrt"
    COMP_COL = "Fzg_Index"
    NUMBER_COL = "n-te Fahrt"

    zugklasse_map = {
        str(r[0]).strip(): str(r[11]).strip() if pd.notna(r[11]) else ""
        for r in zugliste.itertuples(index=False)
    }

    lookup = {
        (str(comp).strip(), int(num)): str(name).strip()
        for comp, num, name in zip(verkn[COMP_COL], verkn[NUMBER_COL], verkn[NAME_COL])
    }

    lauf.iloc[:, COL_START] = pd.to_datetime(
        lauf.iloc[:, COL_START].astype(str),
        format="%H:%M:%S",
        errors="coerce"
    )

    composition_distances = {
        str(r.iloc[COL_FAHRTNAME]).strip(): int(round(float(r.iloc[COL_DISTANCE])/1000.0, 0))
        if pd.notna(r.iloc[COL_DISTANCE]) else 0
        for _, r in lauf.iterrows() if str(r.iloc[COL_FAHRTNAME]).strip()
    }

    all_cols_after_G = list(range(COL_START + 1, lauf.shape[1]))
    balken = []

    for _, r in lauf.iterrows():
        comp = str(r.iloc[COL_FAHRTNAME]).strip()
        if not comp: continue
        start_time = r.iloc[COL_START]
        if pd.isna(start_time): continue

        t_cursor = start_time
        for i, col_idx in enumerate(all_cols_after_G):
            dur = parse_duration(r.iloc[col_idx])
            if dur <= pd.Timedelta(0): continue
            if i % 2 == 0:
                seg_start = t_cursor
                seg_end = t_cursor + dur
                trip_number = 2 + i // 2
                train_name = lookup.get((comp, trip_number), f"Fahrt {trip_number}")
                station_left, station_right = extract_stations(train_name)
                zugklasse = zugklasse_map.get(train_name, "")
                balken.append({
                    "Start": seg_start, "Ende": seg_end, "Y": comp,
                    "Train": train_name, "StationLeft": station_left,
                    "StationRight": station_right, "Nummer": trip_number,
                    "Zugklasse": zugklasse
                })
                t_cursor = seg_end
            else:
                t_cursor += dur

    unique_comps = sorted(list({b["Y"] for b in balken}), key=extract_number)
    y_axis_labels = [f"{comp} ({composition_distances.get(comp,0)} km)" for comp in unique_comps]
    comp_to_label = {comp: lbl for comp, lbl in zip(unique_comps, y_axis_labels)}

    bars_by_color = defaultdict(lambda: {"x": [], "y": [], "text": []})
    annotations = []

    for b in balken:
        comp_label = comp_to_label[b["Y"]]
        color = color_for_zugklasse(b["Zugklasse"])

        bars_by_color[color]["x"].extend([b["Start"], b["Ende"], None])
        bars_by_color[color]["y"].extend([comp_label, comp_label, None])
        bars_by_color[color]["text"].extend([b["Train"], b["Train"], None])

        if b["StationLeft"]:
            annotations.append(dict(
                x=b["Start"] + pd.to_timedelta('00:00:30'),
                y=comp_label,
                text=b["StationLeft"],
                showarrow=False,
                font=dict(size=8, color="white"),
                xanchor="left",
                yanchor="middle"
            ))
        if b["StationRight"]:
            annotations.append(dict(
                x=b["Ende"] - pd.to_timedelta('00:00:30'),
                y=comp_label,
                text=b["StationRight"],
                showarrow=False,
                font=dict(size=8, color="white"),
                xanchor="right",
                yanchor="middle"
            ))

        mid = b["Start"] + (b["Ende"] - b["Start"]) / 2
        annotations.append(dict(
            x=mid,
            y=comp_label,
            text=b["Train"],
            showarrow=False,
            font=dict(size=9, color="navy"),
            yshift=12
        ))

    fig = go.Figure()
    for color, segs in bars_by_color.items():
        fig.add_trace(go.Scatter(
            x=segs["x"],
            y=segs["y"],
            mode="lines",
            line=dict(width=13, color=color),
            text=segs["text"],
            hoverinfo="x",
            showlegend=False
        ))

    n_comps = len(y_axis_labels)
    fig_height = max(600, n_comps * 30)
    fig.update_layout(
        title=dict(text="<b>Schedule of the Trains</b>", font=dict(size=22)),
        xaxis_title="Time",
        yaxis_title="Compositions",
        yaxis=dict(categoryorder="array", categoryarray=y_axis_labels, tickfont=dict(size=12)),
        showlegend=False,
        hovermode="closest",
        height=fig_height,
        margin=dict(l=180, r=40, t=80, b=80),
        annotations=annotations
    )
    fig.update_xaxes(tickformat="%H:%M:%S", showgrid=True, gridwidth=1, gridcolor="lightgray")

    composition_lists = {}
    for comp, num, nm in zip(verkn[COMP_COL], verkn[NUMBER_COL], verkn[NAME_COL]):
        composition_lists.setdefault(str(comp).strip(), []).append((int(num), str(nm).strip()))
    for comp in composition_lists:
        composition_lists[comp] = [name for _, name in sorted(composition_lists[comp], key=lambda x: x[0])]

    # generate HTML with colored1 at the end
    html_filename = "time_schedule_gantt_colored1.html"
    fig.write_html(html_filename, include_plotlyjs='cdn', full_html=True)

    composition_html = '<h2 style="font-size:20px;"><b>Trains in the Compositions</b></h2>\n'
    for comp in sorted(composition_lists.keys(), key=extract_number):
        dist_km = composition_distances.get(comp, 0)
        composition_html += f"<b>{comp} ({dist_km} km):</b> "
        composition_html += " -> ".join(composition_lists[comp])
        composition_html += "<br>\n"

    with open(html_filename, "a", encoding="utf-8") as f:
        f.write(composition_html)

    return html_filename


