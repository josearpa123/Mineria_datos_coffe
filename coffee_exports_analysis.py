#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Coffee Exports Analysis (Colombia) - Full Report (Trade volume-centric)

Instructions:
1) Place this script in the same folder as your Excel file (or adjust EXCEL_PATH).
2) Run:  python coffee_exports_analysis.py
3) Charts and CSV summaries will be written to ./outputs/

Notes:
- Uses only matplotlib for plots (no seaborn).
- Charts use default colors and one plot per figure.
"""

import os
import sys
import math
import argparse
import pandas as pd
import matplotlib.pyplot as plt

# -------------------------
# Config
# -------------------------
DEFAULT_EXCEL_PATH = "colombia_coffee_v1_0_3.xlsx"
OUTPUT_DIR = "outputs"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# -------------------------
# Helpers
# -------------------------
def read_year_sheets(excel_path: str) -> pd.DataFrame:
    """Read all sheets starting with 'Year ' and return a combined DataFrame with a 'Year' column."""
    xls = pd.ExcelFile(excel_path)
    year_sheets = [s for s in xls.sheet_names if s.lower().startswith("year")]
    frames = []
    for sheet in year_sheets:
        df = pd.read_excel(excel_path, sheet_name=sheet)
        # Standardize Year column
        year_label = sheet.replace("Year ", "").strip()
        df["Year"] = year_label
        frames.append(df)
    if not frames:
        raise ValueError("No 'Year ...' sheets found in the workbook.")
    df_all = pd.concat(frames, ignore_index=True)
    return df_all

def coerce_numeric(df: pd.DataFrame, cols):
    """Ensure numeric columns with coercion; non-convertible become NaN."""
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

def save_table(df: pd.DataFrame, name: str):
    """Save a CSV summary to outputs/"""
    out_path = os.path.join(OUTPUT_DIR, f"{name}.csv")
    df.to_csv(out_path, index=False, encoding="utf-8-sig")
    return out_path

def barh_plot(x_labels, values, title, xlabel, ylabel, filename):
    plt.figure(figsize=(10, 6))
    plt.barh(x_labels, values)
    plt.title(title)
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    # Show largest at top
    plt.gca().invert_yaxis()
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, filename), dpi=150)
    plt.close()

def bar_plot(x_labels, values, title, xlabel, ylabel, filename, rotation=45):
    plt.figure(figsize=(10, 6))
    plt.bar(x_labels, values)
    plt.title(title)
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.xticks(rotation=rotation)
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, filename), dpi=150)
    plt.close()

def pie_plot(labels, values, title, filename):
    plt.figure(figsize=(8, 8))
    # autopct default formatting
    plt.pie(values, labels=labels, autopct="%1.1f%%", startangle=140)
    plt.title(title)
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, filename), dpi=150)
    plt.close()

# -------------------------
# Main analytics
# -------------------------
def main(excel_path):
    # 1) Read and combine
    df_all = read_year_sheets(excel_path)

    # 2) Ensure numeric columns
    df_all = coerce_numeric(df_all, ["Trade volume", "Trade value"])

    # 3) Convert Trade volume to tons
    df_all["Trade volume (t)"] = df_all["Trade volume"] / 1000.0

    # Optional: drop rows with no volume
    df_all = df_all.dropna(subset=["Trade volume"])

    # Save base combined file
    save_table(df_all, "combined_all_years")

    # -------------------------
    # A) Ventas por año (volumen)
    # -------------------------
    per_year = (
        df_all.groupby("Year", as_index=False)["Trade volume"]
        .sum()
        .sort_values("Year")
    )
    save_table(per_year, "summary_trade_volume_by_year")
    bar_plot(
        per_year["Year"],
        per_year["Trade volume"],
        "Volumen exportado por año (kg)",
        "Año",
        "Trade volume (kg)",
        "A_trade_volume_by_year.png",
        rotation=45,
    )

    # -------------------------
    # B) Países de destino - % del total (acumulado todos los años)
    # -------------------------
    if "Country of destination" in df_all.columns:
        by_country = (
            df_all.groupby("Country of destination", as_index=False)["Trade volume"]
            .sum()
            .sort_values("Trade volume", ascending=False)
        )
        by_country["Porcentaje"] = by_country["Trade volume"] / by_country["Trade volume"].sum() * 100.0
        save_table(by_country, "summary_trade_volume_by_destination_country")
        # Pie top 10
        top10_c = by_country.head(10)
        pie_plot(
            top10_c["Country of destination"],
            top10_c["Trade volume"],
            "Top 10 países destino por volumen (%)",
            "B_destination_country_pie_top10.png",
        )

    # -------------------------
    # C) Exportadores principales (acumulado) y %
    # -------------------------
    if "Exporter" in df_all.columns:
        by_exporter = (
            df_all.groupby("Exporter", as_index=False)["Trade volume"]
            .sum()
            .sort_values("Trade volume", ascending=False)
        )
        by_exporter["Porcentaje"] = by_exporter["Trade volume"] / by_exporter["Trade volume"].sum() * 100.0
        save_table(by_exporter, "summary_trade_volume_by_exporter")
        top10_e = by_exporter.head(10)
        barh_plot(
            top10_e["Exporter"],
            top10_e["Trade volume"],
            "Top 10 exportadores por volumen (kg)",
            "Trade volume (kg)",
            "Exportador",
            "C_exporters_top10_barh.png",
        )

    # -------------------------
    # D) Importadores top 3 por año
    # -------------------------
    if "Importer" in df_all.columns:
        by_year_importer = (
            df_all.groupby(["Year", "Importer"], as_index=False)["Trade volume"]
            .sum()
            .sort_values(["Year", "Trade volume"], ascending=[True, False])
        )
        # top 3 por año
        top3_imp = by_year_importer.groupby("Year", as_index=False).head(3).reset_index(drop=True)
        save_table(top3_imp, "top3_importers_per_year")
        # Gráfico de barras agrupadas: 3 barras por año
        # Creamos una figura por año para mantener 1 gráfico por figura (requisito)
        years_sorted = sorted(top3_imp["Year"].unique(), key=lambda x: str(x))
        for y in years_sorted:
            temp = top3_imp[top3_imp["Year"] == y]
            bar_plot(
                temp["Importer"],
                temp["Trade volume"],
                f"Top 3 importadores en {y} (kg)",
                "Importer",
                "Trade volume (kg)",
                f"D_top3_importers_{y}.png",
                rotation=30,
            )

    # -------------------------
    # E) Países destino top 3 por año
    # -------------------------
    if "Country of destination" in df_all.columns:
        by_year_country = (
            df_all.groupby(["Year", "Country of destination"], as_index=False)["Trade volume"]
            .sum()
            .sort_values(["Year", "Trade volume"], ascending=[True, False])
        )
        top3_countries = by_year_country.groupby("Year", as_index=False).head(3).reset_index(drop=True)
        save_table(top3_countries, "top3_destination_countries_per_year")
        years_sorted = sorted(top3_countries["Year"].unique(), key=lambda x: str(x))
        for y in years_sorted:
            temp = top3_countries[top3_countries["Year"] == y]
            bar_plot(
                temp["Country of destination"],
                temp["Trade volume"],
                f"Top 3 países destino en {y} (kg)",
                "País de destino",
                "Trade volume (kg)",
                f"E_top3_countries_{y}.png",
                rotation=30,
            )

    # -------------------------
    # F) Coffee bean (acumulado) por volumen
    # -------------------------
    if "Coffee bean" in df_all.columns:
        beans = (
            df_all.groupby("Coffee bean", as_index=False)["Trade volume"]
            .sum()
            .sort_values("Trade volume", ascending=False)
        )
        beans["Porcentaje"] = beans["Trade volume"] / beans["Trade volume"].sum() * 100.0
        save_table(beans, "summary_trade_volume_by_coffee_bean")
        barh_plot(
            beans["Coffee bean"],
            beans["Trade volume"],
            "Tipos de café más exportados (por volumen, kg)",
            "Trade volume (kg)",
            "Tipo de grano",
            "F_beans_barh.png",
        )

        # F2) Coffee bean top por año (top1)
        beans_year = (
            df_all.groupby(["Year", "Coffee bean"], as_index=False)["Trade volume"]
            .sum()
            .sort_values(["Year", "Trade volume"], ascending=[True, False])
        )
        top1_bean_year = beans_year.groupby("Year", as_index=False).head(1).reset_index(drop=True)
        save_table(top1_bean_year, "top1_coffee_bean_per_year")
        # Una figura por año para mantener un gráfico por figura
        years_sorted = sorted(top1_bean_year["Year"].unique(), key=lambda x: str(x))
        for y in years_sorted:
            tmp = top1_bean_year[top1_bean_year["Year"] == y]
            bar_plot(
                tmp["Coffee bean"],
                tmp["Trade volume"],
                f"Grano más exportado en {y} (kg)",
                "Coffee bean",
                "Trade volume (kg)",
                f"F2_top1_bean_{y}.png",
                rotation=0,
            )

    # -------------------------
    # G) Municipios principales y exportadores por municipio
    # -------------------------
    if "Municipality of export" in df_all.columns:
        muni = (
            df_all.groupby("Municipality of export", as_index=False)["Trade volume"]
            .sum()
            .sort_values("Trade volume", ascending=False)
        )
        save_table(muni, "summary_trade_volume_by_municipality")
        top10_muni = muni.head(10)
        barh_plot(
            top10_muni["Municipality of export"],
            top10_muni["Trade volume"],
            "Top 10 municipios exportadores (kg)",
            "Trade volume (kg)",
            "Municipio",
            "G_municipalities_top10_barh.png",
        )

        # Exportadores por municipio (para top 5 municipios)
        if "Exporter" in df_all.columns:
            muni_exp = (
                df_all.groupby(["Municipality of export", "Exporter"], as_index=False)["Trade volume"]
                .sum()
                .sort_values(["Municipality of export", "Trade volume"], ascending=[True, False])
            )
            top5_names = top10_muni.head(5)["Municipality of export"].tolist()
            top_muni_exp = muni_exp[muni_exp["Municipality of export"].isin(top5_names)]
            save_table(top_muni_exp, "exporters_in_top5_municipalities")

            # Generar una figura por municipio top con los 5 exportadores principales
            for m in top5_names:
                tmp = top_muni_exp[top_muni_exp["Municipality of export"] == m].head(5)
                barh_plot(
                    tmp["Exporter"],
                    tmp["Trade volume"],
                    f"Principales exportadores en {m} (kg)",
                    "Trade volume (kg)",
                    "Exportador",
                    f"G2_exporters_top_{m.replace(' ', '_')}.png",
                )

    print(f"Listo. Archivos generados en: {OUTPUT_DIR}")
    return 0

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Coffee Exports Analysis - Full Report (Trade volume)")
    parser.add_argument("--excel", type=str, default=DEFAULT_EXCEL_PATH,
                        help="Ruta al archivo Excel (por defecto: colombia_coffee_v1_0_3.xlsx)")
    args = parser.parse_args()
    sys.exit(main(args.excel))
