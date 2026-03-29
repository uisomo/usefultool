#!/usr/bin/env python3
"""
HHI vs Loss Rate - Excel Formula-Based Monte Carlo Generator

Generates an Excel file where ALL calculations use Excel formulas:
  - User edits obligor weights in the Weights sheet
  - HHI is auto-computed: =SUMPRODUCT(weights^2)
  - Defaults are simulated: =IF(RAND()<PD, 1, 0)
  - Loss stats use AVERAGE, STDEV, PERCENTILE formulas
  - Press F9 to re-simulate

Conditions: correlation=0, PD=AA (0.03%), LGD=100%, 10 obligors
"""

import openpyxl
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Configuration ────────────────────────────────────────────────────────
N_OBLIGORS = 10
N_SIMS = 500          # trials (keeps Excel responsive)
N_SCENARIOS = 6       # different weight distributions to compare
DEFAULT_PD = 0.0003   # 0.03% (AA)
DEFAULT_LGD = 1.0     # 100%

# ── Styles ───────────────────────────────────────────────────────────────
HEADER_FONT = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
TITLE_FONT = Font(name="Calibri", bold=True, size=14, color="2F5496")
SUBTITLE_FONT = Font(name="Calibri", bold=True, size=11, color="2F5496")
PARAM_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
INPUT_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
GOOD_FILL = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
WARN_FILL = PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)


def style_header_row(ws, row, max_col):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = THIN_BORDER


def style_cell(cell, num_format=None):
    cell.border = THIN_BORDER
    cell.alignment = Alignment(horizontal="center")
    if num_format:
        cell.number_format = num_format


# ── Predefined weight scenarios ──────────────────────────────────────────
SCENARIOS = [
    ("均等配分",        [10, 10, 10, 10, 10, 10, 10, 10, 10, 10]),
    ("やや集中",       [20, 15, 12, 11, 10,  8,  8,  7,  5,  4]),
    ("中程度集中",     [30, 20, 15, 10,  8,  5,  5,  3,  2,  2]),
    ("高集中",         [50, 15, 10,  7,  5,  4,  3,  3,  2,  1]),
    ("極端な集中",     [70, 10,  5,  4,  3,  2,  2,  2,  1,  1]),
    ("1社集中",        [91,  1,  1,  1,  1,  1,  1,  1,  1,  1]),
]


def create_parameters_sheet(wb):
    """Sheet 1: Editable parameters (PD, LGD, N_SIMS)."""
    ws = wb.active
    ws.title = "Parameters"
    ws.sheet_properties.tabColor = "2F5496"

    ws["A1"] = "HHI vs 損失率 Monte Carlo分析（数式版）"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:D1")

    ws["A3"] = "パラメータ"
    ws["B3"] = "値"
    ws["C3"] = "説明"
    style_header_row(ws, 3, 3)

    params = [
        ("PD (デフォルト確率)", DEFAULT_PD, "0.00%", "AA格の年間PD（編集可）"),
        ("LGD (損失率)", DEFAULT_LGD, "0%", "デフォルト時の損失割合（編集可）"),
        ("シミュレーション回数", N_SIMS, "#,##0", f"試行回数（{N_SIMS}列）"),
        ("債務者数", N_OBLIGORS, "#,##0", "ポートフォリオの企業数"),
    ]
    for i, (name, val, fmt, desc) in enumerate(params):
        r = 4 + i
        ws.cell(row=r, column=1, value=name).fill = PARAM_FILL
        c = ws.cell(row=r, column=2, value=val)
        c.number_format = fmt
        c.fill = INPUT_FILL
        ws.cell(row=r, column=3, value=desc)
        for col in range(1, 4):
            ws.cell(row=r, column=col).border = THIN_BORDER

    ws["A9"] = "使い方"
    ws["A9"].font = SUBTITLE_FONT
    instructions = [
        "1. Weightsシートで各シナリオの債務者ウェイトを編集",
        "2. PD・LGDはこのシートのB4・B5で変更可能",
        "3. F9キー（またはCtrl+Alt+F9）で再シミュレーション",
        "4. Resultsシートで各シナリオのHHI・損失指標・チャートを確認",
        "5. 全ての計算はExcel数式 — VBAなし",
    ]
    for i, text in enumerate(instructions):
        ws.cell(row=10 + i, column=1, value=text)

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 40


def create_weights_sheet(wb):
    """Sheet 2: Editable weight distributions for each scenario."""
    ws = wb.create_sheet("Weights")
    ws.sheet_properties.tabColor = "7030A0"

    ws["A1"] = "債務者ウェイト配分（各シナリオ）"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:M1")

    ws["A2"] = "※黄色セルは編集可能。ウェイトの合計は100%になるようにしてください。"
    ws["A2"].font = Font(italic=True, color="FF0000")

    # Headers
    headers = ["シナリオ"] + [f"債務者{i+1}" for i in range(N_OBLIGORS)] + ["合計", "HHI"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=4, column=c, value=h)
    style_header_row(ws, 4, len(headers))

    for s_idx, (name, pcts) in enumerate(SCENARIOS):
        r = 5 + s_idx
        ws.cell(row=r, column=1, value=name).fill = PARAM_FILL
        ws.cell(row=r, column=1).border = THIN_BORDER

        for j, pct in enumerate(pcts):
            c = ws.cell(row=r, column=2 + j, value=pct / 100)
            c.number_format = "0.0%"
            c.fill = INPUT_FILL
            c.border = THIN_BORDER
            c.alignment = Alignment(horizontal="center")

        # Sum formula (should = 100%)
        sum_col = N_OBLIGORS + 2  # column L
        sum_cell = ws.cell(row=r, column=sum_col)
        first_col = get_column_letter(2)
        last_col = get_column_letter(N_OBLIGORS + 1)
        sum_cell.value = f"=SUM({first_col}{r}:{last_col}{r})"
        sum_cell.number_format = "0.0%"
        sum_cell.border = THIN_BORDER
        sum_cell.alignment = Alignment(horizontal="center")

        # HHI formula = SUMPRODUCT(weights^2)
        hhi_col = N_OBLIGORS + 3  # column M
        hhi_cell = ws.cell(row=r, column=hhi_col)
        hhi_cell.value = f"=SUMPRODUCT({first_col}{r}:{last_col}{r},{first_col}{r}:{last_col}{r})"
        hhi_cell.number_format = "0.0000"
        hhi_cell.border = THIN_BORDER
        hhi_cell.alignment = Alignment(horizontal="center")

    ws.column_dimensions["A"].width = 16
    for c in range(2, N_OBLIGORS + 4):
        ws.column_dimensions[get_column_letter(c)].width = 10


def create_simulation_sheet(wb, scenario_idx):
    """Create a simulation sheet for one scenario.

    Each cell: =IF(RAND() < PD, 1, 0)  (independent Bernoulli)
    Bottom row: portfolio loss = SUM(default_i * weight_i * LGD)
    """
    name = SCENARIOS[scenario_idx][0]
    ws = wb.create_sheet(f"Sim_{scenario_idx+1}_{name}")
    ws.sheet_properties.tabColor = "C55A11"

    s_row = 5 + scenario_idx  # row in Weights sheet for this scenario

    # Column A: Obligor labels
    ws.cell(row=1, column=1, value="債務者 \\ Trial")
    ws.cell(row=1, column=1).font = HEADER_FONT
    ws.cell(row=1, column=1).fill = HEADER_FILL

    for i in range(N_OBLIGORS):
        c = ws.cell(row=i + 2, column=1, value=f"債務者{i+1}")
        c.font = Font(bold=True)
        c.fill = PARAM_FILL
        c.border = THIN_BORDER

    # Portfolio loss row
    loss_row = N_OBLIGORS + 3
    ws.cell(row=loss_row, column=1, value="Portfolio Loss")
    ws.cell(row=loss_row, column=1).font = Font(bold=True, color="FF0000")
    ws.cell(row=loss_row, column=1).border = THIN_BORDER

    # Trial headers + formulas
    for j in range(1, N_SIMS + 1):
        col = j + 1
        col_letter = get_column_letter(col)

        # Trial header
        c = ws.cell(row=1, column=col, value=f"Trial {j}")
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")

        # Default indicator for each obligor: =IF(RAND() < PD, 1, 0)
        for i in range(N_OBLIGORS):
            r = i + 2
            cell = ws.cell(row=r, column=col)
            cell.value = "=IF(RAND()<Parameters!$B$4,1,0)"
            cell.number_format = "0"

        # Portfolio loss = SUM(default_i * weight_i) * LGD
        # = (D1*w1 + D2*w2 + ...) * LGD
        # Use SUMPRODUCT(defaults, weights) * LGD
        first_r = 2
        last_r = N_OBLIGORS + 1
        weight_first = get_column_letter(2)
        weight_last = get_column_letter(N_OBLIGORS + 1)
        formula = (
            f"=SUMPRODUCT({col_letter}{first_r}:{col_letter}{last_r},"
            f"Weights!{weight_first}${s_row}:{weight_last}${s_row})"
            f"*Parameters!$B$5"
        )
        cell = ws.cell(row=loss_row, column=col, value=formula)
        cell.number_format = "0.00%"
        cell.font = Font(bold=True)

    ws.column_dimensions["A"].width = 16
    return ws


def create_results_sheet(wb):
    """Sheet: Results with formulas referencing all simulation sheets."""
    ws = wb.create_sheet("Results")
    ws.sheet_properties.tabColor = "00B050"

    ws["A1"] = "HHI vs 損失率 シミュレーション結果（全て数式）"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:I1")

    loss_row = N_OBLIGORS + 3  # portfolio loss row in each Sim sheet
    last_col = get_column_letter(N_SIMS + 1)

    # ── Summary Table ────────────────────────────────────────────────────
    ws["A3"] = "シナリオ別 損失指標"
    ws["A3"].font = SUBTITLE_FONT

    headers = [
        "シナリオ", "HHI",
        "期待損失率\nE[L]",
        "標準偏差\nσ",
        "理論値 σ\n=√(HHI·PD·(1-PD))",
        "VaR 99%",
        "VaR 99.9%",
        "最大損失率",
        "最大ウェイト",
    ]
    for c, h in enumerate(headers, 1):
        ws.cell(row=4, column=c, value=h)
    style_header_row(ws, 4, len(headers))

    for s_idx, (name, _) in enumerate(SCENARIOS):
        r = 5 + s_idx
        sim_sheet = f"Sim_{s_idx+1}_{name}"
        s_weight_row = 5 + s_idx  # row in Weights sheet
        w_first = get_column_letter(2)
        w_last = get_column_letter(N_OBLIGORS + 1)

        loss_range = f"'{sim_sheet}'!B{loss_row}:{last_col}{loss_row}"

        # Scenario name
        ws.cell(row=r, column=1, value=name).fill = PARAM_FILL
        style_cell(ws.cell(row=r, column=1))

        # HHI (from Weights sheet)
        hhi_col = get_column_letter(N_OBLIGORS + 3)
        c = ws.cell(row=r, column=2, value=f"=Weights!{hhi_col}{s_weight_row}")
        style_cell(c, "0.0000")

        # Expected Loss = AVERAGE
        c = ws.cell(row=r, column=3, value=f"=AVERAGE({loss_range})")
        style_cell(c, "0.0000%")

        # Std Dev = STDEV
        c = ws.cell(row=r, column=4, value=f"=STDEV({loss_range})")
        style_cell(c, "0.0000%")

        # Theoretical σ = SQRT(HHI * PD * (1-PD)) * LGD
        c = ws.cell(row=r, column=5,
                    value=f"=SQRT(B{r}*Parameters!$B$4*(1-Parameters!$B$4))*Parameters!$B$5")
        style_cell(c, "0.0000%")

        # VaR 99%
        c = ws.cell(row=r, column=6, value=f"=PERCENTILE({loss_range},0.99)")
        style_cell(c, "0.00%")

        # VaR 99.9%
        c = ws.cell(row=r, column=7, value=f"=PERCENTILE({loss_range},0.999)")
        style_cell(c, "0.00%")

        # Max Loss
        c = ws.cell(row=r, column=8, value=f"=MAX({loss_range})")
        style_cell(c, "0.00%")

        # Max Weight = MAX of weights
        c = ws.cell(row=r, column=9,
                    value=f"=MAX(Weights!{w_first}{s_weight_row}:{w_last}{s_weight_row})")
        style_cell(c, "0.0%")

    last_data_row = 4 + N_SCENARIOS

    # ── Theoretical E[L] reference ───────────────────────────────────────
    ws.cell(row=last_data_row + 2, column=1, value="理論値 E[L]").font = SUBTITLE_FONT
    c = ws.cell(row=last_data_row + 2, column=2,
                value="=Parameters!$B$4*Parameters!$B$5")
    style_cell(c, "0.0000%")
    ws.cell(row=last_data_row + 2, column=3,
            value="= PD × LGD（HHIに関係なく一定）")

    # ── Chart 1: σ simulated vs theoretical ──────────────────────────────
    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 10
    chart1.title = "HHI vs 標準偏差（シミュレーション vs 理論値）"
    chart1.x_axis.title = "シナリオ"
    chart1.y_axis.title = "標準偏差"
    chart1.y_axis.numFmt = "0.00%"
    chart1.width = 22
    chart1.height = 13

    cats = Reference(ws, min_col=1, min_row=5, max_row=last_data_row)
    d1 = Reference(ws, min_col=4, min_row=4, max_row=last_data_row)
    d2 = Reference(ws, min_col=5, min_row=4, max_row=last_data_row)
    chart1.add_data(d1, titles_from_data=True)
    chart1.add_data(d2, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.series[0].graphicalProperties.solidFill = "FF9800"
    chart1.series[1].graphicalProperties.solidFill = "FF0000"

    ws.add_chart(chart1, "A" + str(last_data_row + 5))

    # ── Chart 2: VaR 99.9% vs Max Loss (the pitfall) ────────────────────
    chart2 = BarChart()
    chart2.type = "col"
    chart2.style = 10
    chart2.title = "VaRの落とし穴: VaR 99.9% vs 最大損失率"
    chart2.x_axis.title = "シナリオ（集中度↑）"
    chart2.y_axis.title = "損失率"
    chart2.y_axis.numFmt = "0%"
    chart2.width = 22
    chart2.height = 13

    d3 = Reference(ws, min_col=7, min_row=4, max_row=last_data_row)
    d4 = Reference(ws, min_col=8, min_row=4, max_row=last_data_row)
    chart2.add_data(d3, titles_from_data=True)
    chart2.add_data(d4, titles_from_data=True)
    chart2.set_categories(cats)
    chart2.series[0].graphicalProperties.solidFill = "E91E63"
    chart2.series[1].graphicalProperties.solidFill = "9C27B0"

    ws.add_chart(chart2, "A" + str(last_data_row + 21))

    # ── Chart 3: All metrics by HHI ─────────────────────────────────────
    chart3 = BarChart()
    chart3.type = "col"
    chart3.style = 10
    chart3.title = "全指標比較（HHI順）"
    chart3.x_axis.title = "シナリオ（集中度↑）"
    chart3.y_axis.title = "損失率"
    chart3.y_axis.numFmt = "0%"
    chart3.width = 22
    chart3.height = 13

    for col_idx in [3, 4, 7, 8]:
        ref = Reference(ws, min_col=col_idx, min_row=4, max_row=last_data_row)
        chart3.add_data(ref, titles_from_data=True)
    chart3.set_categories(cats)

    colors = ["2196F3", "FF9800", "E91E63", "9C27B0"]
    for i, color in enumerate(colors):
        chart3.series[i].graphicalProperties.solidFill = color

    ws.add_chart(chart3, "A" + str(last_data_row + 37))

    # Column widths
    ws.column_dimensions["A"].width = 16
    for c in range(2, 10):
        ws.column_dimensions[get_column_letter(c)].width = 16

    # ── Key Formulas Reference ───────────────────────────────────────────
    formula_start = last_data_row + 54
    ws.cell(row=formula_start, column=1, value="使用されている数式").font = SUBTITLE_FONT

    formulas = [
        ("HHI (集中度指数)", "=SUMPRODUCT(w₁:w₁₀, w₁:w₁₀)", "Weightsシート M列"),
        ("デフォルト判定", "=IF(RAND()<PD, 1, 0)", "各Simシート（相関=0のBernoulli）"),
        ("ポートフォリオ損失", "=SUMPRODUCT(defaults, weights)*LGD", "各Simシート最下行"),
        ("期待損失率 E[L]", "=AVERAGE(全trial損失)", "= PD × LGD（HHI無関係）"),
        ("標準偏差 σ", "=STDEV(全trial損失)", "≈ √(HHI × PD × (1-PD))"),
        ("理論値 σ", "=SQRT(HHI*PD*(1-PD))*LGD", "再生性による解析解"),
        ("VaR 99.9%", "=PERCENTILE(損失, 0.999)", "上位0.1%の損失額"),
    ]
    ws.cell(row=formula_start + 1, column=1, value="計算項目")
    ws.cell(row=formula_start + 1, column=2, value="Excel数式")
    ws.cell(row=formula_start + 1, column=3, value="説明")
    style_header_row(ws, formula_start + 1, 3)

    for i, (label, formula, desc) in enumerate(formulas):
        r = formula_start + 2 + i
        ws.cell(row=r, column=1, value=label).fill = PARAM_FILL
        ws.cell(row=r, column=1).border = THIN_BORDER
        ws.cell(row=r, column=2, value=formula).border = THIN_BORDER
        ws.cell(row=r, column=3, value=desc).border = THIN_BORDER


def main():
    wb = openpyxl.Workbook()

    print("Creating Parameters sheet...")
    create_parameters_sheet(wb)

    print("Creating Weights sheet...")
    create_weights_sheet(wb)

    print(f"Creating {N_SCENARIOS} simulation sheets ({N_SIMS} trials each)...")
    for s_idx in range(N_SCENARIOS):
        name = SCENARIOS[s_idx][0]
        print(f"  Scenario {s_idx+1}: {name}...")
        create_simulation_sheet(wb, s_idx)

    print("Creating Results sheet...")
    create_results_sheet(wb)

    output_path = "hhi_loss_formula.xlsx"
    wb.save(output_path)
    print(f"\nDone! Saved to: {output_path}")
    print(f"  - {N_SCENARIOS} scenarios x {N_SIMS} trials")
    print(f"  - {N_OBLIGORS} obligors per scenario")
    print("  - Open in Excel and press F9 to re-simulate")
    print("  - Edit weights in the 'Weights' sheet")
    print("  - Edit PD/LGD in the 'Parameters' sheet")


if __name__ == "__main__":
    main()
