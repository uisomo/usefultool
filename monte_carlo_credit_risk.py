#!/usr/bin/env python3
"""
Monte Carlo Credit Risk Simulation - Excel Prototype Generator
Generates an Excel file with formula-based Monte Carlo simulation (no VBA).

Uses a multi-factor Gaussian copula model:
  Z_i = sqrt(w_r)*F_region + sqrt(w_s)*F_sector + sqrt(w_h)*F_hnwi + sqrt(1-w_r-w_s-w_h)*eps_i

Obligor defaults when Z_i < NORM.S.INV(PD_i)
"""

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from copy import copy

# ── Configuration ──────────────────────────────────────────────────────────
N_SIMS = 500
N_BUCKETS = 25  # histogram buckets

# ── Sample Obligor Data ───────────────────────────────────────────────────
OBLIGORS = [
    # (ID, Name, Region, Sector, HNWI, PD%, EAD_millions, LGD%)
    ("OBL-001", "Alpha Industries",       "Asia",     "Manufacturing", "No",  1.5, 120, 45),
    ("OBL-002", "Beta Financial Group",    "Asia",     "Finance",       "No",  0.8, 200, 40),
    ("OBL-003", "Gamma Tech Solutions",    "Asia",     "Tech",          "No",  2.0,  80, 50),
    ("OBL-004", "Delta Real Estate",       "Asia",     "Real Estate",   "Yes", 3.0, 150, 55),
    ("OBL-005", "Epsilon Holdings",        "Asia",     "Finance",       "Yes", 1.2, 180, 40),
    ("OBL-006", "Zeta Automobil AG",       "Europe",   "Manufacturing", "No",  1.0, 160, 45),
    ("OBL-007", "Eta Banque SA",           "Europe",   "Finance",       "No",  0.5, 250, 35),
    ("OBL-008", "Theta PropCo Ltd",        "Europe",   "Real Estate",   "No",  2.5, 100, 55),
    ("OBL-009", "Iota Digital GmbH",       "Europe",   "Tech",          "Yes", 1.8,  90, 50),
    ("OBL-010", "Kappa Energie SA",        "Europe",   "Manufacturing", "No",  0.7, 140, 45),
    ("OBL-011", "Lambda Capital Corp",     "Americas", "Finance",       "No",  0.6, 220, 38),
    ("OBL-012", "Mu Semiconductor Inc",    "Americas", "Tech",          "No",  2.2,  70, 50),
    ("OBL-013", "Nu Development LLC",      "Americas", "Real Estate",   "Yes", 3.5,  95, 58),
    ("OBL-014", "Xi Wealth Partners",      "Americas", "Finance",       "Yes", 1.0, 170, 40),
    ("OBL-015", "Omicron MFG Corp",        "Americas", "Manufacturing", "Yes", 4.0,  60, 48),
]

REGIONS = ["Asia", "Europe", "Americas"]
SECTORS = ["Finance", "Manufacturing", "Tech", "Real Estate"]
HNWI_VALS = ["No", "Yes"]

# ── Styles ─────────────────────────────────────────────────────────────────
HEADER_FONT = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
TITLE_FONT = Font(name="Calibri", bold=True, size=14, color="2F5496")
SUBTITLE_FONT = Font(name="Calibri", bold=True, size=11, color="2F5496")
PARAM_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
INPUT_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
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


def style_data_cell(cell, num_format=None):
    cell.border = THIN_BORDER
    cell.alignment = Alignment(horizontal="center")
    if num_format:
        cell.number_format = num_format


def create_parameters_sheet(wb):
    """Sheet 1: Editable parameters."""
    ws = wb.active
    ws.title = "Parameters"
    ws.sheet_properties.tabColor = "2F5496"

    ws["A1"] = "Monte Carlo Credit Risk Simulation"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:C1")

    ws["A3"] = "Parameter"
    ws["B3"] = "Value"
    ws["C3"] = "Description"
    style_header_row(ws, 3, 3)

    params = [
        ("Number of Simulations", N_SIMS,  "Trials (press F9 to re-simulate)"),
        ("Region Factor Weight (w_r)", 0.15, "Correlation from regional factor"),
        ("Sector Factor Weight (w_s)", 0.10, "Correlation from sector factor"),
        ("HNWI Factor Weight (w_h)",   0.05, "Correlation from HNWI factor"),
        ("VaR Confidence Level",       0.99, "For Value-at-Risk calculation"),
    ]
    for i, (name, val, desc) in enumerate(params):
        r = 4 + i
        ws.cell(row=r, column=1, value=name).fill = PARAM_FILL
        c = ws.cell(row=r, column=2, value=val)
        c.fill = INPUT_FILL
        if isinstance(val, float) and val < 1:
            c.number_format = "0.00%"
        else:
            c.number_format = "#,##0"
        ws.cell(row=r, column=3, value=desc)
        for col in range(1, 4):
            ws.cell(row=r, column=col).border = THIN_BORDER

    # Computed: idiosyncratic weight
    ws.cell(row=9, column=1, value="Idiosyncratic Weight (1-w_r-w_s-w_h)").fill = PARAM_FILL
    c = ws.cell(row=9, column=2)
    c.value = "=1-B5-B6-B7"
    c.number_format = "0.00%"
    c.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    ws.cell(row=9, column=3, value="Remaining weight for idiosyncratic risk")
    for col in range(1, 4):
        ws.cell(row=9, column=col).border = THIN_BORDER

    # Instructions
    ws["A11"] = "How to Use"
    ws["A11"].font = SUBTITLE_FONT
    instructions = [
        "1. Review/edit obligor data in the 'Obligors' sheet",
        "2. Adjust factor weights above to change correlation structure",
        "3. Press F9 (or Ctrl+Alt+F9) to recalculate = run a new simulation",
        "4. View results in the 'Results' sheet",
        "5. The 'Factors' sheet shows systematic factor draws per trial",
        "6. The 'Simulation' sheet shows the correlated latent variables",
        "7. The 'Defaults' sheet shows binary default indicators (1=default)",
        "8. The 'Losses' sheet shows per-obligor loss amounts",
    ]
    for i, text in enumerate(instructions):
        ws.cell(row=12 + i, column=1, value=text)

    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 42

    return ws


def create_obligors_sheet(wb):
    """Sheet 2: Obligor portfolio data."""
    ws = wb.create_sheet("Obligors")
    ws.sheet_properties.tabColor = "548235"

    headers = ["ID", "Name", "Region", "Sector", "HNWI", "PD (%)", "EAD (M)", "LGD (%)", "Default Threshold"]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))

    for i, ob in enumerate(OBLIGORS):
        r = i + 2
        ws.cell(row=r, column=1, value=ob[0])
        ws.cell(row=r, column=2, value=ob[1])
        ws.cell(row=r, column=3, value=ob[2])
        ws.cell(row=r, column=4, value=ob[3])
        ws.cell(row=r, column=5, value=ob[4])
        c_pd = ws.cell(row=r, column=6, value=ob[5] / 100)
        c_pd.number_format = "0.00%"
        c_ead = ws.cell(row=r, column=7, value=ob[6])
        c_ead.number_format = "#,##0"
        c_lgd = ws.cell(row=r, column=8, value=ob[7] / 100)
        c_lgd.number_format = "0.00%"
        # Default threshold = NORM.S.INV(PD)
        c_thr = ws.cell(row=r, column=9)
        c_thr.value = f"=NORM.S.INV(F{r})"
        c_thr.number_format = "0.0000"

        for col in range(1, 10):
            style_data_cell(ws.cell(row=r, column=col))

    # Total exposure
    r_total = len(OBLIGORS) + 3
    ws.cell(row=r_total, column=6, value="Total EAD:").font = Font(bold=True)
    ws.cell(row=r_total, column=7, value=f"=SUM(G2:G{len(OBLIGORS)+1})")
    ws.cell(row=r_total, column=7).number_format = "#,##0"
    ws.cell(row=r_total, column=7).font = Font(bold=True)

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 8
    ws.column_dimensions["F"].width = 10
    ws.column_dimensions["G"].width = 12
    ws.column_dimensions["H"].width = 10
    ws.column_dimensions["I"].width = 16

    return ws


def create_factors_sheet(wb):
    """Sheet 3: Systematic factor draws (NORM.S.INV(RAND())) for each trial.

    Layout:
      Row 1: headers (Trial 1, Trial 2, ...)
      Col A: Factor labels
      Rows 2-4: Region factors (Asia, Europe, Americas)
      Rows 5-8: Sector factors (Finance, Manufacturing, Tech, Real Estate)
      Rows 9-10: HNWI factors (No, Yes)
    """
    ws = wb.create_sheet("Factors")
    ws.sheet_properties.tabColor = "BF8F00"

    # Column A labels
    ws.cell(row=1, column=1, value="Factor \\ Trial")
    ws.cell(row=1, column=1).font = HEADER_FONT
    ws.cell(row=1, column=1).fill = HEADER_FILL

    factor_labels = []
    for r in REGIONS:
        factor_labels.append(f"Region: {r}")
    for s in SECTORS:
        factor_labels.append(f"Sector: {s}")
    for h in HNWI_VALS:
        factor_labels.append(f"HNWI: {h}")

    for i, label in enumerate(factor_labels):
        c = ws.cell(row=i + 2, column=1, value=label)
        c.font = Font(bold=True)
        c.fill = PARAM_FILL
        c.border = THIN_BORDER

    # Trial headers + factor draws
    for j in range(1, N_SIMS + 1):
        col = j + 1
        c = ws.cell(row=1, column=col, value=f"Trial {j}")
        c.font = Font(bold=True, size=9)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center")

        # Each factor gets an independent NORM.S.INV(RAND())
        for i in range(len(factor_labels)):
            cell = ws.cell(row=i + 2, column=col)
            cell.value = "=NORM.S.INV(RAND())"
            cell.number_format = "0.0000"
            cell.border = THIN_BORDER

    ws.column_dimensions["A"].width = 22
    # Keep trial columns narrow
    for j in range(1, N_SIMS + 1):
        ws.column_dimensions[get_column_letter(j + 1)].width = 8

    # Add a lookup helper: map factor names to row numbers (in column A of a helper area)
    # We'll use the row positions directly in formulas instead.
    # Region rows: 2,3,4  (Asia=2, Europe=3, Americas=4)
    # Sector rows: 5,6,7,8 (Finance=5, Manufacturing=6, Tech=7, Real Estate=8)
    # HNWI rows: 9,10 (No=9, Yes=10)

    return ws


def _factor_row_for_obligor(region, sector, hnwi):
    """Return the row numbers in the Factors sheet for this obligor's factors."""
    region_row = 2 + REGIONS.index(region)
    sector_row = 2 + len(REGIONS) + SECTORS.index(sector)
    hnwi_row = 2 + len(REGIONS) + len(SECTORS) + HNWI_VALS.index(hnwi)
    return region_row, sector_row, hnwi_row


def create_simulation_sheet(wb):
    """Sheet 4: Correlated latent variable Z for each obligor x trial.

    Z_i = sqrt(w_r)*F_region + sqrt(w_s)*F_sector + sqrt(w_h)*F_hnwi
          + sqrt(1-w_r-w_s-w_h)*NORM.S.INV(RAND())

    Parameters!B5 = w_r, B6 = w_s, B7 = w_h
    """
    ws = wb.create_sheet("Simulation")
    ws.sheet_properties.tabColor = "C55A11"

    # Column A: obligor IDs
    ws.cell(row=1, column=1, value="Obligor \\ Trial")
    ws.cell(row=1, column=1).font = HEADER_FONT
    ws.cell(row=1, column=1).fill = HEADER_FILL

    for i, ob in enumerate(OBLIGORS):
        c = ws.cell(row=i + 2, column=1, value=ob[0])
        c.font = Font(bold=True)
        c.fill = PARAM_FILL
        c.border = THIN_BORDER

    # Trial headers
    for j in range(1, N_SIMS + 1):
        col = j + 1
        c = ws.cell(row=1, column=col, value=f"Trial {j}")
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")

    # Formulas
    for i, ob in enumerate(OBLIGORS):
        r_row = i + 2
        region_row, sector_row, hnwi_row = _factor_row_for_obligor(ob[2], ob[3], ob[4])

        for j in range(1, N_SIMS + 1):
            col = j + 1
            col_letter = get_column_letter(col)

            # Reference to factor sheet cells for this trial
            f_region = f"Factors!{col_letter}{region_row}"
            f_sector = f"Factors!{col_letter}{sector_row}"
            f_hnwi = f"Factors!{col_letter}{hnwi_row}"

            # Z = sqrt(w_r)*F_r + sqrt(w_s)*F_s + sqrt(w_h)*F_h + sqrt(1-w_r-w_s-w_h)*eps
            formula = (
                f"=SQRT(Parameters!$B$5)*{f_region}"
                f"+SQRT(Parameters!$B$6)*{f_sector}"
                f"+SQRT(Parameters!$B$7)*{f_hnwi}"
                f"+SQRT(Parameters!$B$9)*NORM.S.INV(RAND())"
            )
            cell = ws.cell(row=r_row, column=col, value=formula)
            cell.number_format = "0.0000"

    ws.column_dimensions["A"].width = 12
    return ws


def create_defaults_sheet(wb):
    """Sheet 5: Binary default indicators.
    =IF(Simulation!cell < Obligors!$I$row, 1, 0)
    """
    ws = wb.create_sheet("Defaults")
    ws.sheet_properties.tabColor = "FF0000"

    ws.cell(row=1, column=1, value="Obligor \\ Trial")
    ws.cell(row=1, column=1).font = HEADER_FONT
    ws.cell(row=1, column=1).fill = HEADER_FILL

    for i, ob in enumerate(OBLIGORS):
        c = ws.cell(row=i + 2, column=1, value=ob[0])
        c.font = Font(bold=True)
        c.fill = PARAM_FILL
        c.border = THIN_BORDER

    for j in range(1, N_SIMS + 1):
        col = j + 1
        col_letter = get_column_letter(col)
        c = ws.cell(row=1, column=col, value=f"Trial {j}")
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")

        for i in range(len(OBLIGORS)):
            r = i + 2
            # Default if Z < threshold
            formula = f"=IF(Simulation!{col_letter}{r}<Obligors!$I${r},1,0)"
            cell = ws.cell(row=r, column=col, value=formula)
            cell.number_format = "0"

    ws.column_dimensions["A"].width = 12
    return ws


def create_losses_sheet(wb):
    """Sheet 6: Loss per obligor per trial = Default * EAD * LGD."""
    ws = wb.create_sheet("Losses")
    ws.sheet_properties.tabColor = "7030A0"

    ws.cell(row=1, column=1, value="Obligor \\ Trial")
    ws.cell(row=1, column=1).font = HEADER_FONT
    ws.cell(row=1, column=1).fill = HEADER_FILL

    for i, ob in enumerate(OBLIGORS):
        c = ws.cell(row=i + 2, column=1, value=ob[0])
        c.font = Font(bold=True)
        c.fill = PARAM_FILL
        c.border = THIN_BORDER

    # Portfolio loss row
    loss_total_row = len(OBLIGORS) + 3
    ws.cell(row=loss_total_row, column=1, value="Portfolio Loss")
    ws.cell(row=loss_total_row, column=1).font = Font(bold=True, color="FF0000")

    for j in range(1, N_SIMS + 1):
        col = j + 1
        col_letter = get_column_letter(col)
        c = ws.cell(row=1, column=col, value=f"Trial {j}")
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")

        for i in range(len(OBLIGORS)):
            r = i + 2
            # Loss = Default_indicator * EAD * LGD
            formula = f"=Defaults!{col_letter}{r}*Obligors!$G${r}*Obligors!$H${r}"
            cell = ws.cell(row=r, column=col, value=formula)
            cell.number_format = "#,##0.0"

        # Sum column for portfolio loss
        first_row = 2
        last_row = len(OBLIGORS) + 1
        formula = f"=SUM({col_letter}{first_row}:{col_letter}{last_row})"
        cell = ws.cell(row=loss_total_row, column=col, value=formula)
        cell.number_format = "#,##0.0"
        cell.font = Font(bold=True)

    ws.column_dimensions["A"].width = 16
    return ws


def create_results_sheet(wb):
    """Sheet 7: Summary statistics and histogram."""
    ws = wb.create_sheet("Results")
    ws.sheet_properties.tabColor = "00B050"

    loss_total_row = len(OBLIGORS) + 3  # row in Losses sheet with portfolio totals
    last_col_letter = get_column_letter(N_SIMS + 1)

    # ── Title ──
    ws["A1"] = "Monte Carlo Simulation Results"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:D1")

    # ── Summary Statistics ──
    ws["A3"] = "Summary Statistics"
    ws["A3"].font = SUBTITLE_FONT

    stats = [
        ("Expected Loss (Mean)", f"=AVERAGE(Losses!B{loss_total_row}:{last_col_letter}{loss_total_row})", "#,##0.0"),
        ("Loss Std Dev", f"=STDEV(Losses!B{loss_total_row}:{last_col_letter}{loss_total_row})", "#,##0.0"),
        ("VaR (99%)", f"=PERCENTILE(Losses!B{loss_total_row}:{last_col_letter}{loss_total_row},Parameters!B8)", "#,##0.0"),
        ("Expected Shortfall (CVaR)",
         f"=AVERAGEIF(Losses!B{loss_total_row}:{last_col_letter}{loss_total_row},\">=\"&B6)",
         "#,##0.0"),
        ("Max Loss", f"=MAX(Losses!B{loss_total_row}:{last_col_letter}{loss_total_row})", "#,##0.0"),
        ("Min Loss", f"=MIN(Losses!B{loss_total_row}:{last_col_letter}{loss_total_row})", "#,##0.0"),
        ("Median Loss", f"=MEDIAN(Losses!B{loss_total_row}:{last_col_letter}{loss_total_row})", "#,##0.0"),
        ("Total Portfolio EAD", f"=Obligors!G{len(OBLIGORS)+3}", "#,##0"),
        ("Expected Loss Rate", f"=B4/B11", "0.00%"),
    ]

    ws["A4"] = "Metric"
    ws["B4"] = "Value"
    ws["C4"] = "Unit"
    style_header_row(ws, 4, 3)

    units = ["M", "M", "M", "M", "M", "M", "M", "M", "%"]
    for i, (label, formula, fmt) in enumerate(stats):
        r = 5 + i
        ws.cell(row=r, column=1, value=label).fill = PARAM_FILL
        c = ws.cell(row=r, column=2, value=formula)
        c.number_format = fmt
        c.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        ws.cell(row=r, column=3, value=units[i])
        for col in range(1, 4):
            ws.cell(row=r, column=col).border = THIN_BORDER

    # ── Default Count Statistics ──
    ws["A16"] = "Default Count Statistics"
    ws["A16"].font = SUBTITLE_FONT

    ws["A17"] = "Metric"
    ws["B17"] = "Value"
    style_header_row(ws, 17, 2)

    # Count defaults per trial (sum of Defaults sheet column)
    # We need a helper row. Use row 18-22 for stats based on default counts.
    # Average number of defaults per trial
    first_col = "B"
    # We'll compute from Defaults sheet: each trial column sum
    # Use a trick: SUMPRODUCT across the Defaults range / N_SIMS for avg defaults
    n_ob = len(OBLIGORS)
    ws.cell(row=18, column=1, value="Avg Defaults per Trial").fill = PARAM_FILL
    ws.cell(row=18, column=2,
            value=f"=SUMPRODUCT(Defaults!B2:{last_col_letter}{n_ob+1})/{N_SIMS}")
    ws.cell(row=18, column=2).number_format = "0.00"

    ws.cell(row=19, column=1, value="Avg Default Rate").fill = PARAM_FILL
    ws.cell(row=19, column=2, value=f"=B18/{n_ob}")
    ws.cell(row=19, column=2).number_format = "0.00%"

    for r in range(18, 20):
        for col in range(1, 3):
            ws.cell(row=r, column=col).border = THIN_BORDER

    # ── Loss Distribution (Histogram Data) ──
    ws["A22"] = "Loss Distribution"
    ws["A22"].font = SUBTITLE_FONT

    ws["A23"] = "Bucket Upper"
    ws["B23"] = "Frequency"
    ws["C23"] = "Cumulative %"
    style_header_row(ws, 23, 3)

    # Bucket boundaries: from 0 to Max Loss in N_BUCKETS steps
    # Bucket width = MAX(losses) / N_BUCKETS, but we use a fixed reasonable range
    # Use bin edges based on a formula referencing the max loss
    for b in range(N_BUCKETS):
        r = 24 + b
        # Bucket upper bound = (b+1) * MaxLoss / N_BUCKETS
        ws.cell(row=r, column=1,
                value=f"=({b + 1})*B9/{N_BUCKETS}")
        ws.cell(row=r, column=1).number_format = "#,##0.0"
        ws.cell(row=r, column=1).border = THIN_BORDER

    # FREQUENCY array - in openpyxl we write as CSE (Ctrl+Shift+Enter) array formula
    # FREQUENCY(data_array, bins_array)
    freq_start = 24
    freq_end = 24 + N_BUCKETS - 1
    data_range = f"Losses!B{loss_total_row}:{last_col_letter}{loss_total_row}"
    bins_range = f"A{freq_start}:A{freq_end}"

    for b in range(N_BUCKETS):
        r = 24 + b
        # Use COUNTIFS instead of FREQUENCY for compatibility (no CSE needed)
        if b == 0:
            formula = f'=COUNTIF({data_range},"<="&A{r})'
        else:
            formula = f'=COUNTIFS({data_range},"<="&A{r})-COUNTIFS({data_range},"<="&A{r - 1})'
        ws.cell(row=r, column=2, value=formula)
        ws.cell(row=r, column=2).number_format = "#,##0"
        ws.cell(row=r, column=2).border = THIN_BORDER

        # Cumulative %
        if b == 0:
            ws.cell(row=r, column=3, value=f"=B{r}/{N_SIMS}")
        else:
            ws.cell(row=r, column=3, value=f"=C{r - 1}+B{r}/{N_SIMS}")
        ws.cell(row=r, column=3).number_format = "0.0%"
        ws.cell(row=r, column=3).border = THIN_BORDER

    # ── Histogram Chart ──
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Portfolio Loss Distribution"
    chart.x_axis.title = "Loss Amount (M)"
    chart.y_axis.title = "Frequency"
    chart.width = 24
    chart.height = 14

    cats = Reference(ws, min_col=1, min_row=freq_start, max_row=freq_end)
    data = Reference(ws, min_col=2, min_row=23, max_row=freq_end)  # include header
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4

    # Style the bars
    series = chart.series[0]
    series.graphicalProperties.solidFill = "2F5496"
    chart.legend = None

    ws.add_chart(chart, "E4")

    # ── Cumulative Distribution Chart ──
    chart2 = BarChart()
    chart2.type = "col"
    chart2.style = 10
    chart2.title = "Cumulative Loss Distribution"
    chart2.x_axis.title = "Loss Amount (M)"
    chart2.y_axis.title = "Cumulative Probability"
    chart2.width = 24
    chart2.height = 14

    from openpyxl.chart import LineChart
    chart2 = LineChart()
    chart2.title = "Cumulative Loss Distribution"
    chart2.x_axis.title = "Loss Amount (M)"
    chart2.y_axis.title = "Cumulative Probability"
    chart2.width = 24
    chart2.height = 14

    data2 = Reference(ws, min_col=3, min_row=23, max_row=freq_end)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats)
    series2 = chart2.series[0]
    series2.graphicalProperties.line.solidFill = "C55A11"
    chart2.legend = None

    ws.add_chart(chart2, "E22")

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16

    return ws


def main():
    wb = openpyxl.Workbook()

    print("Creating Parameters sheet...")
    create_parameters_sheet(wb)

    print("Creating Obligors sheet...")
    create_obligors_sheet(wb)

    print("Creating Factors sheet...")
    create_factors_sheet(wb)

    print("Creating Simulation sheet...")
    create_simulation_sheet(wb)

    print("Creating Defaults sheet...")
    create_defaults_sheet(wb)

    print("Creating Losses sheet...")
    create_losses_sheet(wb)

    print("Creating Results sheet...")
    create_results_sheet(wb)

    output_path = "monte_carlo_credit_risk.xlsx"
    wb.save(output_path)
    print(f"\nDone! Saved to: {output_path}")
    print(f"  - {len(OBLIGORS)} obligors x {N_SIMS} trials")
    print("  - Open in Excel and press F9 to recalculate (new simulation)")
    print("  - Adjust weights in 'Parameters' sheet to change correlation structure")


if __name__ == "__main__":
    main()
