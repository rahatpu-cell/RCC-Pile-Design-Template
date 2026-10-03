from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import LineChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

OUTPUT_FILE = "RCC_Pile_Design_Template.xlsx"


def apply_title(ws, title, row=1, col=1, end_col=8):
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=end_col)
    cell = ws.cell(row=row, column=col, value=title)
    cell.font = Font(size=16, bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    cell.alignment = Alignment(horizontal="center")


def set_label(ws, row, label, value=None, col_label=1, col_value=2, bold=True):
    label_cell = ws.cell(row=row, column=col_label, value=label)
    label_cell.font = Font(bold=bold)
    label_cell.fill = PatternFill("solid", fgColor="EAF2F8")
    value_cell = ws.cell(row=row, column=col_value, value=value)
    value_cell.border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    if value is not None and isinstance(value, (int, float)):
        value_cell.number_format = "0.00"
    return label_cell, value_cell


def add_button(ws, cell_ref, label, target_sheet, fill="4F81BD"):
    cell = ws[cell_ref]
    cell.value = label
    cell.font = Font(color="FFFFFF", bold=True)
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = Border(
        left=Side(style="thin", color="FFFFFF"),
        right=Side(style="thin", color="FFFFFF"),
        top=Side(style="thin", color="FFFFFF"),
        bottom=Side(style="thin", color="FFFFFF"),
    )
    cell.hyperlink = f"'{target_sheet}'!A1"
    return cell


def add_section_title(ws, row, title, fill="D9EAF7"):
    cell = ws.cell(row=row, column=1, value=title)
    cell.font = Font(bold=True, color="1F1F1F")
    cell.fill = PatternFill("solid", fgColor=fill)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    cell.alignment = Alignment(horizontal="left")


def create_cover_sheet(wb):
    ws = wb.active
    ws.title = "Cover"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    ws["A1"] = "RCC Cast-in-Situ Pile Design Template"
    ws["A1"].font = Font(size=18, bold=True, color="1F1F1F")
    ws["A2"] = "Ultimate Excel Design Workbook for Pile Systems"
    ws["A2"].font = Font(size=11, italic=True)

    ws["A4"] = "Design Basis"
    ws["A4"].font = Font(size=12, bold=True)
    ws["A5"] = "- BNBC 2020 seismic actions"
    ws["A6"] = "- ACI / ASTM concrete and steel requirements"
    ws["A7"] = "- Eurocode design considerations"
    ws["A8"] = "- Soil investigation inputs including SPT, liquefaction and landslide checks"
    ws["A9"] = "- Cast-in-situ RCC pile axial, flexural, and shear design"

    buttons = [
        ("A12", "Inputs", "Inputs", "4F81BD"),
        ("A14", "Soil Checks", "Soil_Checks", "77933C"),
        ("A16", "Design Checks", "Design_Checks", "C0504D"),
        ("A18", "Seismic", "Seismic", "5B9BD5"),
        ("A20", "BMD/SFD", "BMD_SFD", "8064A2"),
        ("A22", "Sketch", "Sketch", "ED7D31"),
        ("A24", "Summary", "Summary", "70AD47"),
    ]

    for ref, label, target, fill in buttons:
        cell = add_button(ws, ref, label, target, fill=fill)
        ws[ref].alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions["A"].width = 26

    ws["C4"] = "Prepared for:"
    ws["C4"].font = Font(bold=True)
    ws["C5"] = "Project / Client / Agency"
    ws["D5"] = "(enter data)"
    ws["C7"] = "Engineer / Reviewer"
    ws["C8"] = "(enter professional name)"
    ws["C10"] = "Revision / Date"
    ws["C11"] = "(enter date)"

    ws["F4"] = "Key Notes"
    ws["F4"].font = Font(bold=True)
    notes = [
        "- Check soil profile and SPT values against pile capacity.",
        "- Verify liquefaction and landslide conditions for extreme events.",
        "- Update seismic zone and site class values for dynamic design.",
        "- Confirm reinforcement detailing and cover requirements.",
    ]
    for i, n in enumerate(notes, start=5):
        ws[f"F{i}"] = n

    for row in range(1, 30):
        ws.row_dimensions[row].height = 20

    ws["A30"] = "Design assumptions: must be validated by the responsible engineer and governing code authority."
    ws["A30"].font = Font(italic=True, color="4F4F4F")


def create_inputs_sheet(wb):
    ws = wb.create_sheet("Inputs")
    ws.freeze_panes = "A5"
    apply_title(ws, "RCC CAST-IN-SITU PILE DESIGN INPUTS", 1, 1, 8)
    ws["A3"] = "Project / Structure"
    ws["A3"].font = Font(bold=True)
    ws["B3"] = "(enter project code)"

    # Data entry area
    entries = [
        (4, "Pile diameter, D (mm)", 1200),
        (5, "Pile length, L (m)", 18),
        (6, "Concrete strength, f'c (MPa)", 35),
        (7, "Steel yield strength, fy (MPa)", 415),
        (8, "Concrete cover (mm)", 75),
        (9, "Soil unit weight, gamma (kN/m3)", 18.5),
        (10, "SPT N-value (blows/ft)", 25),
        (11, "Cohesion, c (kPa)", 0),
        (12, "Friction angle, phi (deg)", 30),
        (13, "Water table depth (m)", 3),
        (14, "Pile spacing / group center spacing (m)", 3),
        (15, "Axial service load, P (kN)", 1500),
        (16, "Ultimate axial load, Pu (kN)", 2200),
        (17, "Lateral load, H (kN)", 200),
        (18, "Moment at pile head, M (kN.m)", 180),
        (19, "Seismic zone", "Zone 2"),
        (20, "Site class", "D"),
        (21, "PGA (g)", 0.25),
        (22, "Importance factor, I", 1.25),
        (23, "Response reduction factor, R", 5),
        (24, "Liquefaction susceptibility", "Low"),
        (25, "Landslide risk", "Low"),
        (26, "Pile type", "Cast-in-situ RCC"),
    ]

    for row, label, value in entries:
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=value)
        ws.cell(row=row, column=1).font = Font(bold=False)
        ws.cell(row=row, column=1).fill = PatternFill("solid", fgColor="EAF2F8")
        ws.cell(row=row, column=2).border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )

    ws["A30"] = "Design reference values"
    ws["A30"].font = Font(bold=True)
    ref_table = [
        ("Code / Standard", "Design factor / reference"),
        ("BNBC 2020", "Seismic design coefficient and detailing requirements"),
        ("ACI 318", "Flexural / shear capacity, reinforcement limits"),
        ("ASTM", "Concrete and steel material grade verification"),
        ("EC2 / EC8", "Partial factors, detailing, and seismic ductility checks"),
    ]

    for r, (a, b) in enumerate(ref_table, start=31):
        ws.cell(row=r, column=1, value=a)
        ws.cell(row=r, column=2, value=b)
        ws.cell(row=r, column=1).fill = PatternFill("solid", fgColor="F2F2F2")

    # Data validation dropdowns
    dv_zone = DataValidation(type="list", formula1='"Zone 1,Zone 2,Zone 3,Zone 4,Zone 5"', allow_blank=False)
    dv_zone.add("B19")
    ws.add_data_validation(dv_zone)

    dv_site = DataValidation(type="list", formula1='"A,B,C,D,E,F"', allow_blank=False)
    dv_site.add("B20")
    ws.add_data_validation(dv_site)

    dv_liq = DataValidation(type="list", formula1='"Low,Moderate,High"', allow_blank=False)
    dv_liq.add("B24")
    ws.add_data_validation(dv_liq)

    dv_landslide = DataValidation(type="list", formula1='"Low,Moderate,High"', allow_blank=False)
    dv_landslide.add("B25")
    ws.add_data_validation(dv_landslide)

    # Set widths
    for col in [1, 2, 3, 4, 5, 6, 7, 8]:
        ws.column_dimensions[get_column_letter(col)].width = 22
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 18


def create_soil_sheet(wb):
    ws = wb.create_sheet("Soil_Checks")
    apply_title(ws, "SOIL CHECKS AND GEOTECHNICAL SCREENING", 1, 1, 8)

    rows = [
        (3, "SPT N-value", "=Inputs!B10", "blows/ft"),
        (4, "Effective overburden stress", "=0.5*Inputs!B9*Inputs!B13", "kPa"),
        (5, "Liquefaction screening", "=IF(Inputs!B10<15, \"Potential\", \"Low risk\")", "flag"),
        (6, "Landslide screening", "=IF(Inputs!B25=\"High\", \"Critical\", \"Stable\")", "flag"),
        (7, "Pile end bearing estimate", "=0.5*Inputs!B10*Inputs!B9*Inputs!B5", "kN"),
        (8, "Skin friction estimate", "=0.8*Inputs!B10*Inputs!B9*Inputs!B5", "kN"),
        (9, "Required shaft resistance check", "=IF((B7+B8)>Inputs!B16,\"Pass\",\"Review\")", "status"),
    ]

    for row, label, formula, unit in rows:
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=formula)
        ws.cell(row=row, column=3, value=unit)
        ws.cell(row=row, column=1).fill = PatternFill("solid", fgColor="EAF2F8")

    ws["A12"] = "Design note: review liquefaction and landslide hazards carefully for extreme seismic scenarios."
    ws["A12"].font = Font(italic=True)
    for col in [1, 2, 3]:
        ws.column_dimensions[get_column_letter(col)].width = 24


def create_design_sheet(wb):
    ws = wb.create_sheet("Design_Checks")
    apply_title(ws, "PILE DESIGN CHECKS", 1, 1, 8)

    # Basic geometry
    ws["A3"] = "Parameter"
    ws["B3"] = "Value"
    ws["C3"] = "Unit"
    ws["D3"] = "Check"

    rows = [
        (4, "Diameter (D)", "=Inputs!B4", "mm"),
        (5, "Length (L)", "=Inputs!B5", "m"),
        (6, "Area", "=PI()/4*Inputs!B4^2", "mm2"),
        (7, "Moment of inertia", "=(PI()/64)*Inputs!B4^4", "mm4"),
        (8, "Concrete strength", "=Inputs!B6", "MPa"),
        (9, "Steel strength", "=Inputs!B7", "MPa"),
        (10, "Elastic modulus Ec", "=4700*SQRT(Inputs!B6)", "MPa"),
        (11, "Design axial capacity", "=0.4*Inputs!B6*PI()/4*Inputs!B4^2/1000", "kN"),
        (12, "Maximum shear capacity", "=0.17*SQRT(Inputs!B6)*PI()/4*Inputs!B4^2/1000", "kN"),
        (13, "Required axial status", "=IF(B11>Inputs!B16,\"Pass\",\"Review\")", ""),
        (14, "Required flexural status", "=IF(B12>Inputs!B18,\"Pass\",\"Review\")", ""),
    ]

    for row, a, b, c in rows:
        ws.cell(row=row, column=1, value=a)
        ws.cell(row=row, column=2, value=b)
        ws.cell(row=row, column=3, value=c)
        ws.cell(row=row, column=1).fill = PatternFill("solid", fgColor="EAF2F8")

    # Code notes
    ws["F3"] = "BNBC 2020 / ACI / ASTM / Eurocode Notes"
    ws["F3"].font = Font(bold=True)
    notes = [
        "- Check pile capacity using geotechnical resistance with factor of safety.",
        "- Verify concrete cover, reinforcement ratio, and detailing against ACI/Eurocode.",
        "- Include dynamic amplification and seismic demand adjustment as required.",
        "- Confirm material strength and testing per ASTM specifications.",
        "- Evaluate liquefaction / landslide effects for extreme conditions.",
    ]
    for idx, txt in enumerate(notes, start=4):
        ws[f"F{idx}"] = txt

    ws["A18"] = "Summary of design checks"
    ws["A18"].font = Font(bold=True)
    summary = [
        ("Axial capacity", "=IF(B13=\"Pass\",\"OK\",\"Needs review\")"),
        ("Flexural capacity", "=IF(B14=\"Pass\",\"OK\",\"Needs review\")"),
        ("Shear capacity", "=IF(B12>Inputs!B17,\"OK\",\"Needs review\")"),
        ("Seismic demand", "=IF(Inputs!B21>0.15,\"Critical review\",\"Moderate\")"),
    ]
    for i, (a, b) in enumerate(summary, start=19):
        ws[f"A{i}"] = a
        ws[f"B{i}"] = b

    for col in [1, 2, 3, 4, 5, 6, 7, 8]:
        ws.column_dimensions[get_column_letter(col)].width = 18


def create_seismic_sheet(wb):
    ws = wb.create_sheet("Seismic")
    apply_title(ws, "SEISMIC ACTIONS AND DESIGN FACTORS", 1, 1, 8)

    entries = [
        (3, "Seismic zone", "=Inputs!B19", ""),
        (4, "Site class", "=Inputs!B20", ""),
        (5, "Importance factor, I", "=Inputs!B22", ""),
        (6, "Response reduction, R", "=Inputs!B23", ""),
        (7, "PGA", "=Inputs!B21", "g"),
        (8, "Peak spectral acceleration", "=Inputs!B21*1.5", "g"),
        (9, "Equivalent static seismic force", "=Inputs!B17*Inputs!B21*1.2", "kN"),
        (10, "Dynamic amplification factor", "=1.0+0.25*(Inputs!B21/0.2)", ""),
        (11, "Liquefaction risk", "=Inputs!B24", ""),
        (12, "Landslide risk", "=Inputs!B25", ""),
        (13, "Critical review required", "=IF(OR(Inputs!B24=\"High\",Inputs!B25=\"High\"),\"YES\",\"NO\")", ""),
    ]

    for row, label, formula, unit in entries:
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=formula)
        ws.cell(row=row, column=3, value=unit)
        ws.cell(row=row, column=1).fill = PatternFill("solid", fgColor="EAF2F8")

    ws["A16"] = "BNBC 2020 / ACI / EC Notes"
    ws["A16"].font = Font(bold=True)
    notes = [
        "- Seismic zone, site class, and ground motion values must be updated case-by-case.",
        "- Base shear and lateral force effects shall be checked using relevant dynamic analysis methods.",
        "- For liquefiable / landslide-prone sites, adopt additional geotechnical and structural safeguards.",
        "- Use ACI/Eurocode detailing provisions for ductility, confinement, and reinforcement anchorage.",
    ]
    for idx, txt in enumerate(notes, start=17):
        ws[f"A{idx}"] = txt

    for col in [1, 2, 3]:
        ws.column_dimensions[get_column_letter(col)].width = 24


def create_bmd_sheet(wb):
    ws = wb.create_sheet("BMD_SFD")
    apply_title(ws, "BENDING MOMENT AND SHEAR FORCE DIAGRAMS", 1, 1, 8)
    ws["A3"] = "Depth (m)"
    ws["B3"] = "Static BMD (kN.m)"
    ws["C3"] = "Dynamic BMD (kN.m)"
    ws["D3"] = "Static SFD (kN)"
    ws["E3"] = "Dynamic SFD (kN)"

    # Create a simple textbook-style loading distribution for demonstration
    # Approximate demand varies linearly with depth and seismic amplification.
    data = []
    for i in range(0, 21):
        x = i / 2
        static_bmd = max(0, 150 - 8.5 * x)
        dynamic_bmd = static_bmd * (1 + 0.25 * 0.25 / 0.2)
        static_sfd = max(0, 180 - 9 * x)
        dynamic_sfd = static_sfd * (1 + 0.25 * 0.25 / 0.2)
        data.append((x, static_bmd, dynamic_bmd, static_sfd, dynamic_sfd))

    for r, vals in enumerate(data, start=4):
        for c, val in enumerate(vals, start=1):
            ws.cell(row=r, column=c, value=val)

    # Add chart
    chart = LineChart()
    chart.title = "Pile BMD / SFD"
    chart.style = 12
    chart.y_axis.title = "Force"
    chart.x_axis.title = "Depth (m)"

    data_ref = Reference(ws, min_col=2, min_row=3, max_col=5, max_row=23)
    cats = Reference(ws, min_col=1, min_row=4, max_row=23)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "H4")

    ws["A28"] = "Note: The direction and peak values should be calibrated to the project-specific pile-soil model and dynamic analysis result."
    ws["A28"].font = Font(italic=True)

    for col in [1, 2, 3, 4, 5]:
        ws.column_dimensions[get_column_letter(col)].width = 18


def create_sketch_sheet(wb):
    ws = wb.create_sheet("Sketch")
    apply_title(ws, "PILE DETAIL SKETCH", 1, 1, 8)
    ws["A3"] = "Plan / Elevation / Section"
    ws["A3"].font = Font(bold=True)

    # Simple conceptual pile sketch using merged cells and borders
    for row in range(4, 30):
        for col in range(1, 10):
            cell = ws.cell(row=row, column=col)
            cell.border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )

    # Draw a pile elevation
    ws.merge_cells("A5:F18")
    ws["A5"] = "Pile shaft\nCast-in-situ RCC\nD = project input\nL = project input"
    ws["A5"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["A5"].font = Font(bold=True)
    ws["A5"].fill = PatternFill("solid", fgColor="D9EAF7")

    ws.merge_cells("H5:N18")
    ws["H5"] = "Cross-section\nfc' = input\nFy = input\nCover = input\nReinf. cage = detailed design"
    ws["H5"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["H5"].font = Font(bold=True)
    ws["H5"].fill = PatternFill("solid", fgColor="F2F2F2")

    ws["A22"] = "Detailing notes"
    ws["A22"].font = Font(bold=True)
    notes = [
        "- Check minimum reinforcement ratio and detailing against seismic provisions.",
        "- Confirm cover, lap splices, and cage stability.",
        "- Verify soil-structure interaction and pile cap connection design.",
        "- Check group effects and pile spacing if applicable.",
    ]
    for idx, txt in enumerate(notes, start=23):
        ws[f"A{idx}"] = txt

    for col in range(1, 15):
        ws.column_dimensions[get_column_letter(col)].width = 18


def create_summary_sheet(wb):
    ws = wb.create_sheet("Summary")
    apply_title(ws, "DESIGN SUMMARY", 1, 1, 8)

    summary_items = [
        (3, "Project / Name", "(enter project)", ""),
        (4, "Pile diameter", "=Inputs!B4", "mm"),
        (5, "Pile length", "=Inputs!B5", "m"),
        (6, "Critical soil issue", "=IF(Soil_Checks!B5=\"Potential\",\"Liquefaction\",\"Stable\")", ""),
        (7, "Seismic zone", "=Inputs!B19", ""),
        (8, "Site class", "=Inputs!B20", ""),
        (9, "Required review", "=IF(OR(Seismic!B13=\"YES\", Soil_Checks!B6=\"Critical\"),\"YES\",\"NO\")", ""),
        (10, "Final status", "=IF(B9=\"YES\",\"Redesign required\",\"Acceptable preliminary design\")", ""),
    ]

    for row, label, formula, unit in summary_items:
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=formula)
        ws.cell(row=row, column=3, value=unit)
        ws.cell(row=row, column=1).fill = PatternFill("solid", fgColor="EAF2F8")

    ws["A14"] = "Important: This workbook is a design template requiring final professional engineering review and project-specific verification."
    ws["A14"].font = Font(italic=True, color="6B6B6B")

    for col in [1, 2, 3]:
        ws.column_dimensions[get_column_letter(col)].width = 24


def main():
    wb = Workbook()
    create_cover_sheet(wb)
    create_inputs_sheet(wb)
    create_soil_sheet(wb)
    create_design_sheet(wb)
    create_seismic_sheet(wb)
    create_bmd_sheet(wb)
    create_sketch_sheet(wb)
    create_summary_sheet(wb)

    # Add conditional highlighting to summary statuses
    ws = wb["Summary"]
    ws.conditional_formatting.add("B9", CellIsRule(operator="equal", formula=['"YES"'], fill=PatternFill("solid", fgColor="F4CCCC")))
    ws.conditional_formatting.add("B10", CellIsRule(operator="equal", formula=['"Redesign required"'], fill=PatternFill("solid", fgColor="F4CCCC")))

    # Add workbook metadata
    wb.properties.title = "RCC Cast-In-Situ Pile Design Template"
    wb.properties.creator = "Copilot"
    wb.properties.company = "Engineering Design"

    wb.save(OUTPUT_FILE)
    print(f"Workbook generated: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()


