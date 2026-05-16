from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT_FILE = Path("bundle_planning_template.xlsx")
MAX_ROWS = 301


def style_header(ws, cells):
    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in cells:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def set_widths(ws, widths):
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def build_input_sheet(wb):
    ws = wb.active
    ws.title = "Input_Data"
    headers = [
        "Product_ID",
        "Product_Name",
        "Category",
        "Purchased_Qty_M1",
        "Purchased_Qty_M2",
        "Purchased_Qty_M3",
        "Purchased_Qty_M4",
        "Purchased_Qty_M5",
        "Purchased_Qty_M6",
        "Contribution_%_of_Total_Sales",
    ]
    ws.append(headers)
    style_header(ws, ws[1])

    sample_rows = [
        ["P-1001", "Sparkling Water 330ml", "Beverages", 440, 470, 510, 525, 560, 610, 0.091],
        ["P-1002", "Premium Coffee Beans", "Beverages", 210, 205, 198, 183, 176, 169, 0.044],
        ["P-2001", "Dry Fruit Mix 200g", "Snacks", 98, 95, 87, 75, 64, 55, 0.021],
        ["P-2002", "Sea Salt Crackers", "Snacks", 520, 548, 561, 575, 590, 615, 0.102],
        ["P-3001", "Aloe Vera Gel 250ml", "Personal Care", 41, 39, 35, 30, 24, 19, 0.013],
    ]
    for row in sample_rows:
        ws.append(row)

    for row in range(2, MAX_ROWS):
        ws[f"J{row}"].number_format = "0.00%"

    ws.freeze_panes = "A2"
    set_widths(
        ws,
        {
            1: 14,
            2: 30,
            3: 20,
            4: 18,
            5: 18,
            6: 18,
            7: 18,
            8: 18,
            9: 18,
            10: 28,
        },
    )


def build_scoring_sheet(wb):
    ws = wb.create_sheet("Scoring_Model")
    headers = [
        "Product_ID",
        "Product_Name",
        "Category",
        "Total_Qty_6M",
        "Avg_Monthly_Qty",
        "Movement_Percentile",
        "Contribution_%",
        "Contribution_Percentile",
        "Unmoved_Score",
        "Bundle_Priority_Score",
        "Movement_Band",
        "Bundle_Value_Cluster",
        "Anchor_Eligible",
        "Candidate_Eligible",
        "Suggested_Discount_%",
        "Anchor_Key",
        "Anchor_Name",
    ]
    ws.append(headers)
    style_header(ws, ws[1])

    for row in range(2, MAX_ROWS):
        ws[f"A{row}"] = f"=Input_Data!A{row}"
        ws[f"B{row}"] = f"=Input_Data!B{row}"
        ws[f"C{row}"] = f"=Input_Data!C{row}"
        ws[f"D{row}"] = f'=IF(A{row}="","",SUM(Input_Data!D{row}:I{row}))'
        ws[f"E{row}"] = f'=IF(D{row}="","",D{row}/6)'
        ws[f"F{row}"] = (
            f'=IF(D{row}="","",IF(COUNTA($D$2:$D$300)<=1,0,PERCENTRANK.INC($D$2:$D$300,D{row})))'
        )
        ws[f"G{row}"] = f"=Input_Data!J{row}"
        ws[f"H{row}"] = (
            f'=IF(G{row}="","",IF(COUNTA($G$2:$G$300)<=1,0,PERCENTRANK.INC($G$2:$G$300,G{row})))'
        )
        ws[f"I{row}"] = f'=IF(F{row}="","",1-F{row})'
        ws[f"J{row}"] = f'=IF(A{row}="","",0.6*I{row}+0.4*H{row})'
        ws[f"K{row}"] = (
            f'=IF(F{row}="","",IF(F{row}<=0.33,"Low Movement",IF(F{row}<=0.66,"Medium Movement","High Movement")))'
        )
        ws[f"L{row}"] = (
            f'=IF(J{row}="","",IF(J{row}>=0.67,"High Value Bundle",IF(J{row}>=0.34,"Medium Value Bundle","Low Value Bundle")))'
        )
        ws[f"M{row}"] = f'=IF(A{row}="","",IF(AND(F{row}>=0.66,H{row}>=0.66),"Yes","No"))'
        ws[f"N{row}"] = f'=IF(A{row}="","",IF(K{row}="Low Movement","Yes","No"))'
        ws[f"O{row}"] = (
            f'=IF(L{row}="","",IF(L{row}="High Value Bundle",0.10,IF(L{row}="Medium Value Bundle",0.15,0.20)))'
        )
        ws[f"P{row}"] = f'=IF(M{row}="Yes",C{row}&"|anchor","")'
        ws[f"Q{row}"] = f'=IF(M{row}="Yes",B{row},"")'

        for col in ("F", "G", "H", "I", "J", "O"):
            ws[f"{col}{row}"].number_format = "0.00%"

    ws.freeze_panes = "A2"
    set_widths(
        ws,
        {
            1: 14,
            2: 30,
            3: 20,
            4: 14,
            5: 16,
            6: 18,
            7: 14,
            8: 20,
            9: 14,
            10: 20,
            11: 18,
            12: 20,
            13: 15,
            14: 18,
            15: 20,
            16: 18,
            17: 25,
        },
    )

    ws.conditional_formatting.add(
        "L2:L300",
        FormulaRule(formula=['$L2="High Value Bundle"'], fill=PatternFill("solid", fgColor="FCE4D6")),
    )
    ws.conditional_formatting.add(
        "L2:L300",
        FormulaRule(formula=['$L2="Medium Value Bundle"'], fill=PatternFill("solid", fgColor="FFF2CC")),
    )
    ws.conditional_formatting.add(
        "L2:L300",
        FormulaRule(formula=['$L2="Low Value Bundle"'], fill=PatternFill("solid", fgColor="E2F0D9")),
    )


def build_recommendations_sheet(wb):
    ws = wb.create_sheet("Bundle_Recommendations")
    headers = [
        "Product_ID",
        "Candidate_Product",
        "Category",
        "Movement_Band",
        "Bundle_Value_Cluster",
        "Priority_Score",
        "Recommended_Anchor_Product",
        "Suggested_Discount_%",
        "Action_Note",
    ]
    ws.append(headers)
    style_header(ws, ws[1])

    for row in range(2, MAX_ROWS):
        ws[f"A{row}"] = f"=Scoring_Model!A{row}"
        ws[f"B{row}"] = f"=Scoring_Model!B{row}"
        ws[f"C{row}"] = f"=Scoring_Model!C{row}"
        ws[f"D{row}"] = f"=Scoring_Model!K{row}"
        ws[f"E{row}"] = f"=Scoring_Model!L{row}"
        ws[f"F{row}"] = f"=Scoring_Model!J{row}"
        ws[f"G{row}"] = (
            f'=IF(D{row}<>"Low Movement","",IFERROR(XLOOKUP(C{row}&"|anchor",Scoring_Model!$P$2:$P$300,Scoring_Model!$Q$2:$Q$300,'
            f'INDEX(Scoring_Model!$Q$2:$Q$300,MATCH("Yes",Scoring_Model!$M$2:$M$300,0))),"No anchor found"))'
        )
        ws[f"H{row}"] = f"=Scoring_Model!O{row}"
        ws[f"I{row}"] = (
            f'=IF(A{row}="","",IF(D{row}<>"Low Movement","Stable mover - bundle optional",'
            f'"Pair this low-movement item with the suggested anchor"))'
        )
        ws[f"F{row}"].number_format = "0.00%"
        ws[f"H{row}"].number_format = "0.00%"

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:I300"
    set_widths(
        ws,
        {1: 14, 2: 30, 3: 20, 4: 18, 5: 20, 6: 14, 7: 30, 8: 18, 9: 55},
    )


def build_logic_sheet(wb):
    ws = wb.create_sheet("Logic")
    ws["A1"] = "Bundle logic used in this workbook"
    ws["A1"].font = Font(bold=True, size=14)

    lines = [
        "1) Fill Input_Data with the last 6 months purchased quantity and contribution% per product.",
        "2) Movement_Percentile ranks each product by 6-month quantity sold (0 = least moved, 1 = fastest mover).",
        "3) Unmoved_Score = 1 - Movement_Percentile so low-movement products get higher focus score.",
        "4) Contribution_Percentile ranks products by contribution to total sales value.",
        "5) Bundle_Priority_Score = 0.60 * Unmoved_Score + 0.40 * Contribution_Percentile.",
        "6) Value clusters are based on Bundle_Priority_Score:",
        "   - High Value Bundle: score >= 0.67",
        "   - Medium Value Bundle: 0.34 <= score < 0.67",
        "   - Low Value Bundle: score < 0.34",
        "7) Anchor_Eligible products are those with strong movement and strong contribution (>=66th percentile both).",
        "8) Candidate_Eligible products are low-movement products; bundle them with anchor products from same category.",
        "9) Suggested discount defaults: High=10%, Medium=15%, Low=20%. Tune by margin and stock age.",
    ]

    row = 3
    for line in lines:
        ws[f"A{row}"] = line
        row += 1

    ws.column_dimensions["A"].width = 140


def main():
    wb = Workbook()
    build_input_sheet(wb)
    build_scoring_sheet(wb)
    build_recommendations_sheet(wb)
    build_logic_sheet(wb)
    wb.save(OUTPUT_FILE)
    print(f"Workbook created: {OUTPUT_FILE.resolve()}")


if __name__ == "__main__":
    main()
