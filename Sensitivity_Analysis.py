from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.chart import ScatterChart, Reference, Series
import numpy as np

def create_sensitivity_analysis(wb):
    """Create and format the Sensitivity_Analysis sheet."""
    print("Creating Sensitivity_Analysis sheet...")
    
    # Get the Sensitivity_Analysis sheet
    ws = wb["Sensitivity_Analysis"]
    
    # Set the title
    ws['A1'] = "SENSITIVITY ANALYSIS"
    ws['A1'].font = Font(bold=True, size=14)
    
    # Base Case Results section
    ws['A3'] = "Base Case Results"
    ws['A3'].font = Font(bold=True)
    
    base_metrics = [
        ("NPV", "=CAPITAL_BUDGETING!B18", "#,##0"),
        ("IRR", "=CAPITAL_BUDGETING!B22", "0.00%"),
        ("Payback Period", "=CAPITAL_BUDGETING!B20", "0.00"),
        ("Profitability Index", "=CAPITAL_BUDGETING!B24", "0.00")
    ]
    
    for i, (metric, formula, format) in enumerate(base_metrics):
        row = i + 4
        ws[f'A{row}'] = metric
        ws[f'B{row}'] = formula
        ws[f'B{row}'].number_format = format
    
    # One-Variable Sensitivity (Discount Rate) section
    ws['A8'] = "Discount Rate Sensitivity"
    ws['A8'].font = Font(bold=True)
    
    ws['A9'] = "Discount Rate"
    ws['B9'] = "NPV"
    ws['C9'] = "IRR Impact"
    
    # Make headers bold
    for col in ['A', 'B', 'C']:
        ws[f'{col}9'].font = Font(bold=True)
    
    # Discount rate sensitivity analysis
    discount_rates = [0.06, 0.08, 0.1, 0.12, 0.14]
    
    for i, rate in enumerate(discount_rates, 10):
        row = i
        ws[f'A{row}'] = rate
        ws[f'A{row}'].number_format = '0.00%'
        # NPV calculation
        ws[f'B{row}'] = f"=NPV(A{row},CAPITAL_BUDGETING!B11:B15)+CAPITAL_BUDGETING!B10"
        ws[f'B{row}'].number_format = '#,##0'
        # IRR impact (percentage change from base IRR)
        ws[f'C{row}'] = f"=(IRR(CAPITAL_BUDGETING!B10:B15)-CAPITAL_BUDGETING!B22)/CAPITAL_BUDGETING!B22"
        ws[f'C{row}'].number_format = '0.00%'
    
    # Create scatter chart for discount rate sensitivity
    chart1 = ScatterChart()
    chart1.title = "NPV vs Discount Rate"
    chart1.style = 10
    chart1.x_axis.title = "Discount Rate"
    chart1.y_axis.title = "NPV"
    
    xvalues = Reference(ws, min_col=1, min_row=10, max_row=14)
    yvalues = Reference(ws, min_col=2, min_row=10, max_row=14)
    series = Series(yvalues, xvalues, title="NPV")
    chart1.series.append(series)
    
    # Add the chart to the worksheet
    ws.add_chart(chart1, "D8")
    
    # Two-Variable Sensitivity Analysis (NPV) section
    ws['A16'] = "Two-Variable Sensitivity Analysis (NPV)"
    ws['A16'].font = Font(bold=True)
    
    # Column headers
    ws['A17'] = "Annual Cash Flow % Change"
    ws['A17'].font = Font(bold=True)
    
    cf_changes = ["-20%", "-10%", "0%", "+10%", "+20%"]
    for j, change in enumerate(cf_changes, 2):
        col = get_column_letter(j)
        ws[f'{col}17'] = change
        ws[f'{col}17'].font = Font(bold=True)
    
    # Row headers
    ws['A18'] = "Initial Investment % Change"
    ws['A18'].font = Font(bold=True)
    
    inv_changes = ["-20%", "-10%", "0%", "+10%", "+20%"]
    
    # Create the sensitivity matrix with corrected multipliers
    for i, inv_change in enumerate(inv_changes, 19):
        row = i
        ws[f'A{row}'] = inv_change
        
        inv_multipliers = {"-20%": 0.8, "-10%": 0.9, "0%": 1, "+10%": 1.1, "+20%": 1.2}
        inv_multiplier = inv_multipliers[inv_change]
        
        for j, cf_change in enumerate(cf_changes, 2):
            col = get_column_letter(j)
            cf_multipliers = {"-20%": 0.8, "-10%": 0.9, "0%": 1, "+10%": 1.1, "+20%": 1.2}
            cf_multiplier = cf_multipliers[cf_change]
            
            # Updated formula to properly handle both initial investment and cash flow changes
            ws[f'{col}{row}'] = f"=NPV(CAPITAL_BUDGETING!B6,CAPITAL_BUDGETING!B11:B15*{cf_multiplier})+(CAPITAL_BUDGETING!B10*{inv_multiplier})"
            ws[f'{col}{row}'].number_format = '#,##0'
    
    # Add conditional formatting (color scale) to the sensitivity matrix
    red_to_green = ColorScaleRule(start_type='min', start_color='F8696B',
                                 mid_type='percentile', mid_value=50, mid_color='FFEB84',
                                 end_type='max', end_color='63BE7B')
    
    ws.conditional_formatting.add('B19:F23', red_to_green)
    
    # Break-even Analysis section
    ws['A25'] = "Break-even Analysis"
    ws['A25'].font = Font(bold=True)
    
    breakeven_metrics = [
        ("Break-even Annual Cash Flow", "=PMT(CAPITAL_BUDGETING!B6,CAPITAL_BUDGETING!B5,-CAPITAL_BUDGETING!B4,CAPITAL_BUDGETING!B7)", "#,##0"),
        ("% of Base Case Cash Flow", "=B26/AVERAGE(CAPITAL_BUDGETING!B11:B15)", "0.00%"),
        ("Required Growth Rate", "=(B26/CAPITAL_BUDGETING!B11)^(1/CAPITAL_BUDGETING!B5)-1", "0.00%")
    ]
    
    for i, (metric, formula, format) in enumerate(breakeven_metrics):
        row = i + 26
        ws[f'A{row}'] = metric
        ws[f'B{row}'] = formula
        ws[f'B{row}'].number_format = format
    
    # Risk Analysis section
    ws['A30'] = "Risk Analysis"
    ws['A30'].font = Font(bold=True)
    
    risk_metrics = [
        ("NPV Standard Deviation", "=STDEV(B19:F23)", "#,##0"),
        ("Coefficient of Variation", "=ABS(B31/AVERAGE(B19:F23))", "0.00"),
        ("NPV Range", "=MAX(B19:F23)-MIN(B19:F23)", "#,##0"),
        ("Probability of Negative NPV", "=COUNTIF(B19:F23,\"<0\")/25", "0.00%")
    ]
    
    for i, (metric, formula, format) in enumerate(risk_metrics):
        row = i + 31
        ws[f'A{row}'] = metric
        ws[f'B{row}'] = formula
        ws[f'B{row}'].number_format = format
    
    # Scenario Analysis section
    ws['A36'] = "Scenario Analysis"
    ws['A36'].font = Font(bold=True)
    
    scenarios = [
        ("Best Case", "=MAX(B19:F23)"),
        ("Base Case", "=INDEX(B19:F23,3,3)"),
        ("Worst Case", "=MIN(B19:F23)"),
        ("Expected Value", "=AVERAGE(B19:F23)"),
        ("Range of Outcomes", "=MAX(B19:F23)-MIN(B19:F23)")
    ]
    
    for i, (scenario, formula) in enumerate(scenarios):
        row = i + 37
        ws[f'A{row}'] = scenario
        ws[f'B{row}'] = formula
        ws[f'B{row}'].number_format = '#,##0'
    
    # Formatting
    # Add background colors
    header_fill = PatternFill(start_color='DCE6F1', end_color='DCE6F1', fill_type='solid')
    section_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    
    # Apply fills to section headers
    for row in [3, 8, 16, 25, 30, 36]:
        ws[f'A{row}'].fill = header_fill
    
    # Set column widths
    ws.column_dimensions['A'].width = 35
    for col in range(2, 7):
        ws.column_dimensions[get_column_letter(col)].width = 15
    
    print("Sensitivity_Analysis sheet created successfully")