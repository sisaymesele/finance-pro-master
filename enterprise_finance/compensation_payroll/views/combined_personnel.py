from django.shortcuts import render
from django.contrib.auth.decorators import login_required
from django.http import HttpResponse
from openpyxl import Workbook
from decimal import Decimal
from compensation_payroll.models import RegularPayroll
from compensation_payroll.services.combined.personnel_context import get_combined_personnel_payroll_context
from openpyxl.utils import get_column_letter
import openpyxl
from io import BytesIO
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from compensation_payroll.services.excel_export import ExportUtilityService

@login_required
def combined_personnel_detail(request):
    context = get_combined_personnel_payroll_context(request)
    return render(request, 'combined_payroll/personnel/detail.html', context)

@login_required
def combined_personnel_adjustment_list(request):
    context = get_combined_personnel_payroll_context(request)
    return render(request, 'combined_payroll/personnel/adjustment_list.html', context)

@login_required
def combined_personnel_payroll_list(request):
    context = get_combined_personnel_payroll_context(request)
    return render(request, 'combined_payroll/personnel/payroll_list.html', context)

@login_required
def combined_personnel_total(request):
    context = get_combined_personnel_payroll_context(request)
    return render(request, 'combined_payroll/personnel/total_list.html', context)

@login_required
def combined_personnel_expense(request):
    context = get_combined_personnel_payroll_context(request)
    return render(request, 'combined_payroll/personnel/expense_list.html', context)

@login_required
def combined_personnel_net_income(request):
    context = get_combined_personnel_payroll_context(request)
    return render(request, 'combined_payroll/personnel/net_income_list.html', context)


@login_required
def combined_personnel_employment_income_tax(request):
    context = get_combined_personnel_payroll_context(request)
    return render(request, 'combined_payroll/personnel/employment_income_tax_list.html', context)


@login_required
def combined_employee_pension(request):
    context = get_combined_personnel_payroll_context(request)
    return render(request, 'combined_payroll/personnel/pension_list.html', context)


@login_required
def combined_personnel_detail(request):
    context = get_combined_personnel_payroll_context(request)
    return render(request, 'combined_payroll/personnel/detail.html', context)



#common export header



#detail
def export_combined_personnel_detail(request):

    # Get the same data as the regular view
    context = get_combined_personnel_payroll_context(request)
    payroll_data = context['payroll_data']

    # Create workbook and worksheet
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Combined Personnel Payroll"

    # Styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    border = Border(left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin'))

    # Use built-in number format instead of numbers.FORMAT_NUMBER_COMMA_SEPARATED2
    money_format = '#,##0.00'

    row_num = 1

    for item in payroll_data:
        # Header with merged cells
        ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=7)
        header_cell = ws.cell(row=row_num, column=1,
                              value=f"Combined Payslip For {item['payroll'].personnel_full_name} | {item['payroll'].payroll_month}")
        header_cell.font = Font(bold=True, size=14)
        header_cell.alignment = Alignment(horizontal="center")
        row_num += 2

        # Regular Payroll Section
        ws.cell(row=row_num, column=1, value="Regular Payroll").font = Font(bold=True, color="0070C0")
        row_num += 1

        # Regular Payroll Headers
        headers = ["Component", "Amount"]
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=row_num, column=col_num, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = border
        row_num += 1

        # Regular Payroll Data
        for component, amount in item['regular_item_by_component'].items():
            if amount:
                ws.cell(row=row_num, column=1, value=component)
                amount_cell = ws.cell(row=row_num, column=2, value=float(amount))
                amount_cell.number_format = money_format
                row_num += 1

        row_num += 1

        # Earning Adjustment Section (if exists)
        if item['show_earning']:
            ws.cell(row=row_num, column=1, value="Earning Adjustment").font = Font(bold=True, color="00B050")
            row_num += 1

            # Earning Adjustment Headers
            headers = ["Component", "Total", "Taxable", "Non-Taxable",
                       "Employee Pension", "Employer Pension", "Total Pension"]
            for col_num, header in enumerate(headers, 1):
                cell = ws.cell(row=row_num, column=col_num, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
                cell.border = border
            row_num += 1

            # Earning Adjustment Data
            for component, amounts in item['earning_adj_by_component'].items():
                if amounts['earning_amount']:
                    ws.cell(row=row_num, column=1, value=component)
                    ws.cell(row=row_num, column=2, value=float(amounts['earning_amount'])).number_format = money_format
                    ws.cell(row=row_num, column=3, value=float(amounts['taxable'])).number_format = money_format
                    ws.cell(row=row_num, column=4, value=float(amounts['non_taxable'])).number_format = money_format
                    ws.cell(row=row_num, column=5,
                            value=float(amounts['employee_pension_contribution'])).number_format = money_format
                    ws.cell(row=row_num, column=6,
                            value=float(amounts['employer_pension_contribution'])).number_format = money_format
                    ws.cell(row=row_num, column=7, value=float(amounts['total_pension'])).number_format = money_format
                    row_num += 1

            row_num += 1

        # Deduction Adjustment Section (if exists)
        if item['show_deduction']:
            ws.cell(row=row_num, column=1, value="Deduction Adjustment").font = Font(bold=True, color="FF0000")
            row_num += 1

            # Deduction Adjustment Headers
            headers = ["Component", "Amount"]
            for col_num, header in enumerate(headers, 1):
                cell = ws.cell(row=row_num, column=col_num, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
                cell.border = border
            row_num += 1

            # Deduction Adjustment Data
            for component, amount in item['deduction_adj_by_component'].items():
                if amount:
                    ws.cell(row=row_num, column=1, value=component)
                    amount_cell = ws.cell(row=row_num, column=2, value=float(amount))
                    amount_cell.number_format = money_format
                    row_num += 1

            row_num += 1

        # Summary Section
        ws.cell(row=row_num, column=1, value="Total Summary").font = Font(bold=True)
        row_num += 1

        # Summary Headers
        headers = ["Component", "Amount"]
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=row_num, column=col_num, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = border
        row_num += 1

        # Summary Data
        summary_items = [
            ("Taxable Gross Pay", item['totals']['taxable_gross']),
            ("Non-Taxable Gross Pay", item['totals']['non_taxable_gross']),
            ("Total Gross Pay", item['totals']['gross_pay']),
            ("Total Pensionable", item['totals']['pensionable']),
            ("Employee Pension", item['totals']['employee_pension']),
            ("Employer Pension", item['totals']['employer_pension']),
            ("Total Pension", item['totals']['total_pension']),
            ("Income Tax", item['totals']['employment_income_tax']),
            ("Total Deduction", item['totals']['deduction']),
            ("Total Expense", item['totals']['expense']),
            ("Final Net Pay", item['totals']['final_net_pay']),
        ]

        for component, amount in summary_items:
            ws.cell(row=row_num, column=1, value=component)
            amount_cell = ws.cell(row=row_num, column=2, value=float(amount))
            amount_cell.number_format = money_format
            row_num += 1

        # Add space between records
        row_num += 3

    # Auto-size columns
    for col in ws.columns:
        # Skip merged cells in column width calculation
        valid_cells = [cell for cell in col if not isinstance(cell, openpyxl.cell.cell.MergedCell)]
        if not valid_cells:
            continue

        max_length = 0
        column = valid_cells[0].column_letter  # Get letter from first valid cell

        for cell in valid_cells:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass

        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column].width = adjusted_width

    # Create HTTP response
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename=combined_personnel_payroll.xlsx'
    wb.save(response)

    return response



@login_required
def export_combined_personnel_list(request):
    context = get_combined_personnel_payroll_context(request)
    payroll_data = context.get('payroll_data', [])

    wb = Workbook()
    ws = wb.active
    ws.title = "Combined Payroll List Report"

    # Title row (1st)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=12)
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = "Combined Personnel Payroll List"
    title_cell.font = Font(size=14, bold=True)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')

    # Title row (2nd)
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=10)
    title_cell = ws.cell(row=2, column=2)
    title_cell.value = "Regular, Adjustment and Totals Personnel Payroll List"
    title_cell.font = Font(size=14, bold=True)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')

    headers = [
        'Payroll Month', 'Personnel ID', 'First Name', 'Father Name', 'Last Name',
        'Category',
        'Taxable Gross', 'Non-Taxable Gross', 'Gross Pay', 'Pensionable',
        'Employee Pension', 'Employer Pension', 'Total Pension',
        'Income Tax', 'Deductions', 'Net Pay', 'Expense'
    ]

    # call header
    # call from service header decorate
    export_util = ExportUtilityService()
    # Transform headers with line splitting
    ws.append([export_util.split_header_to_lines(h) for h in headers])

    # Header styling

    header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")  # Orange
    header_font = Font(bold=True)
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


    for cell in ws[3]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment

    total_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

    def to_dec(value):
        if value is None:
            return Decimal('0.00')
        if isinstance(value, Decimal):
            return value
        try:
            return Decimal(value)
        except Exception:
            return Decimal('0.00')

    def write_row(personnel, payroll_month, category_name, data, show_personnel_info=False):
        row_data = [
            payroll_month if show_personnel_info else "",
            personnel.personnel_id if show_personnel_info else "",
            personnel.first_name if show_personnel_info else "",
            personnel.father_name if show_personnel_info else "",
            personnel.last_name if show_personnel_info else "",
            category_name,
            to_dec(data.get('taxable_gross')),
            to_dec(data.get('non_taxable_gross')),
            to_dec(data.get('gross_pay')),
            to_dec(data.get('pensionable', data.get('adjusted_pensionable'))),
            to_dec(data.get('employee_pension')),
            to_dec(data.get('employer_pension')),
            to_dec(data.get('total_pension')),
            to_dec(data.get('employment_income_tax')),
            to_dec(data.get('deduction', data.get('total_adjustment_deduction'))),
            to_dec(data.get('net_pay', data.get('net_monthly_adjustment', data.get('final_net_pay')))),
            to_dec(data.get('expense')),
        ]
        ws.append(row_data)
        if category_name.lower() == "total":
            for col_num in range(1, len(row_data) + 1):
                ws.cell(row=ws.max_row, column=col_num).fill = total_fill

    for item in payroll_data:
        payroll = item.get('payroll')
        personnel = getattr(payroll, 'personnel_full_name', None)
        payroll_month = getattr(payroll, 'payroll_month', None)
        earning = item.get('earning_adjustment_item', {})
        regular = item.get('regular_totals', {})
        totals = item.get('totals', {})

        if not personnel or not payroll_month:
            continue

        payroll_month_str = getattr(payroll_month, 'payroll_month', '')

        first_row = True

        if to_dec(regular.get('gross_pay')) > 0:
            write_row(personnel, payroll_month_str, "Regular", regular, show_personnel_info=first_row)
            first_row = False

        if to_dec(earning.get('gross_pay')) > 0:
            write_row(personnel, payroll_month_str, "Adjustment", earning, show_personnel_info=first_row)
            first_row = False

        write_row(personnel, payroll_month_str, "Total", totals, show_personnel_info=first_row)

    # Adjust column widths dynamically
    MIN_WIDTH = 12
    MAX_WIDTH = 15
    for i, col_cells in enumerate(ws.columns, 1):
        max_length = 0
        for cell in col_cells:
            try:
                if cell.value:
                    length = len(str(cell.value))
                    if length > max_length:
                        max_length = length
            except Exception:
                pass
        adjusted_width = max(MIN_WIDTH, min(MAX_WIDTH, max_length + 2))
        ws.column_dimensions[get_column_letter(i)].width = adjusted_width

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    response = HttpResponse(
        output.read(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    response['Content-Disposition'] = 'attachment; filename=combined_payroll_list_report.xlsx'
    return response


@login_required
def export_personnel_total_adjustment(request):
    context = get_combined_personnel_payroll_context(request)
    data = context.get('payroll_data', [])

    wb = Workbook()
    ws = wb.active
    ws.title = "Personnel Total Adjustment"

    # Title row (1st)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=19)
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = "Combined Personnel Total Adjustment"
    title_cell.font = Font(size=14, bold=True)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')

    # Subtitle row (2nd)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=19)
    subtitle_cell = ws.cell(row=2, column=1)
    subtitle_cell.value = "Details of personnel combined total earning and deduction adjustments per payroll month"
    subtitle_cell.font = Font(size=10, italic=True)
    subtitle_cell.alignment = Alignment(horizontal='center', vertical='center')

    headers = [
        "#", "Payroll Month", "Personnel ID", "First Name", "Father Name", "Last Name",
        "Adjustment Taxable Gross", "Adjustment Non-Taxable Gross", "Adjustment Gross Pay",
        "Adjusted Pensionable", "Employee Pension", "Employer Pension", "Total Pension",
        "Employment Income Tax", "Earning Adjustment Deduction", "Other Adjustment Deduction",
        "Total Adjustment Deduction", "Net Adjustment Pay", "Adjustment Expense"
    ]

    #call from service header decorate
    export_util = ExportUtilityService()
    # Transform headers with line splitting
    ws.append([export_util.split_header_to_lines(h) for h in headers])


    # Style header
    header_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFFFF")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for cell in ws[3]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment

    for i, item in enumerate(data, 1):
        payroll = item.get('payroll')
        earning = item.get('earning_adjustment_item', {})
        combined_adj = item.get('combined_adjustment', {})
        deduction_adj = item.get('deduction_adjustment')

        personnel = getattr(payroll, 'personnel_full_name', None)
        payroll_month = getattr(getattr(payroll, 'payroll_month', None), 'payroll_month', '')

        row = [
            i,
            payroll_month,
            getattr(personnel, 'personnel_id', '') if personnel else '',
            getattr(personnel, 'first_name', '') if personnel else '',
            getattr(personnel, 'father_name', '') if personnel else '',
            getattr(personnel, 'last_name', '') if personnel else '',
            earning.get('taxable_gross', Decimal('0.00')),
            earning.get('non_taxable_gross', Decimal('0.00')),
            earning.get('gross_pay', Decimal('0.00')),
            earning.get('adjusted_pensionable', Decimal('0.00')),
            earning.get('employee_pension', Decimal('0.00')),
            earning.get('employer_pension', Decimal('0.00')),
            earning.get('total_pension', Decimal('0.00')),
            earning.get('employment_income_tax', Decimal('0.00')),
            earning.get('earning_adjustment_deduction', Decimal('0.00')),
            getattr(deduction_adj, 'monthly_adjusted_deduction', Decimal('0.00')),
            combined_adj.get('total_adjustment_deduction', Decimal('0.00')),
            combined_adj.get('net_monthly_adjustment', Decimal('0.00')),
            earning.get('expense', Decimal('0.00')),
        ]
        ws.append(row)

    # Adjust column widths with limits
    MIN_WIDTH = 10
    MAX_WIDTH = 25
    for i, col_cells in enumerate(ws.columns, 1):
        max_length = 0
        for cell in col_cells:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception:
                pass
        adjusted_width = max(MIN_WIDTH, min(MAX_WIDTH, max_length + 2))
        ws.column_dimensions[get_column_letter(i)].width = adjusted_width

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    response = HttpResponse(
        output.read(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response["Content-Disposition"] = "attachment; filename=personnel_total_adjustment.xlsx"
    return response




@login_required
def export_combined_personnel_total(request):
    context = get_combined_personnel_payroll_context(request)
    payroll_data = context.get('payroll_data', [])

    wb = Workbook()
    ws = wb.active
    ws.title = "Total Payroll Summary"

    # Title row (1st)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=16)
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = "Combined Personnel Total Payroll Summary"
    title_cell.font = Font(size=14, bold=True)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')


    headers = [
        "Month", "Personnel ID", "First Name", "Father Name", "Last Name",
        "Taxable Gross", "Non-Taxable Gross", "Total Gross Pay", "Total Pensionable",
        "Employee Pension", "Employer Pension", "Total Pension",
        "Employment Income Tax", "Total Deduction", "Final Net Pay", "Total Expense"
    ]
    #call header
    # call from service header decorate
    export_util = ExportUtilityService()
    # Transform headers with line splitting
    ws.append([export_util.split_header_to_lines(h) for h in headers])

    # Style header
    header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")  # Orange fill
    header_font = Font(bold=True)
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for cell in ws[2]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment

    # Append data rows
    for item in payroll_data:
        p = getattr(item['payroll'], 'personnel_full_name', None)
        pm = getattr(item['payroll'], 'payroll_month', None)
        t = item.get('totals', {})

        ws.append([
            getattr(pm, 'payroll_month', '') if pm else '',
            getattr(p, 'personnel_id', '') if p else '',
            getattr(p, 'first_name', '') if p else '',
            getattr(p, 'father_name', '') if p else '',
            getattr(p, 'last_name', '') if p else '',
            float(t.get('taxable_gross', 0)),
            float(t.get('non_taxable_gross', 0)),
            float(t.get('gross_pay', 0)),
            float(t.get('pensionable', 0)),
            float(t.get('employee_pension', 0)),
            float(t.get('employer_pension', 0)),
            float(t.get('total_pension', 0)),
            float(t.get('employment_income_tax', 0)),
            float(t.get('deduction', 0)),
            float(t.get('final_net_pay', 0)),
            float(t.get('expense', 0)),
        ])

    # Adjust column widths with limits
    MIN_WIDTH = 10
    MAX_WIDTH = 15
    for i, col_cells in enumerate(ws.columns, 2):
        max_length = 0
        for cell in col_cells:
            try:
                if cell.value:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
            except Exception:
                pass
        adjusted_width = max(MIN_WIDTH, min(MAX_WIDTH, max_length + 2))
        ws.column_dimensions[get_column_letter(i)].width = adjusted_width

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    response = HttpResponse(
        output.read(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response["Content-Disposition"] = "attachment; filename=total_combined_payroll_summary.xlsx"
    return response




@login_required
def export_combined_personnel_expense(request):
    context = get_combined_personnel_payroll_context(request)
    payroll_data = context.get('payroll_data', [])

    wb = Workbook()
    ws = wb.active
    ws.title = "Total Payroll Expense Summary"

    # Title row (1st)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = "Combined Personnel Expense Summary"
    title_cell.font = Font(size=14, bold=True)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')


    headers = [
        "Month", "Personnel ID", "First Name", "Father Name", "Last Name",
        "Total Gross Pay", "Employer Pension", "Total Expense"
    ]

    # call header
    # call from service header decorate
    export_util = ExportUtilityService()
    # Transform headers with line splitting
    ws.append([export_util.split_header_to_lines(h) for h in headers])

    # Style header
    header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")  # Orange fill
    header_font = Font(bold=True)
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for cell in ws[2]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment

    for item in payroll_data:
        p = getattr(item.get('payroll'), 'personnel_full_name', None)
        pm = getattr(item.get('payroll'), 'payroll_month', None)
        t = item.get('totals', {})

        ws.append([
            getattr(pm, 'payroll_month', '') if pm else '',
            getattr(p, 'personnel_id', '') if p else '',
            getattr(p, 'first_name', '') if p else '',
            getattr(p, 'father_name', '') if p else '',
            getattr(p, 'last_name', '') if p else '',
            float(t.get('gross_pay', 0)),
            float(t.get('employer_pension', 0)),
            float(t.get('expense', 0)),
        ])

    # Adjust column widths with limits
    MIN_WIDTH = 10
    MAX_WIDTH = 15
    for i, col_cells in enumerate(ws.columns, 1):
        max_length = 0
        for cell in col_cells:
            try:
                if cell.value:
                    length = len(str(cell.value))
                    if length > max_length:
                        max_length = length
            except Exception:
                pass
        adjusted_width = max(MIN_WIDTH, min(MAX_WIDTH, max_length + 2))
        ws.column_dimensions[get_column_letter(i)].width = adjusted_width

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    response = HttpResponse(
        output.read(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = "attachment; filename=combined_payroll_expense_summary.xlsx"
    return response

@login_required
def export_combined_personnel_net_income(request):
    context = get_combined_personnel_payroll_context(request)
    payroll_data = context.get('payroll_data', [])

    wb = Workbook()
    ws = wb.active
    ws.title = "Total Net Income Summary"

    # Title row (1st)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = "Combined Personnel Net Income Summary"
    title_cell.font = Font(size=14, bold=True)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')

    headers = [
        "Month", "Personnel ID", "First Name", "Father Name", "Last Name",
        "Total Gross Pay", "Total Deduction", "Final Net Pay"
    ]


    # call header
    # call from service header decorate
    export_util = ExportUtilityService()
    # Transform headers with line splitting
    ws.append([export_util.split_header_to_lines(h) for h in headers])

    # Header styling
    header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")  # Orange
    header_font = Font(bold=True)
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for cell in ws[2]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment

    for item in payroll_data:
        payroll = item.get('payroll')
        personnel = getattr(payroll, 'personnel_full_name', None)
        payroll_month = getattr(payroll, 'payroll_month', None)
        totals = item.get('totals', {})

        ws.append([
            getattr(payroll_month, 'payroll_month', '') if payroll_month else '',
            getattr(personnel, 'personnel_id', '') if personnel else '',
            getattr(personnel, 'first_name', '') if personnel else '',
            getattr(personnel, 'father_name', '') if personnel else '',
            getattr(personnel, 'last_name', '') if personnel else '',
            float(totals.get('gross_pay', 0)),
            float(totals.get('deduction', 0)),
            float(totals.get('final_net_pay', 0)),
        ])

    # Adjust column widths
    MIN_WIDTH = 10
    MAX_WIDTH = 15
    for i, col_cells in enumerate(ws.columns, 1):
        max_length = 0
        for cell in col_cells:
            try:
                if cell.value:
                    length = len(str(cell.value))
                    if length > max_length:
                        max_length = length
            except Exception:
                pass
        adjusted_width = max(MIN_WIDTH, min(MAX_WIDTH, max_length + 2))
        ws.column_dimensions[get_column_letter(i)].width = adjusted_width

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    response = HttpResponse(
        output.read(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response['Content-Disposition'] = 'attachment; filename=combined_net_income_summary.xlsx'
    return response



@login_required
def export_combined_personnel_employment_tax(request):
    context = get_combined_personnel_payroll_context(request)
    payroll_data = context.get('payroll_data', [])

    wb = Workbook()
    ws = wb.active
    ws.title = "Employment Tax Summary"

    # Title row (1st)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = "Combined Personnel Total Employment Income Tax Summary"
    title_cell.font = Font(size=14, bold=True)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')

    headers = [
        "Month", "Personnel ID", "First Name", "Father Name", "Last Name",
        "Total Taxable Gross", "Total Employment Income Tax"
    ]


    # call from service header decorate
    export_util = ExportUtilityService()
    # Transform headers with line splitting
    ws.append([export_util.split_header_to_lines(h) for h in headers])

    # Header style
    header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")  # Orange
    header_font = Font(bold=True)
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for cell in ws[2]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment

    for item in payroll_data:
        payroll = item.get('payroll')
        personnel = getattr(payroll, 'personnel_full_name', None)
        payroll_month = getattr(payroll, 'payroll_month', None)
        totals = item.get('totals', {})

        ws.append([
            getattr(payroll_month, 'payroll_month', '') if payroll_month else '',
            getattr(personnel, 'personnel_id', '') if personnel else '',
            getattr(personnel, 'first_name', '') if personnel else '',
            getattr(personnel, 'father_name', '') if personnel else '',
            getattr(personnel, 'last_name', '') if personnel else '',
            float(totals.get('taxable_gross', 0)),
            float(totals.get('employment_income_tax', 0)),
        ])

    # Adjust column widths
    MIN_WIDTH = 10
    MAX_WIDTH = 15
    for i, col_cells in enumerate(ws.columns, 1):
        max_length = 0
        for cell in col_cells:
            try:
                if cell.value:
                    length = len(str(cell.value))
                    if length > max_length:
                        max_length = length
            except Exception:
                pass
        adjusted_width = max(MIN_WIDTH, min(MAX_WIDTH, max_length + 2))
        ws.column_dimensions[get_column_letter(i)].width = adjusted_width

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    response = HttpResponse(
        output.read(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response['Content-Disposition'] = 'attachment; filename=employment_tax_summary.xlsx'
    return response


#detail export

#pension export
@login_required
def export_combined_personnel_pension(request):
    """
    Export Combined Total Pension Summary (Per Employee Per Month) as Excel.
    """
    context = get_combined_personnel_payroll_context(request)
    payroll_data = context.get('payroll_data', [])

    wb = Workbook()
    ws = wb.active
    ws.title = "Pension Summary"

    # Title row (1st)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = "Combined Personnel Pension Summery"
    title_cell.font = Font(size=14, bold=True)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')

    headers = [
        'Payroll Month', 'Personnel ID', 'First Name', 'Father Name', 'Last Name',
        'Total Pensionable', 'Total Employee Pension',
        'Total Employer Pension', 'Total Pension'
    ]

    # call from service header decorate
    export_util = ExportUtilityService()
    # Transform headers with line splitting
    ws.append([export_util.split_header_to_lines(h) for h in headers])

    # Header style
    header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")  # Orange
    header_font = Font(bold=True)
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for cell in ws[2]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment

    def safe_float(value):
        if value is None:
            return 0.0
        if isinstance(value, Decimal):
            return float(value)
        try:
            return float(value)
        except Exception:
            return 0.0

    for item in payroll_data:
        payroll = item.get('payroll')
        totals = item.get('totals', {})
        personnel = getattr(payroll, 'personnel_full_name', None)
        payroll_month = getattr(payroll.payroll_month, 'payroll_month', '') if payroll else ''

        if not personnel:
            continue

        ws.append([
            getattr(personnel, 'personnel_id', ''),
            getattr(personnel, 'first_name', ''),
            getattr(personnel, 'father_name', ''),
            getattr(personnel, 'last_name', ''),
            payroll_month,
            safe_float(totals.get('pensionable')),
            safe_float(totals.get('employee_pension')),
            safe_float(totals.get('employer_pension')),
            safe_float(totals.get('total_pension')),
        ])

    # Adjust column widths
    MIN_WIDTH = 12
    MAX_WIDTH = 15
    for i, col_cells in enumerate(ws.columns, 1):
        max_len = max((len(str(cell.value)) if cell.value else 0) for cell in col_cells)
        adjusted_width = max(MIN_WIDTH, min(MAX_WIDTH, max_len + 2))
        ws.column_dimensions[get_column_letter(i)].width = adjusted_width

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    response = HttpResponse(
        output.read(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename=pension_summary.xlsx'
    return response