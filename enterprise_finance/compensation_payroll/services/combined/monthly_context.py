from decimal import Decimal
from collections import defaultdict
from django.core.paginator import Paginator
from django.db.models import Sum
from compensation_payroll.models import RegularPayroll, SeverancePay


def get_combined_monthly_detail(request):
    def safe_dec(value):
        return value if value is not None else Decimal('0.00')

    #
    def format_key(key):
        return key.replace('_', ' ').title()

    #
    def has_nonzero_earning(earning_adj_by_component):
        return any(v['earning_amount'] != Decimal('0.00') for v in earning_adj_by_component.values())

    def has_nonzero_deduction(deduction_adj_by_component):
        return any(v != Decimal('0.00') for v in deduction_adj_by_component.values())

    #

    payrolls = RegularPayroll.objects.filter(
        organization_name=request.user.organization_name
    ).select_related('payroll_month') \
        .prefetch_related('earning_adjustments', 'deduction_adjustments')
    payrolls = payrolls.order_by('-payroll_month__payroll_month__year', '-payroll_month__payroll_month__month')

    #
    monthly_summary = defaultdict(lambda: {
        'regular': {
            'taxable_gross': Decimal('0.00'),
            'non_taxable_gross': Decimal('0.00'),
            'gross': Decimal('0.00'),
            'pensionable': Decimal('0.00'),
            'employee_pension': Decimal('0.00'),
            'employer_pension': Decimal('0.00'),
            'total_pension': Decimal('0.00'),
            'employment_income_tax': Decimal('0.00'),
            'total_regular_deduction': Decimal('0.00'),
            'net_pay': Decimal('0.00'),
            'expense': Decimal('0.00'),
            'components': defaultdict(lambda: Decimal('0.00')),
        },
        'adjustment': {
            'taxable_gross': Decimal('0.00'),
            'non_taxable_gross': Decimal('0.00'),
            'gross': Decimal('0.00'),
            'adjusted_pensionable': Decimal('0.00'),
            'employee_pension': Decimal('0.00'),
            'employer_pension': Decimal('0.00'),
            'total_pension': Decimal('0.00'),
            'employment_income_tax': Decimal('0.00'),
            'net_monthly_adjustment': Decimal('0.00'),
            'earning_adjustment_deduction': Decimal('0.00'),
            'expense': Decimal('0.00'),
            'earning_adj_by_component': {
                'taxable': defaultdict(lambda: Decimal('0.00')),
                'non_taxable': defaultdict(lambda: Decimal('0.00')),
                'total_earning_adjustment': defaultdict(lambda: Decimal('0.00')),
                'employee_pension_contribution': defaultdict(lambda: Decimal('0.00')),
                'employer_pension_contribution': defaultdict(lambda: Decimal('0.00')),
                'total_pension': defaultdict(lambda: Decimal('0.00')),
            },
            'individual_adjustment_deduction': Decimal('0.00'),
            'deduction_adj_by_component': defaultdict(lambda: Decimal('0.00')),
        },
        'severance': {
            'taxable_gross': Decimal('0.00'),
            'gross': Decimal('0.00'),
            'employment_income_tax': Decimal('0.00'),
            'total_severance_deduction': Decimal('0.00'),
            'net': Decimal('0.00'),
            'expense': Decimal('0.00'),
        },
        'totals': {
            'taxable_gross': Decimal('0.00'),
            'non_taxable_gross': Decimal('0.00'),
            'gross': Decimal('0.00'),
            'pensionable': Decimal('0.00'),
            'employee_pension': Decimal('0.00'),
            'employer_pension': Decimal('0.00'),
            'total_pension': Decimal('0.00'),
            'employment_income_tax': Decimal('0.00'),
            'total_deduction': Decimal('0.00'),
            'expense': Decimal('0.00'),
            'final_net_pay': Decimal('0.00'),
        }
    })

    # Aggregate payrolls into monthly summary
    for payroll in payrolls:
        if payroll and payroll.payroll_month:
            key = payroll.payroll_month.payroll_month.payroll_month
        else:
            key = None  # or some default/fallback value, or skip the record

        #assign before it's initialized for the current payroll month key

        month_data = monthly_summary[key]

        # Sum regular components
        regular_components = {
            'basic_salary': payroll.basic_salary,
            'overtime': payroll.overtime,
            'housing_allowance': payroll.housing_allowance,
            'position_allowance': payroll.position_allowance,
            'commission': payroll.commission,
            'telephone_allowance': payroll.telephone_allowance,
            'one_time_bonus': payroll.one_time_bonus,
            'causal_labor_wage': payroll.causal_labor_wage,
            #
            'transport_home_to_office_taxable': payroll.transport_home_to_office_taxable,
            'transport_home_to_office_non_taxable': payroll.transport_home_to_office_non_taxable,

            'fuel_home_to_office_taxable': payroll.fuel_home_to_office_taxable,
            'fuel_home_to_office_non_taxable': payroll.fuel_home_to_office_non_taxable,

            'transport_for_work_taxable': payroll.transport_for_work_taxable,
            'transport_for_work_non_taxable': payroll.transport_for_work_non_taxable,

            'fuel_for_work_taxable': payroll.fuel_for_work_taxable,
            'fuel_for_work_non_taxable': payroll.fuel_for_work_non_taxable,

            'per_diem_taxable': payroll.per_diem_taxable,
            'per_diem_non_taxable': payroll.per_diem_non_taxable,

            'hardship_allowance_taxable': payroll.hardship_allowance_taxable,
            'hardship_allowance_non_taxable': payroll.hardship_allowance_non_taxable,

            'public_cash_award': payroll.public_cash_award,
            'incidental_operation_allowance': payroll.incidental_operation_allowance,
            'medical_allowance': payroll.medical_allowance,
            'cash_gift': payroll.cash_gift,
            'tuition_fees': payroll.tuition_fees,
            'personal_injury': payroll.personal_injury,
            'child_support_payment': payroll.child_support_payment,
            'charitable_donation': payroll.charitable_donation,
            'saving_plan': payroll.saving_plan,
            'loan_payment': payroll.loan_payment,
            'court_order': payroll.court_order,
            'workers_association': payroll.workers_association,
            'personnel_insurance_saving': payroll.personnel_insurance_saving,
            'university_cost_share_pay': payroll.university_cost_share_pay,
            'red_cross': payroll.red_cross,
            'party_contribution': payroll.party_contribution,
            'other_deduction': payroll.other_deduction,
        }
        for comp, val in regular_components.items():
            month_data['regular']['components'][comp] += safe_dec(val)

        month_data['regular']['taxable_gross'] += safe_dec(payroll.gross_taxable_pay)
        month_data['regular']['non_taxable_gross'] += safe_dec(payroll.gross_non_taxable_pay)
        month_data['regular']['gross'] += safe_dec(payroll.gross_pay)
        month_data['regular']['pensionable'] += safe_dec(payroll.basic_salary)
        month_data['regular']['employee_pension'] += safe_dec(payroll.employee_pension_contribution)
        month_data['regular']['employer_pension'] += safe_dec(payroll.employer_pension_contribution)
        month_data['regular']['total_pension'] += safe_dec(payroll.total_pension_contribution)
        month_data['regular']['employment_income_tax'] += safe_dec(payroll.employment_income_tax)
        month_data['regular']['total_regular_deduction'] += safe_dec(payroll.total_payroll_deduction)
        month_data['regular']['net_pay'] += safe_dec(payroll.net_pay)
        month_data['regular']['expense'] += safe_dec(payroll.expense)
        #
        # # Aggregate earning adjustments by component
        # Sum adjusted pensionable directly from DB
        pensionable_sum = payroll.earning_adjustments.filter(
            component='basic_salary',
        ).aggregate(total=Sum('earning_amount'))['total'] or Decimal('0.00')

        month_data['adjustment']['adjusted_pensionable'] += safe_dec(pensionable_sum)

        for ea in payroll.earning_adjustments.all():
            #
            c = ea.component
            month_data['adjustment']['earning_adj_by_component']['taxable'][c] += safe_dec(ea.taxable)
            month_data['adjustment']['earning_adj_by_component']['non_taxable'][c] += safe_dec(ea.non_taxable)
            month_data['adjustment']['earning_adj_by_component']['total_earning_adjustment'][c] += safe_dec(
                ea.earning_amount)
            month_data['adjustment']['earning_adj_by_component']['employee_pension_contribution'][c] += safe_dec(
                ea.employee_pension_contribution)
            month_data['adjustment']['earning_adj_by_component']['employer_pension_contribution'][c] += safe_dec(
                ea.employer_pension_contribution)
            month_data['adjustment']['earning_adj_by_component']['total_pension'][c] += safe_dec(ea.total_pension)

        # Aggregate overall adjustment totals
        earning_adj_first = payroll.earning_adjustments.first() or type('Empty', (), {})()

        month_data['adjustment']['taxable_gross'] += safe_dec(
            getattr(earning_adj_first, 'recorded_month_taxable_gross_pay', 0))
        month_data['adjustment']['non_taxable_gross'] += safe_dec(
            getattr(earning_adj_first, 'recorded_month_non_taxable_gross_pay', 0))
        month_data['adjustment']['gross'] += safe_dec(
            getattr(earning_adj_first, 'recorded_month_gross_pay', 0))

        month_data['adjustment']['employer_pension'] += safe_dec(getattr(earning_adj_first, 'recorded_month_employer_pension_contribution', 0))

        month_data['adjustment']['employee_pension'] += safe_dec(
            getattr(earning_adj_first, 'recorded_month_employee_pension_contribution', 0))

        month_data['adjustment']['total_pension'] += safe_dec(
            getattr(earning_adj_first, 'recorded_month_total_pension_contribution', 0))
        month_data['adjustment']['employment_income_tax'] += safe_dec(
            getattr(earning_adj_first, 'recorded_month_employment_income_tax', 0))
        month_data['adjustment']['earning_adjustment_deduction'] += safe_dec(
            getattr(earning_adj_first, 'recorded_month_total_earning_deduction', 0))
        month_data['adjustment']['expense'] += safe_dec(
            getattr(earning_adj_first, 'recorded_month_expense', 0))

        # Aggregate deduction adjustments by component (detailed)
        for da in payroll.deduction_adjustments.all():
            c = da.component
            month_data['adjustment']['deduction_adj_by_component'][c] += safe_dec(getattr(da, 'deduction_amount', 0))

        # Aggregate total adjustment deduction (summary)
        deduction_adj_first = payroll.deduction_adjustments.first() or type('Empty', (), {})()
        month_data['adjustment']['individual_adjustment_deduction'] += safe_dec(
            getattr(deduction_adj_first, 'recorded_month_total_deduction', 0)
        )

    # Derive net monthly adjustment
    # After summing all payrolls for the month, calculate net
    for month_data in monthly_summary.values():
        month_data['adjustment']['total_monthly_adjustment_deduction'] = (
                month_data['adjustment']['earning_adjustment_deduction'] + month_data['adjustment']['individual_adjustment_deduction']
        )
        month_data['adjustment']['net_monthly_adjustment'] = (
                month_data['adjustment']['gross'] - month_data['adjustment']['total_monthly_adjustment_deduction']
        )

    #severance
    severances = SeverancePay.objects.filter(organization_name=request.user.organization_name)
    for sev in severances:
        key = sev.severance_record_month.payroll_month
        month_data = monthly_summary[key]

        gross = safe_dec(sev.gross_severance_pay)
        withholding = safe_dec(sev.employment_income_tax_from_severance_pay)
        net = safe_dec(sev.net_severance_pay)

        month_data['severance']['taxable_gross'] += gross
        month_data['severance']['gross'] += gross
        month_data['severance']['employment_income_tax'] += withholding
        month_data['severance']['total_severance_deduction'] += withholding
        month_data['severance']['net'] += net
        month_data['severance']['expense'] += gross

    # Calculate totals per month
    for summary in monthly_summary.values():
        summary['totals']['taxable_gross'] = (
                summary['regular']['taxable_gross'] + summary['adjustment']['taxable_gross'] + summary['severance'][
            'taxable_gross']
        )
        summary['totals']['non_taxable_gross'] = (
                summary['regular']['non_taxable_gross'] + summary['adjustment']['non_taxable_gross']
        )
        summary['totals']['gross'] = (
                summary['regular']['gross'] + summary['adjustment']['gross'] + summary['severance']['gross']
        )
        summary['totals']['pensionable'] = (
                summary['regular']['pensionable'] + summary['adjustment']['adjusted_pensionable']
        )
        summary['totals']['employee_pension'] = (
                summary['regular']['employee_pension'] + summary['adjustment']['employee_pension']
        )
        summary['totals']['employer_pension'] = (
                summary['regular']['employer_pension'] + summary['adjustment']['employer_pension']
        )
        summary['totals']['total_pension'] = (
                summary['regular']['total_pension'] + summary['adjustment']['total_pension']
        )
        summary['totals']['employment_income_tax'] = (
                summary['regular']['employment_income_tax'] + summary['adjustment']['employment_income_tax'] +
                summary['severance']['employment_income_tax']
        )
        summary['totals']['total_deduction'] = (
                summary['regular']['total_regular_deduction'] + summary['adjustment']['total_monthly_adjustment_deduction'] +
                summary['severance']['total_severance_deduction']
        )
        summary['totals']['expense'] = (
                summary['regular']['expense'] + summary['adjustment']['expense'] + summary['severance']['expense']
        )
        summary['totals']['final_net_pay'] = (
                summary['totals']['gross'] - summary['totals']['total_deduction']
        )

    # order by month for regular with severance based order
    def parse_key(key):
        try:
            month_str, year_str = key.split('-')
            year = int(year_str)
            month_order_full = {
                'September': '01', 'October': '02', 'November': '03', 'December': '04', 'January': '05',
                'February': '06', 'March': '07', 'April': '08', 'May': '09', 'June': '10', 'July': '11', 'August': '12'
            }
            month_num_str = month_order_full.get(month_str.capitalize(), '0')
            month_num = int(month_num_str)  # convert string month number to int
            return (year, month_num)
        except Exception:
            return (0, 0)

    final_monthly_list = []

    # for month order
    for month_key, data in sorted(
            monthly_summary.items(),
            key=lambda item: parse_key(item[0]),
            reverse=True  # latest first
    ):
        # format
        regular_components_fmt = {
            format_key(k): v for k, v in data['regular']['components'].items()
        }
        #
        earning_adj_comp_fmt = {
            format_key(comp): {
                'taxable': data['adjustment']['earning_adj_by_component']['taxable'][comp],
                'non_taxable': data['adjustment']['earning_adj_by_component']['non_taxable'][comp],
                'earning_amount': data['adjustment']['earning_adj_by_component']['total_earning_adjustment'][comp],
                'employee_pension_contribution':
                    data['adjustment']['earning_adj_by_component']['employee_pension_contribution'][comp],
                'employer_pension_contribution':
                    data['adjustment']['earning_adj_by_component']['employer_pension_contribution'][comp],
                'total_pension': data['adjustment']['earning_adj_by_component']['total_pension'][comp],

            }
            for comp in data['adjustment']['earning_adj_by_component']['total_earning_adjustment']
        }
        #
        deduction_adj_comp_fmt = {
            format_key(k): v for k, v in data['adjustment']['deduction_adj_by_component'].items()
        }
        #

        show_earning = has_nonzero_earning(earning_adj_comp_fmt)
        show_deduction = has_nonzero_deduction(deduction_adj_comp_fmt)

        regular_total = sum(regular_components_fmt.values())

        final_monthly_list.append({
            'month': month_key,
            'regular': data['regular'],
            'adjustment': data['adjustment'],
            'severance': data['severance'],
            'totals': data['totals'],
            'regular_item_by_component': regular_components_fmt,
            'earning_adj_by_component': earning_adj_comp_fmt,
            'deduction_adj_by_component': deduction_adj_comp_fmt,
            'show_earning': show_earning,
            'show_deduction': show_deduction,
            'regular_total': regular_total,
        })

    paginator = Paginator(final_monthly_list, 10)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    return {
        'page_obj': page_obj,
        'monthly_summary': monthly_summary,

    }





