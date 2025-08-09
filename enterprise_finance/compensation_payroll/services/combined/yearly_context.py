from decimal import Decimal
from collections import defaultdict
from compensation_payroll.models import RegularPayroll, SeverancePay
from django.core.paginator import Paginator
from django.db.models import Sum


def get_combined_yearly_detail(request):

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

    payrolls = RegularPayroll.objects.filter(
        organization_name=request.user.organization_name
    ).select_related('payroll_month') \
        .prefetch_related('earning_adjustments', 'deduction_adjustments')
    payrolls = payrolls.order_by('-payroll_month__year')

    #
    yearly_summary = defaultdict(lambda: {
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
            'total_adjustment_deduction': Decimal('0.00'),
            'expense': Decimal('0.00'),
            'earning_adj_by_component': {
                'taxable': defaultdict(lambda: Decimal('0.00')),
                'non_taxable': defaultdict(lambda: Decimal('0.00')),
                'total_earning_adjustment': defaultdict(lambda: Decimal('0.00')),
                'employee_pension_contribution': defaultdict(lambda: Decimal('0.00')),
                'employer_pension_contribution': defaultdict(lambda: Decimal('0.00')),
                'total_pension': defaultdict(lambda: Decimal('0.00')),
            },
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

    for payroll in payrolls:
        if payroll and payroll.payroll_month:
            year_key = payroll.payroll_month.year
        else:
            # handle None case — skip, set default, or log warning
            year_key = None  # or continue/skip processing this record

        # year_key = payroll.payroll_month.year
        year_data = yearly_summary[year_key]

        year_data['regular']['taxable_gross'] += safe_dec(payroll.gross_taxable_pay)
        year_data['regular']['non_taxable_gross'] += safe_dec(payroll.gross_non_taxable_pay)
        year_data['regular']['gross'] += safe_dec(payroll.gross_pay)
        year_data['regular']['pensionable'] += safe_dec(payroll.basic_salary)
        year_data['regular']['employee_pension'] += safe_dec(payroll.employee_pension_contribution)
        year_data['regular']['employer_pension'] += safe_dec(payroll.employer_pension_contribution)
        year_data['regular']['total_pension'] += safe_dec(payroll.total_pension_contribution)
        year_data['regular']['employment_income_tax'] += safe_dec(payroll.employment_income_tax)
        year_data['regular']['total_regular_deduction'] += safe_dec(payroll.total_payroll_deduction)
        year_data['regular']['net_pay'] += safe_dec(payroll.net_pay)
        year_data['regular']['expense'] += safe_dec(payroll.expense)

        # Calculate pensionable sum once per payroll
        pensionable_sum = payroll.earning_adjustments.filter(
            component='basic_salary'
        ).aggregate(total=Sum('earning_amount'))['total'] or Decimal('0.00')

        adjusted_pensionable = safe_dec(pensionable_sum)

        # Add it once to monthly summary
        year_data['adjustment']['adjusted_pensionable'] += adjusted_pensionable

        for ea in payroll.earning_adjustments.all():
            c = ea.component
            year_data['adjustment']['earning_adj_by_component']['taxable'][c] += safe_dec(ea.taxable)
            year_data['adjustment']['earning_adj_by_component']['non_taxable'][c] += safe_dec(ea.non_taxable)
            year_data['adjustment']['earning_adj_by_component']['total_earning_adjustment'][c] += safe_dec(ea.earning_amount)
            year_data['adjustment']['earning_adj_by_component']['employee_pension_contribution'][c] += safe_dec(ea.employee_pension_contribution)
            year_data['adjustment']['earning_adj_by_component']['employer_pension_contribution'][c] += safe_dec(ea.employer_pension_contribution)
            year_data['adjustment']['earning_adj_by_component']['total_pension'][c] += safe_dec(ea.total_pension)


        for da in payroll.deduction_adjustments.all():
            c = da.component
            year_data['adjustment']['deduction_adj_by_component'][c] += safe_dec(getattr(da, 'deduction_amount', 0))

        ea_first = payroll.earning_adjustments.first() or type('Empty', (), {})()
        #
        year_data['adjustment']['taxable_gross'] += safe_dec(getattr(ea_first, 'recorded_month_taxable_gross_pay', 0))
        year_data['adjustment']['non_taxable_gross'] += safe_dec(getattr(ea_first, 'recorded_month_non_taxable_gross_pay', 0))
        year_data['adjustment']['gross'] += safe_dec(getattr(ea_first, 'recorded_month_gross_pay', 0))
        year_data['adjustment']['adjusted_pensionable'] += safe_dec(getattr(ea_first, 'recorded_month_adjusted_pensionable', 0))
        year_data['adjustment']['employee_pension'] += safe_dec(getattr(ea_first, 'recorded_month_employee_pension_contribution', 0))
        year_data['adjustment']['employer_pension'] += safe_dec(getattr(ea_first, 'recorded_month_employer_pension_contribution', 0))
        year_data['adjustment']['total_pension'] += safe_dec(getattr(ea_first, 'recorded_month_total_pension_contribution', 0))
        year_data['adjustment']['employment_income_tax'] += safe_dec(getattr(ea_first, 'recorded_month_employment_income_tax', 0))
        year_data['adjustment']['total_adjustment_deduction'] += safe_dec(getattr(ea_first, 'recorded_month_total_earning_deduction', 0))
        year_data['adjustment']['expense'] += safe_dec(getattr(ea_first, 'recorded_month_expense', 0))

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
            year_data['regular']['components'][comp] += safe_dec(val)

    severances = SeverancePay.objects.filter(organization_name=request.user.organization_name)
    for sev in severances:
        year_key = str(sev.year)
        year_data = yearly_summary[year_key]

        gross = safe_dec(sev.gross_severance_pay)
        withholding = safe_dec(sev.employment_income_tax_from_severance_pay)
        net = safe_dec(sev.net_severance_pay)

        year_data['severance']['taxable_gross'] += gross
        year_data['severance']['gross'] += gross
        year_data['severance']['employment_income_tax'] += withholding
        year_data['severance']['total_severance_deduction'] += withholding
        year_data['severance']['net'] += net
        year_data['severance']['expense'] += gross

    for summary in yearly_summary.values():
        summary['totals']['taxable_gross'] = (
            summary['regular']['taxable_gross'] + summary['adjustment']['taxable_gross'] + summary['severance']['taxable_gross']
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
            summary['regular']['employment_income_tax'] + summary['adjustment']['employment_income_tax'] + summary['severance']['employment_income_tax']
        )
        summary['totals']['total_deduction'] = (
            summary['regular']['total_regular_deduction'] + summary['adjustment']['total_adjustment_deduction'] + summary['severance']['total_severance_deduction']
        )
        summary['totals']['expense'] = (
            summary['regular']['expense'] + summary['adjustment']['expense'] + summary['severance']['expense']
        )
        summary['totals']['final_net_pay'] = (
            summary['totals']['gross'] - summary['totals']['total_deduction']
        )

    final_yearly_list = []
    filtered_items = [(k, v) for k, v in yearly_summary.items() if k is not None]
    for year_key, data in sorted(filtered_items, reverse=True):
        #
        regular_components_fmt = {format_key(k): v for k, v in data['regular']['components'].items()}
        #
        earning_adj_comp_fmt = {
            format_key(comp): {
                'earning_amount': data['adjustment']['earning_adj_by_component']['total_earning_adjustment'][comp],
                'taxable': data['adjustment']['earning_adj_by_component']['taxable'][comp],
                'non_taxable': data['adjustment']['earning_adj_by_component']['non_taxable'][comp],
                'employee_pension_contribution': data['adjustment']['earning_adj_by_component']['employee_pension_contribution'][comp],
                'employer_pension_contribution': data['adjustment']['earning_adj_by_component']['employer_pension_contribution'][comp],
                'total_pension': data['adjustment']['earning_adj_by_component']['total_pension'][comp],
            }
            for comp in data['adjustment']['earning_adj_by_component']['total_earning_adjustment']
        }
        #
        deduction_adj_comp_fmt = {
            format_key(k): v for k, v in data['adjustment']['deduction_adj_by_component'].items()
        }

        show_earning = has_nonzero_earning(earning_adj_comp_fmt)
        show_deduction = has_nonzero_deduction(deduction_adj_comp_fmt)
        regular_total = sum(regular_components_fmt.values())

        final_yearly_list.append({
            'year': year_key,
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

    paginator = Paginator(final_yearly_list, 10)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    return {
        'page_obj': page_obj,
        'yearly_summary': yearly_summary,  # ← Add this line

    }

