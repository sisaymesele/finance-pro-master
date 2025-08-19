from django.core.paginator import Paginator
from django.db.models import Q
from compensation_payroll.models import RegularPayroll
from django.db.models import Sum
from collections import defaultdict
from decimal import Decimal


def get_combined_personnel_payroll_context(request):

    def has_nonzero_earning(earning_adj_by_component):
        for v in earning_adj_by_component.values():
            if v['earning_amount'] != Decimal('0.00'):
                return True
        return False

    def has_nonzero_deduction(deduction_adj_by_component):
        for v in deduction_adj_by_component.values():
            if v != Decimal('0.00'):
                return True
        return False

    def format_key(key):
        return key.replace('_', ' ').title()

    def safe_dec(value):
        return value if value is not None else Decimal('0.00')

    payrolls = RegularPayroll.objects.filter(
        organization_name=request.user.organization_name
    ).select_related('personnel_full_name', 'payroll_month') \
     .prefetch_related('earning_adjustments', 'deduction_adjustments')

    search_query = request.GET.get('search', '')
    if search_query:
        payrolls = payrolls.filter(
            Q(payroll_month__payroll_month__payroll_month__icontains=search_query) |
            Q(personnel_full_name__personnel_id__icontains=search_query) |
            Q(personnel_full_name__first_name__icontains=search_query) |
            Q(personnel_full_name__father_name__icontains=search_query) |
            Q(personnel_full_name__last_name__icontains=search_query)
        )

    payrolls = payrolls.order_by(
        '-payroll_month__payroll_month__year',
        '-payroll_month__payroll_month__month'
    )
    paginator = Paginator(payrolls, 15)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    payroll_data = []

    for payroll in page_obj:
        #regular
        regular = {
            'pensionable': safe_dec(payroll.basic_salary),
            'taxable_gross': safe_dec(payroll.gross_taxable_pay),
            'non_taxable_gross': safe_dec(payroll.gross_non_taxable_pay),
            'gross_pay': safe_dec(payroll.gross_pay),
            'employee_pension': safe_dec(payroll.employee_pension_contribution),
            'employer_pension': safe_dec(payroll.employer_pension_contribution),
            'total_pension': safe_dec(payroll.total_pension_contribution),
            'employment_income_tax': safe_dec(payroll.employment_income_tax),
            'deduction': safe_dec(payroll.total_payroll_deduction),
            'net_pay': safe_dec(payroll.net_pay),
            'expense': safe_dec(payroll.expense),
        }

        regular_item_by_component = {

            'basic_salary': safe_dec(payroll.basic_salary),
            'overtime': safe_dec(payroll.overtime),
            'housing_allowance': safe_dec(payroll.housing_allowance),
            'position_allowance': safe_dec(payroll.position_allowance),
            'commission': safe_dec(payroll.commission),
            'telephone_allowance': safe_dec(payroll.telephone_allowance),
            'one_time_bonus': safe_dec(payroll.one_time_bonus),
            'causal_labor_wage': safe_dec(payroll.causal_labor_wage),

            'transport_home_to_office_taxable': safe_dec(payroll.transport_home_to_office_taxable),
            'transport_home_to_office_non_taxable': safe_dec(payroll.transport_home_to_office_non_taxable),

            'fuel_home_to_office_taxable': safe_dec(payroll.fuel_home_to_office_taxable),
            'fuel_home_to_office_non_taxable': safe_dec(payroll.fuel_home_to_office_non_taxable),

            'transport_for_work_taxable': safe_dec(payroll.transport_for_work_taxable),
            'transport_for_work_non_taxable': safe_dec(payroll.transport_for_work_non_taxable),

            'fuel_for_work_taxable': safe_dec(payroll.fuel_for_work_taxable),
            'fuel_for_work_non_taxable': safe_dec(payroll.fuel_for_work_non_taxable),

            'per_diem_taxable': safe_dec(payroll.per_diem_taxable),
            'per_diem_non_taxable': safe_dec(payroll.per_diem_non_taxable),

            'hardship_allowance_taxable': safe_dec(payroll.hardship_allowance_taxable),
            'hardship_allowance_non_taxable': safe_dec(payroll.hardship_allowance_non_taxable),

            'public_cash_award': safe_dec(payroll.public_cash_award),
            'incidental_operation_allowance': safe_dec(payroll.incidental_operation_allowance),
            'medical_allowance': safe_dec(payroll.medical_allowance),
            'cash_gift': safe_dec(payroll.cash_gift),
            'tuition_fees': safe_dec(payroll.tuition_fees),
            'personal_injury': safe_dec(payroll.personal_injury),
            'child_support_payment': safe_dec(payroll.child_support_payment),

            'charitable_donation': safe_dec(payroll.charitable_donation),
            'saving_plan': safe_dec(payroll.saving_plan),
            'loan_payment': safe_dec(payroll.loan_payment),
            'court_order': safe_dec(payroll.court_order),
            'workers_association': safe_dec(payroll.workers_association),
            'personnel_insurance_saving': safe_dec(payroll.personnel_insurance_saving),

            'cost_share_percent_to_basic_salary': safe_dec(payroll.cost_share_percent_to_basic_salary),
            'university_cost_share_pay': safe_dec(payroll.university_cost_share_pay),
            'red_cross': safe_dec(payroll.red_cross),
            'party_contribution': safe_dec(payroll.party_contribution),
            'other_deduction': safe_dec(payroll.other_deduction),
        }
        #adjustment
        earning_adj = payroll.earning_adjustments.order_by(
            '-payroll_to_record__payroll_month__payroll_month__year',
            '-payroll_to_record__payroll_month__payroll_month__month'
        ).first() or type('Empty', (), {})()

        deduction_adj = payroll.deduction_adjustments.order_by(
            '-payroll_to_record__payroll_month__payroll_month__year',
            '-payroll_to_record__payroll_month__payroll_month__month'
        ).first() or type('Empty', (), {})()
         #get total data from deduction adjustment
        adjusted_deduction = safe_dec(getattr(deduction_adj, 'recorded_month_total_deduction', 0))

        #pensionable from adjustment
        pensionable_sum = payroll.earning_adjustments.filter(
            component='basic_salary',
        ).aggregate(total=Sum('earning_amount'))['total'] or Decimal('0.00')

        adjusted_pensionable = safe_dec(pensionable_sum)

        earning_adjustment = {
            'taxable_gross': safe_dec(getattr(earning_adj, 'recorded_month_taxable_gross_pay', 0)),
            'non_taxable_gross': safe_dec(getattr(earning_adj, 'recorded_month_non_taxable_gross_pay', 0)),
            'gross_pay': safe_dec(getattr(earning_adj, 'recorded_month_gross_pay', 0)),
            'adjusted_pensionable': adjusted_pensionable,
            'employee_pension': safe_dec(getattr(earning_adj, 'recorded_month_employee_pension_contribution', 0)),
            'employer_pension': safe_dec(getattr(earning_adj, 'recorded_month_employer_pension_contribution', 0)),
            'total_pension': safe_dec(getattr(earning_adj, 'recorded_month_total_pension_contribution', 0)),
            'employment_income_tax': safe_dec(getattr(earning_adj, 'recorded_month_employment_income_tax', 0)),
            'earning_adjustment_deduction': safe_dec(getattr(earning_adj, 'recorded_month_total_earning_deduction', 0)),
            'adjusted_deduction': adjusted_deduction,
            'expense': safe_dec(getattr(earning_adj, 'recorded_month_expense', 0)),
        }

        combined_adjustment = {
            'total_adjustment_deduction': earning_adjustment['earning_adjustment_deduction'] + adjusted_deduction,
            'net_monthly_adjustment': earning_adjustment['gross_pay'] - earning_adjustment['earning_adjustment_deduction'] - adjusted_deduction,
        }

        # Prepare two dictionaries: one for taxable and one for non-taxable
        earning_adj_by_component = {
            'taxable': defaultdict(lambda: Decimal('0.00')),
            'non_taxable': defaultdict(lambda: Decimal('0.00')),
            'total_earning_adjustment': defaultdict(lambda: Decimal('0.00')),
            'employee_pension_contribution': defaultdict(lambda: Decimal('0.00')),
            'employer_pension_contribution': defaultdict(lambda: Decimal('0.00')),
            'total_pension': defaultdict(lambda: Decimal('0.00')),
        }

        # Loop through adjustments and fill each part
        for ea in payroll.earning_adjustments.all():
            component = ea.component
            earning_adj_by_component['taxable'][component] += safe_dec(ea.taxable)
            earning_adj_by_component['non_taxable'][component] += safe_dec(ea.non_taxable)
            earning_adj_by_component['total_earning_adjustment'][component] += safe_dec(ea.earning_amount)
            earning_adj_by_component['employee_pension_contribution'][component] += safe_dec(ea.employee_pension_contribution)
            earning_adj_by_component['employer_pension_contribution'][component] += safe_dec(ea.employer_pension_contribution)
            earning_adj_by_component['total_pension'][component] += safe_dec(ea.total_pension)


        deduction_adj_by_component = defaultdict(lambda: Decimal('0.00'))
        for da in payroll.deduction_adjustments.all():
            deduction_adj_by_component[da.component] += safe_dec(getattr(da, 'deduction_amount', 0))


        totals = {
            'taxable_gross': regular['taxable_gross'] + earning_adjustment['taxable_gross'],
            'non_taxable_gross': regular['non_taxable_gross'] + earning_adjustment['non_taxable_gross'],
            'gross_pay': regular['gross_pay'] + earning_adjustment['gross_pay'],
            'pensionable': regular['pensionable'] + earning_adjustment['adjusted_pensionable'],
            'employee_pension': regular['employee_pension'] + earning_adjustment['employee_pension'],
            'employer_pension': regular['employer_pension'] + earning_adjustment['employer_pension'],
            'total_pension': regular['total_pension'] + earning_adjustment['total_pension'],
            'employment_income_tax': regular['employment_income_tax'] + earning_adjustment['employment_income_tax'],
            'regular_deduction': regular['deduction'],
            'adjustment_deduction': earning_adjustment['earning_adjustment_deduction'] + adjusted_deduction,
            'deduction': regular['deduction'] + earning_adjustment['earning_adjustment_deduction'] + adjusted_deduction,
            'expense': regular['expense'] + earning_adjustment['expense'],
        }

        totals['final_net_pay'] = totals['gross_pay'] - totals['deduction']

        # Format the keys for display
        regular_item_by_component_formatted = {
            format_key(k): v for k, v in regular_item_by_component.items()
        }
        #

        earning_adj_by_component_formatted = {
            format_key(component): {
                'taxable': earning_adj_by_component['taxable'].get(component, 0),
                'non_taxable': earning_adj_by_component['non_taxable'].get(component, 0),
                'earning_amount': earning_adj_by_component['total_earning_adjustment'].get(component, 0),
                'employee_pension_contribution': earning_adj_by_component['employee_pension_contribution'].get(component, 0),
                'employer_pension_contribution': earning_adj_by_component['employer_pension_contribution'].get(component, 0),
                'total_pension': earning_adj_by_component['total_pension'].get(component, 0),
            }
            for component in earning_adj_by_component['total_earning_adjustment']
        }
        #

        deduction_adj_by_component_formatted = {
            format_key(k): v for k, v in dict(deduction_adj_by_component).items()
        }

        show_earning = has_nonzero_earning(earning_adj_by_component_formatted)
        show_deduction = has_nonzero_deduction(deduction_adj_by_component_formatted)


        payroll_data.append({
            'payroll': payroll,
            'earning_adjustment': earning_adj,
            'deduction_adjustment': deduction_adj,
            'combined_adjustment': combined_adjustment,
            'earning_adjustment_item': earning_adjustment,
            'regular_totals': regular,
            'totals': totals,
            #
            'regular_item_by_component': regular_item_by_component_formatted,
            'earning_adj_by_component': earning_adj_by_component_formatted,
            'deduction_adj_by_component': deduction_adj_by_component_formatted,
             #
            'show_earning': show_earning,
            'show_deduction': show_deduction,
        })

    return {
        'page_obj': page_obj,
        'search_query': search_query,
        'payroll_data': payroll_data,
    }
#
# #
