# services/earning_adjustment_context.py
#view and excel
from django.core.paginator import Paginator
from django.db.models import Q, Count, Sum, F



def get_earning_adjustment_context(request):
    from compensation_payroll.models import EarningAdjustment

    earning_adjustments = EarningAdjustment.objects.filter(
        organization_name=request.user.organization_name
    ).select_related(
        'payroll_to_record__personnel_full_name',
        'payroll_needing_adjustment__personnel_full_name',
        'payroll_to_record__payroll_month',
        'payroll_needing_adjustment__payroll_month',
    )

    search_query = request.GET.get('search', '')
    if search_query:
        earning_adjustments = earning_adjustments.filter(
            Q(payroll_to_record__payroll_month__payroll_month__payroll_month__icontains=search_query) |
            Q(payroll_to_record__personnel_full_name__personnel_id__icontains=search_query) |
            Q(payroll_to_record__personnel_full_name__first_name__icontains=search_query) |
            Q(payroll_to_record__personnel_full_name__father_name__icontains=search_query) |
            Q(payroll_to_record__personnel_full_name__last_name__icontains=search_query)
        )

    earning_adjustments = earning_adjustments.order_by(
        '-payroll_to_record__payroll_month__payroll_month__year',
        '-payroll_to_record__payroll_month__payroll_month__month'
    )

    paginator = Paginator(earning_adjustments, 10)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    earning_per_adjusted_month = (
        earning_adjustments
        .values(
            'payroll_to_record__payroll_month__payroll_month__payroll_month',
            'payroll_to_record__personnel_full_name__personnel_id',
            'payroll_to_record__personnel_full_name__first_name',
            'payroll_to_record__personnel_full_name__father_name',
            'payroll_to_record__personnel_full_name__last_name',

            'payroll_needing_adjustment__payroll_month__payroll_month__payroll_month',
            'payroll_needing_adjustment__personnel_full_name__personnel_id',
            'payroll_needing_adjustment__personnel_full_name__first_name',
            'payroll_needing_adjustment__personnel_full_name__father_name',
            'payroll_needing_adjustment__personnel_full_name__last_name',

            'payroll_needing_adjustment__gross_taxable_pay',
            'payroll_needing_adjustment__employment_income_tax',
            'adjusted_month_gross_taxable_pay',
            'adjusted_month_gross_non_taxable_pay',
            'adjusted_month_gross_pay',
            'adjusted_month_total_taxable_pay',
            'adjusted_month_employment_income_tax_total',
            'adjusted_month_employment_income_tax',
            'adjusted_month_total_earning_deduction',
            'adjusted_month_expense',

            'adjusted_month_employee_pension_contribution',
            'adjusted_month_employer_pension_contribution',
            'adjusted_month_total_pension',
        )
        .annotate(record_count=Count('id'))
        .order_by(
            '-payroll_to_record__payroll_month__payroll_month__year',
        '-payroll_to_record__payroll_month__payroll_month__month'
        )
    )

    #monthly per employee
    monthly_earning_adjustment = (
        earning_adjustments
        .values(
            'payroll_to_record__payroll_month__payroll_month__payroll_month',
            'payroll_to_record__personnel_full_name__personnel_id',
            'payroll_to_record__personnel_full_name__first_name',
            'payroll_to_record__personnel_full_name__father_name',
            'payroll_to_record__personnel_full_name__last_name',

            'recorded_month_taxable_gross_pay',
            'recorded_month_non_taxable_gross_pay',
            'recorded_month_gross_pay',
            'recorded_month_total_taxable_pay',
            'recorded_month_employment_income_tax_total',
            'recorded_month_employment_income_tax',
            'recorded_month_total_earning_deduction',
            'recorded_month_expense',

            'recorded_month_employee_pension_contribution',
            'recorded_month_employer_pension_contribution',
            'recorded_month_total_pension_contribution',
        )
        .annotate(record_count=Count('id'))
        .order_by(
            '-payroll_to_record__payroll_month__payroll_month__year',
        '-payroll_to_record__payroll_month__payroll_month__month'
        )
    )
    # Sort data keys by year and month ascending

    monthly_earning_adjustment_aggregate = (
        earning_adjustments
        .values('payroll_to_record__payroll_month__payroll_month__payroll_month')
        .annotate(
            adjusted_gross_taxable_pay=Sum('adjusted_month_gross_taxable_pay'),
            adjusted_gross_non_taxable_pay=Sum('adjusted_month_gross_non_taxable_pay'),
            adjusted_gross_pay=Sum('adjusted_month_gross_pay'),
            adjusted_taxable_pay_total=Sum('adjusted_month_total_taxable_pay'),
            adjusted_income_tax_total=Sum('adjusted_month_employment_income_tax_total'),
            adjusted_income_tax=Sum('adjusted_month_employment_income_tax'),
            earning_deduction=Sum('adjusted_month_total_earning_deduction'),
            expense=Sum('adjusted_month_expense'),
            employee_pension=Sum('adjusted_month_employee_pension_contribution'),
            employer_pension=Sum('adjusted_month_employer_pension_contribution'),
            total_pension=Sum('adjusted_month_total_pension'),
            record_count=Count('id')
        )
        .order_by(
            '-payroll_to_record__payroll_month__payroll_month__year',
            '-payroll_to_record__payroll_month__payroll_month__month'
        )
    )


    return {
        'page_obj': page_obj,
        'search_query': search_query,
        'earning_adjustments': earning_adjustments,
        'earning_per_adjusted_month': earning_per_adjusted_month,
        'monthly_earning_adjustment': monthly_earning_adjustment,
        'monthly_earning_adjustment_aggregate': monthly_earning_adjustment_aggregate,
    }


