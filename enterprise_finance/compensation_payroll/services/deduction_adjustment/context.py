from decimal import Decimal
from django.db.models import Q, Count, Sum
from django.core.paginator import Paginator

def get_deduction_adjustment_context(request):
    from compensation_payroll.models import DeductionAdjustment

    deduction_adjustments = DeductionAdjustment.objects.filter(
        organization_name=request.user.organization_name
    ).select_related(
        'payroll_to_record__personnel_full_name',
        'payroll_needing_adjustment__personnel_full_name',
        'payroll_to_record__payroll_month',
        'payroll_needing_adjustment__payroll_month',
    )

    search_query = request.GET.get('search', '')
    if search_query:
        deduction_adjustments = deduction_adjustments.filter(
            Q(payroll_to_record__payroll_month__payroll_month__payroll_month__icontains=search_query) |
            Q(payroll_to_record__personnel_full_name__personnel_id__icontains=search_query) |
            Q(payroll_to_record__personnel_full_name__first_name__icontains=search_query) |
            Q(payroll_to_record__personnel_full_name__father_name__icontains=search_query) |
            Q(payroll_to_record__personnel_full_name__last_name__icontains=search_query)
        )

    deduction_adjustments = deduction_adjustments.order_by(
        '-payroll_to_record__payroll_month__payroll_month__year',
        '-payroll_to_record__payroll_month__payroll_month__month'
    )

    paginator = Paginator(deduction_adjustments, 10)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    deduction_per_adjusted_month = (
        deduction_adjustments
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

            'adjusted_month_total_deduction',
        )
        .annotate(record_count=Count('id'))
        .order_by(
            '-payroll_to_record__payroll_month__payroll_month__year',
        '-payroll_to_record__payroll_month__payroll_month__month'
        )
    )

    monthly_deduction_adjustment = (
        deduction_adjustments
        .values(
            'payroll_to_record__payroll_month__payroll_month__payroll_month',
            'payroll_to_record__personnel_full_name__personnel_id',
            'payroll_to_record__personnel_full_name__first_name',
            'payroll_to_record__personnel_full_name__father_name',
            'payroll_to_record__personnel_full_name__last_name',

            'recorded_month_total_deduction',
        )
        .annotate(record_count=Count('id'))
        .order_by(
            '-payroll_to_record__payroll_month__payroll_month__year',
        '-payroll_to_record__payroll_month__payroll_month__month'
        )
    )

    # Aggregate per month
    monthly_deduction_adjustment_aggregate = (
        deduction_adjustments
        .values('payroll_to_record__payroll_month__payroll_month__payroll_month')
        .annotate(
            total_deduction=Sum('adjusted_month_total_deduction'),
            total_record_count=Count('id'),
        )
        .order_by(
            '-payroll_to_record__payroll_month__payroll_month__year',
            '-payroll_to_record__payroll_month__payroll_month__month'
        )
    )

    return {
        'page_obj': page_obj,
        'search_query': search_query,
        'deduction_adjustments': deduction_adjustments,
        'deduction_per_adjusted_month': deduction_per_adjusted_month,
        'monthly_deduction_adjustment': monthly_deduction_adjustment,
        'monthly_deduction_adjustment_aggregate': monthly_deduction_adjustment_aggregate,
    }
