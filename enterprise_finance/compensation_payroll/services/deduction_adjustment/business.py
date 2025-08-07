from decimal import Decimal
from django.db.models import Sum
from django.core.paginator import Paginator
from django.db.models import Q, Count


class DeductionAdjustmentBusinessService:

    def __init__(self, instance=None):
        self.instance = instance

    def perform_all_calculations(self):
        self.calculate_deduction_adjustment_per_adjusted_month()
        self.calculate_monthly_deduction_adjustment()

    def calculate_deduction_adjustment_per_adjusted_month(self):
        if not self.instance or not self.instance.record_month:
            return

        from compensation_payroll.models import DeductionAdjustment
        record_month = self.instance.record_month
        payroll_needing_adjustment = self.instance.payroll_needing_adjustment

        # Get all deduction adjustments for same month and person
        deductions = DeductionAdjustment.objects.filter(
            record_month=record_month,
            payroll_needing_adjustment=payroll_needing_adjustment
        ).select_related(
            'record_month',
            'payroll_needing_adjustment'
        )

        # Calculate total deduction
        total_deduction = deductions.aggregate(
            total=Sum('deduction_amount')
        )['total'] or Decimal('0.00')

        # Update adjusted_deduction for all matching deductions
        deductions.update(
            adjusted_deduction_per_adjusted_month=total_deduction
        )

    def calculate_monthly_deduction_adjustment(self):
        if not self.instance or not self.instance.record_month:
            return

        from decimal import Decimal
        from compensation_payroll.models import DeductionAdjustment

        record_month = self.instance.record_month

        # Get all deductions for the current record_month and personnel
        deductions = DeductionAdjustment.objects.filter(
            record_month=record_month
        ).select_related('payroll_needing_adjustment')

        # Use payroll_needing_adjustment.payroll_month to ensure one per adjusted month
        unique_deductions_by_month = {}
        for deduction in deductions:
            if deduction.payroll_needing_adjustment is None:
                continue
            adjusted_month = deduction.payroll_needing_adjustment.payroll_month
            if adjusted_month not in unique_deductions_by_month:
                unique_deductions_by_month[adjusted_month] = deduction

        # Aggregate across adjusted months
        total_adjusted_deduction = Decimal('0.00')

        for deduction in unique_deductions_by_month.values():
            total_adjusted_deduction += deduction.adjusted_deduction_per_adjusted_month or Decimal('0.00')

        # Update all deductions with the monthly total
        deductions.update(
            monthly_adjusted_deduction=total_adjusted_deduction
        )


