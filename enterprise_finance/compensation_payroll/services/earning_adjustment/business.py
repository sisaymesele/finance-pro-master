from decimal import Decimal
from django.db.models import Sum
from compensation_payroll.services.employment_income_tax import EmploymentIncomeTaxService
from compensation_payroll.services.pension import calculate_pension_contributions
from django.core.paginator import Paginator
from django.db.models import Q, Count
from compensation_payroll.services.constant_componet import FULLY_TAXABLE, PARTIALLY_TAXABLE, NON_TAXABLE, DEFERRED_EARNINGS, DEDUCTIONS, PENSIONABLE

class EarningAdjustmentBusinessService:

    def __init__(self, instance=None):
        self.instance = instance

    def perform_all_calculations(self):
        self.get_pension_contributions()
        taxable, non_taxable = self.classify_component()
        self.instance.taxable = taxable
        self.instance.non_taxable = non_taxable
        self.calculate_earning_adjustment_per_adjusted_month()
        self.calculate_monthly_earning_adjustment()

    def get_pension_contributions(self):
        personnel, employer, total = calculate_pension_contributions(self.instance)

        self.instance.employee_pension_contribution = personnel
        self.instance.employer_pension_contribution = employer
        self.instance.total_pension = total

        return personnel, employer, total

    def classify_component(self):
        """
        Classify the earning adjustment component as fully taxable, non-taxable, or partially taxable.
        Returns a tuple: (taxable_amount, non_taxable_amount)
        """
        component = self.instance.component
        #they should be used as global constants, not self.FULLY_TAXABLE.
        #
        if component in FULLY_TAXABLE or component in DEFERRED_EARNINGS:
            return self.handle_fully_taxable()

        elif component in NON_TAXABLE:
            return self.handle_fully_non_taxable()

        elif component in PARTIALLY_TAXABLE:
            return self.partial_component_taxability()


        # Default return if component not found
        return Decimal('0.00'), Decimal('0.00')

    def handle_fully_taxable(self):
        amount = Decimal(self.instance.earning_amount or 0)
        return amount, Decimal('0.00')

    def handle_fully_non_taxable(self):
        amount = Decimal(self.instance.earning_amount or 0)
        return Decimal('0.00'), amount

    def partial_component_taxability(self):

        component = self.instance.component

        from compensation_payroll.services.regular_payroll import RegularPayrollService

        service = RegularPayrollService(instance=self.instance)

        if component == 'transport_home_to_office':
            taxable, non_taxable = service.calculate_transport_home_to_office(amount=self.instance.earning_amount)

        elif component == 'fuel_home_to_office':
            taxable, non_taxable = service.calculate_fuel_home_to_office(amount=self.instance.earning_amount)

        elif component == 'transport_for_work':
            taxable, non_taxable = service.calculate_transport_for_work(amount=self.instance.earning_amount)

        elif component == 'fuel_for_work':
            taxable, non_taxable = service.calculate_fuel_for_work(amount=self.instance.earning_amount)

        elif component == 'per_diem':
            taxable, non_taxable = service.calculate_per_diem(amount=self.instance.earning_amount)

        elif component == 'hardship_allowance':
            taxable, non_taxable = service.calculate_hardship_allowance(amount=self.instance.earning_amount)

        else:
            taxable = non_taxable = Decimal('0.00')

        return taxable, non_taxable


    #per adjusted month adjustment
    def calculate_earning_adjustment_per_adjusted_month(self):
        from compensation_payroll.models import EarningAdjustment

        # ✅ Prevent AttributeError
        if not self.instance or not self.instance.payroll_to_record or not self.instance.payroll_needing_adjustment:
            return

        # Calculate pension only if component is pensionable
        if self.instance.component in PENSIONABLE:
            employee_pension, employer_pension, total_pension = calculate_pension_contributions(self.instance)
        else:
            employee_pension = employer_pension = total_pension = Decimal('0.00')

        # Save pension contributions on the instance
        self.instance.employee_pension_contribution = employee_pension
        self.instance.employer_pension_contribution = employer_pension
        self.instance.total_pension = total_pension

        payroll_to_record = self.instance.payroll_to_record
        payroll_needing_adjustment = self.instance.payroll_needing_adjustment

        adjustments = EarningAdjustment.objects.filter(
            payroll_to_record=payroll_to_record,
            payroll_needing_adjustment=payroll_needing_adjustment
        ).select_related(
            'payroll_to_record',
            'payroll_needing_adjustment',
        )

        totals = adjustments.aggregate(
            gross_taxable_sum=Sum('taxable'),
            gross_non_taxable_sum=Sum('non_taxable'),

            employee_pension_sum=Sum('employee_pension_contribution'),
            employer_pension_sum=Sum('employer_pension_contribution'),
            total_pension_sum=Sum('total_pension'),
        )

        gross_taxable = totals['gross_taxable_sum'] or Decimal('0.00')
        gross_non_taxable = totals['gross_non_taxable_sum'] or Decimal('0.00')
        gross_total = gross_taxable + gross_non_taxable

        employee_pension_total_per_adjusting = totals['employee_pension_sum'] or Decimal('0.00')
        employer_pension_total_per_adjusting = totals['employer_pension_sum'] or Decimal('0.00')
        total_pension_total_per_adjusting = totals['total_pension_sum'] or Decimal('0.00')

        # Safely get previous values
        previous_gross_taxable = payroll_needing_adjustment.gross_taxable_pay or Decimal('0.00')
        previous_employment_income_tax = payroll_needing_adjustment.employment_income_tax or Decimal('0.00')


        adjusted_month_total_taxable_pay = previous_gross_taxable + gross_taxable
        adjusted_month_employment_income_tax_total = EmploymentIncomeTaxService(
            adjusted_month_total_taxable_pay).calculate()

        adjusted_month_employment_income_tax = adjusted_month_employment_income_tax_total - previous_employment_income_tax
        #total deduction
        earning_adjustment_deduction_total_per_adjusting = (adjusted_month_employment_income_tax + employee_pension_total_per_adjusting)


        # Update all matching adjustments with aggregated sums including pension totals
        adjustments.update(
            adjusted_month_gross_taxable_pay=gross_taxable,
            adjusted_month_gross_non_taxable_pay=gross_non_taxable,
            adjusted_month_gross_pay=gross_total,
            adjusted_month_total_taxable_pay=adjusted_month_total_taxable_pay,
            adjusted_month_employment_income_tax_total=adjusted_month_employment_income_tax_total,
            adjusted_month_employment_income_tax=adjusted_month_employment_income_tax,
            #
            adjusted_month_employee_pension_contribution=employee_pension_total_per_adjusting,
            adjusted_month_employer_pension_contribution=employer_pension_total_per_adjusting,
            adjusted_month_total_pension=total_pension_total_per_adjusting,
            #
            adjusted_month_total_earning_deduction=earning_adjustment_deduction_total_per_adjusting,
            adjusted_month_expense=gross_total + employer_pension_total_per_adjusting,

        )

    #monthly adjustment

    def calculate_monthly_earning_adjustment(self):
        from compensation_payroll.models import EarningAdjustment

        # ✅ Prevent AttributeError
        if not self.instance or not self.instance.payroll_to_record or not self.instance.payroll_needing_adjustment:
            return

        payroll_to_record = self.instance.payroll_to_record

        adjustments = EarningAdjustment.objects.filter(
            payroll_to_record=payroll_to_record
        ).select_related('payroll_needing_adjustment')

        # Ensure one entry per adjusted month
        unique_adjustments_by_month = {}
        for adj in adjustments:
            if adj.payroll_needing_adjustment is None:
                continue
            adjusted_month = adj.payroll_needing_adjustment.payroll_month
            if adjusted_month not in unique_adjustments_by_month:
                unique_adjustments_by_month[adjusted_month] = adj

        total_taxable = Decimal('0.00')
        total_non_taxable = Decimal('0.00')
        total_gross = Decimal('0.00')
        total_taxable_combined = Decimal('0.00')
        total_employment_income_tax = Decimal('0.00')
        total_employment_income_tax_adjustment = Decimal('0.00')
        total_net_adjustment = Decimal('0.00')
        total_adjusted_expense = Decimal('0.00')

        # Pension totals
        total_employee_pension = Decimal('0.00')
        total_employer_pension = Decimal('0.00')
        total_combined_pension = Decimal('0.00')
        #
        total_earning_adjustment_deduction = Decimal('0.00')

        for adj in unique_adjustments_by_month.values():
            total_taxable += adj.adjusted_month_gross_taxable_pay or Decimal('0.00')
            total_non_taxable += adj.adjusted_month_gross_non_taxable_pay or Decimal('0.00')
            total_gross += adj.adjusted_month_gross_pay or Decimal('0.00')
            total_taxable_combined += adj.adjusted_month_total_taxable_pay or Decimal('0.00')
            total_employment_income_tax += adj.adjusted_month_employment_income_tax_total or Decimal('0.00')
            total_employment_income_tax_adjustment += adj.adjusted_month_employment_income_tax or Decimal('0.00')
            #
            total_employee_pension += adj.adjusted_month_employee_pension_contribution or Decimal('0.00')
            total_employer_pension += adj.adjusted_month_employer_pension_contribution or Decimal('0.00')
            total_combined_pension += adj.adjusted_month_total_pension or Decimal('0.00')
            #
            #
            total_earning_adjustment_deduction += adj.adjusted_month_total_earning_deduction or Decimal('0.00')
            total_adjusted_expense += adj.adjusted_month_expense or Decimal('0.00')

        adjustments.update(
            recorded_month_taxable_gross_pay=total_taxable,
            recorded_month_non_taxable_gross_pay=total_non_taxable,
            recorded_month_gross_pay=total_gross,
            recorded_month_total_taxable_pay=total_taxable_combined,
            recorded_month_employment_income_tax_total=total_employment_income_tax,
            recorded_month_employment_income_tax=total_employment_income_tax_adjustment,
            #
            recorded_month_employee_pension_contribution=total_employee_pension,
            recorded_month_employer_pension_contribution=total_employer_pension,
            recorded_month_total_pension_contribution=total_combined_pension,
            #
            recorded_month_total_earning_deduction=total_earning_adjustment_deduction,
            recorded_month_expense=total_adjusted_expense,

        )


