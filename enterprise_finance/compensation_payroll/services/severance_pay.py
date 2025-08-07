# services/severance_pay_service.py

from decimal import Decimal, ROUND_HALF_UP
from dateutil.relativedelta import relativedelta
from datetime import date

from .employment_income_tax import EmploymentIncomeTaxService


class SeverancePayService:
    """
    Service class to encapsulate business logic for SeverancePay model.
    Keeps the model clean and reusable across views, signals, commands, etc.
    """

    def __init__(self, severance_instance):
        self.obj = severance_instance


    def calculate_service_length(self):
        """
        Calculate service years and remaining days (excluding months).
        Caps years at 34 based on business rule.
        """
        if not self.obj.start_date:
            raise ValueError("Start date is required to calculate service length.")

        end_date = self.obj.end_date or date.today()
        delta = relativedelta(end_date, self.obj.start_date)

        # Cap at 34 years max
        if delta.years > 34:
            return 34, 0

        years = delta.years
        remaining_days = (end_date - (self.obj.start_date + relativedelta(years=years))).days
        return years, remaining_days

    def compute(self):
        """
        Main method that calculates and sets all severance-related fields on the model instance.
        """
        # Ensure Decimal values are initialized
        fields_to_zero = [
            'basic_salary', 'prorate_salary', 'employment_income_tax_from_basic_salary',
            'total_employment_income_tax_from_basic_salary', 'employment_income_tax_from_prorate_salary',
            'employment_income_tax_from_severance_pay', 'gross_severance_pay', 'net_severance_pay'
        ]
        for field in fields_to_zero:
            setattr(self.obj, field, Decimal(getattr(self.obj, field) or '0.00'))

        # Calculate service length
        years, days = self.calculate_service_length()
        self.obj.service_years = years
        self.obj.service_days = days

        # Monthly wage based on daily wage
        monthly_wage = self.obj.last_week_daily_wages * Decimal(30)
        base_severance = Decimal('1') / Decimal('3') * monthly_wage

        # Year-based severance (full year)
        self.obj.severance_for_years = (
            self.obj.last_week_daily_wages * Decimal(30)
            + (Decimal(years - 1) * base_severance)
        )

        # Day-based severance (partial year)
        self.obj.severance_for_days = (Decimal(days) / Decimal(365)) * base_severance

        # Total gross before type-based adjustment
        self.obj.gross_severance_pay = self.obj.severance_for_years + self.obj.severance_for_days

        # Adjust gross severance pay by severance type
        extra_days = Decimal(0)
        if self.obj.severance_type == 'no_notice':
            extra_days = Decimal(30)
        elif self.obj.severance_type == 'harassment':
            extra_days = Decimal(90)

        self.obj.gross_severance_pay += self.obj.last_week_daily_wages * extra_days

        # Tax base: full months and remaining prorated part
        self.obj.basic_salary = self.obj.last_week_daily_wages * Decimal(30)
        severance_by_months = self.obj.gross_severance_pay // self.obj.basic_salary
        self.obj.prorate_salary = self.obj.gross_severance_pay - (severance_by_months * self.obj.basic_salary)

        # Tax calculations
        self.obj.employment_income_tax_from_basic_salary = EmploymentIncomeTaxService(self.obj.basic_salary).calculate()
        self.obj.employment_income_tax_from_prorate_salary = EmploymentIncomeTaxService(self.obj.prorate_salary).calculate()

        self.obj.total_employment_income_tax_from_basic_salary = (
            self.obj.employment_income_tax_from_basic_salary * severance_by_months
        )

        self.obj.employment_income_tax_from_severance_pay = (
            self.obj.total_employment_income_tax_from_basic_salary + self.obj.employment_income_tax_from_prorate_salary
        )

        # Final net severance
        self.obj.net_severance_pay = (
            self.obj.gross_severance_pay - self.obj.employment_income_tax_from_severance_pay
        ).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

        return self.obj
