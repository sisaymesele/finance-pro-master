# services/salary.py

from decimal import Decimal

class AbsenceDeductionService:
    def __init__(self, instance):
        self.instance = instance

    def calculate_remaining_salary(self):
        monthly_salary = Decimal(self.instance.monthly_salary or '0.00')
        absence_days = Decimal(self.instance.absence_days or 0)
        return monthly_salary - ((monthly_salary / Decimal(30)) * absence_days)

    def apply_calculations(self):
        """Apply calculated fields to instance."""
        self.instance.remaining_salary = self.calculate_remaining_salary()
        self.instance.absence_deduction_amount = self.instance.monthly_salary - self.instance.remaining_salary
        return self.instance
