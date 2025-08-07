# models.py
from django.db import models
from decimal import Decimal


class OvertimeRate(models.Model):
    COMPONENT_CHOICES = [
        ('evening', '6PM to 4PM'),
        ('night', '4PM to 6AM'),
        ('rest_day', 'Weekly Rest Day'),
        ('public_holiday', 'Public Holiday'),
    ]

    component = models.CharField(max_length=20, choices=COMPONENT_CHOICES, unique=True,
                                 help_text="Type of overtime component")
    multiplier = models.DecimalField(max_digits=5, decimal_places=2, default=1.00,
                                     help_text="Multiplier to apply to base hourly rate")

    def __str__(self):
        return f"{self.get_component_display()} ({self.multiplier}x)"

    class Meta:
        verbose_name = "Overtime Rate"
        verbose_name_plural = "Overtime Rate"


class TransportHomeToOfficeRate(models.Model):
    cap_amount = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('600.00'),
                                     help_text="Maximum non-taxable amount for transport home to office")

    def __str__(self):
        return f"Transport Home to Office Rate Cap: {self.cap_amount}"

    class Meta:
        verbose_name = "Transport Home to Office Rate"
        verbose_name_plural = "Transport Home to Office Rate"


class FuelHomeToOfficeRate(models.Model):
    cap_amount = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('600.00'),
                                     help_text="Maximum non-taxable amount for fuel home to office")

    def __str__(self):
        return f"Fuel Home to Office Rate Cap: {self.cap_amount}"

    class Meta:
        verbose_name = "Fuel Home to Office Rate"
        verbose_name_plural = "Fuel Home to Office Rate"


class TransportForWorkRate(models.Model):
    salary_divisor = models.DecimalField(max_digits=5, decimal_places=2, default=Decimal('4.00'),
                                         help_text="Divisor for salary to compute transport limit (e.g., 4 means salary/4)"
                                         )
    cap_amount = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('2200.00'),
                                     help_text="Maximum non-taxable cap for transport for work")

    def __str__(self):
        return f"Transport for work Rate: {self.salary_divisor}-{self.cap_amount}"

    class Meta:
        verbose_name = "Transport for work Rate"
        verbose_name_plural = "Transport for work Rate"


class FuelForWorkRate(models.Model):
    salary_divisor = models.DecimalField(max_digits=5, decimal_places=2, default=Decimal('4.00'),
                                         help_text="Divisor for salary to compute fuel limit (e.g., 4 means salary/4)"
                                         )
    cap_amount = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('2200.00'),
                                     help_text="Maximum non-taxable cap for fuel for work")



    def __str__(self):
        return f"Fuel for work Rate: {self.salary_divisor}-{self.cap_amount}"

    class Meta:
        verbose_name = "Fuel for work Rate"
        verbose_name_plural = "Fuel for work Rate"


class HardshipAllowanceRate(models.Model):
    WORK_ENVIRONMENT_CHOICES = [
        ('adverse', 'Adverse'),
        ('very_adverse', 'Very Adverse'),
        ('extremely_adverse', 'Extremely Adverse'),
    ]

    environment = models.CharField( max_length=30, choices=WORK_ENVIRONMENT_CHOICES, unique=True)
    limit_percent = models.DecimalField( max_digits=5, decimal_places=2, help_text="Percentage of basic salary allowed as non-taxable")

    def __str__(self):
        return f"{self.environment} ({self.limit_percent * 100:.0f}%)"

    class Meta:
        verbose_name = "Hardship Allowance Rate"
        verbose_name_plural = "Hardship Allowance Rate"


class PerDiemRate(models.Model):

    WORKING_AREA_CHOICES = [
        ('government_official', 'Government Official'), ('non_governmental_manager', 'Non Governmental Manager'),
        ('deputy_non_governmental_manager', 'Deputy Non Governmental Manager'),
        ('non_governmental_expert', 'Non Governmental Expert'),
        ('construction_machinery_operator', 'Construction Machinery Operator'),
        ('other', 'Other'),
    ]

    working_area = models.CharField(max_length=50, choices=WORKING_AREA_CHOICES, unique=True)
    percent_limit = models.DecimalField(max_digits=5, decimal_places=2, null=True, blank=True, help_text="Limit as percentage of salary")
    cap_amount = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True, help_text="Cap in birr")
    fully_non_taxable = models.BooleanField(default=False)

    def __str__(self):
        return self.get_working_area_display()



class PensionRate(models.Model):

    personnel_rate = models.DecimalField( max_digits=5, decimal_places=4, default=Decimal('0.00'), help_text="Percentage of salary contributed by personnel")
    employer_rate = models.DecimalField( max_digits=5, decimal_places=4, default=Decimal('0.00'), help_text="Percentage of salary contributed by employer")
  # You might want to add a timestamp or versioning
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-created_at']
    
    
class EmploymentIncomeTaxBracket(models.Model):
    min_amount = models.DecimalField(max_digits=10, decimal_places=2)
    max_amount = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True)
    rate = models.DecimalField(max_digits=5, decimal_places=2, help_text="Percentage rate (e.g., 10 for 10%)")
    deduction = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))

    class Meta:
        ordering = ['min_amount']

    def __str__(self):
        max_amt = self.max_amount if self.max_amount is not None else "∞"
        return f"{self.min_amount} - {max_amt} @ {self.rate}% - Deduct {self.deduction}"