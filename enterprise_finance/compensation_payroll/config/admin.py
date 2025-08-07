# config_admin.py

from django.contrib import admin
from compensation_payroll.config.models import OvertimeRate, TransportHomeToOfficeRate, FuelHomeToOfficeRate, \
    TransportForWorkRate, FuelForWorkRate, HardshipAllowanceRate, PerDiemRate, PensionRate, EmploymentIncomeTaxBracket  # adjust import path as needed


@admin.register(OvertimeRate)
class OvertimeRateAdmin(admin.ModelAdmin):
    list_display = ('component', 'multiplier')
    ordering = ('component',)
    search_fields = ('component',)


@admin.register(TransportHomeToOfficeRate)
class TransportHomeToOfficeRateAdmin(admin.ModelAdmin):
    list_display = ('cap_amount',)


@admin.register(FuelHomeToOfficeRate)
class FuelHomeToOfficeRateAdmin(admin.ModelAdmin):
    list_display = ('cap_amount',)


# for work
@admin.register(TransportForWorkRate)
class TransportForWorkRateAdmin(admin.ModelAdmin):
    list_display = ('salary_divisor', 'cap_amount')


@admin.register(FuelForWorkRate)
class FuelForWorkRateAdmin(admin.ModelAdmin):
    list_display = ('salary_divisor', 'cap_amount')


@admin.register(HardshipAllowanceRate)
class HardshipAllowanceRateAdmin(admin.ModelAdmin):
    list_display = ('environment', 'limit_percent')

@admin.register(PerDiemRate)
class PerDiemRateAdmin(admin.ModelAdmin):
    list_display = ['working_area', 'percent_limit', 'cap_amount', 'fully_non_taxable']


@admin.register(PensionRate)
class PensionRateAdmin(admin.ModelAdmin):
    list_display = ('personnel_rate', 'employer_rate', 'created_at',)#'organization_type',


@admin.register(EmploymentIncomeTaxBracket)
class EmploymentIncomeTaxBracketAdmin(admin.ModelAdmin):
    list_display = ('min_amount', 'max_amount', 'rate', 'deduction')
    ordering = ('min_amount',)
    search_fields = ('min_amount', 'max_amount')