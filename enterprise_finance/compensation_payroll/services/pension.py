# payroll/services/pension.py
from decimal import Decimal
from compensation_payroll.services.constant_componet import PENSIONABLE
from compensation_payroll.config.fetchers import PensionRateFetcher


def calculate_pension_contributions(instance):
    """
    Supports both RegularPayroll and EarningAdjustment instances.
    Uses:
    - RegularPayroll.basic_salary
    - EarningAdjustment.earning_amount if component is pensionable
    """
    salary = None

    if hasattr(instance, 'basic_salary'):
        salary = instance.basic_salary

    elif hasattr(instance, 'earning_amount') and hasattr(instance, 'component'):
        if instance.component in PENSIONABLE:
            salary = instance.earning_amount

    if not salary:
        return Decimal('0.00'), Decimal('0.00'), Decimal('0.00')

    fetcher = PensionRateFetcher()
    config = fetcher.get_pension_rates()

    employee_pension = salary * config['personnel_rate']
    employer_pension = salary * config['employer_rate']
    total_pension = employee_pension + employer_pension

    return employee_pension, employer_pension, total_pension


# def calculate_pension_contributions(instance):
#     """
#     Supports both RegularPayroll and EarningAdjustment instances.
#     Uses:
#     - RegularPayroll.basic_salary
#     - EarningAdjustment.earning_amount if component is pensionable
#     """
#     from compensation_payroll.services.earning_adjustment import EarningAdjustmentService
#
#     salary = None
#
#     if hasattr(instance, 'basic_salary'):
#         # RegularPayroll model
#         salary = instance.basic_salary
#
#     elif hasattr(instance, 'earning_amount') and hasattr(instance, 'component'):
#         # EarningAdjustment model
#         if instance.component in EarningAdjustmentService.PENSIONABLE:
#             salary = instance.earning_amount
#
#     if not salary:
#         return Decimal('0.00'), Decimal('0.00'), Decimal('0.00')
#
#     fetcher = PensionRateFetcher()
#     config = fetcher.get_pension_rates()
#
#     employee_pension = salary * config['personnel_rate']
#     employer_pension = salary * config['employer_rate']
#     total_pension = employee_pension + employer_pension
#
#     return employee_pension, employer_pension, total_pension



# def calculate_pension_contributions(instance):
#     """
#     Calculates personnel, employer, and total pension contributions
#     for both RegularPayroll and EarningAdjustment instances,
#     using only the instance's own fields (e.g., basic_salary).
#     """
#     # Fetch basic salary directly from the instance
#     salary = getattr(instance, 'basic_salary', None)
#
#     if not salary:
#         return Decimal('0.00'), Decimal('0.00'), Decimal('0.00')
#
#     # Fetch default pension rates using PensionRateFetcher
#     fetcher = PensionRateFetcher()
#     config = fetcher.get_pension_rates()
#
#     # Calculate pension contributions
#     employee_pension = salary * config['personnel_rate']
#     employer_pension = salary * config['employer_rate']
#     total_pension = employee_pension + employer_pension
#
#     return employee_pension, employer_pension, total_pension

#
# def calculate_pension_contributions(instance):
#     """
#     Calculates personnel, employer, and total pension contributions
#     for both RegularPayroll and EarningAdjustment instances.
#     """
#     employee_pension = employer_pension = total_pension_contribution = Decimal('0.00')
#
#     # Try to get personnel and organization info from either direct fields or nested objects
#     personnel = getattr(instance, 'personnel_full_name', None)
#     organization = getattr(instance, 'organization_name', None)
#
#     # If not directly on the instance, try through payroll_to_record (used by EarningAdjustment)
#     if not personnel or not organization:
#         payroll_record = getattr(instance, 'payroll_to_record', None)
#         if payroll_record:
#             personnel = getattr(payroll_record, 'personnel_full_name', None)
#             organization = getattr(payroll_record, 'organization_name', organization)
#
#     # Proceed only if we have the required data
#     if personnel and organization and organization.organization_type and personnel.basic_salary:
#         org_type = organization.organization_type
#         salary = personnel.basic_salary
#
#         fetcher = PensionFetcher()
#         config = fetcher.get_config(org_type)
#
#         employee_pension = salary * config['personnel_rate']
#         employer_pension = salary * config['employer_rate']
#         total_pension_contribution = employee_pension + employer_pension
#
#     return employee_pension, employer_pension, total_pension_contribution
#
#
# #