# config/fetchers.py
from decimal import Decimal
from compensation_payroll.config.models import OvertimeRate, TransportHomeToOfficeRate, FuelHomeToOfficeRate, \
    TransportForWorkRate, FuelForWorkRate, HardshipAllowanceRate, PerDiemRate, PensionRate, EmploymentIncomeTaxBracket

DEFAULT_OVERTIME_MULTIPLIERS = {
    'evening': Decimal('1.25'),
    'night': Decimal('1.50'),
    'rest_day': Decimal('2.00'),
    'public_holiday': Decimal('2.50'),
}


class OvertimeMultiplierFetcher:
    def __init__(self):
        self.defaults = DEFAULT_OVERTIME_MULTIPLIERS

    def get(self, component: str) -> Decimal:
        try:
            return OvertimeRate.objects.get(component=component).multiplier
        except OvertimeRate.DoesNotExist:
            return self.defaults.get(component, Decimal('1.0'))


class TransportHomeCapFetcher:
    def __init__(self):
        self.default = Decimal('600.00')

    def get(self) -> Decimal:
        try:
            return TransportHomeToOfficeRate.objects.first().cap_amount
        except AttributeError:
            return self.default


class FuelHomeCapFetcher:
    def __init__(self):
        self.default = Decimal('600.00')

    def get(self) -> Decimal:
        try:
            return FuelHomeToOfficeRate.objects.first().cap_amount
        except AttributeError:
            return self.default

#


class TransportForWorkFetcher:

    def __init__(self):
        self.defaults = {
            'salary_divisor': Decimal('4.00'),
            'cap_amount': Decimal('2200.00'),
        }

    def get_config(self) -> dict:
        try:
            config = TransportForWorkRate.objects.first()
            return {
                'salary_divisor': config.salary_divisor or self.defaults['salary_divisor'],
                'cap_amount': config.cap_amount or self.defaults['cap_amount'],
            }
        except Exception:
            return self.defaults




class FuelForWorkFetcher:
    def __init__(self):
        self.defaults = {
            'salary_divisor': Decimal('4.00'),
            'cap_amount': Decimal('2200.00'),
        }

    def get_config(self) -> dict:
        try:
            config = FuelForWorkRate.objects.first()
            return {
                'salary_divisor': config.salary_divisor or self.defaults['salary_divisor'],
                'cap_amount': config.cap_amount or self.defaults['cap_amount'],
            }
        except Exception:
            return self.defaults



class HardshipAllowanceFetcher:
    def __init__(self):
        self.default_limits = {
            'adverse': Decimal('0.25'),
            'very_adverse': Decimal('0.40'),
            'extremely_adverse': Decimal('0.60'),
        }

    def get_limit_percent(self, environment):
        try:
            config = HardshipAllowanceRate.objects.get(environment=environment)
            return config.limit_percent
        except HardshipAllowanceRate.DoesNotExist:
            return self.default_limits.get(environment, Decimal('0.00'))


class PerDiemFetcher:

    def __init__(self):

        self.defaults = {
            'non_governmental_manager': {'percent_limit': Decimal('0.05'), 'cap_amount': Decimal('1000.00'), 'fully_non_taxable': False},
            'deputy_non_governmental_manager': {'percent_limit': Decimal('0.05'), 'cap_amount': Decimal('1000.00'), 'fully_non_taxable': False},
            'construction_machinery_operator': {'percent_limit': Decimal('0.00'), 'cap_amount': Decimal('0.00'), 'fully_non_taxable': True},
            'government_official': {'percent_limit': Decimal('0.00'), 'cap_amount': Decimal('0.00'), 'fully_non_taxable': True},
            'non_governmental_expert': {'percent_limit': Decimal('0.04'), 'cap_amount': Decimal('500.00'), 'fully_non_taxable': False},
            'other': {'percent_limit': Decimal('0.04'), 'cap_amount': Decimal('500.00'), 'fully_non_taxable': False},
        }

    def get_config(self, working_area):
        try:
            config = PerDiemRate.objects.get(working_area=working_area)
            return {
                'percent_limit': config.percent_limit or self.defaults[working_area]['percent_limit'],
                'cap_amount': config.cap_amount or self.defaults[working_area]['cap_amount'],
                'fully_non_taxable': config.fully_non_taxable,
            }
        except PerDiemRate.DoesNotExist:
            return self.defaults.get(working_area, {'percent_limit': Decimal('0.00'), 'cap_amount': Decimal('0.00'), 'fully_non_taxable': False})



class PensionRateFetcher:
    def __init__(self):
        # Default pension rates for all organizations
        self.defaults = {
            'personnel_rate': Decimal('0.07'),
            'employer_rate': Decimal('0.11'),
        }

    def get_pension_rates(self):
        try:
            # Get latest PensionRate record (customize as per your app logic)
            pension_rate_obj = PensionRate.objects.latest('id')
            return {
                'personnel_rate': pension_rate_obj.personnel_rate,
                'employer_rate': pension_rate_obj.employer_rate,
            }
        except PensionRate.DoesNotExist:
            # No records found — return defaults
            return self.defaults



class EmploymentIncomeTaxFetcher:
    def __init__(self):
        self.defaults = [
            {'min_amount': Decimal('0.00'), 'max_amount': Decimal('2000.00'), 'rate': Decimal('0.00'), 'deduction': Decimal('0.00')},
            {'min_amount': Decimal('2001.00'), 'max_amount': Decimal('4000.00'), 'rate': Decimal('15.00'), 'deduction': Decimal('300.00')},
            {'min_amount': Decimal('4001.00'), 'max_amount': Decimal('7000.00'), 'rate': Decimal('20.00'), 'deduction': Decimal('500.00')},
            {'min_amount': Decimal('7001.00'), 'max_amount': Decimal('10000.00'), 'rate': Decimal('25.00'), 'deduction': Decimal('850.00')},
            {'min_amount': Decimal('10001.00'), 'max_amount': Decimal('14000.00'), 'rate': Decimal('30.00'), 'deduction': Decimal('1350.00')},
            {'min_amount': Decimal('14001.00'), 'max_amount': None, 'rate': Decimal('35.00'), 'deduction': Decimal('2050.00')},
        ]

    def get_brackets(self):
        try:
            brackets = EmploymentIncomeTaxBracket.objects.all().order_by('min_amount')
            if brackets.exists():
                return brackets
        except Exception:
            pass
        return self.defaults

