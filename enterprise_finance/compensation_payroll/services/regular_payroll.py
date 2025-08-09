from decimal import Decimal
from .employment_income_tax import EmploymentIncomeTaxService  # adjust the import path as needed
from compensation_payroll.config.fetchers import OvertimeMultiplierFetcher, TransportHomeCapFetcher, FuelHomeCapFetcher, \
    TransportForWorkFetcher, FuelForWorkFetcher, HardshipAllowanceFetcher, PerDiemFetcher
from compensation_payroll.services.pension import calculate_pension_contributions


# compensation_payroll/services/regular_payroll.py
def get_basic_salary(i):
    if hasattr(i, 'personnel_full_name'):
        salary = getattr(i.personnel_full_name, 'basic_salary', None)
        if salary:
            return salary
    if hasattr(i, 'payroll_needing_adjustment'):
        adjusted = i.payroll_needing_adjustment
        if hasattr(adjusted, 'personnel_full_name'):
            return getattr(adjusted.personnel_full_name, 'basic_salary', None)
    return None


def get_working_area(i):
    """
    Returns working_area from payroll_needing_adjustment or personnel_full_name.
    Priority is given to payroll_needing_adjustment if available.
    """
    # Priority 1: payroll_needing_adjustment
    if hasattr(i, 'payroll_needing_adjustment') and i.payroll_needing_adjustment:
        adjusted = i.payroll_needing_adjustment
        area = getattr(adjusted, 'working_area', None)
        if area:
            return area

    # Priority 2: personnel_full_name
    if hasattr(i, 'personnel_full_name') and i.personnel_full_name:
        area = getattr(i.personnel_full_name, 'working_area', None)
        return area

    # Default fallback: None
    return None


def get_working_environment(i):
    """Retrieve working environment from either personnel_full_name or payroll_needing_adjustment."""
    if hasattr(i, 'personnel_full_name'):
        env = getattr(i.personnel_full_name, 'working_environment', None)
        if env:
            return env
    if hasattr(i, 'payroll_needing_adjustment'):
        adjusted = i.payroll_needing_adjustment
        if hasattr(adjusted, 'personnel_full_name'):
            return getattr(adjusted.personnel_full_name, 'working_environment', None)
    return None


def get_daily_per_diem(i):
    """Retrieve working environment from either personnel_full_name or payroll_needing_adjustment."""
    if hasattr(i, 'personnel_full_name'):
        env = getattr(i.personnel_full_name, 'daily_per_diem', None)
        if env:
            return env
    if hasattr(i, 'payroll_needing_adjustment'):
        adjusted = i.payroll_needing_adjustment
        if hasattr(adjusted, 'personnel_full_name'):
            return getattr(adjusted.personnel_full_name, 'daily_per_diem', None)
    return None


class RegularPayrollService:
    def __init__(self, instance):
        self.instance = instance
        self.multiplier_fetcher = OvertimeMultiplierFetcher()
        self.transport_home_cap_fetcher = TransportHomeCapFetcher()
        self.fuel_home_cap_fetcher = FuelHomeCapFetcher()
        self.transport_work_fetcher = TransportForWorkFetcher()
        self.fuel_work_fetcher = FuelForWorkFetcher()
        self.hardship_fetcher = HardshipAllowanceFetcher()
        self.per_diem_fetcher = PerDiemFetcher()


    def perform_calculations(self):
        self.calculate_overtime()
        self.calculate_transport_home_to_office()
        self.calculate_fuel_home_to_office()
        self.calculate_transport_for_work()
        self.calculate_fuel_for_work()
        self.calculate_hardship_allowance()
        self.calculate_per_diem()
        self.calculate_university_cost_share()
        self.calculate_gross_pay()
        self.calculate_gross_taxable_pay()
        self.calculate_gross_non_taxable_pay()
        self.calculate_final_save_calculation()


    # set up any other values from instance if needed

    def calculate_overtime(self):
        i = self.instance

        i.overtime_hours_from_six_pm_to_four_pm = Decimal(i.overtime_hours_from_six_pm_to_four_pm or '0.00')
        i.overtime_hours_from_four_pm_to_six_am = Decimal(i.overtime_hours_from_four_pm_to_six_am or '0.00')
        i.overtime_hours_in_weekly_rest_day = Decimal(i.overtime_hours_in_weekly_rest_day or '0.00')
        i.overtime_hours_in_public_holiday = Decimal(i.overtime_hours_in_public_holiday or '0.00')

        basic_salary = get_basic_salary(i)  # Assume this helper is defined elsewhere
        if not basic_salary:
            return Decimal('0.00')

        hourly_rate = basic_salary / Decimal('240')

        total_overtime = (
                i.overtime_hours_from_six_pm_to_four_pm * hourly_rate * self.multiplier_fetcher.get('evening') +
                i.overtime_hours_from_four_pm_to_six_am * hourly_rate * self.multiplier_fetcher.get('night') +
                i.overtime_hours_in_weekly_rest_day * hourly_rate * self.multiplier_fetcher.get('rest_day') +
                i.overtime_hours_in_public_holiday * hourly_rate * self.multiplier_fetcher.get('public_holiday')
        )

        return total_overtime

    def calculate_transport_home_to_office(self, amount=None):
        i = self.instance

        if amount is None:
            amount = Decimal(getattr(i, 'transport_home_to_office', '0.00') or '0.00')
        else:
            amount = Decimal(amount or '0.00')

        amount = max(amount, Decimal('0.00'))

        cap = self.transport_home_cap_fetcher.get()

        if amount <= cap:
            non_taxable = amount
            taxable = Decimal('0.00')
        else:
            non_taxable = cap
            taxable = amount - cap

        return taxable, non_taxable

    def calculate_fuel_home_to_office(self, amount=None):
        i = self.instance

        if amount is None:
            amount = Decimal(getattr(i, 'fuel_home_to_office', '0.00') or '0.00')
        else:
            amount = Decimal(amount or '0.00')

        amount = max(amount, Decimal('0.00'))

        cap = self.fuel_home_cap_fetcher.get()

        if amount <= cap:
            non_taxable = amount
            taxable = Decimal('0.00')
        else:
            non_taxable = cap
            taxable = amount - cap

        return taxable, non_taxable

    def calculate_transport_for_work(self, amount=None):
        i = self.instance

        taxable = Decimal('0.00')
        non_taxable = Decimal('0.00')

        # Get amount from parameter or instance
        amount = Decimal(amount or getattr(i, 'transport_for_work', '0.00') or '0.00')
        amount = max(amount, Decimal('0.00'))

        # Get salary
        basic_salary = get_basic_salary(i)
        if not basic_salary:
            return taxable, non_taxable

        # Fetch divisor and cap from config
        config = self.transport_work_fetcher.get_config()
        divisor = config['salary_divisor']
        cap = config['cap_amount']

        # Compute transport limit based on salary
        transport_limit = basic_salary / divisor

        # Match original logic exactly
        if amount <= transport_limit:
            non_taxable = amount
            taxable = Decimal('0.00')
        elif transport_limit < cap:
            non_taxable = transport_limit
            taxable = amount - transport_limit
        else:
            non_taxable = cap
            taxable = amount - cap

        return taxable, non_taxable

    #


    def calculate_fuel_for_work(self, amount=None):
        i = self.instance

        taxable = Decimal('0.00')
        non_taxable = Decimal('0.00')

        # Get amount from parameter or instance
        amount = Decimal(amount or getattr(i, 'fuel_for_work', '0.00') or '0.00')
        amount = max(amount, Decimal('0.00'))

        # Get salary
        basic_salary = get_basic_salary(i)
        if not basic_salary:
            return taxable, non_taxable

        # Fetch divisor and cap from config
        config = self.fuel_work_fetcher.get_config()
        divisor = config['salary_divisor']
        cap = config['cap_amount']

        # Compute fuel limit based on salary
        fuel_limit = basic_salary / divisor

        # Match original logic exactly
        if amount <= fuel_limit:
            non_taxable = amount
            taxable = Decimal('0.00')
        elif fuel_limit < cap:
            non_taxable = fuel_limit
            taxable = amount - fuel_limit
        else:
            non_taxable = cap
            taxable = amount - cap

        return taxable, non_taxable


    def calculate_hardship_allowance(self, amount=None):
        i = self.instance
        taxable = Decimal('0.00')
        non_taxable = Decimal('0.00')

        amount = Decimal(amount or getattr(i, 'hardship_allowance', '0.00') or '0.00')
        amount = max(amount, Decimal('0.00'))

        env = get_working_environment(i)
        salary = get_basic_salary(i)

        if not salary or not env:
            return taxable, non_taxable

        limit_percent = self.hardship_fetcher.get_limit_percent(env)
        if not limit_percent:
            return taxable, non_taxable

        limit = salary * limit_percent

        if amount <= limit:
            non_taxable = amount
        else:
            non_taxable = limit
            taxable = amount - limit

        return taxable, non_taxable



    def calculate_per_diem(self, amount=None):
        i = self.instance
        fetcher = self.per_diem_fetcher  # Use pre-initialized fetcher

        taxable = Decimal('0.00')
        non_taxable = Decimal('0.00')

        # Get claimed per diem amount
        amount = Decimal(amount or getattr(i, 'per_diem', '0.00') or '0.00')
        amount = max(amount, Decimal('0.00'))

        # Fetch required attributes
        basic_salary = get_basic_salary(i)
        daily_per_diem = get_daily_per_diem(i)
        area = get_working_area(i)

        # Early exit if required values missing
        if not basic_salary or not daily_per_diem:
            return taxable, non_taxable

        # Get config for area
        config = fetcher.get_config(area)

        # Fully non-taxable logic
        if config['fully_non_taxable']:
            return Decimal('0.00'), amount

        # Percentage-based logic
        limit = basic_salary * config['percent_limit']
        cap = config['cap_amount']

        if daily_per_diem <= limit:
            nt_daily = daily_per_diem
            tx_daily = Decimal('0.00')
        elif limit < cap:
            nt_daily = cap
            tx_daily = daily_per_diem - cap
        else:
            nt_daily = limit
            tx_daily = daily_per_diem - limit

        # Scale based on claimed amount
        taxable = (tx_daily / daily_per_diem) * amount if daily_per_diem else Decimal('0.00')
        non_taxable = (nt_daily / daily_per_diem) * amount if daily_per_diem else Decimal('0.00')

        return taxable, non_taxable


    def calculate_university_cost_share(self):
        i = self.instance
        percent = Decimal(i.cost_share_percent_to_basic_salary or '0.00')
        salary = i.personnel_full_name.basic_salary if i.personnel_full_name and i.personnel_full_name.basic_salary else Decimal(
            '0.00')
        return salary * (percent / Decimal('100'))


    def calculate_gross_pay(self):
        fields = [
            'basic_salary', 'overtime', 'transport_home_to_office', 'fuel_home_to_office',
            'transport_for_work', 'fuel_for_work', 'per_diem', 'hardship_allowance',
            'housing_allowance', 'position_allowance', 'commission', 'telephone_allowance',
            'one_time_bonus', 'causal_labor_wage', 'public_cash_award', 'incidental_operation_allowance',
            'medical_allowance', 'cash_gift', 'tuition_fees', 'personal_injury', 'child_support_payment'
        ]

        for field in fields:
            setattr(self.instance, field, Decimal(getattr(self.instance, field) or '0.00'))

        return sum(getattr(self.instance, field) for field in fields)


    def calculate_gross_taxable_pay(self):
        taxable_fields = [
            'basic_salary', 'overtime', 'transport_home_to_office_taxable',
            'fuel_home_to_office_taxable', 'transport_for_work_taxable', 'fuel_for_work_taxable',
            'per_diem_taxable', 'hardship_allowance_taxable', 'housing_allowance',
            'position_allowance', 'commission', 'telephone_allowance', 'one_time_bonus',
            'causal_labor_wage'
        ]

        for field in taxable_fields:
            setattr(self.instance, field, Decimal(getattr(self.instance, field) or '0.00'))

        return sum(getattr(self.instance, field) for field in taxable_fields)


    def calculate_gross_non_taxable_pay(self):
        non_taxable_fields = [
            'transport_home_to_office_non_taxable', 'fuel_home_to_office_non_taxable',
            'transport_for_work_non_taxable', 'fuel_for_work_non_taxable',
            'per_diem_non_taxable', 'hardship_allowance_non_taxable',
            'public_cash_award', 'incidental_operation_allowance',
            'medical_allowance', 'cash_gift', 'tuition_fees', 'personal_injury',
            'child_support_payment',
        ]

        for field in non_taxable_fields:
            setattr(self.instance, field, Decimal(getattr(self.instance, field) or '0.00'))

        return sum(getattr(self.instance, field) for field in non_taxable_fields)



    def calculate_final_save_calculation(self):
        i = self.instance

        i.overtime = self.calculate_overtime()
        i.transport_home_to_office_taxable, i.transport_home_to_office_non_taxable = self.calculate_transport_home_to_office()
        i.fuel_home_to_office_taxable, i.fuel_home_to_office_non_taxable = self.calculate_fuel_home_to_office()
        i.transport_for_work_taxable, i.transport_for_work_non_taxable = self.calculate_transport_for_work()
        i.fuel_for_work_taxable, i.fuel_for_work_non_taxable = self.calculate_fuel_for_work()
        i.per_diem_taxable, i.per_diem_non_taxable = self.calculate_per_diem()
        i.hardship_allowance_taxable, i.hardship_allowance_non_taxable = self.calculate_hardship_allowance()

        i.university_cost_share = self.calculate_university_cost_share()

        (
            i.employee_pension_contribution,
            i.employer_pension_contribution,
            i.total_pension_contribution
        ) = calculate_pension_contributions(i)


        i.gross_pay = self.calculate_gross_pay()
        i.gross_taxable_pay = self.calculate_gross_taxable_pay()
        i.gross_non_taxable_pay = self.calculate_gross_non_taxable_pay()

        i.employment_income_tax = EmploymentIncomeTaxService(i.gross_taxable_pay).calculate()

        # Sanitize deduction fields
        deduction_fields = [
            'employment_income_tax', 'employee_pension_contribution', 'charitable_donation',
            'saving_plan', 'loan_payment', 'court_order', 'workers_association',
            'personnel_insurance_saving', 'university_cost_share_pay', 'red_cross',
            'party_contribution', 'other_deduction'
        ]
        for field in deduction_fields:
            setattr(i, field, Decimal(getattr(i, field) or '0.00'))

        i.total_payroll_deduction = sum([getattr(i, field) for field in deduction_fields])
        i.net_pay = round(i.gross_pay - i.total_payroll_deduction, 2)
        i.expense = i.gross_pay + i.employer_pension_contribution





