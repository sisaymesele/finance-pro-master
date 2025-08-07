
#
from decimal import Decimal
from compensation_payroll.config.fetchers import EmploymentIncomeTaxFetcher

class EmploymentIncomeTaxService:
   def __init__(self, taxable_pay):
        self.taxable_pay = Decimal(taxable_pay or 0)
        self.fetcher = EmploymentIncomeTaxFetcher()

   def calculate(self):
        brackets = self.fetcher.get_brackets()

        for bracket in brackets:
            min_amt = getattr(bracket, 'min_amount', bracket.get('min_amount'))
            max_amt = getattr(bracket, 'max_amount', bracket.get('max_amount')) or Decimal('999999999.99')
            rate = getattr(bracket, 'rate', bracket.get('rate'))
            deduction = getattr(bracket, 'deduction', bracket.get('deduction', Decimal('0.00')))

            if min_amt < self.taxable_pay <= max_amt:
                # Calculate based on total income, not just over min
                tax = self.taxable_pay * rate / Decimal('100') - deduction
                return max(tax, Decimal('0.00')).quantize(Decimal("0.01"))

        return Decimal('0.00')
