from django.db.models.signals import post_save, post_delete
from django.dispatch import receiver
from django.db import transaction
from django.db.models import Q

from compensation_payroll.models import (
    RegularPayroll, EarningAdjustment, DeductionAdjustment
)
from compensation_payroll.services.earning_adjustment.business import EarningAdjustmentBusinessService
from compensation_payroll.services.deduction_adjustment.business import DeductionAdjustmentBusinessService

# -------------------------------
# Earning Adjustment Recalculations
# -------------------------------

def _recalc_earning_adjustment(instance):
    service = EarningAdjustmentBusinessService(instance=instance)
    service.calculate_earning_adjustment_per_adjusted_month()
    service.calculate_monthly_earning_adjustment()

@receiver(post_save, sender=RegularPayroll)
def recalc_adjustments_on_regular_payroll_save(sender, instance, **kwargs):
    def _recalc_all():
        adjustments = EarningAdjustment.objects.filter(
            Q(payroll_to_record=instance) | Q(payroll_needing_adjustment=instance)
        )
        for adj in adjustments:
            _recalc_earning_adjustment(adj)

    transaction.on_commit(_recalc_all)

@receiver(post_save, sender=EarningAdjustment)
def recalc_regular_payroll_on_adjustment_save(sender, instance, **kwargs):
    if instance.payroll_to_record:
        transaction.on_commit(lambda: _recalc_earning_adjustment(instance))

@receiver(post_delete, sender=EarningAdjustment)
def recalc_regular_payroll_on_adjustment_delete(sender, instance, **kwargs):
    if instance.payroll_to_record:
        _recalc_earning_adjustment(instance)

# -------------------------------
# Deduction Adjustment Recalculations
# -------------------------------

def _recalc_deduction_adjustment(instance):
    service = DeductionAdjustmentBusinessService(instance=instance)
    service.calculate_deduction_adjustment_per_adjusted_month()
    service.calculate_monthly_deduction_adjustment()

@receiver(post_save, sender=DeductionAdjustment)
def recalc_deductions_on_save(sender, instance, **kwargs):
    transaction.on_commit(lambda: _recalc_deduction_adjustment(instance))

@receiver(post_delete, sender=DeductionAdjustment)
def recalc_deductions_on_delete(sender, instance, **kwargs):
    _recalc_deduction_adjustment(instance)

#
