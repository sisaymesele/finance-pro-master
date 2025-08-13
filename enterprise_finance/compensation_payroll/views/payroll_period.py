from django.shortcuts import render, get_object_or_404, redirect
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from compensation_payroll.models import PayrollPeriod
from compensation_payroll.forms import PayrollPeriodForm


@login_required
def payroll_period_list(request):
    payroll_periods = PayrollPeriod.objects.filter(
        organization_name=request.user.organization_name
    )

    context = {
        'payroll_periods': payroll_periods,
    }
    return render(request, 'payroll_period/list.html', context)


@login_required
def create_payroll_period(request):
    if request.method == 'POST':
        form = PayrollPeriodForm(request.POST)
        if form.is_valid():
            year = form.cleaned_data['year']
            month = form.cleaned_data['month']

            if PayrollPeriod.objects.filter(
                    organization_name=request.user.organization_name,
                    year=year,
                    month=month
            ).exists():
                messages.error(request, f"A record with {year}-{month} already exists for your account.")
            else:
                period = form.save(commit=False)
                period.organization_name = request.user.organization_name
                period.save()
                messages.success(request, "Payroll period created successfully!")
                return redirect('payroll_period_list')
    else:
        form = PayrollPeriodForm()

    context = {
        'form': form,
        'create_mode': True
    }
    return render(request, 'payroll_period/form.html', context)


@login_required
def update_payroll_period(request, pk):
    period = get_object_or_404(
        PayrollPeriod,
        pk=pk,
        organization_name=request.user.organization_name
    )

    if request.method == 'POST':
        form = PayrollPeriodForm(request.POST, instance=period)
        if form.is_valid():
            year = form.cleaned_data['year']
            month = form.cleaned_data['month']

            if PayrollPeriod.objects.filter(
                    organization_name=request.user.organization_name,
                    year=year,
                    month=month
            ).exclude(id=pk).exists():
                messages.error(request, f"A record with {year}-{month} already exists for your account.")
            else:
                form.save()
                messages.success(request, "Payroll period updated successfully!")
                return redirect('payroll_period_list')
    else:
        form = PayrollPeriodForm(instance=period)

    context = {
        'form': form,
        'edit_mode': True,
        'editing_payroll_period': period,
    }
    return render(request, 'payroll_period/form.html', context)


@login_required
def delete_payroll_period(request, pk):
    period = get_object_or_404(
        PayrollPeriod,
        pk=pk,
        organization_name=request.user.organization_name
    )

    if request.method == 'POST':
        period.delete()
        messages.success(request, "Payroll period deleted successfully!")
        return redirect('payroll_period_list')

    context = {
        'payroll_period': period,
    }
    return render(request, 'payroll_period/delete_confirm.html', context)
