
# payroll process is on last
from django.shortcuts import render, get_object_or_404, redirect
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from compensation_payroll.models import AbsenceDeduction
from compensation_payroll.forms import AbsenceDeductionForm
from django.core.paginator import Paginator


# start absence deduction
@login_required
def absence_deduction_list(request):
    # Filter deductions by the logged-in user
    absence_deductions = AbsenceDeduction.objects.filter(organization_name=request.user.organization_name)
    paginator = Paginator(absence_deductions, 10)  # Show 10 deductions per page

    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    form = AbsenceDeductionForm()  # Instantiate an empty form for creating new deductions

    context = {'page_obj': page_obj, 'form': form}
    return render(request, 'absence_deduction/list.html', context)

#

@login_required
def create_absence_deduction(request):
    if request.method == 'POST':
        form = AbsenceDeductionForm(request.POST)
        if form.is_valid():
            absence_deduction = form.save(commit=False)
            absence_deduction.organization_name = request.user.organization_name
            absence_deduction.save()
            messages.success(request, "Absence Deduction created successfully!")
            return redirect('absence_deduction_list')
        else:
            messages.error(request, "Please correct the errors below.")
    else:
        form = AbsenceDeductionForm()

    return render(request, 'absence_deduction/list.html', {'form': form})


@login_required
def update_absence_deduction(request, pk):
    absence_deduction = get_object_or_404(AbsenceDeduction, pk=pk, organization_name=request.user.organization_name)

    if request.method == 'POST':
        form = AbsenceDeductionForm(request.POST, instance=absence_deduction)
        if form.is_valid():
            form.save()
            messages.success(request, "Absence Deduction updated successfully!")
            return redirect('absence_deduction_list')
        else:
            messages.error(request, "Please correct the errors below.")
    else:
        form = AbsenceDeductionForm(instance=absence_deduction)

    context = {
        'form': form,
        'edit_mode': True,
        'editing_absence_deduction': absence_deduction,
    }
    return render(request, 'absence_deduction/list.html', context)


@login_required
def delete_absence_deduction(request, pk):
    deduction = get_object_or_404(AbsenceDeduction, pk=pk, organization_name=request.user.organization_name)
    deduction.delete()
    messages.success(request, "Absence Deduction deleted successfully!")
    return redirect('absence_deduction_list')

