
from django.shortcuts import render, get_object_or_404, redirect
from django.contrib.auth.decorators import login_required
from compensation_payroll.models import PayrollMonthComponent
from compensation_payroll.forms import PayrollMonthComponentForm
from django.contrib import messages


@login_required
def payroll_month_and_component_list(request):
    # Query PayrollMonthComponent objects filtered by the current user's organization
    qs = PayrollMonthComponent.objects.filter(
        organization_name=request.user.organization_name
    )

    enriched_qs = []
    # Loop through each PayrollMonthComponent object
    for obj in qs:
        selected_fields = []
        # Loop through all fields of the object
        for field in obj._meta.fields:
            # Check if the field name starts with 'use_' and its value is True
            if field.name.startswith('use_') and getattr(obj, field.name, False):
                # Get the verbose name or fallback to the field name
                label = field.verbose_name or field.name
                # Remove 'use ' prefix if it exists at the start of the label
                if label.lower().startswith("use "):
                    label = label[4:]
                # Replace underscores with spaces and capitalize the label for display
                label = label.replace('_', ' ').capitalize()
                selected_fields.append(label)
        # Add a new attribute to the object with the list of selected component labels
        obj.selected_components = selected_fields
        # Add the enriched object to the list
        enriched_qs.append(obj)

    # Prepare the context dictionary to pass to the template
    context = {
        'payroll_month_and_components': enriched_qs,
    }

    # Render the template with the context
    return render(request, 'payroll_month_and_component/list.html', context)



#

@login_required
def create_payroll_month_and_component(request):
    if request.method == 'POST':
        form = PayrollMonthComponentForm(request.POST)
        if form.is_valid():
            payroll_month = form.cleaned_data['payroll_month']

            # Check for existing record with user filter
            if PayrollMonthComponent.objects.filter(
                    organization_name=request.user.organization_name,
                    payroll_month=payroll_month
            ).exists():
                messages.error(request, f"A record with {payroll_month} already exists for your account.")
            else:
                component = form.save(commit=False)
                component.organization_name = request.user.organization_name
                component.save()
                messages.success(request, "Payroll month created successfully!")
                return redirect('payroll_month_and_component_list')
    else:
        form = PayrollMonthComponentForm()

    context = {
        'form': form,
        'create_mode': True
    }
    return render(request, 'payroll_month_and_component/form.html', context)


@login_required
def update_payroll_month_and_component(request, pk):
    # Get component belonging to current user
    component = get_object_or_404(
        PayrollMonthComponent,
        pk=pk,
        organization_name=request.user.organization_name
    )

    if request.method == 'POST':
        form = PayrollMonthComponentForm(
            request.POST,
            instance=component
        )
        if form.is_valid():
            payroll_month = form.cleaned_data['payroll_month']

            # Check for conflicts with user filter
            if PayrollMonthComponent.objects.filter(
                    organization_name=request.user.organization_name,
                    payroll_month=payroll_month
            ).exclude(id=pk).exists():
                messages.error(request, f"A record with {payroll_month} already exists for your account.")
            else:
                form.save()
                messages.success(request, "Payroll month updated successfully!")
                return redirect('payroll_month_and_component_list')
    else:
        form = PayrollMonthComponentForm(
            instance=component
        )

    context = {
        'form': form,
        'edit_mode': True,
        'editing_payroll_month_and_component': component,
    }
    return render(request, 'payroll_month_and_component/form.html', context)


#

@login_required
def delete_payroll_month_and_component(request, pk):
    # Ensure user can only delete their own components
    payroll_month_and_component = get_object_or_404(
        PayrollMonthComponent,
        pk=pk,
        organization_name=request.user.organization_name
    )

    if request.method == 'POST':
        payroll_month_and_component.delete()
        messages.success(request, "Payroll month deleted successfully!")
        return redirect('payroll_month_and_component_list')
    context = {
        'payroll_month_and_component': payroll_month_and_component,
    }
    # If GET request, render confirmation page
    return render(request, 'payroll_month_and_component/delete_confirm.html', context)

