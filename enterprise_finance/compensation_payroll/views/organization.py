from django.shortcuts import render, get_object_or_404, redirect
from django.contrib.auth.decorators import login_required
from compensation_payroll.models import OrganizationalProfile
from compensation_payroll.forms import OrganizationalProfileForm
from django.contrib import messages


#
@login_required
def organizational_profile(request):
    organizational_profiles = OrganizationalProfile.objects.filter(organization_name=request.user.organization_name)

    context = {'organizational_profiles': organizational_profiles, }

    return render(request, 'organizational_profile/list.html', context)


@login_required
def create_organizational_profile(request):
    # Redirect if the user already has an organization
    if request.user.organization_name:
        messages.info(request, "You already have an organization.")
        return redirect('dashboard')

    if request.method == 'POST':
        form = OrganizationalProfileForm(request.POST)
        if form.is_valid():
            org = form.save()  # Save directly, no commit=False
            request.user.organization_name = org
            request.user.save()
            messages.success(request, "Organization created successfully.")
            return redirect('dashboard')
        else:
            messages.error(request, "Please correct the errors below.")
    else:
        form = OrganizationalProfileForm()

    return render(request, 'organizational_profile/form.html', {'form': form})


def update_organizational_profile(request, pk):
    # Fetch the object to update or return 404 if not found
    organizational_profile = get_object_or_404(OrganizationalProfile, pk=pk,
                                               organization_name=request.user.organization_name)

    if request.method == 'POST':
        # Bind form with POST data and the instance to update
        form = OrganizationalProfileForm(request.POST, instance=organizational_profile)
        if form.is_valid():
            form.save()
            messages.success(request, "Organizational Profile updated successfully!")
            return redirect('organizational_profile')

            # return redirect('organizational_profile')  # Redirect to avoid form resubmission
        else:
            messages.error(request, "Please correct the errors below.")
    else:
        # Populate the form with the current instance for GET requests
        form = OrganizationalProfileForm(instance=organizational_profile)

    # Pass the list and the form to the template
    context = {
        'form': form,
        'edit_mode': True,  # Optional: to indicate edit mode in the template
        'editing_profile': organizational_profile,  # Optional: to pass the profile being edited
    }
    return render(request, 'organizational_profile/form.html', context)


# # Ensure the user is logged in before allowing access to this view
#
@login_required
def delete_organizational_profile(request, pk):
    """ Show a confirmation page before deleting """
    profile = get_object_or_404(OrganizationalProfile, pk=pk, organization_name=request.user.organization_name)

    if request.method == "POST":
        profile.delete()
        messages.success(request, "Organizational Profile deleted successfully!")
        return redirect("organizational_profile")  # Redirect after deletion

    return render(request, "organizational_profile/delete_confirm.html", {"profile": profile})


#