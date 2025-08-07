from django.shortcuts import render, redirect
from django.contrib.auth.decorators import login_required
from compensation_payroll.models import OrganizationalProfile


@login_required
def dashboard(request):
    # Check if the user has a related organizational profile
    has_company = OrganizationalProfile.objects.filter(organization_name=request.user.organization_name).exists()

    if not has_company:
        return redirect('create_organizational_profile')

    return render(request, 'dashboard.html')