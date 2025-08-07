from datetime import datetime, timedelta

from django.contrib import messages
from django.contrib.auth import login, logout, get_user_model
from django.contrib.auth.decorators import login_required
from django.contrib.auth.forms import PasswordResetForm, SetPasswordForm
from django.contrib.auth.tokens import default_token_generator as account_activation_token
from django.core.mail import EmailMessage, send_mail
from django.http import HttpResponse
from django.shortcuts import render, redirect, get_object_or_404
from django.template.loader import render_to_string
from django.urls import reverse
from django.utils import timezone
from django.utils.encoding import force_bytes, force_str
from django.utils.http import urlsafe_base64_encode, urlsafe_base64_decode
from django.contrib.sites.shortcuts import get_current_site

from .forms import *
from .models import *
from .tokens import account_activation_token


def register(request):
    if request.method == 'POST':
        form = CustomUserRegistrationForm(request.POST)
        if form.is_valid():
            user = form.save(commit=False)
            user.is_active = False
            user.save()

            current_site = get_current_site(request)
            mail_subject = 'Activate your account'
            message = render_to_string('activation-email.html', {
                'user': user,
                'domain': current_site.domain,
                'uid': urlsafe_base64_encode(force_bytes(user.pk)),
                'token': account_activation_token.make_token(user),
            })
            email = EmailMessage(mail_subject, message, to=[user.email])
            email.send()

            messages.success(request, 'Registration successful! Please check your email to verify your account.')
            messages.info(
                request,
                'Didn\'t receive the activation email? '
                '<a href="/account/resend-activation-email/">Resend it</a>'
            )
            return redirect('login')
    else:
        form = CustomUserRegistrationForm()

    return render(request, 'register.html', {'form': form})


def activate(request, uidb64, token):
    try:
        uid = force_str(urlsafe_base64_decode(uidb64))
        user = CustomUser.objects.get(pk=uid)
    except (TypeError, ValueError, OverflowError, CustomUser.DoesNotExist):
        user = None

    if user and account_activation_token.check_token(user, token):
        user.is_active = True
        user.save()
        login(request, user, backend='django.contrib.auth.backends.ModelBackend')
        messages.success(request, 'Your account has been activated successfully.')
        return redirect('dashboard')

    return render(request, 'activation-invalid.html')


def resend_activation_email(request):
    if request.method == 'POST':
        form = ResendActivationEmailForm(request.POST)
        if form.is_valid():
            email = form.cleaned_data['email']
            User = get_user_model()

            try:
                user = User.objects.get(email=email)
                if not user.is_active:
                    current_site = get_current_site(request)
                    mail_subject = 'Activate your account'
                    message = render_to_string('activation-email.html', {
                        'user': user,
                        'domain': current_site.domain,
                        'uid': urlsafe_base64_encode(force_bytes(user.pk)),
                        'token': account_activation_token.make_token(user),
                    })
                    email = EmailMessage(mail_subject, message, to=[user.email])
                    email.send()
                    messages.success(request, 'Activation email has been resent.')
                else:
                    messages.info(request, 'This account is already activated.')
            except User.DoesNotExist:
                messages.error(request, 'No user with that email exists.')

            return redirect('login')
    else:
        form = ResendActivationEmailForm()

    return render(request, 'activation-resend.html', {'form': form})


def resend_activate(request, uidb64, token):
    try:
        uid = force_str(urlsafe_base64_decode(uidb64))
        user = CustomUser.objects.get(pk=uid)
    except (TypeError, ValueError, OverflowError, CustomUser.DoesNotExist):
        user = None

    if user and account_activation_token.check_token(user, token):
        user.is_active = True
        user.save()
        login(request, user)
        messages.success(request, 'Your account has been activated successfully.')
        return redirect('dashboard')

    messages.error(request, 'Invalid or expired activation link. Please request a new one.')
    return redirect('resend_activation_email')


def login_view(request):
    if request.method == 'POST':
        form = CustomUserLoginForm(request, data=request.POST)
        if form.is_valid():
            user = form.get_user()
            login(request, user)
            return redirect('dashboard')
    else:
        form = CustomUserLoginForm()

    if 'activation_email_sent' in request.session:
        activation_time = datetime.strptime(request.session['activation_email_sent'], '%Y-%m-%d %H:%M:%S')
        time_diff = datetime.now() - activation_time

        if time_diff <= timedelta(minutes=10):
            messages.info(request, 'Didn\'t receive the activation email? '
                                   f'<a href="{reverse("resend_activation_email")}">Resend it</a>')

        del request.session['activation_email_sent']
        del request.session['activation_email_user']

    return render(request, 'login.html', {'form': form})


def logout_view(request):
    logout(request)
    return redirect('login')


def password_reset_request(request):
    if request.method == 'POST':
        form = PasswordResetForm(request.POST)
        if form.is_valid():
            email_address = form.cleaned_data['email']
            associated_users = CustomUser.objects.filter(email=email_address)
            if associated_users.exists():
                for user in associated_users:
                    current_site = get_current_site(request)
                    subject = 'Password Reset Requested'
                    context = {
                        'user': user,
                        'email': user.email,
                        'domain': current_site.domain,
                        'site_name': 'Compensation and Payroll Management',
                        'uid': urlsafe_base64_encode(force_bytes(user.pk)),
                        'token': account_activation_token.make_token(user),
                        'protocol': request.scheme,
                    }
                    email_body = render_to_string('password-reset-email.html', context)
                    try:
                        send_mail(subject, email_body, 'admin@example.com', [user.email], fail_silently=False)
                    except Exception:
                        messages.error(request, 'An error occurred while sending the password reset email.')
                        return redirect('password_reset')
                messages.success(request, 'A password reset email has been sent.')
                return redirect('login')
            else:
                messages.error(request, 'No account found with that email address.')
                return redirect('password_reset')
    else:
        form = PasswordResetForm()

    return render(request, 'password-reset.html', {'form': form})


def password_reset_confirm(request, uidb64, token):
    try:
        uid = force_str(urlsafe_base64_decode(uidb64))
        user = CustomUser.objects.get(pk=uid)
    except (TypeError, ValueError, OverflowError, CustomUser.DoesNotExist):
        user = None

    if user and account_activation_token.check_token(user, token):
        if request.method == 'POST':
            form = SetPasswordForm(user, request.POST)
            if form.is_valid():
                form.save()
                messages.success(request, 'Your password has been reset. You can now log in.')
                return redirect('login')
        else:
            form = SetPasswordForm(user)
        return render(request, 'password-reset-confirm.html', {'form': form})

    messages.error(request, 'The reset link is invalid or has expired.')
    return redirect('password_reset')


def password_reset_done(request):
    return render(request, 'password-reset-done.html')


def password_reset_complete(request):
    return render(request, 'password-reset-complete.html')


@login_required
def announcement_list(request):
    announcements = Announcement.objects.all().order_by('-created_at')
    return render(request, 'announcement/detail.html', {'announcements': announcements})

#



