from django import forms
from django.contrib.auth.forms import UserCreationForm, UserChangeForm, AuthenticationForm
from .models import *

class CustomUserAdminForm(UserChangeForm):
    class Meta:
        model = CustomUser
        fields = [
            'email', 'username', 'first_name', 'last_name',
            'country_code', 'phone_number',
            'is_active', 'is_staff', 'is_superuser',
            'groups', 'user_permissions'
        ]

class CustomUserRegistrationForm(UserCreationForm):
    """Form for user registration with widgets for styling."""

    class Meta:
        model = CustomUser
        fields = [
            # 'organization_name',
            'username', 'first_name', 'last_name', 'email',
            'country_code', 'phone_number',
            # 'is_active', 'is_staff', 'is_superuser',  # ← Add these
        ]
        widgets = {
            'username': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Username'}),
            'first_name': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'First Name'}),
            'last_name': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Last Name'}),
            'email': forms.EmailInput(attrs={'class': 'form-control', 'placeholder': 'Email Address'}),
            'country_code': forms.Select(attrs={'class': 'form-select'}),
            'phone_number': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Phone Number'}),
            'is_active': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'is_staff': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'is_superuser': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # password1 and password2 are NOT part of Meta widgets
        self.fields['password1'].widget.attrs.update({
            'class': 'form-control',
            'placeholder': 'Enter Password'
        })
        self.fields['password2'].widget.attrs.update({
            'class': 'form-control',
            'placeholder': 'Confirm Password'
        })

    def clean_email(self):
        """Ensure the email is unique."""
        email = self.cleaned_data.get("email")
        if CustomUser.objects.filter(email=email).exists():
            raise forms.ValidationError("A user with this email already exists.")
        return email


class CustomUserLoginForm(AuthenticationForm):
    """Login form using email instead of username."""

    username = forms.EmailField(
        label="Email",
        widget=forms.EmailInput(attrs={"autofocus": True, 'class': 'form-control', 'placeholder': 'Email Address'})
    )
    password = forms.CharField(
        label="Password",
        strip=False,
        widget=forms.PasswordInput(attrs={
            "autocomplete": "current-password",
            "class": "form-control",
            "placeholder": "Password"
        }),
    )

    def clean_username(self):
        """Ensure email exists in the system."""
        email = self.cleaned_data.get("username")
        if not email:
            raise forms.ValidationError("Email field is required.")
        return email


class ResendActivationEmailForm(forms.Form):
    """Form for resending the activation email."""

    email = forms.EmailField(
        label="Enter your email",
        widget=forms.EmailInput(attrs={
            'class': 'form-control form-control-lg',  # makes input larger
            'placeholder': 'Enter your registered email',
            'aria-describedby': 'emailHelp',
            'autocomplete': 'email',
        })
    )

    def clean_email(self):
        """Ensure that the email exists in the system."""
        email = self.cleaned_data.get("email")
        if not CustomUser.objects.filter(email=email).exists():
            raise ValidationError("No user found with this email address.")
        return email


class EmailNotificationForm(forms.ModelForm):
    """Form for creating email notifications."""

    class Meta:
        model = EmailNotification
        fields = ['subject', 'message', 'created_at', ]
        widgets = {
            'subject': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Subject'}),
            'message': forms.Textarea(attrs={'class': 'form-control', 'placeholder': 'Message', 'rows': 5}),
            'created_at': forms.DateInput(
                attrs={'class': 'form-control', 'placeholder': 'Select registration date', 'type': 'date'}),
        }

        def clean_subject(self):
            """Ensure subject is not empty."""
            subject = self.cleaned_data.get("subject")
            if not subject:
                raise forms.ValidationError("Subject is required.")
            return subject

        def clean_message(self):
            """Ensure message is not empty."""
            message = self.cleaned_data.get("message")
            if not message:
                raise forms.ValidationError("Message is required.")
            return message

class AnnouncementForm(forms.ModelForm):
    class Meta:
        model = Announcement
        fields = ['title', 'message', 'created_at']
        widgets = {
            'title': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Title'}),
            'message': forms.Textarea(attrs={'class': 'form-control', 'placeholder': 'Message'}),
            'created_at': forms.DateInput(
                attrs={
                    'class': 'form-control',
                    'type': 'date',  # Date picker
                }
            ),
        }