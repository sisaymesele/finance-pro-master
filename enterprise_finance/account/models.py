from django.contrib.auth.models import AbstractUser, BaseUserManager
from django.db import models
from django.core.exceptions import ValidationError
import re
from django.db.models.signals import post_save
from django.dispatch import receiver
from django.core.mail import send_mail
from compensation_payroll.choices.country_code_choices import COUNTRY_CODE_CHOICES
from django.conf import settings
import datetime
from datetime import timedelta
from django.utils import timezone
from tinymce.models import HTMLField

class CustomUserManager(BaseUserManager):
    """Custom manager for CustomUser model."""

    def create_user(self, username, email, password=None, **extra_fields):
        """Create and return a regular user with an email, username, and password."""
        if not email:
            raise ValueError("The Email field must be set")
        if not username:
            raise ValueError("The Username field must be set")

        email = self.normalize_email(email)
        extra_fields.setdefault("is_active", True)

        user = self.model(username=username, email=email, **extra_fields)
        user.set_password(password)
        user.save(using=self._db)
        return user

    def create_superuser(self, username, email, password=None, **extra_fields):
        """Create and return a superuser with an email, username, and password."""
        if not email:
            raise ValueError("Superuser must have an email address")
        if not username:
            raise ValueError("Superuser must have a username")

        email = self.normalize_email(email)
        extra_fields.setdefault("is_staff", True)
        extra_fields.setdefault("is_superuser", True)

        return self.create_user(username, email, password, **extra_fields)


class CustomUser(AbstractUser):
    """Custom user model that requires both username and email for registration but uses email for login."""

    organization_name = models.ForeignKey('compensation_payroll.OrganizationalProfile', on_delete=models.SET_NULL, null=True, blank=True)
    email = models.EmailField("email address", unique=True)
    country_code = models.CharField(max_length=5, choices=COUNTRY_CODE_CHOICES, default='+251')
    phone_number = models.CharField(max_length=15)


    # Use email as the primary authentication field
    USERNAME_FIELD = "email"
    REQUIRED_FIELDS = ["username"]  # Username is required for registration

    objects = CustomUserManager()

    def __str__(self):
        return self.email

    @property
    def full_phone_number(self):
        """Returns the full phone number including the country code."""
        return f"{self.country_code}{self.phone_number}"

    def clean(self):
        """Validates the phone number format."""
        full_number = f"{self.country_code}{self.phone_number}"
        phone_number_regex = r'^\+?\d{12,13}$'  # Updated regex to allow valid formats

        if not re.match(phone_number_regex, full_number):
            raise ValidationError(
                f"The phone number {full_number} is invalid. Ensure it includes only digits and follows a valid format.")

        super().clean()

    def save(self, *args, **kwargs):
        if not self.country_code:
            self.country_code = 'ET'
        super().save(*args, **kwargs)


class EmailNotification(models.Model):
    subject = models.CharField(max_length=200)
    message = HTMLField()  # replaces TextField with TinyMCE editor
    created_at = models.DateField(default=datetime.date.today)  # Default current date

    def __str__(self):
        return self.subject

class Announcement(models.Model):
    title = models.CharField(max_length=255)
    message = models.TextField()
    created_at = models.DateField(default=datetime.date.today)

    def __str__(self):
        return self.title