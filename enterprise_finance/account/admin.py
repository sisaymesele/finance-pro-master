from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from django.utils.translation import gettext_lazy as _
from .forms import *
from .models import *
from axes.models import AccessAttempt, AccessLog
from django.conf import settings
from django.db import transaction
from django.template.loader import render_to_string
from django.core.mail import EmailMultiAlternatives, get_connection

# Custom User Admin

@admin.register(CustomUser)
class CustomUserAdmin(UserAdmin):
    model = CustomUser
    form = CustomUserAdminForm        # For editing users
    # add_form = CustomUserCreationForm  # For adding new users
    add_form = CustomUserRegistrationForm

    list_display = (
        'username', 'email', 'first_name', 'last_name',
        'country_code', 'phone_number', 'is_staff', 'is_active', 'last_login', 'date_joined'
    )
    list_filter = ('is_active', 'is_staff', 'is_superuser', 'groups', 'country_code')

    search_fields = ('email', 'username', 'first_name', 'last_name')
    ordering = ('email',)
    filter_horizontal = ('groups', 'user_permissions')

    fieldsets = (
        (None, {'fields': ('email', 'username', 'password')}),
        ('Personal Info', {'fields': ('organization_name', 'first_name', 'last_name', 'country_code', 'phone_number')}),
        ('Permissions', {'fields': ('is_active', 'is_staff', 'is_superuser', 'groups', 'user_permissions')}),
        ('Important Dates', {'fields': ('last_login', 'date_joined')}),
    )

    add_fieldsets = (
        (None, {
            'classes': ('wide',),
            'fields': (
                'organization_name', 'email', 'username', 'password1', 'password2',
                'first_name', 'last_name', 'country_code', 'phone_number',
                'is_active', 'is_staff', 'is_superuser', 'groups', 'user_permissions'
            ),
        }),
    )



class EmailNotificationAdmin(admin.ModelAdmin):
    list_display = ('subject', 'created_at')

    def save_model(self, request, obj, form, change):
        is_new = obj.pk is None
        super().save_model(request, obj, form, change)

        if is_new:
            transaction.on_commit(lambda: self.send_emails(obj))

    def send_emails(self, obj):
        users = CustomUser.objects.filter(is_active=True).exclude(email__isnull=True).exclude(email__exact='')
        print(f"✅ Found {users.count()} users")

        connection = get_connection(fail_silently=False)
        emails = []

        for user in users:
            print(f"Preparing email for {user.email}")
            context = {
                'user': user,
                'message': obj.message,
            }

            text_content = render_to_string('emails/notification.txt', context)
            html_content = render_to_string('emails/notification.html', context)

            email = EmailMultiAlternatives(
                subject=obj.subject,
                body=text_content,
                from_email=settings.DEFAULT_FROM_EMAIL,
                to=[user.email],
                connection=connection
            )
            email.attach_alternative(html_content, "text/html")
            emails.append(email)

        print(f"Sending {len(emails)} emails now...")
        connection.send_messages(emails)


admin.site.register(EmailNotification, EmailNotificationAdmin)
#

class AnnouncementAdmin(admin.ModelAdmin):
    list_display = ('title', 'message', 'created_at')
    search_fields = ('title', 'message')
    ordering = ('-created_at',)

admin.site.register(Announcement, AnnouncementAdmin)