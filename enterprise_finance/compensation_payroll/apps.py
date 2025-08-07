from django.apps import AppConfig


class PayrollManagementConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'compensation_payroll'



    def ready(self):
        import compensation_payroll.signals  # adjust this path accordingly
