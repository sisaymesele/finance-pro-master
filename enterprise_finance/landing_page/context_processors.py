# landing_page/context_processors.py

from .models import SiteSettings

def footer_settings(request):
    settings = SiteSettings.objects.all()
    return {'footer_settings': settings}
