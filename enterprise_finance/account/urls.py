from django.urls import path
from . import views

urlpatterns = [
    path("register/", views.register, name="register"),
    path("activate/<uidb64>/<token>/", views.activate, name="activate"),
    path('resend-activation-email/', views.resend_activation_email, name='resend_activation_email'),
    path("login/", views.login_view, name="login"),
    path("logout/", views.logout_view, name="logout"),
    path("password-reset/", views.password_reset_request, name="password_reset"),
    path("password-reset/done/", views.password_reset_done, name="password_reset_done"),
    path("password-reset-confirm/<uidb64>/<token>/", views.password_reset_confirm, name="password_reset_confirm"),
    path("password-reset/complete/", views.password_reset_complete, name="password_reset_complete"),

    #announcement
    path('announcement/', views.announcement_list, name='announcement_list'),

]
