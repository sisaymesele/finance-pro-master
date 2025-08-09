from django.urls import path
from . import views

urlpatterns = [
    path('', views.landing_page, name='home'),
    #
    path('feature/', views.feature_list, name='feature_list'),
    path('how-it-work', views.how_it_work, name='how_it_work_list'),
    path('pricing/', views.pricing_list, name='pricing_list'),
    path('video/', views.video_list, name='video_list'),
    path('qa/', views.qa_list, name='qa_list'),
    path('contact/', views.contact_list, name='contact_list'),
    path('privacy/', views.privacy_list, name='privacy_list'),
    path('terms/', views.terms_list, name='terms_list'),

    path('blog-post/', views.blog_post_list, name='blog_post_list'),  # Added trailing slash here
    path('blog-post/<slug:slug>', views.blog_post_detail, name='blog_post_detail'),  # Added trailing slash here
    # Video Post URLs
    path('video-post/', views.video_post_list, name='video_post_list'),
    #documentation
    path('documentation/', views.documentation_list, name='documentation_list'),
    path('documentation/<slug:slug>', views.documentation_detail, name='documentation_detail'),
]
