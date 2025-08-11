from django.shortcuts import render, get_object_or_404, redirect
from .models import BlogPost, Documentation, VideoPost, SiteSettings


# Create your views here.

def landing_page(request):
    return render(request, 'landing_page.html')

def feature_list(request):
    return render(request, 'partial/feature.html')

def how_it_work(request):
    return render(request, 'partial/how_it_work.html')

def pricing_list(request):
    return render(request, 'partial/pricing_list.html')

def video_list(request):
    return render(request, 'partial/video_list.html')

def qa_list(request):
    return render(request, 'partial/qa.html')

def why_choose(request):
    return render(request, 'partial/why_choose.html')

def organization_compatible(request):
    return render(request, 'partial/organization_compatible.html')

def compliance_standards(request):
    return render(request, 'partial/compliance_standard.html')

def blog_post_list(request):
    blog_posts = BlogPost.objects.all().order_by('-created_at')
    context = {
        'blog_posts': blog_posts
    }
    return render(request, 'blog-post/list.html', context)

def blog_post_detail(request, slug):
    blog_post = get_object_or_404(BlogPost, slug=slug)
    context = {
        'blog_post': blog_post,
        'meta_title': blog_post.meta_title or blog_post.title,
        'meta_description': blog_post.meta_description or blog_post.content[:150],
    }
    return render(request, 'blog-post/detail.html', context)

def video_post_list(request):
    """List all video posts"""
    video_posts = VideoPost.objects.all().order_by('-created_at')
    context = {
        'video_posts': video_posts
    }
    return render(request, 'video-post/list.html', context)

def documentation_list(request):
    documents = Documentation.objects.all()
    return render(request, 'documentation/list.html', {'documents': documents})

def documentation_detail(request, slug):
    document = get_object_or_404(Documentation, slug=slug)
    documents = Documentation.objects.all()  # for sidebar
    return render(request, 'documentation/detail.html', {
        'document': document,
        'documents': documents,
    })


def privacy_list(request):
    settings = SiteSettings.objects.all()
    return render(request, 'privacy_list.html', {'settings': settings})

def terms_list(request):
    settings = SiteSettings.objects.all()
    return render(request, 'terms_list.html', {'settings': settings})

def contact_list(request):
    settings = SiteSettings.objects.all()
    return render(request, 'contact_list.html', {'settings': settings})

