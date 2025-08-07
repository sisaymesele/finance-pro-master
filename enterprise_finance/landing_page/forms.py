
from django import forms
from tinymce.widgets import TinyMCE
from .models import *


class BlogPostForm(forms.ModelForm):
    content = forms.CharField(widget=TinyMCE(attrs={'cols': 80, 'rows': 30}))  # Using TinyMCE widget for content field

    class Meta:
        model = BlogPost
        fields = ['title', 'content', 'tags', 'meta_title', 'meta_description', 'slug',]


class VideoPostForm(forms.ModelForm):
    class Meta:
        model = VideoPost
        fields = ['title', 'content', 'video_url', 'tags', 'meta_title', 'meta_description', 'slug',]
        widgets = {
            'content': forms.Textarea(attrs={'class': 'tinymce'}),
            'video_url': forms.URLInput(attrs={
                'placeholder': 'https://youtube.com/watch?v=... or https://vimeo.com/...',
                'class': 'form-control'
            }),
            'tags': forms.TextInput(attrs={
                'placeholder': 'comma,separated,tags',
                'class': 'form-control'
            }),
        }
        help_texts = {
            'video_url': 'Paste YouTube, Vimeo, or other embeddable video URL',
        }



class DocumentationForm(forms.ModelForm):
    content = forms.CharField(widget=TinyMCE(attrs={'cols': 80, 'rows': 30}))

    class Meta:
        model = Documentation
        fields = ['title', 'content', 'tags', 'meta_title', 'meta_description', 'slug',]

