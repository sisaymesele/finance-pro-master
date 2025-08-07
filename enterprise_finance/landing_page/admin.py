
from django.contrib import admin
from .models import BlogPost
from .forms import *
from django.utils.safestring import mark_safe  # Add this import statement

@admin.register(BlogPost)
class BlogPostAdmin(admin.ModelAdmin):
    form = BlogPostForm
    list_display = ('title', 'truncated_content', 'tags', 'meta_title', 'meta_description', 'created_at', 'updated_at')
    list_filter = ('title', 'created_at')
    search_fields = ('title', 'content', 'tags', 'meta_title', 'meta_description')
    prepopulated_fields = {'slug': ('title',)}
    readonly_fields = ('created_at', 'updated_at')
    ordering = ('-created_at',)

    def truncated_content(self, obj):
        """Return a truncated version of the content"""
        return mark_safe(obj.content[:40])  # Show first 40 characters of the content

    truncated_content.short_description = 'Content'

@admin.register(VideoPost)
class VideoPostAdmin(admin.ModelAdmin):
    form = VideoPostForm
    list_display = ('title', 'content', 'tags', 'meta_title', 'meta_description', 'created_at', 'updated_at')
    list_filter = ('tags', 'created_at')
    search_fields = ('title', 'content', 'tags', 'meta_title', 'meta_description')
    prepopulated_fields = {'slug': ('title',)}
    readonly_fields = ('created_at', 'updated_at')
    ordering = ('-created_at',)


@admin.register(Documentation)
class DocumentationAdmin(admin.ModelAdmin):
    form = DocumentationForm
    list_display = ('title', 'tag_list', 'short_content', 'created_at', 'updated_at')
    list_filter = ('created_at', 'tags')
    search_fields = ('title', 'content', 'tags', 'meta_title', 'meta_description')
    prepopulated_fields = {'slug': ('title',)}
    readonly_fields = ('created_at', 'updated_at')
    ordering = ('-created_at',)

    def tag_list(self, obj):
        return obj.tags

    def short_content(self, obj):
        return mark_safe(obj.content[:60] + "...")

    short_content.short_description = 'Content Preview'


@admin.register(SiteSettings)
class SiteSettingsAdmin(admin.ModelAdmin):
    list_display = (
        'site_name', 'privacy', 'copy_right', 'terms',
        'address', 'phone', 'email', 'website',
        'google_map_embed_url', 'facebook', 'twitter', 'linkedin',
        'instagram', 'youtube'
    )

