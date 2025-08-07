from django.db import models
from tinymce.models import HTMLField
from django.utils.text import slugify

class BlogPost(models.Model):
    title = models.CharField(max_length=255)
    content = HTMLField()  # This will use TinyMCE for the content field
    tags = models.CharField(max_length=255, blank=True, help_text="Comma-separated tags")
    meta_title = models.CharField(max_length=255, blank=True)
    meta_description = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    slug = models.SlugField(unique=True, blank=True)

    class Meta:
        ordering = ['-created_at']
        verbose_name = 'Blog Post'
        verbose_name_plural = 'Blog Posts'

    def save(self, *args, **kwargs):
        if not self.slug:
            self.slug = slugify(self.title)
        super().save(*args, **kwargs)

    def tag_list(self):
        return [tag.strip() for tag in self.tags.split(',') if tag.strip()]

    def __str__(self):
        return self.title


class VideoPost(models.Model):
    title = models.CharField(max_length=255)
    content = HTMLField()
    video_url = models.URLField(help_text="Paste YouTube/Vimeo/other video URL")
    tags = models.CharField(max_length=255, blank=True)
    meta_title = models.CharField(max_length=255, blank=True)
    meta_description = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    slug = models.SlugField(unique=True, blank=True)

    class Meta:
        ordering = ['-created_at']
        verbose_name = 'Video Post'
        verbose_name_plural = 'Video Posts'

    def save(self, *args, **kwargs):
        if not self.slug:
            self.slug = slugify(self.title)
        super().save(*args, **kwargs)

    def tag_list(self):
        return [tag.strip() for tag in self.tags.split(',') if tag.strip()]

    def __str__(self):
        return self.title

class Documentation(models.Model):
    title = models.CharField(max_length=255)
    content = HTMLField()  # TinyMCE rich text editor
    tags = models.CharField(
        max_length=255,
        blank=True,
        help_text="Comma-separated tags (e.g., planning, strategy, performance)"
    )
    meta_title = models.CharField(max_length=255, blank=True)
    meta_description = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    slug = models.SlugField(unique=True, blank=True)

    def save(self, *args, **kwargs):
        if not self.slug:
            self.slug = slugify(self.title)  # Auto-generate slug
        super().save(*args, **kwargs)

    def __str__(self):
        return self.title