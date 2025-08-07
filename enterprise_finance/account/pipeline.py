from django.contrib.auth import get_user_model
from social_core.exceptions import AuthException
from social_core.exceptions import AuthAlreadyAssociated
from social_core.pipeline.social_auth import social_user as original_social_user
from django.contrib import messages
from django.shortcuts import redirect

class EmailAlreadyUsed(AuthException):
    def __str__(self):
        return 'This email is already associated with another account'


def social_user(backend, uid, user=None, *args, **kwargs):
    """Handle browser-specific association issues."""
    provider = backend.name
    social = backend.strategy.storage.user.get_social_auth(provider, uid)

    # If social account exists but no user is logged in
    if social and not user:
        return {'social': social, 'user': social.user, 'is_new': False}

    # If social account exists for a different user
    if social and user and social.user != user:
        # Store both users in session for potential merging
        backend.strategy.session_set('old_user_id', social.user.id)
        backend.strategy.session_set('new_user_id', user.id)

        # Add custom message
        request = backend.strategy.request
        messages.error(
            request,
            "🚫 This Google account is already linked to another account. contact support."
        )

        # Redirect to login or error page
        return redirect('login')  # Update 'login' to your actual login or error view name

    return original_social_user(backend, uid, user, *args, **kwargs)




def check_email_unique(backend, details, user=None, *args, **kwargs):
    """Prevent duplicate email registration"""
    if user:  # Existing user
        return None

    email = details.get('email')
    if email and get_user_model().objects.filter(email=email).exists():
        raise EmailAlreadyUsed(backend)


def create_user(backend, details, user=None, *args, **kwargs):
    """Create a user if one doesn't exist yet"""
    if user:
        return None

    User = get_user_model()
    email = details.get('email')
    username = details.get('username', email.split('@')[0])

    # If email already exists, stop user creation
    if User.objects.filter(email=email).exists():
        raise EmailAlreadyUsed(backend)

    new_user = User.objects.create_user(
        email=email,
        username=username,
        password=None  # Password will be handled by social auth backend
    )
    return {'user': new_user}







