"""
ASGI config for mysite project.

It exposes the ASGI callable as a module-level variable named ``application``.

For more information on this file, see
https://docs.djangoproject.com/en/4.0/howto/deployment/asgi/
"""

import os

from django.core.asgi import get_asgi_application
from dotenv import load_dotenv

load_dotenv()

if os.environ['DEBUG'] == 'True':
    os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'app.settings.local')
else:
    os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'app.settings.production')

application = get_asgi_application()
