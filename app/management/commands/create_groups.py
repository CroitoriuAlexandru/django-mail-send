from django.core.management import BaseCommand
from django.core.management.base import CommandParser

from django.contrib.auth.models import Group, Permission

from django.utils import datetime_safe


class Command(BaseCommand):
    
    def handle(self, *args, **options):
    
    
    
        groups = [
            Group(name="admin"),
            Group(name="dev"),
            Group(name="manager"),
            Group(name="provider"),
        ]
        for group in groups:
            
            group.save()
        
        print("Groups created!")
