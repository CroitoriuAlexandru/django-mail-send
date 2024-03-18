from django.core.management import BaseCommand
from django.core.management.base import CommandParser
from app import models


# find . -type d -name '__pycache__' -exec rm -r {} +
from django.core.management import BaseCommand
from django.core.management.base import CommandParser
from app import models





class Command(BaseCommand):
    help = "create superusers"
    
    def add_arguments(self, parser: CommandParser) -> None:
        parser.add_argument("name", type=str)
        parser.add_argument("password", type=str)
    
    def handle(self, *args, **options):
        user = models.User.objects.create_superuser(options["name"], " ", options["password"], image="users/"+options["name"]+".jpg")
        
        print(user, " User created!")