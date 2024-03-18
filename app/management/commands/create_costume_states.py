from django.core.management import BaseCommand
from django.core.management.base import CommandParser
from app import models


class Command(BaseCommand):
    colors = models.Color.objects.all()
    costumeStates = [
        models.CostumeState(name="foarte bun", info="", color=colors[8]),
        models.CostumeState(name="rupt", info="", color=colors[5]),
        models.CostumeState(name="retras", info="", color=colors[4]),
        models.CostumeState(name="nou nout", info="", color=colors[10]),
        ]
    
    def handle(self, *args, **options):
        for costumeState in self.costumeStates:
            costumeState.save()
        
        print("Costume state's created!")