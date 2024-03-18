from django.core.management import BaseCommand
from django.core.management.base import CommandParser
from app import models


class Command(BaseCommand):
    colors = models.Color.objects.all()
    eventTypes = [
        models.EventType(name="animatie", info="Animatie pentru copii, descriere", color=colors[15], position=0),
        models.EventType(name="nunta", info="Nunta, descriere", color=colors[18], position=1),
        models.EventType(name="botez", info="Botez, descriere", color=colors[16], position=1),
        ]
            
    def handle(self, *args, **options):
        for eventType in self.eventTypes:
            eventType.save()
        
        print("Event Type's created!")