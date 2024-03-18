from django.core.management import BaseCommand
from django.core.management.base import CommandParser
from app.models import Client


class Command(BaseCommand):
    clients = [
        Client(name="client_0", phone="0771133250"),
        Client(name="client_1", phone="0771133251"),
        Client(name="client_2", phone="0771133252"),
        Client(name="client_3", phone="0771133253"),
        Client(name="client_4", phone="0771133254"),
        Client(name="client_5", phone="0771133255"),
        Client(name="client_6", phone="0771133256"),
        Client(name="client_7", phone="0771133257"),
        Client(name="client_8", phone="0771133258"),
        Client(name="client_9", phone="0771133259"),
        
        ]
            
    def handle(self, *args, **options):
        for client in self.clients:
            client.save()
        
        print("Clients created!")