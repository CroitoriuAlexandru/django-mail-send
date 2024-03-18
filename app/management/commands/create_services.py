from django.core.management import BaseCommand
from django.core.management.base import CommandParser
from app import models


class Command(BaseCommand):
    def handle(self, *args, **options):
        services = [
            models.Service(
                name="kids entertainer", 
                info="Animatori pentru copii ofero cea mai buna distractie 'long term' pentru copii, ...",
                price='200',
                duration='1',
                can_have_multiple='True',
                position='0'
            ),
            models.Service(
                name="Magician pentru copii", 
                info="Show de magie pentru copii si adulti, in special pentru copii :D",
                price='450',
                duration='0.5',
                can_have_multiple='False',
                position='1'
            ),
            models.Service(
                name="Piniata plina", 
                info="Un cadou special plin cu jucari si dulciuri pentru copii",
                price='400',
                duration='0',
                can_have_multiple='True',
                position='2'
            ),
            models.Service(
                name="Piniata goala", 
                info="alege tu ce sa pui in piniata si noi ti-o livram",
                price='250',
                duration='0',
                can_have_multiple='True',
                position='5'
            )
        ]
        
        for service in services:
            service.save()
            
        print("Services created!")