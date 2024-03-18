from django.core.management import BaseCommand
from django.core.management.base import CommandParser
from app.models import Costume, User, CostumeState
from django.utils import datetime_safe


class Command(BaseCommand):
    costumeStates = CostumeState.objects.all()
    users = User.objects.all()
    
    costumes = [
        Costume(
            user=users[0],
            name="spidernam",
            origin="blue party",
            costumeState=costumeStates[0],
            purchase_price="300",
            parts="salopeta, masca",
            entry_date=datetime_safe.date(2020, 1, 1),
            retier_date=datetime_safe.date(2023, 1, 1),
            image="costumes/spiderman.png",
            ),
            Costume(
                user=users[1],
                name="hulk",
                origin="kiddo",
                costumeState=costumeStates[1],
                purchase_price="300",
                parts="salopeta, masca, bocanci",
                entry_date=datetime_safe.date(2021, 1, 1),
                retier_date=datetime_safe.date(2024, 1, 1),
                image="costumes/hulk.webp",
            ),
            Costume(
                name="micky mouse",
                origin="cufarul fermecat",
                costumeState=costumeStates[2],
                purchase_price="150",
                parts="salopeta, manusi, urechi",
                entry_date=datetime_safe.date(2021, 1, 1),
                retier_date=datetime_safe.date(2023, 1, 1),
                image="costumes/micky.webp",
            ),
            Costume(
                name="flash",
                origin="kiddo",
                costumeState=costumeStates[3],
                purchase_price="120",
                parts="salopeta, masca, pantofy",
                entry_date=datetime_safe.date(2019, 1, 1),
                retier_date=datetime_safe.date(2023, 1, 1),
                image="costumes/flash.png",
            ),
        ]
    
    def handle(self, *args, **options):
        for costume in self.costumes:
            costume.save()
        
        print("Costume's created!")