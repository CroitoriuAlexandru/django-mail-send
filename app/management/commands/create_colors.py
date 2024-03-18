from django.core.management import BaseCommand
from django.core.management.base import CommandParser
from app.models import Color


class Command(BaseCommand):
    colors = [
        Color(name="grayS"),
        Color(name="zincS"),
        Color(name="neutralS"),
        Color(name="stoneS"),
        Color(name="redS"),
        Color(name="orangeS"),
        Color(name="amberS"),
        Color(name="yellowS"),
        Color(name="limeS"),
        Color(name="greenS"),
        Color(name="emeraldS"),
        Color(name="tealS"),
        Color(name="cyanS"),
        Color(name="skyS"),
        Color(name="blueS"),
        Color(name="indigoS"),
        Color(name="violetS"),
        Color(name="purpleS"),
        Color(name="fuchsiaS"),
        Color(name="pinkS")
        ]
            
    def handle(self, *args, **options):
        for color in self.colors:
            color.save()
            
        print("Colors created!")
        