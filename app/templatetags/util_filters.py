from django import template

register = template.Library()

@register.filter(name='addclass')
def addclass(field, cssClasses):
   return field.as_widget(attrs={"class":cssClasses})

@register.filter(name='formatDateForWhatsApp')
def formatDateForWhatsApp(date):
    return date.strftime("%d/%m/%Y")

@register.filter(name='whatsAppServices')
def whatsAppServices(event):
   
   return "test"

