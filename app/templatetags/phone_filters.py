from django import template

register = template.Library()

@register.filter
def phone_format(phone_number):
    phone_str = str(phone_number)
    return f"{phone_str[2:6]} {phone_str[6:9]} {phone_str[9:]}"