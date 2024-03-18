from ipaddress import summarize_address_range
from django.http import JsonResponse, HttpResponse, HttpResponseRedirect
from django.shortcuts import render, redirect, reverse, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.db.models import Value
from django.db.models.functions import Concat
from datetime import datetime, timedelta, date
from django.utils import timezone
import uuid
import win32com.client
import openai

from icecream import ic
from .models import *
from .decorators import unauthentificated_user, allowed_users



api_key = "test"  # Replace with your actual OpenAI API key


def generate_email_body(subject, api_key):
    """
    Generate an email body based on the given subject using OpenAI's GPT-3.5 Turbo model.
    """
    
    openai.api_key = api_key
    conversation = [
        {"role": "user", "content": f"Write a professional email about {subject}."},
        {"role": "assistant", "content": "Sure, here's a draft of an email:"}
    ]
    response = openai.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=conversation
    )
    return response.choices[0].message.content



def send_email_with_outlook(recipient, subject, message, attachment_path):
    """Send an email using Outlook with an attachment."""
    ol = win32com.client.Dispatch("Outlook.Application")
    newmail = ol.CreateItem(0x0)
    newmail.Subject = subject
    newmail.To = recipient
    newmail.Body = message
    # Add attachment
    if attachment_path:
        newmail.Attachments.Add(attachment_path)
    newmail.Send()
    

# @allowed_users(allowed_roles=[])
def getHome(request):
    if request.method == 'GET':
        context = {}
        return render(request, './templates/get.html', context)
    elif request.method == 'POST':

        # Usage Example
        subject = "Job Offer in Developing a Flask App"
        # Generate the email body
        email_body = generate_email_body(subject, api_key)
        # Replace with actual recipient, subject, and attachment path
        recipient = "nykw2002@gmail.com"
        attachment_path = "36.pdf"
        # Send the email
        send_email_with_outlook(recipient, subject, email_body, attachment_path)
        print(request.POST.email)
        return JsonResponse({'message': 'POST request received.'})
    


