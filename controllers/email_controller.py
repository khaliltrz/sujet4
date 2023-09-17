from flask import Blueprint, request, jsonify
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from app.services.email_service import check_and_process_emails

email_controller = Blueprint('email_controller', __name__)

@email_controller.route('/process_emails', methods=['POST'])
def process_emails():
    check_and_process_emails()
    # Send an email with the updated Excel file (implement this part)
    return jsonify({'message': 'Email processing initiated'})

@email_controller.route('/send_email', methods=['POST'])
def send_email():
    # Handle sending email with the Excel file (implement this part)
    return jsonify({'message': 'Email sent successfully'})
