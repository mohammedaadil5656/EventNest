import os
from datetime import timedelta

class Config:
    # Secret key for session management and CSRF security
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'eventnest-secret-key-2026-final-year-project'
    
    # 365-Day Long-term Persistent Session Configuration
    PERMANENT_SESSION_LIFETIME = timedelta(days=365)
    REMEMBER_COOKIE_DURATION = timedelta(days=365)
    REMEMBER_COOKIE_REFRESH_EACH_REQUEST = True
    SESSION_PERMANENT = True
    
    # Base directory and SQLite database path
    BASE_DIR = os.path.abspath(os.path.dirname(__file__))
    SQLALCHEMY_DATABASE_URI = 'sqlite:///' + os.path.join(BASE_DIR, 'database.db')
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    
    # Uploads folder for venue photos
    UPLOAD_FOLDER = os.path.join(BASE_DIR, 'static', 'images')

    # Razorpay Payment Gateway Configuration (Test Mode)
    # Replace with your Razorpay Dashboard keys or set via environment variables
    RAZORPAY_KEY_ID = os.environ.get('RAZORPAY_KEY_ID', 'rzp_test_EventNestDemoKey')
    RAZORPAY_KEY_SECRET = os.environ.get('RAZORPAY_KEY_SECRET', 'EventNestDemoSecret')
    RAZORPAY_CURRENCY = 'INR'

    # Admin Contact & Payment Info
    ADMIN_NAME = 'Mohammed Aadil'
    ADMIN_PHONE = '+91 90352 80291'
    ADMIN_EMAIL = 'mohammedaadil291@gamil.com'
    ADMIN_PHONEPE_QR = 'images/admin_phonepe_qr.jpg'
    REGISTRATION_FEE = 5000.0

    # Executive Leadership Board (CEO, Director, Manager, HR)
    EXECUTIVE_LEADERSHIP = [
        {
            'role': 'CEO',
            'title': 'Chief Executive Officer',
            'name': 'Mr. Ahmad Talha Siddiqui',
            'badge': 'bg-gold text-dark',
            'icon': 'fa-crown',
            'email': 'ahmadtalha@eventnest.com',
            'phone': '+91 98765 43210',
            'responsibilities': 'Executive Leadership, Strategic Expansion, Venue Acquisitions & Corporate Governance'
        },
        {
            'role': 'Director',
            'title': 'Managing Director & Co-Founder',
            'name': 'Mohammed Aadil',
            'badge': 'bg-charcoal text-gold',
            'icon': 'fa-user-tie',
            'email': 'mohammedaadil291@gamil.com',
            'phone': '+91 90352 80291',
            'responsibilities': 'Commercial Operations, PhonePe Payment Desk, Hall Owner Relations & Legal Compliance'
        },
        {
            'role': 'Manager',
            'title': 'Operations & Booking Manager',
            'name': 'Md Huzaif Masood',
            'badge': 'bg-primary text-white',
            'icon': 'fa-clipboard-check',
            'email': 'huzaifmasood68@gmail.com',
            'phone': '+91 97407 80164',
            'responsibilities': 'Venue Slot Approvals, Catering Audits, Schedule Coordination & Concierge Support'
        },
        {
            'role': 'HR',
            'title': 'Head of Human Resources',
            'name': 'Md Farhan Anwar',
            'badge': 'bg-success text-white',
            'icon': 'fa-users-gear',
            'email': 'farhananwar@eventnest.com',
            'phone': '+91 91234 56789',
            'responsibilities': 'Personnel Management, Hall Supervisor Onboarding, Quality Standards & Grievance Redressal'
        }
    ]


