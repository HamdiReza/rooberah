import random
from datetime import datetime, timedelta

def generate_otp():
    return str(random.randint(10000, 99999))

def otp_expiry(minutes=2):
    return datetime.utcnow() + timedelta(minutes=minutes)
