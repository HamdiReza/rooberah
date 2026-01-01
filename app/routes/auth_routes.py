from flask import Blueprint, request, jsonify
from datetime import datetime

from app.extensions import db
from app.models.user import User
from app.utils.otp import generate_otp, otp_expiry

auth_bp = Blueprint("auth", __name__)

@auth_bp.route("/request-otp", methods=["POST"])
def request_otp():
    data = request.get_json()
    mobile = data.get("mobile")

    if not mobile:
        return jsonify({"error": "Mobile is required"}), 400

    user = User.query.filter_by(mobile=mobile).first()

    if not user:
        user = User(mobile=mobile)
        db.session.add(user)

    otp = generate_otp()
    user.otp_code = otp
    user.otp_expires_at = otp_expiry()
    user.is_verified = False

    db.session.commit()

    # 🔴 TEMP: print instead of SMS
    print(f"OTP for {mobile}: {otp}")

    return jsonify({"message": "OTP sent"}), 200

@auth_bp.route("/verify-otp", methods=["POST"])
def verify_otp():
    data = request.get_json()

    mobile = data.get("mobile")
    code = data.get("code")

    user = User.query.filter_by(mobile=mobile).first()

    if not user:
        return jsonify({"error": "User not found"}), 404

    if not user.otp_code or user.otp_code != code:
        return jsonify({"error": "Invalid code"}), 400

    if user.otp_expires_at < datetime.utcnow():
        return jsonify({"error": "OTP expired"}), 400

    user.is_verified = True
    user.otp_code = None
    user.otp_expires_at = None

    db.session.commit()

    return jsonify({
        "message": "Login successful",
        "user": user.to_dict()
    }), 200