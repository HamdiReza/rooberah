from flask import Blueprint, request, jsonify
from app.extensions import db
from app.models.user import User

user_bp = Blueprint("users", __name__)

@user_bp.route("/", methods=["GET"])
def get_users():
    users = User.query.all()
    return jsonify([u.to_dict() for u in users])

@user_bp.route("/", methods=["POST"])
def create_user():
    data = request.get_json()

    if not data:
        return jsonify({"error": "Invalid JSON"}), 400

    user = User(
        name=data.get("name"),
        email=data.get("email")
    )

    db.session.add(user)
    db.session.commit()

    return jsonify(user.to_dict()), 201
