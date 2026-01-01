from datetime import datetime
from app.extensions import db

users_roles = db.Table(
    "users_roles",
    db.Column("user_id", db.Integer, db.ForeignKey("users.id")),
    db.Column("role_id", db.Integer, db.ForeignKey("roles.id")),
)


class User(db.Model):
    __tablename__ = "users"

    id = db.Column(db.Integer, primary_key=True)

    fname = db.Column(db.String(100))
    lname = db.Column(db.String(100))
    email = db.Column(db.String(120), unique=True, index=True)
    mobile = db.Column(db.String(15), unique=True, nullable=False, index=True)
    phone = db.Column(db.String(15))
    address = db.Column(db.Text)

    is_active = db.Column(db.Boolean, default=True)
    is_verified = db.Column(db.Boolean, default=False)

    otp_code = db.Column(db.String(5))
    otp_expires_at = db.Column(db.DateTime)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    roles = db.relationship(
        "Role",
        secondary=users_roles,
        backref=db.backref("users", lazy="dynamic"),
        lazy="dynamic",
    )

    def to_dict(self):
        return {
            "id": self.id,
            "fname": self.fname,
            "lname": self.lname,
            "email": self.email,
            "mobile": self.mobile,
            "phone": self.phone,
            "address": self.address,
            "roles": [r.name for r in self.roles],
            "is_verified": self.is_verified,
        }

    # helper
    def has_role(self, role_name):
        return any(r.name == role_name for r in self.roles)

    def has_permission(self, perm_name):
        for role in self.roles:
            if role.permissions.filter_by(name=perm_name).first():
                return True
        return False
