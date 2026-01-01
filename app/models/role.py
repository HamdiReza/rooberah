from app.extensions import db

roles_permissions = db.Table(
    "roles_permissions",
    db.Column("role_id", db.Integer, db.ForeignKey("roles.id")),
    db.Column("permission_id", db.Integer, db.ForeignKey("permissions.id")),
)


class Role(db.Model):
    __tablename__ = "roles"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), unique=True, nullable=False)
    description = db.Column(db.String(255))

    permissions = db.relationship(
        "Permission",
        secondary=roles_permissions,
        backref=db.backref("roles", lazy="dynamic"),
        lazy="dynamic",
    )
