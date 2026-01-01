# Installation Guide (Linux)

## 1. System Requirements

```bash
sudo apt update
sudo apt install -y python3 python3-venv python3-pip postgresql postgresql-contrib
```

---

## 2. Clone Project

```bash
git clone https://github.com/HamdiReza/rooberah
cd rooberah
```

---

## 3. Create Python Virtual Environment

```bash
python3 -m venv .venv
source .venv/bin/activate
```

---

## 4. Install Python Dependencies

```bash
pip install -r requirements.txt
```

---

## 5. PostgreSQL Setup

```bash
sudo systemctl start postgresql
sudo systemctl enable postgresql
```

```bash
sudo -u postgres psql
```

```sql
ALTER USER postgres PASSWORD 'strongpassword';
CREATE DATABASE rooberah_db;
\q
```

---

## 6. Environment Variables

Create `.env` in project root:

```env
DATABASE_URL=postgresql://postgres:strongpassword@localhost:5432/rooberah_db
SECRET_KEY=dev
```

---

## 7. Database Migrations

```bash
flask db init
flask db migrate -m "init"
flask db upgrade
```

---

## 8. Run Application

```bash
python3 run.py
```

---

## 9. Access

* API: `http://127.0.0.1:5000`
* Dashboard: `http://127.0.0.1:5000/`

---

## 10. OTP Login Flow

1. POST `/api/auth/request-otp`
2. POST `/api/auth/verify-otp`
3. GET  `/api/users`
---

## Done

Project is ready for development or production deployment.
