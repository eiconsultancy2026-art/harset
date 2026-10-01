from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import sqlite3, os
from datetime import datetime
from io import BytesIO
from functools import wraps
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook

BASE = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(BASE, "temple_accounts.db")

app = Flask(__name__)
app.secret_key = os.environ.get("TEMPLE_APP_SECRET", "CHANGE_THIS_SECRET_KEY_BEFORE_PRODUCTION")


# Standard temple accounting categories used by the Reason / Details dropdown.
# Keep these lists in one place so they can be changed without editing the HTML.
INCOME_CATEGORIES = [
    "Nanjai Varavu",
    "Punjai Varavu",
    "Donation",
    "Hundi Collection",
    "Kanikkai / Offering",
    "Archanai Income",
    "Abhishekam Income",
    "Prasadam Sales",
    "Pooja Ticket Income",
    "Festival Collection",
    "Hall / Premises Rent",
    "Lease / Property Income",
    "FD Interest",
    "Savings / Bank Interest",
    "Bank Interest / Other Interest",
    "Sale of Scrap",
    "Sale of Publications / Books",
    "Annadhanam Contribution",
    "Government / Institutional Grant",
    "Other Income",
]

EXPENSE_CATEGORIES = [
    "Pooja Materials",
    "Flowers / Garland",
    "Oil / Ghee / Camphor",
    "Prasadam Materials",
    "Annadhanam / Food Expenses",
    "Temple Staff Salary / Wages",
    "Priest / Archakar Honorarium",
    "Electricity Bill",
    "Water Bill",
    "Telephone / Internet",
    "Cleaning Expenses",
    "Security Expenses",
    "Repairs & Maintenance",
    "Building / Civil Work",
    "Electrical / Plumbing Work",
    "Festival Expenses",
    "Religious / Cultural Programme",
    "Vehicle / Transport Expenses",
    "Office / Stationery Expenses",
    "Bank Charges",
    "Audit / Professional Fees",
    "Insurance",
    "Taxes / Government Charges",
    "Donation / Charity Paid",
    "Miscellaneous Expense",
    "Other Expense",
]

@app.context_processor
def accounting_categories():
    return {"INCOME_CATEGORIES": INCOME_CATEGORIES, "EXPENSE_CATEGORIES": EXPENSE_CATEGORIES}

def db():
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys = ON")
    return c


def column_exists(c, table, column):
    return any(r[1] == column for r in c.execute(f"PRAGMA table_info({table})").fetchall())


def init_db():
    c = db()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS temples (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        address TEXT,
        created_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        temple_id INTEGER NOT NULL,
        username TEXT NOT NULL,
        password TEXT,
        password_hash TEXT,
        role TEXT NOT NULL DEFAULT 'admin',
        UNIQUE(temple_id, username),
        FOREIGN KEY(temple_id) REFERENCES temples(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS super_admins (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT NOT NULL UNIQUE,
        password_hash TEXT NOT NULL,
        created_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS transactions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        temple_id INTEGER NOT NULL,
        txn_date TEXT NOT NULL,
        purpose TEXT NOT NULL,
        donor TEXT,
        txn_type TEXT NOT NULL CHECK(txn_type IN ('donation','expense')),
        amount REAL NOT NULL,
        payment_mode TEXT NOT NULL,
        reference_no TEXT,
        created_at TEXT NOT NULL,
        created_by INTEGER,
        FOREIGN KEY(temple_id) REFERENCES temples(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS audit_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        temple_id INTEGER NOT NULL,
        user_id INTEGER,
        action TEXT NOT NULL,
        details TEXT,
        created_at TEXT NOT NULL,
        FOREIGN KEY(temple_id) REFERENCES temples(id) ON DELETE CASCADE
    );
    """)

    # Migrate older installations that defined users.password as NOT NULL.
    # The new schema keeps only password_hash and allows password to be null so
    # newly-created credentials never need to be stored as plain text.
    user_cols = c.execute("PRAGMA table_info(users)").fetchall()
    password_col = next((r for r in user_cols if r[1] == "password"), None)
    password_hash_exists = any(r[1] == "password_hash" for r in user_cols)
    if password_col and password_col[3] == 1:
        c.execute("PRAGMA foreign_keys=OFF")
        c.execute("ALTER TABLE users RENAME TO users_legacy")
        c.execute("""CREATE TABLE users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            temple_id INTEGER NOT NULL,
            username TEXT NOT NULL,
            password TEXT,
            password_hash TEXT,
            role TEXT NOT NULL DEFAULT 'admin',
            UNIQUE(temple_id, username),
            FOREIGN KEY(temple_id) REFERENCES temples(id) ON DELETE CASCADE
        )""")
        legacy_cols = [r[1] for r in user_cols]
        has_hash = "password_hash" in legacy_cols
        for row in c.execute("SELECT * FROM users_legacy").fetchall():
            values = dict(zip(legacy_cols, row))
            c.execute(
                "INSERT INTO users(id,temple_id,username,password,password_hash,role) VALUES(?,?,?,?,?,?)",
                (values["id"], values["temple_id"], values["username"], values.get("password"),
                 values.get("password_hash") if has_hash else None, values.get("role") or "admin")
            )
        c.execute("DROP TABLE users_legacy")
        c.execute("PRAGMA foreign_keys=ON")
    elif not password_hash_exists:
        c.execute("ALTER TABLE users ADD COLUMN password_hash TEXT")

    # Existing demo temple/user is preserved. Convert its password to a hash.
    if c.execute("SELECT COUNT(*) FROM temples").fetchone()[0] == 0:
        c.execute(
            "INSERT INTO temples(name,address,created_at) VALUES(?,?,?)",
            ("Sri Murugan Temple", "Temple Street, Chennai - 600001", datetime.now().isoformat()),
        )
        tid = c.execute("SELECT last_insert_rowid()").fetchone()[0]
        c.execute(
            "INSERT INTO users(temple_id,username,password,password_hash,role) VALUES(?,?,?,?,?)",
            (tid, "admin", None, generate_password_hash("admin123"), "admin"),
        )
    else:
        # Hash any legacy plain-text user passwords exactly once.
        legacy = c.execute("SELECT id,password,password_hash FROM users").fetchall()
        for u in legacy:
            if u[1] and not u[2]:
                c.execute(
                    "UPDATE users SET password=NULL,password_hash=? WHERE id=?",
                    (generate_password_hash(u[1]), u[0]),
                )

    # Create the single system Super Admin on first run.
    if c.execute("SELECT COUNT(*) FROM super_admins").fetchone()[0] == 0:
        c.execute(
            "INSERT INTO super_admins(username,password_hash,created_at) VALUES(?,?,?)",
            ("superadmin", generate_password_hash("SuperAdmin@123"), datetime.now().isoformat()),
        )

    c.commit()
    c.close()


def query(sql, args=()):
    c = db()
    rows = c.execute(sql, args).fetchall()
    c.close()
    return rows


def is_logged_in():
    return "user_id" in session


def is_super_admin():
    return session.get("role") == "super_admin"


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not is_logged_in():
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper


def super_admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not is_logged_in():
            return redirect(url_for("login"))
        if not is_super_admin():
            flash("Super Admin authorization is required for this action.")
            return redirect(url_for("dashboard"))
        return fn(*args, **kwargs)
    return wrapper


def temple_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not is_logged_in():
            return redirect(url_for("login"))
        if is_super_admin():
            flash("Select a temple account before performing this temple-user action.")
            return redirect(url_for("temples"))
        return fn(*args, **kwargs)
    return wrapper


def log(action, details="", temple_id=None, user_id=None):
    # Audit records are associated with the affected temple. This also lets
    # Super Admin changes remain visible to that temple's audit history.
    tid = temple_id if temple_id is not None else session.get("temple_id")
    uid = user_id if user_id is not None else session.get("user_id")
    if tid is None:
        return
    c = db()
    c.execute(
        "INSERT INTO audit_log(temple_id,user_id,action,details,created_at) VALUES(?,?,?,?,?)",
        (tid, uid, action, details, datetime.now().isoformat()),
    )
    c.commit()
    c.close()


@app.route("/", methods=["GET", "POST"])
def login():
    # The application entry point is always the login screen.  Do not
    # automatically redirect an existing browser session to the dashboard;
    # this ensures the normal startup workflow begins with authentication.
    # Successful authentication below still redirects to the role-appropriate
    # dashboard. Existing protected routes remain guarded by @login_required.
    if request.method == "POST":
        username = request.form["username"].strip()
        password = request.form["password"]

        # Super Admin is a separate system account and has no temple_id.
        admins = query("SELECT * FROM super_admins WHERE username=?", (username,))
        if admins and check_password_hash(admins[0]["password_hash"], password):
            a = admins[0]
            session.clear()
            session.update(
                user_id=a["id"], username=a["username"], role="super_admin",
                temple_id=None, temple_name="Super Admin"
            )
            return redirect(url_for("dashboard"))

        users = query(
            """SELECT u.*, t.name temple_name
               FROM users u JOIN temples t ON t.id=u.temple_id
               WHERE u.username=?""",
            (username,),
        )
        if users:
            u = users[0]
            valid = bool(u["password_hash"] and check_password_hash(u["password_hash"], password))
            if valid:
                session.clear()
                session.update(
                    user_id=u["id"], temple_id=u["temple_id"], username=u["username"],
                    temple_name=u["temple_name"], role="temple_admin"
                )
                log("LOGIN", "Successful temple login")
                return redirect(url_for("dashboard"))

        flash("பயனர் பெயர் அல்லது கடவுச்சொல் தவறாக உள்ளது")

    return render_template("login.html")


@app.route("/logout")
def logout():
    if is_logged_in() and not is_super_admin():
        log("LOGOUT", "User logged out")
    session.clear()
    return redirect(url_for("login"))


@app.route("/dashboard")
@login_required
def dashboard():
    if is_super_admin():
        # Super Admin can choose which temple's transactions to display.
        # No selection means all temples, preserving the existing overview.
        selected_temple_id = request.args.get("temple_id", type=int)
        temples = query("SELECT id, name FROM temples ORDER BY name")
        if selected_temple_id:
            tx = query(
                """SELECT tr.*, t.name temple_name
                   FROM transactions tr JOIN temples t ON t.id=tr.temple_id
                   WHERE tr.temple_id=?
                   ORDER BY tr.txn_date DESC, tr.id DESC""",
                (selected_temple_id,),
            )
        else:
            tx = query(
                """SELECT tr.*, t.name temple_name
                   FROM transactions tr JOIN temples t ON t.id=tr.temple_id
                   ORDER BY tr.txn_date DESC, tr.id DESC"""
            )
        income = sum(r["amount"] for r in tx if r["txn_type"] == "donation")
        expense = sum(r["amount"] for r in tx if r["txn_type"] == "expense")
        temples_count = len(temples)
        selected_temple_name = next(
            (t["name"] for t in temples if t["id"] == selected_temple_id), None
        )
        return render_template(
            "dashboard.html", tx=tx, income=income, expense=expense,
            balance=income-expense, super_admin=True, temples_count=temples_count,
            temples=temples, selected_temple_id=selected_temple_id,
            selected_temple_name=selected_temple_name
        )

    tid = session["temple_id"]
    tx = query("SELECT * FROM transactions WHERE temple_id=? ORDER BY txn_date DESC,id DESC", (tid,))
    income = sum(r["amount"] for r in tx if r["txn_type"] == "donation")
    expense = sum(r["amount"] for r in tx if r["txn_type"] == "expense")
    return render_template(
        "dashboard.html", tx=tx, income=income, expense=expense,
        balance=income-expense, super_admin=False, temples_count=0
    )


@app.route("/transaction", methods=["POST"])
@temple_required
def transaction():
    t = request.form["txn_type"]
    donor = request.form.get("donor", "").strip() if t == "donation" else ""
    amount = float(request.form["amount"])
    c = db()
    c.execute(
        """INSERT INTO transactions
        (temple_id,txn_date,purpose,donor,txn_type,amount,payment_mode,reference_no,created_at,created_by)
        VALUES(?,?,?,?,?,?,?,?,?,?)""",
        (
            session["temple_id"], request.form["txn_date"], request.form["purpose"].strip(),
            donor, t, amount, request.form["payment_mode"],
            request.form.get("reference_no", "").strip(), datetime.now().isoformat(), session["user_id"]
        ),
    )
    c.commit()
    c.close()
    log("CREATE_TRANSACTION", f"{t}: {amount} - {request.form['purpose']}")
    flash("பதிவு வெற்றிகரமாக சேமிக்கப்பட்டது")
    return redirect(url_for("dashboard"))


@app.route("/transaction/<int:txn_id>/edit", methods=["GET", "POST"])
@super_admin_required
def edit_transaction(txn_id):
    row = query("SELECT * FROM transactions WHERE id=?", (txn_id,))
    if not row:
        flash("Transaction not found.")
        return redirect(url_for("dashboard"))
    tx = row[0]

    if request.method == "POST":
        txn_type = request.form["txn_type"]
        donor = request.form.get("donor", "").strip() if txn_type == "donation" else ""
        amount = float(request.form["amount"])
        c = db()
        c.execute(
            """UPDATE transactions
               SET txn_date=?, purpose=?, donor=?, txn_type=?, amount=?, payment_mode=?, reference_no=?
               WHERE id=?""",
            (
                request.form["txn_date"], request.form["purpose"].strip(), donor, txn_type,
                amount, request.form["payment_mode"], request.form.get("reference_no", "").strip(), txn_id
            ),
        )
        c.commit(); c.close()
        log("MODIFY_TRANSACTION", f"Transaction #{txn_id} modified by Super Admin", temple_id=tx["temple_id"])
        flash("Transaction modified successfully.")
        return redirect(url_for("dashboard"))

    return render_template("edit_transaction.html", tx=tx)


@app.route("/transaction/<int:txn_id>/delete", methods=["POST"])
@super_admin_required
def delete_transaction(txn_id):
    row = query("SELECT * FROM transactions WHERE id=?", (txn_id,))
    if not row:
        flash("Transaction not found.")
        return redirect(url_for("dashboard"))
    tx = row[0]
    c = db()
    c.execute("DELETE FROM transactions WHERE id=?", (txn_id,))
    c.commit(); c.close()
    log("DELETE_TRANSACTION", f"Transaction #{txn_id} deleted by Super Admin", temple_id=tx["temple_id"])
    flash("Transaction deleted successfully.")
    return redirect(url_for("dashboard"))


@app.route("/reports")
@login_required
def reports():
    if is_super_admin():
        tid = request.args.get("temple_id", type=int)
        temples = query("SELECT * FROM temples ORDER BY name")
        if not tid and temples:
            tid = temples[0]["id"]
    else:
        tid = session["temple_id"]
        temples = []

    if not tid:
        flash("No temple is registered yet.")
        return redirect(url_for("temples"))

    tx = query("SELECT * FROM transactions WHERE temple_id=? ORDER BY txn_date,id", (tid,))
    temple_rows = query("SELECT * FROM temples WHERE id=?", (tid,))
    if not temple_rows:
        flash("Temple not found.")
        return redirect(url_for("dashboard"))
    temple = temple_rows[0]
    log("VIEW_REPORTS", "Reports opened", temple_id=tid)
    return render_template("reports.html", tx=tx, temple=temple, temples=temples, super_admin=is_super_admin())


@app.route("/export/excel")
@login_required
def export_excel():
    if is_super_admin():
        tid = request.args.get("temple_id", type=int)
    else:
        tid = session["temple_id"]
    if not tid:
        return redirect(url_for("dashboard"))

    tx = query(
        "SELECT txn_date,purpose,donor,txn_type,amount,payment_mode,reference_no FROM transactions WHERE temple_id=? ORDER BY txn_date,id",
        (tid,),
    )
    temple_rows = query("SELECT * FROM temples WHERE id=?", (tid,))
    if not temple_rows:
        return redirect(url_for("dashboard"))
    temple = temple_rows[0]

    wb = Workbook()
    ws = wb.active
    ws.title = "All Transactions"
    ws.append([temple["name"], ""])
    ws.append(["Date", "Purpose", "Donor Name", "Type", "Amount (INR)", "Payment Mode", "Reference No."])
    for r in tx:
        ws.append([
            r["txn_date"], r["purpose"], r["donor"],
            "Donation" if r["txn_type"] == "donation" else "Expense",
            r["amount"], r["payment_mode"], r["reference_no"]
        ])
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 22
    bio = BytesIO(); wb.save(bio); bio.seek(0)
    log("EXPORT_EXCEL", "All transactions exported", temple_id=tid)
    return send_file(
        bio, as_attachment=True, download_name=f"{temple['name']}_Temple_Accounts.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/audit")
@login_required
def audit():
    if is_super_admin():
        rows = query(
            """SELECT a.*, t.name temple_name
               FROM audit_log a JOIN temples t ON t.id=a.temple_id
               ORDER BY a.id DESC LIMIT 500"""
        )
    else:
        rows = query(
            "SELECT * FROM audit_log WHERE temple_id=? ORDER BY id DESC LIMIT 300",
            (session["temple_id"],),
        )
    return render_template("audit.html", rows=rows, super_admin=is_super_admin())


@app.route("/admin/temples", methods=["GET", "POST"])
@super_admin_required
def temples():
    if request.method == "POST":
        name = request.form["name"].strip()
        address = request.form.get("address", "").strip()
        username = request.form["username"].strip()
        password = request.form["password"]

        if len(password) < 8:
            flash("Temple password must contain at least 8 characters.")
            return redirect(url_for("temples"))

        c = db()
        try:
            c.execute(
                "INSERT INTO temples(name,address,created_at) VALUES(?,?,?)",
                (name, address, datetime.now().isoformat()),
            )
            tid = c.execute("SELECT last_insert_rowid()").fetchone()[0]
            c.execute(
                "INSERT INTO users(temple_id,username,password,password_hash,role) VALUES(?,?,?,?,?)",
                (tid, username, None, generate_password_hash(password), "admin"),
            )
            c.commit()
            log("CREATE_TEMPLE", f"Created temple '{name}' and login '{username}'", temple_id=tid)
            flash("Temple account created successfully.")
        except sqlite3.IntegrityError:
            c.rollback()
            flash("Temple name or username already exists.")
        finally:
            c.close()

    rows = query(
        """SELECT t.id,t.name,t.address,t.created_at,u.username
           FROM temples t LEFT JOIN users u ON u.temple_id=t.id
           ORDER BY t.name"""
    )
    return render_template("temples.html", rows=rows)


@app.route("/admin/temples/<int:temple_id>/edit", methods=["GET", "POST"])
@super_admin_required
def edit_temple(temple_id):
    rows = query("SELECT * FROM temples WHERE id=?", (temple_id,))
    if not rows:
        flash("Temple not found.")
        return redirect(url_for("temples"))
    temple = rows[0]
    user_rows = query("SELECT * FROM users WHERE temple_id=? LIMIT 1", (temple_id,))
    user = user_rows[0] if user_rows else None

    if request.method == "POST":
        name = request.form["name"].strip()
        address = request.form.get("address", "").strip()
        username = request.form["username"].strip()
        new_password = request.form.get("password", "")
        c = db()
        try:
            c.execute("UPDATE temples SET name=?, address=? WHERE id=?", (name, address, temple_id))
            if user:
                if new_password:
                    if len(new_password) < 8:
                        raise ValueError("Password must contain at least 8 characters.")
                    c.execute(
                        "UPDATE users SET username=?, password=NULL, password_hash=? WHERE id=?",
                        (username, generate_password_hash(new_password), user["id"]),
                    )
                else:
                    c.execute("UPDATE users SET username=? WHERE id=?", (username, user["id"]))
            c.commit()
            log("MODIFY_TEMPLE", f"Modified temple '{name}' and account", temple_id=temple_id)
            flash("Temple/account details updated successfully.")
            return redirect(url_for("temples"))
        except (sqlite3.IntegrityError, ValueError) as exc:
            c.rollback()
            flash(str(exc) if isinstance(exc, ValueError) else "Temple name or username already exists.")
        finally:
            c.close()

    return render_template("edit_temple.html", temple=temple, user=user)


if __name__ == "__main__":
    init_db()
    # LAN mode: other computers on the same Wi-Fi can connect to the host PC.
    app.run(host="0.0.0.0", port=8000, debug=False)
