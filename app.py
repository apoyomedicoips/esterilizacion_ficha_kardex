import calendar
import json
import os
import secrets
import uuid
from collections import defaultdict
from datetime import date, datetime
from functools import wraps
from typing import Dict, List

import gspread
from dotenv import load_dotenv
from flask import Flask, abort, flash, redirect, render_template, request, session, url_for
from werkzeug.security import check_password_hash, generate_password_hash

load_dotenv()

app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', secrets.token_urlsafe(32))
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['SESSION_COOKIE_SECURE'] = os.getenv('COOKIE_SECURE', '0') == '1'

SPREADSHEET_ID = os.getenv('SPREADSHEET_ID', '161uPwejp2D8YVO7OVgbVyqeUiqzt0fYBzfZG5ERmLLg')


def load_users() -> Dict[str, Dict[str, str]]:
    raw = os.getenv('KARDEX_USERS_JSON', '').strip()
    if raw:
        parsed = json.loads(raw)
        users = {}
        for username, cfg in parsed.items():
            users[username.lower()] = {
                'password_hash': cfg.get('password_hash', ''),
                'role': cfg.get('role', 'viewer'),
                'active': bool(cfg.get('active', True)),
            }
        return users

    # Demo fallback. Change immediately in production.
    return {
        'admin': {
            'password_hash': generate_password_hash('Cambiar123!'),
            'role': 'admin',
            'active': True,
        }
    }


USERS = load_users()


class SheetStore:
    def __init__(self, spreadsheet_id: str):
        self.spreadsheet_id = spreadsheet_id
        self._gc = None
        self._spreadsheet = None

    def _client(self):
        if self._gc is not None:
            return self._gc

        credentials_file = os.getenv('GOOGLE_SERVICE_ACCOUNT_FILE', '').strip()
        credentials_json = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON', '').strip()

        if credentials_file:
            self._gc = gspread.service_account(filename=credentials_file)
        elif credentials_json:
            info = json.loads(credentials_json)
            self._gc = gspread.service_account_from_dict(info)
        else:
            raise RuntimeError(
                'Missing Google credentials. Set GOOGLE_SERVICE_ACCOUNT_FILE or GOOGLE_SERVICE_ACCOUNT_JSON.'
            )
        return self._gc

    def _doc(self):
        if self._spreadsheet is None:
            self._spreadsheet = self._client().open_by_key(self.spreadsheet_id)
        return self._spreadsheet

    def _ws(self, title: str, headers: List[str]):
        doc = self._doc()
        try:
            ws = doc.worksheet(title)
        except gspread.WorksheetNotFound:
            ws = doc.add_worksheet(title=title, rows=1000, cols=max(26, len(headers) + 2))
            ws.append_row(headers)
        if ws.row_count < 2:
            ws.add_rows(2)

        first = ws.row_values(1)
        if not first:
            ws.append_row(headers)
        elif first != headers:
            ws.update('A1', [headers])
        return ws

    def ensure_schema(self):
        self._ws('ITEMS', ['item_code', 'item_name', 'initial_stock', 'active', 'created_at'])
        self._ws('SERVICES', ['service_name', 'active', 'created_at'])
        self._ws(
            'MOVEMENTS',
            [
                'movement_id',
                'movement_date',
                'item_code',
                'service_name',
                'move_type',
                'quantity',
                'ticket_no',
                'notes',
                'username',
                'created_at',
            ],
        )
        self._ws('AUDIT_LOG', ['created_at', 'username', 'action', 'entity', 'entity_id', 'details'])

    def list_items(self):
        ws = self._ws('ITEMS', ['item_code', 'item_name', 'initial_stock', 'active', 'created_at'])
        rows = ws.get_all_records()
        out = []
        for r in rows:
            code = str(r.get('item_code', '')).strip()
            if not code:
                continue
            out.append(
                {
                    'item_code': code,
                    'item_name': str(r.get('item_name', '')).strip(),
                    'initial_stock': int(r.get('initial_stock') or 0),
                    'active': str(r.get('active', 'true')).lower() in ('true', '1', 'si', 'yes'),
                    'created_at': str(r.get('created_at', '')).strip(),
                }
            )
        return sorted(out, key=lambda x: x['item_code'])

    def add_item(self, code: str, name: str, initial_stock: int):
        code = code.strip().upper()
        if not code:
            raise ValueError('Invalid code')
        for item in self.list_items():
            if item['item_code'] == code:
                raise ValueError('Item already exists')
        ws = self._ws('ITEMS', ['item_code', 'item_name', 'initial_stock', 'active', 'created_at'])
        ws.append_row([code, name.strip(), int(initial_stock), 'true', datetime.utcnow().isoformat()])

    def list_services(self):
        ws = self._ws('SERVICES', ['service_name', 'active', 'created_at'])
        rows = ws.get_all_records()
        out = []
        for r in rows:
            name = str(r.get('service_name', '')).strip()
            if not name:
                continue
            out.append(
                {
                    'service_name': name,
                    'active': str(r.get('active', 'true')).lower() in ('true', '1', 'si', 'yes'),
                    'created_at': str(r.get('created_at', '')).strip(),
                }
            )
        return sorted(out, key=lambda x: x['service_name'])

    def add_service(self, service_name: str):
        name = service_name.strip().upper()
        if not name:
            raise ValueError('Invalid service')
        for svc in self.list_services():
            if svc['service_name'].upper() == name:
                raise ValueError('Service already exists')
        ws = self._ws('SERVICES', ['service_name', 'active', 'created_at'])
        ws.append_row([name, 'true', datetime.utcnow().isoformat()])

    def list_movements(self):
        ws = self._ws(
            'MOVEMENTS',
            [
                'movement_id',
                'movement_date',
                'item_code',
                'service_name',
                'move_type',
                'quantity',
                'ticket_no',
                'notes',
                'username',
                'created_at',
            ],
        )
        rows = ws.get_all_records()
        out = []
        for r in rows:
            movement_id = str(r.get('movement_id', '')).strip()
            if not movement_id:
                continue
            try:
                quantity = int(r.get('quantity') or 0)
            except Exception:
                quantity = 0
            out.append(
                {
                    'movement_id': movement_id,
                    'movement_date': str(r.get('movement_date', '')).strip(),
                    'item_code': str(r.get('item_code', '')).strip(),
                    'service_name': str(r.get('service_name', '')).strip(),
                    'move_type': str(r.get('move_type', 'OUT')).strip().upper(),
                    'quantity': quantity,
                    'ticket_no': str(r.get('ticket_no', '')).strip(),
                    'notes': str(r.get('notes', '')).strip(),
                    'username': str(r.get('username', '')).strip(),
                    'created_at': str(r.get('created_at', '')).strip(),
                }
            )
        return sorted(out, key=lambda x: x['created_at'], reverse=True)

    def add_movement(self, movement: Dict[str, str]):
        ws = self._ws(
            'MOVEMENTS',
            [
                'movement_id',
                'movement_date',
                'item_code',
                'service_name',
                'move_type',
                'quantity',
                'ticket_no',
                'notes',
                'username',
                'created_at',
            ],
        )
        ws.append_row(
            [
                str(uuid.uuid4()),
                movement['movement_date'],
                movement['item_code'],
                movement['service_name'],
                movement['move_type'],
                int(movement['quantity']),
                movement.get('ticket_no', ''),
                movement.get('notes', ''),
                movement['username'],
                datetime.utcnow().isoformat(),
            ]
        )

    def audit(self, username: str, action: str, entity: str, entity_id: str = '', details: str = ''):
        ws = self._ws('AUDIT_LOG', ['created_at', 'username', 'action', 'entity', 'entity_id', 'details'])
        ws.append_row([datetime.utcnow().isoformat(), username, action, entity, entity_id, details[:2000]])

    def list_audit(self, limit: int = 500):
        ws = self._ws('AUDIT_LOG', ['created_at', 'username', 'action', 'entity', 'entity_id', 'details'])
        rows = ws.get_all_records()
        out = []
        for r in rows:
            ts = str(r.get('created_at', '')).strip()
            if not ts:
                continue
            out.append(
                {
                    'created_at': ts,
                    'username': str(r.get('username', '')).strip(),
                    'action': str(r.get('action', '')).strip(),
                    'entity': str(r.get('entity', '')).strip(),
                    'entity_id': str(r.get('entity_id', '')).strip(),
                    'details': str(r.get('details', '')).strip(),
                }
            )
        return sorted(out, key=lambda x: x['created_at'], reverse=True)[:limit]


store = SheetStore(SPREADSHEET_ID)


@app.before_request
def ensure_csrf_token():
    if 'csrf_token' not in session:
        session['csrf_token'] = secrets.token_hex(16)


@app.context_processor
def inject_globals():
    return {
        'csrf_token': session.get('csrf_token', ''),
        'current_user': session.get('user', {}),
        'is_authenticated': bool(session.get('user')),
    }


def validate_csrf():
    sent = request.form.get('csrf_token', '')
    token = session.get('csrf_token', '')
    if not token or sent != token:
        abort(400, description='Invalid CSRF token')


def current_user():
    return session.get('user') or {}


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get('user'):
            return redirect(url_for('login'))
        return fn(*args, **kwargs)

    return wrapper


def require_roles(*roles):
    def deco(fn):
        @wraps(fn)
        def wrapper(*args, **kwargs):
            user = session.get('user')
            if not user:
                return redirect(url_for('login'))
            if user.get('role') not in roles:
                abort(403)
            return fn(*args, **kwargs)

        return wrapper

    return deco


def parse_month(value: str):
    try:
        year, month = value.split('-')
        year = int(year)
        month = int(month)
        if year < 2020 or year > 2100 or month < 1 or month > 12:
            raise ValueError
        return year, month
    except Exception:
        today = date.today()
        return today.year, today.month


@app.route('/health')
def health():
    return {'ok': True, 'time': datetime.utcnow().isoformat()}


@app.route('/login', methods=['GET', 'POST'])
def login():
    if session.get('user'):
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        validate_csrf()
        username = request.form.get('username', '').strip().lower()
        password = request.form.get('password', '')
        user = USERS.get(username)

        if not user or not user.get('active'):
            flash('Invalid username or password.', 'error')
            return render_template('login.html')

        if not check_password_hash(user['password_hash'], password):
            flash('Invalid username or password.', 'error')
            return render_template('login.html')

        session['user'] = {'username': username, 'role': user['role']}
        try:
            store.audit(username, 'LOGIN', 'User', username, 'Successful login')
        except Exception:
            pass
        return redirect(url_for('dashboard'))

    return render_template('login.html')


@app.route('/logout', methods=['POST'])
@login_required
def logout():
    validate_csrf()
    user = current_user()
    username = user.get('username', 'unknown')
    session.pop('user', None)
    try:
        store.audit(username, 'LOGOUT', 'User', username, 'Logout')
    except Exception:
        pass
    return redirect(url_for('login'))


@app.route('/')
@login_required
def index():
    return redirect(url_for('dashboard'))


@app.route('/dashboard')
@login_required
def dashboard():
    try:
        month_str = request.args.get('month') or date.today().strftime('%Y-%m')
        year, month = parse_month(month_str)
        first_day = date(year, month, 1)
        last_day = date(year, month, calendar.monthrange(year, month)[1])
        days = list(range(1, last_day.day + 1))

        items = [i for i in store.list_items() if i['active']]
        movements = store.list_movements()

        net_before = defaultdict(int)
        daily_delta = defaultdict(int)

        for m in movements:
            try:
                m_date = datetime.strptime(m['movement_date'], '%Y-%m-%d').date()
            except Exception:
                continue

            sign = 1 if m['move_type'] == 'IN' else -1
            item_code = m['item_code']

            if m_date < first_day:
                net_before[item_code] += sign * int(m['quantity'])
            elif first_day <= m_date <= last_day:
                daily_delta[(item_code, m_date.day)] += sign * int(m['quantity'])

        rows = []
        for item in items:
            code = item['item_code']
            start_balance = int(item['initial_stock']) + int(net_before.get(code, 0))
            running = start_balance
            day_values = []
            for d in days:
                running += daily_delta.get((code, d), 0)
                day_values.append(running)
            rows.append(
                {
                    'item_code': code,
                    'item_name': item['item_name'],
                    'start_balance': start_balance,
                    'day_values': day_values,
                    'end_balance': running,
                }
            )

        recent = movements[:40]
        return render_template(
            'dashboard.html', month_value=f'{year:04d}-{month:02d}', days=days, rows=rows, recent=recent
        )
    except Exception as exc:
        return render_template('error.html', error=str(exc), hint='Check Google credentials and sheet sharing.')


@app.route('/movements', methods=['GET', 'POST'])
@login_required
@require_roles('admin', 'operator')
def movements():
    try:
        items = [i for i in store.list_items() if i['active']]
        services = [s for s in store.list_services() if s['active']]

        if request.method == 'POST':
            validate_csrf()
            movement_date = request.form.get('movement_date', '').strip()
            item_code = request.form.get('item_code', '').strip().upper()
            service_name = request.form.get('service_name', '').strip().upper()
            move_type = request.form.get('move_type', 'OUT').strip().upper()
            quantity = int(request.form.get('quantity', '0'))
            ticket_no = request.form.get('ticket_no', '').strip()[:80]
            notes = request.form.get('notes', '').strip()[:1000]

            datetime.strptime(movement_date, '%Y-%m-%d')
            if move_type not in ('IN', 'OUT'):
                raise ValueError('Invalid movement type')
            if quantity <= 0:
                raise ValueError('Quantity must be > 0')

            store.add_movement(
                {
                    'movement_date': movement_date,
                    'item_code': item_code,
                    'service_name': service_name,
                    'move_type': move_type,
                    'quantity': quantity,
                    'ticket_no': ticket_no,
                    'notes': notes,
                    'username': current_user().get('username', 'unknown'),
                }
            )
            store.audit(
                current_user().get('username', 'unknown'),
                'CREATE',
                'Movement',
                item_code,
                f'{move_type} {quantity} on {movement_date} service={service_name}',
            )
            flash('Movement saved to Google Sheets.', 'ok')
            return redirect(url_for('movements'))

        listing = store.list_movements()[:250]
        return render_template('movements.html', items=items, services=services, listing=listing)
    except Exception as exc:
        return render_template('error.html', error=str(exc), hint='Cannot read/write Google Sheets now.')


@app.route('/admin/items', methods=['GET', 'POST'])
@login_required
@require_roles('admin')
def admin_items():
    try:
        if request.method == 'POST':
            validate_csrf()
            code = request.form.get('code', '').strip().upper()[:30]
            name = request.form.get('name', '').strip()[:255]
            initial_stock = int(request.form.get('initial_stock', '0') or 0)
            store.add_item(code, name, initial_stock)
            store.audit(current_user().get('username', 'unknown'), 'CREATE', 'Item', code, name)
            flash('Item created in Google Sheets.', 'ok')
            return redirect(url_for('admin_items'))

        items = store.list_items()
        return render_template('items.html', items=items)
    except Exception as exc:
        return render_template('error.html', error=str(exc), hint='Cannot access sheet ITEMS.')


@app.route('/admin/services', methods=['GET', 'POST'])
@login_required
@require_roles('admin')
def admin_services():
    try:
        if request.method == 'POST':
            validate_csrf()
            service_name = request.form.get('service_name', '').strip().upper()
            store.add_service(service_name)
            store.audit(current_user().get('username', 'unknown'), 'CREATE', 'Service', service_name, '')
            flash('Service created in Google Sheets.', 'ok')
            return redirect(url_for('admin_services'))

        services = store.list_services()
        return render_template('services.html', services=services)
    except Exception as exc:
        return render_template('error.html', error=str(exc), hint='Cannot access sheet SERVICES.')


@app.route('/admin/audit')
@login_required
@require_roles('admin')
def admin_audit():
    try:
        logs = store.list_audit(limit=500)
        return render_template('audit.html', logs=logs)
    except Exception as exc:
        return render_template('error.html', error=str(exc), hint='Cannot access sheet AUDIT_LOG.')


@app.route('/admin/init-sheet', methods=['POST'])
@login_required
@require_roles('admin')
def admin_init_sheet():
    validate_csrf()
    try:
        store.ensure_schema()
        store.audit(current_user().get('username', 'unknown'), 'INIT', 'Spreadsheet', SPREADSHEET_ID, 'Schema ensured')
        flash('Google Sheets schema initialized.', 'ok')
    except Exception as exc:
        flash(f'Error initializing schema: {exc}', 'error')
    return redirect(url_for('dashboard'))


if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5050, debug=False)
