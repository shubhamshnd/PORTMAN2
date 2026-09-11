import base64
import smtplib
import threading
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from database import get_db, get_cursor

# Display name on every outgoing mail, unless smtp_config overrides it.
FROM_NAME = 'JSW JNPA Liquid Terminals - Portbird JNPA'


def build_approval_mail_html(
    approver_name,
    action_label,
    subtitle,
    details,
    action_url,
    action_btn_label,
    submitted_by,
    badge_color='#2544a7',
):
    """
    Build a styled HTML approval/notification email body.

    :param approver_name:    Recipient's display name (e.g. 'John')
    :param action_label:     Short status badge text (e.g. 'Pending Approval')
    :param subtitle:         One-line description shown below the header title
    :param details:          List of (label, value) tuples for the info rows
    :param action_url:       CTA button URL
    :param action_btn_label: CTA button text
    :param submitted_by:     Username who triggered the action
    :param badge_color:      Hex colour for the status badge background
    """
    rows_html = ''.join(
        f'''<tr>
              <td style="padding:6px 0;color:#718096;font-size:12px;width:38%;vertical-align:top;">{label}</td>
              <td style="padding:6px 0;color:#2d3748;font-size:12px;font-weight:600;vertical-align:top;">{value}</td>
            </tr>'''
        for label, value in details
    )
    return f"""
<div style="font-family:'Segoe UI',Arial,sans-serif;max-width:520px;margin:0 auto;padding:32px 20px;background:#f7fafc;border-radius:10px;">

  <!-- Header -->
  <div style="text-align:center;margin-bottom:24px;">
    <h2 style="color:#2d3748;font-size:20px;margin:0;letter-spacing:-0.3px;">Portbird <span style="color:#ec1c24;">&#x2022;</span> DPPL</h2>
    <p style="color:#718096;font-size:12px;margin:4px 0 0;">{subtitle}</p>
  </div>

  <!-- Card -->
  <div style="background:#fff;border-radius:10px;padding:28px 24px;border:1px solid #e2e8f0;box-shadow:0 1px 3px rgba(0,0,0,.06);">

    <!-- Status badge -->
    <div style="margin-bottom:20px;">
      <span style="display:inline-block;background:{badge_color};color:#fff;font-size:11px;font-weight:700;letter-spacing:.6px;text-transform:uppercase;padding:4px 12px;border-radius:20px;">{action_label}</span>
    </div>

    <p style="color:#2d3748;font-size:14px;margin:0 0 18px;">
      Hi <strong>{approver_name or 'Approver'}</strong>,
    </p>
    <p style="color:#4a5568;font-size:13px;margin:0 0 20px;line-height:1.6;">
      The following record has been submitted and requires your attention.
    </p>

    <!-- Details table -->
    <div style="background:#f7fafc;border-radius:8px;padding:16px 18px;margin:0 0 24px;">
      <table style="width:100%;border-collapse:collapse;">
        {rows_html}
        <tr>
          <td style="padding:6px 0;color:#718096;font-size:12px;width:38%;vertical-align:top;">Submitted by</td>
          <td style="padding:6px 0;color:#2d3748;font-size:12px;font-weight:600;vertical-align:top;">{submitted_by or '—'}</td>
        </tr>
      </table>
    </div>

    <!-- CTA button -->
    <div style="text-align:center;margin-bottom:6px;">
      <a href="{action_url}"
         style="display:inline-block;background:linear-gradient(135deg,#ec1c24,#2544a7);color:#fff;padding:12px 32px;border-radius:8px;text-decoration:none;font-size:14px;font-weight:600;letter-spacing:.4px;">
        {action_btn_label}
      </a>
    </div>

  </div>

  <!-- Footer -->
  <p style="text-align:center;color:#a0aec0;font-size:10px;margin:18px 0 0;">
    Portbird &mdash; DPPL &nbsp;|&nbsp; Port Management System
    <br>This is an automated notification. Please do not reply to this email.
  </p>

</div>
"""


def get_smtp_config():
    """Return smtp_config row as dict, or None if table empty/missing."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT * FROM smtp_config ORDER BY id LIMIT 1')
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None


def queue_mail(to_email, to_name, subject, body_html, module_code=None, ref_id=None,
               attachment_name=None, attachment_bytes=None):
    """Insert a pending mail into mail_queue. Safe to call from any view.

    A single optional attachment rides along base64-encoded in the row, so the
    queue stays self-contained — no temp files to lose between queueing and the
    retry that sends three attempts later."""
    if not to_email:
        return
    b64 = base64.b64encode(attachment_bytes).decode('ascii') if attachment_bytes else None
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("""
        INSERT INTO mail_queue (to_email, to_name, subject, body_html, module_code, ref_id,
                                attachment_name, attachment_b64)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
    """, [to_email, to_name, subject, body_html, module_code, ref_id,
          attachment_name if b64 else None, b64])
    conn.commit()
    conn.close()


def get_user_email_by_id(user_id):
    """Return (email, username) for a user id."""
    if not user_id:
        return None, None
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT email, username FROM users WHERE id=%s', [user_id])
    row = cur.fetchone()
    conn.close()
    return (row['email'], row['username']) if row else (None, None)


def get_module_approver_info(module_code, fallback_module=None):
    """Return approver config and contact details for a module."""
    from database import get_module_config

    cfg = get_module_config(module_code) or {}
    approver_id = cfg.get('approver_id')
    approval_add = cfg.get('approval_add')

    if fallback_module:
        fallback_cfg = get_module_config(fallback_module) or {}
        if not approver_id:
            approver_id = fallback_cfg.get('approver_id')
        if approval_add is None:
            approval_add = fallback_cfg.get('approval_add')

    email, username = get_user_email_by_id(approver_id)
    return {
        'approver_id': approver_id,
        'approval_add': bool(approval_add),
        'email': email,
        'username': username,
    }


def trigger_mail_processing():
    """Kick off an asynchronous mail send attempt."""
    threading.Thread(target=process_mail_queue, daemon=True).start()


def notify_module_approver(module_code, subject, body_html, ref_id=None, fallback_module=None):
    """Queue an approval mail to the configured approver and trigger processing."""
    info = get_module_approver_info(module_code, fallback_module=fallback_module)
    if not info.get('email'):
        return False
    queue_mail(
        info['email'],
        info.get('username'),
        subject,
        body_html,
        module_code,
        ref_id
    )
    trigger_mail_processing()
    return True


def process_mail_queue():
    """Called by scheduler. Reads smtp_config, sends all pending mails."""
    cfg = get_smtp_config()
    if not cfg or not cfg.get('is_enabled'):
        return  # kill-switch off

    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("""
        SELECT * FROM mail_queue
        WHERE status = 'pending' AND retry_count < max_retries
        ORDER BY created_at
    """)
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()

    for mail in rows:
        _send_one(mail, cfg)


def build_message(mail, cfg):
    """Compose one mail_queue row into a MIME message, with its attachment if
    it has one. Split out from _send_one so it can be checked without SMTP."""
    body = MIMEText(mail['body_html'], 'html')
    if mail.get('attachment_b64'):
        msg = MIMEMultipart('mixed')
        msg.attach(body)
        part = MIMEApplication(base64.b64decode(mail['attachment_b64']), _subtype='pdf')
        part.add_header('Content-Disposition', 'attachment',
                        filename=mail.get('attachment_name') or 'attachment.pdf')
        msg.attach(part)
    else:
        msg = MIMEMultipart('alternative')
        msg.attach(body)
    msg['Subject'] = mail['subject']
    msg['From'] = f"{cfg.get('from_name') or FROM_NAME} <{cfg['from_email']}>"
    msg['To'] = mail['to_email']
    return msg


def _send_one(mail, cfg):
    """Send a single mail row. Updates status in DB."""
    conn = get_db()
    cur = get_cursor(conn)
    try:
        msg = build_message(mail, cfg)

        server = smtplib.SMTP(cfg['host'], cfg['port'], timeout=15)
        if cfg.get('use_tls'):
            server.starttls()
        server.login(cfg['username'], cfg['password'])
        server.sendmail(cfg['from_email'], [mail['to_email']], msg.as_string())
        server.quit()

        cur.execute("""
            UPDATE mail_queue SET status='sent', sent_at=%s, error_message=NULL WHERE id=%s
        """, [datetime.now(), mail['id']])
    except Exception as e:
        new_retry = mail['retry_count'] + 1
        new_status = 'failed' if new_retry >= mail['max_retries'] else 'pending'
        cur.execute("""
            UPDATE mail_queue SET retry_count=%s, status=%s, error_message=%s WHERE id=%s
        """, [new_retry, new_status, str(e)[:500], mail['id']])
    finally:
        conn.commit()
        conn.close()
