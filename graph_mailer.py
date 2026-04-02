"""
graph_mailer.py — Microsoft Graph Email Sender (Windows Edition)
=================================================================
Sends the Azure VM Performance Report via Microsoft Graph API.
Replaces SendGrid — fully M365-native, no third-party dependency.

Auth: Client Credentials (application permission: Mail.Send)
      Credentials read from Windows environment variables — NEVER hard-coded.

Required App Registration permissions (application, not delegated):
    Mail.Send  (Microsoft Graph)   — admin consent required

The sender UPN must be a licensed Exchange Online mailbox in the tenant.
"""

import os
import base64
import datetime
import requests


# ──────────────────────────────────────────────
# Token Acquisition
# ──────────────────────────────────────────────

def get_graph_token(tenant_id: str,
                     client_id: str,
                     client_secret: str) -> str:
    """
    Client Credentials grant for Microsoft Graph.
    All three values must come from Windows environment variables — see setup_env.ps1.
    Never pass hard-coded strings.
    """
    url  = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    resp = requests.post(url, data={
        "grant_type":    "client_credentials",
        "client_id":     client_id,
        "client_secret": client_secret,
        "scope":         "https://graph.microsoft.com/.default",
    }, timeout=20)
    resp.raise_for_status()
    print("[INFO] Microsoft Graph token acquired")
    return resp.json()["access_token"]


# ──────────────────────────────────────────────
# HTML Email Body Builder
# ──────────────────────────────────────────────

def _build_email_body(client_name: str,
                       period_str: str,
                       all_findings: list) -> str:
    """Build a professional HTML email body summarising the report."""
    alerts   = [f for f in all_findings if f["status"] != "NORMAL"]
    critical = [f for f in all_findings if f["status"] == "CRITICAL"]

    if critical:
        banner_bg   = "#C00000"
        banner_icon = "&#9888;"
        banner_text = f"{len(critical)} VM(s) in CRITICAL state — immediate attention required"
    elif alerts:
        banner_bg   = "#BF8F00"
        banner_icon = "&#9888;"
        banner_text = f"{len(alerts)} VM(s) require attention"
    else:
        banner_bg   = "#375623"
        banner_icon = "&#10003;"
        banner_text = "All VMs operating within normal thresholds"

    status_styles = {
        "NORMAL":   ("background:#E2EFDA;color:#375623", "Normal"),
        "WARNING":  ("background:#FFF2CC;color:#BF8F00", "Warning"),
        "CRITICAL": ("background:#FFDFD9;color:#C00000", "Critical"),
    }

    rows_html = ""
    for f in all_findings:
        style, label = status_styles.get(f["status"], ("", f["status"]))
        issues_html  = "<br>".join(f["issues"]) if f["issues"] else "—"
        rows_html += f"""
        <tr>
          <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                     font-weight:600;color:#1F4E79">{f['vm_name']}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                     color:#595959;font-size:13px">{f.get('sku','—')}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8">
            <span style="padding:3px 10px;border-radius:4px;font-size:12px;
                         font-weight:600;{style}">{label}</span>
          </td>
          <td style="padding:8px 12px;border-bottom:1px solid #e8e8e8;
                     font-size:13px;color:#333">{issues_html}</td>
        </tr>"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;font-family:Segoe UI,Arial,sans-serif;background:#F0F4F8">
<table width="100%" cellpadding="0" cellspacing="0" style="padding:30px 0">
<tr><td align="center">
<table width="640" cellpadding="0" cellspacing="0"
       style="background:#fff;border-radius:8px;
              box-shadow:0 2px 12px rgba(0,0,0,.10);overflow:hidden">

  <!-- Header bar -->
  <tr><td style="background:#1F4E79;padding:28px 32px">
    <div style="color:#fff;font-size:21px;font-weight:700;letter-spacing:-.3px">
      Azure VM Performance Report
    </div>
    <div style="color:#A8C8E8;font-size:13px;margin-top:6px">
      {client_name}&nbsp;&nbsp;|&nbsp;&nbsp;{period_str}
    </div>
  </td></tr>

  <!-- Status banner -->
  <tr><td style="background:{banner_bg};padding:13px 32px">
    <span style="color:#fff;font-size:14px;font-weight:600">
      {banner_icon}&nbsp; {banner_text}
    </span>
  </td></tr>

  <!-- Body -->
  <tr><td style="padding:28px 32px">
    <p style="margin:0 0 20px;color:#333;font-size:15px;line-height:1.65">
      Please find the automated monthly Azure VM Performance Report attached.
      The report contains CPU, memory, disk, and network metric charts for all
      monitored virtual machines, along with detailed findings and recommendations.
    </p>

    <!-- VM status table -->
    <table width="100%" cellpadding="0" cellspacing="0"
           style="border:1px solid #e0e0e0;border-radius:6px;
                  overflow:hidden;margin-bottom:24px">
      <thead>
        <tr style="background:#DEEAF1">
          <th style="padding:9px 12px;text-align:left;font-size:12px;
                     color:#1F4E79;font-weight:700;border-bottom:2px solid #BDD7EE">
            VM Name</th>
          <th style="padding:9px 12px;text-align:left;font-size:12px;
                     color:#1F4E79;font-weight:700;border-bottom:2px solid #BDD7EE">
            SKU</th>
          <th style="padding:9px 12px;text-align:left;font-size:12px;
                     color:#1F4E79;font-weight:700;border-bottom:2px solid #BDD7EE">
            Status</th>
          <th style="padding:9px 12px;text-align:left;font-size:12px;
                     color:#1F4E79;font-weight:700;border-bottom:2px solid #BDD7EE">
            Issues</th>
        </tr>
      </thead>
      <tbody>{rows_html}</tbody>
    </table>

    <p style="margin:0;color:#999;font-size:12px;line-height:1.5">
      Generated automatically on
      {datetime.datetime.utcnow().strftime('%d %B %Y at %H:%M UTC')}
      using Azure Monitor Metrics API (Read-Only access via Azure Lighthouse).
    </p>
  </td></tr>

  <!-- Footer -->
  <tr><td style="background:#F5F7FA;padding:14px 32px;
                 border-top:1px solid #eee">
    <p style="margin:0;color:#bbb;font-size:11px;text-align:center">
      Azure Infrastructure Monitoring — Automated Monthly Report
    </p>
  </td></tr>

</table>
</td></tr>
</table>
</body>
</html>"""


# ──────────────────────────────────────────────
# Graph sendMail
# ──────────────────────────────────────────────

def send_report_via_graph(graph_token: str,
                           sender_upn: str,
                           to_recipients: list,
                           cc_recipients: list,
                           subject: str,
                           body_html: str,
                           attachment_path: str) -> None:
    """
    POST /v1.0/users/{sender_upn}/sendMail

    sender_upn:      Sending mailbox UPN (must exist in M365 tenant).
    to_recipients:   List of "To" email addresses.
    cc_recipients:   List of "Cc" email addresses (can be empty).
    attachment_path: Local path to the .docx report.

    Graph permission: Mail.Send (Application) with admin consent.
    HTTP 202 = success (no response body from Graph).
    """
    with open(attachment_path, "rb") as fh:
        attachment_b64 = base64.b64encode(fh.read()).decode("utf-8")

    attachment_name = os.path.basename(attachment_path)

    message = {
        "subject": subject,
        "importance": "normal",
        "body": {
            "contentType": "HTML",
            "content": body_html,
        },
        "toRecipients": [
            {"emailAddress": {"address": addr}} for addr in to_recipients
        ],
        "attachments": [
            {
                "@odata.type":  "#microsoft.graph.fileAttachment",
                "name":         attachment_name,
                "contentType":  (
                    "application/vnd.openxmlformats-officedocument"
                    ".wordprocessingml.document"
                ),
                "contentBytes": attachment_b64,
            }
        ],
    }
    if cc_recipients:
        message["ccRecipients"] = [
            {"emailAddress": {"address": a}} for a in cc_recipients
        ]

    resp = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{sender_upn}/sendMail",
        headers={
            "Authorization":  f"Bearer {graph_token}",
            "Content-Type":   "application/json",
        },
        json={"message": message, "saveToSentItems": True},
        timeout=90,   # large attachments need more time
    )

    if resp.status_code not in (200, 202):
        raise RuntimeError(
            f"Graph sendMail failed: HTTP {resp.status_code}\n{resp.text[:500]}"
        )

    print(f"[INFO] Email sent via Microsoft Graph to: {', '.join(to_recipients)}")


# ──────────────────────────────────────────────
# High-level helper called by main.py
# ──────────────────────────────────────────────

def distribute_report(graph_token: str,
                       sender_upn: str,
                       to_recipients: list,
                       cc_recipients: list,
                       report_path: str,
                       client_name: str,
                       period_str: str,
                       all_findings: list,
                       report_month: str) -> None:
    """
    Build the email subject + HTML body and send via Graph.
    Called once per subscription (one email per client per month).
    """
    alerts   = [f for f in all_findings if f["status"] != "NORMAL"]
    critical = [f for f in all_findings if f["status"] == "CRITICAL"]

    if critical:
        tag = " [ACTION REQUIRED]"
    elif alerts:
        tag = " [ATTENTION]"
    else:
        tag = ""

    subject   = f"Azure VM Performance Report — {client_name} {report_month}{tag}"
    body_html = _build_email_body(client_name, period_str, all_findings)

    send_report_via_graph(
        graph_token=graph_token,
        sender_upn=sender_upn,
        to_recipients=to_recipients,
        cc_recipients=cc_recipients,
        subject=subject,
        body_html=body_html,
        attachment_path=report_path,
    )
