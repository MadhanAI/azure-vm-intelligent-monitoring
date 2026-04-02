"""
main.py — Windows Server Entrypoint
=====================================
Replaces runbook_main.py (Azure Automation).
Runs on a Windows Server, scheduled via Windows Task Scheduler.

Authentication:
    Azure ARM / Metrics  →  Service Principal (env vars AZURE_*)
    Microsoft Graph      →  App Registration  (env vars GRAPH_*)
    No Managed Identity, no Azure Key Vault.

All configuration is read from Windows System Environment Variables.
See setup_env.ps1 for the full list.

Usage (direct):
    python main.py

Usage (backfill a specific month):
    set REPORT_YEAR=2025 && set REPORT_MONTH=9 && python main.py

Usage (via Task Scheduler):
    run_report.bat  (calls this script, writes to log file)
"""

import os
import sys
import datetime
import logging
import sys, os
sys.path.insert(0, os.getcwd())

# Configure file + console logging before any imports that emit output
LOG_DIR = os.environ.get("REPORT_LOG_DIR",
                          os.path.join(os.path.dirname(__file__), "logs"))
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(
    LOG_DIR,
    f"report_{datetime.datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.log"
)

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ]
)
log = logging.getLogger(__name__)


# ── Local module imports ───────────────────────────────────────────────
from config import load_config
from collect_metrics import (
    ReportConfig,
    collect_all_tenants,
    get_arm_token,
    get_log_analytics_token,
)
from generate_report import generate_report
from graph_mailer import get_graph_token, distribute_report


# ──────────────────────────────────────────────
# Optional Blob Storage upload (requires azure-storage-blob)
# ──────────────────────────────────────────────

def _upload_to_blob(local_path: str, account: str,
                    container: str, blob_name: str,
                    tenant_id: str, client_id: str, client_secret: str) -> str:
    """
    Upload report to Azure Blob Storage using Service Principal.
    Only called if STORAGE_ACCOUNT_NAME is set.
    Returns the blob URL.
    """
    try:
        from azure.identity import ClientSecretCredential
        from azure.storage.blob import BlobServiceClient
    except ImportError:
        log.warning("azure-storage-blob not installed — skipping Blob upload")
        return local_path

    credential   = ClientSecretCredential(tenant_id, client_id, client_secret)
    account_url  = f"https://{account}.blob.core.windows.net"
    blob_client  = BlobServiceClient(account_url, credential).get_blob_client(
        container=container, blob=blob_name
    )
    with open(local_path, "rb") as f:
        blob_client.upload_blob(f, overwrite=True)

    url = f"{account_url}/{container}/{blob_name}"
    log.info(f"Blob uploaded: {url}")
    return url


# ──────────────────────────────────────────────
# Optional Teams notification
# ──────────────────────────────────────────────

def _post_teams(webhook_url: str, report_url: str,
                 findings: list, period_str: str, client_name: str):
    import requests
    alerts = [f for f in findings if f["status"] != "NORMAL"]
    color  = "attention" if alerts else "good"
    text   = f"{len(alerts)} VM(s) need attention" if alerts else "All VMs healthy"
    card = {
        "type": "message",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard", "version": "1.4",
                "body": [
                    {"type": "TextBlock",
                     "text": f"Azure VM Report — {client_name}",
                     "weight": "Bolder", "size": "Large", "color": color},
                    {"type": "TextBlock", "text": period_str, "isSubtle": True},
                    {"type": "TextBlock", "text": text, "weight": "Bolder"},
                    {"type": "FactSet",
                     "facts": [{"title": f["vm_name"], "value": f["status"]}
                                for f in findings]},
                ],
                "actions": [{"type": "Action.OpenUrl",
                              "title": "Download Report", "url": report_url}],
            },
        }]
    }
    try:
        requests.post(webhook_url, json=card, timeout=15).raise_for_status()
        log.info(f"Teams notification sent for {client_name}")
    except Exception as e:
        log.warning(f"Teams notification failed: {e}")


# ──────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────

def main() -> int:
    """
    Returns 0 on success, 1 on any error (for Task Scheduler exit code checking).
    """
    log.info("=" * 60)
    log.info("Azure VM Performance Report — starting")
    log.info("=" * 60)

    # ── 1. Load config from Windows env vars ──────────────────────────
    try:
        cfg = load_config()
    except (EnvironmentError, ValueError) as e:
        log.error(f"Configuration error:\n{e}")
        return 1

    period_str = (
        f"{cfg['report_start_dt'].strftime('%d %b %Y')} to "
        f"{cfg['report_end_dt'].strftime('%d %b %Y')}"
    )
    log.info(f"Report period : {period_str}")
    log.info(f"Target subs   : {len(cfg['target_subscription_ids'])}")
    for sid in cfg["target_subscription_ids"]:
        cname = cfg["client_configs"].get(sid, {}).get("client_name", sid)
        log.info(f"               [{sid}] {cname}")

    # ── 2. Acquire Azure ARM + LA tokens (Service Principal) ──────────
    log.info("Acquiring Azure ARM token (Service Principal)...")
    try:
        arm_token = get_arm_token(
            tenant_id=cfg["azure_tenant_id"],
            client_id=cfg["azure_client_id"],
            client_secret=cfg["azure_client_secret"],
        )
        la_token = get_log_analytics_token(
            tenant_id=cfg["azure_tenant_id"],
            client_id=cfg["azure_client_id"],
            client_secret=cfg["azure_client_secret"],
        )
    except Exception as e:
        log.error(f"Azure token acquisition failed: {e}")
        return 1

    # ── 3. Acquire Microsoft Graph token ─────────────────────────────
    log.info("Acquiring Microsoft Graph token (App Registration)...")
    try:
        graph_token = get_graph_token(
            tenant_id=cfg["graph_tenant_id"],
            client_id=cfg["graph_client_id"],
            client_secret=cfg["graph_client_secret"],
        )
    except Exception as e:
        log.error(f"Graph token acquisition failed: {e}")
        return 1

    # ── 4. Build ReportConfig ─────────────────────────────────────────
    os.makedirs(cfg["output_dir"], exist_ok=True)

    report_config = ReportConfig(
        target_subscription_ids=cfg["target_subscription_ids"],
        target_resource_groups=cfg["target_resource_groups"],
        log_analytics_workspace_ids=cfg["log_analytics_workspace_ids"],
        client_configs=cfg["client_configs"],
        report_start=cfg["report_start_dt"],
        report_end=cfg["report_end_dt"],
        report_month_name=cfg["report_month_name"],
        report_period_label=cfg["report_period_label"],
        cpu_alert_threshold=cfg["cpu_threshold"],
        memory_min_threshold=cfg["memory_threshold"],
        disk_util_threshold=cfg["disk_threshold"],
        output_dir=cfg["output_dir"],
    )

    # ── 5. Collect metrics from all target Lighthouse subscriptions ───
    log.info("Collecting metrics from Lighthouse-delegated subscriptions...")
    try:
        tenant_results = collect_all_tenants(
            config=report_config,
            arm_token=arm_token,
            la_token=la_token,
        )
    except Exception as e:
        log.error(f"Metrics collection failed: {e}")
        return 1

    if not tenant_results:
        log.warning("No data collected — no target subscriptions found or accessible")
        return 1

    log.info(f"Metrics collected from {len(tenant_results)} subscription(s)")

    # ── 6. Per-subscription: generate report + distribute ─────────────
    report_month = cfg["report_end_dt"].strftime("%B %Y")
    errors = 0

    for subscription_id, (all_vm_metrics, all_findings) in tenant_results.items():
        client_cfg   = cfg["client_configs"].get(subscription_id, {})
        client_name  = client_cfg.get("client_name", subscription_id)
        to_recipients = client_cfg.get("to_recipients", [])
        cc_recipients = list(set(
            client_cfg.get("cc_recipients", []) + cfg.get("internal_cc", [])
        ))

        log.info(f"\n── {client_name} ({subscription_id})")
        log.info(f"   VMs: {len(all_vm_metrics)}  |  "
                 f"Alerts: {sum(1 for f in all_findings if f['status'] != 'NORMAL')}")

        # Build report filename: safe for Windows filesystem
        safe_name = (
            client_name
            .replace(" ", "_")
            .replace("/", "-")
            .replace("\\", "-")
            .replace(":", "")
        )
        report_fname = (
            f"Azure_VM_Report_{safe_name}"
            f"_{cfg['report_month_name'].replace(' ', '_')}.docx"
        )
        report_path = os.path.join(cfg["output_dir"], report_fname)

        # Subscription display name (from first VM or sub ID)
        sub_display = (
            all_vm_metrics[0].subscription_id if all_vm_metrics else subscription_id
        )
        # Try to get the readable name from the metrics
        # (stored in tenant_id field — use client_name instead)
        sub_display = client_name

        try:
            generate_report(
                all_vm_metrics=all_vm_metrics,
                all_findings=all_findings,
                config=report_config,
                client_name=client_name,
                subscription_name=sub_display,
                output_path=report_path,
            )
        except Exception as e:
            log.error(f"Report generation failed for {client_name}: {e}")
            errors += 1
            continue

        report_url = report_path

        # Optional: upload to Blob Storage
        if cfg.get("storage_account_name"):
            try:
                blob_name = (
                    f"{subscription_id}/"
                    f"{cfg['report_end_dt'].strftime('%Y/%m')}/"
                    f"{report_fname}"
                )
                report_url = _upload_to_blob(
                    local_path=report_path,
                    account=cfg["storage_account_name"],
                    container=cfg["storage_container"],
                    blob_name=blob_name,
                    tenant_id=cfg["azure_tenant_id"],
                    client_id=cfg["azure_client_id"],
                    client_secret=cfg["azure_client_secret"],
                )
            except Exception as e:
                log.warning(f"Blob upload failed (report still saved locally): {e}")

        # Send via Microsoft Graph
        if to_recipients:
            try:
                distribute_report(
                    graph_token=graph_token,
                    sender_upn=cfg["graph_sender_upn"],
                    to_recipients=to_recipients,
                    cc_recipients=cc_recipients,
                    report_path=report_path,
                    client_name=client_name,
                    period_str=period_str,
                    all_findings=all_findings,
                    report_month=report_month,
                )
            except Exception as e:
                log.error(f"Email send failed for {client_name}: {e}")
                errors += 1
        else:
            log.warning(f"No to_recipients configured for {client_name} — email skipped")

        # Optional: Teams notification
        if cfg.get("teams_webhook_url"):
            _post_teams(
                webhook_url=cfg["teams_webhook_url"],
                report_url=report_url,
                findings=all_findings,
                period_str=period_str,
                client_name=client_name,
            )

    # ── 7. Summary ───────────────────────────────────────────────────
    log.info("\n" + "=" * 60)
    if errors == 0:
        log.info(f"DONE — all {len(tenant_results)} report(s) generated and sent")
    else:
        log.warning(f"DONE with {errors} error(s) — check log: {LOG_FILE}")
    log.info("=" * 60)

    return 0 if errors == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
