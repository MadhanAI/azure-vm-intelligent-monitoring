"""
config.py — Windows Server Configuration Loader
=================================================
Environment variables hold CREDENTIALS ONLY (short strings — no size limit issues).
Large JSON configurations live in files under VM_REPORT_CONFIG_DIR.

Why files instead of env vars for JSON configs?
  Windows setx / System env vars have a hard limit of 1,024 characters per variable.
  Large CLIENT_CONFIGS_JSON with 19 subscriptions exceeds this limit and causes
  "Unterminated string" parse errors. Storing JSON in files has no size limit and
  is easier to edit (open in Notepad, validate with a JSON linter).

Directory layout (VM_REPORT_CONFIG_DIR, default: <scripts>\\config\\):
  client_configs.json          REQUIRED — per-subscription client configuration
  subscriptions.json           REQUIRED — list of subscription IDs to process
  rg_filter.json               optional — resource group filter per subscription

Environment variables (CREDENTIALS — set via setup_env.ps1 as Administrator):
  AZURE_TENANT_ID              Managing tenant ID
  AZURE_CLIENT_ID              Service Principal client ID
  AZURE_CLIENT_SECRET          Service Principal secret   [sensitive]
  GRAPH_CLIENT_ID              App Registration client ID
  GRAPH_CLIENT_SECRET          App Registration secret    [sensitive]
  GRAPH_SENDER_UPN             Sending mailbox UPN

Optional env vars:
  VM_REPORT_CONFIG_DIR         Path to the config folder   (default: <scripts>\\config)
  REPORT_OUTPUT_DIR            Path for .docx output       (default: <scripts>\\reports)
  REPORT_LOG_DIR               Path for log files          (default: <scripts>\\logs)
  GRAPH_TENANT_ID              Override if M365 is in a different tenant
  INTERNAL_CC_EMAILS           Comma-separated CC addresses
  CPU_ALERT_THRESHOLD          Default: 80
  MEMORY_MIN_THRESHOLD         Default: 10
  DISK_UTIL_THRESHOLD          Default: 85
  STORAGE_ACCOUNT_NAME         Azure Blob Storage account (optional)
  STORAGE_CONTAINER            Blob container name (default: vm-reports)
  TEAMS_WEBHOOK_URL            Teams notification webhook (optional)
  REPORT_YEAR / REPORT_MONTH   Backfill override (default: previous calendar month)
"""

import os
import json
import datetime
import calendar


# ──────────────────────────────────────────────────────────────────────────────
# Required environment variables (credentials only — all short strings)
# ──────────────────────────────────────────────────────────────────────────────
REQUIRED_CREDENTIAL_VARS = [
    "AZURE_TENANT_ID",
    "AZURE_CLIENT_ID",
    "AZURE_CLIENT_SECRET",
    "GRAPH_CLIENT_ID",
    "GRAPH_CLIENT_SECRET",
    "GRAPH_SENDER_UPN",
]


def _req_env(name: str) -> str:
    val = os.environ.get(name, "").strip()
    if not val:
        raise EnvironmentError(
            f"Required credential env var '{name}' is not set or empty.\n"
            f"Run setup_env.ps1 as Administrator, then restart any open terminals."
        )
    return val


def _config_dir() -> str:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.environ.get(
        "VM_REPORT_CONFIG_DIR",
        os.path.join(script_dir, "config")
    ).strip()


def _load_json_file(config_dir: str, filename: str, required: bool = True):
    path = os.path.join(config_dir, filename)
    if not os.path.isfile(path):
        if required:
            raise FileNotFoundError(
                f"Required config file not found:\n  {path}\n\n"
                f"Create it with the correct JSON content.\n"
                f"Config directory: {config_dir}"
            )
        return None
    with open(path, encoding="utf-8") as f:
        raw = f.read().strip()
    if not raw:
        if required:
            raise ValueError(f"Config file is empty: {path}")
        return None
    try:
        return json.loads(raw)
    except json.JSONDecodeError as e:
        raise ValueError(
            f"JSON parse error in {path}\n"
            f"  Error: {e}\n\n"
            f"Tip: validate at https://jsonlint.com or run:\n"
            f"  python -c \"import json; json.load(open(r'{path}'))\""
        )


# ──────────────────────────────────────────────────────────────────────────────
# Calendar Month Period
# ──────────────────────────────────────────────────────────────────────────────

def get_report_period(override_year=None, override_month=None):
    today = datetime.date.today()
    if override_year and override_month:
        year, month = override_year, override_month
    else:
        prev  = today.replace(day=1) - datetime.timedelta(days=1)
        year, month = prev.year, prev.month

    _, days = calendar.monthrange(year, month)
    start_dt = datetime.datetime(year, month, 1,    0, 0, 0)
    end_dt   = datetime.datetime(year, month, days, 23, 59, 59)
    abbr     = datetime.date(year, month, 1).strftime("%b").upper()
    return (
        start_dt, end_dt,
        datetime.date(year, month, 1).strftime("%B %Y"),
        f"01 {abbr} {year} TO {days:02d} {abbr} {year}",
    )


# ──────────────────────────────────────────────────────────────────────────────
# Main Config Loader
# ──────────────────────────────────────────────────────────────────────────────

def load_config() -> dict:
    # ── 1. Credential env vars ────────────────────────────────────────────────
    missing = [v for v in REQUIRED_CREDENTIAL_VARS
               if not os.environ.get(v, "").strip()]
    if missing:
        raise EnvironmentError(
            "Missing credential env vars:\n"
            + "\n".join(f"  - {v}" for v in missing)
            + "\n\nRun setup_env.ps1 as Administrator, then restart your terminal."
        )

    # ── 2. Config directory ───────────────────────────────────────────────────
    config_dir = _config_dir()
    if not os.path.isdir(config_dir):
        raise FileNotFoundError(
            f"Config directory not found: {config_dir}\n\n"
            f"Create it and add:\n"
            f"  {config_dir}\\subscriptions.json\n"
            f"  {config_dir}\\client_configs.json\n\n"
            f"Or set VM_REPORT_CONFIG_DIR to an existing directory."
        )

    # ── 3. subscriptions.json — list of subscription IDs to process ───────────
    target_sub_ids = _load_json_file(config_dir, "subscriptions.json")
    if not isinstance(target_sub_ids, list) or not target_sub_ids:
        raise ValueError(
            "subscriptions.json must be a non-empty JSON array.\n"
            f"File: {os.path.join(config_dir, 'subscriptions.json')}\n"
            'Expected: ["sub-id-1", "sub-id-2"]'
        )

    # ── 4. client_configs.json — per-subscription config ─────────────────────
    client_configs = _load_json_file(config_dir, "client_configs.json")
    if not isinstance(client_configs, dict):
        raise ValueError("client_configs.json must be a JSON object (not array).")

    missing_cfgs = [sid for sid in target_sub_ids if sid not in client_configs]
    if missing_cfgs:
        raise ValueError(
            "These subscription IDs are in subscriptions.json but missing from "
            "client_configs.json:\n"
            + "\n".join(f"  - {sid}" for sid in missing_cfgs)
        )

    # ── 5. rg_filter.json — optional RG filter ────────────────────────────────
    rg_filter = _load_json_file(config_dir, "rg_filter.json", required=False) or {}

    # ── 6. Log Analytics workspace IDs ───────────────────────────────────────
    la_workspaces = {
        sid: cfg["log_analytics_workspace_id"]
        for sid, cfg in client_configs.items()
        if cfg.get("log_analytics_workspace_id", "").strip()
    }

    # ── 7. Calendar month period ──────────────────────────────────────────────
    override_year  = int(os.environ.get("REPORT_YEAR",  "0") or "0") or None
    override_month = int(os.environ.get("REPORT_MONTH", "0") or "0") or None
    start_dt, end_dt, month_name, period_label = get_report_period(
        override_year, override_month
    )

    # ── 8. Paths ──────────────────────────────────────────────────────────────
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.environ.get(
        "REPORT_OUTPUT_DIR", os.path.join(script_dir, "reports")
    ).strip()

    graph_tenant = (
        os.environ.get("GRAPH_TENANT_ID", "").strip()
        or os.environ["AZURE_TENANT_ID"]
    )
    internal_cc = [
        e.strip()
        for e in os.environ.get("INTERNAL_CC_EMAILS", "").split(",")
        if e.strip()
    ]

    return {
        "azure_tenant_id":     _req_env("AZURE_TENANT_ID"),
        "azure_client_id":     _req_env("AZURE_CLIENT_ID"),
        "azure_client_secret": _req_env("AZURE_CLIENT_SECRET"),
        "graph_tenant_id":     graph_tenant,
        "graph_client_id":     _req_env("GRAPH_CLIENT_ID"),
        "graph_client_secret": _req_env("GRAPH_CLIENT_SECRET"),
        "graph_sender_upn":    _req_env("GRAPH_SENDER_UPN"),
        "target_subscription_ids":      target_sub_ids,
        "target_resource_groups":       rg_filter,
        "log_analytics_workspace_ids":  la_workspaces,
        "client_configs":               client_configs,
        "report_start_dt":     start_dt,
        "report_end_dt":       end_dt,
        "report_month_name":   month_name,
        "report_period_label": period_label,
        "cpu_threshold":    float(os.environ.get("CPU_ALERT_THRESHOLD",  "80")),
        "memory_threshold": float(os.environ.get("MEMORY_MIN_THRESHOLD", "10")),
        "disk_threshold":   float(os.environ.get("DISK_UTIL_THRESHOLD",  "85")),
        "output_dir":       output_dir,
        "internal_cc":      internal_cc,
        "storage_account_name": os.environ.get("STORAGE_ACCOUNT_NAME", "").strip(),
        "storage_container":    os.environ.get("STORAGE_CONTAINER", "vm-reports"),
        "teams_webhook_url":    os.environ.get("TEAMS_WEBHOOK_URL", "").strip(),
    }
