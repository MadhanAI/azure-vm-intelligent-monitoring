"""
collect_metrics.py — Azure VM Metrics Collector (Windows Server / Lighthouse Edition)
======================================================================================
Auth:    Service Principal via Windows environment variables (no Managed Identity).
Access:  Lighthouse-delegated Reader + Monitoring Reader on customer subscriptions.
Period:  Exact calendar month passed in as explicit start/end datetimes.
Filter:  ONLY processes subscriptions in config.target_subscription_ids (required).
         If a subscription is NOT in the list it is silently skipped even if delegated.
"""

import os
import datetime
import requests
from dataclasses import dataclass, field
from typing import Optional


# ──────────────────────────────────────────────
# Data Models
# ──────────────────────────────────────────────

@dataclass
class TenantSubscription:
    """One Lighthouse-delegated subscription visible to the SP."""
    subscription_id: str
    subscription_name: str
    tenant_id: str


@dataclass
class ReportConfig:
    # REQUIRED — explicit list; empty list is a hard error (see config.py)
    target_subscription_ids: list

    # Exact report period (datetime objects, calendar month boundaries)
    report_start: datetime.datetime = None
    report_end: datetime.datetime = None
    report_month_name: str = ""        # e.g. "September 2025"
    report_period_label: str = ""      # e.g. "01 SEP 2025 TO 30 SEP 2025"

    # Optional RG filter: {"sub-id": ["rg1", "rg2"]}
    target_resource_groups: dict = field(default_factory=dict)

    # LA workspace per subscription: {"sub-id": "workspace-id"}
    log_analytics_workspace_ids: dict = field(default_factory=dict)

    # Per-client metadata: {"sub-id": {"client_name": ..., "to_recipients": [...]}}
    client_configs: dict = field(default_factory=dict)

    # Alert thresholds
    cpu_alert_threshold: float = 80.0
    memory_min_threshold: float = 10.0
    disk_util_threshold: float = 85.0

    output_dir: str = "./reports"


@dataclass
class VMMetrics:
    vm_name: str
    sku: str
    vcpus: int
    memory_gib: float
    resource_id: str
    location: str
    subscription_id: str
    resource_group: str
    tenant_id: str = ""
    # Time-series: list of (iso_timestamp_str, float_value)
    # Average series — used for trend lines and alert threshold comparisons
    cpu_percent: list = field(default_factory=list)
    available_memory_bytes: list = field(default_factory=list)
    disk_read_iops: list = field(default_factory=list)
    disk_write_iops: list = field(default_factory=list)
    network_bytes_sent: list = field(default_factory=list)
    network_bytes_received: list = field(default_factory=list)
    # Maximum series — collected in the same API call, shows peak within each hour
    cpu_percent_max: list = field(default_factory=list)
    available_memory_bytes_max: list = field(default_factory=list)
    disk_read_iops_max: list = field(default_factory=list)
    disk_write_iops_max: list = field(default_factory=list)
    network_bytes_sent_max: list = field(default_factory=list)
    network_bytes_received_max: list = field(default_factory=list)
    disk_utilization: dict = field(default_factory=dict)   # {drive: used_pct}
    # Explains where disk data came from, or why it is absent:
    #   "InsightsMetrics"  — Azure Monitor Agent (new VM Insights / AMA)
    #   "Perf"             — Legacy MMA agent (classic Log Analytics agent)
    #   "no_workspace"     — log_analytics_workspace_id not configured in CLIENT_CONFIGS_JSON
    #   "no_data"          — workspace configured but no rows in InsightsMetrics or Perf
    disk_source: str = "no_workspace"
    # Data source for performance metrics (CPU, Mem, IOPS, Network)
    #   "LA-AMA"  — Log Analytics, InsightsMetrics table (Azure Monitor Agent)
    #   "LA-MMA"  — Log Analytics, Perf table (legacy MMA agent)
    #   "ARM"     — Azure Monitor REST API (no Log Analytics workspace)
    #   "mixed"   — different metrics came from different sources (rare)
    #   "none"    — no data collected
    metric_source: str = "none"
    # NSG rules collected per VM (populated by fetch_vm_nsg_rules)
    nsg_rules: list = field(default_factory=list)   # list of NSGRule objects
    nsg_names: list = field(default_factory=list)   # e.g. ["NIC: lfg-nsg", "Subnet: sub-nsg"]
    # Computed summary stats (populated by analyze_vm_metrics)
    cpu_max: float = 0.0
    cpu_avg: float = 0.0
    cpu_threshold_breaches: list = field(default_factory=list)  # [(date_str, pct)]
    memory_min_available_gb: float = 0.0
    has_alerts: bool = False


# ──────────────────────────────────────────────
# Authentication — Service Principal (Windows env vars)
# ──────────────────────────────────────────────

def _sp_token(tenant_id: str, client_id: str,
              client_secret: str, resource: str) -> str:
    """
    Acquire a token via Service Principal client credentials grant.
    Called with credentials read from Windows environment variables — never
    pass hard-coded values.
    """
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/token"
    resp = requests.post(url, data={
        "grant_type":    "client_credentials",
        "client_id":     client_id,
        "client_secret": client_secret,
        "resource":      resource,
    }, timeout=20)
    resp.raise_for_status()
    return resp.json()["access_token"]


def get_arm_token(tenant_id: str = None,
                   client_id: str = None,
                   client_secret: str = None) -> str:
    """
    ARM / Azure Monitor token.
    Lighthouse delegation: single token covers all delegated customer subscriptions.
    ARM resolves the cross-tenant delegation server-side; no per-tenant token needed.
    """
    return _sp_token(
        tenant_id     or os.environ["AZURE_TENANT_ID"],
        client_id     or os.environ["AZURE_CLIENT_ID"],
        client_secret or os.environ["AZURE_CLIENT_SECRET"],
        "https://management.azure.com/",
    )


def get_log_analytics_token(tenant_id: str = None,
                              client_id: str = None,
                              client_secret: str = None) -> str:
    """
    Log Analytics API token (different audience from ARM).
    Same SP, same Lighthouse delegation covers Log Analytics Reader.
    """
    return _sp_token(
        tenant_id     or os.environ["AZURE_TENANT_ID"],
        client_id     or os.environ["AZURE_CLIENT_ID"],
        client_secret or os.environ["AZURE_CLIENT_SECRET"],
        "https://api.loganalytics.io/",
    )


# ──────────────────────────────────────────────
# Lighthouse — Subscription Discovery
# ──────────────────────────────────────────────

def list_lighthouse_subscriptions(arm_token: str,
                                   allowed_ids: list) -> list:
    """
    GET /subscriptions returns the managing tenant's own subscriptions AND all
    Lighthouse-delegated customer subscriptions in one paginated call.

    IMPORTANT: only subscriptions whose ID is in allowed_ids are returned.
    allowed_ids must be a non-empty list (enforced by config.py).
    Subscriptions NOT in the list are silently ignored even if delegated.

    Permission: Reader delegation via Lighthouse (ARM resolves automatically).
    """
    url     = "https://management.azure.com/subscriptions?api-version=2022-12-01"
    headers = {"Authorization": f"Bearer {arm_token}"}
    subs    = []

    while url:
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        body = resp.json()
        for sub in body.get("value", []):
            if sub.get("state") != "Enabled":
                continue
            sid = sub["subscriptionId"]
            if sid not in allowed_ids:
                continue   # strict filter — not in the explicit list → skip
            subs.append(TenantSubscription(
                subscription_id=sid,
                subscription_name=sub.get("displayName", sid),
                tenant_id=sub.get("tenantId", ""),
            ))
        url = body.get("nextLink")

    found_ids = {s.subscription_id for s in subs}
    missed    = [sid for sid in allowed_ids if sid not in found_ids]
    if missed:
        print(f"[WARN] The following subscription IDs were not found / not delegated:")
        for sid in missed:
            print(f"[WARN]   - {sid}")

    print(f"[INFO] Lighthouse: {len(subs)} of {len(allowed_ids)} target subscription(s) found")
    return subs


# ──────────────────────────────────────────────
# VM Discovery
# ──────────────────────────────────────────────

def list_all_vms_in_subscription(subscription_id: str,
                                  arm_token: str) -> list:
    """Subscription-wide VM list (across all resource groups). Read-only."""
    url     = (
        f"https://management.azure.com/subscriptions/{subscription_id}"
        f"/providers/Microsoft.Compute/virtualMachines?api-version=2023-03-01"
    )
    headers = {"Authorization": f"Bearer {arm_token}"}
    vms     = []
    while url:
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        body = resp.json()
        for vm in body.get("value", []):
            rg = vm["id"].split("/resourceGroups/")[1].split("/")[0]
            vms.append({
                "name":           vm["name"],
                "resource_id":    vm["id"],
                "location":       vm["location"],
                "vm_size":        vm["properties"]["hardwareProfile"]["vmSize"],
                "resource_group": rg,
            })
        url = body.get("nextLink")
    return vms


def list_vms_in_resource_group(subscription_id: str, resource_group: str,
                                arm_token: str) -> list:
    """List VMs in a specific resource group. Read-only."""
    url = (
        f"https://management.azure.com/subscriptions/{subscription_id}"
        f"/resourceGroups/{resource_group}"
        f"/providers/Microsoft.Compute/virtualMachines?api-version=2023-03-01"
    )
    resp = requests.get(url, headers={"Authorization": f"Bearer {arm_token}"}, timeout=30)
    resp.raise_for_status()
    return [
        {
            "name":           vm["name"],
            "resource_id":    vm["id"],
            "location":       vm["location"],
            "vm_size":        vm["properties"]["hardwareProfile"]["vmSize"],
            "resource_group": resource_group,
        }
        for vm in resp.json().get("value", [])
    ]


def get_vm_size_details(subscription_id: str, location: str,
                        vm_size: str, arm_token: str) -> dict:
    """Get vCPU count and memory for a VM SKU. Read-only."""
    url = (
        f"https://management.azure.com/subscriptions/{subscription_id}"
        f"/providers/Microsoft.Compute/locations/{location}"
        f"/vmSizes?api-version=2023-03-01"
    )
    resp = requests.get(url, headers={"Authorization": f"Bearer {arm_token}"}, timeout=30)
    resp.raise_for_status()
    for s in resp.json().get("value", []):
        if s["name"] == vm_size:
            return {"vcpus": s["numberOfCores"], "memory_mb": s["memoryInMB"]}
    return {"vcpus": 0, "memory_mb": 0}


# ──────────────────────────────────────────────
# Metrics Fetching (Azure Monitor REST API)
# ──────────────────────────────────────────────

def fetch_metric(resource_id: str, metric_name: str,
                 start_dt: datetime.datetime, end_dt: datetime.datetime,
                 arm_token: str,
                 aggregation: str = "Average",
                 interval: str = "PT1H") -> dict:
    """
    Fetch one metric from Azure Monitor Metrics API requesting BOTH Average
    and Maximum in a single HTTP call (comma-separated aggregation parameter).

    Azure Monitor returns both values per data point when both are requested —
    no extra API calls needed, no extra quota consumed.

    Returns a dict:
        {
          "Average": [(iso_ts_str, float), ...],
          "Maximum": [(iso_ts_str, float), ...],
        }
    The "Total" aggregation (used for network counters) populates "Average" only
    since Maximum on a Total counter is not meaningful for trending.

    Permission: Monitoring Reader (delegated via Lighthouse).
    """
    start_str = start_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
    end_str   = end_dt.strftime("%Y-%m-%dT%H:%M:%SZ")

    # Build aggregation request.
    # For Average-type metrics: request Average + Maximum together.
    # For Total-type metrics (Network In/Out):
    #   Request Total + Maximum only — NOT "Average,Maximum,Total".
    #   Reason: Azure Monitor's "Average" aggregation for a Total-type metric
    #   at PT1H interval returns the average of the per-minute samples, where
    #   each per-minute sample = bytes transferred in that minute.
    #   That is bytes/minute, NOT bytes/second.
    #   "Total" aggregation = sum of all per-minute samples = total bytes/hour.
    #   Dividing total by 3600 gives the correct avg bytes/sec.
    if aggregation == "Total":
        agg_request = "Total,Maximum"
    else:
        agg_request = "Average,Maximum"

    url = f"https://management.azure.com{resource_id}/providers/microsoft.insights/metrics"
    params = {
        "api-version": "2023-10-01",
        "metricnames": metric_name,
        "aggregation": agg_request,
        "interval":    interval,
        "timespan":    f"{start_str}/{end_str}",
    }
    resp = requests.get(
        url,
        headers={"Authorization": f"Bearer {arm_token}"},
        params=params,
        timeout=30,
    )
    resp.raise_for_status()

    avg_pts, max_pts = [], []
    for metric in resp.json().get("value", []):
        for ts in metric.get("timeseries", []):
            for dp in ts.get("data", []):
                timestamp = dp["timeStamp"]
                # For Total-type metrics: dp["total"] = bytes/hour (no "average" key present)
                # For Average-type metrics: dp["average"] = the average value
                avg_val = dp.get("average") if dp.get("average") is not None                           else dp.get("total")
                if avg_val is not None:
                    avg_pts.append((timestamp, float(avg_val)))
                # Maximum value (present for all aggregation types)
                max_val = dp.get("maximum")
                if max_val is not None:
                    max_pts.append((timestamp, float(max_val)))

    return {"Average": avg_pts, "Maximum": max_pts}


def _run_kql(workspace_id: str, kql: str, la_token: str) -> dict:
    """
    Execute a KQL query expecting columns (Drive, UsedPct).
    Returns {drive: used_pct} or {} if no rows / query fails.
    Internal helper — not called directly.
    """
    resp = requests.post(
        f"https://api.loganalytics.io/v1/workspaces/{workspace_id}/query",
        headers={
            "Authorization": f"Bearer {la_token}",
            "Content-Type":  "application/json",
        },
        json={"query": kql},
        timeout=30,
    )
    resp.raise_for_status()
    result = {}
    for tbl in resp.json().get("tables", []):
        cols = [c["name"] for c in tbl["columns"]]
        for row in tbl["rows"]:
            rec = dict(zip(cols, row))
            drive    = str(rec.get("Drive", "")).strip()
            used_pct = rec.get("UsedPct")
            if drive and used_pct is not None:
                result[drive] = round(float(used_pct), 1)
    return result


def fetch_disk_utilization_via_log_analytics(workspace_id: str,
                                              vm_name: str,
                                              start_dt: datetime.datetime,
                                              end_dt: datetime.datetime,
                                              la_token: str) -> dict:
    """
    Fetch disk utilisation from Log Analytics using a three-strategy cascade.

    IMPORTANT — why we do NOT restrict by the report period:
        The Azure Portal's "CURRENT USED (%)" column in the Logical Disk
        Performance table is a live reading from the VM agent, independent of
        whatever time range is selected. It always shows today's value.

        If we restrict the KQL to the report period (e.g. Feb 1-28), arg_max
        returns the last reading within that window. For mount points with
        sparse data (/boot/efi, /mnt) or mounts where usage changed after the
        period ended, the last in-window reading can differ significantly from
        the current portal value.

        Solution: query the last 7 days relative to NOW and use arg_max to get
        the absolute most recent reading. This matches the portal exactly and
        is the correct semantics for a "current disk health" metric — you want
        to know if a disk is nearly full TODAY, not what it was a month ago.

    Strategy 1 — InsightsMetrics (AMA / new VM Insights)
    Strategy 2 — Perf table (legacy MMA)
    Strategy 3 — Empty dict (no data found)

    Returns: {mount_point: used_pct_float}
             e.g. {"C:": 67.0, "D:": 47.0, "/": 10.0, "/boot/efi": 10.0}
    """
    # ── Filesystem exclusion filters ─────────────────────────────────────────
    # Goal: show exactly what the Azure Portal "Logical Disk Performance" table shows.
    #
    # The portal SHOWS:
    #   /             — root filesystem (ext4/xfs, real storage)
    #   /boot/efi     — EFI system partition (real storage)
    #   /mnt          — any attached data disk (real storage)
    #   /snap/core20  — snap squashfs (real block device, always 100% — read-only by design)
    #   /snap/lxd     — snap squashfs (real block device, always 100%)
    #   C:, D:        — Windows NTFS volumes
    #
    # The portal does NOT show (and neither should we):
    #   /sys, /sys/*  — kernel sysfs / cgroupfs  (virtual, no backing device)
    #   /proc         — kernel procfs             (virtual)
    #   /dev/*        — devtmpfs / udev           (virtual kernel device nodes)
    #   /run, /run/*  — tmpfs runtime             (in-memory, not persistent storage)
    #   /dev/shm      — POSIX shared memory       (in-memory)
    #   overlay       — Docker/container OverlayFS (container layers, not host disk)
    #
    # Note: /snap/* are NOT excluded. They are real squashfs block devices.
    # Showing them 100% is correct — squashfs is always 100% used (read-only, compressed).
    _LINUX_EXCLUDE_AMA = (
        '| where Drive !startswith "HarddiskVolume"\n'    # Windows: recovery partitions
        '| where not(Drive startswith "/sys")\n'           # Linux: kernel sysfs/cgroupfs
        '| where not(Drive startswith "/proc")\n'          # Linux: kernel procfs
        '| where not(Drive startswith "/dev/")\n'          # Linux: kernel devtmpfs
        '| where Drive != "/run"\n'                        # Linux: tmpfs runtime root
        '| where not(Drive startswith "/run/")\n'          # Linux: tmpfs runtime children
        '| where Drive != "/dev/shm"\n'                    # Linux: POSIX shared memory
        '| where Drive !contains "overlay"\n'              # Linux: Docker overlay layers
    )
    _LINUX_EXCLUDE_PERF = (
        '| where InstanceName !in ("_Total", "HarddiskVolume1")\n'  # Windows: aggregate + recovery
        '| where not(InstanceName startswith "/sys")\n'
        '| where not(InstanceName startswith "/proc")\n'
        '| where not(InstanceName startswith "/dev/")\n'
        '| where InstanceName != "/run"\n'
        '| where not(InstanceName startswith "/run/")\n'
        '| where InstanceName != "/dev/shm"\n'
        '| where InstanceName !contains "overlay"\n'
    )

    # ── Strategy 1: InsightsMetrics (AMA / new VM Insights) ──────────────────
    kql_insights = (
        f'InsightsMetrics\n'
        # Lookback window: 30 days from NOW (not the report period).
        #
        # The portal "CURRENT USED (%)" is a live reading, not historical.
        # We use arg_max to get the most recent reading per drive, regardless
        # of when it was written. 30 days ensures we capture sparse-write mounts
        # like /boot/efi and /mnt which may only be sampled every few days.
        f'| where TimeGenerated >= ago(30d)\n'
        f'| where Computer contains "{vm_name}"\n'
        f'| where Namespace == "LogicalDisk" and Name == "FreeSpacePercentage"\n'
        f'| extend Drive = tostring(parse_json(Tags)["vm.azm.ms/mountId"])\n'
        f'| where isnotempty(Drive)\n'
    ) + _LINUX_EXCLUDE_AMA + (
        # arg_max(TimeGenerated, Val): returns one row per drive with the
        # latest TimeGenerated, giving the most recent FreeSpacePercentage.
        f'| summarize arg_max(TimeGenerated, Val) by Drive\n'
        f'| project Drive, UsedPct = round(100.0 - Val, 1)\n'
        # Sanity guard: 0–100 only. Outside range = corrupt/missing data.
        f'| where UsedPct >= 0.0 and UsedPct <= 100.0'
    )
    try:
        result = _run_kql(workspace_id, kql_insights, la_token)
        if result:
            print(f"[INFO]       Disk util source: InsightsMetrics (AMA) — "
                  f"{len(result)} drive(s)  [most recent reading]")
            return result
    except Exception as e:
        print(f"[INFO]       InsightsMetrics query skipped: {e}")

    # ── Strategy 2: Perf table (legacy MMA) ──────────────────────────────────
    kql_perf = (
        f'Perf\n'
        f'| where TimeGenerated >= ago(30d)\n'
        f'| where Computer contains "{vm_name}"\n'
        f'| where ObjectName == "Logical Disk" and CounterName == "% Free Space"\n'
        f'| where isnotempty(InstanceName)\n'
    ) + _LINUX_EXCLUDE_PERF + (
        f'| summarize arg_max(TimeGenerated, CounterValue) by InstanceName\n'
        f'| project Drive = InstanceName, UsedPct = round(100.0 - CounterValue, 1)\n'
        f'| where UsedPct >= 0.0 and UsedPct <= 100.0'
    )
    try:
        result = _run_kql(workspace_id, kql_perf, la_token)
        if result:
            print(f"[INFO]       Disk util source: Perf table (MMA) — "
                  f"{len(result)} drive(s)  [most recent reading]")
            return result
    except Exception as e:
        print(f"[INFO]       Perf query skipped: {e}")

    print(f"[WARN]       Disk util: no data in InsightsMetrics or Perf for '{vm_name}'. "
          f"Ensure the VM's Data Collection Rule sends LogicalDisk performance counters "
          f"to workspace {workspace_id}.")
    return {}


def _run_la_metrics_kql(workspace_id: str, kql: str,
                         la_token: str) -> dict:
    """
    Execute a metrics KQL query that returns (MetricName, TimeGenerated, AvgVal, MaxVal).
    Returns empty METRIC_TEMPLATE on any error.
    Internal helper — not called directly.
    """
    _EMPTY = {
        "Percentage CPU":            {"Average": [], "Maximum": []},
        "Available Memory Bytes":    {"Average": [], "Maximum": []},
        "Disk Read Operations/Sec":  {"Average": [], "Maximum": []},
        "Disk Write Operations/Sec": {"Average": [], "Maximum": []},
        "Network In Total":          {"Average": [], "Maximum": []},
        "Network Out Total":         {"Average": [], "Maximum": []},
    }
    try:
        resp = requests.post(
            f"https://api.loganalytics.io/v1/workspaces/{workspace_id}/query",
            headers={"Authorization": f"Bearer {la_token}",
                     "Content-Type":  "application/json"},
            json={"query": kql},
            timeout=60,
        )
        resp.raise_for_status()
    except Exception as e:
        raise RuntimeError(f"LA query failed: {e}")

    data = {k: {"Average": [], "Maximum": []} for k in _EMPTY}
    for tbl in resp.json().get("tables", []):
        cols = [c["name"] for c in tbl["columns"]]
        for row in tbl["rows"]:
            rec   = dict(zip(cols, row))
            mname = rec.get("MetricName")
            ts    = rec.get("TimeGenerated")
            avg_v = rec.get("AvgVal")
            max_v = rec.get("MaxVal")
            if mname in data and ts and avg_v is not None and max_v is not None:
                data[mname]["Average"].append((ts, float(avg_v)))
                data[mname]["Maximum"].append((ts, float(max_v)))
    return data


def _fetch_metrics_ama(workspace_id: str, vm_name: str,
                        start_str: str, end_str: str,
                        la_token: str) -> dict:
    """
    Query InsightsMetrics (Azure Monitor Agent / new VM Insights) for all
    performance metrics.

    AMA stores metrics in InsightsMetrics with structured Namespace/Name fields.
    Metrics are already in per-second rates where applicable — no unit conversion
    needed for CPU%, memory (MiB → bytes), IOPS (ops/sec), or network (bytes/sec).

    Aggregation strategy (matches what Azure Portal reports):
      Step 1: avg per-instance per 1-minute bin  — handles 10s/15s polling frequency
      Step 2: sum instances per minute           — adds C: + D: disks, eth0 + eth1 NICs
      Step 3: avg and max into 1-hour bins       — hourly Average and peak-minute Maximum
    """
    kql = f"""
let startT = datetime("{start_str}");
let endT   = datetime("{end_str}");
InsightsMetrics
| where TimeGenerated between (startT .. endT)
| where Computer contains "{vm_name}"
| where (Namespace == "Processor"   and Name == "UtilizationPercentage")
     or (Namespace == "Memory"      and Name == "AvailableMB")
     or (Namespace == "LogicalDisk" and Name in ("ReadsPerSecond", "WritesPerSecond"))
     or (Namespace == "Network"     and Name in ("BytesReceivedPerSecond",
                                                  "BytesTransmittedPerSecond"))
| extend MetricName = case(
    Namespace == "Processor"   and Name == "UtilizationPercentage",  "Percentage CPU",
    Namespace == "Memory"      and Name == "AvailableMB",            "Available Memory Bytes",
    Namespace == "LogicalDisk" and Name == "ReadsPerSecond",         "Disk Read Operations/Sec",
    Namespace == "LogicalDisk" and Name == "WritesPerSecond",        "Disk Write Operations/Sec",
    Namespace == "Network"     and Name == "BytesReceivedPerSecond", "Network In Total",
    Namespace == "Network"     and Name == "BytesTransmittedPerSecond", "Network Out Total",
    "")
| where isnotempty(MetricName)
// Drive/NIC instance from Tags JSON — used to sum per-instance then aggregate
| extend Instance = tostring(parse_json(Tags)["vm.azm.ms/mountId"])
// Exclude _Total aggregate for disk (we sum individual drives ourselves)
| where not(MetricName startswith "Disk" and Instance == "_Total")
// Convert AvailableMB → bytes (MiB to bytes: ×1,048,576)
| extend Value = iff(MetricName == "Available Memory Bytes", Val * 1048576.0, Val)
// Step 1: per-instance per-minute average (normalises sub-minute polling)
| summarize InstanceVal = avg(Value) by MetricName, Instance, bin(TimeGenerated, 1m)
// Step 2: sum instances per minute (total across all drives / NICs)
| summarize MinTotal = sum(InstanceVal) by MetricName, TimeGenerated
// Step 3: roll up to 1-hour buckets — AvgVal = hourly average, MaxVal = peak 1-min within hour
| summarize AvgVal = avg(MinTotal), MaxVal = max(MinTotal)
    by MetricName, bin(TimeGenerated, 1h)
| order by TimeGenerated asc
"""
    return _run_la_metrics_kql(workspace_id, kql, la_token)


def _fetch_metrics_mma(workspace_id: str, vm_name: str,
                        start_str: str, end_str: str,
                        la_token: str) -> dict:
    """
    Query Perf table (legacy Microsoft Monitoring Agent / MMA) for all
    performance metrics.

    MMA stores metrics using ObjectName/CounterName/InstanceName semantics.
    CPU uses _Total instance; disk and network use per-drive/per-NIC instances
    which are summed to get VM totals.

    Unit notes:
      CPU         — % (0-100), no conversion
      Memory      — Available MBytes (MiB) → bytes ×1,048,576
      Disk        — Reads/Writes per sec, no conversion
      Network     — Bytes/sec, no conversion
    """
    kql = f"""
let startT = datetime("{start_str}");
let endT   = datetime("{end_str}");
Perf
| where TimeGenerated between (startT .. endT)
| where Computer contains "{vm_name}"
| where (ObjectName == "Processor"          and CounterName == "% Processor Time"
            and InstanceName == "_Total")
     or (ObjectName == "Memory"             and CounterName == "Available MBytes")
     or (ObjectName == "LogicalDisk"        and CounterName in ("Disk Reads/sec",
                                                                "Disk Writes/sec")
            and InstanceName !in ("_Total", "HarddiskVolume1") and isnotempty(InstanceName))
     or (ObjectName == "Network Interface"  and CounterName in (
                "Bytes Received/sec",           // Windows MMA
                "Bytes Sent/sec",               // Windows MMA
                "Total Bytes Received/sec",     // Linux MMA (some agent versions)
                "Total Bytes Transmitted/sec"   // Linux MMA (some agent versions)
            )
            and InstanceName != "_Total")
| extend MetricName = case(
    ObjectName == "Processor"         and CounterName == "% Processor Time",  "Percentage CPU",
    ObjectName == "Memory"            and CounterName == "Available MBytes",   "Available Memory Bytes",
    ObjectName == "LogicalDisk"       and CounterName == "Disk Reads/sec",     "Disk Read Operations/Sec",
    ObjectName == "LogicalDisk"       and CounterName == "Disk Writes/sec",    "Disk Write Operations/Sec",
    ObjectName == "Network Interface" and CounterName in ("Bytes Received/sec", "Total Bytes Received/sec"),     "Network In Total",
    ObjectName == "Network Interface" and CounterName in ("Bytes Sent/sec", "Total Bytes Transmitted/sec"),  "Network Out Total",
    "")
| where isnotempty(MetricName)
// Convert MMA memory (MiB) → bytes
| extend Value = iff(MetricName == "Available Memory Bytes",
                     CounterValue * 1048576.0, CounterValue)
// Step 1: per-instance per-minute average
| summarize InstanceVal = avg(Value) by MetricName, Instance = InstanceName,
    bin(TimeGenerated, 1m)
// Step 2: sum instances per minute (disk C: + D:, NIC eth0 + eth1)
| summarize MinTotal = sum(InstanceVal) by MetricName, TimeGenerated
// Step 3: hourly avg and peak
| summarize AvgVal = avg(MinTotal), MaxVal = max(MinTotal)
    by MetricName, bin(TimeGenerated, 1h)
| order by TimeGenerated asc
"""
    return _run_la_metrics_kql(workspace_id, kql, la_token)


def _has_data(metrics_dict: dict) -> bool:
    """Return True if at least one metric in the dict has data points."""
    return any(
        len(v["Average"]) > 0
        for v in metrics_dict.values()
    )


def fetch_all_metrics_via_log_analytics(workspace_id: str,
                                         vm_name: str,
                                         start_dt: datetime.datetime,
                                         end_dt: datetime.datetime,
                                         la_token: str) -> tuple:
    """
    Fetch all performance metrics from Log Analytics using a strict priority cascade:

      Priority 1 — InsightsMetrics (AMA / new VM Insights, DCR-based)
        Used when the VM has the Azure Monitor Agent installed and VM Insights enabled.

      Priority 2 — Perf table (MMA / legacy Log Analytics agent)
        Fallback for VMs still using the classic agent.

    IMPORTANT: The two sources are NEVER mixed for the same VM.
    Mixing them would double-count metrics because:
      - Both agents poll the same counters (CPU, memory, disk, network)
      - The aggregation pipeline sums per-instance values
      - If both AMA and MMA run on the same VM, every metric would be doubled

    Returns:
        (metrics_dict, source_label)
        metrics_dict: {metric_name: {"Average": [...], "Maximum": [...]}}
        source_label: "LA-AMA" | "LA-MMA" | "none"
    """
    start_str = start_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
    end_str   = end_dt.strftime("%Y-%m-%dT%H:%M:%SZ")

    # ── Priority 1: AMA / InsightsMetrics ────────────────────────────────────
    try:
        ama_data = _fetch_metrics_ama(workspace_id, vm_name, start_str, end_str, la_token)
        if _has_data(ama_data):
            filled = sum(1 for v in ama_data.values() if v["Average"])
            print(f"[INFO]       Metrics source: LA-AMA (InsightsMetrics) — "
                  f"{filled}/6 metrics")
            return ama_data, "LA-AMA"
    except Exception as e:
        print(f"[INFO]       AMA query skipped: {e}")

    # ── Priority 2: MMA / Perf table ─────────────────────────────────────────
    try:
        mma_data = _fetch_metrics_mma(workspace_id, vm_name, start_str, end_str, la_token)
        if _has_data(mma_data):
            filled = sum(1 for v in mma_data.values() if v["Average"])
            print(f"[INFO]       Metrics source: LA-MMA (Perf table) — "
                  f"{filled}/6 metrics")
            return mma_data, "LA-MMA"
    except Exception as e:
        print(f"[INFO]       MMA query skipped: {e}")

    print(f"[INFO]       No LA data found for {vm_name} — will use ARM API")
    return {}, "none"



@dataclass
class NSGRule:
    """One security rule from a Network Security Group."""
    name: str
    priority: int
    direction: str          # "Inbound" or "Outbound"
    access: str             # "Allow" or "Deny"
    protocol: str           # "TCP", "UDP", "*", "ICMP"
    source_prefix: str      # "*", CIDR, or service tag e.g. "VirtualNetwork"
    source_port: str        # "*" or port/range
    dest_prefix: str        # "*", CIDR, or service tag
    dest_port: str          # "*", single port, range, or comma-joined list
    description: str = ""
    nsg_name: str = ""      # which NSG this rule belongs to
    nsg_level: str = ""     # "NIC" or "Subnet"
    # Risk assessment (filled by _assess_nsg_rule_risk)
    risk: str = "INFO"      # CRITICAL | HIGH | MEDIUM | LOW | INFO
    risk_reason: str = ""
    recommendation: str = ""


# ──────────────────────────────────────────────
# NSG Collection
# ──────────────────────────────────────────────

# Well-known sensitive ports and their labels
_SENSITIVE_PORTS = {
    "22":   "SSH",
    "3389": "RDP",
    "3306": "MySQL",
    "1433": "MSSQL",
    "1434": "MSSQL Browser",
    "5432": "PostgreSQL",
    "1521": "Oracle DB",
    "27017": "MongoDB",
    "27018": "MongoDB",
    "6379": "Redis",
    "5984": "CouchDB",
    "9200": "Elasticsearch",
    "9300": "Elasticsearch",
    "5900": "VNC",
    "23":   "Telnet",
    "21":   "FTP",
    "20":   "FTP Data",
    "445":  "SMB",
    "139":  "NetBIOS",
    "137":  "NetBIOS",
    "138":  "NetBIOS",
    "135":  "Windows RPC",
    "593":  "HTTP RPC",
    "4848": "GlassFish Admin",
    "2375": "Docker (unencrypted)",
    "2376": "Docker TLS",
    "6443": "Kubernetes API",
    "2379": "etcd",
    "2380": "etcd",
}

# Source prefixes that represent "any internet source"
_INTERNET_SOURCES = {"*", "0.0.0.0/0", "internet", "any"}

# Service tags that are still broad (not locked to org network)
_BROAD_TAGS = {"internet", "any", "*"}

# Ports that are intentionally exposed to the internet for public-facing applications.
# Opening these to "any" source is by design for web servers and load balancers.
# These ports are rated INFO when the source is the internet, not HIGH.
_PUBLIC_WEB_PORTS = {
    "80",    # HTTP
    "443",   # HTTPS
    "8080",  # HTTP alternate (common for app servers behind a load balancer)
    "8443",  # HTTPS alternate
    "8000",  # HTTP alternate (Django, etc.)
    "8001",  # HTTP alternate
    "8888",  # HTTP alternate (Jupyter, etc.)
}


def _is_internet_source(prefix: str) -> bool:
    """
    True only if the source prefix means the entire internet / any source.

    Handles:
      "*"           → True   (wildcard — any source)
      "0.0.0.0/0"  → True   (all IPv4)
      "Internet"    → True   (Azure service tag meaning public internet)
      "Any"         → True   (explicit any)
      "14.194.28.250, 118.185.0.1"  → False  (specific whitelisted IPs)
      "10.0.0.0/24"                 → False  (private CIDR)
      "VirtualNetwork"              → False  (Azure service tag)

    When multiple IPs are comma-joined (from sourceAddressPrefixes),
    ALL tokens would have to be wildcards to be considered internet-open,
    which is never true for a comma-separated list of real IPs.
    """
    raw = prefix.strip()
    # Multi-value (comma-separated) — never qualifies as a single internet wildcard
    if "," in raw:
        return False
    return raw.lower() in _INTERNET_SOURCES


def _is_public_web_port(dest_port: str) -> bool:
    """
    Return True if ALL destination ports in the rule are standard public web ports
    (HTTP, HTTPS, and common application-server alternates).

    Rationale: ports 80, 443, 8080, 8443 etc. are intentionally internet-facing
    on web servers and load balancers. An NSG rule opening these to any source
    is by design and should be rated INFO, not HIGH.

    Returns False if ANY port in the rule is NOT in _PUBLIC_WEB_PORTS,
    so a rule like "80,3389" is still rated HIGH/CRITICAL for the RDP port.
    """
    raw = dest_port.strip()
    if raw == "*":
        return False   # all-port rules handled separately as CRITICAL
    # Parse comma-separated ports/ranges
    segments = [s.strip() for s in raw.split(",") if s.strip()]
    for seg in segments:
        if "-" in seg:
            # A range e.g. "8000-8010" — only INFO if the entire range maps to
            # public web ports. Conservative: reject ranges unless they collapse
            # to a single known port.
            try:
                lo, hi = seg.split("-", 1)
                if lo == hi and lo in _PUBLIC_WEB_PORTS:
                    continue
            except ValueError:
                pass
            return False   # ranges are not automatically public-web
        if seg not in _PUBLIC_WEB_PORTS:
            return False
    return len(segments) > 0


def _normalise_port(raw) -> str:
    """Normalise a port value from the API (may be list or string)."""
    if isinstance(raw, list):
        return ", ".join(str(p) for p in raw)
    return str(raw) if raw else "*"


def _normalise_prefix(single: str, multi: list) -> str:
    """
    Resolve the source/destination address prefix from the Azure API response.

    Azure stores addresses in TWO mutually exclusive fields:
      sourceAddressPrefix       (str)  — used when there is exactly ONE value
                                         e.g. "*", "10.0.0.1", "VirtualNetwork"
      sourceAddressPrefixes     (list) — used when there are MULTIPLE values
                                         e.g. ["14.194.28.250", "118.185.0.0/16"]

    When the portal shows "14.194.28.250,118.185...." the rule is using the
    plural field. Reading only the singular field gives "" which we were
    defaulting to "*" — causing the report to show a whitelisted rule as CRITICAL.

    Returns a display string:
      Single value  → the value itself        e.g. "*" / "10.0.0.0/24"
      Multiple IPs  → comma-joined            e.g. "14.194.28.250, 118.185.0.1"
    """
    # Plural field (list) takes precedence when non-empty
    if multi and isinstance(multi, list):
        cleaned = [str(p).strip() for p in multi if str(p).strip()]
        if cleaned:
            return ", ".join(cleaned)
    # Fall back to singular field
    val = str(single).strip() if single else ""
    return val if val else "*"


def _port_hits_sensitive(dest_port: str) -> list:
    """
    Return a list of (port_str, service_name) tuples where the rule's
    destination port range covers a known sensitive port.
    Handles "*", single ports, and ranges like "3000-4000".
    """
    if dest_port.strip() == "*":
        return list(_SENSITIVE_PORTS.items())

    hits = []
    for segment in dest_port.replace(" ", "").split(","):
        segment = segment.strip()
        if "-" in segment:
            try:
                lo, hi = segment.split("-", 1)
                lo, hi = int(lo), int(hi)
                for p_str, svc in _SENSITIVE_PORTS.items():
                    if lo <= int(p_str) <= hi:
                        hits.append((p_str, svc))
            except ValueError:
                pass
        else:
            if segment in _SENSITIVE_PORTS:
                hits.append((segment, _SENSITIVE_PORTS[segment]))
    return hits


def _assess_nsg_rule_risk(rule: NSGRule) -> None:
    """
    Assign risk, risk_reason, and recommendation to a rule in-place.

    Risk matrix for INBOUND ALLOW rules:
      CRITICAL  — internet source AND (sensitive port OR any port)
      HIGH      — internet source AND specific non-sensitive port exposed
      MEDIUM    — broad CIDR (not full internet) AND sensitive port
      LOW       — broad CIDR AND any port, or unusual outbound allow to internet
      INFO      — everything else (deny rules, service-tag-restricted allows, outbound)
    """
    # Deny rules and outbound rules to non-internet are always INFO
    if rule.access.lower() == "deny":
        rule.risk = "INFO"
        rule.risk_reason = "Deny rule — no exposure"
        return

    internet_src = _is_internet_source(rule.source_prefix)
    inbound      = rule.direction.lower() == "inbound"

    if not inbound:
        # Outbound allow to internet — LOW if any-port, else INFO
        if internet_src and rule.dest_port.strip() == "*":
            rule.risk = "LOW"
            rule.risk_reason = "Outbound Allow to internet on all ports"
            rule.recommendation = (
                f"Review outbound rule '{rule.name}' on {rule.nsg_name}. "
                f"Restrict destination port range to required protocols only."
            )
        else:
            rule.risk = "INFO"
            rule.risk_reason = "Outbound allow — low risk"
        return

    # ── Inbound Allow rules below ─────────────────────────────────────
    any_port   = rule.dest_port.strip() == "*"
    sens_hits  = _port_hits_sensitive(rule.dest_port)

    if internet_src:
        if any_port:
            rule.risk = "CRITICAL"
            rule.risk_reason = "Inbound Allow from ANY source on ALL ports"
            rule.recommendation = (
                f"URGENT: Rule '{rule.name}' on {rule.nsg_name} allows unrestricted "
                f"inbound access from the internet on all ports. "
                f"Remove or replace with rules specifying exact source IPs/CIDRs "
                f"and required destination ports only."
            )
        elif sens_hits:
            svc_list = ", ".join(f"{p} ({s})" for p, s in sens_hits[:3])
            rule.risk = "CRITICAL"
            rule.risk_reason = f"Inbound Allow from internet to sensitive port(s): {svc_list}"
            # Generate port-specific guidance
            port_recs = []
            for p_str, svc in sens_hits:
                if svc == "RDP":
                    port_recs.append(
                        "Disable public RDP (3389). Enable Microsoft Defender for Cloud "
                        "Just-in-Time (JIT) VM access, or use Azure Bastion."
                    )
                elif svc == "SSH":
                    port_recs.append(
                        "Disable public SSH (22). Use Azure Bastion, JIT access, "
                        "or restrict source to a specific management IP/CIDR."
                    )
                elif svc in ("MySQL", "MSSQL", "MSSQL Browser", "PostgreSQL",
                              "Oracle DB", "MongoDB", "Redis", "CouchDB",
                              "Elasticsearch"):
                    port_recs.append(
                        f"Database port {p_str} ({svc}) must not be exposed to the internet. "
                        f"Remove the inbound rule and restrict access to the application "
                        f"subnet CIDR or a private endpoint."
                    )
                elif svc in ("SMB", "NetBIOS", "Windows RPC"):
                    port_recs.append(
                        f"Windows file sharing port {p_str} ({svc}) is exposed to the internet. "
                        f"This is a critical vulnerability — block immediately."
                    )
                elif svc in ("Telnet", "FTP", "FTP Data"):
                    port_recs.append(
                        f"Unencrypted protocol port {p_str} ({svc}) is exposed. "
                        f"Remove this rule. Replace with SSH/SFTP where applicable."
                    )
                elif svc in ("Docker (unencrypted)", "Docker TLS",
                              "Kubernetes API", "etcd"):
                    port_recs.append(
                        f"Container orchestration port {p_str} ({svc}) must never be "
                        f"exposed to the internet. Restrict to VNet CIDR only."
                    )
                else:
                    port_recs.append(
                        f"Restrict port {p_str} ({svc}) source from '*' to a specific "
                        f"IP address or CIDR. Apply the principle of least privilege."
                    )
            rule.recommendation = " | ".join(dict.fromkeys(port_recs))  # dedupe
        elif _is_public_web_port(rule.dest_port):
            # Port is intentionally public-facing (HTTP/HTTPS/common web ports).
            # Opening these to any internet source is expected for web servers.
            rule.risk = "INFO"
            rule.risk_reason = (
                f"Inbound Allow from internet to public web port {rule.dest_port} "
                f"— expected for public-facing application hosting"
            )
        else:
            rule.risk = "HIGH"
            rule.risk_reason = (
                f"Inbound Allow from internet to port {rule.dest_port}"
            )
            rule.recommendation = (
                f"Rule '{rule.name}' on {rule.nsg_name} allows inbound access from "
                f"the internet to port {rule.dest_port}. "
                f"Whitelist specific source IP addresses or CIDRs instead of '*'."
            )
    else:
        # Non-internet source — assess broad CIDRs
        # Heuristic: a /8 or /16 prefix is considered broad
        src = rule.source_prefix
        is_broad_cidr = False
        try:
            if "/" in src:
                prefix_len = int(src.split("/")[1])
                is_broad_cidr = prefix_len <= 16
        except (ValueError, IndexError):
            pass

        if is_broad_cidr and any_port:
            # any_port is a superset — check it first (it's the broader risk)
            rule.risk = "LOW"
            rule.risk_reason = f"Inbound Allow from broad CIDR {src} on all ports"
            rule.recommendation = (
                f"Rule '{rule.name}' on {rule.nsg_name} allows access from a broad "
                f"network range ({src}) on all ports. "
                f"Restrict to specific required destination port(s)."
            )
        elif is_broad_cidr and sens_hits:
            svc_list = ", ".join(f"{p} ({s})" for p, s in sens_hits[:3])
            rule.risk = "MEDIUM"
            rule.risk_reason = (
                f"Inbound Allow from broad CIDR {src} to sensitive port(s): {svc_list}"
            )
            rule.recommendation = (
                f"Rule '{rule.name}' on {rule.nsg_name} allows access from a broad "
                f"CIDR ({src}) to sensitive port(s) {svc_list}. "
                f"Tighten the source CIDR to the minimum required subnet."
            )
        else:
            rule.risk = "INFO"
            rule.risk_reason = "Source is restricted to a specific subnet or service tag"


def _get_nsg_rules_from_id(nsg_id: str, arm_token: str,
                             nsg_level: str) -> tuple:
    """
    Fetch security rules from a single NSG resource ID.
    Returns (nsg_name, [NSGRule, ...]).
    Only returns custom rules (not Azure default rules — those are always INFO).
    """
    url = f"https://management.azure.com{nsg_id}?api-version=2023-11-01"
    resp = requests.get(url, headers={"Authorization": f"Bearer {arm_token}"}, timeout=30)
    resp.raise_for_status()
    nsg_body = resp.json()
    nsg_name = nsg_body.get("name", nsg_id.split("/")[-1])

    rules = []
    for r in nsg_body.get("properties", {}).get("securityRules", []):
        props = r.get("properties", {})
        # Normalise destination port: prefer portRanges list, fall back to single range
        dest_port = _normalise_port(
            props.get("destinationPortRanges") or props.get("destinationPortRange") or "*"
        )
        src_port = _normalise_port(
            props.get("sourcePortRanges") or props.get("sourcePortRange") or "*"
        )
        source_prefix = _normalise_prefix(
            props.get("sourceAddressPrefix", ""),
            props.get("sourceAddressPrefixes", []),
        )
        dest_prefix = _normalise_prefix(
            props.get("destinationAddressPrefix", ""),
            props.get("destinationAddressPrefixes", []),
        )
        rule = NSGRule(
            name=r.get("name", "unnamed"),
            priority=props.get("priority", 0),
            direction=props.get("direction", ""),
            access=props.get("access", ""),
            protocol=props.get("protocol", "*"),
            source_prefix=source_prefix,
            source_port=src_port,
            dest_prefix=dest_prefix,
            dest_port=dest_port,
            description=props.get("description", ""),
            nsg_name=nsg_name,
            nsg_level=nsg_level,
        )
        _assess_nsg_rule_risk(rule)
        rules.append(rule)

    # Sort by priority ascending so report reads naturally
    rules.sort(key=lambda x: x.priority)
    return nsg_name, rules


def fetch_vm_nsg_rules(vm_resource_id: str, arm_token: str) -> tuple:
    """
    Collect all NSG rules that apply to a VM by traversing:
      VM → NIC(s) → NIC-level NSG (if any)
                  → Subnet → Subnet-level NSG (if any)

    Both NIC-level and subnet-level NSGs are returned and labelled clearly.
    Only custom rules are included (default Azure rules are excluded to keep
    the report focused on administrator-controlled configuration).

    Permission: Reader (delegated via Lighthouse) — fully read-only.

    Returns:
        nsg_names: list of display strings e.g. ["NIC: lfg-web-nsg", "Subnet: sub-nsg"]
        all_rules: combined list of NSGRule objects, sorted by (level, priority)
    """
    nsg_names = []
    all_rules = []
    seen_nsg_ids = set()   # avoid double-collecting if NIC and subnet share the same NSG

    try:
        # 1. Get VM details → NIC IDs
        vm_url = f"https://management.azure.com{vm_resource_id}?api-version=2023-03-01"
        vm_resp = requests.get(
            vm_url, headers={"Authorization": f"Bearer {arm_token}"}, timeout=30
        )
        vm_resp.raise_for_status()
        vm_body = vm_resp.json()

        nic_refs = (
            vm_body.get("properties", {})
                   .get("networkProfile", {})
                   .get("networkInterfaces", [])
        )

        for nic_ref in nic_refs:
            nic_id = nic_ref.get("id", "")
            if not nic_id:
                continue

            try:
                # 2. Get NIC details
                nic_url = f"https://management.azure.com{nic_id}?api-version=2023-11-01"
                nic_resp = requests.get(
                    nic_url, headers={"Authorization": f"Bearer {arm_token}"}, timeout=30
                )
                nic_resp.raise_for_status()
                nic_body = nic_resp.json()
                nic_props = nic_body.get("properties", {})

                # 3. NIC-level NSG
                nic_nsg = nic_props.get("networkSecurityGroup", {})
                nic_nsg_id = nic_nsg.get("id", "") if nic_nsg else ""
                if nic_nsg_id and nic_nsg_id not in seen_nsg_ids:
                    seen_nsg_ids.add(nic_nsg_id)
                    nsg_name, rules = _get_nsg_rules_from_id(nic_nsg_id, arm_token, "NIC")
                    nsg_names.append(f"NIC: {nsg_name}")
                    all_rules.extend(rules)

                # 4. Subnet-level NSG (from first ipConfiguration)
                for ip_cfg in nic_props.get("ipConfigurations", []):
                    subnet_id = (
                        ip_cfg.get("properties", {})
                              .get("subnet", {})
                              .get("id", "")
                    )
                    if not subnet_id:
                        continue
                    try:
                        sub_url = (
                            f"https://management.azure.com{subnet_id}"
                            f"?api-version=2023-11-01"
                        )
                        sub_resp = requests.get(
                            sub_url,
                            headers={"Authorization": f"Bearer {arm_token}"},
                            timeout=30,
                        )
                        sub_resp.raise_for_status()
                        sub_nsg = (
                            sub_resp.json()
                                    .get("properties", {})
                                    .get("networkSecurityGroup", {})
                        )
                        sub_nsg_id = sub_nsg.get("id", "") if sub_nsg else ""
                        if sub_nsg_id and sub_nsg_id not in seen_nsg_ids:
                            seen_nsg_ids.add(sub_nsg_id)
                            nsg_name, rules = _get_nsg_rules_from_id(
                                sub_nsg_id, arm_token, "Subnet"
                            )
                            nsg_names.append(f"Subnet: {nsg_name}")
                            all_rules.extend(rules)
                    except Exception as e:
                        print(f"[WARN]       Subnet NSG fetch failed: {e}")
                    break   # one subnet per NIC is sufficient

            except Exception as e:
                print(f"[WARN]       NIC {nic_id.split('/')[-1]} NSG fetch failed: {e}")

    except Exception as e:
        print(f"[WARN]       VM NIC list fetch failed: {e}")

    return nsg_names, all_rules


def _cpu_max_consecutive_breach(cpu_series: list, threshold: float) -> int:
    """
    Scan the hourly CPU series and return the longest run of consecutive
    data points that all exceed the threshold.
    Used to distinguish isolated spikes (run=1) from sustained pressure (run>3).
    """
    max_run = 0
    current = 0
    for _, val in cpu_series:
        if val >= threshold:
            current += 1
            max_run = max(max_run, current)
        else:
            current = 0
    return max_run


# SKU upgrade suggestions — maps (vcpus, mem_gib) to a suggested next tier
# Covers the most common D/E/B series VMs used in LFG environments.
_SKU_UPGRADES = {
    # Standard_B series
    ("Standard_B2s",    2,  4):  "Standard_B4ms (4 vCPUs, 16 GiB)",
    ("Standard_B2ms",   2,  8):  "Standard_B4ms (4 vCPUs, 16 GiB)",
    ("Standard_B4ms",   4, 16):  "Standard_B8ms (8 vCPUs, 32 GiB)",
    ("Standard_B8ms",   8, 32):  "Standard_D8s_v3 (8 vCPUs, 32 GiB)",
    # Standard_D v3 series
    ("Standard_D2_v3",  2,  8):  "Standard_D4_v3 (4 vCPUs, 16 GiB)",
    ("Standard_D4_v3",  4, 16):  "Standard_D8_v3 (8 vCPUs, 32 GiB)",
    ("Standard_D8_v3",  8, 32):  "Standard_D16_v3 (16 vCPUs, 64 GiB)",
    ("Standard_D16_v3",16, 64):  "Standard_D32_v3 (32 vCPUs, 128 GiB)",
    ("Standard_D32_v3",32,128):  "Standard_D48_v3 (48 vCPUs, 192 GiB)",
    # Standard_Ds v3 series
    ("Standard_D2s_v3", 2,  8):  "Standard_D4s_v3 (4 vCPUs, 16 GiB)",
    ("Standard_D4s_v3", 4, 16):  "Standard_D8s_v3 (8 vCPUs, 32 GiB)",
    ("Standard_D8s_v3", 8, 32):  "Standard_D16s_v3 (16 vCPUs, 64 GiB)",
    ("Standard_D16s_v3",16, 64): "Standard_D32s_v3 (32 vCPUs, 128 GiB)",
    ("Standard_D32s_v3",32,128): "Standard_D48s_v3 (48 vCPUs, 192 GiB)",
    # Standard_E v3 series
    ("Standard_E2_v3",  2, 16):  "Standard_E4_v3 (4 vCPUs, 32 GiB)",
    ("Standard_E4_v3",  4, 32):  "Standard_E8_v3 (8 vCPUs, 64 GiB)",
    ("Standard_E8_v3",  8, 64):  "Standard_E16_v3 (16 vCPUs, 128 GiB)",
    ("Standard_E16_v3",16,128):  "Standard_E32_v3 (32 vCPUs, 256 GiB)",
}


def _suggest_sku_upgrade(sku: str, vcpus: int, mem_gib: float) -> str:
    """Return a suggested next-tier SKU string, or empty string if unknown."""
    # Try exact match first
    key = (sku, vcpus, int(mem_gib))
    if key in _SKU_UPGRADES:
        return _SKU_UPGRADES[key]
    # Fall back to vcpu/memory match (handles minor SKU name variations)
    for (_, v, m), suggestion in _SKU_UPGRADES.items():
        if v == vcpus and m == int(mem_gib):
            return suggestion
    return ""


def _cpu_recommendations(vm_name: str, sku: str, vcpus: int, mem_gib: float,
                          breach_count: int, breach_day_count: int,
                          breach_peak: float, breach_avg: float,
                          overall_avg: float, max_consecutive: int,
                          breach_dates: list, threshold: float) -> list:
    """
    Generate context-aware, tiered CPU recommendations.

    Logic:
      CRITICAL (peak ≥ 95%):
        → Always urgent right-sizing recommendation with specific SKU suggestion
        → If sustained (max_consecutive > 3): add immediate scaling advisory

      HIGH BREACH COUNT (≥ 5 days or ≥ 20 hourly breaches):
        → Persistent pressure — recommend right-sizing + Azure Monitor alert
        → Include SKU suggestion if one is known
        → If overall avg also high (> 60%): note VM appears consistently under-resourced

      MODERATE BREACH COUNT (breaches on 2-4 distinct days):
        → Semi-regular pattern — recommend workload review for specific dates
        → Suggest scheduled task / batch job investigation

      LOW BREACH COUNT (1 day or very few breaches):
        → Likely isolated spike — recommend log review for specific date(s)
        → Suggest Azure Monitor alert rule to catch future occurrences

      SUSTAINED (max_consecutive > 3 hours, any severity):
        → Add explicit recommendation: prolonged CPU saturation is more harmful
          than brief spikes — process-level investigation needed
    """
    recs = []
    upgrade = _suggest_sku_upgrade(sku, vcpus, mem_gib)

    # ── Near-saturation (peak ≥ 95%) ─────────────────────────────────────────
    if breach_peak >= 95:
        upgrade_text = (
            f" Consider upgrading to {upgrade}."
            if upgrade else
            " Evaluate a larger SKU in the same series or move to a compute-optimised tier."
        )
        recs.append(
            f"URGENT — {vm_name} reached {breach_peak}% CPU (near-saturation). "
            f"At this level the OS scheduler queues threads, causing latency spikes "
            f"and potential request timeouts.{upgrade_text}"
        )

    # ── Sustained pressure (consecutive hours above threshold) ────────────────
    if max_consecutive >= 4:
        recs.append(
            f"{vm_name} sustained CPU above {threshold}% for {max_consecutive} "
            f"consecutive hours. Sustained high CPU is more harmful than brief "
            f"spikes — identify the process(es) responsible using Azure Monitor "
            f"Metrics Explorer (split by Process) or connect via Azure Bastion and "
            f"run 'top' / Task Manager during the next occurrence."
        )
    elif max_consecutive >= 2:
        recs.append(
            f"{vm_name} had CPU above {threshold}% for {max_consecutive} consecutive "
            f"hours. Investigate whether a scheduled job or batch process is running "
            f"without CPU throttling or concurrency limits."
        )

    # ── Persistent pressure (many breach days) ────────────────────────────────
    if breach_day_count >= 5:
        upgrade_note = f" Recommended next tier: {upgrade}." if upgrade else ""
        recs.append(
            f"{vm_name} exceeded the {threshold}% CPU threshold on {breach_day_count} "
            f"separate days (overall average: {overall_avg}%). This is a persistent "
            f"pattern, not an isolated event.{upgrade_note} "
            f"Actions: (1) Profile the application to identify CPU-heavy code paths. "
            f"(2) Review auto-scaling rules if this VM is part of an availability set. "
            f"(3) Create an Azure Monitor Alert rule on Percentage CPU > {threshold}% "
            f"to receive real-time notifications."
        )
    elif breach_day_count >= 2:
        date_list = ", ".join(breach_dates[:5])
        if len(breach_dates) > 5:
            date_list += f" (+ {len(breach_dates)-5} more)"
        recs.append(
            f"{vm_name} had CPU breaches on {breach_day_count} days: {date_list}. "
            f"Review application logs and Azure Activity Log for these dates to "
            f"identify deployments, configuration changes, or batch jobs that "
            f"triggered the spikes. Set up an Azure Monitor Alert to detect future "
            f"breaches before they impact users."
        )
    else:
        # Single day or very few isolated breaches
        date_str = breach_dates[0] if breach_dates else "the reported date"
        recs.append(
            f"{vm_name} had an isolated CPU spike on {date_str} "
            f"(peak: {breach_peak}%). "
            f"Review application logs, deployment history, and scheduled tasks "
            f"for that date. If this recurs, consider creating an Azure Monitor "
            f"Alert on Percentage CPU > {threshold}% with a 5-minute evaluation window."
        )

    # ── High overall average alongside breaches ───────────────────────────────
    if overall_avg >= 60 and breach_day_count >= 3:
        recs.append(
            f"{vm_name} overall monthly average CPU is {overall_avg}%. "
            f"Combined with threshold breaches on {breach_day_count} days, "
            f"this indicates the VM is consistently under-resourced for its "
            f"current workload. A right-sizing review is recommended."
        )

    return recs


# ──────────────────────────────────────────────
# Analysis & Findings
# ──────────────────────────────────────────────

def analyze_vm_metrics(vm: VMMetrics, config: ReportConfig) -> dict:
    """
    Evaluate all metrics against configured thresholds.
    Populates vm computed fields and returns a structured findings dict.
    """
    findings = {
        "vm_name":             vm.vm_name,
        "sku":                 vm.sku,
        "subscription_id":     vm.subscription_id,
        "issues":              [],
        "recommendations":     [],
        "nsg_issues":          [],   # list of NSGRule objects with risk != INFO
        "cpu_breach_summary":  None, # populated only if CPU threshold was breached
        "status":              "NORMAL",
        "cpu_max":             0.0,  # to be populated
    }

    # ── CPU ──────────────────────────────────────────────────────────
    if vm.cpu_percent:
        values     = [v for _, v in vm.cpu_percent]
        vm.cpu_avg = round(sum(values) / len(values), 2)
        
        # Calculate overall vm.cpu_max
        if vm.cpu_percent_max:
            values_max = [v for _, v in vm.cpu_percent_max]
            vm.cpu_max = round(max(values_max), 2)
        else:
            vm.cpu_max = round(max(values), 2)
            
        findings["cpu_max"] = vm.cpu_max

        # Collect breaches with full ISO timestamp (not just date) for hour-level detail
        # Check breaches against the Max series since both Avg and Max lines are drawn,
        # and users expect alerts when spikes cross the threshold.
        series_to_check = vm.cpu_percent_max if vm.cpu_percent_max else vm.cpu_percent
        
        for ts, val in series_to_check:
            if val >= config.cpu_alert_threshold:
                vm.cpu_threshold_breaches.append((ts, round(val, 2)))

        if vm.cpu_threshold_breaches:
            # ── Derived statistics ────────────────────────────────────────
            breach_count   = len(vm.cpu_threshold_breaches)
            breach_values  = [v for _, v in vm.cpu_threshold_breaches]
            breach_peak    = round(max(breach_values), 2)
            breach_avg     = round(sum(breach_values) / breach_count, 2)

            # Unique calendar days that had at least one breach
            breach_dates   = sorted(set(ts[:10] for ts, _ in vm.cpu_threshold_breaches))
            breach_day_count = len(breach_dates)

            # Detect consecutive-hour runs (sustained vs spike)
            # A "run" = consecutive hourly data points all above threshold
            max_consecutive = _cpu_max_consecutive_breach(series_to_check,
                                                          config.cpu_alert_threshold)

            # Average CPU across the full period
            overall_avg = vm.cpu_avg

            # ── Severity classification ───────────────────────────────────
            if breach_peak >= 95:
                findings["status"] = "CRITICAL"
                severity_label     = "CRITICAL"
            else:
                findings["status"] = "WARNING"
                severity_label     = "WARNING"

            # ── Issue text ───────────────────────────────────────────────
            first_5 = "; ".join(
                f"{ts[:10]} {ts[11:16]} ({v}%)"
                for ts, v in vm.cpu_threshold_breaches[:5]
            )
            suffix = f" — showing first 5 of {breach_count}" if breach_count > 5 else ""
            findings["issues"].append(
                f"[{severity_label}] CPU exceeded {config.cpu_alert_threshold}% threshold "
                f"{breach_count} time(s) across {breach_day_count} day(s). "
                f"Peak: {breach_peak}%, Avg during breach: {breach_avg}%. "
                f"Breach timestamps: {first_5}{suffix}"
            )

            # ── Context-aware recommendations ────────────────────────────
            recs = _cpu_recommendations(
                vm_name=vm.vm_name,
                sku=vm.sku,
                vcpus=vm.vcpus,
                mem_gib=vm.memory_gib,
                breach_count=breach_count,
                breach_day_count=breach_day_count,
                breach_peak=breach_peak,
                breach_avg=breach_avg,
                overall_avg=overall_avg,
                max_consecutive=max_consecutive,
                breach_dates=breach_dates,
                threshold=config.cpu_alert_threshold,
            )
            findings["recommendations"].extend(recs)

            # Store structured breach summary for the report table
            findings["cpu_breach_summary"] = {
                "breach_count":    breach_count,
                "breach_day_count": breach_day_count,
                "breach_peak":     breach_peak,
                "breach_avg":      breach_avg,
                "overall_avg":     overall_avg,
                "max_consecutive": max_consecutive,
                "breach_dates":    breach_dates,
                "threshold":       config.cpu_alert_threshold,
                "severity":        severity_label,
            }

    # ── Memory ───────────────────────────────────────────────────────
    if vm.available_memory_bytes:
        values_gb             = [v / (1024 ** 3) for _, v in vm.available_memory_bytes]
        vm.memory_min_available_gb = round(min(values_gb), 2)
        free_pct = (
            (vm.memory_min_available_gb / vm.memory_gib * 100)
            if vm.memory_gib else 0
        )
        if free_pct < config.memory_min_threshold:
            findings["status"] = (
                "WARNING" if findings["status"] == "NORMAL" else findings["status"]
            )
            findings["issues"].append(
                f"Available memory dropped to {vm.memory_min_available_gb:.1f} GiB "
                f"({free_pct:.1f}% of {vm.memory_gib} GiB total)"
            )
            findings["recommendations"].append(
                f"Memory pressure detected on {vm.vm_name}. "
                f"Review memory-intensive processes or upgrade VM memory tier."
            )

    # ── Disk ─────────────────────────────────────────────────────────
    for drive, used_pct in vm.disk_utilization.items():
        if used_pct >= config.disk_util_threshold:
            findings["status"] = (
                "WARNING" if findings["status"] == "NORMAL" else findings["status"]
            )
            findings["issues"].append(
                f"Drive {drive} at {used_pct}% utilisation "
                f"(threshold: {config.disk_util_threshold}%)"
            )
            findings["recommendations"].append(
                f"Disk on {vm.vm_name} drive {drive} is approaching capacity. "
                f"Clean up old logs/data or expand disk size."
            )

    # ── NSG Security ─────────────────────────────────────────────────
    critical_rules = [r for r in vm.nsg_rules if r.risk == "CRITICAL"]
    high_rules     = [r for r in vm.nsg_rules if r.risk == "HIGH"]
    medium_rules   = [r for r in vm.nsg_rules if r.risk == "MEDIUM"]

    risky_rules = critical_rules + high_rules + medium_rules
    if critical_rules:
        findings["status"] = "CRITICAL"
    elif high_rules and findings["status"] != "CRITICAL":
        findings["status"] = "WARNING"
    elif medium_rules and findings["status"] == "NORMAL":
        findings["status"] = "WARNING"

    for rule in risky_rules:
        findings["nsg_issues"].append(rule)
        if rule.recommendation:
            findings["recommendations"].append(rule.recommendation)
        findings["issues"].append(
            f"NSG {rule.nsg_name} rule '{rule.name}' [{rule.risk}]: {rule.risk_reason}"
        )

    vm.has_alerts = findings["status"] != "NORMAL"
    return findings


# ──────────────────────────────────────────────
# Main Orchestrator — Multi-Tenant via Lighthouse
# ──────────────────────────────────────────────

def collect_all_tenants(config: ReportConfig,
                         arm_token: str,
                         la_token: str) -> dict:
    """
    Collect VM metrics for ALL subscriptions in config.target_subscription_ids.
    Uses Lighthouse: single ARM token, ARM resolves delegation per resource.

    Returns:
        {subscription_id: ([VMMetrics, ...], [findings_dict, ...])}
    """
    start_dt = config.report_start
    end_dt   = config.report_end

    print(f"[INFO] Report period: "
          f"{start_dt.strftime('%d %b %Y')} to {end_dt.strftime('%d %b %Y')}")

    # Only enumerate delegated subscriptions that match the explicit target list
    subscriptions = list_lighthouse_subscriptions(
        arm_token, allowed_ids=config.target_subscription_ids
    )

    results = {}

    for sub in subscriptions:
        sid  = sub.subscription_id
        name = sub.subscription_name
        print(f"\n[INFO] ── Subscription: {name} ({sid})")

        # VM discovery: use RG filter if configured, else subscription-wide
        filter_rgs = config.target_resource_groups.get(sid)
        if filter_rgs:
            vms_raw = []
            for rg in filter_rgs:
                rg_vms = list_vms_in_resource_group(sid, rg, arm_token)
                vms_raw.extend(rg_vms)
                print(f"[INFO]   RG '{rg}': {len(rg_vms)} VM(s)")
        else:
            vms_raw = list_all_vms_in_subscription(sid, arm_token)
            print(f"[INFO]   All RGs: {len(vms_raw)} VM(s)")

        if not vms_raw:
            print(f"[WARN]   No VMs found — skipping subscription")
            continue

        workspace_id     = config.log_analytics_workspace_ids.get(sid)
        all_vm_metrics   = []
        all_findings     = []

        for vm_info in vms_raw:
            vm_name = vm_info["name"]
            print(f"[INFO]     → {vm_name}")

            size_info = get_vm_size_details(
                sid, vm_info["location"], vm_info["vm_size"], arm_token
            )
            vm = VMMetrics(
                vm_name=vm_name,
                sku=vm_info["vm_size"],
                vcpus=size_info["vcpus"],
                memory_gib=round(size_info["memory_mb"] / 1024, 1),
                resource_id=vm_info["resource_id"],
                location=vm_info["location"],
                subscription_id=sid,
                resource_group=vm_info["resource_group"],
                tenant_id=sub.tenant_id,
            )

            # ── Metrics collection: LA priority, ARM fallback ────────────────
            # Priority order:
            #   1. Log Analytics InsightsMetrics (AMA) — highest accuracy
            #   2. Log Analytics Perf table (MMA)      — legacy agent
            #   3. Azure Monitor REST API              — no LA workspace / not enrolled
            #
            # Sources are NEVER mixed for the same VM — prevents double-counting.

            la_metrics  = {}
            la_source   = "none"
            arm_metrics = set()   # track which metrics came from ARM for the label

            if workspace_id:
                try:
                    la_metrics, la_source = fetch_all_metrics_via_log_analytics(
                        workspace_id, vm_name, start_dt, end_dt, la_token
                    )
                except Exception as e:
                    print(f"[WARN]       LA metrics query failed: {e}")
                    la_source = "none"

            metric_map = {
                # (avg_attr, max_attr): (azure_metric_name, arm_aggregation)
                ("cpu_percent",            "cpu_percent_max"):
                    ("Percentage CPU",            "Average"),
                ("available_memory_bytes", "available_memory_bytes_max"):
                    ("Available Memory Bytes",    "Average"),
                ("disk_read_iops",         "disk_read_iops_max"):
                    ("Disk Read Operations/Sec",  "Average"),
                ("disk_write_iops",        "disk_write_iops_max"):
                    ("Disk Write Operations/Sec", "Average"),
                ("network_bytes_sent",     "network_bytes_sent_max"):
                    ("Network Out Total",         "Total"),
                ("network_bytes_received", "network_bytes_received_max"):
                    ("Network In Total",          "Total"),
            }

            for (avg_attr, max_attr), (mname, agg) in metric_map.items():
                if la_metrics and la_metrics.get(mname, {}).get("Average"):
                    # LA data: already in correct units (bytes/sec for network,
                    # bytes for memory, % for CPU, ops/sec for disk)
                    setattr(vm, avg_attr, la_metrics[mname]["Average"])
                    setattr(vm, max_attr, la_metrics[mname]["Maximum"])
                else:
                    # ARM API fallback
                    try:
                        result = fetch_metric(
                            vm.resource_id, mname, start_dt, end_dt, arm_token, agg
                        )
                        if mname in ("Network In Total", "Network Out Total"):
                            # ARM Total-type metric conversions:
                            #   result["Average"] = dp["total"] = total bytes in 1 hour
                            #     → ÷ 3600 = average bytes/sec  (matches portal Avg KB/s)
                            #
                            #   result["Maximum"] = dp["maximum"] = peak bytes in 1 minute
                            #     → ÷ 60   = peak bytes/sec      (matches portal Max MB/s)
                            #
                            # Both the portal Average and Maximum are displayed as
                            # bytes/sec rates (KB/s, MB/s). These conversions match.
                            avg_bps = [(ts, v / 3600) for ts, v in result["Average"]]
                            max_bps = [(ts, v / 60)   for ts, v in result["Maximum"]]
                            setattr(vm, avg_attr, avg_bps)
                            setattr(vm, max_attr, max_bps)
                        else:
                            setattr(vm, avg_attr, result["Average"])
                            setattr(vm, max_attr, result["Maximum"])
                        arm_metrics.add(mname)
                    except Exception as e:
                        print(f"[WARN]       {mname} (ARM API): {e}")

            # Set the definitive metric_source label for the report
            if la_source in ("LA-AMA", "LA-MMA"):
                # Some metrics may have still fallen back to ARM (missing counters)
                if arm_metrics:
                    vm.metric_source = f"{la_source} + ARM({len(arm_metrics)})"
                else:
                    vm.metric_source = la_source
            elif arm_metrics:
                vm.metric_source = "ARM"
            else:
                vm.metric_source = "none"

            print(f"[INFO]       Metric source: {vm.metric_source}")

            # Disk utilisation via Log Analytics (if workspace configured)
            if workspace_id:
                try:
                    result = fetch_disk_utilization_via_log_analytics(
                        workspace_id, vm_name, start_dt, end_dt, la_token
                    )
                    vm.disk_utilization = result
                    if result:
                        # Source tag is set inside fetch_disk_utilization_via_log_analytics
                        # via the print lines; derive it here for the report
                        vm.disk_source = "found"   # generate_report checks disk_utilization dict
                    else:
                        vm.disk_source = "no_data"
                except Exception as e:
                    print(f"[WARN]       Disk util (Log Analytics): {e}")
                    vm.disk_source = "no_data"
            else:
                vm.disk_source = "no_workspace"   # workspace ID not in CLIENT_CONFIGS_JSON

            # NSG rules (read-only Network API — no write access)
            try:
                vm.nsg_names, vm.nsg_rules = fetch_vm_nsg_rules(
                    vm.resource_id, arm_token
                )
                risky = sum(1 for r in vm.nsg_rules if r.risk not in ("INFO", "LOW"))
                total = len(vm.nsg_rules)
                flag  = f"  ⚠ {risky} risky rule(s)" if risky else ""
                print(f"[INFO]       NSG: {len(vm.nsg_names)} NSG(s), "
                      f"{total} rule(s){flag}")
            except Exception as e:
                print(f"[WARN]       NSG collection failed: {e}")

            findings = analyze_vm_metrics(vm, config)
            all_vm_metrics.append(vm)
            all_findings.append(findings)

        results[sid] = (all_vm_metrics, all_findings)
        print(f"[INFO]   Subscription complete: "
              f"{len(all_vm_metrics)} VMs, "
              f"{sum(1 for f in all_findings if f['status'] != 'NORMAL')} alerts")

    return results
