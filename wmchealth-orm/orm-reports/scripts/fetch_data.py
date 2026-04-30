"""
fetch_data.py
-------------
Authenticates with the Press Ganey API and pulls all review + engagement
data needed to build the WMCHealth CTL monthly report.

Environment variables required:
  PG_APP_ID      - Press Ganey Application ID
  PG_APP_SECRET  - Press Ganey Application Secret

Usage:
  python orm-reports/scripts/fetch_data.py --month 2026-03
  python orm-reports/scripts/fetch_data.py   # defaults to previous calendar month
"""

import os
import json
import argparse
import requests
from datetime import datetime, timezone
from dateutil.relativedelta import relativedelta


# ── API ENDPOINTS ─────────────────────────────────────────────────────────────
BASE_URL       = "https://api1.consumerism.pressganey.com"
TOKEN_URL      = f"{BASE_URL}/api/service/v1/token/create"
REVIEWS_URL    = f"{BASE_URL}/api/reputation/v1/reviews"
SOURCES_URL    = f"{BASE_URL}/api/reputation/v1/sources"
ENGAGEMENT_URL = f"{BASE_URL}/api/service/v1/engagements/statistics"


# ── AUTHENTICATION ────────────────────────────────────────────────────────────
def get_access_token(app_id: str, app_secret: str) -> str:
    print("🔑 Authenticating with Press Ganey...")
    response = requests.post(
        TOKEN_URL,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        data={"appId": app_id, "appSecret": app_secret},
        timeout=30,
    )
    response.raise_for_status()
    data = response.json()
    if data.get("status", {}).get("code") != 200:
        raise ValueError(f"Auth failed: {data.get('status', {}).get('message')}")
    print(f"✅ Token obtained. Expires: {data.get('expiresIn')}")
    return data["accessToken"]


# ── SOURCES ───────────────────────────────────────────────────────────────────
def fetch_sources(token: str) -> list:
    print("📋 Fetching sources...")
    response = requests.get(SOURCES_URL, headers={"accessToken": token}, timeout=30)
    response.raise_for_status()
    sources = response.json()
    print(f"   Found {len(sources)} sources")
    return sources


# ── REVIEWS (paginated) ───────────────────────────────────────────────────────
def fetch_reviews(token: str, from_date: str, to_date: str) -> list:
    print(f"📥 Fetching reviews {from_date[:10]} → {to_date[:10]}...")
    all_reviews = []
    offset = 0
    limit = 1000

    while True:
        payload = {
            "mentionTimeAfter": from_date,
            "mentionTimeBefore": to_date,
            "modifiedTimeFrom": from_date,
            "additionalFields": ["brand", "visitRecordId"],
            "offset": offset,
            "limit": limit,
        }
        response = requests.post(
            REVIEWS_URL,
            headers={"Content-Type": "application/json", "accessToken": token},
            json=payload,
            timeout=60,
        )
        response.raise_for_status()
        batch = response.json().get("reviews", [])
        all_reviews.extend(batch)
        print(f"   Fetched {len(all_reviews)} reviews so far...")
        if len(batch) < limit:
            break
        offset += limit

    print(f"✅ Total reviews: {len(all_reviews)}")
    return all_reviews


# ── ENGAGEMENT RATE ───────────────────────────────────────────────────────────
def fetch_engagement_rate(token: str, node_id: str, from_date: str, to_date: str) -> dict:
    print(f"📊 Fetching engagement rate for node {node_id}...")
    response = requests.get(
        ENGAGEMENT_URL,
        headers={"accessToken": token},
        params={"accessToken": token, "node": node_id, "fromDate": from_date, "toDate": to_date},
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


# ── REGION MAPPING ────────────────────────────────────────────────────────────
def load_region_map() -> tuple:
    map_path = os.path.join(os.path.dirname(__file__), "region_map.json")
    with open(map_path) as f:
        data = json.load(f)

    name_map = {}
    for region, locations in data["regions"].items():
        for loc in locations:
            name_map[loc.lower().strip()] = region

    id_map = {k: v for k, v in data.get("location_id_map", {}).items() if not k.startswith("_")}
    return name_map, id_map


def assign_region(location_name: str, location_id: str, name_map: dict, id_map: dict) -> str:
    if str(location_id) in id_map:
        return id_map[str(location_id)]
    if location_name:
        key = location_name.lower().strip()
        if key in name_map:
            return name_map[key]
        for map_name, region in name_map.items():
            if map_name in key or key in map_name:
                return region
    return "Unknown"


# ── METRICS ───────────────────────────────────────────────────────────────────
def calculate_metrics(reviews: list, name_map: dict, id_map: dict) -> dict:
    print("🧮 Calculating metrics...")

    region_keys = ["North", "South", "West", "Unknown"]
    region_tallies = {
        r: {"total": 0, "closed": 0, "times_to_close": [], "positive": 0, "neutral": 0, "negative": 0}
        for r in region_keys
    }
    source_tallies = {}
    owner_tallies  = {}
    location_tallies = {}
    open_tasks = []
    critical_alerts = []

    for review in reviews:
        loc_name  = review.get("location", {}).get("name", "")
        loc_id    = str(review.get("location", {}).get("id", ""))
        region    = assign_region(loc_name, loc_id, name_map, id_map)
        task      = review.get("taskDetails") or {}
        status    = task.get("status", "pending").lower()
        owner     = task.get("ownerName") or "Unassigned"
        source    = review.get("source", {}).get("name", "Unknown")
        score     = review.get("reviewScore") or 0
        mention_t = review.get("mentionTime", "")

        sentiment = "positive" if score >= 4.0 else ("neutral" if score >= 2.5 else "negative")

        # Time to close
        ttc_hours = None
        engagements = review.get("engagements") or []
        if engagements and mention_t:
            try:
                from dateutil import parser as dtp
                t0 = dtp.parse(mention_t)
                t1 = dtp.parse(engagements[0].get("engagementTime", ""))
                ttc_hours = abs((t1 - t0).total_seconds() / 3600)
            except Exception:
                pass

        # Region tallies
        rt = region_tallies[region]
        rt["total"] += 1
        if status == "completed":
            rt["closed"] += 1
        if ttc_hours is not None:
            rt["times_to_close"].append(ttc_hours)
        rt[sentiment] += 1

        # Source tallies
        source_tallies.setdefault(source, {"total": 0, "closed": 0})
        source_tallies[source]["total"] += 1
        if status == "completed":
            source_tallies[source]["closed"] += 1

        # Owner tallies
        owner_tallies.setdefault(owner, {"total": 0, "closed": 0, "times_to_close": []})
        owner_tallies[owner]["total"] += 1
        if status == "completed":
            owner_tallies[owner]["closed"] += 1
        if ttc_hours is not None:
            owner_tallies[owner]["times_to_close"].append(ttc_hours)

        # Location tallies
        location_tallies.setdefault(loc_name, {"total": 0, "closed": 0, "region": region})
        location_tallies[loc_name]["total"] += 1
        if status == "completed":
            location_tallies[loc_name]["closed"] += 1

        # Open tasks
        if status != "completed":
            try:
                from dateutil import parser as dtp
                days_open = (datetime.now(timezone.utc) - dtp.parse(mention_t).astimezone(timezone.utc)).days
            except Exception:
                days_open = 0

            text = ""
            contents = review.get("contents") or []
            if contents:
                text = contents[0].get("response", "")

            task_entry = {
                "review_id":   review.get("id"),
                "location":    loc_name,
                "region":      region,
                "source":      source,
                "score":       score,
                "owner":       owner,
                "days_open":   days_open,
                "mention_time": mention_t,
                "text":        text,
            }
            open_tasks.append(task_entry)
            if days_open >= 14 or score < 1.5:
                critical_alerts.append(task_entry)

    # Network totals
    net_total  = len(reviews)
    net_closed = sum(1 for r in reviews if (r.get("taskDetails") or {}).get("status", "").lower() == "completed")
    all_ttc    = [t for rt in region_tallies.values() for t in rt["times_to_close"]]
    net_pos    = sum(rt["positive"] for rt in region_tallies.values())
    net_neu    = sum(rt["neutral"]  for rt in region_tallies.values())
    net_neg    = sum(rt["negative"] for rt in region_tallies.values())

    def eng_rate(closed, total):
        return round(closed / total * 100, 1) if total else 0

    def avg_ttc(times):
        return round(sum(times) / len(times), 1) if times else 0

    metrics = {
        "total_tasks": net_total,
        "network": {
            "total_tasks":             net_total,
            "closed_tasks":            net_closed,
            "engagement_rate":         eng_rate(net_closed, net_total),
            "avg_time_to_close_hours": avg_ttc(all_ttc),
            "positive": net_pos,
            "neutral":  net_neu,
            "negative": net_neg,
        },
        "regions": {
            region: {
                "total_tasks":             rt["total"],
                "closed_tasks":            rt["closed"],
                "engagement_rate":         eng_rate(rt["closed"], rt["total"]),
                "avg_time_to_close_hours": avg_ttc(rt["times_to_close"]),
                "positive": rt["positive"],
                "neutral":  rt["neutral"],
                "negative": rt["negative"],
            }
            for region, rt in region_tallies.items()
        },
        "by_source": {
            src: {
                "total":           st["total"],
                "closed":          st["closed"],
                "engagement_rate": eng_rate(st["closed"], st["total"]),
            }
            for src, st in source_tallies.items()
        },
        "by_owner": {
            owner: {
                "total":                   ot["total"],
                "closed":                  ot["closed"],
                "engagement_rate":         eng_rate(ot["closed"], ot["total"]),
                "avg_time_to_close_hours": avg_ttc(ot["times_to_close"]),
            }
            for owner, ot in owner_tallies.items()
        },
        "by_location": {
            loc: {
                "total":           lt["total"],
                "closed":          lt["closed"],
                "region":          lt["region"],
                "engagement_rate": eng_rate(lt["closed"], lt["total"]),
            }
            for loc, lt in location_tallies.items()
        },
        "open_tasks":      open_tasks,
        "critical_alerts": critical_alerts,
        "reviews_raw":     reviews,
    }

    print(f"✅ Network: {net_total} tasks, {metrics['network']['engagement_rate']}% engagement")
    for region in ["North", "South", "West"]:
        rm = metrics["regions"][region]
        if rm["total_tasks"] > 0:
            print(f"   {region}: {rm['total_tasks']} tasks, {rm['engagement_rate']}%")
    return metrics


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--month",  default="", help="YYYY-MM (default: previous month)")
    parser.add_argument("--node",   default=os.environ.get("PG_NODE_ID", ""))
    parser.add_argument("--output", default="orm-reports/scripts/data.json")
    args = parser.parse_args()

    if args.month:
        report_date = datetime.strptime(args.month, "%Y-%m").replace(tzinfo=timezone.utc)
    else:
        report_date = (datetime.now(timezone.utc).replace(day=1) - relativedelta(months=1))

    from_date   = report_date.strftime("%Y-%m-01T00:00:00.000Z")
    last_day    = report_date + relativedelta(months=1) - relativedelta(days=1)
    to_date     = last_day.strftime("%Y-%m-%dT23:59:59.000Z")
    month_label = report_date.strftime("%B %Y")

    print(f"\n{'='*55}")
    print(f"  WMCHealth ORM CTL Report — {month_label}")
    print(f"  Range: {from_date[:10]}  →  {to_date[:10]}")
    print(f"{'='*55}\n")

    app_id     = os.environ.get("PG_APP_ID")
    app_secret = os.environ.get("PG_APP_SECRET")
    if not app_id or not app_secret:
        raise EnvironmentError(
            "PG_APP_ID and PG_APP_SECRET must be set.\n"
            "Add them in GitHub → Settings → Secrets → Actions."
        )

    token   = get_access_token(app_id, app_secret)
    sources = fetch_sources(token)
    reviews = fetch_reviews(token, from_date, to_date)

    engagement_data = {}
    if args.node:
        engagement_data = fetch_engagement_rate(
            token, args.node,
            from_date[:10] + "T00:00:00Z",
            to_date[:10]   + "T23:59:59Z",
        )

    name_map, id_map = load_region_map()
    metrics = calculate_metrics(reviews, name_map, id_map)

    metrics["meta"] = {
        "month_label":        month_label,
        "from_date":          from_date,
        "to_date":            to_date,
        "generated_at":       datetime.now(timezone.utc).isoformat(),
        "sources":            sources,
        "engagement_api_data": engagement_data,
    }

    os.makedirs(os.path.dirname(args.output), exist_ok=True)
    with open(args.output, "w") as f:
        json.dump(metrics, f, indent=2, default=str)

    print(f"\n✅ Data saved → {args.output}")
    print(f"   Open tasks:      {len(metrics['open_tasks'])}")
    print(f"   Critical alerts: {len(metrics['critical_alerts'])}")


if __name__ == "__main__":
    main()
