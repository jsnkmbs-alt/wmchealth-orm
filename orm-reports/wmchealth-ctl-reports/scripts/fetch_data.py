"""
fetch_data.py
-------------
Authenticates with the Press Ganey API and pulls all review + engagement
data needed to build the WMCHealth CTL monthly report.

Environment variables required:
  PG_APP_ID      - Press Ganey Application ID
  PG_APP_SECRET  - Press Ganey Application Secret

Usage:
  python scripts/fetch_data.py --month 2026-03
  python scripts/fetch_data.py          # defaults to previous calendar month
"""

import os
import json
import argparse
import requests
from datetime import datetime, timezone
from dateutil.relativedelta import relativedelta


# ── API BASE ──────────────────────────────────────────────────────────────────
BASE_URL = "https://api1.consumerism.pressganey.com"
TOKEN_URL = f"{BASE_URL}/api/service/v1/token/create"
REVIEWS_URL = f"{BASE_URL}/api/reputation/v1/reviews"
DELETED_URL = f"{BASE_URL}/api/reputation/v1/deleted-reviews"
SOURCES_URL = f"{BASE_URL}/api/reputation/v1/sources"
ENGAGEMENT_URL = f"{BASE_URL}/api/service/v1/engagements/statistics"


# ── AUTHENTICATION ────────────────────────────────────────────────────────────
def get_access_token(app_id: str, app_secret: str) -> str:
    """Request a new access token from Press Ganey."""
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

    token = data["accessToken"]
    expires = data.get("expiresIn", "unknown")
    print(f"✅ Token obtained. Expires: {expires}")
    return token


# ── SOURCES LOOKUP ────────────────────────────────────────────────────────────
def fetch_sources(token: str) -> list:
    """Fetch all available review sources (Google, Yelp, etc.)"""
    print("📋 Fetching sources...")
    response = requests.get(
        SOURCES_URL,
        headers={"accessToken": token},
        timeout=30,
    )
    response.raise_for_status()
    sources = response.json()
    print(f"   Found {len(sources)} sources")
    return sources


# ── REVIEWS ───────────────────────────────────────────────────────────────────
def fetch_reviews(token: str, from_date: str, to_date: str) -> list:
    """
    Fetch all reviews for the given date range using pagination.
    Uses mentionTimeAfter/Before so we get reviews by when they were posted,
    not by when PG last modified them internally.
    """
    print(f"📥 Fetching reviews from {from_date} to {to_date}...")
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
            headers={
                "Content-Type": "application/json",
                "accessToken": token,
            },
            json=payload,
            timeout=60,
        )
        response.raise_for_status()
        data = response.json()
        batch = data.get("reviews", [])
        all_reviews.extend(batch)
        print(f"   Fetched {len(all_reviews)} reviews so far (offset={offset})...")

        # Stop paginating when we get fewer results than the limit
        if len(batch) < limit:
            break
        offset += limit

    print(f"✅ Total reviews fetched: {len(all_reviews)}")
    return all_reviews


# ── ENGAGEMENT RATE ───────────────────────────────────────────────────────────
def fetch_engagement_rate(token: str, node_id: str, from_date: str, to_date: str) -> dict:
    """Fetch engagement/response rate statistics for a given node."""
    print(f"📊 Fetching engagement rate for node {node_id}...")
    response = requests.get(
        ENGAGEMENT_URL,
        headers={"accessToken": token},
        params={
            "accessToken": token,
            "node": node_id,
            "fromDate": from_date,
            "toDate": to_date,
        },
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


# ── REGION MAPPING ────────────────────────────────────────────────────────────
def load_region_map() -> dict:
    """Load the region mapping from region_map.json."""
    map_path = os.path.join(os.path.dirname(__file__), "region_map.json")
    with open(map_path, "r") as f:
        data = json.load(f)

    # Build a flat name → region lookup
    flat_map = {}
    for region, locations in data["regions"].items():
        for loc_name in locations:
            flat_map[loc_name.lower().strip()] = region

    # Also build ID-based map if populated
    id_map = {}
    for loc_id, region in data.get("location_id_map", {}).items():
        if not loc_id.startswith("_"):
            id_map[loc_id] = region

    return flat_map, id_map


def assign_region(location_name: str, location_id: str, name_map: dict, id_map: dict) -> str:
    """Assign a region to a review based on location name or ID."""
    # Try ID first (more reliable)
    if str(location_id) in id_map:
        return id_map[str(location_id)]
    # Fall back to name matching
    if location_name:
        key = location_name.lower().strip()
        if key in name_map:
            return name_map[key]
        # Partial match
        for map_name, region in name_map.items():
            if map_name in key or key in map_name:
                return region
    return "Unknown"


# ── METRICS CALCULATION ───────────────────────────────────────────────────────
def calculate_metrics(reviews: list, name_map: dict, id_map: dict) -> dict:
    """
    Calculate all metrics needed for the PowerPoint deck from raw reviews.
    Returns a structured dict mirroring the slide data requirements.
    """
    print("🧮 Calculating metrics...")

    metrics = {
        "total_tasks": len(reviews),
        "network": {},
        "regions": {"North": {}, "South": {}, "West": {}, "Unknown": {}},
        "by_source": {},
        "by_owner": {},
        "by_location": {},
        "reviews_raw": reviews,  # Keep raw for sentiment/alert slides
        "open_tasks": [],
        "critical_alerts": [],
    }

    # Tally counters
    region_tallies = {r: {"total": 0, "closed": 0, "times_to_close": [], "positive": 0, "neutral": 0, "negative": 0} for r in ["North", "South", "West", "Unknown"]}
    source_tallies = {}
    owner_tallies = {}
    location_tallies = {}

    for review in reviews:
        location_name = review.get("location", {}).get("name", "")
        location_id = str(review.get("location", {}).get("id", ""))
        region = assign_region(location_name, location_id, name_map, id_map)

        task = review.get("taskDetails", {}) or {}
        status = task.get("status", "pending").lower()
        owner_name = task.get("ownerName", "Unassigned") or "Unassigned"
        source_name = review.get("source", {}).get("name", "Unknown")
        review_score = review.get("reviewScore", 0) or 0
        mention_time = review.get("mentionTime", "")

        # Sentiment bucket
        if review_score >= 4.0:
            sentiment = "positive"
        elif review_score >= 2.5:
            sentiment = "neutral"
        else:
            sentiment = "negative"

        # Time to close
        ttc_hours = None
        engagements = review.get("engagements") or review.get("engagement") or []
        if engagements and mention_time:
            try:
                from dateutil import parser as dtparser
                first_eng = engagements[0].get("engagementTime", "") if isinstance(engagements, list) else ""
                if first_eng:
                    t_mention = dtparser.parse(mention_time)
                    t_engage = dtparser.parse(first_eng)
                    ttc_hours = abs((t_engage - t_mention).total_seconds() / 3600)
            except Exception:
                pass

        # ── Region tallies
        rt = region_tallies[region]
        rt["total"] += 1
        if status == "completed":
            rt["closed"] += 1
        if ttc_hours is not None:
            rt["times_to_close"].append(ttc_hours)
        rt[sentiment] += 1

        # ── Source tallies
        if source_name not in source_tallies:
            source_tallies[source_name] = {"total": 0, "closed": 0}
        source_tallies[source_name]["total"] += 1
        if status == "completed":
            source_tallies[source_name]["closed"] += 1

        # ── Owner tallies
        if owner_name not in owner_tallies:
            owner_tallies[owner_name] = {"total": 0, "closed": 0, "times_to_close": [], "region": region}
        owner_tallies[owner_name]["total"] += 1
        if status == "completed":
            owner_tallies[owner_name]["closed"] += 1
        if ttc_hours is not None:
            owner_tallies[owner_name]["times_to_close"].append(ttc_hours)

        # ── Location tallies
        if location_name not in location_tallies:
            location_tallies[location_name] = {"total": 0, "closed": 0, "region": region}
        location_tallies[location_name]["total"] += 1
        if status == "completed":
            location_tallies[location_name]["closed"] += 1

        # ── Open task tracking
        if status != "completed":
            from datetime import datetime, timezone
            try:
                from dateutil import parser as dtparser
                days_open = (datetime.now(timezone.utc) - dtparser.parse(mention_time).astimezone(timezone.utc)).days if mention_time else 0
            except Exception:
                days_open = 0

            open_task = {
                "review_id": review.get("id"),
                "location": location_name,
                "region": region,
                "source": source_name,
                "score": review_score,
                "owner": owner_name,
                "days_open": days_open,
                "mention_time": mention_time,
                "text": (review.get("contents") or [{}])[0].get("response", ""),
            }
            metrics["open_tasks"].append(open_task)

            # Flag as critical alert
            if days_open >= 14 or review_score < 1.5:
                metrics["critical_alerts"].append(open_task)

    # ── Network totals
    net_total = len(reviews)
    net_closed = sum(1 for r in reviews if (r.get("taskDetails") or {}).get("status", "").lower() == "completed")
    all_ttc = [t for rt in region_tallies.values() for t in rt["times_to_close"]]

    metrics["network"] = {
        "total_tasks": net_total,
        "closed_tasks": net_closed,
        "engagement_rate": round((net_closed / net_total * 100), 1) if net_total else 0,
        "avg_time_to_close_hours": round(sum(all_ttc) / len(all_ttc), 1) if all_ttc else 0,
        "positive": sum(rt["positive"] for rt in region_tallies.values()),
        "neutral": sum(rt["neutral"] for rt in region_tallies.values()),
        "negative": sum(rt["negative"] for rt in region_tallies.values()),
    }

    # ── Region summaries
    for region, rt in region_tallies.items():
        ttc_avg = round(sum(rt["times_to_close"]) / len(rt["times_to_close"]), 1) if rt["times_to_close"] else 0
        metrics["regions"][region] = {
            "total_tasks": rt["total"],
            "closed_tasks": rt["closed"],
            "engagement_rate": round((rt["closed"] / rt["total"] * 100), 1) if rt["total"] else 0,
            "avg_time_to_close_hours": ttc_avg,
            "positive": rt["positive"],
            "neutral": rt["neutral"],
            "negative": rt["negative"],
        }

    # ── Source summaries
    for src, st in source_tallies.items():
        metrics["by_source"][src] = {
            "total": st["total"],
            "closed": st["closed"],
            "engagement_rate": round((st["closed"] / st["total"] * 100), 1) if st["total"] else 0,
        }

    # ── Owner summaries
    for owner, ot in owner_tallies.items():
        ttc_avg = round(sum(ot["times_to_close"]) / len(ot["times_to_close"]), 1) if ot["times_to_close"] else 0
        metrics["by_owner"][owner] = {
            "total": ot["total"],
            "closed": ot["closed"],
            "engagement_rate": round((ot["closed"] / ot["total"] * 100), 1) if ot["total"] else 0,
            "avg_time_to_close_hours": ttc_avg,
        }

    # ── Location summaries
    for loc, lt in location_tallies.items():
        metrics["by_location"][loc] = {
            "total": lt["total"],
            "closed": lt["closed"],
            "region": lt["region"],
            "engagement_rate": round((lt["closed"] / lt["total"] * 100), 1) if lt["total"] else 0,
        }

    print(f"✅ Metrics calculated:")
    print(f"   Network: {metrics['network']['total_tasks']} tasks, {metrics['network']['engagement_rate']}% engaged")
    for region, rm in metrics["regions"].items():
        if rm.get("total_tasks", 0) > 0:
            print(f"   {region}: {rm['total_tasks']} tasks, {rm['engagement_rate']}% engaged")

    return metrics


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Fetch Press Ganey CTL data")
    parser.add_argument(
        "--month",
        type=str,
        help="Month to fetch in YYYY-MM format (default: previous month)",
    )
    parser.add_argument(
        "--node",
        type=str,
        default=os.environ.get("PG_NODE_ID", ""),
        help="Press Ganey node ID for engagement rate (optional)",
    )
    parser.add_argument(
        "--output",
        type=str,
        default="scripts/data.json",
        help="Output file path for the metrics JSON",
    )
    args = parser.parse_args()

    # Determine date range
    if args.month:
        report_date = datetime.strptime(args.month, "%Y-%m")
    else:
        report_date = datetime.now(timezone.utc).replace(day=1) - relativedelta(months=1)

    from_date = report_date.strftime("%Y-%m-01T00:00:00.000Z")
    # Last day of month
    last_day = (report_date + relativedelta(months=1) - relativedelta(days=1))
    to_date = last_day.strftime(f"%Y-%m-%dT23:59:59.000Z")
    month_label = report_date.strftime("%B %Y")

    print(f"\n{'='*55}")
    print(f"  WMCHealth CTL Report — {month_label}")
    print(f"  Date range: {from_date[:10]} → {to_date[:10]}")
    print(f"{'='*55}\n")

    # Get credentials from environment
    app_id = os.environ.get("PG_APP_ID")
    app_secret = os.environ.get("PG_APP_SECRET")
    if not app_id or not app_secret:
        raise EnvironmentError(
            "PG_APP_ID and PG_APP_SECRET environment variables are required.\n"
            "Set them in your GitHub repository Secrets."
        )

    # Authenticate
    token = get_access_token(app_id, app_secret)

    # Fetch sources
    sources = fetch_sources(token)

    # Fetch reviews
    reviews = fetch_reviews(token, from_date, to_date)

    # Fetch engagement rate (if node ID provided)
    engagement_data = {}
    if args.node:
        engagement_data = fetch_engagement_rate(token, args.node, from_date[:10] + "T00:00:00Z", to_date[:10] + "T23:59:59Z")

    # Load region map and calculate metrics
    name_map, id_map = load_region_map()
    metrics = calculate_metrics(reviews, name_map, id_map)

    # Add metadata
    metrics["meta"] = {
        "month_label": month_label,
        "from_date": from_date,
        "to_date": to_date,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "sources": sources,
        "engagement_api_data": engagement_data,
    }

    # Save to JSON
    output_path = args.output
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w") as f:
        json.dump(metrics, f, indent=2, default=str)

    print(f"\n✅ Data saved to {output_path}")
    print(f"   Open tasks requiring attention: {len(metrics['open_tasks'])}")
    print(f"   Critical alerts flagged: {len(metrics['critical_alerts'])}")
    print(f"\nNext step: python scripts/build_report.py --data {output_path}")


if __name__ == "__main__":
    main()
