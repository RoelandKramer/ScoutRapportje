# app.py
"""
Streamlit app:
- Generate SciSports token
- Search + select a player (name, age, position, club)
- Fill a PPTX scouting template by replacing {Placeholders}

Template expected placeholders (from uploaded PPTX):
{ DD/MM/YYYY }, { Place }, { Country }, { Height }, { Preferred Foot }, { Position },
{ League }, { Club }, {Season_m}, {season_min}, {season_g}, {season_a},
{Career_m}, {career_min}, {career_g}, {career_a}, {con_DD/MM/YYYY},
{TV}, {MV}, {Agency}, {Agent}
"""

from __future__ import annotations

import json
import os
import re
import tempfile
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any, Dict, Iterable, List, Optional, Tuple

import requests
import streamlit as st
from pptx import Presentation  # python-pptx


API_BASE = "https://api-recruitment.scisports.app/api"
TOKEN_URL = "https://identity.scisports.app/connect/token"
TEMPLATE_PATH = "TemplateScoutingsRapport.pptx"


# -----------------------------
# Models
# -----------------------------
@dataclass(frozen=True)
class PlayerOption:
    player_id: int
    name: str
    age: Optional[int]
    position: str
    club: str
    league: str

    def label(self) -> str:
        age_str = "?" if self.age is None else str(self.age)
        pos = self.position or "UnknownPos"
        club = self.club or "UnknownClub"
        league = self.league or "UnknownLeague"
        return f"{self.name} — {age_str} — {pos} — {club} ({league})"


# -----------------------------
# Helpers
# -----------------------------
def _http_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"Accept": "application/json"})
    return s


def _auth_headers(access_token: str) -> Dict[str, str]:
    return {"Authorization": f"Bearer {access_token}"}


def _safe_get(d: Dict[str, Any], path: str, default: Any = None) -> Any:
    cur: Any = d
    for part in path.split("."):
        if not isinstance(cur, dict) or part not in cur:
            return default
        cur = cur[part]
    return cur


def _iso_to_ddmmyyyy(value: Optional[str]) -> str:
    if not value:
        return ""
    try:
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
        return dt.strftime("%d/%m/%Y")
    except Exception:
        # Also accepts yyyy-mm-dd
        try:
            dt = datetime.strptime(value[:10], "%Y-%m-%d")
            return dt.strftime("%d/%m/%Y")
        except Exception:
            return value


def _fmt_int(value: Any) -> str:
    try:
        if value is None:
            return ""
        return f"{int(value)}"
    except Exception:
        return ""


def _fmt_money_eur(value: Any) -> str:
    try:
        if value is None:
            return ""
        v = float(value)
        # 12_345_678 -> €12.35M
        if abs(v) >= 1_000_000:
            return f"€{v/1_000_000:.2f}M"
        if abs(v) >= 1_000:
            return f"€{v/1_000:.1f}K"
        return f"€{v:.0f}"
    except Exception:
        return ""


def _first_position(player_info: Dict[str, Any]) -> str:
    positions = player_info.get("positions") or []
    if isinstance(positions, list) and positions:
        return str(positions[0])
    return ""


def _as_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (int, float, bool)):
        return str(value)
    return str(value)


# -----------------------------
# SciSports API
# -----------------------------
def token_password_grant(
    username: str,
    password: str,
    client_id: str,
    client_secret: str,
    scope: str = "api recruitment",
    timeout_s: int = 30,
) -> str:
    """
    Uses the password grant as you provided.
    Returns access_token or raises.
    """
    payload = {
        "grant_type": "password",
        "username": username,
        "password": password,
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": scope,
    }
    resp = requests.post(TOKEN_URL, data=payload, timeout=timeout_s)
    resp.raise_for_status()
    data = resp.json()
    token = data.get("access_token")
    if not token:
        raise RuntimeError(f"Token response missing access_token: {json.dumps(data)[:500]}")
    return token


@st.cache_data(show_spinner=False, ttl=60 * 15)
def search_players(
    access_token: str,
    search_text: str,
    limit: int = 50,
    offset: int = 0,
    context: str = "Male",
) -> Tuple[int, List[PlayerOption]]:
    """
    Returns (total, options) for a search query.
    Cached to avoid re-fetching the same searches.
    """
    s = _http_session()
    params = {
        "offset": offset,
        "limit": limit,
        "context": context,
    }
    if search_text.strip():
        params["searchText"] = search_text.strip()

    url = f"{API_BASE}/v2/players"
    resp = s.get(url, headers=_auth_headers(access_token), params=params, timeout=30)
    resp.raise_for_status()
    payload = resp.json()
    total = int(payload.get("total", 0))
    items = payload.get("items") or []

    options: List[PlayerOption] = []
    for it in items:
        info = it.get("info") or {}
        team = it.get("team") or {}
        league = it.get("league") or {}
        player_id = info.get("id")
        if player_id is None:
            continue

        options.append(
            PlayerOption(
                player_id=int(player_id),
                name=_as_text(info.get("name") or info.get("footballName") or ""),
                age=info.get("age"),
                position=_first_position(info),
                club=_as_text(team.get("name") or ""),
                league=_as_text(league.get("name") or ""),
            )
        )

    return total, options


def get_player(access_token: str, player_id: int) -> Dict[str, Any]:
    s = _http_session()
    url = f"{API_BASE}/v2/players/{player_id}"
    resp = s.get(url, headers=_auth_headers(access_token), timeout=30)
    resp.raise_for_status()
    return resp.json()


def get_sciskill(access_token: str, player_id: int, context: str = "Male") -> Optional[Dict[str, Any]]:
    s = _http_session()
    url = f"{API_BASE}/v2/metrics/players/sciskill"
    params = {"offset": 0, "limit": 1, "playerIds": player_id, "context": context}
    resp = s.get(url, headers=_auth_headers(access_token), params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    items = data.get("items") or []
    return items[0] if items else None


def get_latest_transfer_fee(access_token: str, player_id: int, context: str = "Male") -> Optional[Dict[str, Any]]:
    s = _http_session()
    url = f"{API_BASE}/v2/metrics/players/transfer-fees"
    params = {
        "offset": 0,
        "limit": 1,
        "playerIds": player_id,
        "latestTransferFee": "true",
        "context": context,
    }
    resp = s.get(url, headers=_auth_headers(access_token), params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    items = data.get("items") or []
    return items[0] if items else None


def get_career_stats_all(access_token: str, player_id: int, context: str = "Male") -> List[Dict[str, Any]]:
    """
    Pulls all career-stats items for a player by paginating.
    Endpoint name matches your docs example: /api/v2/metrics/career-stats/Players
    (we use lowercase 'players' in the path, but keep a fallback).
    """
    s = _http_session()
    items: List[Dict[str, Any]] = []
    offset = 0
    limit = 50

    # Try both variants just in case the API is case-sensitive upstream.
    paths = [
        f"{API_BASE}/v2/metrics/career-stats/players",
        f"{API_BASE}/v2/metrics/career-stats/Players",
    ]

    url: Optional[str] = None
    last_err: Optional[Exception] = None

    for candidate in paths:
        try:
            test = s.get(
                candidate,
                headers=_auth_headers(access_token),
                params={"offset": 0, "limit": 1, "playerIds": player_id, "context": context},
                timeout=30,
            )
            if test.status_code < 400:
                url = candidate
                break
        except Exception as e:
            last_err = e

    if not url:
        raise RuntimeError(f"Could not reach career-stats endpoint. Last error: {last_err}")

    while True:
        resp = s.get(
            url,
            headers=_auth_headers(access_token),
            params={"offset": offset, "limit": limit, "playerIds": player_id, "context": context},
            timeout=30,
        )
        resp.raise_for_status()
        payload = resp.json()
        batch = payload.get("items") or []
        items.extend(batch)

        total = int(payload.get("total", len(items)))
        offset += limit
        if len(items) >= total or not batch:
            break

    return items


def summarize_career_stats(items: List[Dict[str, Any]]) -> Tuple[Dict[str, int], Dict[str, int]]:
    """
    Returns (latest_season_summary, career_total_summary)
    with keys: matches, minutes, goals, assists

    We infer keys from the response:
    - minutesPlayed
    - matchesPlayed
    - goal
    - assist
    """
    def extract_stats(it: Dict[str, Any]) -> Dict[str, int]:
        stats = it.get("stats") or {}
        return {
            "matches": int(stats.get("matchesPlayed", 0) or 0),
            "minutes": int(stats.get("minutesPlayed", 0) or 0),
            "goals": int(stats.get("goal", 0) or 0),
            "assists": int(stats.get("assist", 0) or 0),
        }

    def season_start(it: Dict[str, Any]) -> str:
        # prefer season.startDate if present
        season = it.get("season") or {}
        sd = season.get("startDate")
        if isinstance(sd, str):
            return sd
        return ""

    if not items:
        zero = {"matches": 0, "minutes": 0, "goals": 0, "assists": 0}
        return zero, zero

    # Latest season = max by season.startDate; fallback to first item.
    items_sorted = sorted(items, key=season_start)
    latest = items_sorted[-1]
    latest_stats = extract_stats(latest)

    career = {"matches": 0, "minutes": 0, "goals": 0, "assists": 0}
    for it in items:
        s = extract_stats(it)
        for k in career:
            career[k] += s[k]

    return latest_stats, career


# -----------------------------
# PPTX templating
# -----------------------------
PLACEHOLDER_RE = re.compile(r"\{[^{}]+\}")

def replace_placeholders_in_pptx(
    template_path: str,
    output_path: str,
    replacements: Dict[str, str],
) -> None:
    """
    Replaces placeholders in text runs to preserve formatting as much as possible.
    replacements keys must match the placeholder text exactly (including braces/spaces),
    e.g. "{ Place }" -> "Rosario".
    """
    prs = Presentation(template_path)

    for slide in prs.slides:
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            tf = shape.text_frame
            for paragraph in tf.paragraphs:
                for run in paragraph.runs:
                    txt = run.text
                    if "{" not in txt or "}" not in txt:
                        continue
                    for k, v in replacements.items():
                        if k in txt:
                            run.text = txt.replace(k, v)
                            txt = run.text

    prs.save(output_path)


def build_replacements(
    player: Dict[str, Any],
    sciskill: Optional[Dict[str, Any]],
    transfer_fee: Optional[Dict[str, Any]],
    latest_season: Dict[str, int],
    career_total: Dict[str, int],
) -> Dict[str, str]:
    info = player.get("info") or {}
    team = player.get("team") or {}
    league = player.get("league") or {}
    contract = player.get("contract") or {}

    today = date.today().strftime("%d/%m/%Y")

    birth_place = _as_text(info.get("birthPlace") or "")
    birth_country = _as_text(_safe_get(info, "birthCountry.name", "") or "")
    height = _fmt_int(info.get("height"))
    preferred_foot = _as_text(info.get("preferredFoot") or "")
    position = _first_position(info)
    league_name = _as_text(league.get("name") or "")
    club_name = _as_text(team.get("name") or "")

    contract_end = _iso_to_ddmmyyyy(contract.get("contractEnd"))
    market_value = _fmt_money_eur(contract.get("marketValue"))

    tv = ""
    if transfer_fee:
        tv = _fmt_money_eur(transfer_fee.get("valueEstimateEur"))

    # Agency/Agent often not present in this API; keep empty by default.
    agency = ""
    agent = ""

    # Optional: if you want to show SciSkill somewhere later, you can add placeholders.
    _ = sciskill

    return {
        "{ DD/MM/YYYY }": today,
        "{ Place }": birth_place,
        "{ Country }": birth_country,
        "{ Height }": height,
        "{ Preferred Foot }": preferred_foot,
        "{ Position }": position,
        "{ League }": league_name,
        "{ Club }": club_name,
        "{Season_m}": _fmt_int(latest_season.get("matches")),
        "{season_min}": _fmt_int(latest_season.get("minutes")),
        "{season_g}": _fmt_int(latest_season.get("goals")),
        "{season_a}": _fmt_int(latest_season.get("assists")),
        "{Career_m}": _fmt_int(career_total.get("matches")),
        "{career_min}": _fmt_int(career_total.get("minutes")),
        "{career_g}": _fmt_int(career_total.get("goals")),
        "{career_a}": _fmt_int(career_total.get("assists")),
        "{con_DD/MM/YYYY}": contract_end,
        "{TV}": tv,
        "{MV}": market_value,
        "{Agency}": agency,
        "{Agent}": agent,
    }


# -----------------------------
# Streamlit UI
# -----------------------------
def main() -> None:
    st.set_page_config(page_title="SciSports Scouting Form Generator", layout="centered")
    st.title("SciSports Scouting Form Generator")

    with st.expander("1) Generate API key (access token)", expanded=True):
        col1, col2 = st.columns(2)
        username = col1.text_input("Username", value=st.secrets.get("SCISPORTS_USERNAME", ""))
        password = col2.text_input("Password", type="password", value=st.secrets.get("SCISPORTS_PASSWORD", ""))

        col3, col4 = st.columns(2)
        client_id = col3.text_input("Client ID", value=st.secrets.get("SCISPORTS_CLIENT_ID", ""))
        client_secret = col4.text_input("Client Secret", type="password", value=st.secrets.get("SCISPORTS_CLIENT_SECRET", ""))

        scope = st.text_input("Scope", value=st.secrets.get("SCISPORTS_SCOPE", "api recruitment"))
        context = st.selectbox("Context", options=["Male", "Female"], index=0)

        gen = st.button("Generate API key", type="primary")

        if gen:
            try:
                token = token_password_grant(
                    username=username.strip(),
                    password=password,
                    client_id=client_id.strip(),
                    client_secret=client_secret,
                    scope=scope.strip(),
                )
                st.session_state["SCISPORTS_ACCESS_TOKEN"] = token
                st.session_state["SCISPORTS_CONTEXT"] = context
                os.environ["SCISPORTS_ACCESS_TOKEN"] = token
                st.success("Access token generated and stored for this session.")
            except Exception as e:
                st.error(f"Failed to generate token: {e}")

        if st.session_state.get("SCISPORTS_ACCESS_TOKEN"):
            st.info("Token present ✅ You can now search and select a player.")

    token = st.session_state.get("SCISPORTS_ACCESS_TOKEN", "")
    context = st.session_state.get("SCISPORTS_CONTEXT", "Male")

    st.divider()

    st.subheader("2) Select player")
    if not token:
        st.warning("Generate the API key first.")
        st.stop()

    q = st.text_input("Search player name", value="", placeholder="e.g. Messi")
    colp1, colp2, colp3 = st.columns([1, 1, 2])
    limit = colp1.selectbox("Limit", options=[25, 50, 100], index=1)
    offset = colp2.number_input("Offset", min_value=0, value=0, step=limit)
    do_search = colp3.button("Search", type="secondary")

    if do_search or "player_options" not in st.session_state:
        try:
            total, options = search_players(token, q, limit=int(limit), offset=int(offset), context=context)
            st.session_state["player_total"] = total
            st.session_state["player_options"] = options
        except Exception as e:
            st.error(f"Player search failed: {e}")
            st.stop()

    total = st.session_state.get("player_total", 0)
    options: List[PlayerOption] = st.session_state.get("player_options", [])
    st.caption(f"Results: showing {len(options)} of total {total} (use Offset/Limit for paging)")

    if not options:
        st.warning("No players found. Try a different search.")
        st.stop()

    selected_label = st.selectbox(
        "Player (Name — Age — Position — Club (League))",
        options=[o.label() for o in options],
        index=0,
    )
    selected = next(o for o in options if o.label() == selected_label)
    st.session_state["selected_player_id"] = selected.player_id

    st.divider()

    st.subheader("3) Generate Scoutings Form")
    st.caption(f"Template: {TEMPLATE_PATH}")

    if not os.path.exists(TEMPLATE_PATH):
        st.error(
            f"Template not found at {TEMPLATE_PATH}. "
            "In Streamlit Cloud, ensure the PPTX is in the repo root."
        )
        st.stop()

    if st.button("Generate Scoutings Form", type="primary"):
        with st.spinner("Fetching data and generating PPTX..."):
            try:
                player = get_player(token, selected.player_id)
                sciskill = get_sciskill(token, selected.player_id, context=context)
                transfer_fee = get_latest_transfer_fee(token, selected.player_id, context=context)

                career_items = get_career_stats_all(token, selected.player_id, context=context)
                latest_season, career_total = summarize_career_stats(career_items)

                replacements = build_replacements(
                    player=player,
                    sciskill=sciskill,
                    transfer_fee=transfer_fee,
                    latest_season=latest_season,
                    career_total=career_total,
                )

                with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
                    output_path = tmp.name

                replace_placeholders_in_pptx(TEMPLATE_PATH, output_path, replacements)

                with open(output_path, "rb") as f:
                    ppt_bytes = f.read()

                safe_name = re.sub(r"[^a-zA-Z0-9_-]+", "_", selected.name)[:80]
                out_filename = f"ScoutingsRapport_{safe_name}_{selected.player_id}.pptx"

                st.success("Generated ✅")
                st.download_button(
                    "Download filled Scoutings Rapport (.pptx)",
                    data=ppt_bytes,
                    file_name=out_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )

                with st.expander("Filled values (debug)"):
                    st.json(replacements)

            except Exception as e:
                st.error(f"Failed to generate PPTX: {e}")


if __name__ == "__main__":
    main()
