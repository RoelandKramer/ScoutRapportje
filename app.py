# app.py
"""
SciSports Scouting Form Generator

Repo layout:
.
├─ app.py
├─ requirements.txt
└─ TemplateScoutingsRapport.pptx

Streamlit Secrets (flat keys):
SCISPORTS_USERNAME="..."
SCISPORTS_PASSWORD="..."
SCISPORTS_CLIENT_ID="..."
SCISPORTS_CLIENT_SECRET="..."
SCISPORTS_SCOPE="api recruitment"  # optional (defaults to this)

Template placeholders (detected in TemplateScoutingsRapport.pptx):
{Name}
{ DD/MM/YYYY }
{ Place }
{Nationalities}
{ Height }
{ Preferred Foot }
{ Position }
{ League }
{ Club }
{Season_m} {season_min} {season_g} {season_a}
{Career_m} {career_min} {career_g} {career_a}
{con_DD/MM/YYYY} {TV} {MV} {Agency} {Agent}

Notes:
- Token replacement handles placeholders split across PowerPoint runs (common PPTX quirk).
- Position coloring colors the "1..11" position boxes in the bottom-left area.
"""

from __future__ import annotations

import json
import os
import re
import tempfile
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import requests
import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor

API_BASE = "https://api-recruitment.scisports.app/api"
TOKEN_URL = "https://identity.scisports.app/connect/token"
TEMPLATE_PATH = "TemplateScoutingsRapport.pptx"

# Hardcode; you said context is not needed. Keep if API requires it, else we omit.
DEFAULT_CONTEXT: Optional[str] = None  # set to "Male" if SciSports requires the param


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
# HTTP helpers
# -----------------------------
def _http_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"Accept": "application/json"})
    return s


def _auth_headers(access_token: str) -> Dict[str, str]:
    return {"Authorization": f"Bearer {access_token}"}


def _safe_get(d: Any, path: str, default: Any = None) -> Any:
    cur = d
    for part in path.split("."):
        if not isinstance(cur, dict) or part not in cur:
            return default
        cur = cur[part]
    return cur


def _as_text(value: Any) -> str:
    return "" if value is None else str(value)


def _fmt_int(value: Any) -> str:
    try:
        return "" if value is None else str(int(value))
    except Exception:
        return ""


def _fmt_money_eur(value: Any) -> str:
    try:
        if value is None:
            return ""
        v = float(value)
        if abs(v) >= 1_000_000:
            return f"€ {v/1_000_000:.2f}M"
        if abs(v) >= 1_000:
            return f"€ {v/1_000:.0f}K"
        return f"€ {v:.0f}"
    except Exception:
        return ""


def _parse_iso_date_to_ddmmyyyy(value: Optional[str]) -> str:
    if not value:
        return ""
    # Handles "YYYY-MM-DD" and ISO datetimes
    try:
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
        return dt.strftime("%d-%m-%Y")
    except Exception:
        try:
            dt = datetime.strptime(value[:10], "%Y-%m-%d")
            return dt.strftime("%d-%m-%Y")
        except Exception:
            return value


def _first_position(info: Dict[str, Any]) -> str:
    positions = info.get("positions") or []
    if isinstance(positions, list) and positions:
        return _as_text(positions[0])
    return ""


# -----------------------------
# Secrets + token
# -----------------------------
def _require_secrets() -> Dict[str, str]:
    required = [
        "SCISPORTS_USERNAME",
        "SCISPORTS_PASSWORD",
        "SCISPORTS_CLIENT_ID",
        "SCISPORTS_CLIENT_SECRET",
    ]
    missing = [k for k in required if not st.secrets.get(k)]
    if missing:
        st.error(
            "Missing Streamlit Secrets. Add these keys in Streamlit Cloud → App → Settings → Secrets:\n"
            + "\n".join([f"- {k}" for k in missing])
        )
        st.stop()

    return {
        "username": st.secrets["SCISPORTS_USERNAME"],
        "password": st.secrets["SCISPORTS_PASSWORD"],
        "client_id": st.secrets["SCISPORTS_CLIENT_ID"],
        "client_secret": st.secrets["SCISPORTS_CLIENT_SECRET"],
        "scope": st.secrets.get("SCISPORTS_SCOPE", "api recruitment"),
    }


def token_password_grant_from_secrets(timeout_s: int = 30) -> str:
    creds = _require_secrets()
    payload = {
        "grant_type": "password",
        "username": creds["username"],
        "password": creds["password"],
        "client_id": creds["client_id"],
        "client_secret": creds["client_secret"],
        "scope": creds["scope"],
    }
    resp = requests.post(TOKEN_URL, data=payload, timeout=timeout_s)
    resp.raise_for_status()
    data = resp.json()
    token = data.get("access_token")
    if not token:
        raise RuntimeError(f"Token response missing access_token: {json.dumps(data)[:500]}")
    return token


# -----------------------------
# SciSports API
# -----------------------------
@st.cache_data(show_spinner=False, ttl=60 * 15)
def search_players(
    access_token: str,
    search_text: str,
    limit: int,
    offset: int,
) -> Tuple[int, List[PlayerOption]]:
    s = _http_session()
    params: Dict[str, Any] = {"offset": offset, "limit": limit}
    if DEFAULT_CONTEXT:
        params["context"] = DEFAULT_CONTEXT
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
        pid = info.get("id")
        if pid is None:
            continue

        options.append(
            PlayerOption(
                player_id=int(pid),
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


def get_latest_transfer_fee(access_token: str, player_id: int) -> Optional[Dict[str, Any]]:
    s = _http_session()
    url = f"{API_BASE}/v2/metrics/players/transfer-fees"
    params: Dict[str, Any] = {
        "offset": 0,
        "limit": 1,
        "playerIds": player_id,
        "latestTransferFee": "true",
    }
    if DEFAULT_CONTEXT:
        params["context"] = DEFAULT_CONTEXT
    resp = s.get(url, headers=_auth_headers(access_token), params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    items = data.get("items") or []
    return items[0] if items else None


def get_career_stats_all(access_token: str, player_id: int) -> List[Dict[str, Any]]:
    """
    Fetch all career-stats items for a player (paginates).
    We then derive:
      - "this season": latest season by season.startDate
      - "career": sum across all returned seasons
    """
    s = _http_session()
    offset = 0
    limit = 100
    items: List[Dict[str, Any]] = []

    candidates = [
        f"{API_BASE}/v2/metrics/career-stats/players",
        f"{API_BASE}/v2/metrics/career-stats/Players",
    ]
    url = None
    for c in candidates:
        try:
            test = s.get(
                c,
                headers=_auth_headers(access_token),
                params={"offset": 0, "limit": 1, "playerIds": player_id},
                timeout=30,
            )
            if test.status_code < 400:
                url = c
                break
        except Exception:
            pass
    if not url:
        raise RuntimeError("Could not reach career-stats endpoint.")

    while True:
        params: Dict[str, Any] = {"offset": offset, "limit": limit, "playerIds": player_id}
        if DEFAULT_CONTEXT:
            params["context"] = DEFAULT_CONTEXT
        resp = s.get(url, headers=_auth_headers(access_token), params=params, timeout=30)
        resp.raise_for_status()
        payload = resp.json()
        batch = payload.get("items") or []
        items.extend(batch)

        total = int(payload.get("total", len(items)))
        offset += limit
        if len(items) >= total or not batch:
            break

    return items


def summarize_season_and_career(items: List[Dict[str, Any]]) -> Tuple[Dict[str, int], Dict[str, int], Optional[Dict[str, Any]]]:
    """
    Returns:
      (latest_season_stats, career_total_stats, latest_item)

    Keys:
      matches, minutes, goals, assists
    """
    def extract_stats(it: Dict[str, Any]) -> Dict[str, int]:
        stats = it.get("stats") or {}
        return {
            "matches": int(stats.get("matchesPlayed", 0) or 0),
            "minutes": int(stats.get("minutesPlayed", 0) or 0),
            "goals": int(stats.get("goal", 0) or 0),
            "assists": int(stats.get("assist", 0) or 0),
        }

    def season_sort_key(it: Dict[str, Any]) -> str:
        season = it.get("season") or {}
        sd = season.get("startDate")
        return sd if isinstance(sd, str) else ""

    if not items:
        z = {"matches": 0, "minutes": 0, "goals": 0, "assists": 0}
        return z, z, None

    latest_item = sorted(items, key=season_sort_key)[-1]
    latest = extract_stats(latest_item)

    career = {"matches": 0, "minutes": 0, "goals": 0, "assists": 0}
    for it in items:
        s = extract_stats(it)
        for k in career:
            career[k] += s[k]

    return latest, career, latest_item


# -----------------------------
# Agency / agent extraction (best-effort)
# -----------------------------
def extract_agent_and_agency(player_obj: Dict[str, Any]) -> Tuple[str, str]:
    """
    SciSports fields may differ per contract model; try a few common patterns.
    Returns (agency, agent).
    """
    contract = player_obj.get("contract") or {}
    info = player_obj.get("info") or {}

    # Common guesses
    agency = (
        _as_text(contract.get("agencyName"))
        or _as_text(_safe_get(contract, "agency.name", ""))
        or _as_text(info.get("agencyName"))
        or _as_text(_safe_get(info, "agency.name", ""))
    )
    agent = (
        _as_text(contract.get("agentName"))
        or _as_text(_safe_get(contract, "agent.name", ""))
        or _as_text(info.get("agentName"))
        or _as_text(_safe_get(info, "agent.name", ""))
    )

    return agency.strip(), agent.strip()


# -----------------------------
# PPTX placeholder replacement (run-safe)
# -----------------------------
TOKEN_RE = re.compile(r"\{[^{}]+\}")

def _replace_tokens_in_shape(shape, replacements: Dict[str, str]) -> bool:
    if not getattr(shape, "has_text_frame", False):
        return False

    changed = False
    for paragraph in shape.text_frame.paragraphs:
        if not paragraph.runs:
            continue

        # 1) run-level replace
        for run in paragraph.runs:
            t = run.text or ""
            if "{" not in t:
                continue
            new_t = t
            for k, v in replacements.items():
                if k in new_t:
                    new_t = new_t.replace(k, v)
            if new_t != t:
                run.text = new_t
                changed = True

        # 2) cross-run replace (token split across runs)
        combined = "".join((r.text or "") for r in paragraph.runs)
        if "{" not in combined:
            continue

        new_combined = combined
        for k, v in replacements.items():
            if k in new_combined:
                new_combined = new_combined.replace(k, v)

        if new_combined != combined:
            paragraph.runs[0].text = new_combined
            for r in paragraph.runs[1:]:
                r.text = ""
            changed = True

    return changed


def replace_placeholders_in_pptx(template_path: str, output_path: str, replacements: Dict[str, str]) -> None:
    prs = Presentation(template_path)

    for slide in prs.slides:
        for shape in slide.shapes:
            _replace_tokens_in_shape(shape, replacements)

    prs.save(output_path)


# -----------------------------
# Position coloring (bottom-left 1..11)
# -----------------------------
MAIN_BLUE = RGBColor(0, 83, 159)
SECOND_BLUE = RGBColor(0, 142, 204)

POSITION_TO_NUMBER: Dict[str, int] = {
    "Goalkeeper": 1,
    "RightBack": 2,
    "RightFullback": 2,
    "Right Wing": 7,
    "RightWing": 7,
    "LeftBack": 5,
    "Left Wing": 11,
    "LeftWing": 11,
    "CentreMidfield": 8,
    "AttackingMidfield": 10,
    "DefensiveMidfield": 6,
    "CentreForward": 9,
    # CentreBack handled dynamically (3/4)
    "CentreBack": -1,
}

def _resolve_position_number(position: str, preferred_foot: str) -> Optional[int]:
    if position in {"CentreBack", "Centre Back"}:
        pf = (preferred_foot or "").strip().lower()
        is_left = "left" in pf or pf.startswith("l")
        return 4 if is_left else 3
    if position in POSITION_TO_NUMBER and POSITION_TO_NUMBER[position] != -1:
        return POSITION_TO_NUMBER[position]
    return None


def apply_position_coloring(
    prs: Presentation,
    slide,
    ordered_positions: List[str],
    preferred_foot: str,
) -> None:
    if not ordered_positions:
        return

    main = _resolve_position_number(ordered_positions[0], preferred_foot)
    secondary: List[int] = []
    for p in ordered_positions[1:3]:
        n = _resolve_position_number(p, preferred_foot)
        if n is not None:
            secondary.append(n)

    if main is None and not secondary:
        return

    W = prs.slide_width
    H = prs.slide_height

    # bottom-left area (tuned to your template screenshot)
    x_max = int(W * 0.50)
    y_min = int(H * 0.55)

    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        if not (shape.left <= x_max and shape.top >= y_min):
            continue

        txt = (shape.text_frame.text or "").strip()
        if not txt.isdigit():
            continue

        num = int(txt)
        if not (1 <= num <= 11):
            continue

        if main is not None and num == main:
            shape.fill.solid()
            shape.fill.fore_color.rgb = MAIN_BLUE
        elif num in secondary:
            shape.fill.solid()
            shape.fill.fore_color.rgb = SECOND_BLUE


# -----------------------------
# Replacement mapping
# -----------------------------
def build_replacements(
    player: Dict[str, Any],
    transfer_fee: Optional[Dict[str, Any]],
    season_stats: Dict[str, int],
    career_stats: Dict[str, int],
) -> Dict[str, str]:
    info = player.get("info") or {}
    team = player.get("team") or {}
    league = player.get("league") or {}
    contract = player.get("contract") or {}

    name = _as_text(info.get("footballName") or info.get("name") or "")
    dob = _parse_iso_date_to_ddmmyyyy(_as_text(info.get("birthDate")))
    place = _as_text(info.get("birthPlace") or "")

    nats = info.get("nationalities") or []
    nat_names = [str(n.get("name", "")).strip() for n in nats if isinstance(n, dict) and n.get("name")]
    nationalities = ", ".join([n for n in nat_names if n])

    height = _fmt_int(info.get("height"))
    preferred_foot = _as_text(info.get("preferredFoot") or "")
    position = _first_position(info)
    league_name = _as_text(league.get("name") or "")
    club_name = _as_text(team.get("name") or "")

    contract_end = _parse_iso_date_to_ddmmyyyy(_as_text(contract.get("contractEnd")))
    market_value = _fmt_money_eur(contract.get("marketValue"))
    tv = _fmt_money_eur(transfer_fee.get("valueEstimateEur")) if transfer_fee else ""

    agency, agent = extract_agent_and_agency(player)

    # Provide multiple key variants to tolerate spaces in tokens
    def variants(token: str) -> List[str]:
        base = token.strip()
        inner = base.strip("{}").strip()
        return list({f"{{{inner}}}", f"{{ {inner} }}", base})

    mapping_pairs = [
        ("{Name}", name),
        ("{ DD/MM/YYYY }", dob),        # date of birth in your template
        ("{Place}", place),
        ("{ Place }", place),
        ("{Nationalities}", nationalities),
        ("{Country}", nationalities),    # if you ever used {Country}
        ("{ Height }", height),
        ("{Preferred foot}", preferred_foot),
        ("{ Preferred Foot }", preferred_foot),
        ("{Position}", position),
        ("{ Position }", position),
        ("{League}", league_name),
        ("{ League }", league_name),
        ("{Club}", club_name),
        ("{ Club }", club_name),
        ("{Season_m}", _fmt_int(season_stats.get("matches"))),
        ("{season_min}", _fmt_int(season_stats.get("minutes"))),
        ("{season_g}", _fmt_int(season_stats.get("goals"))),
        ("{season_a}", _fmt_int(season_stats.get("assists"))),
        ("{Career_m}", _fmt_int(career_stats.get("matches"))),
        ("{career_min}", _fmt_int(career_stats.get("minutes"))),
        ("{career_g}", _fmt_int(career_stats.get("goals"))),
        ("{career_a}", _fmt_int(career_stats.get("assists"))),
        ("{con_DD/MM/YYYY}", contract_end),
        ("{TV}", tv),
        ("{MV}", market_value),
        ("{Agency}", agency),
        ("{Agent}", agent),
    ]

    out: Dict[str, str] = {}
    for k, v in mapping_pairs:
        for kk in variants(k):
            out[kk] = v

    return out


# -----------------------------
# UI
# -----------------------------
def render_player_card(p: PlayerOption) -> None:
    age_str = "?" if p.age is None else str(p.age)
    st.markdown(
        """
<style>
.player-card {
  border: 1px solid rgba(49, 51, 63, 0.2);
  border-radius: 10px;
  padding: 14px 16px;
  background: rgba(255,255,255,0.02);
}
.player-name { font-size: 16px; font-weight: 700; line-height: 1.2; }
.player-meta { margin-top: 6px; font-size: 13px; opacity: 0.85; }
</style>
""",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"""
<div class="player-card">
  <div class="player-name">{p.name}</div>
  <div class="player-meta">{age_str} - {p.position or "UnknownPos"} - {p.club or "UnknownClub"} ({p.league or "Unknown"})</div>
</div>
""",
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(page_title="SciSports Scouting Form Generator", layout="centered")
    st.title("SciSports Scouting Form Generator")

    with st.expander("1) Generate API key (access token)", expanded=True):
        if st.button("Generate API key", type="primary"):
            try:
                token = token_password_grant_from_secrets()
                st.session_state["SCISPORTS_ACCESS_TOKEN"] = token
                os.environ["SCISPORTS_ACCESS_TOKEN"] = token
                st.success("Access token generated and stored for this session.")
            except Exception as e:
                st.error(f"Failed to generate token: {e}")

        if st.session_state.get("SCISPORTS_ACCESS_TOKEN"):
            st.info("Token present ✅ You can now search and select a player.")

    token = st.session_state.get("SCISPORTS_ACCESS_TOKEN", "")
    st.divider()

    st.subheader("2) Select player")
    if not token:
        st.warning("Generate the API key first.")
        st.stop()

    q = st.text_input("Search player name", placeholder="e.g. Ibrahim El Kadiri")
    c1, c2, c3 = st.columns([1, 1, 2])
    limit = c1.selectbox("Limit", options=[25, 50, 100], index=1)
    offset = c2.number_input("Offset", min_value=0, value=0, step=int(limit))

    if c3.button("Search", type="secondary") or "player_options" not in st.session_state:
        total, options = search_players(token, q, int(limit), int(offset))
        st.session_state["player_total"] = total
        st.session_state["player_options"] = options

    total = int(st.session_state.get("player_total", 0))
    options: List[PlayerOption] = st.session_state.get("player_options", [])
    st.caption(f"Results: showing {len(options)} of total {total} (use Offset/Limit to page)")

    if not options:
        st.warning("No players found. Try another search.")
        st.stop()

    selected_label = st.selectbox("Player", options=[o.label() for o in options], index=0)
    selected = next(o for o in options if o.label() == selected_label)
    st.session_state["selected_player_id"] = selected.player_id

    render_player_card(selected)

    st.divider()
    st.subheader("3) Generate Scoutings Form")

    if not os.path.exists(TEMPLATE_PATH):
        st.error(f"Template not found: {TEMPLATE_PATH}. Commit it into the repo root.")
        st.stop()

    if st.button("Generate Scoutings Form", type="primary"):
        with st.spinner("Fetching data and generating PPTX..."):
            try:
                player = get_player(token, selected.player_id)
                transfer_fee = get_latest_transfer_fee(token, selected.player_id)

                career_items = get_career_stats_all(token, selected.player_id)
                season_stats, career_stats, latest_item = summarize_season_and_career(career_items)

                replacements = build_replacements(
                    player=player,
                    transfer_fee=transfer_fee,
                    season_stats=season_stats,
                    career_stats=career_stats,
                )

                prs = Presentation(TEMPLATE_PATH)

                ordered_positions = (player.get("info") or {}).get("positions") or []
                preferred_foot = _as_text(_safe_get(player, "info.preferredFoot", ""))

                for slide in prs.slides:
                    for shape in slide.shapes:
                        _replace_tokens_in_shape(shape, replacements)

                    apply_position_coloring(
                        prs=prs,
                        slide=slide,
                        ordered_positions=[_as_text(p) for p in ordered_positions if p],
                        preferred_foot=preferred_foot,
                    )

                with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
                    out_path = tmp.name
                prs.save(out_path)

                with open(out_path, "rb") as f:
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

                # Helpful debug if Agency/Agent are empty
                if not replacements.get("{Agency}") and not replacements.get("{Agent}"):
                    st.info(
                        "Agency/Agent came back empty. If you can point me to the exact SciSports field/endpoint "
                        "for representation, I’ll wire it in."
                    )

            except Exception as e:
                st.error(f"Failed to generate PPTX: {e}")


if __name__ == "__main__":
    main()
