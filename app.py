# app.py
"""
Streamlit app:
- Uses Streamlit Secrets for SciSports credentials (no credential inputs in UI)
- Generates SciSports token
- Search + select a player
- Generate filled PPTX by replacing placeholders in TemplateScoutingsRapport.pptx
"""

from __future__ import annotations

import json
import os
import re
import tempfile
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any, Dict, List, Optional, Tuple

import requests
import streamlit as st
from pptx import Presentation

API_BASE = "https://api-recruitment.scisports.app/api"
TOKEN_URL = "https://identity.scisports.app/connect/token"
TEMPLATE_PATH = "TemplateScoutingsRapport.pptx"


@dataclass(frozen=True)
class PlayerOption:
    player_id: int
    name: str
    age: Optional[int]
    position: str
    club: str
    league: str
    country: str

    def dropdown_label(self) -> str:
        age_str = "?" if self.age is None else str(self.age)
        pos = self.position or "UnknownPos"
        club = self.club or "UnknownClub"
        league = self.league or "UnknownLeague"
        return f"{self.name} | {age_str} | {pos} | {club} ({league})"


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


def _as_text(value: Any) -> str:
    return "" if value is None else str(value)


def _fmt_int(value: Any) -> str:
    try:
        return "" if value is None else f"{int(value)}"
    except Exception:
        return ""


def _fmt_money_eur(value: Any) -> str:
    try:
        if value is None:
            return ""
        v = float(value)
        if abs(v) >= 1_000_000:
            return f"€{v/1_000_000:.2f}M"
        if abs(v) >= 1_000:
            return f"€{v/1_000:.1f}K"
        return f"€{v:.0f}"
    except Exception:
        return ""


def _iso_to_ddmmyyyy(value: Optional[str]) -> str:
    if not value:
        return ""
    try:
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
        return dt.strftime("%d/%m/%Y")
    except Exception:
        try:
            dt = datetime.strptime(value[:10], "%Y-%m-%d")
            return dt.strftime("%d/%m/%Y")
        except Exception:
            return value


def _first_position(info: Dict[str, Any]) -> str:
    positions = info.get("positions") or []
    if isinstance(positions, list) and positions:
        return _as_text(positions[0])
    return ""


def _require_secrets() -> Dict[str, str]:
    """
    Reads required SciSports secrets from Streamlit Secrets.
    Raises a Streamlit error if missing.
    """
    required = [
        "SCISPORTS_USERNAME",
        "SCISPORTS_PASSWORD",
        "SCISPORTS_CLIENT_ID",
        "SCISPORTS_CLIENT_SECRET",
    ]
    missing = [k for k in required if not st.secrets.get(k)]
    if missing:
        st.error(
            "Missing Streamlit Secrets. Please add these keys in Streamlit Cloud → App → Settings → Secrets:\n"
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


@st.cache_data(show_spinner=False, ttl=60 * 15)
def search_players(
    access_token: str,
    search_text: str,
    limit: int,
    offset: int,
    context: str,
) -> Tuple[int, List[PlayerOption]]:
    s = _http_session()
    params: Dict[str, Any] = {"offset": offset, "limit": limit, "context": context}
    if search_text.strip():
        params["searchText"] = search_text.strip()

    url = f"{API_BASE}/v2/players"
    resp = s.get(url, headers=_auth_headers(access_token), params=params, timeout=30)
    resp.raise_for_status()
    payload = resp.json()

    total = int(payload.get("total", 0))
    items = payload.get("items") or []

    out: List[PlayerOption] = []
    for it in items:
        info = it.get("info") or {}
        team = it.get("team") or {}
        league = it.get("league") or {}
        country = _as_text(_safe_get(league, "nation.name", "") or _safe_get(info, "birthCountry.name", "") or "")

        pid = info.get("id")
        if pid is None:
            continue

        out.append(
            PlayerOption(
                player_id=int(pid),
                name=_as_text(info.get("name") or info.get("footballName") or ""),
                age=info.get("age"),
                position=_first_position(info),
                club=_as_text(team.get("name") or ""),
                league=_as_text(league.get("name") or ""),
                country=country,
            )
        )

    return total, out


def get_player(access_token: str, player_id: int) -> Dict[str, Any]:
    s = _http_session()
    url = f"{API_BASE}/v2/players/{player_id}"
    resp = s.get(url, headers=_auth_headers(access_token), timeout=30)
    resp.raise_for_status()
    return resp.json()


def get_latest_transfer_fee(access_token: str, player_id: int, context: str) -> Optional[Dict[str, Any]]:
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


def get_career_stats_all(access_token: str, player_id: int, context: str) -> List[Dict[str, Any]]:
    s = _http_session()
    items: List[Dict[str, Any]] = []
    offset = 0
    limit = 50

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
                params={"offset": 0, "limit": 1, "playerIds": player_id, "context": context},
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
    def extract(it: Dict[str, Any]) -> Dict[str, int]:
        stats = it.get("stats") or {}
        return {
            "matches": int(stats.get("matchesPlayed", 0) or 0),
            "minutes": int(stats.get("minutesPlayed", 0) or 0),
            "goals": int(stats.get("goal", 0) or 0),
            "assists": int(stats.get("assist", 0) or 0),
        }

    def season_start(it: Dict[str, Any]) -> str:
        season = it.get("season") or {}
        sd = season.get("startDate")
        return sd if isinstance(sd, str) else ""

    if not items:
        z = {"matches": 0, "minutes": 0, "goals": 0, "assists": 0}
        return z, z

    latest = sorted(items, key=season_start)[-1]
    latest_stats = extract(latest)

    career = {"matches": 0, "minutes": 0, "goals": 0, "assists": 0}
    for it in items:
        s = extract(it)
        for k in career:
            career[k] += s[k]

    return latest_stats, career


def replace_placeholders_in_pptx(template_path: str, output_path: str, replacements: Dict[str, str]) -> None:
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
    transfer_fee: Optional[Dict[str, Any]],
    latest_season: Dict[str, int],
    career_total: Dict[str, int],
) -> Dict[str, str]:
    contract = (player.get("contract") or {})
    contract_end = _iso_to_ddmmyyyy(contract.get("contractEnd"))
    market_value = _fmt_money_eur(contract.get("marketValue"))
    tv = _fmt_money_eur(transfer_fee.get("valueEstimateEur")) if transfer_fee else ""

    return {
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
        "{Agency}": "",
        "{Agent}": "",
    }


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
  <div class="player-meta">{age_str} - {p.position or "UnknownPos"} - {p.club or "UnknownClub"} ({p.country or p.league or "Unknown"})</div>
</div>
""",
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(page_title="SciSports Scouting Form Generator", layout="centered")
    st.title("SciSports Scouting Form Generator")

    with st.expander("1) Generate API key (access token)", expanded=True):
        context = st.selectbox("Context", options=["Male", "Female"], index=0)

        if st.button("Generate API key", type="primary"):
            try:
                token = token_password_grant_from_secrets()
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

    q = st.text_input("Search player name", placeholder="e.g. Messi", value=st.session_state.get("player_search", ""))
    st.session_state["player_search"] = q

    a, b, c = st.columns([1, 1, 2])
    limit = a.selectbox("Limit", options=[25, 50, 100], index=1)
    offset = b.number_input("Offset", min_value=0, value=int(st.session_state.get("player_offset", 0)), step=int(limit))
    st.session_state["player_offset"] = int(offset)

    if c.button("Search", type="secondary") or "player_options" not in st.session_state:
        total, options = search_players(token, q, int(limit), int(offset), context)
        st.session_state["player_total"] = total
        st.session_state["player_options"] = options

    total = int(st.session_state.get("player_total", 0))
    options: List[PlayerOption] = st.session_state.get("player_options", [])
    st.caption(f"Results: showing {len(options)} of total {total} (use Offset/Limit to page)")

    if not options:
        st.warning("No players found. Try another search.")
        st.stop()

    labels = [o.dropdown_label() for o in options]
    selected_label = st.selectbox("Pick player", options=labels, index=0)
    selected = next(o for o in options if o.dropdown_label() == selected_label)
    st.session_state["selected_player_id"] = selected.player_id

    render_player_card(selected)

    st.divider()
    st.subheader("3) Generate Scoutings Form")

    if not os.path.exists(TEMPLATE_PATH):
        st.error(
            f"Template not found at {TEMPLATE_PATH}. "
            "Commit TemplateScoutingsRapport.pptx into the repo root."
        )
        st.stop()

    if st.button("Generate Scoutings Form", type="primary"):
        with st.spinner("Fetching data and generating PPTX..."):
            player = get_player(token, selected.player_id)
            transfer_fee = get_latest_transfer_fee(token, selected.player_id, context=context)
            career_items = get_career_stats_all(token, selected.player_id, context=context)
            latest_season, career_total = summarize_career_stats(career_items)

            replacements = build_replacements(player, transfer_fee, latest_season, career_total)

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


if __name__ == "__main__":
    main()
