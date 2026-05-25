import os
import requests
import streamlit as st
import pandas as pd
import io
import Xlsxwriter

from datetime import date, timedelta, datetime
from requests.auth import HTTPBasicAuth
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Optional


# ======================
# STREAMLIT CONFIG
# ======================

st.set_page_config(
    page_title="Jira Worklog Dashboard",
    layout="wide"
)

st.title("Jira Worklog Dashboard")


# ======================
# BASIC AUTH
# ======================

APP_USERNAME = st.secrets["APP_USERNAME"]
APP_PASSWORD = st.secrets["APP_PASSWORD"]


def check_login():

    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    st.subheader("Login")

    with st.form("login_form"):

        username = st.text_input("Username")
        password = st.text_input("Password", type="password")

        submit = st.form_submit_button("Accedi")

        if submit:

            if (
                username == APP_USERNAME
                and password == APP_PASSWORD
            ):

                st.session_state.authenticated = True
                st.rerun()

            else:
                st.error("Credenziali non valide")

    return False


if not check_login():
    st.stop()


# ======================
# LOGOUT
# ======================

with st.sidebar:

    st.success("Autenticato")

    if st.button("Logout"):

        st.session_state.authenticated = False
        st.rerun()


# ======================
# CONFIG (ENV VARS)
# ======================

jira_domain = st.secrets["JIRA_DOMAIN"]
email = st.secrets["JIRA_EMAIL"]
api_token = st.secrets["JIRA_API_TOKEN"]

default_jql = "project = KAN"

if not jira_domain or not email or not api_token:
    st.error("Imposta JIRA_DOMAIN JIRA_EMAIL JIRA_API_TOKEN")
    st.stop()

BASE_URL = f"https://{jira_domain}/rest/api/3"

AUTH = HTTPBasicAuth(email, api_token)

HEADERS_JSON = {
    "Accept": "application/json",
    "Content-Type": "application/json"
}

HEADERS_GET = {
    "Accept": "application/json"
}

MAX_WORKERS = 10
MARGIN_DAYS = 3

TTL_SEARCH = 30 * 60
TTL_WORKLOG = 60 * 60
TTL_EPIC = 6 * 60 * 60


# ======================
# JIRA API
# ======================

def search_issues_jql_v3(base_url, auth, jql, fields=None):

    if fields is None:
        fields = ["summary", "issuetype"]

    url = f"{base_url}/search/jql"

    issues = []
    next_page_token = None

    while True:

        payload = {
            "jql": jql,
            "fields": fields
        }

        if next_page_token:
            payload["nextPageToken"] = next_page_token

        resp = requests.post(
            url,
            json=payload,
            headers=HEADERS_JSON,
            auth=auth,
            timeout=60,
        )

        if not resp.ok:
            raise RuntimeError(resp.text)

        data = resp.json()

        issues.extend(data.get("issues", []))

        next_page_token = data.get("nextPageToken")

        if not next_page_token:
            break

    return issues


def get_issue_worklogs_v3(base_url, auth, issue_key):

    url = f"{base_url}/issue/{issue_key}/worklog"

    start_at = 0
    max_results = 100

    out = []

    while True:

        params = {
            "startAt": start_at,
            "maxResults": max_results
        }

        resp = requests.get(
            url,
            params=params,
            headers=HEADERS_GET,
            auth=auth,
            timeout=60,
        )

        if not resp.ok:
            raise RuntimeError(resp.text)

        data = resp.json()

        wls = data.get("worklogs", [])

        out.extend(wls)

        start_at += len(wls)

        if start_at >= data.get("total", 0):
            break

    return out


# ======================
# DETECT EPIC FIELD
# ======================

@st.cache_data(ttl=24 * 60 * 60)
def detect_epic_link_field():

    url = f"{BASE_URL}/field"

    resp = requests.get(
        url,
        headers=HEADERS_GET,
        auth=AUTH
    )

    if not resp.ok:
        return None

    for f in resp.json():

        if (f.get("name") or "").lower() == "epic link":
            return f.get("id")

    return None


EPIC_LINK_FIELD_ID = detect_epic_link_field()


# ======================
# SIDEBAR
# ======================

st.sidebar.header("Filtri")

today = date.today()

date_from = st.sidebar.date_input(
    "Dal",
    today - timedelta(days=7)
)

date_to = st.sidebar.date_input(
    "Al",
    today
)

if date_from > date_to:
    st.sidebar.error("Range non valido")
    st.stop()

refresh = st.sidebar.button("Aggiorna cache")

if refresh:
    st.cache_data.clear()

pref_from = date_from - timedelta(days=MARGIN_DAYS)

jql_effective = (
    f"({default_jql}) "
    f'AND updated >= "{pref_from.isoformat()}"'
)


# ======================
# CACHE
# ======================

@st.cache_data(ttl=TTL_SEARCH)
def cached_search_issues(jql):

    fields = [
        "summary",
        "issuetype",
        "timetracking",
        "status",
        "assignee",
        "parent",
    ]

    if EPIC_LINK_FIELD_ID:
        fields.append(EPIC_LINK_FIELD_ID)

    return search_issues_jql_v3(
        BASE_URL,
        AUTH,
        jql,
        fields
    )


@st.cache_data(ttl=TTL_WORKLOG)
def cached_issue_worklogs(issue_key):

    return get_issue_worklogs_v3(
        BASE_URL,
        AUTH,
        issue_key
    )


@st.cache_data(ttl=TTL_EPIC)
def cached_issue_summary(issue_key):

    url = f"{BASE_URL}/issue/{issue_key}"

    resp = requests.get(
        url,
        params={"fields": "summary"},
        headers=HEADERS_GET,
        auth=AUTH,
        timeout=60,
    )

    if not resp.ok:
        return ""

    data = resp.json() or {}

    fields = data.get("fields", {}) or {}

    return fields.get("summary", "") or ""


# ======================
# HELPERS
# ======================

def estimate_hours(fields):

    tt = fields.get("timetracking") or {}

    sec = tt.get("originalEstimateSeconds")

    if sec is None:
        sec = fields.get("timeoriginalestimate")

    return round((sec or 0) / 3600, 2)


def epic_key_from_issue(fields):

    if EPIC_LINK_FIELD_ID:

        v = fields.get(EPIC_LINK_FIELD_ID)

        if isinstance(v, str) and v.strip():
            return v

    parent = fields.get("parent") or {}

    return parent.get("key", "")


# ======================
# BUILD DATAFRAME
# ======================

def build_dataframe(issues):

    rows = []
    epic_keys = set()

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:

        futures = {
            ex.submit(cached_issue_worklogs, i["key"]): i
            for i in issues
        }

        for fut in as_completed(futures):

            issue = futures[fut]

            fields = issue.get("fields") or {}

            key = issue.get("key")

            summary = fields.get("summary", "")

            issue_type = (
                fields.get("issuetype") or {}
            ).get("name", "")

            status = (
                fields.get("status") or {}
            ).get("name", "")

            owner = (
                fields.get("assignee") or {}
            ).get("displayName", "")

            est = estimate_hours(fields)

            epic = epic_key_from_issue(fields)

            epic_keys.add(epic)

            worklogs = fut.result()

            for wl in worklogs:

                started = wl.get("started")

                if not started:
                    continue

                wl_day = datetime.strptime(
                    started[:10],
                    "%Y-%m-%d"
                ).date()

                if wl_day < date_from or wl_day > date_to:
                    continue

                rows.append(
                    {
                        "Data": wl_day,
                        "Utente": (
                            wl.get("author") or {}
                        ).get("displayName", ""),

                        "IssueType": issue_type,
                        "Issue": key,
                        "Summary": summary,
                        "Owner": owner,
                        "EpicKey": epic,
                        "Stato": status,
                        "StimaOre": est,

                        "Ore": round(
                            (
                                wl.get(
                                    "timeSpentSeconds",
                                    0
                                )
                            ) / 3600,
                            2
                        ),
                    }
                )

    df = pd.DataFrame(rows)

    if df.empty:
        return df

    epic_map = {
        k: cached_issue_summary(k)
        for k in epic_keys
        if str(k).strip()
    }

    df["EpicName"] = (
        df["EpicKey"]
        .map(epic_map)
        .fillna("")
    )

    return df


def create_excel_export(df):

    cols_export = [
        "Data",
        "Utente",
        "IssueType",
        "Issue",
        "Summary",
        "EpicKey",
        "EpicName",
        "StimaOre",
        "Ore",
        "Stato",
    ]

    df_export = df.copy()

    df_export["Data"] = pd.to_datetime(
        df_export["Data"]
    )

    df_export["StimaOre"] = (
        df_export["StimaOre"]
        .astype(float)
    )

    df_export["Ore"] = (
        df_export["Ore"]
        .astype(float)
    )

    df_export = df_export[cols_export]

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter"
    ) as writer:

        df_export.to_excel(
            writer,
            index=False,
            sheet_name="Worklog"
        )

        workbook = writer.book
        worksheet = writer.sheets["Worklog"]

        date_format = workbook.add_format(
            {"num_format": "dd/mm/yyyy"}
        )

        number_format = workbook.add_format(
            {"num_format": "0.00"}
        )

        worksheet.set_column(
            "A:A",
            12,
            date_format
        )

        worksheet.set_column(
            "H:I",
            12,
            number_format
        )

        worksheet.set_column("B:B", 25)
        worksheet.set_column("E:E", 50)

    output.seek(0)

    return output


# ======================
# LOAD DATA
# ======================

with st.spinner("Ricerca issue..."):
    issues = cached_search_issues(jql_effective)

if not issues:
    st.info("Nessuna issue trovata")
    st.stop()

with st.spinner("Caricamento worklog..."):
    df = build_dataframe(issues)

if df.empty:
    st.info("Nessun worklog")
    st.stop()


# ======================
# FILTRI UI
# ======================

statuses = ["(tutti)"] + sorted(
    df["Stato"].dropna().unique()
)

status_sel = st.sidebar.selectbox(
    "Stato",
    statuses
)

types = ["(tutti)"] + sorted(
    df["IssueType"].dropna().unique()
)

type_sel = st.sidebar.selectbox(
    "Issue Type",
    types
)

if df["EpicName"].astype(str).str.strip().any():

    epics = ["(tutte)"] + sorted(
        df["EpicName"].dropna().unique()
    )

    epic_sel = st.sidebar.selectbox(
        "Epic",
        epics
    )

else:

    epics = ["(tutte)"] + sorted(
        df["EpicKey"].dropna().unique()
    )

    epic_sel = st.sidebar.selectbox(
        "Epic (key)",
        epics
    )

users = ["(tutti)"] + sorted(
    df["Utente"].dropna().unique()
)

user_sel = st.sidebar.selectbox(
    "Utente",
    users
)

df_view = df.copy()

if status_sel != "(tutti)":
    df_view = df_view[
        df_view["Stato"] == status_sel
    ]

if type_sel != "(tutti)":
    df_view = df_view[
        df_view["IssueType"] == type_sel
    ]

if epic_sel != "(tutte)":

    if "EpicName" in df_view.columns:

        df_view = df_view[
            df_view["EpicName"] == epic_sel
        ]

    else:

        df_view = df_view[
            df_view["EpicKey"] == epic_sel
        ]

if user_sel != "(tutti)":

    df_view = df_view[
        df_view["Utente"] == user_sel
    ]


# ======================
# KPI
# ======================

c1, c2, c3, c4 = st.columns(4)

c1.metric(
    "Totale ore",
    f"{df_view['Ore'].sum():.2f}"
)

c2.metric(
    "Worklog",
    len(df_view)
)

c3.metric(
    "Issue",
    df_view["Issue"].nunique()
)

c4.metric(
    "Utenti",
    df_view["Utente"].nunique()
)

st.divider()


# ======================
# DETTAGLIO
# ======================

st.subheader("Dettaglio worklog")

df_show = df_view.copy()

df_show["Data"] = pd.to_datetime(
    df_show["Data"]
).dt.strftime("%d/%m/%Y")

st.dataframe(
    df_show,
    use_container_width=True,
    hide_index=True
)

excel_file = create_excel_export(df_view)

st.download_button(
    label="📥 Scarica Excel",
    data=excel_file,
    file_name="jira_worklog.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
