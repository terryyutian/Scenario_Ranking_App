"""FastAPI app for ranking generated scenarios.

The app reads scenario pairs from an Excel file, presents three scenarios in a
randomized order, and stores each user's ranking decisions in SQLite.
"""

from __future__ import annotations

import ast
import io
import json
import random
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd
from fastapi import FastAPI, HTTPException, Request, Response
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates


BASE_DIR = Path(__file__).resolve().parent
SOURCE_FILE = BASE_DIR / "scenario_comparison_example_questions.xlsx"
DATABASE_FILE = BASE_DIR / "scenario_rankings.sqlite3"

REQUIRED_COLUMNS = [
    "pair_id",
    "gold_scenario",
    "topic",
    "target_attribute",
    "web_search_scenario",
    "baseline_scenario",
    "web_search_example_question",
    "baseline_example_question",
]

COGNITIVE_LEVEL_COLUMNS = ["cognitive_level", "congitive_level"]
SCENARIO_TYPES = ["gold", "web_search", "baseline"]

app = FastAPI(title="Scenario Ranking App")
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")
templates = Jinja2Templates(directory=BASE_DIR / "templates")


def get_db_connection() -> sqlite3.Connection:
    """Return a SQLite connection with row dictionaries enabled."""
    connection = sqlite3.connect(DATABASE_FILE)
    connection.row_factory = sqlite3.Row
    return connection


def init_db() -> None:
    """Create the SQLite tables if they do not already exist."""
    with get_db_connection() as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS rankings (
                username TEXT NOT NULL,
                pair_id TEXT NOT NULL,
                gold_scenario TEXT NOT NULL,
                topic TEXT NOT NULL,
                target_attribute TEXT NOT NULL,
                web_search_scenario TEXT NOT NULL,
                baseline_scenario TEXT NOT NULL,
                web_search_example_question TEXT NOT NULL,
                baseline_example_question TEXT NOT NULL,
                cognitive_level TEXT NOT NULL DEFAULT '',
                gold_score INTEGER NOT NULL,
                web_search_score INTEGER NOT NULL,
                baseline_score INTEGER NOT NULL,
                gold_comment TEXT NOT NULL DEFAULT '',
                web_search_comment TEXT NOT NULL DEFAULT '',
                baseline_comment TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL,
                PRIMARY KEY (username, pair_id)
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS user_pair_order (
                username TEXT NOT NULL,
                position INTEGER NOT NULL,
                pair_id TEXT NOT NULL,
                PRIMARY KEY (username, position),
                UNIQUE (username, pair_id)
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS user_scenario_order (
                username TEXT NOT NULL,
                pair_id TEXT NOT NULL,
                scenario_order TEXT NOT NULL,
                PRIMARY KEY (username, pair_id)
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS app_state (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
            """
        )
        ensure_column(connection, "rankings", "cognitive_level", "TEXT NOT NULL DEFAULT ''")
        ensure_column(connection, "rankings", "gold_comment", "TEXT NOT NULL DEFAULT ''")
        ensure_column(connection, "rankings", "web_search_comment", "TEXT NOT NULL DEFAULT ''")
        ensure_column(connection, "rankings", "baseline_comment", "TEXT NOT NULL DEFAULT ''")


def ensure_column(
    connection: sqlite3.Connection,
    table_name: str,
    column_name: str,
    column_definition: str,
) -> None:
    """Add a column to an existing SQLite table if it is missing."""
    existing_columns = {
        row["name"]
        for row in connection.execute(f"PRAGMA table_info({table_name})").fetchall()
    }
    if column_name not in existing_columns:
        connection.execute(
            f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_definition}"
        )


def cell_to_text(value: Any) -> str:
    """Convert an Excel cell value to clean text."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def load_source_records() -> list[dict[str, str]]:
    """Read the source Excel file and return one dictionary per row."""
    if not SOURCE_FILE.exists():
        raise FileNotFoundError(
            f"Could not find {SOURCE_FILE.name} in the app root directory."
        )

    data = pd.read_excel(SOURCE_FILE)
    data.columns = data.columns.astype(str).str.strip()

    missing_columns = [col for col in REQUIRED_COLUMNS if col not in data.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns: {missing_columns}")

    cognitive_level_column = next(
        (column for column in COGNITIVE_LEVEL_COLUMNS if column in data.columns),
        None,
    )

    records: list[dict[str, str]] = []
    for _, row in data.iterrows():
        record = {col: cell_to_text(row[col]) for col in REQUIRED_COLUMNS}
        record["cognitive_level"] = (
            cell_to_text(row[cognitive_level_column])
            if cognitive_level_column
            else ""
        )
        records.append(record)

    return records


def get_records_by_pair_id() -> dict[str, dict[str, str]]:
    """Return source records indexed by pair_id."""
    records = load_source_records()
    return {record["pair_id"]: record for record in records}


def parse_scenario(value: str) -> dict[str, Any]:
    """Parse a scenario stored as JSON or a Python-dictionary-like string."""
    if not value:
        return {}

    if isinstance(value, dict):
        return value

    text = str(value).strip()
    if not text:
        return {}

    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        try:
            parsed = ast.literal_eval(text)
        except (ValueError, SyntaxError):
            return {"scenario_text": text}

    return parsed if isinstance(parsed, dict) else {"scenario_text": text}


def normalize_data_table(data_table: Any) -> dict[str, Any] | None:
    """Prepare a scenario data table for template rendering."""
    if not isinstance(data_table, dict):
        return None

    headers = data_table.get("column_headers") or []
    rows = data_table.get("rows") or []
    title = cell_to_text(data_table.get("table_title", ""))
    notes = cell_to_text(data_table.get("notes", ""))

    if not headers and not rows and not title and not notes:
        return None

    normalized_rows = []
    max_values = 0
    has_row_labels = False

    for row in rows:
        if not isinstance(row, dict):
            continue

        row_label = cell_to_text(row.get("row_label", ""))
        values = [cell_to_text(value) for value in row.get("values", [])]
        max_values = max(max_values, len(values))
        has_row_labels = has_row_labels or bool(row_label)
        normalized_rows.append({"row_label": row_label, "values": values})

    if not headers and max_values:
        headers = [f"Column {index}" for index in range(1, max_values + 1)]

    return {
        "title": title,
        "headers": [cell_to_text(header) for header in headers],
        "rows": normalized_rows,
        "notes": notes,
        "has_row_labels": has_row_labels,
    }


def build_scenario_card(
    record: dict[str, str],
    scenario_type: str,
    saved_ranking: sqlite3.Row | None = None,
) -> dict[str, Any]:
    """Build one scenario card for display."""
    if scenario_type == "gold":
        raw_scenario = record["gold_scenario"]
        parsed = parse_scenario(raw_scenario)
        example_question = (
            cell_to_text(parsed.get("question", ""))
            or cell_to_text(parsed.get("example_question", ""))
        )
    elif scenario_type == "web_search":
        raw_scenario = record["web_search_scenario"]
        parsed = parse_scenario(raw_scenario)
        example_question = record["web_search_example_question"]
    elif scenario_type == "baseline":
        raw_scenario = record["baseline_scenario"]
        parsed = parse_scenario(raw_scenario)
        example_question = record["baseline_example_question"]
    else:
        raise ValueError(f"Unknown scenario type: {scenario_type}")

    return {
        "type": scenario_type,
        "scenario_text": cell_to_text(parsed.get("scenario_text", "")),
        "data_table": normalize_data_table(parsed.get("data_table")),
        "formula": cell_to_text(parsed.get("formula", "")),
        "example_question": example_question,
        "comment": saved_ranking[f"{scenario_type}_comment"] if saved_ranking else "",
    }


def get_username(request: Request) -> str | None:
    """Read the current username from the browser cookie."""
    username = request.cookies.get("username")
    if username:
        return username.strip()
    return None


def require_username(request: Request) -> str:
    """Return the current username or redirect the user to log in."""
    username = get_username(request)
    if not username:
        raise HTTPException(status_code=303, headers={"Location": "/"})
    return username


def get_last_username() -> str | None:
    """Return the most recent username stored by the app."""
    with get_db_connection() as connection:
        row = connection.execute(
            "SELECT value FROM app_state WHERE key = 'last_username'"
        ).fetchone()
        if row:
            return row["value"]

        row = connection.execute(
            """
            SELECT username
            FROM rankings
            ORDER BY updated_at DESC
            LIMIT 1
            """
        ).fetchone()
        if row:
            return row["username"]

        row = connection.execute(
            """
            SELECT username
            FROM user_pair_order
            ORDER BY rowid DESC
            LIMIT 1
            """
        ).fetchone()
        return row["username"] if row else None


def set_last_username(username: str) -> None:
    """Store the most recent username in SQLite."""
    with get_db_connection() as connection:
        connection.execute(
            """
            INSERT INTO app_state (key, value)
            VALUES ('last_username', ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            (username,),
        )


def get_or_create_pair_order(username: str, pair_ids: list[str]) -> list[str]:
    """Return this user's fixed randomized pair order."""
    with get_db_connection() as connection:
        rows = connection.execute(
            """
            SELECT pair_id
            FROM user_pair_order
            WHERE username = ?
            ORDER BY position
            """,
            (username,),
        ).fetchall()

        existing_order = [row["pair_id"] for row in rows]
        current_pair_ids = set(pair_ids)
        kept_order = [pair_id for pair_id in existing_order if pair_id in current_pair_ids]
        missing_pair_ids = [pair_id for pair_id in pair_ids if pair_id not in kept_order]

        if missing_pair_ids:
            random.shuffle(missing_pair_ids)
            kept_order.extend(missing_pair_ids)

            connection.execute(
                "DELETE FROM user_pair_order WHERE username = ?",
                (username,),
            )
            connection.executemany(
                """
                INSERT INTO user_pair_order (username, position, pair_id)
                VALUES (?, ?, ?)
                """,
                [
                    (username, position, pair_id)
                    for position, pair_id in enumerate(kept_order)
                ],
            )

    return kept_order


def get_or_create_scenario_order(username: str, pair_id: str) -> list[str]:
    """Return this user's fixed randomized scenario order for one pair."""
    with get_db_connection() as connection:
        row = connection.execute(
            """
            SELECT scenario_order
            FROM user_scenario_order
            WHERE username = ? AND pair_id = ?
            """,
            (username, pair_id),
        ).fetchone()

        if row:
            return json.loads(row["scenario_order"])

        scenario_order = SCENARIO_TYPES.copy()
        random.shuffle(scenario_order)
        connection.execute(
            """
            INSERT INTO user_scenario_order (username, pair_id, scenario_order)
            VALUES (?, ?, ?)
            """,
            (username, pair_id, json.dumps(scenario_order)),
        )

    return scenario_order


def get_ranking(username: str, pair_id: str) -> sqlite3.Row | None:
    """Return a saved ranking for the user and pair, if one exists."""
    with get_db_connection() as connection:
        return connection.execute(
            """
            SELECT *
            FROM rankings
            WHERE username = ? AND pair_id = ?
            """,
            (username, pair_id),
        ).fetchone()


def get_completed_pair_ids(username: str, pair_ids: list[str]) -> set[str]:
    """Return the pair IDs already ranked by this user."""
    if not pair_ids:
        return set()

    placeholders = ",".join("?" for _ in pair_ids)
    query = f"""
        SELECT pair_id
        FROM rankings
        WHERE username = ? AND pair_id IN ({placeholders})
    """

    with get_db_connection() as connection:
        rows = connection.execute(query, [username, *pair_ids]).fetchall()

    return {row["pair_id"] for row in rows}


def get_first_incomplete_position(username: str, pair_order: list[str]) -> int | None:
    """Return the zero-based position of the first incomplete pair."""
    completed_pair_ids = get_completed_pair_ids(username, pair_order)
    for position, pair_id in enumerate(pair_order):
        if pair_id not in completed_pair_ids:
            return position
    return None


def get_display_order(username: str, pair_id: str) -> list[str]:
    """Return saved ranking order if available, otherwise the initial random order."""
    ranking = get_ranking(username, pair_id)
    if ranking:
        scores = {
            "gold": ranking["gold_score"],
            "web_search": ranking["web_search_score"],
            "baseline": ranking["baseline_score"],
        }
        return sorted(SCENARIO_TYPES, key=lambda item: scores[item], reverse=True)

    return get_or_create_scenario_order(username, pair_id)


def save_ranking(
    username: str,
    record: dict[str, str],
    ranking_order: list[str],
    comments: dict[str, str],
) -> None:
    """Save or update one ranking decision in SQLite."""
    if sorted(ranking_order) != sorted(SCENARIO_TYPES):
        raise ValueError("Ranking order must contain gold, web_search, and baseline.")

    scores = {
        scenario_type: len(SCENARIO_TYPES) - 1 - position
        for position, scenario_type in enumerate(ranking_order)
    }

    with get_db_connection() as connection:
        connection.execute(
            """
            INSERT INTO rankings (
                username,
                pair_id,
                gold_scenario,
                topic,
                target_attribute,
                web_search_scenario,
                baseline_scenario,
                web_search_example_question,
                baseline_example_question,
                cognitive_level,
                gold_score,
                web_search_score,
                baseline_score,
                gold_comment,
                web_search_comment,
                baseline_comment,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(username, pair_id) DO UPDATE SET
                gold_scenario = excluded.gold_scenario,
                topic = excluded.topic,
                target_attribute = excluded.target_attribute,
                web_search_scenario = excluded.web_search_scenario,
                baseline_scenario = excluded.baseline_scenario,
                web_search_example_question = excluded.web_search_example_question,
                baseline_example_question = excluded.baseline_example_question,
                cognitive_level = excluded.cognitive_level,
                gold_score = excluded.gold_score,
                web_search_score = excluded.web_search_score,
                baseline_score = excluded.baseline_score,
                gold_comment = excluded.gold_comment,
                web_search_comment = excluded.web_search_comment,
                baseline_comment = excluded.baseline_comment,
                updated_at = excluded.updated_at
            """,
            (
                username,
                record["pair_id"],
                record["gold_scenario"],
                record["topic"],
                record["target_attribute"],
                record["web_search_scenario"],
                record["baseline_scenario"],
                record["web_search_example_question"],
                record["baseline_example_question"],
                record["cognitive_level"],
                scores["gold"],
                scores["web_search"],
                scores["baseline"],
                comments["gold"],
                comments["web_search"],
                comments["baseline"],
                datetime.now(timezone.utc).isoformat(),
            ),
        )


def build_export_rows(username: str) -> list[dict[str, Any]]:
    """Return saved rankings for export, ordered by the user's pair order."""
    with get_db_connection() as connection:
        rows = connection.execute(
            """
            SELECT
                r.pair_id,
                r.gold_scenario,
                r.topic,
                r.target_attribute,
                r.web_search_scenario,
                r.baseline_scenario,
                r.web_search_example_question,
                r.baseline_example_question,
                r.cognitive_level,
                r.gold_score,
                r.web_search_score,
                r.baseline_score,
                r.gold_comment,
                r.web_search_comment,
                r.baseline_comment,
                r.username,
                r.updated_at
            FROM rankings r
            LEFT JOIN user_pair_order o
                ON r.username = o.username AND r.pair_id = o.pair_id
            WHERE r.username = ?
            ORDER BY o.position
            """,
            (username,),
        ).fetchall()

    return [dict(row) for row in rows]


@app.on_event("startup")
def startup() -> None:
    """Initialize the local database when the app starts."""
    init_db()


@app.get("/", response_class=HTMLResponse)
def login_page(request: Request) -> HTMLResponse:
    """Display the username login page."""
    return templates.TemplateResponse(
        name="login.html",
        context={
            "request": request,
            "previous_username": get_last_username(),
        },
        request=request,
    )


@app.post("/login")
async def login(request: Request) -> Response:
    """Store the username in a cookie and continue to the ranking task."""
    form_data = await request.form()
    username = str(form_data.get("username", "")).strip()
    action = str(form_data.get("action", "continue")).strip()
    confirmed_different_username = (
        str(form_data.get("confirmed_different_username", "")).strip() == "yes"
    )
    previous_username = get_last_username()

    if not username:
        return templates.TemplateResponse(
            name="login.html",
            context={
                "request": request,
                "error": "Please enter a username.",
                "previous_username": previous_username,
            },
            request=request,
            status_code=400,
        )

    if action == "use_previous" and previous_username:
        username = previous_username
    elif (
        previous_username
        and username != previous_username
        and not confirmed_different_username
    ):
        return templates.TemplateResponse(
            name="login.html",
            context={
                "request": request,
                "previous_username": previous_username,
                "requested_username": username,
                "username_warning": True,
            },
            request=request,
        )

    set_last_username(username)
    response = RedirectResponse(url="/rank", status_code=303)
    response.set_cookie("username", username, httponly=True, samesite="lax")
    return response


@app.get("/logout")
def logout() -> Response:
    """Clear the username cookie."""
    response = RedirectResponse(url="/", status_code=303)
    response.delete_cookie("username")
    return response


@app.get("/rank")
def resume_ranking(request: Request) -> Response:
    """Send the user to the first pair they have not completed."""
    username = require_username(request)
    records = load_source_records()
    pair_ids = [record["pair_id"] for record in records]
    pair_order = get_or_create_pair_order(username, pair_ids)
    first_incomplete = get_first_incomplete_position(username, pair_order)

    if first_incomplete is None:
        return RedirectResponse(url="/thanks", status_code=303)

    return RedirectResponse(url=f"/rank/{first_incomplete + 1}", status_code=303)


@app.get("/rank/{position}", response_class=HTMLResponse)
def ranking_page(request: Request, position: int, confirmed: int = 0) -> HTMLResponse:
    """Display one scenario-ranking page."""
    username = require_username(request)
    records_by_pair_id = get_records_by_pair_id()
    pair_ids = list(records_by_pair_id.keys())
    pair_order = get_or_create_pair_order(username, pair_ids)

    if not pair_order:
        raise HTTPException(status_code=404, detail="No scenario pairs found.")

    if position < 1:
        return RedirectResponse(url="/rank/1", status_code=303)

    if position > len(pair_order):
        return RedirectResponse(url="/thanks", status_code=303)

    pair_id = pair_order[position - 1]
    record = records_by_pair_id[pair_id]
    display_order = get_display_order(username, pair_id)
    saved_ranking = get_ranking(username, pair_id)
    scenarios = [
        build_scenario_card(record, scenario_type, saved_ranking)
        for scenario_type in display_order
    ]

    completed_count = len(get_completed_pair_ids(username, pair_order))
    remaining_count = len(pair_order) - completed_count

    return templates.TemplateResponse(
        name="rank.html",
        context={
            "request": request,
            "username": username,
            "position": position,
            "total_count": len(pair_order),
            "completed_count": completed_count,
            "remaining_count": remaining_count,
            "topic": record["topic"],
            "target_attribute": record["target_attribute"],
            "scenarios": scenarios,
            "confirmed": bool(confirmed),
            "has_saved_ranking": saved_ranking is not None,
            "is_first": position == 1,
            "is_last": position == len(pair_order),
        },
        request=request,
    )


@app.post("/rank/{position}")
async def submit_ranking(request: Request, position: int) -> Response:
    """Save one ranking decision and return to the page with navigation."""
    username = require_username(request)
    records_by_pair_id = get_records_by_pair_id()
    pair_ids = list(records_by_pair_id.keys())
    pair_order = get_or_create_pair_order(username, pair_ids)

    if position < 1 or position > len(pair_order):
        raise HTTPException(status_code=404, detail="Scenario pair not found.")

    form_data = await request.form()
    order_text = str(form_data.get("order", "")).strip()
    ranking_order = [item.strip() for item in order_text.split(",") if item.strip()]
    comments = {
        scenario_type: str(form_data.get(f"comment_{scenario_type}", "")).strip()
        for scenario_type in SCENARIO_TYPES
    }

    pair_id = pair_order[position - 1]
    record = records_by_pair_id[pair_id]

    try:
        save_ranking(username, record, ranking_order, comments)
    except ValueError as error:
        raise HTTPException(status_code=400, detail=str(error)) from error

    if position == len(pair_order):
        return RedirectResponse(url="/thanks", status_code=303)

    return RedirectResponse(
        url=f"/rank/{position}?confirmed=1#ranking-navigation",
        status_code=303,
    )


@app.get("/thanks", response_class=HTMLResponse)
def thanks_page(request: Request) -> HTMLResponse:
    """Display the thank-you page after all scenario pairs are ranked."""
    username = require_username(request)
    records = load_source_records()
    pair_ids = [record["pair_id"] for record in records]
    pair_order = get_or_create_pair_order(username, pair_ids)
    completed_count = len(get_completed_pair_ids(username, pair_order))

    return templates.TemplateResponse(
        name="thanks.html",
        context={
            "request": request,
            "username": username,
            "completed_count": completed_count,
            "total_count": len(pair_order),
        },
        request=request,
    )


@app.get("/export")
def export_results(request: Request) -> StreamingResponse:
    """Export the current user's ranking results as an Excel file."""
    username = require_username(request)
    export_rows = build_export_rows(username)

    export_columns = [
        "pair_id",
        "gold_scenario",
        "topic",
        "target_attribute",
        "web_search_scenario",
        "baseline_scenario",
        "web_search_example_question",
        "baseline_example_question",
        "cognitive_level",
        "gold_score",
        "web_search_score",
        "baseline_score",
        "gold_comment",
        "web_search_comment",
        "baseline_comment",
        "username",
        "updated_at",
    ]
    export_data = pd.DataFrame(export_rows, columns=export_columns)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_data.to_excel(writer, index=False, sheet_name="rankings")
    output.seek(0)

    safe_username = "".join(
        character if character.isalnum() or character in {"-", "_"} else "_"
        for character in username
    )
    file_name = f"scenario_rankings_{safe_username}.xlsx"

    return StreamingResponse(
        output,
        media_type=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition": f'attachment; filename="{file_name}"',
        },
    )
