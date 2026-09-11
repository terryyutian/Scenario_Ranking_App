# Scenario Ranking App

This is a small FastAPI app for ranking generated scenarios. It reads
`scenario_comparison_example_questions.xlsx` from the project root and stores
ranking results locally in `scenario_rankings.sqlite3`.

## Setup

Install the dependencies:

```bash
pip install -r requirements.txt
```

Make sure the Excel file is in this same directory:

```text
scenario_comparison_example_questions.xlsx
```

Run the app:

```bash
uvicorn main:app --reload
```

Then open:

```text
http://127.0.0.1:8000
```

## Output Stored in SQLite

Each saved ranking includes:

- `username`
- `pair_id`
- `gold_scenario`
- `topic`
- `target_attribute`
- `web_search_scenario`
- `baseline_scenario`
- `web_search_example_question`
- `baseline_example_question`
- `gold_score`
- `web_search_score`
- `baseline_score`
- `updated_at`

Scores are assigned based on the final order on the page:

- Top scenario: `2`
- Middle scenario: `1`
- Bottom scenario: `0`
