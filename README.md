# Scenario Ranking App

This is a small FastAPI app for ranking generated scenarios. It reads
`scenario_comparison_example_questions.xlsx` from the project root and stores
ranking results locally in `scenario_rankings.sqlite3`.

## Holistic Judgement Rubric

### Overall Scenario Quality

Experts should rank the three scenarios based on their overall usefulness for
downstream college statistics question generation.

A high-quality scenario should:

- Clearly support the intended statistics topic and target attribute.
- Present a realistic and instructionally appropriate context.
- Provide enough relevant information to support a clear and solvable
  problem-solving question.
- Be written in a clear, self-contained way for the intended student audience.


## Setup

Create a virtual environment in the project folder:

```bash
python -m venv .venv
```

Activate the virtual environment.

On Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

On Windows Command Prompt:

```cmd
.\.venv\Scripts\activate.bat
```

On macOS or Linux:

```bash
source .venv/bin/activate
```

After activation, your terminal should show `(.venv)` near the prompt. Then
upgrade `pip` and install the dependencies inside the virtual environment:

```bash
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Make sure the Excel file is in the same directory:

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


## Next Time You Open the Repo

You do not need to create the virtual environment again. The `.venv` folder
will stay in the project folder unless you delete it.

When you open the repo again in VS Code, open a new terminal and activate the
existing virtual environment.

On Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

On Windows Command Prompt:

```cmd
.\.venv\Scripts\activate.bat
```

On macOS or Linux:

```bash
source .venv/bin/activate
```

After the virtual environment is activated, run the app:

```bash
uvicorn main:app --reload
```

## Export

After completing the ranking task, the thank-you page includes an export button.
The export downloads the current user's saved results as an `.xlsx` file with
the original fields, ranking scores, comments, username, and timestamp.

## Username Behavior

The app stores the most recent username in SQLite. If someone enters a
different username on the login page, the app asks whether to continue with the
new username or use the previous one. Returning users are sent to the first
scenario pair they have not completed.
