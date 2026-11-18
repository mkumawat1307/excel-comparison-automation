# Installation

This guide walks you through installing and running Excel Comparison Automation on Windows using PowerShell.

## Prerequisites

| Requirement | Detail |
|-------------|--------|
| Python | 3.10 or later (required by `pyproject.toml`: `requires-python = ">=3.10"`) |
| pip | Python's package installer (bundled with Python) |
| Git | Optional. Only needed if you clone the repository from GitHub. Not required to run the application if you already have the source files. |
| Windows PowerShell | 5.1 or later (built into Windows 10/11) |

## Step 1: Get the Source Code

### Option A: Clone with Git (optional)

```powershell
git clone https://github.com/mkumawat1307/excel-comparison-automation.git
cd excel-comparison-automation
```

The repository is at https://github.com/mkumawat1307/excel-comparison-automation.git.

### Option B: Use existing source files

If you already have the project folder, open PowerShell in that directory:

```powershell
cd <cloned-repository-directory>
```

Verify the project files are present:

```powershell
Get-ChildItem -Name
```

You should see `pyproject.toml`, `README.md`, `app/`, `src/`, and `tests/`.

## Step 2: Check Python Version

Run:

```powershell
python --version
```

If `python` is not recognized, try `py --version` (Python Launcher for Windows) or `python3 --version`. You need Python 3.10 or later.

## Step 3: Create a Virtual Environment

It is strongly recommended to install dependencies into an isolated virtual environment so they do not conflict with other Python projects.

From the project root directory:

```powershell
python -m venv .venv
```

Activate the virtual environment:

```powershell
.\.venv\Scripts\Activate.ps1
```

You should see `(.venv)` prepended to your PowerShell prompt, indicating the environment is active.

> **Note:** If you receive an execution-policy error when activating, run the following once in an Administrator PowerShell session, then retry activation:
> ```powershell
> Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
> ```

## Step 4: Install Dependencies

With the virtual environment activated, install the project and its runtime dependencies in editable mode:

```powershell
pip install -e ".[dev]"
```

This installs:

- `pandas>=2.0.0`
- `openpyxl>=3.1.0`
- `streamlit>=1.30.0`
- `pytest>=8.0.0` (dev)
- `ruff>=0.5.0` (dev)

Verify the installation:

```powershell
pip list
```

You should see `excel-comparison-automation`, `pandas`, `openpyxl`, `streamlit`, `pytest`, and `ruff` in the list.

## Step 5: Launch the Application

From the project root directory (with `.venv` activated):

```powershell
streamlit run app/streamlit_app.py
```

Streamlit will start a local web server and print a URL (typically `http://localhost:8501`). Open that URL in your browser to use the application.

To stop the application, press `Ctrl + C` in the PowerShell terminal.

## How to Run Tests

All tests are written with pytest and live in `tests/unit/`.

### Run the full test suite

```powershell
python -m pytest -q
```

Expected output: all tests pass (234 tests as of the current release).

### Run an individual test file

```powershell
python -m pytest tests/unit/test_writer.py -q
```

Replace `test_writer.py` with any other test file name (for example, `test_matcher.py`, `test_comparer.py`, `test_security.py`).

### Run tests matching a keyword

```powershell
python -m pytest -q -k "duplicate"
```

## How to Run Ruff Checks

Ruff is used for linting and formatting. It is configured in `pyproject.toml` (`line-length = 100`, target `py310`).

### Lint check

```powershell
python -m ruff check .
```

### Format check

```powershell
python -m ruff format --check .
```

### Apply formatting (if needed)

```powershell
python -m ruff format .
```

## Common Installation Problems and Solutions

| Problem | Likely cause | Solution |
|---------|--------------|----------|
| `python` is not recognized | Python not installed or not on PATH | Install Python from https://python.org, ensuring "Add Python to PATH" is checked during installation. |
| `streamlit` is not recognized | Virtual environment not activated | Run `.\.venv\Scripts\Activate.ps1`, then retry. |
| Execution policy prevents script activation | PowerShell security policy | Run `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser` in an Administrator PowerShell, then retry activation. |
| `pip install -e ".[dev]"` fails | Build tools missing | Ensure you have an internet connection and pip is up to date: `python -m pip install --upgrade pip setuptools wheel`. |
| Tests cannot import `excel_compare` | Project not installed in editable mode | Run `pip install -e ".[dev]"` from the project root. |
| Streamlit port already in use | Another app is using port 8501 | Use a different port: `streamlit run app/streamlit_app.py --server.port 8502`. |

## Next Steps

After installation, see [USER_GUIDE.md](USER_GUIDE.md) for step-by-step instructions on using the application, and [PROJECT_STRUCTURE.md](PROJECT_STRUCTURE.md) for a developer-oriented overview of the codebase.