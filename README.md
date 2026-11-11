# Personal AI Assistant

A single-file, LLM-driven Windows desktop assistant built with Python, Ollama, PowerShell, and SQLite.

The assistant accepts natural-language requests. Simple `open <name>` requests are handled by a deterministic PowerShell workflow. Other Windows tasks are sent to the local Ollama model, which proposes a PowerShell script.

**Every proposed script is displayed and requires explicit approval before it runs.** Successful operations can save reusable applications, files, folders, projects, and supported URIs in a local SQLite database.

> **Security warning:** Approval is not a sandbox. An approved PowerShell script runs with the permissions of your current Windows user. Review the complete script before typing `YES`.

## Features

- Natural-language interaction through a locally running Ollama model.
- One generic `execute_windows_action` tool for LLM-generated Windows operations.
- Deterministic handling for simple `open <name>` requests.
- Dynamic application discovery through Windows Start Menu application metadata.
- SQLite storage for reusable resources, aliases, usage information, and execution history.
- Reuse of saved resources instead of rediscovering them on every request.
- Explicit approval before executing a proposed PowerShell script.
- Basic local command checks to reject selected risky PowerShell patterns.
- Resource-marker parsing and validation before saving discovered resources.
- Self-tests that run without contacting Ollama or executing PowerShell.

## Requirements

- Windows 10 or Windows 11.
- Python 3.10 or later.
- [Ollama](https://ollama.com/) installed and running locally.
- The `qwen2.5:3b` model pulled in Ollama.
- Windows PowerShell (`powershell.exe`).

The application uses Python standard-library modules. No `pip install` step is required by the current `agent.py`.

## Project structure

```text
personal_ai_assistant/
├── agent.py          # Assistant application
├── assistant.db      # Created automatically; local SQLite database
├── README.md
└── .gitignore
```

`assistant.db` is created beside `agent.py` when the application initializes. It stores reusable resources and execution history. It is local application data, not a source file.

## Setup

### 1. Install and start Ollama

Install Ollama from [ollama.com](https://ollama.com/), then open PowerShell and pull the model:

```powershell
ollama pull qwen2.5:3b
```

Confirm the model is available:

```powershell
ollama list
```

Make sure Ollama is running. Its local API is expected at:

```text
http://localhost:11434
```

### 2. Open the project directory

```powershell
cd D:\personal_ai_assistant
```

Use the path where you actually saved the project if it is different.

### 3. Check Python

```powershell
py --version
```

The application requires Python 3.10 or later.

### 4. Run the self-test

```powershell
py agent.py --self-test
```

The self-test checks local functionality such as SQLite operations, resource validation, marker parsing, command guardrails, plain-text PowerShell fallback, deterministic open scripts, and the tool schema. It does not contact Ollama or execute PowerShell actions.

### 5. Start the assistant

```powershell
py agent.py
```

## Configuration

The application supports two optional environment variables:

| Variable | Default | Purpose |
|---|---|---|
| `OLLAMA_MODEL` | `qwen2.5:3b` | Ollama model used for chat |
| `OLLAMA_URL` | `http://localhost:11434/api/chat` | Ollama chat API endpoint |

Set them in the current PowerShell session before starting the assistant:

```powershell
$env:OLLAMA_MODEL = "qwen2.5:3b"
$env:OLLAMA_URL = "http://localhost:11434/api/chat"

py agent.py
```

If you do not set these variables, the defaults are used.

## Using the assistant

At startup, the application displays the selected model and SQLite database path.

Enter a request at the `You:` prompt. Examples:

```text
open chrome
open calculator
open visual studio code
```

Simple `open`, `launch`, or `start` requests for a single application/resource name use the host-generated PowerShell workflow. The assistant searches Windows Start Menu application metadata when the resource is not already saved. If multiple applications match, it reports the candidates instead of choosing one silently.

Other requests are sent to Ollama so the model can propose a PowerShell script. For example, you can describe a Windows task in natural language. The exact operation depends on the request and the script proposed by the model.

### Approval flow

For a proposed PowerShell action, the assistant displays:

1. The purpose of the operation.
2. The complete PowerShell script.
3. A warning that the script runs as your current Windows user.
4. An approval prompt.

Type the exact text below to approve:

```text
YES
```

Anything else cancels the operation. A rejected script is not executed.

Do not approve a script unless you understand what it does. The local checks are guardrails, not a complete security boundary.

### Built-in commands

| Input | Action |
|---|---|
| `exit` or `quit` | Stop the assistant |
| `:resources` | List saved resources |
| `open <name>` | Open a saved resource or try to discover an installed application |

## SQLite resource memory

The database is stored at:

```text
assistant.db
```

It is created automatically beside `agent.py`.

### `resources` table

Stores reusable targets, including:

- Resource name and aliases.
- Launch type (`path`, `uri`, or `shell`).
- Target path or URI.
- Optional arguments and metadata.
- Creation/update timestamps and last-used information.
- Usage count.

Supported resource types are validated before being stored:

- `path`: an existing absolute file-system path.
- `uri`: a URI using one of the supported schemes: `ms-settings`, `https`, `http`, or `mailto`.
- `shell`: a Windows `shell:AppsFolder\...` target.

### `execution_history` table

Stores a record of proposed operations, including the user request, purpose, command, whether it was approved, return code, output excerpt, and timestamp.

### View saved resources

Start the assistant and enter:

```text
:resources
```

The database is local to this project. Back it up separately if you need to preserve its contents.

## How it works

1. The application receives a natural-language request.
2. A simple `open <name>` request is checked first and handled by the deterministic host workflow.
3. Other requests are sent to Ollama with the generic `execute_windows_action` tool definition.
4. The model may return a structured tool call or a PowerShell script as plain text. Both paths are routed through the same local checks and approval workflow.
5. The application checks the proposed script against a limited set of blocked patterns.
6. The user must approve the exact script by typing `YES`.
7. The host runs the approved script in a temporary `.ps1` file using `powershell.exe`.
8. If the script succeeds, valid resource records emitted by the script can be saved to SQLite.
9. The execution result is recorded and returned to the assistant.

## Safety and limitations

- **Scripts are not sandboxed.** They run with the current Windows user's permissions.
- The approval prompt does not make an unsafe script safe. Inspect every line before approving.
- Command checks block selected patterns, but they cannot detect every harmful or unintended PowerShell operation.
- The system prompt instructs the model not to perform destructive operations, download/install software, access credentials, transmit private data, change security settings, or request elevation. These model instructions are not a security boundary.
- Do not use this assistant to run scripts you do not understand.
- The current implementation is for Windows and uses `powershell.exe`.
- The assistant depends on Ollama being available for LLM-driven requests. Simple deterministic `open <name>` requests do not require an LLM response.
- Long-running development servers are not managed as persistent background services by this implementation.

## Troubleshooting

### Cannot connect to Ollama

Check that Ollama is running:

```powershell
ollama list
```

Check the local API endpoint:

```powershell
Invoke-RestMethod -Uri "http://localhost:11434/api/tags"
```

If you use a different Ollama endpoint, set `OLLAMA_URL` before starting `agent.py`.

### Model not found

Pull the configured model:

```powershell
ollama pull qwen2.5:3b
```

If you use another model, set `OLLAMA_MODEL` to its exact local model name.

### Application does not open

- Check whether the application is installed.
- For simple `open <name>` requests, use a more specific application name if multiple matches are found.
- If a resource was saved previously, inspect it with `:resources`.
- A saved path may no longer exist. The assistant checks saved file-system paths before opening them.

### Self-test fails

Run:

```powershell
py -m py_compile agent.py
py agent.py --self-test
```

Review the first reported error and confirm that the project file has not been partially overwritten or edited.

## Development checks

Before committing changes, run:

```powershell
py -m py_compile agent.py
py agent.py --self-test
```

The self-test is a local code check. It does not prove that every generated PowerShell script is safe or that every Windows application can be discovered.

## Git

The project `.gitignore` should exclude local runtime data such as:

```gitignore
assistant.db
*.db
*.sqlite
*.sqlite3
```

If the database was already committed, adding it to `.gitignore` does not remove it from Git tracking. To stop tracking it while keeping the local file:

```powershell
git rm --cached assistant.db
git add .gitignore
git commit -m "Stop tracking local assistant database"
```

Do not commit private database contents, credentials, or other local user data.
