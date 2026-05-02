"""
keynote.py — Utilities for interacting with the macOS Keynote application via
osascript (AppleScript).

Provides helpers to:
- List open Keynote documents
- Interactively pick one (with fuzzy default selection)
- Export a Keynote document to a PPTX file
"""

from __future__ import annotations

import re
import subprocess
from pathlib import Path

from rich.console import Console
from rich.prompt import Prompt

console = Console()

# Suffixes that are stripped when fuzzy-matching a Keynote filename against a
# PDF stem (case-insensitive).
_FUZZY_STRIP_SUFFIXES = [" slides", " slide deck", " deck", " presentation", " talk"]


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _run_osascript(script: str) -> subprocess.CompletedProcess[str]:
    """Run *script* via ``osascript`` and return the CompletedProcess."""
    return subprocess.run(
        ["osascript", "-e", script],
        capture_output=True,
        text=True,
    )


def _normalise_for_fuzzy(name: str) -> str:
    """Lower-case *name* and strip common presentation suffixes."""
    result = name.lower()
    for suffix in _FUZZY_STRIP_SUFFIXES:
        result = re.sub(re.escape(suffix) + r"$", "", result).rstrip()
    return result


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def list_open_keynote_documents() -> list[Path]:
    """Return the paths of all currently open Keynote documents.

    Uses ``osascript`` to query Keynote.  Returns an empty list if Keynote is
    not running or no documents are open.
    """
    script = """
tell application "Keynote"
    set docList to {}
    set n to count of documents
    repeat with i from 1 to n
        set f to file of document i
        set end of docList to POSIX path of f
    end repeat
    return docList
end tell
"""
    result = _run_osascript(script)
    if result.returncode != 0:
        # Keynote might not be running — that's fine
        return []

    raw = result.stdout.strip()
    if not raw:
        return []

    # osascript returns a comma-separated list of POSIX paths.
    # Paths themselves don't normally contain commas, but split carefully.
    paths: list[Path] = []
    for part in raw.split(", "):
        part = part.strip()
        if part:
            paths.append(Path(part))
    return paths


def pick_keynote_document(pdf_path: Path) -> Path | None:
    """Interactively pick an open Keynote document to use as the notes source.

    - If no documents are open, prints an error and returns ``None``.
    - If exactly one document is open, returns it directly (no prompt).
    - If multiple documents are open, shows a numbered menu and prompts the
      user to choose.  Attempts to pre-select a default by fuzzy-matching
      *pdf_path.stem* against document names.
    """
    documents = list_open_keynote_documents()

    if not documents:
        console.print(
            "[bold red]Error:[/bold red] No Keynote documents are currently open. "
            "Please open the relevant Keynote file and try again."
        )
        return None

    if len(documents) == 1:
        console.print(f"[dim]Using open Keynote document: {documents[0].name}[/dim]")
        return documents[0]

    # --- Multiple documents open: show a menu ---
    console.print("\n[bold]Open Keynote documents:[/bold]")
    for i, doc in enumerate(documents, start=1):
        console.print(f"  [cyan]{i}[/cyan]. {doc.name}")

    # Try to find a best default by fuzzy-matching the PDF stem
    pdf_stem_norm = _normalise_for_fuzzy(pdf_path.stem)
    best_index: int | None = None
    best_score = -1
    for i, doc in enumerate(documents):
        doc_norm = _normalise_for_fuzzy(doc.stem)
        # Simple scoring: exact match wins; otherwise check containment
        if doc_norm == pdf_stem_norm:
            best_index = i
            best_score = 2
            break
        if pdf_stem_norm in doc_norm or doc_norm in pdf_stem_norm:
            score = len(set(pdf_stem_norm.split()) & set(doc_norm.split()))
            if score > best_score:
                best_score = score
                best_index = i

    default_display: str | None = None
    if best_index is not None:
        default_display = str(best_index + 1)  # 1-based for display
        console.print(f"[dim](Best match: {documents[best_index].name})[/dim]")

    while True:
        prompt_text = "Select document number"
        if default_display is not None:
            prompt_text += f" (default: {default_display})"

        raw = Prompt.ask(prompt_text, console=console, default=default_display or "")
        raw = raw.strip()

        if not raw and default_display is not None:
            raw = default_display

        try:
            choice = int(raw)
        except ValueError:
            console.print("[red]Please enter a valid number.[/red]")
            continue

        if 1 <= choice <= len(documents):
            return documents[choice - 1]
        else:
            console.print(f"[red]Please enter a number between 1 and {len(documents)}.[/red]")


def export_keynote_to_pptx(keynote_path: Path, output_path: Path) -> None:
    """Export a Keynote document to a PPTX file via ``osascript``.

    *keynote_path* and *output_path* must be absolute POSIX paths.  The
    document is opened (Keynote handles it gracefully if already open) and
    exported; it is **not** closed afterwards so the user's working state is
    preserved.

    Raises:
        RuntimeError: if ``osascript`` exits with a non-zero return code.
    """
    script = f"""tell application "Keynote"
    set doc to open POSIX file "{keynote_path}"
    set outputFile to POSIX file "{output_path}"
    export doc to file outputFile as Microsoft PowerPoint
end tell"""

    result = _run_osascript(script)
    if result.returncode != 0:
        stderr = result.stderr.strip()
        raise RuntimeError(
            f"osascript failed (exit {result.returncode}) while exporting Keynote to PPTX"
            + (f": {stderr}" if stderr else "")
        )
