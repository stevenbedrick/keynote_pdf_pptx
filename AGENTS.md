# AGENTS.md — keynote_pdf_pptx Project Context

This file provides context for AI coding agents working in this repository.

---

## Project Overview

This is a working directory for a small utility to manage a workflow for converting presentations from Keynote to PowerPoint while preserving fonts, speaker notes, etc.

---

## Runtime & Dependency Management

This project uses **`uv`** for all dependency management and script execution.

- Run scripts: `uv run python <script>.py`
- Add dependencies: `uv add <package>`
- Remove dependencies: `uv remove <package>`
- Python version: 3.13+

**Never use bare `python` or `pip` directly.** Always use `uv run` or `uv add`.

---

## Python Coding Style

### Enums and StrEnum

Prefer enums over bare strings for any value that has a fixed set of valid options.

- Use `StrEnum` (Python 3.11+) for string-valued enums — it inherits from `str` so values serialise as plain strings in JSON and are accepted directly by Pydantic without extra configuration:
  ```python
  from enum import StrEnum

  class PlausibilityConcern(StrEnum):
      NONE = "none"
      COUNT_TOO_LOW = "count_too_low"
      COUNT_TOO_HIGH = "count_too_high"
  ```
- Use plain `Enum` for non-string values (ints, etc.).
- Never use `str, Enum` multiple inheritance — `StrEnum` does that for you.

### Making Illegal States Unrepresentable

Prefer types that make invalid states impossible to construct over runtime checks that catch them after the fact.

- Use `str | None` (and check `is not None`) rather than using an empty string as a sentinel for "no value".
- Use enums instead of `str` for fields with a closed set of values — this prevents invalid values from ever entering the system.
- Use Pydantic `Field(...)` (required, no default) for fields that must always be present rather than giving them a default that would allow them to be silently omitted.
- Prefer narrow, specific types: a function returning `list[OMOPConcept]` is better than one returning `list[dict]`.
- If a function can fail in two meaningfully different ways, return a typed union (e.g. `ValidationError | ValidationTimeout | bool`) rather than a single `str | bool` where the caller must parse the string to understand what happened.
- The Python `match` operator is your friend.

### Pydantic BaseModel

Use `pydantic.BaseModel` for any structured data that crosses a boundary (tool return values, agent outputs, config, file-format records).

- All agent tool return types should be `BaseModel` subclasses or lists thereof — this gives the LLM a well-defined schema to reason about.
- Use `Field(description=...)` on every field in models that are exposed to an LLM — the description becomes part of the schema the model sees.
- Prefer `model_dump_json()` / `model_validate_json()` for serialisation rather than manual `json.dumps` / `json.loads`.
- Keep models flat and single-purpose; avoid deeply nested models unless the nesting reflects genuine domain structure.

### General

- We prefer **Polars** over Pandas for any DataFrame work.
- For new command-line tools, use `jsonargparse` where it makes sense to do so.
  - Default to `auto_cli()` unless there's a real reason to use the more detailed APIs.
- For nicer console output, use `rich`, but use it in moderation and where it adds something, not for its own sake.
- Always use `uv run` / `uv add` — never bare `python` or `pip`.
- When parsing input files, default to using libraries (i.e. a proper docx-parsing library for Word documents, openpyxl for Excel, a bibtex parser for bibtex, etc.) rather than regexes or brittler methods

---

## General Notes

- Rather than diving straight in to writing code, your first step should be a planning session. 
  - Keep things brief and high-level to start, and make sure to ask clarifying questions. 
  - Don't start coding or editing files straight away unless specifically told to do so.
- Don't assume that "the user knows best"; if something doesn't make sense or seems like a bad idea, check in about that rather than "going with the flow."
  - That said, the user _is_ driving the bus; after the user has heard you out, their judgment should take priority.
- Azure things change rapidly so remember that your information may be out of date.
- When writing documentation, don't rewrite existing prose (e.g. in a README file) unless specifically asked to- the user likely wrote it the way they did for a reason. 
  - If it's out of date and needs to be updated, check with the user first and then make your edits in the most parsimonious way possible.
