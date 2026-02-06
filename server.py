"""
OneNote MCP Server (Local Files)
=================================
An MCP (Model Context Protocol) server that reads local OneNote (.one) files
directly from disk and writes to OneNote via the COM API.
No Azure registration or authentication needed.

Reading: parses backup files at:
    C:\\Users\\<user>\\AppData\\Local\\Microsoft\\OneNote\\16.0\\Backup\\

Writing: uses the OneNote COM API via PowerShell (requires OneNote desktop app).

It exposes tools for Claude Code to:
    - List all notebooks and sections
    - Read page text content
    - Search across all pages
    - Create new pages in any notebook/section
    - Append content to existing pages

Prerequisites:
    pip install "mcp[cli]" pyOneNote
    (or: uv add "mcp[cli]" pyOneNote)
    + OneNote desktop app (for write features)

Usage with Claude Code:
    claude mcp add --transport stdio onenote -- uv --directory "path/to/this/project" run server.py
"""

import logging
import os
import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path

from pyOneNote.OneDocument import OneDocment
from mcp.server.fastmcp import FastMCP

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Where OneNote stores local backup files.
# Override with ONENOTE_BACKUP_DIR environment variable if yours is elsewhere.
DEFAULT_BACKUP_DIR = Path(
    os.environ.get("APPDATA", ""),
).parent / "Local" / "Microsoft" / "OneNote" / "16.0" / "Backup"

ONENOTE_DIR = Path(
    os.environ.get("ONENOTE_BACKUP_DIR", str(DEFAULT_BACKUP_DIR))
)

# ---------------------------------------------------------------------------
# Logging (to stderr so it doesn't break stdio MCP transport)
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    stream=sys.stderr,
)
log = logging.getLogger("onenote-mcp")

# ---------------------------------------------------------------------------
# OneNote file parsing helpers
# ---------------------------------------------------------------------------


def _discover_notebooks() -> dict[str, dict]:
    """
    Scan the OneNote backup directory and build a notebook → section → files map.

    Returns a dict like:
    {
        "My Notebook": {
            "path": Path(...),
            "sections": {
                "Algorithm": {
                    "files": [Path("Algorithm (On 1-4-2026).one"), ...],
                    "latest": Path(...)   # most recently modified
                },
                ...
            }
        },
        ...
    }
    """
    if not ONENOTE_DIR.exists():
        log.error("OneNote backup directory not found: %s", ONENOTE_DIR)
        return {}

    notebooks = {}
    for notebook_dir in ONENOTE_DIR.iterdir():
        if not notebook_dir.is_dir():
            continue

        notebook_name = notebook_dir.name
        sections: dict[str, dict] = {}

        # Walk all .one files in this notebook (including subdirectories)
        for one_file in notebook_dir.rglob("*.one"):
            # Skip recycle bin
            if "RecycleBin" in str(one_file):
                continue

            # Extract the base section name (strip the date suffix)
            # e.g. "Algorithm (On 1-4-2026).one" → "Algorithm"
            # e.g. "Python.one (On 12-6-2025).one" → "Python"
            fname = one_file.name
            # Remove .one extension(s) and date suffixes
            section_name = re.sub(r"\.one$", "", fname)
            section_name = re.sub(r"\s*\(On \d+-\d+-\d+\)$", "", section_name)
            section_name = re.sub(r"\.one$", "", section_name)  # handle double .one
            section_name = section_name.strip()

            if not section_name:
                section_name = "(unnamed)"

            # Build relative path for context (subfolder within notebook)
            rel_parts = one_file.parent.relative_to(notebook_dir).parts
            if rel_parts:
                section_key = "/".join(rel_parts) + "/" + section_name
            else:
                section_key = section_name

            if section_key not in sections:
                sections[section_key] = {"files": [], "latest": None}

            sections[section_key]["files"].append(one_file)

        # For each section, determine the latest (most recently modified) file
        for sec_info in sections.values():
            sec_info["files"].sort(key=lambda p: p.stat().st_mtime, reverse=True)
            sec_info["latest"] = sec_info["files"][0]

        if sections:
            notebooks[notebook_name] = {
                "path": notebook_dir,
                "sections": sections,
            }

    return notebooks


def _parse_one_file(filepath: Path) -> list[str]:
    """
    Parse a .one file and extract all text content.

    Returns a list of text strings found in the file.
    """
    texts = []
    try:
        with open(filepath, "rb") as f:
            doc = OneDocment(f)

        props = doc.get_properties()
        for prop in props:
            ptype = prop.get("type", "")
            val = prop.get("val", {})
            if not isinstance(val, dict):
                continue

            # Extract RichEditTextUnicode (the actual text content)
            text = val.get("RichEditTextUnicode", "")
            if text and isinstance(text, str) and text.strip():
                texts.append(text.strip())

    except Exception as e:
        log.warning("Failed to parse %s: %s", filepath, e)

    return texts


def _get_page_titles_from_props(filepath: Path) -> list[str]:
    """Extract page titles from a .one file."""
    titles = []
    try:
        with open(filepath, "rb") as f:
            doc = OneDocment(f)

        props = doc.get_properties()
        for prop in props:
            if prop.get("type") == "jcidTitleNode":
                val = prop.get("val", {})
                if isinstance(val, dict):
                    text = val.get("RichEditTextUnicode", "")
                    if text and text.strip():
                        titles.append(text.strip())
    except Exception as e:
        log.warning("Failed to extract titles from %s: %s", filepath, e)

    return titles


# ---------------------------------------------------------------------------
# MCP Server
# ---------------------------------------------------------------------------

mcp = FastMCP("onenote")


@mcp.tool()
async def list_notebooks() -> str:
    """List all locally available OneNote notebooks.

    Shows notebook names and how many sections each one has.
    """
    notebooks = _discover_notebooks()
    if not notebooks:
        return f"No notebooks found in {ONENOTE_DIR}"

    lines = []
    for name, info in sorted(notebooks.items()):
        section_count = len(info["sections"])
        lines.append(f"- {name}  ({section_count} sections)")
    return "\n".join(lines)


@mcp.tool()
async def list_sections(notebook_name: str) -> str:
    """List all sections in a specific notebook.

    Args:
        notebook_name: The name of the notebook (from list_notebooks).
    """
    notebooks = _discover_notebooks()
    if notebook_name not in notebooks:
        # Try case-insensitive match
        for key in notebooks:
            if key.lower() == notebook_name.lower():
                notebook_name = key
                break
        else:
            available = ", ".join(sorted(notebooks.keys()))
            return f"Notebook '{notebook_name}' not found. Available: {available}"

    sections = notebooks[notebook_name]["sections"]
    lines = []
    for sec_name, sec_info in sorted(sections.items()):
        latest = sec_info["latest"]
        size_kb = latest.stat().st_size / 1024
        lines.append(f"- {sec_name}  ({size_kb:.0f} KB)")
    return "\n".join(lines)


@mcp.tool()
async def read_section(notebook_name: str, section_name: str) -> str:
    """Read all text content from a specific section of a notebook.

    Args:
        notebook_name: The name of the notebook.
        section_name: The name of the section (from list_sections).
    """
    notebooks = _discover_notebooks()

    # Case-insensitive notebook match
    nb = None
    for key, val in notebooks.items():
        if key.lower() == notebook_name.lower():
            nb = val
            break
    if nb is None:
        available = ", ".join(sorted(notebooks.keys()))
        return f"Notebook '{notebook_name}' not found. Available: {available}"

    # Case-insensitive section match
    sec_info = None
    for key, val in nb["sections"].items():
        if key.lower() == section_name.lower():
            sec_info = val
            break
    if sec_info is None:
        available = ", ".join(sorted(nb["sections"].keys()))
        return f"Section '{section_name}' not found. Available: {available}"

    filepath = sec_info["latest"]
    texts = _parse_one_file(filepath)

    if not texts:
        return f"No text content found in section '{section_name}'."

    return "\n\n".join(texts)


@mcp.tool()
async def search_notes(query: str) -> str:
    """Search for text across ALL notebooks and sections.

    Searches through the text content of every section for the given query.
    Returns matching sections with a snippet of the matched text.

    Args:
        query: The text to search for (case-insensitive).
    """
    query_lower = query.lower()
    notebooks = _discover_notebooks()
    results = []

    for nb_name, nb_info in sorted(notebooks.items()):
        for sec_name, sec_info in sorted(nb_info["sections"].items()):
            filepath = sec_info["latest"]
            texts = _parse_one_file(filepath)

            for text in texts:
                if query_lower in text.lower():
                    # Build a snippet around the match
                    idx = text.lower().index(query_lower)
                    start = max(0, idx - 80)
                    end = min(len(text), idx + len(query) + 80)
                    snippet = text[start:end].strip()
                    if start > 0:
                        snippet = "..." + snippet
                    if end < len(text):
                        snippet = snippet + "..."

                    results.append(
                        f"[{nb_name} / {sec_name}]\n  {snippet}"
                    )

    if not results:
        return f"No results found for '{query}'."

    header = f"Found {len(results)} match(es) for '{query}':\n\n"
    return header + "\n\n".join(results[:30])  # limit to 30 results


@mcp.tool()
async def list_all_sections() -> str:
    """List ALL sections across ALL notebooks.

    Useful for getting a complete overview of everything in your OneNote.
    """
    notebooks = _discover_notebooks()
    if not notebooks:
        return f"No notebooks found in {ONENOTE_DIR}"

    lines = []
    for nb_name, nb_info in sorted(notebooks.items()):
        lines.append(f"\n## {nb_name}")
        for sec_name, sec_info in sorted(nb_info["sections"].items()):
            latest = sec_info["latest"]
            size_kb = latest.stat().st_size / 1024
            lines.append(f"  - {sec_name}  ({size_kb:.0f} KB)")

    return "\n".join(lines)


@mcp.tool()
async def get_notebook_summary(notebook_name: str) -> str:
    """Get a summary of a notebook: its sections and a preview of each section's content.

    Args:
        notebook_name: The name of the notebook.
    """
    notebooks = _discover_notebooks()

    nb = None
    for key, val in notebooks.items():
        if key.lower() == notebook_name.lower():
            nb = val
            notebook_name = key
            break
    if nb is None:
        available = ", ".join(sorted(notebooks.keys()))
        return f"Notebook '{notebook_name}' not found. Available: {available}"

    lines = [f"# {notebook_name}\n"]

    for sec_name, sec_info in sorted(nb["sections"].items()):
        filepath = sec_info["latest"]
        texts = _parse_one_file(filepath)

        lines.append(f"## {sec_name}")
        if texts:
            # Show first ~200 chars as preview
            preview = " | ".join(texts)
            if len(preview) > 200:
                preview = preview[:200] + "..."
            lines.append(f"  Preview: {preview}")
        else:
            lines.append("  (no text content)")
        lines.append("")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# OneNote COM API helpers (for writing)
# ---------------------------------------------------------------------------

ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"


def _run_powershell(script: str) -> tuple[bool, str]:
    """Run a PowerShell script and return (success, output)."""
    try:
        result = subprocess.run(
            ["powershell.exe", "-Command", script],
            capture_output=True, text=True, timeout=30,
        )
        output = result.stdout.strip()
        if result.returncode != 0:
            return False, result.stderr.strip() or output
        return True, output
    except subprocess.TimeoutExpired:
        return False, "PowerShell command timed out"
    except FileNotFoundError:
        return False, "PowerShell not found (write features require Windows)"


def _com_get_hierarchy(level: int = 3) -> ET.Element | None:
    """
    Get OneNote hierarchy via COM API.
    Levels: 0=Notebooks, 1=SectionGroups, 2=Sections, 3=Sections(full), 4=Pages
    """
    tmpfile = os.path.join(tempfile.gettempdir(), "onenote_hierarchy.xml")
    # Escape backslashes for PowerShell string
    tmpfile_ps = tmpfile.replace("\\", "\\\\")
    script = (
        f'$onenote = New-Object -ComObject OneNote.Application; '
        f'$h = ""; '
        f'$onenote.GetHierarchy("", {level}, [ref]$h); '
        f'$h | Out-File -FilePath "{tmpfile_ps}" -Encoding UTF8; '
        f'Write-Output "OK"'
    )
    ok, msg = _run_powershell(script)
    if not ok:
        log.warning("COM GetHierarchy failed: %s", msg)
        return None
    try:
        with open(tmpfile, "r", encoding="utf-8-sig") as f:
            xml_content = f.read()
        return ET.fromstring(xml_content)
    except Exception as e:
        log.warning("Failed to parse hierarchy XML: %s", e)
        return None
    finally:
        try:
            os.remove(tmpfile)
        except OSError:
            pass


def _com_find_section_id(notebook_name: str, section_name: str) -> str | None:
    """Find a section ID by notebook and section name (case-insensitive)."""
    root = _com_get_hierarchy(3)
    if root is None:
        return None

    for nb in root.iter(f"{{{ONE_NS}}}Notebook"):
        if nb.get("name", "").lower() != notebook_name.lower():
            continue
        for sec in nb.iter(f"{{{ONE_NS}}}Section"):
            if sec.get("isInRecycleBin") == "true":
                continue
            if sec.get("name", "").lower() == section_name.lower():
                return sec.get("ID")
    return None


def _run_powershell_file(script: str) -> tuple[bool, str]:
    """Write a PowerShell script to a temp file and execute it."""
    ps_file = os.path.join(tempfile.gettempdir(), "onenote_mcp_cmd.ps1")
    try:
        with open(ps_file, "w", encoding="utf-8") as f:
            f.write(script)
        result = subprocess.run(
            ["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", ps_file],
            capture_output=True, text=True, timeout=30,
        )
        output = result.stdout.strip()
        if result.returncode != 0:
            return False, result.stderr.strip() or output
        return True, output
    except subprocess.TimeoutExpired:
        return False, "PowerShell command timed out"
    except FileNotFoundError:
        return False, "PowerShell not found (write features require Windows)"
    finally:
        try:
            os.remove(ps_file)
        except OSError:
            pass


def _com_create_page(section_id: str, title: str, body_html: str) -> tuple[bool, str]:
    """Create a new page in a section using the OneNote COM API."""
    # Escape for PowerShell string literals (double up single quotes)
    section_id_esc = section_id.replace("'", "''")
    title_esc = title.replace("'", "''")
    body_esc = body_html.replace("'", "''")

    script = f'''
$onenote = New-Object -ComObject OneNote.Application
$pageId = ""
$onenote.CreateNewPage('{section_id_esc}', [ref]$pageId, 0)

# Get the new page's XML
$pageXml = ""
$onenote.GetPageContent($pageId, [ref]$pageXml, 0)
$xml = [xml]$pageXml

# Set title
$nsMgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$nsMgr.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$titleNode = $xml.SelectSingleNode("//one:Title/one:OE/one:T", $nsMgr)
if ($titleNode) {{
    $titleNode.InnerXml = "<![CDATA[{title_esc}]]>"
}}

# Update title
$onenote.UpdatePageContent($xml.OuterXml)

# Re-fetch to add body
$pageXml2 = ""
$onenote.GetPageContent($pageId, [ref]$pageXml2, 0)
$xml2 = [xml]$pageXml2

# Add body outline
$outline = $xml2.CreateElement("one", "Outline", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$oeChildren = $xml2.CreateElement("one", "OEChildren", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$oe = $xml2.CreateElement("one", "OE", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$t = $xml2.CreateElement("one", "T", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$cdata = $xml2.CreateCDataSection('{body_esc}')
$t.AppendChild($cdata) | Out-Null
$oe.AppendChild($t) | Out-Null
$oeChildren.AppendChild($oe) | Out-Null
$outline.AppendChild($oeChildren) | Out-Null
$xml2.DocumentElement.AppendChild($outline) | Out-Null

$onenote.UpdatePageContent($xml2.OuterXml)
Write-Output $pageId
'''
    ok, output = _run_powershell_file(script)
    if ok and output:
        return True, f"Page '{title}' created successfully (ID: {output})"
    return False, f"Failed to create page: {output}"


def _com_append_to_page(page_id: str, body_html: str) -> tuple[bool, str]:
    """Append content to an existing page using the OneNote COM API."""
    page_id_esc = page_id.replace("'", "''")
    body_esc = body_html.replace("'", "''")

    script = f'''
$onenote = New-Object -ComObject OneNote.Application
$pageXml = ""
$onenote.GetPageContent('{page_id_esc}', [ref]$pageXml, 0)
$xml = [xml]$pageXml

$outline = $xml.CreateElement("one", "Outline", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$oeChildren = $xml.CreateElement("one", "OEChildren", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$oe = $xml.CreateElement("one", "OE", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$t = $xml.CreateElement("one", "T", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$cdata = $xml.CreateCDataSection('{body_esc}')
$t.AppendChild($cdata) | Out-Null
$oe.AppendChild($t) | Out-Null
$oeChildren.AppendChild($oe) | Out-Null
$outline.AppendChild($oeChildren) | Out-Null
$xml.DocumentElement.AppendChild($outline) | Out-Null

$onenote.UpdatePageContent($xml.OuterXml)
Write-Output "OK"
'''
    ok, output = _run_powershell_file(script)
    if ok:
        return True, "Content appended successfully."
    return False, f"Failed to append content: {output}"


def _com_list_pages(section_id: str) -> list[dict]:
    """List pages in a section via COM API. Returns list of {id, name}."""
    root = _com_get_hierarchy(4)
    if root is None:
        return []

    pages = []
    for sec in root.iter(f"{{{ONE_NS}}}Section"):
        if sec.get("ID") == section_id:
            for page in sec.iter(f"{{{ONE_NS}}}Page"):
                if page.get("isInRecycleBin") == "true":
                    continue
                pages.append({
                    "id": page.get("ID", ""),
                    "name": page.get("name", "(untitled)"),
                })
            break
    return pages


# ---------------------------------------------------------------------------
# MCP Write Tools
# ---------------------------------------------------------------------------


@mcp.tool()
async def list_live_notebooks() -> str:
    """List notebooks from the running OneNote app (live, not backup files).

    This uses the OneNote COM API and shows the notebooks currently open in
    the OneNote desktop app, including their sections. Use this to find
    where to create new pages.
    """
    root = _com_get_hierarchy(3)
    if root is None:
        return "Could not connect to OneNote. Make sure the OneNote desktop app is installed."

    lines = []
    for nb in root.findall(f"{{{ONE_NS}}}Notebook"):
        nb_name = nb.get("name", "?")
        lines.append(f"\n## {nb_name}")
        for sec in nb.iter(f"{{{ONE_NS}}}Section"):
            if sec.get("isInRecycleBin") == "true":
                continue
            sec_name = sec.get("name", "?")
            locked = " (locked)" if sec.get("locked") == "true" else ""
            lines.append(f"  - {sec_name}{locked}")

    if not lines:
        return "No notebooks found in OneNote."
    return "\n".join(lines)


@mcp.tool()
async def create_page(notebook_name: str, section_name: str, title: str, content: str) -> str:
    """Create a new page in a OneNote notebook section.

    The content is written as HTML. You can use basic HTML tags like
    <b>, <i>, <br>, <ul>, <li>, <h1>-<h6>, etc.

    Requires the OneNote desktop app to be installed.

    Args:
        notebook_name: Name of the notebook (from list_live_notebooks).
        section_name: Name of the section within the notebook.
        title: Title for the new page.
        content: The page content (plain text or HTML).
    """
    section_id = _com_find_section_id(notebook_name, section_name)
    if section_id is None:
        return (
            f"Could not find section '{section_name}' in notebook '{notebook_name}'. "
            f"Use list_live_notebooks to see available notebooks and sections."
        )

    ok, msg = _com_create_page(section_id, title, content)
    return msg


@mcp.tool()
async def list_live_pages(notebook_name: str, section_name: str) -> str:
    """List pages in a section from the running OneNote app.

    Use this to find page IDs for appending content to existing pages.

    Args:
        notebook_name: Name of the notebook.
        section_name: Name of the section.
    """
    section_id = _com_find_section_id(notebook_name, section_name)
    if section_id is None:
        return (
            f"Could not find section '{section_name}' in notebook '{notebook_name}'. "
            f"Use list_live_notebooks to see available notebooks and sections."
        )

    pages = _com_list_pages(section_id)
    if not pages:
        return "No pages found in this section."

    lines = []
    for p in pages:
        lines.append(f"- {p['name']}  (id: {p['id']})")
    return "\n".join(lines)


@mcp.tool()
async def append_to_page(page_id: str, content: str) -> str:
    """Append content to an existing OneNote page.

    The content is added as a new outline block at the bottom of the page.
    Supports HTML formatting (<b>, <i>, <br>, <ul>, <li>, etc.).

    Args:
        page_id: The page ID (from list_live_pages).
        content: The content to append (plain text or HTML).
    """
    ok, msg = _com_append_to_page(page_id, content)
    return msg


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    if not ONENOTE_DIR.exists():
        log.error(
            "OneNote backup directory not found: %s\n"
            "Set the ONENOTE_BACKUP_DIR environment variable to the correct path.",
            ONENOTE_DIR,
        )
        sys.exit(1)

    log.info("Starting OneNote MCP server (local files)...")
    log.info("Reading from: %s", ONENOTE_DIR)
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
