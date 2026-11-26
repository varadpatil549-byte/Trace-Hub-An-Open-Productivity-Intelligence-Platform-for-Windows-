"""
TraceHub v0.4 - Windows activity collector

Phase 1 collector for the future TraceHub HTML dashboard.

Captures:
- Username / computer name
- Timestamp
- Current foreground application / process / PID / window title
- Active / Idle + idle seconds
- Browser website name and best-effort page/song title
- Website URL/domain when Windows UI Automation exposes it
- Application / website category
- Productive / Non-Productive classification
- ONLY user-facing background/open applications (not thousands of system processes)
- One CSV file designed to become the Phase 2 HTML dashboard data source
- Small always-on-top live status window with a clearly visible Stop Tracking button

Privacy scope:
- No keystroke recording
- No mouse-coordinate recording
- No screenshots
- No password/page-content capture
- Website capture is limited to browser URL/title metadata exposed by Windows

Windows only.
"""

from __future__ import annotations

import csv
import ctypes
import io
import os
import re
import socket
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple

import psutil
import tkinter as tk
from tkinter import ttk

import win32con
import win32file
import win32gui
import win32process

try:
    import uiautomation as auto
except ImportError:
    auto = None


# ============================================================
# CONFIGURATION
# ============================================================

SAMPLE_SECONDS = 5
IDLE_THRESHOLD_SECONDS = 60

# Performance tuning: expensive Windows UI Automation and background-window
# enumeration are cached instead of being executed on every 5-second sample.
URL_CACHE_SECONDS = 15
BACKGROUND_APPS_CACHE_SECONDS = 10

# Save CSV beside the EXE when packaged with PyInstaller; save beside the .py when
# running from source. This avoids PyInstaller one-file's temporary extraction
# folder, which is deleted when the EXE exits.
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
ACTIVITY_CSV = BASE_DIR / "TraceHub_Activity.csv"

BROWSER_PROCESSES = {
    "chrome.exe", "msedge.exe", "firefox.exe", "brave.exe",
    "opera.exe", "vivaldi.exe"
}

# Keep this list focused on real user-facing application windows.
# It is NOT a list of every Windows process.
SYSTEM_WINDOW_PROCESS_EXCLUSIONS = {
    "applicationframehost.exe",
    "dwm.exe",
    "explorer.exe",  # handled separately as Windows Explorer
    "searchhost.exe",
    "searchapp.exe",
    "shellexperiencehost.exe",
    "startmenuexperiencehost.exe",
    "textinputhost.exe",
    "lockapp.exe",
    "securityhealthsystray.exe",
    "wmiprvse.exe",
    "runtimebroker.exe",
}

# Process-level application classification. The list is intentionally broader
# so common enterprise, AI, development and productivity applications are
# classified instead of falling into "Uncategorized / UNKNOWN".
APPLICATION_RULES: Dict[str, Dict[str, object]] = {
    "Microsoft Office": {
        "processes": {
            "winword.exe", "excel.exe", "powerpnt.exe", "outlook.exe",
            "onenote.exe", "msaccess.exe", "visio.exe", "winproj.exe",
            "msproject.exe"
        },
        "productive": True,
    },
    "AI & Productivity": {
        "processes": {
            "chatgpt.exe", "claude.exe", "copilot.exe", "gemini.exe",
            "perplexity.exe"
        },
        "productive": True,
    },
    "Development": {
        "processes": {
            "code.exe", "devenv.exe", "pycharm64.exe", "pycharm.exe",
            "idea64.exe", "eclipse.exe", "notepad++.exe", "rider64.exe",
            "webstorm64.exe", "clion64.exe", "datagrip64.exe",
            "sublime_text.exe", "atom.exe", "cursor.exe", "windsurf.exe",
            "terminal.exe", "windowsterminal.exe", "cmd.exe",
            "powershell.exe", "pwsh.exe", "git-bash.exe", "python.exe",
            "pythonw.exe", "java.exe"
        },
        "productive": True,
    },
    "Communication": {
        "processes": {
            "teams.exe", "ms-teams.exe", "msteams.exe", "slack.exe",
            "zoom.exe", "webex.exe", "skype.exe", "thunderbird.exe",
            "signal.exe", "whatsapp.exe", "discord.exe"
        },
        "productive": True,
    },
    "ERP & Business": {
        "processes": {
            "saplogon.exe", "saplgpad.exe", "sapgui.exe", "pbi.exe",
            "pbidesktop.exe", "powerbi.exe", "celonis.exe",
            "tableau.exe", "tableaudesktop.exe", "oracle.exe"
        },
        "productive": True,
    },
    "File & Productivity": {
        "processes": {
            "explorer.exe", "acrobat.exe", "acrord32.exe", "notepad.exe",
            "wordpad.exe", "mspaint.exe", "winrar.exe", "7zfm.exe",
            "7zfm.exe", "everything.exe", "sumatrapdf.exe"
        },
        "productive": True,
    },
    "Design & Collaboration": {
        "processes": {
            "figma.exe", "canva.exe", "miro.exe", "drawio.exe"
        },
        "productive": True,
    },
    "Browser": {
        "processes": BROWSER_PROCESSES,
        "productive": None,
    },
    "Media": {
        "processes": {
            "vlc.exe", "wmplayer.exe", "spotify.exe", "music.ui.exe",
            "foobar2000.exe", "itunes.exe", "groove.exe"
        },
        "productive": False,
    },
    "Gaming": {
        "processes": {
            "steam.exe", "steamwebhelper.exe", "epicgameslauncher.exe",
            "riotclientservices.exe", "riotclientux.exe", "gamebar.exe",
            "xboxapp.exe", "battle.net.exe", "goggalaxy.exe"
        },
        "productive": False,
    },
}

# Corporate/internal business domains are handled automatically so new
# Datamatics portals inherit productive classification without adding each URL.
INTERNAL_BUSINESS_DOMAINS = {
    "datamatics.com",
}

# Website/domain classification rules. These intentionally separate business,
# AI, research, collaboration, social, entertainment and shopping activity.
WEBSITE_RULES = {
    "AI & Productivity": {
        "domains": {
            "chatgpt.com", "chat.openai.com", "openai.com", "claude.ai",
            "anthropic.com", "copilot.microsoft.com", "gemini.google.com",
            "bard.google.com", "perplexity.ai", "deepseek.com", "huggingface.co"
        },
        "productive": True,
    },
    "Productivity": {
        "domains": {
            "office.com", "microsoft365.com", "sharepoint.com",
            "onedrive.live.com", "drive.google.com", "docs.google.com",
            "sheets.google.com", "slides.google.com", "notion.so",
            "notion.site", "atlassian.net", "jira.com", "confluence.com",
            "trello.com", "asana.com", "clickup.com", "monday.com",
            "airtable.com", "todoist.com"
        },
        "productive": True,
    },
    "Communication": {
        "domains": {
            "teams.microsoft.com", "slack.com", "zoom.us", "webex.com",
            "mail.google.com", "outlook.office.com", "outlook.live.com",
            "mail.yahoo.com", "proton.me", "protonmail.com"
        },
        "productive": True,
    },
    "Development": {
        "domains": {
            "github.com", "gitlab.com", "bitbucket.org", "stackoverflow.com",
            "developer.mozilla.org", "learn.microsoft.com", "pypi.org",
            "python.org", "npmjs.com", "docker.com", "kubernetes.io",
            "dev.azure.com", "visualstudio.com"
        },
        "productive": True,
    },
    "ERP & Business": {
        "domains": {
            "sap.com", "celonis.com", "powerbi.com", "tableau.com",
            "servicenow.com", "salesforce.com", "workday.com",
            "oracle.com", "dynamics.com"
        },
        "productive": True,
    },
    "Knowledge & Research": {
        "domains": {
            "wikipedia.org", "scholar.google.com", "wolframalpha.com",
            "britannica.com", "researchgate.net", "arxiv.org"
        },
        "productive": True,
    },
    "Design & Collaboration": {
        "domains": {
            "figma.com", "canva.com", "miro.com", "drawio.com",
            "lucidchart.com", "lucid.co"
        },
        "productive": True,
    },
    "Social": {
        "domains": {
            "facebook.com", "instagram.com", "x.com", "twitter.com",
            "reddit.com", "tiktok.com", "snapchat.com", "threads.net",
            "linkedin.com", "pinterest.com"
        },
        "productive": False,
    },
    "Entertainment": {
        "domains": {
            "youtube.com", "youtu.be", "youtube-nocookie.com", "netflix.com",
            "primevideo.com", "hotstar.com", "spotify.com", "twitch.tv",
            "soundcloud.com", "disneyplus.com", "zee5.com"
        },
        "productive": False,
    },
    "Shopping": {
        "domains": {
            "amazon.in", "amazon.com", "flipkart.com", "myntra.com",
            "ebay.com", "meesho.com", "ajio.com", "walmart.com"
        },
        "productive": False,
    },
}

# Title-based website detection. This is particularly important for getting
# a useful website/page name even when the browser address bar cannot be read.
TITLE_SITE_RULES = [
    ("ChatGPT", ("chatgpt", "chat gpt", "openai")),
    ("Claude", ("claude.ai", "anthropic", "claude")),
    ("Microsoft Copilot", ("microsoft copilot", "copilot")),
    ("Google Gemini", ("gemini", "bard")),
    ("Perplexity", ("perplexity",)),
    ("DeepSeek", ("deepseek",)),
    ("YouTube", (" - youtube", "youtube music", "youtube")),
    ("Netflix", ("netflix",)),
    ("Gmail", ("gmail", " - mail.google")),
    ("Google Drive", ("google drive",)),
    ("Google Docs", ("google docs",)),
    ("Google Sheets", ("google sheets",)),
    ("Google Slides", ("google slides",)),
    ("GitHub", ("github",)),
    ("Jira", ("jira",)),
    ("Confluence", ("confluence",)),
    ("Microsoft Teams", ("microsoft teams", "teams")),
    ("Slack", ("slack",)),
    ("WhatsApp", ("whatsapp",)),
    ("Instagram", ("instagram",)),
    ("Facebook", ("facebook",)),
    ("LinkedIn", ("linkedin",)),
    ("Reddit", ("reddit",)),
    ("TikTok", ("tiktok",)),
    ("X", (" - x", "x.com", "twitter")),
    ("Wikipedia", ("wikipedia",)),
    ("Google Search", ("google search", "search - google")),
    ("Microsoft Bing", ("bing",)),
    ("DuckDuckGo", ("duckduckgo",)),
    ("Figma", ("figma",)),
    ("Canva", ("canva",)),
    ("Miro", ("miro",)),
]

# Title-based application detection is the fallback for enterprise software
# whose executable name can vary by version or installation method.
APPLICATION_TITLE_RULES = [
    ("AI & Productivity", ("chatgpt", "openai", "claude", "anthropic", "copilot", "gemini", "perplexity", "deepseek"), True),
    ("ERP & Business", ("sap easy access", "sap logon", "sap gui", "celonis", "power bi", "tableau", "servicenow", "salesforce", "workday"), True),
    ("Development", ("visual studio code", "vs code", "pycharm", "intellij", "eclipse", "notepad++", "powershell", "command prompt", "windows terminal", "git bash"), True),
    ("Communication", ("microsoft teams", "slack", "zoom", "webex", "skype", "outlook", "whatsapp"), True),
    ("Microsoft Office", ("microsoft word", "microsoft excel", "microsoft powerpoint", "microsoft outlook", "microsoft access", "microsoft onenote", "microsoft visio"), True),
    ("File & Productivity", ("file explorer", "windows explorer", "adobe acrobat", "notepad"), True),
    ("Design & Collaboration", ("figma", "canva", "miro", "draw.io", "lucidchart"), True),
    ("Media", ("spotify", "vlc media player", "windows media player", "youtube music"), False),
    ("Gaming", ("steam", "epic games", "riot client", "xbox"), False),
]

FRIENDLY_PROCESS_NAMES = {
    "winword.exe": "Microsoft Word",
    "excel.exe": "Microsoft Excel",
    "powerpnt.exe": "Microsoft PowerPoint",
    "outlook.exe": "Microsoft Outlook",
    "onenote.exe": "Microsoft OneNote",
    "msaccess.exe": "Microsoft Access",
    "visio.exe": "Microsoft Visio",
    "chatgpt.exe": "ChatGPT",
    "claude.exe": "Claude",
    "copilot.exe": "Microsoft Copilot",
    "gemini.exe": "Google Gemini",
    "perplexity.exe": "Perplexity",
    "chrome.exe": "Google Chrome",
    "msedge.exe": "Microsoft Edge",
    "firefox.exe": "Mozilla Firefox",
    "brave.exe": "Brave",
    "opera.exe": "Opera",
    "vivaldi.exe": "Vivaldi",
    "code.exe": "Visual Studio Code",
    "cursor.exe": "Cursor",
    "windsurf.exe": "Windsurf",
    "devenv.exe": "Visual Studio",
    "teams.exe": "Microsoft Teams",
    "ms-teams.exe": "Microsoft Teams",
    "msteams.exe": "Microsoft Teams",
    "slack.exe": "Slack",
    "powershell.exe": "PowerShell",
    "pwsh.exe": "PowerShell",
    "cmd.exe": "Command Prompt",
    "windowsterminal.exe": "Windows Terminal",
    "git-bash.exe": "Git Bash",
    "notepad.exe": "Notepad",
    "spotify.exe": "Spotify",
    "vlc.exe": "VLC",
    "explorer.exe": "Windows Explorer",
    "acrobat.exe": "Adobe Acrobat",
    "acrord32.exe": "Adobe Acrobat Reader",
    "saplogon.exe": "SAP Logon",
    "saplgpad.exe": "SAP Logon",
    "sapgui.exe": "SAP GUI",
    "pbidesktop.exe": "Power BI Desktop",
    "powerbi.exe": "Power BI",
    "tableau.exe": "Tableau",
}

CSV_COLUMNS = [
    "Timestamp",
    "Date",
    "Time",
    "Username",
    "ComputerName",
    "ActivityStatus",
    "IdleSeconds",
    "ForegroundApplication",
    "ProcessName",
    "PID",
    "WindowTitle",
    "ApplicationCategory",
    "Productivity",
    "WebsiteName",
    "PageTitle",
    "WebsiteURL",
    "WebsiteDomain",
    "WebsiteCategory",
    "BackgroundOpenApps",
    "BackgroundAppCount",
    "EventType",
    "TerminationReason",
]


# ============================================================
# USER / IDLE
# ============================================================

def get_username() -> str:
    return os.environ.get("USERNAME") or os.environ.get("USER") or "Unknown"


def get_computer_name() -> str:
    return socket.gethostname()


def get_windows_idle_seconds() -> int:
    """Return seconds since the last keyboard/mouse input using Windows API."""
    class LASTINPUTINFO(ctypes.Structure):
        _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]

    try:
        info = LASTINPUTINFO()
        info.cbSize = ctypes.sizeof(LASTINPUTINFO)
        if ctypes.windll.user32.GetLastInputInfo(ctypes.byref(info)) == 0:
            return 0
        tick = ctypes.windll.kernel32.GetTickCount()
        elapsed_ms = (tick - info.dwTime) & 0xFFFFFFFF
        return max(0, int(elapsed_ms / 1000))
    except Exception:
        return 0


# ============================================================
# FOREGROUND WINDOW
# ============================================================

def friendly_application_name(process_name: str) -> str:
    p = (process_name or "").lower()
    if p in FRIENDLY_PROCESS_NAMES:
        return FRIENDLY_PROCESS_NAMES[p]
    if not p:
        return "Unknown"
    return os.path.splitext(p)[0].replace("_", " ").title()


def get_foreground_window_details() -> Dict[str, object]:
    result = {
        "hwnd": 0,
        "window_title": "",
        "process_name": "",
        "pid": 0,
        "application": "Unknown",
    }
    try:
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            return result

        result["hwnd"] = hwnd
        result["window_title"] = win32gui.GetWindowText(hwnd) or ""

        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        result["pid"] = pid

        try:
            proc = psutil.Process(pid)
            name = proc.name() or ""
            result["process_name"] = name
            result["application"] = friendly_application_name(name)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    except Exception:
        pass
    return result


# ============================================================
# WEBSITE / BROWSER TITLE
# ============================================================

def clean_possible_url(text: str) -> str:
    if not text:
        return ""
    text = text.strip()
    match = re.search(
        r"(?i)(https?://)?((?:[a-z0-9-]+\.)+[a-z]{2,})(?::\d+)?(?:[/\?#][^\s]*)?",
        text,
    )
    if not match:
        return ""
    domain = match.group(2).lower()
    if domain in {"microsoft.com", "google.com"} and not match.group(0).lower().startswith("http"):
        # Avoid treating generic words in titles as URLs.
        return ""
    full = match.group(0)
    return full if full.lower().startswith(("http://", "https://")) else "https://" + full


def extract_domain(url: str) -> str:
    if not url:
        return ""
    value = re.sub(r"^https?://", "", url.lower().strip())
    value = value.split("/")[0].split(":")[0]
    return value[4:] if value.startswith("www.") else value


def domain_matches(domain: str, candidate: str) -> bool:
    return domain == candidate or domain.endswith("." + candidate)


def classify_website(domain: str) -> Tuple[str, str]:
    if not domain:
        return "Web Browsing", "NEUTRAL"

    # Any Datamatics subdomain is treated as an internal business portal.
    # Example: epm.datamatics.com, portal.datamatics.com, rpa.datamatics.com
    for candidate in INTERNAL_BUSINESS_DOMAINS:
        if domain_matches(domain, candidate):
            return "Internal Business", "PRODUCTIVE"

    for category, rule in WEBSITE_RULES.items():
        for candidate in rule["domains"]:
            if domain_matches(domain, candidate):
                return category, "PRODUCTIVE" if rule["productive"] else "NON-PRODUCTIVE"
    return "Web Browsing", "NEUTRAL"


# Small caches keep UI Automation from being queried every five seconds when
# the user remains on the same browser page.
_browser_url_cache: Dict[int, Tuple[float, str, str]] = {}
_browser_url_cache_lock = threading.Lock()


def get_browser_url_uia(hwnd: int, window_title: str = "") -> str:
    """Best-effort URL extraction from Chrome/Edge/Firefox UI Automation.

    The old implementation depended almost entirely on one localized control
    name. On some enterprise Chrome builds that control is exposed with a
    different Name/AutomationId, which caused the URL to come back blank and
    therefore made internal portals fall back to NEUTRAL.

    This version first tries the common address-bar names and then performs a
    small bounded descendant scan for an EditControl whose value actually
    looks like a URL/domain. It never focuses, clicks, or types into the browser.
    """
    if not hwnd or auto is None:
        return ""

    now = time.monotonic()
    with _browser_url_cache_lock:
        cached = _browser_url_cache.get(hwnd)
        if cached:
            cached_at, cached_title, cached_url = cached
            if cached_title == window_title and (now - cached_at) < URL_CACHE_SECONDS:
                return cached_url

    result = ""
    try:
        root = auto.ControlFromHandle(hwnd)
        if root is None:
            return ""

        # Fast path: common Chrome/Edge address-bar names.
        names = [
            "Address and search bar",
            "Address bar",
            "Search or enter web address",
            "Search or enter address",
            "Search or enter web address",
        ]

        for name in names:
            try:
                ctrl = root.EditControl(searchDepth=8, Name=name)
                if ctrl.Exists(0.20):
                    try:
                        value = ctrl.GetValuePattern().Value or ""
                        url = clean_possible_url(value)
                        if url:
                            result = url
                            break
                    except Exception:
                        pass
            except Exception:
                continue

        # Robust fallback: inspect a bounded number of descendants and accept
        # only EditControls whose ValuePattern contains an actual domain/URL.
        if not result:
            def walk(control: object, depth: int = 0) -> str:
                if depth > 12:
                    return ""

                try:
                    control_type = str(getattr(control, "ControlTypeName", ""))
                except Exception:
                    control_type = ""

                if control_type == "EditControl":
                    value = ""
                    try:
                        value = control.GetValuePattern().Value or ""
                    except Exception:
                        try:
                            pattern = control.GetPattern(auto.PatternId.ValuePattern)
                            value = (pattern.Value or "") if pattern else ""
                        except Exception:
                            value = ""

                    url = clean_possible_url(value)
                    if url:
                        return url

                try:
                    children = control.GetChildren()
                except Exception:
                    children = []

                # Browser chrome is usually shallow. Stop after a small number
                # of children so this fallback remains cheap.
                for child in children[:80]:
                    found = walk(child, depth + 1)
                    if found:
                        return found
                return ""

            result = walk(root)

    except Exception:
        result = ""

    with _browser_url_cache_lock:
        _browser_url_cache[hwnd] = (now, window_title, result)
        # Keep the cache tiny in long-running sessions.
        if len(_browser_url_cache) > 128:
            oldest = sorted(_browser_url_cache.items(), key=lambda item: item[1][0])[:32]
            for old_hwnd, _ in oldest:
                _browser_url_cache.pop(old_hwnd, None)

    return result


def detect_site_from_title(title: str) -> str:
    value = f" {(title or '').lower()} "
    for site, keywords in TITLE_SITE_RULES:
        if any(keyword.lower() in value for keyword in keywords):
            return site
    return ""


def strip_browser_suffix(title: str) -> str:
    """Turn 'Song - YouTube - Google Chrome' into 'Song'."""
    title = (title or "").strip()
    if not title:
        return ""

    suffixes = [
        " - Google Chrome",
        " - Microsoft Edge",
        " - Mozilla Firefox",
        " - Brave",
        " - Opera",
        " - Vivaldi",
        " - YouTube",
        " - YouTube Music",
        " | YouTube",
        " | YouTube Music",
    ]

    cleaned = title
    for suffix in suffixes:
        if cleaned.lower().endswith(suffix.lower()):
            cleaned = cleaned[:-len(suffix)].strip()

    # A second pass handles 'Song - YouTube - Google Chrome'.
    browser_markers = [
        " - YouTube", " | YouTube", " - YouTube Music", " | YouTube Music"
    ]
    for marker in browser_markers:
        idx = cleaned.lower().find(marker.lower())
        if idx > 0:
            cleaned = cleaned[:idx].strip()
            break

    return cleaned or title


def website_classification_from_name(website_name: str) -> Tuple[str, str]:
    mapping = {
        "Datamatics Internal": ("Internal Business", "PRODUCTIVE"),
        "ChatGPT": ("AI & Productivity", "PRODUCTIVE"),
        "Claude": ("AI & Productivity", "PRODUCTIVE"),
        "Microsoft Copilot": ("AI & Productivity", "PRODUCTIVE"),
        "Google Gemini": ("AI & Productivity", "PRODUCTIVE"),
        "Perplexity": ("AI & Productivity", "PRODUCTIVE"),
        "DeepSeek": ("AI & Productivity", "PRODUCTIVE"),
        "YouTube": ("Entertainment", "NON-PRODUCTIVE"),
        "Netflix": ("Entertainment", "NON-PRODUCTIVE"),
        "Instagram": ("Social", "NON-PRODUCTIVE"),
        "Facebook": ("Social", "NON-PRODUCTIVE"),
        "Reddit": ("Social", "NON-PRODUCTIVE"),
        "TikTok": ("Social", "NON-PRODUCTIVE"),
        "X": ("Social", "NON-PRODUCTIVE"),
        "LinkedIn": ("Social", "NEUTRAL"),
        "Gmail": ("Communication", "PRODUCTIVE"),
        "Microsoft Teams": ("Communication", "PRODUCTIVE"),
        "Slack": ("Communication", "PRODUCTIVE"),
        "WhatsApp": ("Communication", "PRODUCTIVE"),
        "Google Drive": ("Productivity", "PRODUCTIVE"),
        "Google Docs": ("Productivity", "PRODUCTIVE"),
        "Google Sheets": ("Productivity", "PRODUCTIVE"),
        "Google Slides": ("Productivity", "PRODUCTIVE"),
        "GitHub": ("Development", "PRODUCTIVE"),
        "Jira": ("Productivity", "PRODUCTIVE"),
        "Confluence": ("Productivity", "PRODUCTIVE"),
        "Wikipedia": ("Knowledge & Research", "PRODUCTIVE"),
        "Google Search": ("Knowledge & Research", "PRODUCTIVE"),
        "Microsoft Bing": ("Knowledge & Research", "PRODUCTIVE"),
        "DuckDuckGo": ("Knowledge & Research", "PRODUCTIVE"),
        "Figma": ("Design & Collaboration", "PRODUCTIVE"),
        "Canva": ("Design & Collaboration", "PRODUCTIVE"),
        "Miro": ("Design & Collaboration", "PRODUCTIVE"),
    }
    return mapping.get(website_name, ("Web Browsing", "NEUTRAL"))


def get_website_info(process_name: str, hwnd: int, window_title: str) -> Tuple[str, str, str, str, str, str]:
    """Return website name, page title, URL, domain, category, productivity."""
    process = (process_name or "").lower()
    is_browser = process in BROWSER_PROCESSES

    if not is_browser:
        return "", "", "", "", "", "UNKNOWN"

    url = get_browser_url_uia(hwnd, window_title)
    domain = extract_domain(url)

    title_site = detect_site_from_title(window_title)
    website_name = title_site

    if domain:
        pretty = {
            "chatgpt.com": "ChatGPT", "chat.openai.com": "ChatGPT", "openai.com": "OpenAI",
            "claude.ai": "Claude", "anthropic.com": "Claude",
            "copilot.microsoft.com": "Microsoft Copilot", "gemini.google.com": "Google Gemini",
            "bard.google.com": "Google Gemini", "perplexity.ai": "Perplexity",
            "deepseek.com": "DeepSeek", "youtube.com": "YouTube", "youtu.be": "YouTube",
            "netflix.com": "Netflix", "github.com": "GitHub", "instagram.com": "Instagram",
            "facebook.com": "Facebook", "linkedin.com": "LinkedIn", "reddit.com": "Reddit",
            "mail.google.com": "Gmail", "drive.google.com": "Google Drive",
            "docs.google.com": "Google Docs", "sheets.google.com": "Google Sheets",
            "slack.com": "Slack", "teams.microsoft.com": "Microsoft Teams",
            "spotify.com": "Spotify", "twitch.tv": "Twitch", "wikipedia.org": "Wikipedia",
            "figma.com": "Figma", "canva.com": "Canva", "miro.com": "Miro",
        }
        if any(domain_matches(domain, candidate) for candidate in INTERNAL_BUSINESS_DOMAINS):
            website_name = "Datamatics Internal"
        else:
            website_name = pretty.get(domain, domain)

    # Prefer title/site recognition because it remains available when the URL
    # is blocked by browser security or UI Automation.
    if title_site:
        website_name = title_site

    page_title = strip_browser_suffix(window_title)

    markers = [
        " - YouTube", " | YouTube", " — YouTube", " – YouTube",
        " - YouTube Music", " | YouTube Music",
        " - Google Chrome", " - Microsoft Edge", " - Mozilla Firefox",
        " - Brave", " - Opera", " - Vivaldi",
    ]
    for marker in markers:
        idx = page_title.lower().find(marker.lower())
        if idx > 0:
            page_title = page_title[:idx].strip()
            break

    website_category, website_productivity = classify_website(domain)

    # Strong title-based fallback when the URL cannot be read.
    name_category, name_productivity = website_classification_from_name(website_name)
    if website_name:
        website_category, website_productivity = name_category, name_productivity

    # Enterprise-page fallback for internal portals when Chrome/Edge does not
    # expose the address-bar value to UI Automation. These are intentionally
    # business-process terms rather than a list of individual portal URLs.
    internal_title_hints = (
        "invoice processing", "invoice management", "purchase order",
        "accounts payable", "accounts receivable", "vendor management",
        "procurement", "grir", "vim workplace", "vim analytics",
        "erp", "enterprise portal", "business portal", "rpa portal"
    )
    title_lower = (window_title or "").lower()
    if not domain and any(hint in title_lower for hint in internal_title_hints):
        website_name = "Datamatics Internal"
        website_category = "Internal Business"
        website_productivity = "PRODUCTIVE"

    # Site-specific corrections.
    if website_name == "Spotify":
        website_category, website_productivity = "Entertainment", "NON-PRODUCTIVE"
    elif website_name == "OpenAI":
        website_category, website_productivity = "AI & Productivity", "PRODUCTIVE"
    elif website_name == "LinkedIn":
        website_category, website_productivity = "Social", "NEUTRAL"

    return website_name, page_title, url, domain, website_category, website_productivity


# ============================================================
# APPLICATION CLASSIFICATION
# ============================================================

def classify_application(process_name: str, application: str, window_title: str = "") -> Tuple[str, str]:
    p = (process_name or "").lower()

    for category, rule in APPLICATION_RULES.items():
        if p in rule["processes"]:
            productive = rule["productive"]
            if productive is None:
                return category, "NEUTRAL"
            return category, "PRODUCTIVE" if productive else "NON-PRODUCTIVE"

    text = f"{application} {process_name} {window_title}".lower()
    for category, keywords, productive in APPLICATION_TITLE_RULES:
        if any(keyword.lower() in text for keyword in keywords):
            return category, "PRODUCTIVE" if productive else "NON-PRODUCTIVE"

    if any(x in text for x in ("calculator", "clock", "settings", "task manager", "snipping tool")):
        return "System & Utilities", "NEUTRAL"
    if any(x in text for x in ("remote desktop", "mstsc", "citrix", "vmware horizon")):
        return "Enterprise Access", "PRODUCTIVE"

    return "Uncategorized", "UNKNOWN"


# ============================================================
# BACKGROUND OPEN APPLICATIONS
# ============================================================

_background_cache_lock = threading.Lock()
_background_cache_at = 0.0
_background_cache_foreground_pid = 0
_background_cache_values: List[str] = []


def get_visible_open_applications(foreground_pid: int) -> List[str]:
    """Get user-facing apps with visible top-level windows, not all processes."""
    global _background_cache_at, _background_cache_foreground_pid, _background_cache_values

    now = time.monotonic()
    with _background_cache_lock:
        if (
            _background_cache_at > 0
            and _background_cache_foreground_pid == foreground_pid
            and (now - _background_cache_at) < BACKGROUND_APPS_CACHE_SECONDS
        ):
            return list(_background_cache_values)

    results: Dict[int, str] = {}

    def callback(hwnd: int, _extra: object) -> bool:
        try:
            if not win32gui.IsWindowVisible(hwnd) or win32gui.IsIconic(hwnd):
                return True

            title = (win32gui.GetWindowText(hwnd) or "").strip()
            if not title:
                return True

            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if not pid or pid == foreground_pid or pid == os.getpid():
                return True

            try:
                proc = psutil.Process(pid)
                process_name = (proc.name() or "").lower()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                return True

            if process_name in SYSTEM_WINDOW_PROCESS_EXCLUSIONS:
                # Explorer is a real open app when it has a visible named window.
                if process_name == "explorer.exe":
                    app = "Windows Explorer"
                else:
                    return True
            else:
                app = friendly_application_name(process_name)

            results[pid] = app
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(callback, None)
    except Exception:
        pass

    values = sorted(set(results.values()), key=str.lower)
    with _background_cache_lock:
        _background_cache_at = now
        _background_cache_foreground_pid = foreground_pid
        _background_cache_values = list(values)
    return values


# ============================================================
# CSV
# ============================================================

class LockedCsvWriter:
    """Keep TraceHub_Activity.csv open with Windows share mode = 0.

    This makes the live CSV inaccessible to normal user applications while
    TraceHub is running. The OS releases the handle automatically if TraceHub
    is terminated unexpectedly.
    """

    def __init__(self, path: Path):
        self.path = path
        self.handle = None
        self.records_since_flush = 0

    def _write_bytes(self, payload: bytes) -> None:
        if self.handle is None:
            raise RuntimeError("CSV file is not open")
        win32file.WriteFile(self.handle, payload)

    def _write_text(self, text: str) -> None:
        self._write_bytes(text.encode("utf-8"))

    def open(self) -> None:
        prepare_csv_schema()
        try:
            self.handle = win32file.CreateFile(
                str(self.path),
                win32con.GENERIC_READ | win32con.GENERIC_WRITE,
                0,  # IMPORTANT: no read/write/delete sharing while tracking.
                None,
                win32con.OPEN_ALWAYS,
                win32con.FILE_ATTRIBUTE_NORMAL,
                None,
            )
        except Exception as exc:
            raise RuntimeError(
                "TraceHub could not lock TraceHub_Activity.csv. "
                "Close the CSV in Excel/Notepad and start TraceHub again."
            ) from exc

        size = win32file.GetFileSize(self.handle)
        win32file.SetFilePointer(self.handle, 0, win32file.FILE_END)

        if size == 0:
            header_buffer = io.StringIO()
            writer = csv.writer(header_buffer, lineterminator="\r\n")
            writer.writerow(CSV_COLUMNS)
            self._write_bytes(b"\xef\xbb\xbf" + header_buffer.getvalue().encode("utf-8"))
            win32file.FlushFileBuffers(self.handle)

    def append(self, row: Dict[str, object]) -> None:
        if self.handle is None:
            raise RuntimeError("CSV file is not open")
        clean = [row.get(column, "") for column in CSV_COLUMNS]
        buffer = io.StringIO()
        writer = csv.writer(buffer, lineterminator="\r\n")
        writer.writerow(clean)
        self._write_text(buffer.getvalue())
        self.records_since_flush += 1

        # One flush per minute keeps crash recovery reasonable without forcing
        # a disk flush on every 5-second sample.
        if self.records_since_flush >= 12:
            win32file.FlushFileBuffers(self.handle)
            self.records_since_flush = 0

    def close(self) -> None:
        if self.handle is not None:
            try:
                win32file.FlushFileBuffers(self.handle)
            except Exception:
                pass
            try:
                win32file.CloseHandle(self.handle)
            except Exception:
                pass
            self.handle = None


def prepare_csv_schema() -> None:
    """Archive an older TraceHub CSV if its header doesn't match v0.4."""
    if not ACTIVITY_CSV.exists() or ACTIVITY_CSV.stat().st_size == 0:
        return

    try:
        with ACTIVITY_CSV.open("r", newline="", encoding="utf-8-sig") as f:
            header = next(csv.reader(f), [])
    except Exception as exc:
        raise RuntimeError(f"Cannot inspect {ACTIVITY_CSV.name}: {exc}") from exc

    if header == CSV_COLUMNS:
        return

    archive_name = f"TraceHub_Activity_Archive_{datetime.now():%Y%m%d_%H%M%S}.csv"
    archive_path = ACTIVITY_CSV.with_name(archive_name)
    ACTIVITY_CSV.replace(archive_path)
    print(f"[TraceHub] Older CSV schema archived as: {archive_path.name}")


# ============================================================
# LIVE MINI DASHBOARD
# ============================================================

class TraceHubMiniWindow:
    def __init__(self) -> None:
        self.running = True
        self.start_time = datetime.now()
        self.records = 0
        self.manual_stop_event = threading.Event()
        self.manual_stop_time = None
        self.manual_stop_reason = ""

        self.root = tk.Tk()
        self.root.title("TraceHub • Live Monitor")
        self.root.geometry("410x520")
        self.root.minsize(410, 520)
        self.root.maxsize(410, 520)
        self.root.attributes("-topmost", True)
        self.root.configure(bg="#10151c")

        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.window_close)

    def _build_ui(self) -> None:
        bg = "#10151c"
        panel = "#171e27"
        text = "#ecf2f8"
        muted = "#9eacb9"
        green = "#39d98a"
        red = "#ff5c69"

        outer = tk.Frame(self.root, bg=bg, padx=14, pady=12)
        outer.pack(fill="both", expand=True)

        header = tk.Frame(outer, bg=bg)
        header.pack(fill="x")

        tk.Label(
            header, text="TRACE HUB", font=("Segoe UI", 18, "bold"),
            fg=text, bg=bg
        ).pack(side="left")

        self.live_dot = tk.Label(
            header, text="●  LIVE", font=("Segoe UI", 10, "bold"),
            fg=green, bg=bg
        )
        self.live_dot.pack(side="right", pady=3)

        tk.Label(
            outer, text="Endpoint activity monitor", font=("Segoe UI", 9),
            fg=muted, bg=bg
        ).pack(anchor="w", pady=(0, 12))

        # Status card
        card = tk.Frame(outer, bg=panel, padx=13, pady=11)
        card.pack(fill="x")

        self.status_var = tk.StringVar(value="ACTIVE")
        self.status_label = tk.Label(
            card, textvariable=self.status_var, font=("Segoe UI", 13, "bold"),
            fg=green, bg=panel
        )
        self.status_label.pack(anchor="w")

        self.app_var = tk.StringVar(value="Current App  •  -")
        self.site_var = tk.StringVar(value="Website  •  -")
        self.page_var = tk.StringVar(value="Page •  -")
        self.cat_var = tk.StringVar(value="Category  •  -")
        self.prod_var = tk.StringVar(value="Productivity  •  -")
        self.idle_var = tk.StringVar(value="Idle  •  0 sec")

        for var in [self.app_var, self.site_var, self.page_var, self.cat_var, self.prod_var, self.idle_var]:
            tk.Label(
                card, textvariable=var, font=("Segoe UI", 9),
                fg=text, bg=panel, anchor="w"
            ).pack(fill="x", pady=2)

        # Background apps card
        bgcard = tk.Frame(outer, bg=panel, padx=13, pady=10)
        bgcard.pack(fill="x", pady=10)

        tk.Label(
            bgcard, text="OPEN BACKGROUND APPS", font=("Segoe UI", 8, "bold"),
            fg=muted, bg=panel
        ).pack(anchor="w")

        self.background_var = tk.StringVar(value="-  •  0 apps")
        tk.Label(
            bgcard, textvariable=self.background_var, font=("Segoe UI", 9),
            fg=text, bg=panel, anchor="w", justify="left", wraplength=445
        ).pack(fill="x", pady=(5, 0))

        # Footer stats
        stats = tk.Frame(outer, bg=bg)
        stats.pack(fill="x", pady=(0, 10))
        self.records_var = tk.StringVar(value="Records: 0")
        self.runtime_var = tk.StringVar(value="Running: 00:00:00")
        self.csv_var = tk.StringVar(value=f"CSV: {ACTIVITY_CSV.name}")

        tk.Label(stats, textvariable=self.records_var, font=("Segoe UI", 9), fg=muted, bg=bg).pack(side="left")
        tk.Label(stats, textvariable=self.runtime_var, font=("Segoe UI", 9), fg=muted, bg=bg).pack(side="right")
        tk.Label(outer, textvariable=self.csv_var, font=("Segoe UI", 8), fg="#71808f", bg=bg, anchor="w").pack(fill="x", pady=(0, 10))

        # Strong, always-visible stop button.
        # Compact Stop Tracking button.
        self.stop_button = tk.Button(
            outer,
            text="■  STOP TRACKING",
            command=self.manual_stop,
            font=("Segoe UI", 9, "bold"),
            fg="white",
            bg="#d94755",
            activeforeground="white",
            activebackground="#ef5967",
            relief="flat",
            bd=0,
            cursor="hand2",
            width=17,
            height=1,
            padx=6,
            pady=2,
        )
        self.stop_button.pack(anchor="center", pady=(2, 4))

    def update_data(self, row: Dict[str, object]) -> None:
        try:
            status = str(row.get("ActivityStatus", "UNKNOWN"))
            status_color = "#39d98a" if status == "ACTIVE" else "#ffbf3f"
            if status == "IDLE":
                status_color = "#ff8a4c"

            self.status_var.set(status)
            self.status_label.configure(fg=status_color)
            self.live_dot.configure(text="●  LIVE", fg="#39d98a")

            app = str(row.get("ForegroundApplication") or "-")
            website = str(row.get("WebsiteName") or "-")
            page = str(row.get("PageTitle") or "-")
            category = str(row.get("ApplicationCategory") or "-")
            productivity = str(row.get("Productivity") or "-")
            idle = str(row.get("IdleSeconds") or 0)
            backgrounds = str(row.get("BackgroundOpenApps") or "-")
            bg_count = str(row.get("BackgroundAppCount") or 0)

            self.app_var.set(f"Current App  •  {app}")
            self.site_var.set(f"Website  •  {website}")
            self.page_var.set(f"Page  •  {page}")
            self.cat_var.set(f"Category  •  {category}")
            self.prod_var.set(f"Productivity  •  {productivity}")
            self.idle_var.set(f"Idle  •  {idle} sec")
            self.background_var.set(f"{backgrounds}  •  {bg_count} apps")

            self.records += 1
            self.records_var.set(f"Records: {self.records:,}")
            elapsed = int((datetime.now() - self.start_time).total_seconds())
            h, rem = divmod(elapsed, 3600)
            m, s = divmod(rem, 60)
            self.runtime_var.set(f"Running: {h:02d}:{m:02d}:{s:02d}")
        except tk.TclError:
            pass

    def _request_stop(self, reason: str) -> None:
        if not self.running:
            return
        self.manual_stop_time = datetime.now()
        self.manual_stop_reason = reason
        self.manual_stop_event.set()
        self.running = False
        try:
            self.root.after(0, self.root.destroy)
        except Exception:
            pass

    def manual_stop(self) -> None:
        self._request_stop("User clicked STOP TRACKING")

    def window_close(self) -> None:
        self._request_stop("User closed the TraceHub window")

    def stop(self) -> None:
        # Backward-compatible helper for Ctrl+C / internal callers.
        self._request_stop("Tracking stopped by user")

    def run(self) -> None:
        self.root.mainloop()


# ============================================================
# COLLECTION ENGINE
# ============================================================

def create_activity_record() -> Dict[str, object]:
    now = datetime.now()
    username = get_username()
    computer = get_computer_name()

    window = get_foreground_window_details()
    hwnd = int(window["hwnd"] or 0)
    pid = int(window["pid"] or 0)
    process_name = str(window["process_name"] or "")
    application = str(window["application"] or "Unknown")
    window_title = str(window["window_title"] or "")

    idle_seconds = get_windows_idle_seconds()
    activity_status = "IDLE" if idle_seconds >= IDLE_THRESHOLD_SECONDS else "ACTIVE"

    app_category, app_productivity = classify_application(process_name, application, window_title)

    # Website fields are deliberately independent from app fields so the
    # dashboard can show both 'Google Chrome' and 'YouTube / Song Name'.
    website_name, page_title, website_url, website_domain, website_category, website_productivity = (
        get_website_info(process_name, hwnd, window_title)
    )

    final_category = app_category
    final_productivity = app_productivity

    if website_name:
        final_category = website_category or app_category
        if website_productivity != "UNKNOWN":
            final_productivity = website_productivity

    background_apps = get_visible_open_applications(pid)
    background_summary = ", ".join(background_apps) if background_apps else "None"

    return {
        "Timestamp": now.strftime("%Y-%m-%d %H:%M:%S"),
        "Date": now.strftime("%Y-%m-%d"),
        "Time": now.strftime("%H:%M:%S"),
        "Username": username,
        "ComputerName": computer,
        "ActivityStatus": activity_status,
        "IdleSeconds": idle_seconds,
        "ForegroundApplication": application,
        "ProcessName": process_name,
        "PID": pid,
        "WindowTitle": window_title,
        "ApplicationCategory": final_category,
        "Productivity": final_productivity,
        "WebsiteName": website_name,
        "PageTitle": page_title,
        "WebsiteURL": website_url,
        "WebsiteDomain": website_domain,
        "WebsiteCategory": website_category,
        "BackgroundOpenApps": background_summary,
        "BackgroundAppCount": len(background_apps),
        "EventType": "ACTIVITY",
        "TerminationReason": "",
    }


def create_manual_termination_record(mini: TraceHubMiniWindow) -> Dict[str, object]:
    stamp = mini.manual_stop_time or datetime.now()
    return {
        "Timestamp": stamp.strftime("%Y-%m-%d %H:%M:%S"),
        "Date": stamp.strftime("%Y-%m-%d"),
        "Time": stamp.strftime("%H:%M:%S"),
        "Username": get_username(),
        "ComputerName": get_computer_name(),
        "ActivityStatus": "MANUALLY TERMINATED",
        "IdleSeconds": "",
        "ForegroundApplication": "TraceHub",
        "ProcessName": Path(sys.argv[0]).name,
        "PID": os.getpid(),
        "WindowTitle": "TraceHub • User terminated tracking",
        "ApplicationCategory": "System & Utilities",
        "Productivity": "NEUTRAL",
        "WebsiteName": "",
        "PageTitle": "",
        "WebsiteURL": "",
        "WebsiteDomain": "",
        "WebsiteCategory": "",
        "BackgroundOpenApps": "",
        "BackgroundAppCount": "",
        "EventType": "USER_TERMINATION",
        "TerminationReason": mini.manual_stop_reason or "User stopped tracking",
    }


def collector_loop(mini: TraceHubMiniWindow) -> None:
    uia_context = None
    csv_writer = LockedCsvWriter(ACTIVITY_CSV)

    if auto is not None:
        try:
            # uiautomation explicitly requires this object when its controls
            # are used from a worker thread.
            uia_context = auto.UIAutomationInitializerInThread(debug=False)
        except Exception as exc:
            print(f"[TraceHub] UI Automation thread initialization unavailable: {exc}")
            uia_context = None

    print("=" * 100)
    print("TRACEHUB v0.4  |  Windows activity collector")
    print("=" * 100)
    print(f"User       : {get_username()}")
    print(f"Computer   : {get_computer_name()}")
    print(f"Interval   : {SAMPLE_SECONDS} seconds")
    print(f"Idle limit : {IDLE_THRESHOLD_SECONDS} seconds")
    print(f"CSV        : {ACTIVITY_CSV}")
    print("CSV lock   : EXCLUSIVE while TraceHub is running")
    print("Background : visible/open user apps only (cached every 10 sec)")
    print("URL/UIA    : cached for 15 sec to reduce Windows UI Automation overhead")
    print("Press STOP TRACKING in the TraceHub window or Ctrl+C in terminal.")
    print("=" * 100)

    try:
        csv_writer.open()

        while mini.running:
            try:
                row = create_activity_record()

                # If the user clicked Stop while this sample was being built,
                # don't create a normal post-stop activity row.
                if mini.manual_stop_event.is_set():
                    break

                csv_writer.append(row)

                print(
                    f"[{row['Timestamp']}] "
                    f"{row['ActivityStatus']:<6} | "
                    f"{str(row['ForegroundApplication'])[:22]:<22} | "
                    f"Site={str(row['WebsiteName'])[:16]:<16} | "
                    f"Page={str(row['PageTitle'])[:28]:<28} | "
                    f"{row['Productivity']:<13} | "
                    f"BG Apps={row['BackgroundAppCount']}"
                )

                try:
                    mini.root.after(0, mini.update_data, row)
                except Exception:
                    pass

            except Exception as exc:
                print(f"[TraceHub] Collection error: {exc}")

            # Use small sleeps so the Stop button responds quickly without
            # waking the collector unnecessarily.
            for _ in range(max(1, SAMPLE_SECONDS * 10)):
                if not mini.running:
                    break
                time.sleep(0.1)

        # Manual termination is an explicit audit event. A system crash or
        # power loss will not manufacture this row because this code isn't run.
        if mini.manual_stop_event.is_set():
            termination_row = create_manual_termination_record(mini)
            try:
                csv_writer.append(termination_row)
                print(
                    f"[{termination_row['Timestamp']}] "
                    f"USER_TERMINATION | {termination_row['TerminationReason']}"
                )
            except Exception as exc:
                print(f"[TraceHub] Could not write manual termination event: {exc}")

    except Exception as exc:
        print(f"[TraceHub] Startup/CSV error: {exc}")

    finally:
        try:
            csv_writer.close()
        except Exception:
            pass

        # Release UI Automation COM state on this worker thread.
        try:
            if uia_context is not None:
                uia_context.__del__()
        except Exception:
            pass

    print("TraceHub stopped.")
    print(f"CSV saved at: {ACTIVITY_CSV}")


def main() -> None:
    if sys.platform != "win32":
        print("TraceHub is designed for Windows only.")
        return

    mini = TraceHubMiniWindow()
    worker = threading.Thread(target=collector_loop, args=(mini,), daemon=True)
    worker.start()

    try:
        mini.run()
    except KeyboardInterrupt:
        mini.stop()


if __name__ == "__main__":
    main()
