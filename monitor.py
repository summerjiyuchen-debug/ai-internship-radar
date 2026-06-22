from __future__ import annotations

import argparse
import base64
import csv
import datetime as dt
import email
import hashlib
import html
import imaplib
import json
import os
import re
import socket
import struct
import subprocess
import textwrap
import time
import urllib.request
from urllib.parse import parse_qs, unquote, urljoin, urlparse
from dataclasses import dataclass, field
from pathlib import Path
import tempfile
from typing import Iterable


USER_AGENT = "InternshipMonitor/1.0 (+personal student job search)"
BROWSER_CANDIDATES = [
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]


@dataclass
class Job:
    title: str
    company: str = ""
    location: str = ""
    url: str = ""
    source: str = ""
    source_type: str = ""
    posted_date: str = ""
    deadline: str = ""
    description: str = ""
    raw_id: str = ""
    score: int = 0
    level: str = ""
    reasons: list[str] = field(default_factory=list)
    risks: list[str] = field(default_factory=list)
    source_priority: int = 0
    source_error: bool = False

    @property
    def stable_id(self) -> str:
        base = self.raw_id or "|".join([self.title, self.company, self.location, self.url])
        return hashlib.sha256(base.encode("utf-8", errors="ignore")).hexdigest()[:16]


def load_json(path: Path, default):
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def get_env_var(name: str) -> str:
    value = os.environ.get(name, "")
    if value or os.name != "nt":
        return value
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as key:
            stored, _ = winreg.QueryValueEx(key, name)
            return str(stored)
    except Exception:
        return ""


def strip_tags(value: str) -> str:
    value = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", value)
    value = re.sub(r"(?s)<[^>]+>", " ", value)
    value = html.unescape(value)
    return re.sub(r"\s+", " ", value).strip()


def clean_job_title(value: str) -> str:
    value = re.sub(
        r"\b(featured|recently posted|strong applicant|urgently hiring|new|view job|browse|瀏覽)\b",
        " ",
        value,
        flags=re.I,
    )
    value = re.sub(r"\s+", " ", value).strip(" -|:·")
    return value[:140]


def split_linkedin_label(label: str) -> tuple[str, str, str]:
    label = clean_job_title(label)
    match = re.match(r"(?P<title>.*?)\s+(?P<company>[A-Z][^·]{2,80}?)\s*·\s*(?P<location>.+)$", label)
    if not match:
        return label, "", ""
    return (
        clean_job_title(match.group("title")),
        match.group("company").strip(),
        match.group("location").strip(),
    )


def infer_company_and_location_from_context(text: str, label: str, next_label: str = "") -> tuple[str, str]:
    label = strip_tags(label)
    index = text.find(label)
    if index < 0:
        return "", ""
    tail = text[index + len(label) : index + len(label) + 260]
    if next_label:
        next_index = tail.find(strip_tags(next_label))
        if next_index > 0:
            tail = tail[:next_index]
    tail = re.sub(
        r"\b(featured|recently posted|strong applicant|urgently hiring|view more jobs|apply now)\b",
        " ",
        tail,
        flags=re.I,
    )
    tail = re.sub(r"\s+", " ", tail).strip(" -|:·")
    location_pattern = (
        r"(Causeway Bay,?\s*Wan Chai District,?\s*HK|Tsim Sha Tsui,?\s*Yau Tsim Mong District,?\s*HK|"
        r"Mong Kok,?\s*Yau Tsim Mong District,?\s*HK|Central and Western District,?\s*HK|"
        r"Yau Tsim Mong District,?\s*HK|Wan Chai District,?\s*HK|Hong Kong Island,?\s*HK|"
        r"Sheung Wan|Causeway Bay|Wan Chai|Tsim Sha Tsui|Mong Kok|Kowloon|Shenzhen|Singapore|"
        r"\bHK\b|Remote|Hybrid)"
    )
    match = re.search(location_pattern, tail, re.I)
    if not match:
        return "", ""
    company = tail[: match.start()].strip(" -|:·,")
    location = match.group(1).strip(" -|:·,")
    after_location = tail[match.end() : match.end() + 30]
    remote_match = re.match(r"\s*(\((?:Remote|Hybrid)\))", after_location, re.I)
    if remote_match:
        location = f"{location} {remote_match.group(1)}"
    company_words = company.split()
    if len(company_words) > 8:
        company = " ".join(company_words[:8])
    return company[:100], location[:120]


def message_to_html_and_text(message: email.message.Message) -> tuple[str, str]:
    html_parts: list[str] = []
    text_parts: list[str] = []
    if message.is_multipart():
        parts = message.walk()
    else:
        parts = [message]
    for part in parts:
        content_type = part.get_content_type()
        if content_type not in {"text/html", "text/plain"}:
            continue
        payload = part.get_payload(decode=True)
        if not payload:
            continue
        charset = part.get_content_charset() or "utf-8"
        body = payload.decode(charset, errors="ignore")
        if content_type == "text/html":
            html_parts.append(body)
        else:
            text_parts.append(body)
    return "\n".join(html_parts), "\n".join(text_parts)


def extract_linkedin_email_jobs(page: str, source: dict, fallback_text: str = "") -> list[Job]:
    content = page or fallback_text
    text = strip_tags(content)
    links = re.findall(r'(?is)<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', page)
    jobs: list[Job] = []
    seen_urls: set[str] = set()
    for href, label_html in links:
        label = strip_tags(label_html)
        url = html.unescape(href)
        if "linkedin.com" not in url.lower():
            continue
        if not re.search(r"/jobs/|currentJobId|jobId", url, re.I):
            continue
        if url in seen_urls:
            continue
        seen_urls.add(url)
        title, company, location = split_linkedin_label(label or "LinkedIn job update")
        jobs.append(
            Job(
                title=title[:140],
                company=company,
                location=location,
                url=url,
                source=source.get("name", "LinkedIn email"),
                source_type=source.get("type", "linkedin_email"),
                description=text[:2500],
                raw_id=url,
                source_priority=int(source.get("source_priority_bonus", 0)),
            )
        )
    if jobs:
        return jobs
    if source.get("allow_whole_email_fallback") and re.search(r"\b(intern|internship|summer analyst|placement|trainee)\b", text, re.I):
        return [
            Job(
                title="LinkedIn email update",
                source=source.get("name", "LinkedIn email"),
                source_type=source.get("type", "linkedin_email"),
                description=text[:2500],
                raw_id=hashlib.sha256(text.encode("utf-8", errors="ignore")).hexdigest(),
            )
        ]
    return []


def extract_generic_email_jobs(page: str, source: dict, fallback_text: str = "") -> list[Job]:
    content = page or fallback_text
    text = strip_tags(content)
    allowed_domains = [domain.lower() for domain in source.get("job_domains", [])]
    allowed_tracking_domains = [domain.lower() for domain in source.get("tracking_domains", [])]
    job_keywords = source.get(
        "job_keywords",
        ["intern", "internship", "summer", "analyst", "assistant", "trainee", "placement"],
    )
    keyword_pattern = r"\b(" + "|".join(re.escape(word) for word in job_keywords) + r")\b"

    links = re.findall(r'(?is)<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', page)
    candidates: list[tuple[str, str]] = []
    for href, label_html in links:
        label = strip_tags(label_html)
        url = html.unescape(href)
        lowered_url = url.lower()
        allowed_by_domain = any(domain in lowered_url for domain in allowed_domains)
        allowed_by_tracking = any(domain in lowered_url for domain in allowed_tracking_domains)
        if (allowed_domains or allowed_tracking_domains) and not (allowed_by_domain or allowed_by_tracking):
            continue
        if not re.search(keyword_pattern, f"{label} {url}", re.I):
            continue
        if re.search(r"\b(view more jobs|recommendations|unsubscribe|privacy|terms|account)\b", label, re.I):
            continue
        candidates.append((href, label_html))
    jobs: list[Job] = []
    seen_urls: set[str] = set()
    for index, (href, label_html) in enumerate(candidates):
        label = strip_tags(label_html)
        url = html.unescape(href)
        if url in seen_urls:
            continue
        seen_urls.add(url)
        title = clean_job_title(label or source.get("name", "Job alert"))
        next_label = strip_tags(candidates[index + 1][1]) if index + 1 < len(candidates) else ""
        company, location = infer_company_and_location_from_context(text, label, next_label)
        jobs.append(
            Job(
                title=title,
                company=company,
                location=location,
                url=url,
                source=source.get("name", "Job alert email"),
                source_type=source.get("type", "job_alert_email"),
                description=text[:2500],
                raw_id="|".join([title.lower(), company.lower(), location.lower()]),
                source_priority=int(source.get("source_priority_bonus", 0)),
            )
        )

    if jobs:
        return jobs
    if source.get("allow_whole_email_fallback") and re.search(keyword_pattern, text, re.I):
        return [
            Job(
                title=source.get("fallback_title", f"{source.get('name', 'Job alert')} email update"),
                source=source.get("name", "Job alert email"),
                source_type=source.get("type", "job_alert_email"),
                description=text[:2500],
                raw_id=hashlib.sha256(text.encode("utf-8", errors="ignore")).hexdigest(),
            )
        ]
    return []


def extract_hku_cdt_email_jobs(page: str, source: dict, fallback_text: str = "") -> list[Job]:
    content = page or fallback_text
    text = strip_tags(content)
    links = re.findall(r'(?is)<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', page)
    jobs: list[Job] = []
    seen_urls: set[str] = set()

    for index, (href, label_html) in enumerate(links):
        label = clean_job_title(strip_tags(label_html))
        url = html.unescape(href)
        parsed = urlparse(url)
        target_url = parse_qs(parsed.query).get("url", [url])[0]
        target_url = unquote(target_url)
        lowered_target = target_url.lower()
        if "careers.fbe.hku.hk" not in lowered_target or "/careers/jobs/" not in lowered_target:
            continue
        if not label or re.search(r"\b(here|unsubscribe|view it in a web browser)\b", label, re.I):
            continue
        if target_url in seen_urls:
            continue
        seen_urls.add(target_url)

        next_label = clean_job_title(strip_tags(links[index + 1][1])) if index + 1 < len(links) else ""
        company = ""
        location = "HKU FBE Career Portal"
        text_index = text.find(label)
        if text_index >= 0:
            tail = text[text_index + len(label) : text_index + len(label) + 220]
            if next_label:
                next_index = tail.find(next_label)
                if next_index > 0:
                    tail = tail[:next_index]
            company = re.split(r"Application Deadline|Apply by|Deadline", tail, flags=re.I)[0]
            company = re.sub(r"\s+", " ", company).strip(" -|:·")
            if len(company.split()) > 8:
                company = " ".join(company.split()[:8])

        jobs.append(
            Job(
                title=label,
                company=company,
                location=location,
                url=url,
                source=source.get("name", "HKU CDT emails"),
                source_type=source.get("type", "hku_cdt_email"),
                description=f"{label} {company} {target_url}",
                raw_id="|".join([label.lower(), company.lower(), target_url.lower()]),
                source_priority=int(source.get("source_priority_bonus", 0)),
            )
        )
    return jobs


def extract_json_ld_jobs(page: str, source: dict) -> list[Job]:
    jobs: list[Job] = []
    blocks = re.findall(
        r'(?is)<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
        page,
    )
    for block in blocks:
        try:
            data = json.loads(html.unescape(block).strip())
        except json.JSONDecodeError:
            continue
        items = data if isinstance(data, list) else [data]
        for item in items:
            if isinstance(item, dict) and item.get("@graph"):
                items.extend(item["@graph"])
                continue
            if not isinstance(item, dict):
                continue
            if item.get("@type") != "JobPosting":
                continue
            org = item.get("hiringOrganization") or {}
            loc = item.get("jobLocation") or {}
            address = loc.get("address") if isinstance(loc, dict) else {}
            jobs.append(
                Job(
                    title=item.get("title", "").strip(),
                    company=(org.get("name") if isinstance(org, dict) else "") or "",
                    location=(
                        address.get("addressLocality", "")
                        if isinstance(address, dict)
                        else ""
                    ),
                    url=item.get("url") or source.get("url", ""),
                    source=source.get("name", ""),
                    source_type=source.get("type", ""),
                    posted_date=item.get("datePosted", ""),
                    deadline=item.get("validThrough", ""),
                    description=strip_tags(item.get("description", "")),
                    raw_id=item.get("identifier", {}).get("value", "")
                    if isinstance(item.get("identifier"), dict)
                    else "",
                )
            )
    return [job for job in jobs if job.title]


def extract_fallback_jobs(page: str, source: dict) -> list[Job]:
    title_match = re.search(r"(?is)<title[^>]*>(.*?)</title>", page)
    title = strip_tags(title_match.group(1)) if title_match else source.get("name", "Job page")
    text = strip_tags(page)
    if not re.search(r"\b(intern|internship|summer analyst|placement|trainee)\b", text, re.I):
        return []
    return [
        Job(
            title=title[:140],
            url=source.get("url", ""),
            source=source.get("name", ""),
            source_type=source.get("type", ""),
            description=text[:2500],
        )
    ]


def extract_anchor_jobs(page: str, source: dict) -> list[Job]:
    jobs: list[Job] = []
    seen: set[str] = set()
    source_url = source.get("url", "")
    links = re.findall(r'(?is)<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', page)
    for href, label_html in links:
        label = strip_tags(label_html)
        if not label or len(label) < 8:
            continue
        combined = f"{label} {href}"
        if not re.search(r"\b(intern|internship|summer|trainee|placement|analyst|assistant)\b", combined, re.I):
            continue
        url = urljoin(source_url, html.unescape(href))
        if url in seen:
            continue
        seen.add(url)
        jobs.append(
            Job(
                title=label[:140],
                url=url,
                source=source.get("name", ""),
                source_type=source.get("type", ""),
                description=label,
                raw_id=url,
            )
        )
    return jobs


def extract_text_jobs(page: str, source: dict) -> list[Job]:
    text = strip_tags(page)
    chunks = re.split(r"(?<=[.!?])\s+|\s{2,}", text)
    jobs: list[Job] = []
    seen: set[str] = set()
    for chunk in chunks:
        title = chunk.strip(" -|•\t\r\n")
        if not 10 <= len(title) <= 140:
            continue
        if not re.search(r"\b(intern|internship|summer|trainee|placement|analyst|assistant)\b", title, re.I):
            continue
        if re.search(r"\b(cookie|privacy|login|sign in|subscribe|newsletter|search jobs)\b", title, re.I):
            continue
        key = title.lower()
        if key in seen:
            continue
        seen.add(key)
        jobs.append(
            Job(
                title=title,
                url=source.get("url", ""),
                source=source.get("name", ""),
                source_type=source.get("type", ""),
                description=title,
                raw_id=f"{source.get('url', '')}|{title}",
            )
        )
        if len(jobs) >= 50:
            break
    return jobs


def fetch_url_with_curl(url: str) -> str:
    result = subprocess.run(
        [
            "curl.exe",
            "-L",
            "--silent",
            "--show-error",
            "--max-time",
            "25",
            "--ssl-no-revoke",
            "-A",
            USER_AGENT,
            url,
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="ignore",
        check=True,
    )
    return result.stdout


def find_browser_exes() -> list[str]:
    configured = os.environ.get("INTERNSHIP_BROWSER_EXE")
    if configured and Path(configured).exists():
        return [configured]
    found = []
    for candidate in BROWSER_CANDIDATES:
        if Path(candidate).exists():
            found.append(candidate)
    return found


def fetch_url_with_browser(source: dict) -> str:
    browser_exes = find_browser_exes()
    if not browser_exes:
        raise RuntimeError("Chrome or Edge was not found on this computer.")
    browser_exes = browser_exes[: int(source.get("browser_max_attempts", 1))]

    url = source["url"]
    wait_ms = int(source.get("browser_wait_ms", 12000))
    errors: list[str] = []
    headless_modes = ["--headless=new"]
    for browser_exe in browser_exes:
        for headless_mode in headless_modes:
            with tempfile.TemporaryDirectory(prefix="internship-browser-") as profile_dir:
                args = [
                    browser_exe,
                    headless_mode,
                    "--no-first-run",
                    "--no-default-browser-check",
                    "--disable-crash-reporter",
                    "--disable-background-networking",
                    "--disable-background-timer-throttling",
                    "--disable-renderer-backgrounding",
                    "--disable-sync",
                    "--disable-gpu",
                    "--disable-gpu-sandbox",
                    "--disable-software-rasterizer",
                    "--disable-dev-shm-usage",
                    "--disable-extensions",
                    "--disable-features=VizDisplayCompositor,UseSkiaRenderer,CalculateNativeWinOcclusion",
                    "--hide-scrollbars",
                    f"--user-data-dir={profile_dir}",
                    f"--virtual-time-budget={wait_ms}",
                    "--dump-dom",
                    url,
                ]
                try:
                    result = subprocess.run(
                        args,
                        capture_output=True,
                        text=True,
                        encoding="utf-8",
                        errors="ignore",
                        timeout=int(source.get("browser_timeout_seconds", 8)),
                    )
                except subprocess.TimeoutExpired as exc:
                    errors.append(f"{Path(browser_exe).name} {headless_mode} timed out")
                    continue
            if result.returncode != 0:
                stderr = textwrap.shorten(result.stderr.strip(), width=600, placeholder="...")
                errors.append(f"{Path(browser_exe).name} {headless_mode} exited {result.returncode}: {stderr}")
                continue
            if not result.stdout.strip():
                errors.append(f"{Path(browser_exe).name} {headless_mode} returned an empty page")
                continue
            return result.stdout
    raise RuntimeError(" / ".join(errors[-4:]))


def get_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def websocket_send(sock: socket.socket, payload: str) -> None:
    data = payload.encode("utf-8")
    key = os.urandom(4)
    header = bytearray([0x81])
    length = len(data)
    if length < 126:
        header.append(0x80 | length)
    elif length < 65536:
        header.append(0x80 | 126)
        header.extend(struct.pack("!H", length))
    else:
        header.append(0x80 | 127)
        header.extend(struct.pack("!Q", length))
    header.extend(key)
    masked = bytes(byte ^ key[index % 4] for index, byte in enumerate(data))
    sock.sendall(bytes(header) + masked)


def websocket_recv(sock: socket.socket) -> str:
    first = sock.recv(2)
    if len(first) < 2:
        raise RuntimeError("WebSocket closed early.")
    length = first[1] & 0x7F
    if length == 126:
        length = struct.unpack("!H", sock.recv(2))[0]
    elif length == 127:
        length = struct.unpack("!Q", sock.recv(8))[0]
    masked = bool(first[1] & 0x80)
    mask = sock.recv(4) if masked else b""
    payload = bytearray()
    while len(payload) < length:
        payload.extend(sock.recv(length - len(payload)))
    if masked:
        payload = bytearray(byte ^ mask[index % 4] for index, byte in enumerate(payload))
    return payload.decode("utf-8", errors="ignore")


def read_page_via_devtools(port: int, timeout_seconds: int) -> str:
    deadline = time.time() + timeout_seconds
    tabs = []
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/list", timeout=2) as response:
                tabs = json.loads(response.read().decode("utf-8", errors="ignore"))
            if tabs:
                break
        except Exception:
            time.sleep(0.5)
    if not tabs:
        raise RuntimeError("Browser DevTools endpoint did not become ready.")

    tab = next((item for item in tabs if item.get("type") == "page"), tabs[0])
    ws_url = tab["webSocketDebuggerUrl"]
    match = re.match(r"ws://([^:/]+):(\d+)(/.*)", ws_url)
    if not match:
        raise RuntimeError("Unexpected DevTools WebSocket URL.")
    host, raw_port, path = match.groups()

    key = base64.b64encode(os.urandom(16)).decode("ascii")
    with socket.create_connection((host, int(raw_port)), timeout=timeout_seconds) as sock:
        request = (
            f"GET {path} HTTP/1.1\r\n"
            f"Host: {host}:{raw_port}\r\n"
            "Upgrade: websocket\r\n"
            "Connection: Upgrade\r\n"
            f"Sec-WebSocket-Key: {key}\r\n"
            "Sec-WebSocket-Version: 13\r\n\r\n"
        )
        sock.sendall(request.encode("ascii"))
        handshake = sock.recv(4096)
        if b"101" not in handshake.split(b"\r\n", 1)[0]:
            raise RuntimeError("DevTools WebSocket handshake failed.")

        command = {
            "id": 1,
            "method": "Runtime.evaluate",
            "params": {
                "expression": "document.documentElement.outerHTML",
                "returnByValue": True,
                "awaitPromise": True,
            },
        }
        websocket_send(sock, json.dumps(command))
        while time.time() < deadline:
            message = json.loads(websocket_recv(sock))
            if message.get("id") == 1:
                return message.get("result", {}).get("result", {}).get("value", "")
    raise RuntimeError("Timed out reading page HTML from DevTools.")


def fetch_url_with_visible_browser(source: dict) -> str:
    browser_exes = find_browser_exes()
    if not browser_exes:
        raise RuntimeError("Chrome or Edge was not found on this computer.")
    browser_exes = sorted(browser_exes, key=lambda path: 0 if "Chrome" in path else 1)
    browser_exe = browser_exes[0]
    port = get_free_port()
    wait_seconds = int(source.get("visible_browser_wait_seconds", 12))
    timeout_seconds = int(source.get("visible_browser_timeout_seconds", 25))
    process = None
    profile_dir = tempfile.mkdtemp(prefix="internship-visible-browser-")
    try:
        args = [
            browser_exe,
            "--new-window",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-extensions",
            f"--remote-debugging-port={port}",
            f"--user-data-dir={profile_dir}",
            source["url"],
        ]
        process = subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        try:
            time.sleep(wait_seconds)
            page = read_page_via_devtools(port, timeout_seconds)
        finally:
            if process and process.poll() is None:
                process.terminate()
                try:
                    process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    process.kill()
        if not page.strip():
            raise RuntimeError("Visible browser returned an empty page.")
        return page
    finally:
        pass


def save_debug_page(base_dir: Path, source: dict, method: str, page: str) -> None:
    debug_dir = base_dir / "debug_pages"
    debug_dir.mkdir(exist_ok=True)
    safe_name = re.sub(r"[^A-Za-z0-9]+", "_", source.get("name", "source")).strip("_")
    path = debug_dir / f"{safe_name}_{method}.html"
    path.write_text(page, encoding="utf-8", errors="ignore")


def fetch_url(source: dict, base_dir: Path) -> list[Job]:
    errors: list[str] = []
    headers = {
        "User-Agent": USER_AGENT,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9,zh-HK;q=0.8",
    }
    if source.get("use_browser"):
        try:
            page = fetch_url_with_browser(source)
            save_debug_page(base_dir, source, "browser", page)
            jobs = extract_json_ld_jobs(page, source)
            parsed = jobs or extract_anchor_jobs(page, source) or extract_text_jobs(page, source) or extract_fallback_jobs(page, source)
            if parsed:
                return parsed
            errors.append("Browser loaded the page but no job-like entries were found.")
        except Exception as exc:
            errors.append(f"Browser fetch failed: {exc}")
        if source.get("use_visible_browser_fallback", True):
            try:
                page = fetch_url_with_visible_browser(source)
                save_debug_page(base_dir, source, "visible_browser", page)
                jobs = extract_json_ld_jobs(page, source)
                parsed = jobs or extract_anchor_jobs(page, source) or extract_text_jobs(page, source) or extract_fallback_jobs(page, source)
                if parsed:
                    return parsed
                errors.append("Visible browser loaded the page but no job-like entries were found.")
            except Exception as exc:
                errors.append(f"Visible browser fetch failed: {exc}")

    request = urllib.request.Request(source["url"], headers=headers)
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            page = response.read().decode("utf-8", errors="ignore")
            save_debug_page(base_dir, source, "urllib", page)
    except Exception:
        try:
            page = fetch_url_with_curl(source["url"])
            save_debug_page(base_dir, source, "curl", page)
        except Exception as exc:
            if errors:
                raise RuntimeError("; ".join(errors + [f"curl fetch failed: {exc}"])) from exc
            raise
    jobs = extract_json_ld_jobs(page, source)
    return jobs or extract_anchor_jobs(page, source) or extract_text_jobs(page, source) or extract_fallback_jobs(page, source)


def parse_local_html(source: dict, base_dir: Path) -> list[Job]:
    directory = base_dir / source["path"]
    jobs: list[Job] = []
    if not directory.exists():
        return jobs
    for path in directory.glob("*.html"):
        page = path.read_text(encoding="utf-8", errors="ignore")
        local_source = {**source, "url": str(path), "name": f"{source['name']} / {path.name}"}
        jobs.extend(extract_json_ld_jobs(page, local_source) or extract_fallback_jobs(page, local_source))
    return jobs


def parse_local_eml(source: dict, base_dir: Path) -> list[Job]:
    directory = base_dir / source["path"]
    jobs: list[Job] = []
    if not directory.exists():
        return jobs
    for path in directory.glob("*.eml"):
        raw = path.read_bytes()
        message = email.message_from_bytes(raw)
        page, text = message_to_html_and_text(message)
        local_source = {**source, "name": f"{source['name']} / {path.name}"}
        jobs.extend(extract_linkedin_email_jobs(page, local_source, text))
    return jobs


def fetch_imap_emails(source: dict) -> list[Job]:
    username = source.get("username") or get_env_var(source.get("username_env", "INTERNSHIP_EMAIL_USER"))
    password = get_env_var(source.get("password_env", "INTERNSHIP_EMAIL_PASSWORD"))
    if not username or not password:
        return [
            Job(
                title=f"Email source not configured: {source.get('name', 'IMAP')}",
                source=source.get("name", ""),
                source_type=source.get("type", "imap_email"),
                risks=["Missing email username/password environment variables."],
                source_error=True,
            )
        ]

    host = source["host"]
    mailbox = source.get("mailbox", "INBOX")
    sender = source.get("from", "")
    subject_contains = source.get("subject_contains", "")
    text_contains = source.get("text_contains", "")
    since_days = int(source.get("since_days", 7))
    since_date = (dt.date.today() - dt.timedelta(days=since_days)).strftime("%d-%b-%Y")
    query_parts = [f'SINCE "{since_date}"']
    if sender:
        query_parts.append(f'FROM "{sender}"')
    if subject_contains:
        query_parts.append(f'SUBJECT "{subject_contains}"')
    if text_contains:
        query_parts.append(f'TEXT "{text_contains}"')
    query = "(" + " ".join(query_parts) + ")"

    jobs: list[Job] = []
    with imaplib.IMAP4_SSL(host) as client:
        client.login(username, password)
        client.select(mailbox)
        status, data = client.search(None, query)
        if status != "OK":
            return jobs
        message_ids = data[0].split()[-int(source.get("max_messages", 20)) :]
        for message_id in message_ids:
            status, msg_data = client.fetch(message_id, "(RFC822)")
            if status != "OK":
                continue
            for item in msg_data:
                if not isinstance(item, tuple):
                    continue
                message = email.message_from_bytes(item[1])
                page, text = message_to_html_and_text(message)
                if source.get("type") == "linkedin_email":
                    jobs.extend(extract_linkedin_email_jobs(page, source, text))
                elif source.get("type") == "hku_cdt_email":
                    jobs.extend(extract_hku_cdt_email_jobs(page, source, text))
                else:
                    jobs.extend(extract_generic_email_jobs(page, source, text))
    return jobs


def parse_local_json(source: dict, base_dir: Path) -> list[Job]:
    directory = base_dir / source["path"]
    jobs: list[Job] = []
    if not directory.exists():
        return jobs
    for path in directory.glob("*.json"):
        data = load_json(path, [])
        if isinstance(data, dict):
            data = data.get("jobs", data.get("items", []))
        for item in data if isinstance(data, list) else []:
            if not isinstance(item, dict):
                continue
            jobs.append(
                Job(
                    title=str(item.get("title", "")).strip(),
                    company=str(item.get("company", item.get("companyName", ""))).strip(),
                    location=str(item.get("location", "")).strip(),
                    url=str(item.get("url", item.get("link", ""))).strip(),
                    source=source.get("name", ""),
                    source_type=source.get("type", ""),
                    posted_date=str(item.get("posted_date", item.get("datePosted", ""))).strip(),
                    deadline=str(item.get("deadline", item.get("validThrough", ""))).strip(),
                    description=str(item.get("description", item.get("summary", ""))).strip(),
                    raw_id=str(item.get("id", "")).strip(),
                )
            )
    return [job for job in jobs if job.title]


def collect_jobs(config: dict, base_dir: Path) -> list[Job]:
    sources = config.get("sources", {})
    jobs: list[Job] = []
    for source in sources.get("public_urls", []):
        if "example.com" in source.get("url", ""):
            continue
        try:
            jobs.extend(fetch_url(source, base_dir))
        except Exception as exc:
            jobs.append(
                Job(
                    title=f"Source unavailable: {source.get('name', 'unknown')}",
                    source=source.get("name", ""),
                    source_type=source.get("type", ""),
                    description=str(exc),
                    risks=["Source could not be fetched today."],
                    source_error=True,
                )
            )
    for source in sources.get("local_html_dirs", []):
        jobs.extend(parse_local_html(source, base_dir))
    for source in sources.get("local_eml_dirs", []):
        jobs.extend(parse_local_eml(source, base_dir))
    for source in sources.get("local_json_dirs", []):
        jobs.extend(parse_local_json(source, base_dir))
    for source in sources.get("imap_email", []):
        try:
            jobs.extend(fetch_imap_emails(source))
        except Exception as exc:
            jobs.append(
                Job(
                    title=f"Email source unavailable: {source.get('name', 'IMAP')}",
                    source=source.get("name", ""),
                    source_type=source.get("type", "imap_email"),
                    description=str(exc),
                    risks=["Email source could not be fetched today."],
                    source_error=True,
                )
            )
    return jobs


def tokenize(text: str) -> set[str]:
    return {
        token.lower()
        for token in re.findall(r"[A-Za-z][A-Za-z0-9+#.\-]{1,}", text)
        if len(token) >= 2
    }


def score_job(job: Job, cv_text: str, config: dict) -> Job:
    profile = config.get("student_profile", {})
    scoring = config.get("scoring", {})
    cv_tokens = tokenize(cv_text)
    focus_text = " ".join([job.title, job.company, job.location]).lower()
    job_text = focus_text
    job_tokens = tokenize(job_text)

    score = 20
    matched_cv_terms = sorted((cv_tokens & job_tokens), key=len, reverse=True)[:12]
    score += min(30, len(matched_cv_terms) * 3)

    role_hits = [role for role in profile.get("target_roles", []) if role.lower() in job_text]
    location_hits = [loc for loc in profile.get("target_locations", []) if loc.lower() in job_text]
    preferred_hits = [word for word in profile.get("preferred_keywords", []) if word.lower() in job_text]
    avoid_hits = [word for word in profile.get("avoid_keywords", []) if word.lower() in job_text]
    experience_hits = []
    for match in re.finditer(
        r"\b(?:minimum|at least|required|requires?)\s+(\d+)\+?\s+years?\b|\b(\d+)\+?\s+years?\s+of\s+experience\b",
        focus_text,
        re.I,
    ):
        for value in match.groups():
            if value:
                experience_hits.append(int(value))

    is_student_role = bool(
        re.search(r"\b(intern|internship|student|summer|placement|trainee|graduate trainee|programme|program)\b", focus_text, re.I)
    )

    if re.search(r"\b(intern|internship|summer analyst|placement|trainee)\b", focus_text):
        score += 20
        job.reasons.append("实习/学生项目关键词匹配")
    source_priority = int(job.source_priority or 0)
    if not source_priority and job.source_type == "hku_cdt_email":
        source_priority = int(scoring.get("hku_source_priority_bonus", 18))
    if source_priority:
        score += source_priority
        job.reasons.append(f"Source priority +{source_priority}: {job.source}")
    if preferred_hits:
        score += min(18, len(preferred_hits) * 6)
        job.reasons.append("Priority student/programme fit: " + ", ".join(preferred_hits))
    if role_hits:
        score += min(20, len(role_hits) * 8)
        job.reasons.append("目标岗位匹配：" + ", ".join(role_hits))
    if location_hits:
        score += 10
        job.reasons.append("目标地点匹配：" + ", ".join(location_hits))
    if matched_cv_terms:
        job.reasons.append("CV 关键词匹配：" + ", ".join(matched_cv_terms[:8]))
    if avoid_hits:
        score -= min(35, len(avoid_hits) * 12)
        job.risks.append("可能不适合：" + ", ".join(avoid_hits))
    if experience_hits:
        max_years = max(experience_hits)
        if max_years >= 2:
            score -= 45 if max_years >= 3 else 30
            job.risks.append(f"Experience requirement too high: {max_years}+ years")
    if not is_student_role and re.search(r"\b(analyst|research|manager|consultant|officer)\b", focus_text, re.I):
        score -= 25
        job.risks.append("Not clearly a student internship/programme")
    if re.search(r"\b(final year|2027 graduate|2028 graduate|penultimate)\b", focus_text):
        score += 8
        job.reasons.append("年级/毕业时间可能匹配")

    job.score = max(0, min(100, score))
    if job.score >= scoring.get("excellent_threshold", 75):
        job.level = "强烈推荐"
    elif job.score >= scoring.get("good_threshold", 55):
        job.level = "可申请"
    else:
        job.level = "低优先级"
    if not job.reasons:
        job.reasons.append("信息不足，需要手动查看岗位详情")
    return job


def is_blocked_job(job: Job, config: dict) -> bool:
    blocked_keywords = [word.lower() for word in config.get("student_profile", {}).get("blocked_keywords", [])]
    haystack = " ".join([job.title, job.company, job.location, job.description, job.url]).lower()
    return any(word and word in haystack for word in blocked_keywords)


def dedupe_new_jobs(jobs: Iterable[Job], seen_path: Path) -> list[Job]:
    seen = set(load_json(seen_path, []))
    unique: dict[str, Job] = {}
    for job in jobs:
        unique.setdefault(job.stable_id, job)
    new_jobs = [job for job_id, job in unique.items() if job_id not in seen]
    seen.update(unique.keys())
    seen_path.write_text(json.dumps(sorted(seen), indent=2), encoding="utf-8")
    return new_jobs


def unique_jobs(jobs: Iterable[Job]) -> list[Job]:
    unique: dict[str, Job] = {}
    for job in jobs:
        unique.setdefault(job.stable_id, job)
    return list(unique.values())


def write_markdown(jobs: list[Job], path: Path, source_errors: list[Job] | None = None) -> None:
    source_errors = source_errors or []
    today = dt.date.today().isoformat()
    lines = [f"# Internship Daily Report - {today}", ""]
    if not jobs:
        lines.append("今天没有发现新的岗位。")
    for index, job in enumerate(jobs, 1):
        lines.extend(
            [
                f"## {index}. {job.title}",
                f"- 推荐等级：{job.level}",
                f"- 匹配分数：{job.score}/100",
                f"- 公司：{job.company or 'N/A'}",
                f"- 地点：{job.location or 'N/A'}",
                f"- 来源：{job.source or job.source_type or 'N/A'}",
                f"- 截止日期：{job.deadline or 'N/A'}",
                f"- 链接：{job.url or 'N/A'}",
                f"- 推荐理由：{'; '.join(job.reasons)}",
                f"- 风险提示：{'; '.join(job.risks) if job.risks else '暂无明显风险'}",
                "",
            ]
        )
        if job.description:
            summary = textwrap.shorten(job.description, width=420, placeholder="...")
            lines.extend([f"> {summary}", ""])
    if source_errors:
        lines.extend(["## Source Issues", ""])
        for error in source_errors:
            message = textwrap.shorten(error.description or "; ".join(error.risks), width=360, placeholder="...")
            lines.extend(
                [
                    f"- {error.source or error.title}: {message or 'source unavailable'}",
                ]
            )
    path.write_text("\n".join(lines), encoding="utf-8")


def write_csv(jobs: list[Job], path: Path) -> None:
    fields = [
        "level",
        "score",
        "title",
        "company",
        "location",
        "source",
        "deadline",
        "url",
        "reasons",
        "risks",
    ]
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for job in jobs:
            writer.writerow(
                {
                    "level": job.level,
                    "score": job.score,
                    "title": job.title,
                    "company": job.company,
                    "location": job.location,
                    "source": job.source,
                    "deadline": job.deadline,
                    "url": job.url,
                    "reasons": "; ".join(job.reasons),
                    "risks": "; ".join(job.risks),
                }
            )


def write_html(jobs: list[Job], path: Path) -> None:
    today = dt.date.today().isoformat()

    def esc(value: str) -> str:
        return html.escape(str(value or ""), quote=True)

    rows: list[str] = []
    for index, job in enumerate(jobs, 1):
        link = (
            f'<a class="apply" href="{esc(job.url)}" target="_blank" rel="noopener">Open</a>'
            if job.url
            else '<span class="muted">N/A</span>'
        )
        reasons = esc("; ".join(job.reasons))
        rows.append(
            "<tr>"
            f"<td>{index}</td>"
            f"<td><strong>{esc(job.title)}</strong><div class=\"reasons\">{reasons}</div></td>"
            f"<td>{esc(job.company) or '<span class=\"muted\">N/A</span>'}</td>"
            f"<td>{esc(job.location) or '<span class=\"muted\">N/A</span>'}</td>"
            f"<td><span class=\"score\">{job.score}</span><div>{esc(job.level)}</div></td>"
            f"<td>{link}</td>"
            "</tr>"
        )
    if not rows:
        rows.append('<tr><td colspan="6" class="empty">No job recommendations found today.</td></tr>')

    document = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Internship Daily Report - {today}</title>
  <style>
    body {{
      margin: 0;
      font-family: "Segoe UI", Arial, sans-serif;
      background: #f6f7f9;
      color: #20242a;
    }}
    header {{
      padding: 28px 32px 18px;
      background: #ffffff;
      border-bottom: 1px solid #dde2e8;
    }}
    h1 {{
      margin: 0 0 8px;
      font-size: 24px;
      font-weight: 700;
    }}
    .summary {{
      color: #667085;
      font-size: 14px;
    }}
    main {{
      padding: 24px 32px 40px;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      background: #ffffff;
      border: 1px solid #dde2e8;
    }}
    th, td {{
      padding: 13px 14px;
      border-bottom: 1px solid #e8edf2;
      text-align: left;
      vertical-align: top;
      font-size: 14px;
    }}
    th {{
      background: #eef3f8;
      color: #344054;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: .04em;
    }}
    tr:hover td {{
      background: #fbfdff;
    }}
    .reasons {{
      max-width: 560px;
      margin-top: 6px;
      color: #667085;
      font-size: 12px;
      line-height: 1.45;
    }}
    .score {{
      display: inline-block;
      min-width: 36px;
      padding: 4px 8px;
      border-radius: 999px;
      background: #e9f7ef;
      color: #137547;
      font-weight: 700;
      text-align: center;
    }}
    .apply {{
      display: inline-block;
      padding: 7px 11px;
      border-radius: 6px;
      background: #175cd3;
      color: #ffffff;
      text-decoration: none;
      font-weight: 600;
      white-space: nowrap;
    }}
    .muted, .empty {{
      color: #98a2b3;
    }}
  </style>
</head>
<body>
  <header>
    <h1>Internship Daily Report - {today}</h1>
    <div class="summary">{len(jobs)} recommendations. Sorted by CV match score.</div>
  </header>
  <main>
    <table>
      <thead>
        <tr>
          <th>#</th>
          <th>Job</th>
          <th>Company</th>
          <th>Location</th>
          <th>Match</th>
          <th>Link</th>
        </tr>
      </thead>
      <tbody>
        {''.join(rows)}
      </tbody>
    </table>
  </main>
</body>
</html>
"""
    path.write_text(document, encoding="utf-8")


def write_ics(jobs: list[Job], path: Path) -> None:
    now = dt.datetime.now(dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    tomorrow = (dt.date.today() + dt.timedelta(days=1)).strftime("%Y%m%d")
    events = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Internship Monitor//EN"]
    for job in jobs:
        if job.level == "低优先级":
            continue
        uid = f"{job.stable_id}@internship-monitor"
        title = f"Apply: {job.title[:60]}"
        description = f"{job.company} {job.url}".replace("\n", " ")
        events.extend(
            [
                "BEGIN:VEVENT",
                f"UID:{uid}",
                f"DTSTAMP:{now}",
                f"DTSTART;VALUE=DATE:{tomorrow}",
                f"SUMMARY:{title}",
                f"DESCRIPTION:{description}",
                "END:VEVENT",
            ]
        )
    events.append("END:VCALENDAR")
    path.write_text("\n".join(events), encoding="utf-8")


def write_source_status(collected_jobs: list[Job], path: Path, config: dict) -> None:
    expected_sources = []
    sources = config.get("sources", {})
    for bucket in ("imap_email", "public_urls", "local_html_dirs", "local_eml_dirs", "local_json_dirs"):
        for source in sources.get(bucket, []):
            expected_sources.append(source.get("name", source.get("type", bucket)))

    lines = [f"Source Status - {dt.date.today().isoformat()}", ""]
    for source_name in expected_sources:
        source_jobs = [
            job for job in collected_jobs
            if (job.source == source_name or job.source.startswith(f"{source_name} /"))
            and not job.source_error
        ]
        source_errors = [
            job for job in collected_jobs
            if job.source == source_name and job.source_error
        ]
        lines.append(f"{source_name}")
        lines.append(f"  jobs_found: {len(source_jobs)}")
        if source_errors:
            for error in source_errors:
                message = error.description or "; ".join(error.risks) or "unknown error"
                lines.append(f"  issue: {message}")
        elif not source_jobs:
            lines.append("  issue: no job-like entries extracted")
        lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def cleanup_old_reports(report_dir: Path, keep_days: int = 3) -> None:
    dated_files: list[tuple[dt.date, Path]] = []
    patterns = (
        "daily_report_*.md",
        "daily_report_*.csv",
        "daily_report_*.html",
        "apply_reminders_*.ics",
        "source_status_*.txt",
    )
    for pattern in patterns:
        for path in report_dir.glob(pattern):
            match = re.search(r"(\d{4}-\d{2}-\d{2})", path.name)
            if not match:
                continue
            try:
                file_date = dt.date.fromisoformat(match.group(1))
            except ValueError:
                continue
            dated_files.append((file_date, path))

    keep_dates = set(sorted({file_date for file_date, _ in dated_files}, reverse=True)[:keep_days])
    for file_date, path in dated_files:
        if file_date not in keep_dates:
            try:
                path.unlink()
            except OSError:
                pass


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--include-seen", action="store_true")
    args = parser.parse_args()

    base_dir = Path(__file__).resolve().parent
    config_path = (base_dir / args.config).resolve()
    config = load_json(config_path, {})
    cv_text = (base_dir / "cv.txt").read_text(encoding="utf-8", errors="ignore")

    report_dir = base_dir / "reports"
    report_dir.mkdir(exist_ok=True)
    data_dir = base_dir / "data"
    data_dir.mkdir(exist_ok=True)

    collected_jobs = collect_jobs(config, base_dir)
    source_errors = [job for job in collected_jobs if job.source_error]
    jobs = [
        score_job(job, cv_text, config)
        for job in collected_jobs
        if job.title and not job.source_error and not is_blocked_job(job, config)
    ]
    jobs.sort(key=lambda job: job.score, reverse=True)
    jobs = unique_jobs(jobs)
    if not args.include_seen:
        jobs = dedupe_new_jobs(jobs, data_dir / "seen_jobs.json")
    jobs = jobs[: config.get("scoring", {}).get("max_jobs_in_report", 30)]

    today = dt.date.today().isoformat()
    write_markdown(jobs, report_dir / f"daily_report_{today}.md", [])
    write_csv(jobs, report_dir / f"daily_report_{today}.csv")
    write_html(jobs, report_dir / f"daily_report_{today}.html")
    write_ics(jobs, report_dir / f"apply_reminders_{today}.ics")
    write_source_status(collected_jobs, report_dir / f"source_status_{today}.txt", config)
    cleanup_old_reports(report_dir, keep_days=3)
    if args.include_seen:
        print(f"Generated {len(jobs)} job recommendations in {report_dir}")
    else:
        print(f"Generated {len(jobs)} new job recommendations in {report_dir}")
    if not jobs and source_errors:
        print(f"No reportable jobs were produced. See {report_dir / f'source_status_{today}.txt'} for diagnostics.")


if __name__ == "__main__":
    main()
