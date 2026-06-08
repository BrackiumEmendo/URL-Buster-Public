import datetime
import random
import time
import re
from pathlib import Path
from typing import List, Optional
from collections import defaultdict
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

import pandas as pd
import streamlit as st
import requests
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font

# --------------------------------------------------------------------------- #
# CONFIG
# --------------------------------------------------------------------------- #
AUDIT_DIR = Path("audit")
AUDIT_DIR.mkdir(exist_ok=True)

URLLIST_FILE     = AUDIT_DIR / "urllist.csv"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/122.0 Safari/537.36"
)

HTTP_TIMEOUT     = 10
RETRIES          = 2
DOMAIN_DELAY     = 1.5
BATCH_SIZE       = 5000
CHECKPOINT_EVERY = 10

TRACKING_PARAMS = {
    'cid', 'src', 'its', 'l', 'f', 'step', 'site', 'fh', 'sign', 'add',
    'ppid', 'platform', 'see', 'locale', 'lang', 'type', 'sub', 'filter',
    'chang', 'i', 'daddr', 'target', 'show', 'area', 'caller', 'product',
    'device-type',
}

session = requests.Session()
session.headers.update({"User-Agent": USER_AGENT})

domain_last_hit = defaultdict(float)


# --------------------------------------------------------------------------- #
# CHECKPOINT
# --------------------------------------------------------------------------- #
def save_checkpoint(results: list, batch_idx: int):
    path = AUDIT_DIR / f"checkpoint_batch{batch_idx}.csv"
    pd.DataFrame(results).to_csv(path, index=False)


def load_checkpoint(batch_idx: int) -> list:
    path = AUDIT_DIR / f"checkpoint_batch{batch_idx}.csv"
    if path.exists():
        return pd.read_csv(path).fillna("").to_dict("records")
    return []


def clear_checkpoint(batch_idx: int):
    path = AUDIT_DIR / f"checkpoint_batch{batch_idx}.csv"
    if path.exists():
        path.unlink()


def clear_all_checkpoints():
    for f in AUDIT_DIR.glob("checkpoint_batch*.csv"):
        f.unlink()
    if URLLIST_FILE.exists():
        URLLIST_FILE.unlink()


def save_urllist(urls: List[str]):
    pd.DataFrame({"url": urls}).to_csv(URLLIST_FILE, index=False)


def load_urllist() -> List[str]:
    if URLLIST_FILE.exists():
        return pd.read_csv(URLLIST_FILE)["url"].dropna().tolist()
    return []


def any_checkpoint_exists() -> bool:
    return any(AUDIT_DIR.glob("checkpoint_batch*.csv"))


def get_saved_batch_indices() -> List[int]:
    indices = []
    for f in AUDIT_DIR.glob("checkpoint_batch*.csv"):
        try:
            idx = int(f.stem.replace("checkpoint_batch", ""))
            indices.append(idx)
        except ValueError:
            pass
    return sorted(indices)


# --------------------------------------------------------------------------- #
# URL CLEANER
# --------------------------------------------------------------------------- #
def clean_url(url: str) -> str:
    url = url.strip()
    if not url:
        return url
    return re.sub(
        r"/\d+(?:\.\d+)*/([a-z][a-z0-9-]*)(?:/\d+(?:\.\d+)*)?",
        r"/\1",
        url,
        flags=re.IGNORECASE,
    )


def strip_tracking_params(url: str) -> str:
    try:
        parsed = urlparse(url)
        if not parsed.query:
            return url
        params    = parse_qs(parsed.query, keep_blank_values=True)
        cleaned   = {k: v for k, v in params.items()
                     if k.lower() not in TRACKING_PARAMS}
        new_query = urlencode(cleaned, doseq=True)
        return urlunparse(parsed._replace(query=new_query))
    except Exception:
        return url


def get_platform_url(original: str, redirect: str) -> str:
    original = original.strip() if original else ""
    cleaned_original = clean_url(original)
    if cleaned_original != original:
        return cleaned_original
    return ""


def get_final_url(platform_redirect_url: str, platform_url: str,
                  redirect_url: str, original_url: str) -> str:
    for candidate in [platform_redirect_url, platform_url, redirect_url, original_url]:
        val = str(candidate).strip() if candidate else ""
        if val:
            return strip_tracking_params(clean_url(val))
    return ""


# --------------------------------------------------------------------------- #
# DEDUPLICATION
# --------------------------------------------------------------------------- #
def deduplicate_needs_scraping(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Scrape URL"] = df["Scrape URL"].astype(str).str.strip()
    df = df[df["Scrape URL"] != ""].copy()
    df = df[df["Scrape URL"].str.lower() != "nan"].copy()
    df = df.drop_duplicates(subset=["Scrape URL"], keep="first").reset_index(drop=True)

    seen_titles: set = set()
    keep: List[bool] = []

    for _, row in df.iterrows():
        title     = str(row.get("Title", "")).strip()
        title_key = title.lower()
        if title and title_key not in ("", "nan") and title in seen_titles:
            keep.append(False)
            continue
        if title and title_key not in ("", "nan"):
            seen_titles.add(title)
        keep.append(True)

    return df[keep].reset_index(drop=True)


# --------------------------------------------------------------------------- #
# EXCEL EXPORT
# --------------------------------------------------------------------------- #
def write_excel(df_all: pd.DataFrame, path: Path):
    wb = Workbook()
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    yellow_fill = PatternFill("solid", fgColor="FFFF00")
    red_fill    = PatternFill("solid", fgColor="FFC7CE")

    ws1 = wb.active
    ws1.title = "All URLs"
    for r in dataframe_to_rows(df_all, index=False, header=True):
        ws1.append(r)
    for cell in ws1[1]:
        cell.fill = header_fill
        cell.font = header_font

    url_col_idx       = next(
        (cell.column for cell in ws1[1] if cell.value == "URL"), None
    )
    final_url_col_idx = next(
        (cell.column for cell in ws1[1] if cell.value == "Final URL"), None
    )
    error_col_idx     = next(
        (cell.column for cell in ws1[1] if cell.value == "Error"), None
    )

    for row in ws1.iter_rows(min_row=2):
        error_val = (
            str(row[error_col_idx - 1].value).strip()
            if error_col_idx else ""
        )
        if error_val and error_val.lower() not in ("", "nan"):
            for cell in row:
                cell.fill = red_fill
        elif url_col_idx and final_url_col_idx:
            orig_val  = str(row[url_col_idx - 1].value).strip()
            final_val = str(row[final_url_col_idx - 1].value).strip()
            if orig_val != final_val and final_val:
                for cell in row:
                    cell.fill = yellow_fill

    for col in ws1.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws1.column_dimensions[col[0].column_letter].width = min(max_len + 4, 80)

    needs = df_all[df_all["Error"].astype(str).str.strip() == ""].copy()
    has_final = needs["Final URL"].astype(str).str.strip() != ""
    needs = needs[has_final].copy()
    needs["Scrape URL"] = needs["Final URL"]
    needs_out = needs[["Scrape URL", "Title"]].reset_index(drop=True)
    needs_out = deduplicate_needs_scraping(needs_out)

    ws2 = wb.create_sheet(title="Needs Scraping")
    for r in dataframe_to_rows(needs_out, index=False, header=True):
        ws2.append(r)
    for cell in ws2[1]:
        cell.fill = header_fill
        cell.font = header_font
    for col in ws2.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws2.column_dimensions[col[0].column_letter].width = min(max_len + 4, 80)

    wb.save(path)


# --------------------------------------------------------------------------- #
# HELPERS
# --------------------------------------------------------------------------- #
def now():
    return datetime.datetime.now().strftime("%Y%m%d_%H%M")


def wait_for_domain(url: str):
    domain   = url.split("/")[2]
    now_time = time.time()
    elapsed  = now_time - domain_last_hit[domain]
    if elapsed < DOMAIN_DELAY:
        time.sleep(DOMAIN_DELAY - elapsed)
    domain_last_hit[domain] = time.time()


# --------------------------------------------------------------------------- #
# CONTENT DETECTION
# --------------------------------------------------------------------------- #
def has_real_guide_content(html: str) -> bool:
    if not html:
        return False
    lower   = html.lower()
    signals = [
        "<h1", "role=\"main\"", "table of contents", "learn more",
        "controls", "parameters", "midi", "audio",
    ]
    return sum(1 for s in signals if s in lower) >= 3


def has_not_found_text(html: str) -> bool:
    if not html:
        return False
    lower   = html.lower()
    signals = [
        "the page you're looking for can't be found",
        "we can't find the page",
        "page-not-found",
    ]
    return any(s in lower for s in signals)


def is_probably_throttled(html: str, status: int) -> bool:
    if not html:
        return True
    if status == 503:
        return True
    if status == 200 and len(html) < 3000:
        return True
    lower = html.lower()
    if "akamai" in lower or "access denied" in lower:
        return True
    return False


def is_real_404(html: str) -> bool:
    if not html:
        return False
    lower = html.lower()
    if not has_not_found_text(html):
        return False
    if "table of contents" in lower:
        return False
    if "role=\"main\"" in lower:
        return False
    if "<h1" in lower and "not found" not in lower:
        return False
    return True


# --------------------------------------------------------------------------- #
# TITLE EXTRACTION
# --------------------------------------------------------------------------- #
def extract_title(html: str) -> str:
    if not html:
        return ""
    m     = re.search(r"<title[^>]*>(.*?)</title>", html, re.I | re.S)
    title = m.group(1).strip() if m else ""
    if not title or "page not found" in title.lower():
        og = re.search(r'<meta property="og:title" content="(.*?)"', html, re.I)
        if og:
            title = og.group(1).strip()
    if not title or "page not found" in title.lower():
        h1 = re.search(r"<h1[^>]*>(.*?)</h1>", html, re.I | re.S)
        if h1:
            clean = re.sub("<.*?>", "", h1.group(1)).strip()
            if clean:
                title = clean
    return title


def fix_title(url: str, title: str, html: str) -> str:
    if "/guide/" in url:
        if has_real_guide_content(html) and "page not found" in title.lower():
            h1 = re.search(r"<h1[^>]*>(.*?)</h1>", html, re.I | re.S)
            if h1:
                clean = re.sub("<.*?>", "", h1.group(1)).strip()
                if clean:
                    return clean
        if not has_real_guide_content(html) and has_not_found_text(html):
            return "Page Not Found"
    return title


# --------------------------------------------------------------------------- #
# CLASSIFIER
# --------------------------------------------------------------------------- #
def classify(url: str, html: str, status: int) -> Optional[str]:
    if is_probably_throttled(html, status):
        return None
    if has_real_guide_content(html):
        return None
    if is_real_404(html) and status == 404:
        return "404"
    if status == 404 and not has_not_found_text(html):
        return None
    if status == 404:
        return "404"
    if status >= 400:
        return str(status)
    return None


# --------------------------------------------------------------------------- #
# FETCH
# --------------------------------------------------------------------------- #
def fetch_url(url: str) -> dict:
    for attempt in range(RETRIES + 1):
        try:
            r        = session.get(url, timeout=HTTP_TIMEOUT, allow_redirects=True)
            html     = r.text
            title    = extract_title(html)
            title    = fix_title(url, title, html)
            redirect = r.url if r.url != url else ""
            if redirect:
                redirect = strip_tracking_params(redirect)
            error    = classify(url, html, r.status_code) or ""
            return {
                "url":          url,
                "redirect_url": redirect,
                "status":       r.status_code,
                "title":        title,
                "error":        error,
            }
        except Exception:
            if attempt == RETRIES:
                return {
                    "url":          url,
                    "redirect_url": "",
                    "status":       "ERROR",
                    "title":        "",
                    "error":        "ERROR",
                }
            time.sleep(random.uniform(1.0, 2.5) * (attempt + 1))


# --------------------------------------------------------------------------- #
# CRAWL
# --------------------------------------------------------------------------- #
def crawl(
    urls: List[str],
    batch_idx: int,
    existing_results: list = None,
    platform_check: bool = True,
) -> tuple:
    done_urls = set()
    results   = []

    if existing_results:
        results   = existing_results
        done_urls = {r["url"] for r in results}

    remaining  = [u for u in urls if u not in done_urls]
    total      = len(urls)
    completed  = len(results)
    start_time = time.time()

    progress_bar = st.session_state.get("progress")
    status_text  = st.session_state.get("status_text")

    # ── Pass 1 -- Fetch all URLs ─────────────────────────────────────────────
    for i, url in enumerate(remaining, 1):
        try:
            wait_for_domain(url)
            res = fetch_url(url)
            results.append(res)
            completed += 1

            if i % CHECKPOINT_EVERY == 0:
                save_checkpoint(results, batch_idx)

            elapsed = time.time() - start_time
            rate    = i / elapsed if elapsed else 0
            eta     = (len(remaining) - i) / rate if rate else 0

            if progress_bar:
                progress_bar.progress(
                    completed / (total * 2) if platform_check else completed / total
                )
            if status_text:
                status_text.text(
                    f"Pass 1 — Batch {batch_idx} — {completed:,}/{total:,} | "
                    f"{rate:.1f} URLs/s | ETA {int(eta)}s"
                )
        except Exception as exc:
            save_checkpoint(results, batch_idx)
            if status_text:
                status_text.warning(
                    f"Batch {batch_idx} interrupted at {completed:,} URLs — checkpoint saved."
                )
            raise exc

    save_checkpoint(results, batch_idx)

    # ── Pass 2 -- Platform URL check ─────────────────────────────────────────
    platform_check_results = {}

    if platform_check:
        if status_text:
            status_text.text(f"Pass 2 — Batch {batch_idx} — Checking Platform URLs...")

        pass1_map = {r["url"]: r for r in results}

        platform_urls_to_check = []
        for url in urls:
            r            = pass1_map.get(url, {})
            raw_redirect = r.get("redirect_url", "")
            purl         = get_platform_url(url, raw_redirect)
            if purl and purl != url:
                platform_urls_to_check.append(purl)

        unique_platform = list(dict.fromkeys(platform_urls_to_check))

        for j, purl in enumerate(unique_platform, 1):
            try:
                wait_for_domain(purl)
                res = fetch_url(purl)
                platform_check_results[purl] = res
            except Exception:
                platform_check_results[purl] = {
                    "url": purl, "redirect_url": "", "title": "", "error": "ERROR"
                }

            if progress_bar:
                progress_bar.progress(
                    min((total + j) / (total * 2), 1.0)
                )
            if status_text:
                status_text.text(
                    f"Pass 2 — Batch {batch_idx} — {j:,}/{len(unique_platform):,} platform URLs checked"
                )

    if progress_bar:
        progress_bar.progress(1.0)
    if status_text:
        if platform_check:
            status_text.text(
                f"Done — Batch {batch_idx} — {total:,} original + "
                f"{len(platform_check_results):,} platform URLs checked"
            )
        else:
            status_text.text(
                f"Done — Batch {batch_idx} — {total:,} URLs checked (simple mode)"
            )

    return results, platform_check_results


# --------------------------------------------------------------------------- #
# PROCESS ONE BATCH
# --------------------------------------------------------------------------- #
def process_batch(
    batch_urls: List[str],
    batch_idx: int,
    num_batches: int,
    timestamp: str,
    resume: bool,
    platform_check: bool = True,
    download_placeholder=None,
) -> Optional[Path]:

    st.markdown(f"---\n### Batch {batch_idx} of {num_batches} — {len(batch_urls):,} URLs")

    if not platform_check:
        st.caption("Simple mode — Platform URL cleanup and Pass 2 check skipped.")

    existing = load_checkpoint(batch_idx) if resume else []
    if existing:
        st.info(f"Resuming batch {batch_idx} from {len(existing):,} already crawled.")

    st.session_state.progress    = st.progress(0)
    st.session_state.status_text = st.empty()

    try:
        results, platform_check_results = crawl(
            batch_urls,
            batch_idx,
            existing_results=existing,
            platform_check=platform_check,
        )
    except Exception as e:
        st.error(
            f"Batch {batch_idx} stopped: `{e}`\n\n"
            "Progress saved — use **Resume** to continue."
        )
        return None

    clear_checkpoint(batch_idx)

    out = pd.DataFrame(results)
    for col in ["url", "redirect_url", "title", "error"]:
        if col not in out.columns:
            out[col] = ""
    out = out.fillna("")

    out["raw_redirect"] = out["redirect_url"]

    if platform_check:
        out["platform_url"] = out.apply(
            lambda row: get_platform_url(row["url"], row["raw_redirect"]), axis=1
        )

        def resolve_platform_redirect(row):
            purl = row["platform_url"]
            if not purl:
                return ""
            check = platform_check_results.get(purl)
            if not check:
                return ""
            actual_redirect = check.get("redirect_url", "")
            if actual_redirect and actual_redirect != purl:
                return actual_redirect
            return ""

        out["platform_redirect_url"] = out.apply(resolve_platform_redirect, axis=1)

        def resolve_title(row):
            purl = row["platform_url"]
            if purl and purl in platform_check_results:
                t = platform_check_results[purl].get("title", "")
                if t:
                    return t
            return row["title"]

        out["title"] = out.apply(resolve_title, axis=1)

        out["redirect_url"] = out.apply(
            lambda row: row["raw_redirect"]
            if (row["raw_redirect"] and not row["platform_url"])
            else "",
            axis=1,
        )
        out = out.drop(columns=["raw_redirect"])

        out["final_url"] = out.apply(
            lambda row: get_final_url(
                row["platform_redirect_url"],
                row["platform_url"],
                row["redirect_url"],
                row["url"],
            ),
            axis=1,
        )
    else:
        out["platform_url"]          = ""
        out["platform_redirect_url"] = ""
        out["redirect_url"]          = out["raw_redirect"]
        out = out.drop(columns=["raw_redirect"])
        out["final_url"] = out.apply(
            lambda row: row["redirect_url"] if row["redirect_url"] else row["url"],
            axis=1,
        )

    out = out.rename(columns={
        "url":                   "URL",
        "platform_url":          "Platform URL",
        "platform_redirect_url": "Platform Redirect URL",
        "redirect_url":          "Redirect URL",
        "final_url":             "Final URL",
        "title":                 "Title",
        "error":                 "Error",
    })

    out = out[[
        "URL", "Platform URL", "Platform Redirect URL",
        "Redirect URL", "Final URL", "Title", "Error"
    ]]

    original_order = pd.DataFrame({"URL": batch_urls})
    out = original_order.merge(out, on="URL", how="left").fillna("")

    out["Final URL"] = out["Final URL"].apply(
        lambda u: strip_tracking_params(str(u).strip()) if str(u).strip() else u
    )

    path = (
        AUDIT_DIR / f"run-{timestamp}.xlsx"
        if num_batches == 1
        else AUDIT_DIR / f"run-{timestamp}-batch{batch_idx}of{num_batches}.xlsx"
    )
    write_excel(out, path)

    csv_path = path.with_suffix(".csv")
    out.to_csv(csv_path, index=False)

    st.success(f"Batch {batch_idx} complete — `{path.name}`")

    if download_placeholder is not None:
        with open(path, "rb") as f:
            download_placeholder.download_button(
                label=f"Download Batch {batch_idx} — {path.name}",
                data=f,
                file_name=path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"dl_batch_{batch_idx}_{timestamp}",
            )

    return path


# --------------------------------------------------------------------------- #
# UI
# --------------------------------------------------------------------------- #
st.set_page_config(layout="wide")
st.title("URL Buster 5000")
st.caption("Upload a CSV of URLs to check for redirects, errors, and platform URL cleanup.")

# ── Upload CSV ────────────────────────────────────────────────────────────────
st.subheader("Upload URLs")
uploaded = st.file_uploader("Upload a CSV file -- first column should contain URLs")
if uploaded:
    df_up = pd.read_csv(uploaded)
    st.session_state.urls = df_up.iloc[:, 0].dropna().tolist()
    save_urllist(st.session_state.urls)
    st.success(f"Loaded {len(st.session_state.urls):,} URLs from CSV")

# ── Auto-restore URL list after page refresh ─────────────────────────────────
if "urls" not in st.session_state and URLLIST_FILE.exists():
    st.session_state.urls = load_urllist()

# ── Checkpoint resume banner ──────────────────────────────────────────────────
resume = False
if any_checkpoint_exists() and URLLIST_FILE.exists():
    all_urls      = load_urllist()
    saved_batches = get_saved_batch_indices()
    total_saved   = sum(len(load_checkpoint(i)) for i in saved_batches)
    num_batches   = (len(all_urls) + BATCH_SIZE - 1) // BATCH_SIZE

    st.warning(
        f"**Interrupted crawl detected** — "
        f"checkpoints found for batch(es): **{saved_batches}** "
        f"({total_saved:,} URLs saved across {len(saved_batches)} batch(es) "
        f"of {num_batches} total)."
    )
    if all_urls:
        progress_val = min(total_saved / len(all_urls), 1.0)
        st.progress(
            progress_val,
            text=f"{progress_val:.0%} complete before interruption",
        )

    col_r, col_f = st.columns(2)
    resume = col_r.button("Resume from checkpoint", use_container_width=True, type="primary")
    if col_f.button("Start fresh", use_container_width=True):
        clear_all_checkpoints()
        st.rerun()

# ── Run ───────────────────────────────────────────────────────────────────────
if "urls" in st.session_state:
    urls        = st.session_state.urls
    total_urls  = len(urls)
    num_batches = (total_urls + BATCH_SIZE - 1) // BATCH_SIZE

    st.divider()
    st.write(f"**{total_urls:,} URLs ready to crawl**")

    if num_batches > 1:
        st.info(
            f"{total_urls:,} URLs split into **{num_batches} batches** of up to "
            f"{BATCH_SIZE:,}. Each batch produces its own Excel file available for "
            f"download as soon as it finishes."
        )

    st.divider()
    mode = st.radio(
        "**Run mode**",
        options=["Full — Platform URL cleanup + Pass 2 check",
                 "Simple — Check URLs as-is, no cleanup"],
        index=0,
        horizontal=True,
        help=(
            "**Full mode:** strips version segments, derives Platform URLs, "
            "and runs a second-pass fetch to resolve where each platform URL lands.\n\n"
            "**Simple mode:** crawls the URL list exactly as provided — "
            "no version cleaning, no Platform URL derivation, no Pass 2."
        ),
    )
    platform_check = mode.startswith("Full")

    if not platform_check:
        st.info(
            "**Simple mode selected** — URLs will be checked as-is. "
            "Platform URL and Platform Redirect URL columns will be blank in the report.",
            icon="⚡",
        )

    st.divider()
    run_btn = st.button(
        "Ready, set, go & RUN!",
        type="primary",
        use_container_width=True,
    )

    if run_btn or resume:
        timestamp        = now()
        completed_paths: List[Path] = []

        url_batches = [
            urls[i : i + BATCH_SIZE]
            for i in range(0, total_urls, BATCH_SIZE)
        ]

        st.divider()
        st.subheader("Downloads")
        placeholders = {}
        for i in range(1, num_batches + 1):
            placeholders[i] = st.empty()

        st.divider()

        for batch_idx, batch_urls in enumerate(url_batches, start=1):
            if resume and not load_checkpoint(batch_idx) and batch_idx in get_saved_batch_indices():
                st.info(f"Batch {batch_idx} already complete — skipping.")
                continue

            path = process_batch(
                batch_urls           = batch_urls,
                batch_idx            = batch_idx,
                num_batches          = num_batches,
                timestamp            = timestamp,
                resume               = resume,
                platform_check       = platform_check,
                download_placeholder = placeholders[batch_idx],
            )
            if path:
                completed_paths.append(path)

        if completed_paths:
            st.session_state.completed_paths = completed_paths

        clear_all_checkpoints()
        st.balloons()
        st.success("All batches complete!")

# ── Persistent download section ───────────────────────────────────────────────
if "completed_paths" in st.session_state:
    st.divider()
    st.subheader("Downloads")
    for p in st.session_state.completed_paths:
        if Path(p).exists():
            with open(p, "rb") as f:
                st.download_button(
                    label=f"Download {Path(p).name}",
                    data=f,
                    file_name=Path(p).name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"dl_persist_{Path(p).name}",
                )
