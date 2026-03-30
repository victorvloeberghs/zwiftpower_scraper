import json
import os
import random
import time
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

load_dotenv()
from rich.console import Console
from rich.progress import (
    BarColumn,
    MofNCompleteColumn,
    Progress,
    TextColumn,
    TimeElapsedColumn,
    TimeRemainingColumn,
)

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------
INPUT_CSV       = "extracted_ids.csv"
OUTPUT_EXCEL    = "zwift_results.xlsx"
CHECKPOINT_FILE = "checkpoint.json"
USER_EMAIL      = os.environ["ZWIFT_EMAIL"]
USER_PASS       = os.environ["ZWIFT_PASS"]

RACE_NAMES = [
    "Stage 1: Zwift Games: Kaze Kicker",
    "Stage 2: Zwift Games: Hudson Hustle",
    "Stage 3: Zwift Games: Cobbled Crown",
    "Stage 4: Zwift Games: Peaky Pave",
    "Stage 5: Zwift Games: Three Step Sisters",
    "Stage 6a: Zwift Games: Epiloch",
]

NAV_DELAY   = (1.0, 2.0)   # seconds between individual page loads
BATCH_PAUSE = 15            # seconds every 15 riders
SHORT_PAUSE = 2             # seconds every 5 riders

# ---------------------------------------------------------------------------
console = Console()


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def time_to_excel(t_str: str):
    """Convert 'MM:SS' or 'HH:MM:SS' to an Excel serial time fraction."""
    if not t_str:
        return None
    if any(x in t_str.upper() for x in ("DNS", "DQ", "DNF", "ERR")):
        return None
    try:
        parts = list(map(int, t_str.split(":")))
        if len(parts) == 2:
            return (parts[0] * 60 + parts[1]) / 86400
        if len(parts) == 3:
            return (parts[0] * 3600 + parts[1] * 60 + parts[2]) / 86400
    except ValueError:
        pass
    return None


def col_letter(n: int) -> str:
    """Convert a 0-based column index to an Excel column letter (A, B, ... AA, ...)."""
    result = ""
    n += 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def load_checkpoint() -> dict:
    if Path(CHECKPOINT_FILE).exists():
        with open(CHECKPOINT_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_checkpoint(checkpoint: dict) -> None:
    with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
        json.dump(checkpoint, f, indent=2)


def navigate(page, url: str, retries: int = 2) -> bool:
    """Navigate to *url* with up to *retries* retry attempts on timeout."""
    for attempt in range(retries + 1):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=30_000)
            return True
        except PWTimeoutError:
            if attempt < retries:
                console.print(f"    [yellow]Timeout - retry {attempt + 1}/{retries}[/]")
                time.sleep(2)
    console.print(f"    [red]Failed to load:[/] {url}")
    return False


# ---------------------------------------------------------------------------
# SCRAPING
# ---------------------------------------------------------------------------

def login(page) -> bool:
    try:
        page.goto("https://zwiftpower.com/login.php", wait_until="networkidle", timeout=30_000)
        page.locator('a.btn-lg:has-text("Login with Zwift")').first.click()
        page.wait_for_selector("#username", timeout=45_000)
        page.fill("#username", USER_EMAIL)
        page.fill("#password", USER_PASS)
        page.click("button[type=submit]")
        try:
            page.wait_for_load_state("domcontentloaded", timeout=15_000)
        except PWTimeoutError:
            pass
        time.sleep(1)
        console.print("[green]Login successful.[/]")
        return True
    except Exception as exc:
        console.print(f"[red]Login failed:[/] {exc}")
        return False


def scrape_rider(page, rid: int) -> dict:
    profile_url = f"https://zwiftpower.com/profile.php?z={rid}"
    r_data: dict = {
        "rider_id":   rid,
        "rider_name": str(rid),
        "pace_group": "N/A",
        "profile":    profile_url,
    }

    try:
        if not navigate(page, profile_url):
            raise RuntimeError("Profile page unreachable")
        page.wait_for_selector("table#profile_results td.padright24", timeout=20_000)

        # Rider name
        name_el = page.locator('a[data-toggle="tab"]').first
        if name_el.count() > 0:
            r_data["rider_name"] = name_el.inner_text().strip()

        # Category / pace group
        pace_el = page.locator(
            "span.label-cat-A, span.label-cat-B, span.label-cat-C, span.label-cat-D"
        ).first
        if pace_el.count() > 0:
            val = pace_el.inner_text().strip()
            if not val.isdigit():
                r_data["pace_group"] = val

        # Collect all relevant race URLs in a single DOM pass
        links = page.locator("table#profile_results a.no_under")
        race_urls: dict = {rn: [] for rn in RACE_NAMES}
        for i in range(links.count()):
            link = links.nth(i)
            text = link.inner_text().strip()
            href = link.get_attribute("href") or ""
            url  = href if href.startswith("http") else "https://zwiftpower.com/" + href.lstrip("/")
            for rn in RACE_NAMES:
                if rn in text:
                    race_urls[rn].append(url)
                    break

        # Fetch each race result (best time across multiple attempts)
        for rn in RACE_NAMES:
            urls = race_urls[rn]
            if not urls:
                r_data[rn] = "DNS"
                continue

            if len(urls) > 1:
                console.print(f"    [dim]{rn}: {len(urls)} attempts - keeping best[/]")

            times_found: list = []
            for url in urls:
                time.sleep(random.uniform(*NAV_DELAY))
                try:
                    if not navigate(page, url):
                        continue
                    row_sel = f'tr:has(a[href*="z={rid}"])'
                    try:
                        page.wait_for_selector(row_sel, timeout=4_000)
                    except PWTimeoutError:
                        # Fallback: use the search box
                        try:
                            page.locator("input[type='search']").first.fill(str(rid))
                            page.wait_for_selector(row_sel, timeout=4_000)
                        except PWTimeoutError:
                            continue

                    time_cell = (
                        page.locator(row_sel).first
                        .locator("td.padright24 div.pull-left").first
                    )
                    if time_cell.count() > 0:
                        excel_t = time_to_excel(time_cell.inner_text().strip())
                        if excel_t is not None:
                            times_found.append(excel_t)
                except Exception as exc:
                    console.print(f"      [red]Attempt error:[/] {exc}")

            r_data[rn] = min(times_found) if times_found else "DNS"

    except Exception as exc:
        console.print(f"  [red]Error on rider {rid}:[/] {exc}")
        for rn in RACE_NAMES:
            r_data.setdefault(rn, "DNS")

    return r_data


# ---------------------------------------------------------------------------
# EXCEL OUTPUT
# ---------------------------------------------------------------------------

def save_excel(results: list) -> None:
    df   = pd.DataFrame(results)
    cols = ["rider_id", "rider_name", "pace_group", "profile"] + RACE_NAMES
    for col in cols:
        if col not in df.columns:
            df[col] = "DNS"
    for rn in RACE_NAMES:
        df[rn] = df[rn].where(df[rn].notna(), other="DNS")
    df = df[cols]

    with pd.ExcelWriter(OUTPUT_EXCEL, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="ZwiftResults")
        wb    = writer.book
        ws    = writer.sheets["ZwiftResults"]
        t_fmt = wb.add_format({"num_format": "hh:mm:ss"})

        ws.set_column("A:A", 12)
        ws.set_column("B:B", 25)
        ws.set_column("C:C", 12)
        ws.set_column("D:D", 40)

        # Dynamically compute the race column range (handles any number of stages)
        race_start = cols.index(RACE_NAMES[0])
        race_end   = race_start + len(RACE_NAMES) - 1
        col_range  = f"{col_letter(race_start)}:{col_letter(race_end)}"
        ws.set_column(col_range, 20, t_fmt)

    console.print(f"[green]Saved {len(results)} riders to {OUTPUT_EXCEL}[/]")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def run_scraper() -> None:
    df_ids    = pd.read_csv(INPUT_CSV)
    rider_ids = df_ids["ID"].tolist()

    # Resume from a previous run if a checkpoint exists
    checkpoint  = load_checkpoint()
    done_ids    = {int(k) for k in checkpoint}
    pending     = [rid for rid in rider_ids if rid not in done_ids]
    all_results = list(checkpoint.values())

    if done_ids:
        console.print(
            f"[cyan]Resuming:[/] {len(done_ids)} done, {len(pending)} remaining."
        )

    if not pending:
        console.print("[green]All riders already processed - generating Excel...[/]")
        save_excel(all_results)
        return

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge", headless=False)
        context = browser.new_context(viewport={"width": 1280, "height": 800})
        page    = context.new_page()

        if not login(page):
            browser.close()
            return

        with Progress(
            TextColumn("[bold blue]{task.description}"),
            BarColumn(),
            MofNCompleteColumn(),
            TimeElapsedColumn(),
            TimeRemainingColumn(),
            console=console,
        ) as progress:
            task = progress.add_task(
                "Scraping riders",
                total=len(rider_ids),
                completed=len(done_ids),
            )

            for idx, rid in enumerate(pending):
                progress.update(task, description=f"Rider {rid}")
                r_data = scrape_rider(page, rid)

                all_results.append(r_data)
                checkpoint[str(rid)] = r_data
                save_checkpoint(checkpoint)
                progress.advance(task)

                # Pacing: short pause every 5, longer pause every 15
                processed = len(done_ids) + idx + 1
                if processed % 15 == 0 and processed < len(rider_ids):
                    console.print(f"[dim]Pausing {BATCH_PAUSE}s...[/]")
                    time.sleep(BATCH_PAUSE)
                elif processed % 5 == 0:
                    time.sleep(SHORT_PAUSE)

        browser.close()

    save_excel(all_results)
    Path(CHECKPOINT_FILE).unlink(missing_ok=True)
    console.print("[green]Done! Checkpoint cleared.[/]")


if __name__ == "__main__":
    run_scraper()
