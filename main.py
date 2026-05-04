import os
import json
import logging
import smtplib
import datetime
import pandas as pd
import requests
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from github import Github, GithubException

# ── CONFIG ────────────────────────────────────────────────────────────────────
GEMINI_API_KEY  = os.environ.get("GEMINI_API_KEY")
EMAIL_ADDRESS   = os.environ.get("EMAIL_ADDRESS")
EMAIL_PASSWORD  = os.environ.get("EMAIL_PASSWORD")
GITHUB_TOKEN    = os.environ.get("GH_TOKEN")
GITHUB_USERNAME = os.environ.get("GH_USERNAME")
GITHUB_REPO     = os.environ.get("GH_REPO")    # your games repo name
EXCEL_FILE      = "game_plan.csv"
SAVE_FOLDER     = r"D:\game upload automation details"
MAX_RETRIES     = 3
# ─────────────────────────────────────────────────────────────────────────────

# ── LOGGING ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)
log = logging.getLogger(__name__)
# ─────────────────────────────────────────────────────────────────────────────


# ── 1. LOCAL STORAGE ──────────────────────────────────────────────────────────
def setup_local_storage(date_str):
    """Create local folder structure if not exists."""
    base = SAVE_FOLDER
    paths = {
        "game":  os.path.join(base, "games",  date_str),
        "email": os.path.join(base, "emails"),
        "logs":  os.path.join(base, "logs"),
    }
    for p in paths.values():
        os.makedirs(p, exist_ok=True)
    log.info(f"Local storage ready at {base}")
    return paths


def save_local(paths, html, metadata, email_html, logs_text):
    """Save game files locally."""
    try:
        with open(os.path.join(paths["game"], "index.html"), "w", encoding="utf-8") as f:
            f.write(html)
        with open(os.path.join(paths["game"], "metadata.json"), "w", encoding="utf-8") as f:
            json.dump(metadata, f, indent=2)
        with open(os.path.join(paths["game"], "logs.txt"), "w", encoding="utf-8") as f:
            f.write(logs_text)
        with open(os.path.join(paths["email"], f"{metadata['date']}.html"), "w", encoding="utf-8") as f:
            f.write(email_html)
        sys_log = os.path.join(paths["logs"], "system.log")
        with open(sys_log, "a", encoding="utf-8") as f:
            f.write(f"{metadata['date']} | {metadata['game_name']} | {metadata['status']}\n")
        log.info("Files saved locally.")
    except Exception as e:
        log.warning(f"Local save skipped (may not be Windows): {e}")


# ── 2. EXCEL READER ───────────────────────────────────────────────────────────
def read_today_plan():
    """Read today's game plan from Excel."""
    if not os.path.exists(EXCEL_FILE):
        raise FileNotFoundError(f"Excel file not found: {EXCEL_FILE}")

    df = pd.read_excel(EXCEL_FILE)
    df.columns = ["Day", "Game Type", "Theme", "Difficulty"]

    today = datetime.date.today()
    # Match by day number (1-30) based on how many days since first row date
    # Use row index matching today's date order
    day_number = (datetime.date.today() - datetime.date.today().replace(day=1)).days + 1

    # Try to find row where Day == day_number
    row = df[df["Day"] == day_number]
    if row.empty:
        # fallback: cycle through rows
        idx = (datetime.date.today().timetuple().tm_yday - 1) % len(df)
        row = df.iloc[[idx]]

    plan = row.iloc[0].to_dict()
    log.info(f"Today's plan: {plan}")
    return plan


# ── 3. GAME GENERATOR ─────────────────────────────────────────────────────────
def generate_game(plan, attempt=1):
    """Use Gemini API to generate a complete HTML/CSS/JS game."""
    log.info(f"Generating game (attempt {attempt})...")

    prompt = f"""
You are an expert game developer. Create a complete, fully working, single-file web game.

Game Details:
- Game Type: {plan['Game Type']}
- Theme: {plan['Theme']}
- Difficulty: {plan['Difficulty']}
- Attempt: {attempt} (make it different from previous attempts if attempt > 1)

Requirements:
1. Single HTML file with embedded CSS and JavaScript
2. Must be fully playable in browser with no external dependencies
3. Must have: start screen, game loop, score display, game over screen
4. Must respond to keyboard/mouse/touch input
5. Clean, modern UI with the theme applied to colors and visuals
6. Include a game title that fits the theme
7. Make sure there are NO JavaScript errors
8. Game must have clear win/lose conditions
9. Add instructions on screen so player knows how to play

Return ONLY the complete HTML code, nothing else. Start with <!DOCTYPE html>
"""

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={GEMINI_API_KEY}"
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.8 + (attempt * 0.1), "maxOutputTokens": 8192}
    }

    response = requests.post(url, json=payload)
    data = response.json()

    if "candidates" not in data:
        raise Exception(f"Gemini error: {data}")

    html = data["candidates"][0]["content"]["parts"][0]["text"].strip()

    # Clean markdown code blocks if present
    if html.startswith("```"):
        html = html.split("```")[1]
        if html.startswith("html"):
            html = html[4:]
    if html.endswith("```"):
        html = html[:-3]

    html = html.strip()
    log.info("Game generated successfully.")
    return html


# ── 4. GAME NAME EXTRACTOR ────────────────────────────────────────────────────
def extract_game_name(html):
    """Extract game title from HTML."""
    import re
    match = re.search(r"<title>(.*?)</title>", html, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    match = re.search(r"<h1[^>]*>(.*?)</h1>", html, re.IGNORECASE | re.DOTALL)
    if match:
        return re.sub(r"<.*?>", "", match.group(1)).strip()
    return "Daily Game"


# ── 5. VALIDATOR ──────────────────────────────────────────────────────────────
def validate_game(html):
    """Basic validation of generated HTML game."""
    errors = []

    if "<!DOCTYPE html>" not in html.upper() and "<!doctype html>" not in html.lower():
        errors.append("Missing DOCTYPE")
    if "<html" not in html.lower():
        errors.append("Missing <html> tag")
    if "<body" not in html.lower():
        errors.append("Missing <body> tag")
    if "<script" not in html.lower():
        errors.append("Missing JavaScript")
    if len(html) < 2000:
        errors.append("HTML too short — game likely incomplete")

    # Check for common JS error patterns
    bad_patterns = ["undefined is not", "cannot read property", "syntaxerror"]
    for pattern in bad_patterns:
        if pattern in html.lower():
            errors.append(f"Possible JS error pattern found: {pattern}")

    if errors:
        log.warning(f"Validation failed: {errors}")
        return False, errors

    log.info("Validation passed!")
    return True, []


# ── 6. GITHUB DEPLOYMENT ──────────────────────────────────────────────────────
def deploy_to_github(html, date_str):
    """Push game to GitHub Pages repo."""
    log.info("Deploying to GitHub Pages...")

    g = Github(GITHUB_TOKEN)
    repo = g.get_user(GITHUB_USERNAME).get_repo(GITHUB_REPO)
    file_path = f"game-{date_str}/index.html"
    commit_msg = f"Add game for {date_str}"

    try:
        existing = repo.get_contents(file_path)
        repo.update_file(file_path, commit_msg, html, existing.sha)
        log.info(f"Updated existing file: {file_path}")
    except GithubException:
        repo.create_file(file_path, commit_msg, html)
        log.info(f"Created new file: {file_path}")

    url = f"https://{GITHUB_USERNAME}.github.io/{GITHUB_REPO}/game-{date_str}/"
    log.info(f"Game deployed at: {url}")
    return url


# ── 7. EMAIL SENDER ───────────────────────────────────────────────────────────
def build_email_html(plan, game_name, game_url, date_str, description):
    """Build HTML email body."""
    return f"""
<html><body style="font-family:Arial,sans-serif;max-width:600px;margin:auto;padding:20px">
  <h1 style="color:#6200ea">🎮 Daily Game Ready!</h1>
  <h2>{game_name}</h2>
  <table style="width:100%;border-collapse:collapse">
    <tr><td style="padding:8px;background:#f3e5f5"><b>📅 Date</b></td><td style="padding:8px">{date_str}</td></tr>
    <tr><td style="padding:8px;background:#f3e5f5"><b>🎯 Type</b></td><td style="padding:8px">{plan['Game Type']}</td></tr>
    <tr><td style="padding:8px;background:#f3e5f5"><b>🎨 Theme</b></td><td style="padding:8px">{plan['Theme']}</td></tr>
    <tr><td style="padding:8px;background:#f3e5f5"><b>⚡ Difficulty</b></td><td style="padding:8px">{plan['Difficulty']}</td></tr>
    <tr><td style="padding:8px;background:#f3e5f5"><b>🛠 Tech Stack</b></td><td style="padding:8px">HTML5, CSS3, JavaScript</td></tr>
  </table>
  <br>
  <p><b>📝 Description:</b> {description}</p>
  <br>
  <a href="{game_url}" style="background:#6200ea;color:white;padding:12px 24px;text-decoration:none;border-radius:6px;font-size:16px">
    ▶ Play Now
  </a>
  <br><br>
  <p style="color:#888;font-size:12px">Powered by Gemini AI + GitHub Pages | Game Automation System</p>
</body></html>
"""


def generate_description(plan, game_name):
    """Use Gemini to write a short game description."""
    prompt = f"Write a 2-sentence exciting game description for a {plan['Difficulty']} difficulty {plan['Game Type']} game called '{game_name}' with a {plan['Theme']} theme. Be enthusiastic and fun!"
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={GEMINI_API_KEY}"
    payload = {"contents": [{"parts": [{"text": prompt}]}]}
    response = requests.post(url, json=payload).json()
    try:
        return response["candidates"][0]["content"]["parts"][0]["text"].strip()
    except:
        return f"A {plan['Difficulty']} {plan['Game Type']} game with a {plan['Theme']} theme. Have fun!"


def send_email(plan, game_name, game_url, date_str, html, email_html):
    """Send game email with HTML attachment."""
    log.info("Sending email...")

    msg = MIMEMultipart("alternative")
    msg["From"]    = EMAIL_ADDRESS
    msg["To"]      = EMAIL_ADDRESS
    msg["Subject"] = f"🎮 Daily Game: {game_name} | Ready to Play!"

    plain = f"Today's game: {game_name}\nPlay here: {game_url}\nType: {plan['Game Type']} | Theme: {plan['Theme']} | Difficulty: {plan['Difficulty']}"
    msg.attach(MIMEText(plain, "plain"))
    msg.attach(MIMEText(email_html, "html"))

    # Attach HTML file
    attachment = MIMEBase("application", "octet-stream")
    attachment.set_payload(html.encode("utf-8"))
    encoders.encode_base64(attachment)
    attachment.add_header("Content-Disposition", f"attachment; filename=game-{date_str}.html")
    msg.attach(attachment)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.sendmail(EMAIL_ADDRESS, EMAIL_ADDRESS, msg.as_string())

    log.info("Email sent successfully!")


def send_failure_email(date_str, reason):
    """Send failure alert email."""
    log.info("Sending failure alert...")
    msg = MIMEMultipart()
    msg["From"]    = EMAIL_ADDRESS
    msg["To"]      = EMAIL_ADDRESS
    msg["Subject"] = f"❌ Game Automation Failed | {date_str}"
    body = f"Game generation failed on {date_str}.\n\nReason: {reason}\n\nPlease check your setup."
    msg.attach(MIMEText(body, "plain"))
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.sendmail(EMAIL_ADDRESS, EMAIL_ADDRESS, msg.as_string())
    except Exception as e:
        log.error(f"Failed to send failure email: {e}")


# ── 8. MAIN PIPELINE ──────────────────────────────────────────────────────────
def main():
    date_str = datetime.date.today().strftime("%Y-%m-%d")
    log.info(f"\n{'='*50}\n   GAME AUTOMATION PIPELINE | {date_str}\n{'='*50}")

    build_log = []

    # Step 1: Setup local storage
    paths = setup_local_storage(date_str)

    # Step 2: Read Excel plan
    try:
        plan = read_today_plan()
        build_log.append(f"Plan loaded: {plan}")
    except FileNotFoundError as e:
        log.error(str(e))
        send_failure_email(date_str, str(e))
        return

    # Step 3: Generate + Validate (max 3 attempts)
    html = None
    valid = False
    errors = []

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            html = generate_game(plan, attempt)
            valid, errors = validate_game(html)
            build_log.append(f"Attempt {attempt}: {'PASSED' if valid else 'FAILED - ' + str(errors)}")
            if valid:
                break
        except Exception as e:
            build_log.append(f"Attempt {attempt} error: {e}")
            log.error(f"Attempt {attempt} failed: {e}")

    if not valid or html is None:
        reason = f"Validation failed after {MAX_RETRIES} attempts. Last errors: {errors}"
        log.error(reason)
        send_failure_email(date_str, reason)
        return

    # Step 4: Extract game name & description
    game_name   = extract_game_name(html)
    description = generate_description(plan, game_name)
    build_log.append(f"Game name: {game_name}")

    # Step 5: Deploy to GitHub Pages
    try:
        game_url = deploy_to_github(html, date_str)
        build_log.append(f"Deployed: {game_url}")
    except Exception as e:
        log.error(f"Deployment failed: {e}")
        send_failure_email(date_str, f"GitHub deployment failed: {e}")
        return

    # Step 6: Build email
    email_html = build_email_html(plan, game_name, game_url, date_str, description)

    # Step 7: Save locally
    metadata = {
        "date":        date_str,
        "game_name":   game_name,
        "game_type":   plan["Game Type"],
        "theme":       plan["Theme"],
        "difficulty":  plan["Difficulty"],
        "url":         game_url,
        "description": description,
        "status":      "SUCCESS",
        "build_log":   build_log,
    }
    save_local(paths, html, metadata, email_html, "\n".join(build_log))

    # Step 8: Send email
    try:
        send_email(plan, game_name, game_url, date_str, html, email_html)
    except Exception as e:
        log.error(f"Email failed: {e}")

    log.info(f"\n{'='*50}\n   DONE! Game live at: {game_url}\n{'='*50}")


if __name__ == "__main__":
    main()
