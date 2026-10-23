import os
import re
import io
import logging
from urllib.parse import urlparse

from flask import Flask, render_template, request, jsonify, send_file
import instaloader
from openpyxl import Workbook

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 2 * 1024 * 1024
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("vt_instagram")


def extract_shortcode(url):
    parsed = urlparse(url.strip())
    host = (parsed.hostname or "").lower()
    if host not in {"instagram.com", "www.instagram.com", "m.instagram.com"}:
        raise ValueError("لطفاً لینک معتبر instagram.com وارد کنید.")
    match = re.match(r"^/(p|reel|tv)/([A-Za-z0-9_-]+)/?", parsed.path)
    if not match:
        raise ValueError("لینک باید مربوط به یک پست، ریلز یا ویدئوی اینستاگرام باشد.")
    return match.group(2)


def get_loader():
    loader = instaloader.Instaloader(
        quiet=True,
        download_pictures=False,
        download_videos=False,
        download_video_thumbnails=False,
        download_geotags=False,
        download_comments=False,
        save_metadata=False,
        compress_json=False,
    )
    # Optional session file configured privately on the hosting service.
    session_file = os.environ.get("INSTAGRAM_SESSION_FILE", "").strip()
    username = os.environ.get("INSTAGRAM_USERNAME", "").strip()
    if session_file and username and os.path.exists(session_file):
        loader.load_session_from_file(username, session_file)
        logger.info("Instagram session loaded for configured account.")
    else:
        logger.info("Instagram request is running without a configured session file.")
    return loader


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/api/health")
def health():
    # Deliberately excludes secrets, usernames, and session contents.
    session_file = os.environ.get("INSTAGRAM_SESSION_FILE", "").strip()
    username = os.environ.get("INSTAGRAM_USERNAME", "").strip()
    configured = bool(session_file and username and os.path.exists(session_file))
    return jsonify({
        "ok": True,
        "service": "V.T Instagram",
        "instagram_session_configured": configured,
        "message": "این بررسی فقط وضعیت تنظیمات سرور را می‌سنجد؛ دسترسی اینستاگرام را تضمین نمی‌کند."
    })


@app.post("/api/comments")
def comments():
    data = request.get_json(silent=True) or {}
    url = (data.get("url") or "").strip()
    shortcode = "unknown"
    try:
        shortcode = extract_shortcode(url)
        loader = get_loader()
        post = instaloader.Post.from_shortcode(loader.context, shortcode)
        results = []
        max_comments = min(max(int(data.get("limit") or 500), 1), 5000)
        for i, comment in enumerate(post.get_comments()):
            if i >= max_comments:
                break
            results.append({
                "username": getattr(comment.owner, "username", "") or "",
                "text": getattr(comment, "text", "") or "",
                "date": comment.created_at_utc.strftime("%Y-%m-%d %H:%M:%S") if getattr(comment, "created_at_utc", None) else "",
                "likes": getattr(comment, "likes_count", 0),
            })
        logger.info("Fetched %s comments for shortcode %s.", len(results), shortcode)
        return jsonify({
            "ok": True,
            "shortcode": shortcode,
            "post_owner": getattr(post.owner, "username", "") or "",
            "count": len(results),
            "comments": results,
            "note": "تعداد دریافتی به دسترسی اینستاگرام و وضعیت پست بستگی دارد."
        })
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except Exception as e:
        # Full traceback stays in private hosting logs; never expose secrets to the browser.
        logger.exception("Instagram fetch failed for shortcode %s (%s)", shortcode, type(e).__name__)
        msg = str(e).lower()
        if "login" in msg or "401" in msg or "403" in msg or "checkpoint" in msg:
            friendly = "اینستاگرام دسترسی را محدود کرده یا ورود لازم است. کد خطا: " + type(e).__name__
        elif "not found" in msg or "404" in msg:
            friendly = "پست پیدا نشد یا عمومی/قابل‌دسترسی نیست. کد خطا: " + type(e).__name__
        elif "429" in msg or "too many requests" in msg or "rate limit" in msg:
            friendly = "اینستاگرام درخواست‌های سرور را موقتاً محدود کرده است. بعداً دوباره امتحان کنید. کد خطا: " + type(e).__name__
        else:
            friendly = "دریافت کامنت‌ها ناموفق بود. کد خطا: " + type(e).__name__
        return jsonify({"ok": False, "error": friendly, "error_code": type(e).__name__}), 502


@app.post("/api/followers")
def followers():
    # Instagram does not provide a supported way to enumerate followers of arbitrary
    # third-party profiles through this app. Do not pretend a scrape succeeded.
    data = request.get_json(silent=True) or {}
    profile = (data.get("url") or "").strip()
    if not profile:
        return jsonify({"ok": False, "error": "ابتدا لینک پروفایل یا نام کاربری را وارد کن."}), 400
    return jsonify({
        "ok": False,
        "error": "دریافت خودکار فالوورهای هر پروفایل در این نسخه پشتیبانی نمی‌شود؛ اینستاگرام دسترسی آن را محدود می‌کند. برای اینکه دکمه واقعاً کار کند، باید مسیر رسمی و مجاز برای حسابی که مدیریت می‌کنی پیاده‌سازی شود."
    }), 501


@app.post("/api/export.xlsx")
def export_xlsx():
    data = request.get_json(silent=True) or {}
    comments = data.get("comments") or []
    if not isinstance(comments, list) or not comments:
        return jsonify({"ok": False, "error": "برای خروجی گرفتن، ابتدا کامنت دریافت کنید."}), 400
    wb = Workbook()
    ws = wb.active
    ws.title = "V.T Comments"
    ws.append(["نام کاربری", "متن کامنت", "تاریخ", "تعداد لایک"])
    for item in comments:
        ws.append([
            str(item.get("username", "")),
            str(item.get("text", "")),
            str(item.get("date", "")),
            int(item.get("likes", 0) or 0),
        ])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col, width in {"A": 24, "B": 80, "C": 22, "D": 14}.items():
        ws.column_dimensions[col].width = width
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return send_file(
        bio,
        as_attachment=True,
        download_name="VT_Instagram_Comments.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8080")), debug=False)
