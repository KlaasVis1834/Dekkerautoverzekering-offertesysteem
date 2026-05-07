import os
import requests

ROLLS_BASE = "https://vergelijken.rolls.nl"
KENTEKEN = "HFK79S"

print("ROLLS_CAKEPHP env:", "SET" if os.getenv("ROLLS_CAKEPHP") else "MISSING")

ck = (os.getenv("ROLLS_CAKEPHP") or "").strip()
print("cookie len:", len(ck))

s = requests.Session()
s.headers.update({
    "accept": "*/*",
    "referer": f"{ROLLS_BASE}/",
    "user-agent": "Mozilla/5.0",
    "x-requested-with": "XMLHttpRequest",
})

# cookie zetten
if ck:
    s.cookies.set("CAKEPHP", ck, domain="vergelijken.rolls.nl", path="/")

url = f"{ROLLS_BASE}/beheer/data/kiwa/{KENTEKEN}"

try:
    r = s.get(url, timeout=20, allow_redirects=False)
    print("status:", r.status_code)
    print("location:", r.headers.get("location"))
    print("content-type:", r.headers.get("content-type"))
    print("set-cookie:", r.headers.get("set-cookie"))
    print("len:", len(r.text or ""))
    print("first 400 chars:\n", (r.text or "")[:400])
except Exception as e:
    print("EXCEPTION:", repr(e))
