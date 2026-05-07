import requests

CAKEPHP = "c1ccf89dee7b71b6ba0b6817669597a5"  # jouw waarde

s = requests.Session()
s.headers.update({
    "accept": "*/*",
    "accept-language": "nl,en;q=0.9,en-GB;q=0.8,en-US;q=0.7",
    "referer": "https://vergelijken.rolls.nl/",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36 Edg/144.0.0.0",
    "x-requested-with": "XMLHttpRequest",
})

# ✅ cookie correct zetten (niet via headers)
s.cookies.set("CAKEPHP", CAKEPHP, domain="vergelijken.rolls.nl", path="/")

url = "https://vergelijken.rolls.nl/beheer/data/kiwa/HFK79S"
r = s.get(url, timeout=20, allow_redirects=False)

print("status:", r.status_code)
print("location:", r.headers.get("location"))
print("content-type:", r.headers.get("content-type"))
print("set-cookie:", r.headers.get("set-cookie"))
print("first-300:", r.text[:300])

if r.headers.get("content-type", "").startswith("application/json"):
    data = r.json()
    print("OB_MLDCODE:", data.get("OB_MLDCODE"))
