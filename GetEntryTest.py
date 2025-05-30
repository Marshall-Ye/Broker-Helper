import json, textwrap
import requests

API   = "https://cert-api.acelynk.com"
USER  = "cert_acelynk_D9T"
PWD   = "Cg4fz9Sk0jNuL6"

ENTRY_NO   = "00000012"
INVOICE_NO = "369-93302311-A"

# --------------------------------------------------
def get_token() -> str:
    r = requests.post(f"{API}/token",
        data={"userName": USER, "password": PWD, "grant_type": "password"})
    r.raise_for_status()
    return r.json()["access_token"]

hdr = {"Authorization": f"Bearer {get_token()}"}

# -------- candidate routes to probe ----------
routes = [
    # entry-number look-ups
    f"/api/Invoices/GetEntrySummaryByEntryNumber/{ENTRY_NO}",
    f"/api/Invoices/GetEntrySummaryByEntryNo/{ENTRY_NO}",
    f"/api/Invoices/GetEntrySummary/{ENTRY_NO}",                      # some tenants use this
    f"/api/Invoices/GetEntrySummaryByEntryNumber?entryNumber={ENTRY_NO}",
    f"/api/Invoices/GetEntrySummaryByEntryNo?entryNo={ENTRY_NO}",

    # invoice-number look-ups
    f"/api/Invoices/GetEntrySummary/{INVOICE_NO}",
    f"/api/Invoices/GetEntrySummary?invoiceNumber={INVOICE_NO}",
]

print("🔎 Probing possible routes…\n")
for rel in routes:
    url = API + rel
    resp = requests.get(url, headers=hdr)
    print(f"{url}\n  → {resp.status_code} {resp.reason}")

    ctype = resp.headers.get("Content-Type","")
    if ctype.startswith("application/json"):
        try:
            print(json.dumps(resp.json(), indent=2)[:800])
        except Exception as e:
            print("   (JSON parse error)", e)
    else:
        print(textwrap.shorten(resp.text, width=300, placeholder=" …"))
    print("-" * 70)

# -------- optional: list routes published by /help --------
print("\n📜 /help endpoint (first 1200 bytes)…\n")
help_resp = requests.get(f"{API}/help")
print(textwrap.shorten(help_resp.text, width=1200, placeholder=" …"))
