"""Customer follow-up messages with WhatsApp click-to-chat links (any product).

Upload the monthly sales invoice item export. Pick a product preset (or Custom),
and the page finds customers who bought that product, optionally checks whether
they also bought a cross-sell product, writes a follow-up message for each, and
gives a WhatsApp link per customer. Staff click the link, WhatsApp opens with
the message typed, and they press Send.
"""
import io
import re
from datetime import datetime
from urllib.parse import quote

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Follow-Up Messages", page_icon="💬", layout="wide")

REQUIRED = ["Sales Date", "Client Name", "Client Mobile", "Item Code", "Item Name", "Sales Quantity"]

B12_TOKENS = "73051, 58193, 73110, 100196963, T73051, L73110, 100155670, L58193, 20485, 12590"

# ------------------------------------------------------------------ presets
# Message fields: {name} {outlet} {product} {item} {greeting} {promotion} {cross}
# WhatsApp-style wording: short paragraphs separated by blank lines.
# {benefit_line} is filled from the sidebar "What customers may feel" box and is left out when that box is empty.
PROMO_EN = (
    "Hi {name}, {greeting}! \U0001F44B\n"
    "This is {outlet}.\n"
    "\n"
    "How have you been finding {item}? Hope it is treating you well \U0001F60A\n"
    "\n"
    "{benefit_line}We have a promotion on it at the moment:\n"
    "\U0001F3F7\uFE0F {promotion}\n"
    "\n"
    "Would you like to top up? Just reply here and we will get it ready for you \U0001F6CD\uFE0F\n"
    "\n"
    "Thank you!"
)
PROMO_MS = (
    "Hai {name}, {greeting}! \U0001F44B\n"
    "Ini {outlet}.\n"
    "\n"
    "Bagaimana pengalaman anda menggunakan {item}? Harap anda selesa menggunakannya \U0001F60A\n"
    "\n"
    "{benefit_line}Kami ada promosi untuk produk ini sekarang:\n"
    "\U0001F3F7\uFE0F {promotion}\n"
    "\n"
    "Anda mahu top up? Balas mesej ini dan kami sediakan untuk anda \U0001F6CD\uFE0F\n"
    "\n"
    "Terima kasih!"
)

PROMO_SW = (
    "Hai {name}, {greeting}! \U0001F44B\n"
    "Kamek dari {outlet}.\n"
    "\n"
    "Kemadah pengalaman kitak ngan {item}? Harap kitak selesa guna nya \U0001F60A\n"
    "\n"
    "{benefit_line}Kami ada promosi untuk produk tok sekarang:\n"
    "\U0001F3F7\uFE0F {promotion}\n"
    "\n"
    "Kitak nak top up? Balas mesej tok, kami sediakan untuk kitak \U0001F6CD\uFE0F\n"
    "\n"
    "Terima kasih!"
)

PRESETS = {
    "Promotion follow-up (any product)": {
        "product": "",
        "keywords": "",
        "exclude": "",
        "cross_on": False,
        "cross": "",
        "cross_tokens": "",
        "templates": {
            "English": {"no_cross": PROMO_EN, "has_cross": ""},
            "Bahasa Malaysia": {"no_cross": PROMO_MS, "has_cross": ""},
            "Bahasa Sarawak": {"no_cross": PROMO_SW, "has_cross": ""},
        },
    },
    "Metformin users (Vitamin B12 cross-sell)": {
        "product": "metformin",
        "keywords": "GLUMET, GLUCOPHAGE, GLUCOXIT, JANUMET, XIGDUO, GALVUS MET, GLUCOVANCE, JARDIANCE DUO, METFORMIN",
        "exclude": "",
        "cross_on": True,
        "cross": "Vitamin B12",
        "cross_tokens": B12_TOKENS,
        "templates": {
            "English": {
                "no_cross": (
                    "Hi {name}, this is {outlet}. How are you doing? "
                    "Your {product} may be due for a refill. Would you like us to prepare it for you? "
                    "Long-term {product} use can lower vitamin B12 levels, so please let us know if you "
                    "feel any numbness or tingling in your hands or feet. Our pharmacist can advise you, "
                    "and you may consider a vitamin B12 supplement. Thank you!"
                ),
                "has_cross": (
                    "Hi {name}, this is {outlet}. How are you doing? "
                    "Your {product} may be due for a refill. Would you like us to prepare it for you? "
                    "Please let us know if you feel any numbness or tingling in your hands or feet. "
                    "Our pharmacist is happy to advise you. Thank you!"
                ),
            },
            "Bahasa Malaysia": {
                "no_cross": (
                    "Salam {name}, ini {outlet}. Apa khabar? "
                    "Ubat {product} anda mungkin hampir habis. Adakah anda mahu kami sediakannya? "
                    "Penggunaan {product} jangka panjang boleh merendahkan paras vitamin B12. "
                    "Sila beritahu kami jika anda rasa kebas atau kesemutan pada tangan atau kaki. "
                    "Ahli farmasi kami boleh memberi nasihat, dan anda boleh mempertimbangkan "
                    "suplemen vitamin B12. Terima kasih!"
                ),
                "has_cross": (
                    "Salam {name}, ini {outlet}. Apa khabar? "
                    "Ubat {product} anda mungkin hampir habis. Adakah anda mahu kami sediakannya? "
                    "Sila beritahu kami jika anda rasa kebas atau kesemutan pada tangan atau kaki. "
                    "Ahli farmasi kami sedia membantu. Terima kasih!"
                ),
            },
        },
    },
    "Omega-3 / fish oil users": {
        "product": "omega-3",
        "keywords": "OMEGA, FISH OIL",
        "exclude": "GUMMYKIDS",
        "cross_on": False,
        "cross": "",
        "cross_tokens": "",
        "templates": {
            "English": {
                "no_cross": (
                    "Hi {name}, this is {outlet}. How are you doing? "
                    "Your {product} supplement may be due for a refill. Would you like us to prepare it for you? "
                    "Let us know if you have any questions about how to take it. Thank you!"
                ),
                "has_cross": "",
            },
            "Bahasa Malaysia": {
                "no_cross": (
                    "Salam {name}, ini {outlet}. Apa khabar? "
                    "Suplemen {product} anda mungkin hampir habis. Adakah anda mahu kami sediakannya? "
                    "Sila hubungi kami jika anda ada soalan tentang cara pengambilannya. Terima kasih!"
                ),
                "has_cross": "",
            },
        },
    },
    "Custom (set your own)": {
        "product": "",
        "keywords": "",
        "exclude": "",
        "cross_on": False,
        "cross": "",
        "cross_tokens": "",
        "templates": {
            "English": {"no_cross": PROMO_EN, "has_cross": ""},
            "Bahasa Malaysia": {"no_cross": PROMO_MS, "has_cross": ""},
            "Bahasa Sarawak": {"no_cross": PROMO_SW, "has_cross": ""},
        },
    },
}


# ---------------------------------------------------------------- helpers
def split_list(text: str) -> list[str]:
    return [t.strip().upper() for t in re.split(r"[,\n]", text) if t.strip()]


def clean_name(raw) -> str:
    """Remove prefixes like (DR) or (QKSAM1) and tidy capitalisation."""
    s = re.sub(r"^\s*(\([^)]*\)\s*)+", "", str(raw)).strip()
    return s.title() if s else "Customer"


def normalise_mobile(raw) -> str | None:
    """Return Malaysian number as 60XXXXXXXXX, or None if it looks invalid."""
    if pd.isna(raw):
        return None
    digits = re.sub(r"\D", "", str(raw))
    if digits.startswith("0"):
        digits = "6" + digits
    elif digits.startswith("1"):
        digits = "60" + digits
    if digits.startswith("60") and 11 <= len(digits) <= 12:
        return digits
    return None


def build_followups(df, keywords, exclude, cross_tokens, promo=None, use_main_pack=False):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    desc = df["Description"].fillna("") if "Description" in df.columns else pd.Series("", index=df.index)
    text_up = df["Item Name"].fillna("").str.upper() + " " + desc.astype(str).str.upper()
    code_up = df["Item Code"].astype(str).str.strip().str.upper()

    df["is_prod"] = text_up.apply(
        lambda t: any(k in t for k in keywords) and not any(x in t for x in exclude)
    )
    if cross_tokens:
        df["is_cross"] = code_up.isin(cross_tokens) | text_up.apply(
            lambda t: any(k in t for k in cross_tokens if not k.isdigit())
        )
    else:
        df["is_cross"] = False
    df["Sales Date"] = pd.to_datetime(df["Sales Date"], errors="coerce")
    df["mobile"] = df["Client Mobile"].apply(normalise_mobile)

    prod = df[df["is_prod"]]
    skipped = prod[prod["mobile"].isna()][["Client Name", "Client Mobile", "Item Name", "Sales Date"]]
    skipped = skipped.drop_duplicates()
    prod = prod[prod["mobile"].notna()]

    # one row per customer; returns (negative qty) are netted off
    net = prod.groupby("mobile")["Sales Quantity"].sum()
    prod = prod[prod["mobile"].isin(net[net > 0].index)].copy()
    promo = promo or {}
    codes = prod["Item Code"].astype(str).str.strip().str.upper()
    prod["promo_rule"] = codes.map(promo).fillna("")
    if use_main_pack:
        # single-pack codes such as T73051 or L73110 fall back to the main item's rule (73051, 73110)
        base = codes.str.replace(r"^[A-Z](?=\d)", "", regex=True)
        prod.loc[prod["promo_rule"] == "", "promo_rule"] = base.map(promo).fillna("")[prod["promo_rule"] == ""]
    prod["has_promo"] = prod["promo_rule"] != ""
    # per customer: prefer the latest purchased item that has a promotion rule, else the latest item
    prod = prod.sort_values(["has_promo", "Sales Date"], kind="stable")
    last = prod.groupby("mobile").tail(1).set_index("mobile")

    cross_mobiles = set(df[df["is_cross"] & df["mobile"].notna()]["mobile"])
    out = pd.DataFrame({
        "name": last["Client Name"].apply(clean_name),
        "last_date": last["Sales Date"],
        "last_item": last["Item Name"],
        "item": last["Item Name"].apply(clean_item),
        "promo_rule": last["promo_rule"],
        "has_cross": [m in cross_mobiles for m in last.index],
    })
    out.index.name = "mobile"
    out = out.reset_index()
    out["days_since"] = (pd.Timestamp(datetime.now().date()) - out["last_date"]).dt.days
    return out, skipped


def clean_item(raw) -> str:
    """Tidy an item name for messages: drop markers like *[RX]* or (NEW), soften capitals."""
    s = re.sub(r"^(?:\s*(?:\*|#|\[[^\]]*\]|\([^)]*\)))+\s*", "", str(raw)).strip()
    words = [w if (len(w) <= 3 or any(c.isdigit() for c in w)) else w.capitalize() for w in s.split()]
    return " ".join(words) or str(raw)


_MONTHS = r"(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[A-Z]{0,6}"
_DATE_PART = re.compile(
    r"\b(?:EFFECTIVE|EFF|FROM|TILL|UNTIL|VALID|W\.?E\.?F)\b"
    r"|\d{1,2}\s*[./-]\s*\d{1,2}\s*[./-]\s*\d{2,4}"
    rf"|\b\d{{1,2}}\s?{_MONTHS}\b(?:\s?\d{{2,4}})?"
    rf"|\b{_MONTHS}\.?\s?\d{{2,4}}\b",
    re.IGNORECASE,
)
_SMALL_WORDS = {"at", "for", "each", "can", "mix", "with", "only", "of", "and", "unit", "units", "per"}


def _strip_dates(text: str) -> str:
    """Remove date / effective-from wording; keep other bracket notes like (can mix with X)."""
    def fix_group(m):
        parts = re.split(r"\s+[-\u2013]\s+|;", m.group(1))
        keep = [x.strip() for x in parts if x.strip() and not _DATE_PART.search(x)]
        return f"({' - '.join(keep)})" if keep else ""

    text = re.sub(r"\(([^()]*)\)", fix_group, text)
    text = re.sub(r"\b(?:EFFECTIVE|EFF)\b.*$", "", text, flags=re.IGNORECASE)  # trailing, no brackets
    text = re.sub(r"\s*\*+#?\s*", " / ", text)  # stray **# separators in the file
    return re.sub(r"\s+", " ", text).strip(" -,;/")


def _tidy_case(text: str) -> str:
    def fix(m):
        w = m.group(0)
        if any(c.isdigit() for c in w):
            return w.upper()
        return w.lower() if w.lower() in _SMALL_WORDS else w.capitalize()

    text = re.sub(r"\s*@\s*", " @ ", text)
    text = re.sub(r"(RM\d+)\.00\b", r"\1", text, flags=re.IGNORECASE)
    text = re.sub(r"[A-Za-z][A-Za-z0-9]*", fix, text)
    return re.sub(r"\s+", " ", text).strip()


def _promo_end_date(text: str):
    m = re.search(r"\b(?:TILL|UNTIL)\s*(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})", text, flags=re.IGNORECASE)
    if not m:
        return None
    d, mo, y = (int(x) for x in m.groups())
    y += 2000 if y < 100 else 0
    try:
        return datetime(y, mo, d).date()
    except ValueError:
        return None


def clean_promo(raw) -> str:
    """Turn an Attribute value into customer-friendly wording without dates."""
    s = "" if pd.isna(raw) else str(raw).strip()
    if s in ("", "0", "-"):
        return ""
    return _tidy_case(_strip_dates(s))


def load_promo_rules(file):
    """Return (rules, review). rules: Item Code -> promotion text. review: table of original vs message wording."""
    df = pd.read_excel(file, dtype={"Item Code": str})
    df.columns = [str(c).strip() for c in df.columns]
    if "Item Code" not in df.columns or "Attribute" not in df.columns:
        raise ValueError("The promotion file needs the columns 'Item Code' and 'Attribute'.")
    df = df[df["Attribute"].notna()].copy()
    df["Item Code"] = df["Item Code"].astype(str).str.strip().str.upper()
    df["Original"] = df["Attribute"].astype(str).str.strip()
    df["In message"] = df["Original"].apply(clean_promo)
    today = datetime.now().date()
    end = df["Original"].apply(_promo_end_date)
    df["Status"] = ["Ended " + e.strftime("%d %b %Y") if (e and e < today) else "OK" for e in end]
    df = df[df["In message"] != ""]
    live = df[df["Status"] == "OK"]
    review = df[["Item Code"] + (["Name"] if "Name" in df.columns else []) + ["Original", "In message", "Status"]]
    return dict(zip(live["Item Code"], live["In message"])), review.reset_index(drop=True)


def make_greeting(language: str, mode: str, custom: str) -> str:
    if mode == "Custom text" and custom.strip():
        return custom.strip()
    h = datetime.now().hour
    if language in ("Bahasa Malaysia", "Bahasa Sarawak"):
        return ("Selamat pagi" if h < 12 else "Selamat tengah hari" if h < 14
                else "Selamat petang" if h < 19 else "Selamat malam")
    return "Good morning" if h < 12 else "Good afternoon" if h < 18 else "Good evening"


LINK_TYPES = {
    "WhatsApp Web direct (default)": "https://web.whatsapp.com/send?phone={m}&text={t}",
    "wa.me": "https://wa.me/{m}?text={t}",
    "api.whatsapp.com": "https://api.whatsapp.com/send?phone={m}&text={t}",
}


def make_link(mobile: str, message: str, kind: str = "WhatsApp Web direct (default)") -> str:
    return LINK_TYPES[kind].format(m=mobile, t=quote(message, safe="", encoding="utf-8"))


# --------------------------------------------------------------------- UI
st.title("💬 Customer Follow-Up Messages")
st.caption("Upload the sales export, review the messages, then click each WhatsApp link and press Send.")

with st.sidebar:
    st.header("Settings")
    preset_name = st.selectbox("Follow-up type", list(PRESETS))
    P = PRESETS[preset_name]
    k = preset_name  # widget keys are per preset so each keeps its own values

    language = st.selectbox("Message language", ["English", "Bahasa Malaysia", "Bahasa Sarawak"])
    outlet = st.text_input("Outlet name in message", value="Alpro Pharmacy")

    st.subheader("Product")
    product = st.text_input("Product name used in message", P["product"], key=f"prod_{k}")
    kw_text = st.text_area("Item name keywords (any match counts)", P["keywords"], height=90, key=f"kw_{k}",
                           help="Matched against item name and description. Separate with commas.")
    ex_text = st.text_area("Exclude items containing", P["exclude"], height=60, key=f"ex_{k}",
                           help="For example GUMMYKIDS to leave out children's products.")

    benefit = st.text_area("What customers may feel (used where the message has {benefit_line})", "", height=70,
                           key=f"ben_{k}", placeholder="e.g. more energy through the day",
                           help="Use wording from the product label or approved claims. Leave empty to skip this sentence.")

    st.subheader("Promotion and greeting")
    promotion = st.text_area("Promotion for everyone (used when no rule file is uploaded or an item has no rule)", "", height=80, key=f"promo_{k}",
                             placeholder="e.g. Buy 2 get 10% off until 30 Nov")
    link_kind = st.selectbox("WhatsApp link type", list(LINK_TYPES),
                             help="If emoji or line breaks come out wrong, try another type.")
    photo_url = st.text_input("Photo link to add under the message (optional)", "", key=f"purl_{k}",
                              placeholder="https://... (a public link to the photo)",
                              help="Customers can open this link. A link on this PC or the outlet network will not work for them.")
    skip_no_promo = st.checkbox("Leave out customers whose item has no promotion rule", True, key=f"snp_{k}",
                                help="Only applies when you upload the promotion rule file.")
    use_main_pack = st.checkbox("If a pack has no rule, use the rule of the main pack", False, key=f"ump_{k}",
                                help="Example: T73051 (1S) uses the rule of 73051 (30S). The rule may be for a "
                                     "bigger pack than the one the customer bought.")
    greet_mode = st.radio("Greeting (used where the message has {greeting})",
                          ["Auto (by time of day)", "Custom text"], key=f"gm_{k}")
    greet_custom = st.text_input("Custom greeting", "", key=f"gc_{k}", disabled=greet_mode != "Custom text")

    st.subheader("Cross-sell check (optional)")
    cross_on = st.checkbox("Check if they also bought another product", P["cross_on"], key=f"con_{k}")
    cross = st.text_input("Cross-sell product name", P["cross"], key=f"cn_{k}", disabled=not cross_on)
    cross_text = st.text_area("Cross-sell item codes or name keywords", P["cross_tokens"], height=90,
                              key=f"ct_{k}", disabled=not cross_on)

    st.subheader("Timing")
    min_days = st.number_input("Only customers whose last purchase was at least this many days ago",
                               min_value=0, value=0, step=1)

if not kw_text.strip():
    st.info("Enter at least one item keyword in the sidebar to find customers.")
    st.stop()

upload = st.file_uploader("Sales invoice item export (.xlsx)", type=["xlsx"])
if upload is None:
    st.info("Upload the sales export to begin.")
    st.stop()

promo_upload = st.file_uploader(
    "In-house promotion rule file (.xlsx), optional: gives each item its own promotion from the Attribute column",
    type=["xlsx"])
promo_rules = {}
if promo_upload is not None:
    try:
        promo_rules, promo_review = load_promo_rules(promo_upload)
    except Exception as e:  # noqa: BLE001
        st.error(f"Could not read the promotion file: {e}")
        st.stop()
    n_ended = int((promo_review["Status"] != "OK").sum())
    st.caption(f"{len(promo_rules)} items with a running promotion loaded"
               + (f"; {n_ended} item(s) skipped because their promotion has already ended." if n_ended else "."))
    with st.expander("Check promotion wording (original vs what the customer sees)"):
        st.dataframe(promo_review, hide_index=True, use_container_width=True)

try:
    raw = pd.read_excel(upload, dtype={"Client Mobile": str, "Item Code": str})
except Exception as e:  # noqa: BLE001
    st.error(f"Could not read the file: {e}")
    st.stop()

raw.columns = [str(c).strip() for c in raw.columns]
missing = [c for c in REQUIRED if c not in raw.columns]
if missing:
    st.error(f"Missing column(s): {', '.join(missing)}. Please upload the original sales export.")
    st.stop()

cross_tokens = set(split_list(cross_text)) if cross_on else set()
cust, skipped = build_followups(raw, split_list(kw_text), split_list(ex_text), cross_tokens, promo_rules,
                                use_main_pack)
cust = cust[cust["days_since"] >= min_days].reset_index(drop=True)

no_promo_left_out = 0
if promo_rules and skip_no_promo:
    no_promo_left_out = int((cust["promo_rule"] == "").sum())
    cust = cust[cust["promo_rule"] != ""].reset_index(drop=True)

if cust.empty:
    if no_promo_left_out:
        st.warning(f"{no_promo_left_out} customer(s) found, but none of their items has a promotion rule. "
                   "Untick 'Leave out customers whose item has no promotion rule' to include them.")
    else:
        st.warning("No customers found with these settings.")
    st.stop()

tpl_set = P["templates"].get(language) or P["templates"]["Bahasa Malaysia"]  # presets without a Sarawak version use Malay
with st.expander("Edit message wording (use {name}, {outlet}, {product}, {item}, {greeting}, {promotion}, {benefit_line}, {cross})"):
    if cross_on:
        tpl_no = st.text_area(f"For customers who have NOT bought {cross or 'the cross-sell product'}",
                              tpl_set["no_cross"], height=170, key=f"tn_{k}_{language}")
        tpl_yes = st.text_area(f"For customers who HAVE bought {cross or 'the cross-sell product'}",
                               tpl_set["has_cross"] or tpl_set["no_cross"], height=140, key=f"ty_{k}_{language}")
    else:
        tpl_no = st.text_area("Message", tpl_set["no_cross"], height=170, key=f"tn_{k}_{language}")
        tpl_yes = tpl_no


greeting = make_greeting(language, greet_mode, greet_custom)
if benefit.strip():
    _b = benefit.strip()
    benefit_line = {
        "Bahasa Malaysia": f"Anda mungkin rasa {_b} jika mengambilnya secara berterusan \u2728\n\n",
        "Bahasa Sarawak": f"Kitak mungkin rasa {_b} kalau guna nya selalu \u2728\n\n",
    }.get(language, f"You may feel {_b} when you take it regularly \u2728\n\n")
else:
    benefit_line = ""


def message_for(row) -> str:
    tpl = tpl_yes if (cross_on and row["has_cross"]) else tpl_no
    msg = tpl.format(benefit_line=benefit_line, name=row["name"], outlet=outlet.strip() or "Alpro Pharmacy",
                      product=product.strip() or "medicine", item=row["item"],
                      greeting=greeting, promotion=row["promo_rule"] or promotion.strip() or "(promotion to be added)",
                      cross=cross.strip())
    return msg + (f"\n{photo_url.strip()}" if photo_url.strip() else "")


if "{promotion}" in (tpl_no + tpl_yes) and (cust["promo_rule"] == "").any() and not promotion.strip():
    st.warning("Some customers have no promotion rule and the Promotion box in the sidebar is empty, "
               "so their message shows '(promotion to be added)'. Fill the box or leave those customers out.")
if no_promo_left_out:
    st.caption(f"{no_promo_left_out} customer(s) left out because their item has no promotion rule.")
if promo_rules:
    st.caption("Dates are removed from the promotion wording. Please confirm each promotion is still running "
               "before sending.")

try:
    cust["message"] = cust.apply(message_for, axis=1)
except (KeyError, IndexError, ValueError):
    st.error("The message wording has a stray { or } or an unknown field. "
             "Only {name}, {outlet}, {product}, {item}, {greeting}, {promotion}, {benefit_line} and {cross} are allowed.")
    st.stop()

cust["link"] = [make_link(m, msg, link_kind) for m, msg in zip(cust["mobile"], cust["message"])]

if cross_on:
    m1, m2, m3 = st.columns(3)
    m1.metric("Customers", len(cust))
    m2.metric(f"Already buy {cross or 'cross-sell'}", int(cust["has_cross"].sum()))
    m3.metric("Not yet", int((~cust["has_cross"]).sum()))
else:
    st.metric("Customers", len(cust))

photo_file = st.file_uploader("Photo to send with the message (optional)", type=["png", "jpg", "jpeg", "webp"])
if photo_file is not None:
    pc1, pc2 = st.columns([1, 3])
    with pc1:
        st.image(photo_file, width=220)
    with pc2:
        st.info("WhatsApp links can only carry text, so the photo cannot attach by itself. "
                "Right-click the photo and choose **Copy image**. After each WhatsApp chat opens, "
                "press **Ctrl + V** to paste the photo, then press Send. "
                "Works best in WhatsApp Web or WhatsApp Desktop on a PC.")
        st.download_button("⬇️ Download photo", photo_file.getvalue(), file_name=photo_file.name,
                           mime=photo_file.type)

cross_col = f"Bought {cross or 'cross-sell'}"
view = pd.DataFrame({
    "Sent": False,
    "Name": cust["name"],
    "Mobile": cust["mobile"],
    "Last purchase": cust["last_date"].dt.strftime("%Y-%m-%d"),
    "Days since": cust["days_since"],
    "Item": cust["last_item"],
})
if cross_on:
    view[cross_col] = cust["has_cross"].map({True: "Yes", False: "No"})
if promo_rules or promotion.strip():
    view["Promotion"] = [r or promotion.strip() for r in cust["promo_rule"]]
view["WhatsApp"] = cust["link"]

edited = st.data_editor(
    view,
    hide_index=True,
    use_container_width=True,
    disabled=[c for c in view.columns if c != "Sent"],
    column_config={
        "Sent": st.column_config.CheckboxColumn("Sent ✔", help="Tick after you send, to keep track"),
        "WhatsApp": st.column_config.LinkColumn("WhatsApp", display_text="Open WhatsApp"),
    },
)
st.caption(f"Sent so far: {int(edited['Sent'].sum())} of {len(edited)} "
           "(ticks are lost if you refresh or close the page)")

st.subheader("Message preview")
labels = cust["name"] + " – " + cust["mobile"]
pick = st.selectbox("Customer", labels)
st.info(cust.loc[labels == pick, "message"].iloc[0].replace("\n", "  \n"))

if not skipped.empty:
    with st.expander(f"⚠️ {len(skipped)} purchase(s) skipped: missing or invalid mobile number"):
        st.dataframe(skipped, hide_index=True, use_container_width=True)

export = view.drop(columns=["Sent", "WhatsApp"]).copy()
export["Message"] = cust["message"]
export["WhatsApp link"] = cust["link"]
buf = io.BytesIO()
export.to_excel(buf, index=False, sheet_name="Follow-up")
st.download_button("⬇️ Download list (.xlsx)", buf.getvalue(),
                   file_name=f"FollowUp_{(product or 'list').replace(' ', '_')}_{datetime.now():%Y%m%d}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
