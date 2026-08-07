#!/usr/bin/env python3
"""Generate a dedicated detail page for each Stride wheelchair model.

Product data below is the single source of truth for the detail pages.
Run from the repo root:  python3 scripts/generate_products.py
Outputs: products/<slug>.html
"""
import os
from urllib.parse import quote

WA = "923001234567"  # TODO: replace with the real WhatsApp number
GENERAL_WA = f"https://wa.me/{WA}?text=" + quote(
    "Hi Stride, I'd like to know more about your electric wheelchairs."
)

BADGES = {
    "in": ("badge-in", "In stock"),
    "limited": ("badge-limited", "Limited stock"),
    "pre": ("badge-pre", "Pre-order"),
}

PRODUCTS = [
    {
        "slug": "glide-s1", "name": "Stride Glide S1", "tag": "Folding · Entry",
        "badge": "in", "price": "PKR 185,000", "media": "media-teal",
        "lead": "A lightweight folding power chair built for indoors and short trips. Simple, "
                "reliable and genuinely easy to store — it folds in one motion and fits in a car boot, "
                "so getting out of the house never feels like a project.",
        "quick": [("Range", "15 km"), ("Top speed", "6 km/h"), ("Max load", "100 kg"), ("Foldable", "Yes")],
        "specs": [
            ("Range per charge", "15 km"), ("Top speed", "6 km/h"), ("Max load", "100 kg"),
            ("Kerb weight", "24 kg"), ("Battery", "Li-ion 24V 12Ah"), ("Motor", "Brushless 250W"),
            ("Charge time", "6–8 hrs"), ("Foldable", "Yes (one-fold)"),
            ("Tyres", "Solid, puncture-free"), ("Controls", "Joystick"), ("Warranty", "1 year"),
        ],
        "highlights": ["Folds in one motion for car boots", "Lightweight, easy-to-handle frame",
                       "Puncture-free solid tyres", "Detachable battery for easy charging",
                       "Intuitive joystick control"],
        "bestfor": "Indoor use and short daily outings.",
    },
    {
        "slug": "urban-u2", "name": "Stride Urban U2", "tag": "City commuter",
        "badge": "in", "price": "PKR 235,000", "media": "media-indigo",
        "lead": "A balanced everyday chair for city life. Comfortable seating, responsive controls and a "
                "dependable battery make it a confident companion for errands, visits and daily commutes.",
        "quick": [("Range", "22 km"), ("Top speed", "8 km/h"), ("Max load", "120 kg"), ("Foldable", "Yes")],
        "specs": [
            ("Range per charge", "22 km"), ("Top speed", "8 km/h"), ("Max load", "120 kg"),
            ("Kerb weight", "27 kg"), ("Battery", "Li-ion 24V 20Ah"), ("Motor", "Brushless 300W"),
            ("Charge time", "6–8 hrs"), ("Foldable", "Yes"), ("Suspension", "Front"),
            ("Tyres", "Anti-slip solid"), ("Controls", "Joystick"), ("Warranty", "1 year"),
        ],
        "highlights": ["Comfortable all-day seating", "Front suspension for smoother rides",
                       "Responsive, easy handling", "Foldable for transport", "Reliable city range"],
        "bestfor": "Everyday city travel and errands.",
    },
    {
        "slug": "cruise-c3", "name": "Stride Cruise C3", "tag": "Comfort · Mid",
        "badge": "in", "price": "PKR 265,000", "media": "media-teal",
        "lead": "Extra padding, adjustable armrests and a smoother ride for longer daily use — indoors and "
                "out. The Cruise C3 is our sweet spot of comfort, range and value.",
        "quick": [("Range", "25 km"), ("Top speed", "8 km/h"), ("Max load", "120 kg"), ("Foldable", "Yes")],
        "specs": [
            ("Range per charge", "25 km"), ("Top speed", "8 km/h"), ("Max load", "120 kg"),
            ("Kerb weight", "30 kg"), ("Battery", "Li-ion 24V 20Ah"), ("Motor", "Brushless 300W"),
            ("Charge time", "6–8 hrs"), ("Foldable", "Yes"), ("Seat", "Padded, adjustable armrests"),
            ("Tyres", "PU solid"), ("Controls", "Joystick"), ("Warranty", "1 year"),
        ],
        "highlights": ["Padded seat with adjustable armrests", "Smooth, stable ride quality",
                       "Strong 25 km range", "Comfortable for longer sessions", "Foldable frame"],
        "bestfor": "Longer daily use with extra comfort.",
    },
    {
        "slug": "compact-air", "name": "Stride Compact Air", "tag": "Ultra-light · Travel",
        "badge": "in", "price": "PKR 320,000", "media": "media-amber",
        "lead": "Just 18 kg and airline-friendly. The Compact Air folds in seconds and lifts easily — ideal "
                "for travel, hotels and getting in and out of taxis without a struggle.",
        "quick": [("Range", "20 km"), ("Top speed", "6 km/h"), ("Max load", "110 kg"), ("Weight", "18 kg")],
        "specs": [
            ("Range per charge", "20 km"), ("Top speed", "6 km/h"), ("Max load", "110 kg"),
            ("Kerb weight", "18 kg"), ("Battery", "Li-ion (airline-safe)"), ("Motor", "Brushless 250W"),
            ("Charge time", "5–7 hrs"), ("Foldable", "Yes (auto-fold option)"),
            ("Tyres", "Solid"), ("Controls", "Joystick"), ("Warranty", "1 year"),
        ],
        "highlights": ["Only 18 kg — easy to lift", "Airline-safe lithium battery",
                       "Folds in seconds", "Perfect for travel and taxis", "Optional auto-fold remote"],
        "bestfor": "Travel, flights and frequent transfers.",
    },
    {
        "slug": "terra-x", "name": "Stride Terra X", "tag": "All-terrain · Dual motor",
        "badge": "limited", "price": "PKR 420,000", "media": "media-indigo",
        "lead": "Dual-motor power and rugged pneumatic tyres for rough roads, ramps and outdoor use. Built "
                "for range and stability, the Terra X keeps going where lighter chairs stop.",
        "quick": [("Range", "35 km"), ("Top speed", "10 km/h"), ("Max load", "150 kg"), ("Motor", "Dual")],
        "specs": [
            ("Range per charge", "35 km"), ("Top speed", "10 km/h"), ("Max load", "150 kg"),
            ("Kerb weight", "42 kg"), ("Battery", "Li-ion 24V 30Ah"), ("Motor", "Dual 2×350W"),
            ("Charge time", "8–10 hrs"), ("Foldable", "No"), ("Suspension", "Full"),
            ("Tyres", "Pneumatic off-road"), ("Controls", "Joystick"), ("Warranty", "1 year"),
        ],
        "highlights": ["Dual 350W motors for tough terrain", "Long 35 km range",
                       "Full suspension for stability", "High 150 kg load capacity",
                       "Pneumatic off-road tyres"],
        "bestfor": "Rough roads, ramps and outdoor terrain.",
    },
    {
        "slug": "recline-r5", "name": "Stride Recline R5", "tag": "Full recline · High support",
        "badge": "pre", "price": "PKR 495,000", "media": "media-amber",
        "lead": "A full reclining backrest and elevating legrests for maximum comfort and pressure relief "
                "during long sitting. The Recline R5 is our highest-support chair, built for extended use.",
        "quick": [("Range", "30 km"), ("Top speed", "8 km/h"), ("Max load", "135 kg"), ("Recline", "Full")],
        "specs": [
            ("Range per charge", "30 km"), ("Top speed", "8 km/h"), ("Max load", "135 kg"),
            ("Kerb weight", "38 kg"), ("Battery", "Li-ion 24V 25Ah"), ("Motor", "Brushless 350W"),
            ("Charge time", "8–10 hrs"), ("Backrest", "Full recline"), ("Legrests", "Elevating"),
            ("Controls", "Joystick + recline"), ("Warranty", "1 year"),
        ],
        "highlights": ["Full reclining backrest", "Elevating legrests", "Excellent pressure relief",
                       "High-support seating", "Powerful 350W motor"],
        "bestfor": "Extended sitting and maximum pressure relief.",
    },
]

SYMBOLS = """  <svg width="0" height="0" style="position:absolute" aria-hidden="true">
    <symbol id="ic-wheelchair" viewBox="0 0 24 24">
      <circle cx="8.5" cy="18" r="3.2" fill="none" stroke="currentColor" stroke-width="1.4"/>
      <circle cx="8.5" cy="18" r="0.6" fill="currentColor"/>
      <circle cx="6.5" cy="6" r="1.8" fill="currentColor"/>
      <path d="M6.5 8.4v4.2h4.4l2.6 4.6" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>
      <path d="M6.9 10.4h4" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/>
      <path d="M13.5 17.2h3.2l1.8-4" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>
    </symbol>
    <symbol id="ic-wa" viewBox="0 0 24 24">
      <path d="M12 2a10 10 0 0 0-8.6 15.1L2 22l5.1-1.3A10 10 0 1 0 12 2Zm5.4 14.3c-.2.6-1.3 1.2-1.8 1.2-.5.1-1 .2-3.3-.7-2.8-1.1-4.5-3.9-4.7-4.1-.1-.2-1.1-1.4-1.1-2.7 0-1.3.7-1.9.9-2.2.2-.2.5-.3.7-.3h.5c.2 0 .4 0 .6.5l.8 1.9c.1.2.1.4 0 .5l-.4.6c-.2.2-.3.4-.1.7.2.3.9 1.4 1.9 2.3 1.3 1.1 2.3 1.5 2.6 1.6.3.1.5.1.7-.1l.9-1c.2-.3.4-.2.7-.1l1.9.9c.3.1.5.2.6.3.1.3.1.8-.1 1.4Z" fill="currentColor"/>
    </symbol>
    <symbol id="ic-check" viewBox="0 0 24 24">
      <path d="M20 6 9 17l-5-5" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>
    </symbol>
    <symbol id="ic-truck" viewBox="0 0 24 24">
      <path d="M2.5 6h11v9h-11z" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/>
      <path d="M13.5 9h4l3.5 3.2V15h-7.5z" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/>
      <circle cx="7" cy="17.5" r="1.9" fill="none" stroke="currentColor" stroke-width="1.6"/>
      <circle cx="17.5" cy="17.5" r="1.9" fill="none" stroke="currentColor" stroke-width="1.6"/>
    </symbol>
  </svg>"""


def wa_link(p):
    if p["badge"] == "pre":
        msg = f"Hi Stride, I'd like to pre-order the {p['name']}. Please share details and timeline."
    else:
        msg = f"Hi Stride, I'm interested in the {p['name']} electric wheelchair. Please share availability and final price."
    return f"https://wa.me/{WA}?text={quote(msg)}"


def render(p):
    badge_class, badge_text = BADGES[p["badge"]]
    quick = "\n".join(
        f'              <li><span>{k}</span><strong>{v}</strong></li>' for k, v in p["quick"]
    )
    specs = "\n".join(
        f'              <tr><th>{k}</th><td>{v}</td></tr>' for k, v in p["specs"]
    )
    highlights = "\n".join(
        f'              <li><svg class="icon" aria-hidden="true"><use href="#ic-check"/></svg><span>{h}</span></li>'
        for h in p["highlights"]
    )
    cta_label = "Pre-order on WhatsApp" if p["badge"] == "pre" else "Enquire on WhatsApp"
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{p['name']} — Stride Electric Wheelchairs</title>
  <meta name="description" content="{p['lead'][:150]}">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Google+Sans+Flex:wght@400..800&display=swap">
  <link rel="stylesheet" href="../style.css">
</head>
<body>

{SYMBOLS}

  <div class="topbar" id="top">
    <div class="container topbar-inner">
      <span class="topbar-msg"><svg class="icon" aria-hidden="true"><use href="#ic-truck"/></svg> Free delivery across Pakistan</span>
      <span class="topbar-contact">
        <a href="{GENERAL_WA}" target="_blank" rel="noopener"><svg class="icon" aria-hidden="true"><use href="#ic-wa"/></svg> +92 300 1234567</a>
        <span class="topbar-sep" aria-hidden="true">·</span>
        <span class="topbar-hours">Mon–Sat, 10am–8pm</span>
      </span>
    </div>
  </div>
  <header class="site-header" id="headerMain">
    <div class="container header-inner">
      <a href="../index.html" class="brand" aria-label="Stride home">
        <span class="brand-mark"><svg class="icon" aria-hidden="true"><use href="#ic-wheelchair"/></svg></span>
        <span class="brand-name">stride</span>
      </a>
      <nav class="nav" aria-label="Primary">
        <a href="../index.html#models">Models</a>
        <a href="../index.html#why">Why Stride</a>
        <a href="../index.html#contact">Contact</a>
      </nav>
      <a class="btn btn-wa header-cta" href="{wa_link(p)}" target="_blank" rel="noopener">
        <svg class="icon" aria-hidden="true"><use href="#ic-wa"/></svg>
        <span>Chat on WhatsApp</span>
      </a>
      <button class="nav-toggle" id="navToggle" aria-label="Toggle menu" aria-expanded="false">
        <span></span><span></span><span></span>
      </button>
    </div>
    <nav class="mobile-nav" id="mobileNav" aria-label="Mobile">
      <a href="../index.html#models">Models</a>
      <a href="../index.html#why">Why Stride</a>
      <a href="../index.html#contact">Contact</a>
      <a class="btn btn-wa" href="{wa_link(p)}" target="_blank" rel="noopener">Chat on WhatsApp</a>
    </nav>
  </header>

  <main>
    <section class="section pdp">
      <div class="container">
        <a class="back-link" href="../index.html#models">← All models</a>

        <div class="pdp-hero">
          <div class="pdp-media {p['media']}">
            <span class="badge {badge_class}">{badge_text}</span>
            <svg class="pdp-chair"><use href="#ic-wheelchair"/></svg>
          </div>
          <div class="pdp-info">
            <span class="tag">{p['tag']}</span>
            <h1>{p['name']}</h1>
            <p class="pdp-lead">{p['lead']}</p>
            <div class="pdp-price"><span>From</span>{p['price']}</div>
            <a class="btn btn-wa btn-lg" href="{wa_link(p)}" target="_blank" rel="noopener">
              <svg class="icon" aria-hidden="true"><use href="#ic-wa"/></svg><span>{cta_label}</span>
            </a>
            <ul class="pdp-quick">
{quick}
            </ul>
          </div>
        </div>

        <div class="pdp-grid">
          <section class="pdp-block">
            <h2>Full specifications</h2>
            <table class="spec-table">
{specs}
            </table>
          </section>
          <section class="pdp-block">
            <h2>Highlights</h2>
            <ul class="highlight-list">
{highlights}
            </ul>
            <div class="bestfor"><strong>Best for:</strong> {p['bestfor']}</div>
          </section>
        </div>
      </div>
    </section>

    <section class="section contact" id="contact">
      <div class="container">
        <div class="contact-card">
          <div class="contact-copy">
            <h2>Questions about the {p['name']}?</h2>
            <p>Message us on WhatsApp for live stock, delivery time to your city and the final quote. We'll send real photos and answer anything.</p>
            <div class="contact-meta">
              <p><strong>WhatsApp / Phone:</strong> +92 300 1234567</p>
              <p><strong>Hours:</strong> Mon–Sat, 10am–8pm · Serving all of Pakistan</p>
            </div>
          </div>
          <div class="contact-cta">
            <a class="btn btn-wa btn-lg" href="{wa_link(p)}" target="_blank" rel="noopener">
              <svg class="icon" aria-hidden="true"><use href="#ic-wa"/></svg>
              <span>{cta_label}</span>
            </a>
          </div>
        </div>
      </div>
    </section>
  </main>

  <footer class="site-footer">
    <div class="container footer-inner">
      <div class="footer-brand">
        <span class="brand-mark"><svg class="icon" aria-hidden="true"><use href="#ic-wheelchair"/></svg></span>
        <span class="brand-name">stride</span>
      </div>
      <p class="footer-tag">Electric wheelchairs for independent living, delivered across Pakistan.</p>
      <nav class="footer-nav" aria-label="Footer">
        <a href="../index.html#models">Models</a>
        <a href="../index.html#why">Why Stride</a>
        <a href="../index.html#contact">Contact</a>
      </nav>
    </div>
    <div class="container footer-bottom">
      <span>© <span id="year">2026</span> Stride. All rights reserved.</span>
      <span>Prices in PKR and subject to change. Contact us for current pricing.</span>
    </div>
  </footer>

  <a class="wa-float" href="{wa_link(p)}" target="_blank" rel="noopener" aria-label="Chat on WhatsApp">
    <svg class="icon" aria-hidden="true"><use href="#ic-wa"/></svg>
  </a>

  <script src="../app.js"></script>
</body>
</html>
"""


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out_dir = os.path.join(root, "products")
    os.makedirs(out_dir, exist_ok=True)
    for p in PRODUCTS:
        path = os.path.join(out_dir, f"{p['slug']}.html")
        with open(path, "w") as f:
            f.write(render(p))
        print(f"wrote products/{p['slug']}.html")


if __name__ == "__main__":
    main()
