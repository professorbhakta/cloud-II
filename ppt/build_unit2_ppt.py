#!/usr/bin/env python3
"""Build a cloud-themed Unit II PowerPoint from course notes + real-life examples."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import nsmap
from pptx.oxml import parse_xml
from copy import deepcopy
import os

# ── Cloud theme palette (deep sky / navy — not purple-AI defaults) ──────────
NAVY = RGBColor(0x0A, 0x1B, 0x2E)
SKY = RGBColor(0x1A, 0x6F, 0xB5)
SKY_LIGHT = RGBColor(0x3D, 0x9B, 0xE0)
TEAL = RGBColor(0x0D, 0xA5, 0xA0)
CORAL = RGBColor(0xE8, 0x6A, 0x4C)
CLOUD = RGBColor(0xF4, 0xF8, 0xFC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SOFT = RGBColor(0xE8, 0xF1, 0xF8)
INK = RGBColor(0x1C, 0x2B, 0x3A)
MUTED = RGBColor(0x5A, 0x6B, 0x7C)
GOLD = RGBColor(0xD4, 0xA0, 0x17)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H


def add_rect(slide, left, top, width, height, fill, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
    return shape


def add_rounded(slide, left, top, width, height, fill, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
    # Softer corners
    try:
        shape.adjustments[0] = 0.08
    except Exception:
        pass
    return shape


def set_run(run, text, size=18, bold=False, color=INK, font="Calibri"):
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font


def tb(slide, left, top, width, height, lines, default_size=16, default_color=INK,
       align=PP_ALIGN.LEFT, font="Calibri"):
    """lines: list of str or (text, size, bold, color) tuples."""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for item in lines:
        if isinstance(item, str):
            text, size, bold, color = item, default_size, False, default_color
        else:
            text = item[0]
            size = item[1] if len(item) > 1 else default_size
            bold = item[2] if len(item) > 2 else False
            color = item[3] if len(item) > 3 else default_color
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(4)
        run = p.add_run()
        set_run(run, text, size=size, bold=bold, color=color, font=font)
    return box


def footer(slide, page, total=None):
    add_rect(slide, Inches(0), Inches(7.15), SLIDE_W, Inches(0.35), NAVY)
    tb(slide, Inches(0.4), Inches(7.18), Inches(8), Inches(0.3),
       [("CC-II  ·  Unit II  ·  Networking in the Cloud  ·  Parul University", 10, False, SOFT)])
    label = f"{page}" if total is None else f"{page} / {total}"
    tb(slide, Inches(11.5), Inches(7.18), Inches(1.5), Inches(0.3),
       [(label, 10, False, SOFT)], align=PP_ALIGN.RIGHT)


def header_bar(slide, title, subtitle=None):
    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.15), NAVY)
    add_rect(slide, Inches(0), Inches(1.15), SLIDE_W, Inches(0.08), TEAL)
    tb(slide, Inches(0.5), Inches(0.22), Inches(12), Inches(0.5),
       [(title, 28, True, WHITE)], font="Calibri")
    if subtitle:
        tb(slide, Inches(0.5), Inches(0.7), Inches(12), Inches(0.35),
           [(subtitle, 13, False, SKY_LIGHT)])


def blank_content_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    add_rect(slide, Inches(0), Inches(0), SLIDE_W, SLIDE_H, CLOUD)
    return slide


def card(slide, left, top, width, height, title, body_lines, accent=SKY):
    add_rounded(slide, left, top, width, height, WHITE)
    add_rect(slide, left, top, Inches(0.1), height, accent)
    tb(slide, left + Inches(0.25), top + Inches(0.15), width - Inches(0.4), Inches(0.4),
       [(title, 15, True, NAVY)])
    formatted = []
    for line in body_lines:
        if isinstance(line, str):
            formatted.append((line, 12, False, MUTED))
        else:
            formatted.append(line)
    tb(slide, left + Inches(0.25), top + Inches(0.55), width - Inches(0.4), height - Inches(0.7),
       formatted, default_size=12, default_color=MUTED)


def example_banner(slide, top, text):
    """Highlight real-life example strip."""
    add_rounded(slide, Inches(0.5), top, Inches(12.3), Inches(0.55), RGBColor(0xFF, 0xF4, 0xEC))
    tb(slide, Inches(0.7), top + Inches(0.1), Inches(12), Inches(0.4),
       [(f"🌍  Real life:  {text}", 13, True, CORAL)])


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE BUILDERS
# ═══════════════════════════════════════════════════════════════════════════

PAGE = 0


def next_page():
    global PAGE
    PAGE += 1
    return PAGE


# ── 1. Title ───────────────────────────────────────────────────────────────
s = blank_content_slide()
add_rect(s, Inches(0), Inches(0), SLIDE_W, SLIDE_H, NAVY)
# Decorative sky bands
add_rect(s, Inches(0), Inches(0), SLIDE_W, Inches(0.15), TEAL)
add_rect(s, Inches(0), Inches(7.35), SLIDE_W, Inches(0.15), SKY)
# Soft cloud shapes (ellipses)
for left, top, w, h in [
    (10.2, 0.8, 2.8, 1.2), (0.3, 5.8, 3.2, 1.0), (9.5, 5.5, 3.5, 1.3),
]:
    sh = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(left), Inches(top), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = RGBColor(0x14, 0x2E, 0x4A)
    sh.line.fill.background()

tb(s, Inches(0.8), Inches(1.8), Inches(11.5), Inches(0.4),
   [("CLOUD COMPUTING – II", 16, True, TEAL)], font="Calibri")
tb(s, Inches(0.8), Inches(2.3), Inches(11.5), Inches(1.0),
   [("Networking in the Cloud", 44, True, WHITE)], font="Calibri")
tb(s, Inches(0.8), Inches(3.4), Inches(11.5), Inches(0.5),
   [("Unit II  ·  Base Presentation", 20, False, SKY_LIGHT)])
tb(s, Inches(0.8), Inches(4.3), Inches(11.5), Inches(0.8),
   [
       ("Google App Engine  ·  Compute Engine  ·  Load Balancing  ·  GKE", 14, False, SOFT),
       ("Cloud Functions  ·  DR  ·  IaC  ·  CI/CD  + Real-world case studies", 14, False, SOFT),
   ])
tb(s, Inches(0.8), Inches(5.8), Inches(11.5), Inches(0.5),
   [("FIT & CS  ·  Parul University", 14, False, MUTED)])
next_page()


# ── 2. Agenda ──────────────────────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Learning Roadmap", "Eleven topics that build a production-ready cloud networking mindset")
footer(s, next_page())

topics = [
    ("01", "App Engine & Service Models", SKY),
    ("02", "Compute Engine & Scalable Web Apps", TEAL),
    ("03", "MIG, Load Balancers, CDN & Autohealing", CORAL),
    ("04", "App Engine Flexible & PaaS Deep Dive", SKY),
    ("05", "Event-Driven Cloud Functions", TEAL),
    ("06–07", "GKE Overview & Container Deployments", CORAL),
    ("08", "Cloud Principles & Well-Architected", SKY),
    ("09", "Disaster Recovery in the Cloud", TEAL),
    ("10–11", "Infrastructure as Code & CI/CD", CORAL),
]
for i, (num, title, color) in enumerate(topics):
    col = i % 3
    row = i // 3
    left = Inches(0.5 + col * 4.2)
    top = Inches(1.55 + row * 1.7)
    add_rounded(s, left, top, Inches(3.9), Inches(1.4), WHITE)
    add_rect(s, left, top, Inches(3.9), Inches(0.12), color)
    tb(s, left + Inches(0.25), top + Inches(0.35), Inches(3.4), Inches(0.4),
       [(num, 22, True, color)])
    tb(s, left + Inches(0.25), top + Inches(0.8), Inches(3.4), Inches(0.45),
       [(title, 13, True, NAVY)])


# ── 3. Why Unit II matters ─────────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Why “Networking in the Cloud” Matters", "Unit II connects compute, traffic, resilience, and delivery")
footer(s, next_page())
points = [
    ("Compute choices", "PaaS (App Engine), IaaS (Compute Engine), FaaS (Functions), containers (GKE) — pick the right abstraction."),
    ("Traffic fabric", "Load balancers, named ports, health checks, and CDN move users to healthy backends worldwide."),
    ("Resilience", "MIG autohealing, rolling updates, multi-zone GKE, and DR strategies keep apps alive under failure."),
    ("Delivery speed", "IaC + CI/CD turn Fancy Store / GKE labs into repeatable production pipelines."),
]
for i, (t, b) in enumerate(points):
    top = Inches(1.45 + i * 1.25)
    add_rounded(s, Inches(0.5), top, Inches(12.3), Inches(1.1), WHITE)
    add_rect(s, Inches(0.5), top, Inches(0.12), Inches(1.1), [SKY, TEAL, CORAL, GOLD][i])
    tb(s, Inches(0.9), top + Inches(0.2), Inches(11.5), Inches(0.35), [(t, 16, True, NAVY)])
    tb(s, Inches(0.9), top + Inches(0.55), Inches(11.5), Inches(0.4), [(b, 13, False, MUTED)])


# ── 4. Topic 1 — GAE ───────────────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Topic 1 · Google App Engine (GAE)", "Fully managed PaaS — write code, Google runs the platform")
footer(s, next_page())
card(s, Inches(0.5), Inches(1.45), Inches(6.0), Inches(3.6), "What is GAE?", [
    "• Platform-as-a-Service on Google Cloud",
    "• Google handles servers, OS, scaling, LB",
    "• You deploy app code — not VMs",
    "• Sandboxed runtimes + auto-scaling",
    "• Built-in APIs: Memcache, URL Fetch,",
    "  Cloud Storage, Blobstore (legacy)",
    "• Free tier for CPU, storage, API calls",
], SKY)
card(s, Inches(6.8), Inches(1.45), Inches(6.0), Inches(3.6), "Standard vs Flexible", [
    "STANDARD",
    "• Strict sandbox, fast cold starts",
    "• Fine-grained auto scale, lower cost",
    "",
    "FLEXIBLE",
    "• Docker containers on GCE VMs",
    "• Custom runtimes & longer requests",
    "• Full VPC networking options",
], TEAL)
example_banner(s, Inches(5.3),
               "Snapchat, Rovio (Angry Birds), and Khan Academy scaled early on App Engine — proving PaaS can serve millions without ops teams babysitting VMs.")


# ── 5. IaaS / PaaS / SaaS ──────────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Cloud Service Models — Where GAE Fits", "Architecture decisions start with who manages what")
footer(s, next_page())
models = [
    ("IaaS", "Infrastructure", "You manage OS, runtime, apps\nProvider: hardware + virt",
     "Compute Engine, Amazon EC2", CORAL),
    ("PaaS", "Platform", "You deploy code only\nProvider: OS, scale, runtime",
     "App Engine, Heroku, Azure App Service", SKY),
    ("SaaS", "Software", "Ready-to-use product\nYou just use the UI / API",
     "Gmail, Docs, Salesforce, Zoom", TEAL),
]
for i, (abbr, name, desc, examples, color) in enumerate(models):
    left = Inches(0.5 + i * 4.2)
    add_rounded(s, left, Inches(1.5), Inches(3.95), Inches(4.2), WHITE)
    add_rect(s, left, Inches(1.5), Inches(3.95), Inches(0.7), color)
    tb(s, left + Inches(0.2), Inches(1.6), Inches(3.5), Inches(0.5),
       [(abbr, 26, True, WHITE)], align=PP_ALIGN.CENTER)
    tb(s, left + Inches(0.25), Inches(2.45), Inches(3.5), Inches(0.4),
       [(name, 16, True, NAVY)], align=PP_ALIGN.CENTER)
    lines = desc.split("\n")
    tb(s, left + Inches(0.3), Inches(3.0), Inches(3.4), Inches(1.2),
       [(ln, 13, False, MUTED) for ln in lines], align=PP_ALIGN.CENTER)
    tb(s, left + Inches(0.3), Inches(4.5), Inches(3.4), Inches(0.9),
       [("Examples", 11, True, color), (examples, 12, False, INK)], align=PP_ALIGN.CENTER)


# ── 6. Real-life service model picker ──────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Real-Life Architecture Choices", "How top products map to IaaS / PaaS / FaaS / Containers")
footer(s, next_page())
cases = [
    ("Netflix", "IaaS + Containers", "Chaos engineering on VMs/K8s; global CDN (Open Connect) — need full control of media pipelines."),
    ("Shopify", "PaaS / Containers", "Merchant storefronts surge on Black Friday → auto-scale platforms beat hand-tuned VMs."),
    ("Duolingo", "Functions + K8s", "Lesson XP events & notifications via event-driven functions; core APIs on containers."),
    ("Zomato / Swiggy", "Microservices", "Order, menu, delivery tracking as separate services behind LBs — Fancy Store pattern at city scale."),
]
for i, (name, stack, why) in enumerate(cases):
    top = Inches(1.4 + i * 1.3)
    add_rounded(s, Inches(0.5), top, Inches(12.3), Inches(1.15), WHITE)
    add_rect(s, Inches(0.5), top, Inches(2.4), Inches(1.15), NAVY)
    tb(s, Inches(0.65), top + Inches(0.25), Inches(2.1), Inches(0.35), [(name, 14, True, WHITE)])
    tb(s, Inches(0.65), top + Inches(0.65), Inches(2.1), Inches(0.35), [(stack, 11, False, SKY_LIGHT)])
    tb(s, Inches(3.2), top + Inches(0.35), Inches(9.3), Inches(0.6), [(why, 13, False, MUTED)])


# ── 7. Topic 2 — Compute Engine ────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Topic 2 · Compute Engine — Scalable Web Apps", "IaaS foundation: VMs, Cloud Shell, storage, firewalls")
footer(s, next_page())
card(s, Inches(0.5), Inches(1.4), Inches(4.0), Inches(4.5), "Core building blocks", [
    "• Regions & zones for locality",
    "• Compute Engine API enabled",
    "• Cloud Storage buckets for code",
    "• Startup scripts automate boot:",
    "    install runtime → pull code",
    "    → Supervisor → non-root user",
    "• Network tags + firewall rules",
    "• gcloud CLI / Cloud Shell",
], SKY)
card(s, Inches(4.7), Inches(1.4), Inches(4.0), Inches(4.5), "Fancy Store lab pattern", [
    "Monolith → microservices:",
    "• Frontend  :8080",
    "• Orders    :8081",
    "• Products  :8082",
    "",
    "Code lives in GCS; VMs pull at",
    "startup (stateless image).",
    "Firewall tags: frontend / backend",
], TEAL)
card(s, Inches(8.9), Inches(1.4), Inches(4.0), Inches(4.5), "Real-life parallels", [
    "• Banks lift legacy Java apps",
    "  onto GCE before containerizing",
    "• Media encode farms use custom",
    "  machine types / GPUs",
    "• “Pets → Cattle”: identical VMs",
    "  from templates, not unique pets",
    "• Same idea as AWS EC2 + S3",
    "  user-data scripts",
], CORAL)


# ── 8. Topic 3 opener ──────────────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Topic 3 · Beyond a Single VM", "MIG · Load Balancers · Autohealing · CDN — the heart of cloud networking")
footer(s, next_page())
problems = [
    ("Single point of failure", "Managed Instance Groups"),
    ("No load distribution", "HTTP(S) Load Balancer"),
    ("Manual scaling", "Autoscaling policies"),
    ("No auto recovery", "Health checks + autohealing"),
    ("Slow global media", "Cloud CDN at the edge"),
    ("Risky deploy nights", "Rolling updates / replace"),
]
for i, (prob, sol) in enumerate(problems):
    col = i % 3
    row = i // 3
    left = Inches(0.5 + col * 4.2)
    top = Inches(1.5 + row * 2.4)
    add_rounded(s, left, top, Inches(3.95), Inches(2.1), WHITE)
    tb(s, left + Inches(0.25), top + Inches(0.35), Inches(3.5), Inches(0.5),
       [("Problem", 11, True, CORAL), (prob, 15, True, NAVY)])
    tb(s, left + Inches(0.25), top + Inches(1.2), Inches(3.5), Inches(0.6),
       [("Solution", 11, True, TEAL), (sol, 15, True, NAVY)])


# ── 9. MIG & Templates ─────────────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Instance Templates & Managed Instance Groups", "Blueprints + self-healing fleets")
footer(s, next_page())
card(s, Inches(0.5), Inches(1.4), Inches(6.0), Inches(3.8), "Instance Template", [
    "Immutable blueprint: machine type, disk,",
    "network, startup script, tags.",
    "",
    "Change? Create a NEW template, then roll it.",
    "Create from an existing VM (stop first for",
    "consistent disk state).",
], SKY)
card(s, Inches(6.8), Inches(1.4), Inches(6.0), Inches(3.8), "Managed Instance Group (MIG)", [
    "Identical VMs managed as one entity.",
    "• High availability & autohealing",
    "• Load balancing ready",
    "• Autoscaling & rolling updates",
    "• Named ports (frontend:8080, …)",
    "",
    "MIG (managed) ≫ unmanaged ad-hoc groups",
    "for production.",
], TEAL)
example_banner(s, Inches(5.5),
               "Flipkart / Amazon Big Billion / Prime Day: fleets grow with traffic and shrink overnight — MIG-style autoscaling is how flash sales stay online.")


# ── 10. Load Balancer architecture ─────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "HTTP(S) Load Balancer Architecture", "Global Layer-7 traffic → healthy backends")
footer(s, next_page())
steps = [
    ("1", "Forwarding Rule", "Public IP + port 80/443"),
    ("2", "Target HTTP Proxy", "Terminates HTTP, reads URL map"),
    ("3", "URL Map", "/ → FE · /api/orders · /api/products"),
    ("4", "Backend Services", "Health checks + capacity"),
    ("5", "MIG Instances", "Actual app VMs / Pods"),
]
for i, (n, title, desc) in enumerate(steps):
    left = Inches(0.4 + i * 2.55)
    add_rounded(s, left, Inches(1.7), Inches(2.4), Inches(2.8), WHITE)
    circ = s.shapes.add_shape(MSO_SHAPE.OVAL, left + Inches(0.85), Inches(1.95), Inches(0.7), Inches(0.7))
    circ.fill.solid()
    circ.fill.fore_color.rgb = SKY if i % 2 == 0 else TEAL
    circ.line.fill.background()
    tb(s, left + Inches(0.85), Inches(2.05), Inches(0.7), Inches(0.5),
       [(n, 18, True, WHITE)], align=PP_ALIGN.CENTER)
    tb(s, left + Inches(0.1), Inches(2.9), Inches(2.2), Inches(0.7),
       [(title, 13, True, NAVY)], align=PP_ALIGN.CENTER)
    tb(s, left + Inches(0.1), Inches(3.6), Inches(2.2), Inches(0.7),
       [(desc, 11, False, MUTED)], align=PP_ALIGN.CENTER)
example_banner(s, Inches(5.0),
               "Google Search / YouTube Front Ends: one anycast IP worldwide — same idea as GCP’s global HTTP(S) LB routing users to the nearest healthy backend.")
tb(s, Inches(0.5), Inches(5.8), Inches(12.3), Inches(0.8),
   [("Pro tip: lab often uses HTTP for simplicity — production must terminate HTTPS and protect health-check IP ranges.", 13, False, INK)])


# ── 11. CDN + Autohealing ──────────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Autohealing, Rolling Updates & Cloud CDN", "Recover, release, and accelerate")
footer(s, next_page())
card(s, Inches(0.5), Inches(1.4), Inches(4.0), Inches(4.5), "Health & Autohealing", [
    "Health checks probe / or /api/*",
    "Unhealthy? MIG recreates the VM.",
    "",
    "Separate LB health checks from",
    "autohealing checks when needed.",
    "",
    "Lab: supervisorctl stop → watch",
    "healing recreate the instance.",
], CORAL)
card(s, Inches(4.7), Inches(1.4), Inches(4.0), Inches(4.5), "Rolling Updates", [
    "maxSurge / maxUnavailable control",
    "how many new vs old VMs exist.",
    "",
    "Keeps minimum capacity healthy.",
    "Safer than “stop all → redeploy”.",
    "",
    "Also: rolling replace after you",
    "upload new code to GCS.",
], SKY)
card(s, Inches(8.9), Inches(1.4), Inches(4.0), Inches(4.5), "Cloud CDN", [
    "Cache static assets at Google’s",
    "edge — images, JS, CSS, video.",
    "",
    "Enable on frontend backend service.",
    "Cuts origin load & latency.",
    "",
    "Twitch / Disney+ / Cloudflare-class",
    "patterns: cache near the viewer.",
], TEAL)


# ── 12. Topic 4 Flex ───────────────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Topic 4 · App Engine Flexible & Exploring PaaS", "Docker-powered PaaS when sandbox Standard is too tight")
footer(s, next_page())
card(s, Inches(0.5), Inches(1.4), Inches(6.0), Inches(3.6), "When Flexible wins", [
    "• Custom OS packages / native libs",
    "• Long requests (up to ~60 minutes)",
    "• SSH into underlying VMs for debug",
    "• Full VPC integration",
    "• Pay for vCPU + memory used",
    "",
    "Region choice is permanent — pick close",
    "to users on day one.",
], SKY)
card(s, Inches(6.8), Inches(1.4), Inches(6.0), Inches(3.6), "PaaS in practice", [
    "Languages: Java, Python, Go, PHP,",
    "Node.js, Ruby… (modern runtimes).",
    "",
    "Workflow: local SDK → gcloud app deploy",
    "→ Google scales & load balances.",
    "",
    "Data services: Cloud SQL, Firestore,",
    "Cloud Storage, Memcache.",
], TEAL)
example_banner(s, Inches(5.3),
               "Khan Academy used App Engine to teach millions with a small platform team — PaaS trades control for shipping speed, which is perfect for education & content sites.")


# ── 13. Topic 5 Functions ──────────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Topic 5 · Event-Driven Cloud Functions", "Serverless FaaS — pay per invocation, scale 0 → N")
footer(s, next_page())

flow = [
    ("Event", "Upload / HTTP /\nPub/Sub / Firestore"),
    ("Trigger", "Cloud Function\nwakes automatically"),
    ("Execute", "Run business logic\nreturn or fan-out"),
]
for i, (t, d) in enumerate(flow):
    left = Inches(0.6 + i * 4.2)
    add_rounded(s, left, Inches(1.45), Inches(3.8), Inches(1.8), WHITE)
    add_rect(s, left, Inches(1.45), Inches(3.8), Inches(0.45), [SKY, TEAL, CORAL][i])
    tb(s, left + Inches(0.2), Inches(1.5), Inches(3.4), Inches(0.35),
       [(t, 14, True, WHITE)], align=PP_ALIGN.CENTER)
    tb(s, left + Inches(0.2), Inches(2.1), Inches(3.4), Inches(0.9),
       [(d, 13, False, MUTED)], align=PP_ALIGN.CENTER)

card(s, Inches(0.5), Inches(3.5), Inches(6.0), Inches(2.4), "2nd gen (recommended)", [
    "Built on Cloud Run • longer timeouts • more concurrency",
    "CloudEvents standard across runtimes",
    "Triggers: HTTP, Storage, Pub/Sub, Firestore, Audit logs",
    "Watch cold starts — keep functions small; min instances help",
], SKY)
card(s, Inches(6.8), Inches(3.5), Inches(6.0), Inches(2.4), "Real-life triggers", [
    "• Instagram/WhatsApp-style: resize image on Storage upload",
    "• UPI banks: Pub/Sub fraud-scoring function on each txn",
    "• IoT fleet: telemetry → function → BigQuery",
    "• Webhooks: Stripe / Razorpay payment events",
], CORAL)


# ── 14. Compare compute options ────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Choose Your Compute: Functions vs App Engine vs GCE", "One of the most asked exam & interview questions")
footer(s, next_page())

# Table-like cards
headers = ["Aspect", "Cloud Functions", "App Engine", "Compute Engine"]
rows = [
    ["Model", "FaaS / Serverless", "PaaS", "IaaS"],
    ["Deploy unit", "Single function", "Application", "Virtual machine"],
    ["Scaling", "0 → N auto", "Auto", "Manual / MIG"],
    ["Billing", "Per invocation", "Instance-hours", "VM-hours"],
    ["Best for", "Events & glue code", "Web apps & APIs", "Custom / legacy"],
]
# Header row
for i, h in enumerate(headers):
    left = Inches(0.4 + i * 3.2)
    add_rect(s, left, Inches(1.5), Inches(3.05), Inches(0.55), NAVY)
    tb(s, left + Inches(0.1), Inches(1.58), Inches(2.85), Inches(0.4),
       [(h, 13, True, WHITE)], align=PP_ALIGN.CENTER)
for r, row in enumerate(rows):
    for c, cell in enumerate(row):
        left = Inches(0.4 + c * 3.2)
        top = Inches(2.1 + r * 0.7)
        bg = WHITE if r % 2 == 0 else SOFT
        add_rect(s, left, top, Inches(3.05), Inches(0.65), bg)
        color = NAVY if c == 0 else INK
        bold = c == 0
        tb(s, left + Inches(0.1), top + Inches(0.15), Inches(2.85), Inches(0.4),
           [(cell, 12, bold, color)], align=PP_ALIGN.CENTER)
tb(s, Inches(0.5), Inches(5.8), Inches(12.3), Inches(0.8),
   [("Rule of thumb: start managed (Functions / App Engine / Cloud Run). Drop to GCE/GKE only when you need control that PaaS cannot give.", 13, True, TEAL)])


# ── 15. Topic 6 GKE ────────────────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Topic 6 · Google Kubernetes Engine (GKE)", "Managed Kubernetes born from Google’s Borg")
footer(s, next_page())
card(s, Inches(0.5), Inches(1.4), Inches(6.0), Inches(3.5), "When to use GKE", [
    "• Multi-container apps & microservices",
    "• Portable workloads (cloud ↔ on-prem)",
    "• Service discovery + advanced networking",
    "• CI/CD of container images at scale",
    "",
    "Prefer App Engine/Cloud Run for simple web;",
    "Functions for events; GCE for legacy VMs.",
], SKY)
card(s, Inches(6.8), Inches(1.4), Inches(6.0), Inches(3.5), "Architecture essentials", [
    "Control plane (Google-managed): API server,",
    "scheduler, controllers, etcd",
    "",
    "Nodes = GCE VMs with kubelet",
    "Pod = smallest deployable unit",
    "Deployment = replicas + rolling updates",
    "Service = stable network endpoint",
], TEAL)
example_banner(s, Inches(5.2),
               "Pokémon GO famously rode Kubernetes during launch spikes; Spotify & Airbnb run large microservice meshes on K8s for portability and density.")


# ── 16. Autopilot vs Standard ──────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "GKE Modes — Autopilot vs Standard", "Choose ops burden vs knobs")
footer(s, next_page())
card(s, Inches(0.5), Inches(1.45), Inches(6.0), Inches(4.5), "Autopilot (recommended default)", [
    "• Google manages nodes AND control plane",
    "• Pay only for Pod CPU/memory requests",
    "• Hardened defaults + node auto-repair",
    "• Pod-level SLA, less ops work",
    "• Great for most production microservices",
    "",
    "gcloud container clusters create-auto …",
], TEAL)
card(s, Inches(6.8), Inches(1.45), Inches(6.0), Inches(4.5), "Standard", [
    "• You size & tune nodes",
    "• Pay for whole node (even idle)",
    "• GPUs, custom machines, taints…",
    "• Node auto-provisioning optional",
    "• Release channels: Rapid / Regular / Stable",
    "",
    "Pick when you need exotic hardware or",
    "cluster-level customization.",
], SKY)


# ── 17. Topic 7 Deploy on GKE ──────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Topic 7 · Deploying Containers on GKE", "Artifact Registry → Cluster → Deployment → HPA → Service")
footer(s, next_page())
pipeline = [
    ("1. Build", "docker build → image"),
    ("2. Store", "Artifact Registry"),
    ("3. Cluster", "Autopilot / Standard"),
    ("4. Deploy", "kubectl Deployment"),
    ("5. Scale", "HPA on CPU %"),
    ("6. Expose", "Service / LB / Ingress"),
]
for i, (t, d) in enumerate(pipeline):
    col = i % 3
    row = i // 3
    left = Inches(0.5 + col * 4.2)
    top = Inches(1.5 + row * 2.2)
    add_rounded(s, left, top, Inches(3.95), Inches(1.9), WHITE)
    add_rect(s, left, top, Inches(0.12), Inches(1.9), [SKY, TEAL, CORAL, SKY, TEAL, CORAL][i])
    tb(s, left + Inches(0.35), top + Inches(0.4), Inches(3.4), Inches(0.45), [(t, 18, True, NAVY)])
    tb(s, left + Inches(0.35), top + Inches(1.0), Inches(3.4), Inches(0.5), [(d, 14, False, MUTED)])


# ── 18. Topic 8 Principles ─────────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Topic 8 · Cloud Principles & Well-Architected", "NIST fundamentals + Google’s design pillars")
footer(s, next_page())
pillars = [
    ("On-demand", "Self-service provisioning"),
    ("Broad access", "Any device over network"),
    ("Pooling", "Multi-tenant efficiency"),
    ("Elasticity", "Scale up/down fast"),
    ("Measured", "Utility pay-as-you-go"),
]
for i, (t, d) in enumerate(pillars):
    left = Inches(0.4 + i * 2.55)
    add_rounded(s, left, Inches(1.45), Inches(2.4), Inches(1.5), WHITE)
    tb(s, left + Inches(0.1), Inches(1.65), Inches(2.2), Inches(0.5),
       [(t, 14, True, SKY)], align=PP_ALIGN.CENTER)
    tb(s, left + Inches(0.1), Inches(2.2), Inches(2.2), Inches(0.5),
       [(d, 11, False, MUTED)], align=PP_ALIGN.CENTER)

wa = [
    ("Operational Excellence", "Automate, monitor, learn from failure"),
    ("Security", "IAM, least privilege, encrypt in transit"),
    ("Reliability", "HA, autohealing, assume pods die"),
    ("Performance", "Right-size, HPA, CDN, Memcache"),
    ("Cost", "Delete idle, Spot Pods, Autopilot pay-per-pod"),
    ("Sustainability", "No idle waste; serverless when possible"),
]
for i, (t, d) in enumerate(wa):
    col = i % 3
    row = i // 3
    left = Inches(0.5 + col * 4.2)
    top = Inches(3.3 + row * 1.55)
    add_rounded(s, left, top, Inches(3.95), Inches(1.35), WHITE)
    add_rect(s, left, top, Inches(3.95), Inches(0.1), [SKY, TEAL, CORAL, GOLD, SKY, TEAL][i])
    tb(s, left + Inches(0.2), top + Inches(0.25), Inches(3.5), Inches(0.35), [(t, 13, True, NAVY)])
    tb(s, left + Inches(0.2), top + Inches(0.7), Inches(3.5), Inches(0.45), [(d, 12, False, MUTED)])


# ── 19. Shared responsibility + real life ─────────────────────────────────
s = blank_content_slide()
header_bar(s, "Shared Responsibility — Who Owns What?", "Cloud ≠ “Google secures everything”")
footer(s, next_page())
rows2 = [
    ("Always yours", "App code, data classification, IAM identity design, secrets"),
    ("GAE / Functions", "You: app + data · Google: runtime, OS, infra"),
    ("Compute Engine", "You: OS patches, apps, network config · Google: hardware"),
    ("GKE", "Google: control plane · You: workloads (nodes shared/Autopilot)"),
]
for i, (t, d) in enumerate(rows2):
    top = Inches(1.45 + i * 1.1)
    add_rounded(s, Inches(0.5), top, Inches(12.3), Inches(0.95), WHITE)
    add_rect(s, Inches(0.5), top, Inches(3.0), Inches(0.95), NAVY)
    tb(s, Inches(0.7), top + Inches(0.3), Inches(2.6), Inches(0.4), [(t, 13, True, WHITE)])
    tb(s, Inches(3.8), top + Inches(0.3), Inches(8.7), Inches(0.45), [(d, 13, False, MUTED)])


# ── 20. Topic 9 DR ─────────────────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Topic 9 · Disaster Recovery in the Cloud", "Backup ≠ DR — define RTO & RPO first")
footer(s, next_page())
card(s, Inches(0.5), Inches(1.4), Inches(6.0), Inches(2.6), "Key metrics", [
    "RTO — Recovery Time Objective: max downtime",
    "RPO — Recovery Point Objective: max data loss window",
    "",
    "Lower RTO/RPO ⇒ more $ (hot standby / multi-site)",
], CORAL)
card(s, Inches(6.8), Inches(1.4), Inches(6.0), Inches(2.6), "Strategy ladder", [
    "Backup & Restore → Pilot Light → Warm Standby",
    "→ Hot Standby → Multi-site Active-Active",
    "",
    "HA patterns: MIG, replicas, LB, multi-zone/region",
], SKY)
example_banner(s, Inches(4.3),
               "Airlines & UPI apps target near-zero RTO with active-active regions. A campus portal may accept hours (cheap backup/restore). Match strategy to business pain.")
tb(s, Inches(0.5), Inches(5.2), Inches(12.3), Inches(1.3),
   [
       ("GCP toolkit from the unit:", 14, True, NAVY),
       ("Cloud Storage versioning · GCE disk snapshots · Cloud SQL automated backups · multi-zone MIG/GKE · global LB failover", 13, False, MUTED),
   ])


# ── 21. Topic 10 IaC ───────────────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Topic 10 · Infrastructure as Code (IaC)", "Machines from files — repeatable, reviewable, recoverable")
footer(s, next_page())
card(s, Inches(0.5), Inches(1.4), Inches(4.0), Inches(4.5), "Why IaC", [
    "• Repeatability",
    "• Git history = audit trail",
    "• Fewer console typos",
    "• Docs that can’t rot",
    "• Rebuild after disaster",
    "",
    "Imperative: gcloud create…",
    "Declarative: desired state=2",
], SKY)
card(s, Inches(4.7), Inches(1.4), Inches(4.0), Inches(4.5), "Tool map", [
    "Terraform — multi-cloud HCL",
    "Deployment Manager — GCP YAML",
    "CloudFormation — AWS",
    "Ansible — config over SSH",
    "Pulumi — Python/TS/Go IaC",
    "",
    "Templates & MIG size = declarative",
    "mindset you already practiced.",
], TEAL)
card(s, Inches(8.9), Inches(1.4), Inches(4.0), Inches(4.5), "Real-life", [
    "• HashiCorp Terraform at banks",
    "  for regulated change control",
    "• Spotify Backstage + IaC for",
    "  self-service environments",
    "• Your Fancy Store firewall →",
    "  google_compute_firewall{}",
    "• Never commit secrets — use",
    "  Secret Manager",
], CORAL)


# ── 22. Topic 11 CI/CD ─────────────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Topic 11 · CI/CD & DevOps", "From git push to verified production — automatically")
footer(s, next_page())
stages = ["Code", "Build", "Test", "Artifact", "Deploy", "Verify"]
for i, stg in enumerate(stages):
    left = Inches(0.45 + i * 2.15)
    add_rounded(s, left, Inches(1.45), Inches(2.0), Inches(0.85), NAVY if i % 2 == 0 else SKY)
    tb(s, left, Inches(1.6), Inches(2.0), Inches(0.55),
       [(stg, 14, True, WHITE)], align=PP_ALIGN.CENTER)

card(s, Inches(0.5), Inches(2.6), Inches(4.0), Inches(3.3), "Concepts", [
    "CI — integrate often; build+test every push",
    "Continuous Delivery — always deployable",
    "Continuous Deployment — auto to prod",
    "",
    "Tools: Cloud Build, Cloud Deploy,",
    "GitHub Actions, GitLab CI, Jenkins",
], SKY)
card(s, Inches(4.7), Inches(2.6), Inches(4.0), Inches(3.3), "Strategies", [
    "Rolling — MIG / GKE default",
    "Blue-Green — swap environments",
    "Canary — % traffic to new version",
    "Recreate — simple but downtime",
    "",
    "Health checks gate traffic;",
    "rollback on failure.",
], TEAL)
card(s, Inches(8.9), Inches(2.6), Inches(4.0), Inches(3.3), "Real-life pipelines", [
    "• Google ships internally with",
    "  huge automated test matrices",
    "• GitHub (Actions) powering OSS",
    "• Indian unicorns: canary on",
    "  payments before 100% cutover",
    "• Lab map: gsutil + rolling",
    "  replace = CD in miniature",
], CORAL)


# ── 23. Fancy Store architecture wrap ──────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Capstone Pattern — “Fancy Store” Networking Blueprint", "Unit labs ≈ how modern e-commerce fleets are wired")
footer(s, next_page())
arch = [
    ("Users", "Browsers worldwide"),
    ("Global LB + CDN", "URL map to services"),
    ("Frontend MIG", "Port 8080 · cached assets"),
    ("Orders / Products MIG", "8081 / 8082 APIs"),
    ("GCS + Templates", "Stateless code delivery"),
    ("Health + Autoscale", "Heal, roll, grow, shrink"),
]
for i, (t, d) in enumerate(arch):
    col = i % 3
    row = i // 3
    left = Inches(0.5 + col * 4.2)
    top = Inches(1.5 + row * 2.3)
    add_rounded(s, left, top, Inches(3.95), Inches(2.0), WHITE)
    add_rect(s, left, top, Inches(3.95), Inches(0.55), [NAVY, SKY, TEAL, CORAL, NAVY, SKY][i])
    tb(s, left + Inches(0.2), top + Inches(0.12), Inches(3.5), Inches(0.4),
       [(t, 14, True, WHITE)], align=PP_ALIGN.CENTER)
    tb(s, left + Inches(0.2), top + Inches(0.9), Inches(3.5), Inches(0.7),
       [(d, 14, False, MUTED)], align=PP_ALIGN.CENTER)


# ── 24. Extra real-life gallery ────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "More Real-Life Cloud Stories (Beyond the Notes)", "Connect theory to products students already use")
footer(s, next_page())
stories = [
    ("Hotstar / Disney+", "IPL finals: traffic multiplies 50–100×. Autoscale + CDN + multi-region = stay up during last-over sixes."),
    ("PhonePe / GPay", "Payment spikes on salary day. Event pipelines + multi-zone HA + aggressive RTO for money movement."),
    ("Zoom (COVID peak)", "Elastic capacity overnight. Measured service + rapid elasticity saved classrooms & offices."),
    ("Spotify Wrapped", "Batch + streaming analytics on containers; publish once a year but build on always-on data platforms."),
    ("IRCTC window", "Ticket rush = classic thundering herd. Queues, caching, rate limits, and graceful degradation."),
    ("Campus ERP", "Exam result day mimics cloud black Friday — plan HPA/MIG before the PDF drop."),
]
for i, (t, d) in enumerate(stories):
    col = i % 2
    row = i // 2
    left = Inches(0.45 + col * 6.4)
    top = Inches(1.4 + row * 1.7)
    add_rounded(s, left, top, Inches(6.15), Inches(1.5), WHITE)
    add_rect(s, left, top, Inches(0.12), Inches(1.5), [SKY, TEAL, CORAL, GOLD, SKY, TEAL][i])
    tb(s, left + Inches(0.35), top + Inches(0.2), Inches(5.6), Inches(0.35), [(t, 14, True, NAVY)])
    tb(s, left + Inches(0.35), top + Inches(0.65), Inches(5.6), Inches(0.65), [(d, 12, False, MUTED)])


# ── 25. Exam rapid revision ────────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Exam Rapid Revision — Must Remember", "High-frequency concepts from Unit II")
footer(s, next_page())
bullets = [
    "GAE = PaaS; GCE = IaaS; Cloud Functions = FaaS; GKE = managed Kubernetes (from Borg).",
    "Instance templates are immutable; MIGs deliver HA, autoscaling, rolling updates, autohealing.",
    "HTTP(S) LB chain: Forwarding rule → Target proxy → URL map → Backend service → MIG.",
    "Named ports tell the LB which port to hit (frontend:8080, orders:8081, products:8082).",
    "Autopilot pays per Pod resources; Standard pays for nodes — pick based on control needs.",
    "RTO = downtime budget; RPO = data-loss budget — drive DR tier selection.",
    "IaC: Terraform/DM/Ansible; startup scripts & MIG size already teach the mindset.",
    "CI vs CD vs Continuous Deployment; rolling / blue-green / canary for safe releases.",
]
for i, b in enumerate(bullets):
    top = Inches(1.35 + i * 0.65)
    add_rect(s, Inches(0.5), top, Inches(0.18), Inches(0.18), TEAL if i % 2 == 0 else SKY)
    tb(s, Inches(0.9), top - Inches(0.05), Inches(11.8), Inches(0.55), [(b, 13, False, INK)])


# ── 26. Key takeaways ──────────────────────────────────────────────────────
s = blank_content_slide()
header_bar(s, "Key Takeaways", "What Unit II wants you to design for")
footer(s, next_page())
takes = [
    ("Design for failure", "Instances and Pods die. Health checks, MIG/GKE replicas, and multi-zone beat hero servers."),
    ("Automate the path", "Startup scripts → templates → IaC → CI/CD. Humans review; machines apply."),
    ("Pick abstractions wisely", "Don’t run Kubernetes for a brochure site. Don’t trap a banking core on a toy PaaS sandbox."),
    ("Measure & cache", "Autoscaling needs signals; CDN and Memcache protect origins when the world shows up."),
]
for i, (t, d) in enumerate(takes):
    top = Inches(1.45 + i * 1.25)
    add_rounded(s, Inches(0.5), top, Inches(12.3), Inches(1.1), WHITE)
    add_rect(s, Inches(0.5), top, Inches(0.12), Inches(1.1), [SKY, TEAL, CORAL, GOLD][i])
    tb(s, Inches(0.9), top + Inches(0.2), Inches(11.5), Inches(0.35), [(t, 16, True, NAVY)])
    tb(s, Inches(0.9), top + Inches(0.55), Inches(11.5), Inches(0.4), [(d, 13, False, MUTED)])


# ── 27. Closing ────────────────────────────────────────────────────────────
s = blank_content_slide()
add_rect(s, Inches(0), Inches(0), SLIDE_W, SLIDE_H, NAVY)
add_rect(s, Inches(0), Inches(0), SLIDE_W, Inches(0.15), TEAL)
tb(s, Inches(0.8), Inches(2.2), Inches(11.5), Inches(0.8),
   [("Thank you", 44, True, WHITE)])
tb(s, Inches(0.8), Inches(3.2), Inches(11.5), Inches(0.5),
   [("Questions · Discussion · Lab mapping", 20, False, SKY_LIGHT)])
tb(s, Inches(0.8), Inches(4.2), Inches(11.5), Inches(1.0),
   [
       ("Source: CC-II Unit II notes (Networking in the Cloud)", 14, False, SOFT),
       ("Enriched with industry case studies for classroom discussion", 14, False, SOFT),
       ("Branch: cc-ii-ppt  ·  FIT & CS, Parul University", 14, False, MUTED),
   ])
next_page()


# ── Save ───────────────────────────────────────────────────────────────────
out = os.path.join(os.path.dirname(__file__), "CC-II-Unit-II-Networking-in-the-Cloud.pptx")
prs.save(out)
print(f"Saved {out}")
print(f"Slides: {len(prs.slides)}")
