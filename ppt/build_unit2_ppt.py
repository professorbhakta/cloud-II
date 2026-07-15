#!/usr/bin/env python3
"""Unit II PowerPoint — Disaster Recovery in Cloud (syllabus-aligned)."""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
import os

# Cloud theme (navy / sky — DR-focused)
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
GREEN = RGBColor(0x1F, 0x8A, 0x5F)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H


def add_rect(slide, left, top, width, height, fill):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    return shape


def add_rounded(slide, left, top, width, height, fill):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
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
       align=PP_ALIGN.LEFT):
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
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.space_after = Pt(4)
        run = p.add_run()
        set_run(run, text, size=size, bold=bold, color=color)
    return box


def footer(slide, page):
    add_rect(slide, Inches(0), Inches(7.15), SLIDE_W, Inches(0.35), NAVY)
    tb(slide, Inches(0.4), Inches(7.18), Inches(10), Inches(0.3),
       [("CC-II  ·  Unit II  ·  Disaster Recovery in Cloud  ·  Parul University", 10, False, SOFT)])
    tb(slide, Inches(11.5), Inches(7.18), Inches(1.5), Inches(0.3),
       [(str(page), 10, False, SOFT)], align=PP_ALIGN.RIGHT)


def header_bar(slide, title, subtitle=None):
    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.15), NAVY)
    add_rect(slide, Inches(0), Inches(1.15), SLIDE_W, Inches(0.08), TEAL)
    tb(slide, Inches(0.5), Inches(0.22), Inches(12), Inches(0.5),
       [(title, 26, True, WHITE)])
    if subtitle:
        tb(slide, Inches(0.5), Inches(0.7), Inches(12), Inches(0.35),
           [(subtitle, 13, False, SKY_LIGHT)])


def blank():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, Inches(0), Inches(0), SLIDE_W, SLIDE_H, CLOUD)
    return slide


def card(slide, left, top, width, height, title, body_lines, accent=SKY):
    add_rounded(slide, left, top, width, height, WHITE)
    add_rect(slide, left, top, Inches(0.1), height, accent)
    tb(slide, left + Inches(0.25), top + Inches(0.15), width - Inches(0.4), Inches(0.4),
       [(title, 15, True, NAVY)])
    formatted = [(ln, 12, False, MUTED) if isinstance(ln, str) else ln for ln in body_lines]
    tb(slide, left + Inches(0.25), top + Inches(0.55), width - Inches(0.4), height - Inches(0.7),
       formatted, default_size=12, default_color=MUTED)


def example_banner(slide, top, text):
    add_rounded(slide, Inches(0.5), top, Inches(12.3), Inches(0.55), RGBColor(0xFF, 0xF4, 0xEC))
    tb(slide, Inches(0.7), top + Inches(0.1), Inches(12), Inches(0.4),
       [(f"Real life:  {text}", 13, True, CORAL)])


PAGE = 0


def p():
    global PAGE
    PAGE += 1
    return PAGE


# ═══════════════════════════════════════════════════════════════════════════
# 1. TITLE
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
add_rect(s, Inches(0), Inches(0), SLIDE_W, SLIDE_H, NAVY)
add_rect(s, Inches(0), Inches(0), SLIDE_W, Inches(0.15), TEAL)
add_rect(s, Inches(0), Inches(7.35), SLIDE_W, Inches(0.15), SKY)
for left, top, w, h in [(10.2, 0.8, 2.8, 1.2), (0.3, 5.8, 3.2, 1.0), (9.5, 5.5, 3.5, 1.3)]:
    sh = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(left), Inches(top), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = RGBColor(0x14, 0x2E, 0x4A)
    sh.line.fill.background()

tb(s, Inches(0.8), Inches(1.7), Inches(11.5), Inches(0.4),
   [("CLOUD COMPUTING – II  ·  UNIT II", 15, True, TEAL)])
tb(s, Inches(0.8), Inches(2.25), Inches(11.5), Inches(1.0),
   [("Disaster Recovery", 44, True, WHITE),
    ("in the Cloud", 44, True, WHITE)])
tb(s, Inches(0.8), Inches(4.3), Inches(11.5), Inches(1.0),
   [("Backup & DR Concepts  ·  High Availability & Fault Tolerance", 14, False, SOFT),
    ("RPO & RTO Strategies  ·  Cloud Backups & Snapshots  ·  Case Studies", 14, False, SOFT)])
tb(s, Inches(0.8), Inches(5.8), Inches(11.5), Inches(0.4),
   [("FIT & CS  ·  Parul University", 14, False, MUTED)])
p()


# ═══════════════════════════════════════════════════════════════════════════
# 2. AGENDA — exact syllabus topics
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Unit II Syllabus — Agenda", "Disaster Recovery in Cloud")
footer(s, p())
topics = [
    ("01", "Disaster Recovery in Cloud", "Why DR matters; what a cloud disaster looks like"),
    ("02", "Backup & DR Concepts", "Backup, disaster, DR, business continuity"),
    ("03", "HA & Fault Tolerance", "Redundancy, SPOF, multi-zone / multi-region"),
    ("04", "DR Strategies · RPO & RTO", "Backup-restore → pilot → warm → hot → active-active"),
    ("05", "Cloud Backups & Snapshots", "GCE snapshots, GCS, Cloud SQL, GKE state"),
    ("06", "Case Studies", "Lab + industry DR implementations"),
]
for i, (num, title, sub) in enumerate(topics):
    col = i % 3
    row = i // 3
    left = Inches(0.5 + col * 4.2)
    top = Inches(1.55 + row * 2.4)
    add_rounded(s, left, top, Inches(3.95), Inches(2.1), WHITE)
    add_rect(s, left, top, Inches(3.95), Inches(0.12), [SKY, TEAL, CORAL, SKY, TEAL, CORAL][i])
    tb(s, left + Inches(0.25), top + Inches(0.4), Inches(3.4), Inches(0.4),
       [(num, 22, True, [SKY, TEAL, CORAL, SKY, TEAL, CORAL][i])])
    tb(s, left + Inches(0.25), top + Inches(0.95), Inches(3.4), Inches(0.4),
       [(title, 14, True, NAVY)])
    tb(s, left + Inches(0.25), top + Inches(1.4), Inches(3.4), Inches(0.5),
       [(sub, 12, False, MUTED)])


# ═══════════════════════════════════════════════════════════════════════════
# 3. What is Disaster Recovery in Cloud?
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "What is Disaster Recovery in Cloud?", "Policies + people + technology to restore operations after a disaster")
footer(s, p())

card(s, Inches(0.5), Inches(1.4), Inches(6.0), Inches(3.5), "In the cloud, a “disaster” can be…", [
    "• Regional outage (power, network, cooling)",
    "• Zone failure (single datacenter impact)",
    "• Cyberattack / ransomware encrypting data",
    "• Accidental delete or bad deploy",
    "• Natural disaster near a region",
    "• Vendor / dependency failure cascading out",
], CORAL)

card(s, Inches(6.8), Inches(1.4), Inches(6.0), Inches(3.5), "DR answers three questions", [
    "1. How fast must we be back?  → RTO",
    "2. How much data can we lose? → RPO",
    "3. How do we actually recover?",
    "     backups, replicas, failover, runbooks",
    "",
    "Cloud advantage: spare capacity on demand,",
    "snapshots, multi-region, IaC rebuilds.",
], SKY)

example_banner(s, Inches(5.2),
               "When AWS us-east-1 wobbles, half the internet feels it — DR is not optional for money, health, or exam-result systems.")


# ═══════════════════════════════════════════════════════════════════════════
# 4. Backup vs DR vs Business Continuity
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Concepts — Backup, Disaster, DR & Continuity", "These four words are related but not the same")
footer(s, p())

defs = [
    ("Backup", "A copy of data kept so you can restore after loss or corruption.", SKY),
    ("Disaster", "An event that makes systems unavailable — outage, flood, cyberattack.", CORAL),
    ("Disaster Recovery (DR)", "Policies & procedures to restore IT operations after a disaster.", TEAL),
    ("Business Continuity", "Keeping critical business services running for users / revenue.", GOLD),
]
for i, (t, d, c) in enumerate(defs):
    top = Inches(1.4 + i * 1.25)
    add_rounded(s, Inches(0.5), top, Inches(12.3), Inches(1.1), WHITE)
    add_rect(s, Inches(0.5), top, Inches(3.4), Inches(1.1), c if i < 3 else NAVY)
    tb(s, Inches(0.7), top + Inches(0.35), Inches(3.0), Inches(0.45), [(t, 14, True, WHITE)])
    tb(s, Inches(4.2), top + Inches(0.35), Inches(8.3), Inches(0.5), [(d, 14, False, MUTED)])


# ═══════════════════════════════════════════════════════════════════════════
# 5. Backup types
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Backup Types You Must Know", "Full · Incremental · Snapshot — trade space, time, and restore complexity")
footer(s, p())

types = [
    ("Full Backup", "Entire dataset every run",
     "Simple restore\nSlow + expensive\nOften weekly baseline", SKY),
    ("Incremental", "Only changes since last backup",
     "Fast & cheap to take\nRestore = full + chain\nCommon daily schedule", TEAL),
    ("Snapshot", "Point-in-time disk/image copy",
     "Near-instant capture\nIdeal for VMs & disks\nGCP: GCE / PV snapshots", CORAL),
]
for i, (t, sub, body, c) in enumerate(types):
    left = Inches(0.5 + i * 4.2)
    add_rounded(s, left, Inches(1.5), Inches(3.95), Inches(4.2), WHITE)
    add_rect(s, left, Inches(1.5), Inches(3.95), Inches(0.7), c)
    tb(s, left + Inches(0.2), Inches(1.6), Inches(3.5), Inches(0.5),
       [(t, 18, True, WHITE)], align=PP_ALIGN.CENTER)
    tb(s, left + Inches(0.25), Inches(2.5), Inches(3.5), Inches(0.5),
       [(sub, 13, True, NAVY)], align=PP_ALIGN.CENTER)
    lines = body.split("\n")
    tb(s, left + Inches(0.35), Inches(3.3), Inches(3.3), Inches(2.0),
       [(ln, 13, False, MUTED) for ln in lines], align=PP_ALIGN.CENTER)

example_banner(s, Inches(6.0),
               "PhonePe / banks: full weekly + frequent incrementals/snapshots so a 2 PM corruption can roll back near 1:30 PM (tight RPO).")


# ═══════════════════════════════════════════════════════════════════════════
# 6. Why backup alone is not enough
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Backup ≠ Disaster Recovery", "Having a zip file is not a recovery plan")
footer(s, p())

rows = [
    ("Only backups", "You still need people, runbooks, spare infra, tested restore time, and a decision of WHEN to fail over."),
    ("Untested restores", "Backups that never restore successfully are false comfort — practice game days."),
    ("Same-region copies", "A region outage can wipe primary AND local backups. Cross-region matters."),
    ("No RTO/RPO", "Without targets you overspend on gold-plated DR or underserve a critical app."),
]
for i, (t, d) in enumerate(rows):
    top = Inches(1.4 + i * 1.25)
    add_rounded(s, Inches(0.5), top, Inches(12.3), Inches(1.1), WHITE)
    add_rect(s, Inches(0.5), top, Inches(0.12), Inches(1.1), [CORAL, SKY, TEAL, GOLD][i])
    tb(s, Inches(0.9), top + Inches(0.2), Inches(11.5), Inches(0.35), [(t, 15, True, NAVY)])
    tb(s, Inches(0.9), top + Inches(0.55), Inches(11.5), Inches(0.4), [(d, 13, False, MUTED)])


# ═══════════════════════════════════════════════════════════════════════════
# 7. HA & Fault Tolerance definitions
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "High Availability & Fault Tolerance", "Two goals that often work together — but mean different things")
footer(s, p())

card(s, Inches(0.5), Inches(1.4), Inches(6.0), Inches(4.5), "High Availability (HA)", [
    "Minimize downtime through redundancy.",
    "",
    "Goal: service stays reachable for users.",
    "Often measured as uptime % (e.g. 99.9%).",
    "",
    "Cloud examples:",
    "• MIG with 2+ instances",
    "• Load balancer + health checks",
    "• Multi-zone deployments",
], SKY)

card(s, Inches(6.8), Inches(1.4), Inches(6.0), Inches(4.5), "Fault Tolerance (FT)", [
    "Continue operating despite a component",
    "failure — ideally without the user noticing.",
    "",
    "Cloud examples:",
    "• Autohealing recreates a bad VM",
    "• Replica Pods serve while one dies",
    "• Stateless apps pull code from GCS",
    "",
    "HA often uses FT building blocks.",
], TEAL)


# ═══════════════════════════════════════════════════════════════════════════
# 8. Key HA concepts table
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "HA Building Blocks", "Concept → meaning → unit / cloud example")
footer(s, p())

headers = ["Concept", "Meaning", "Cloud / Unit Example"]
rows_ha = [
    ["High Availability", "Minimize downtime via redundancy", "MIG with 2+ instances"],
    ["Fault Tolerance", "Survive component failure", "Autohealing recreates VM"],
    ["Redundancy", "Duplicate components", "Kubernetes replicas = 3"],
    ["SPOF", "One failure stops everything", "Single VM store → fixed by MIG"],
]
for i, h in enumerate(headers):
    left = Inches(0.5 + i * 4.2)
    add_rect(s, left, Inches(1.45), Inches(4.05), Inches(0.55), NAVY)
    tb(s, left + Inches(0.15), Inches(1.55), Inches(3.75), Inches(0.4),
       [(h, 13, True, WHITE)], align=PP_ALIGN.CENTER)
for r, row in enumerate(rows_ha):
    for c, cell in enumerate(row):
        left = Inches(0.5 + c * 4.2)
        top = Inches(2.05 + r * 0.9)
        bg = WHITE if r % 2 == 0 else SOFT
        add_rect(s, left, top, Inches(4.05), Inches(0.85), bg)
        tb(s, left + Inches(0.15), top + Inches(0.25), Inches(3.75), Inches(0.45),
           [(cell, 12, c == 0, NAVY if c == 0 else MUTED)], align=PP_ALIGN.CENTER)

tb(s, Inches(0.5), Inches(5.9), Inches(12.3), Inches(0.7),
   [("SPOF reminder: if losing ONE thing kills the app, you do not have HA yet.", 14, True, CORAL)])


# ═══════════════════════════════════════════════════════════════════════════
# 9. HA patterns
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "HA Patterns in the Cloud", "From one zone to planet-scale")
footer(s, p())

patterns = [
    ("Multi-instance", "MIG / K8s replicas so one crash is not an outage"),
    ("Load balancing", "Send traffic only to healthy backends"),
    ("Health checks", "Detect failures early; remove bad nodes"),
    ("Autohealing", "Replace unhealthy VMs/Pods automatically"),
    ("Multi-zone", "Survive one datacenter failure in a region"),
    ("Multi-region", "Survive a whole region outage (true DR)"),
]
for i, (t, d) in enumerate(patterns):
    col = i % 3
    row = i // 3
    left = Inches(0.5 + col * 4.2)
    top = Inches(1.45 + row * 2.4)
    add_rounded(s, left, top, Inches(3.95), Inches(2.15), WHITE)
    add_rect(s, left, top, Inches(3.95), Inches(0.55), [SKY, TEAL, CORAL, GOLD, SKY, TEAL][i])
    tb(s, left + Inches(0.2), top + Inches(0.12), Inches(3.5), Inches(0.4),
       [(t, 14, True, WHITE)], align=PP_ALIGN.CENTER)
    tb(s, left + Inches(0.25), top + Inches(0.9), Inches(3.5), Inches(0.9),
       [(d, 14, False, MUTED)], align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════
# 10. RTO & RPO
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "RTO & RPO — The Two Numbers That Drive DR Cost", "Decide these BEFORE picking a DR strategy")
footer(s, p())

card(s, Inches(0.5), Inches(1.4), Inches(6.0), Inches(3.6), "RTO — Recovery Time Objective", [
    "Maximum acceptable downtime.",
    "",
    "“How long can we be offline?”",
    "",
    "Example: payment UPI app RTO ≈ seconds–minutes.",
    "Campus portal after exams may accept hours.",
], CORAL)

card(s, Inches(6.8), Inches(1.4), Inches(6.0), Inches(3.6), "RPO — Recovery Point Objective", [
    "Maximum acceptable data loss (time window).",
    "",
    "“How much recent work can we afford to lose?”",
    "",
    "Hourly backup ⇒ RPO up to ~1 hour.",
    "Sync replication ⇒ RPO near zero.",
], SKY)

example_banner(s, Inches(5.3),
               "Lower RTO + lower RPO = more money (hot standby / active-active). Match the strategy to business pain, not to fashion.")


# ═══════════════════════════════════════════════════════════════════════════
# 11. Visual timeline RTO RPO
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "How RPO and RTO Sit on a Timeline", "Disaster → restore → back in business")
footer(s, p())

# Timeline visual using boxes
add_rounded(s, Inches(0.5), Inches(2.2), Inches(12.3), Inches(2.8), WHITE)
# Events
events = [
    (1.0, "Last good\nbackup", TEAL),
    (4.5, "DISASTER\nhits", CORAL),
    (8.0, "Service\nrestored", SKY),
    (11.0, "Users\nback", GREEN),
]
for x, label, c in events:
    circ = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(3.1), Inches(0.45), Inches(0.45))
    circ.fill.solid()
    circ.fill.fore_color.rgb = c
    circ.line.fill.background()
    tb(s, Inches(x - 0.55), Inches(3.7), Inches(1.6), Inches(0.8),
       [(label, 11, True, NAVY)], align=PP_ALIGN.CENTER)

# Line
add_rect(s, Inches(1.2), Inches(3.28), Inches(9.9), Inches(0.08), SOFT)

tb(s, Inches(1.2), Inches(1.5), Inches(10), Inches(0.5),
   [("RPO = time from last good backup → disaster  (data you may lose)", 14, True, TEAL)])
tb(s, Inches(1.2), Inches(5.3), Inches(10), Inches(0.5),
   [("RTO = time from disaster → service restored  (downtime users feel)", 14, True, CORAL)])
tb(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.7),
   [("Exam tip: never swap the two. RPO = data loss window · RTO = downtime window.", 13, True, NAVY)])


# ═══════════════════════════════════════════════════════════════════════════
# 12. DR Strategy tiers table
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Disaster Recovery Strategy Tiers", "From cheapest/slowest to near-zero downtime")
footer(s, p())

headers2 = ["Strategy", "Typical RTO", "Typical RPO", "Idea"]
rows2 = [
    ["Backup & Restore", "Hours", "Hours", "Restore when disaster hits — lowest cost"],
    ["Pilot Light", "Tens of minutes", "Minutes", "Minimal core always on; scale out on failover"],
    ["Warm Standby", "Minutes", "Minutes", "Scaled-down copy always running"],
    ["Hot Standby", "Near zero", "Near zero", "Full duplicate active — costly"],
    ["Multi-site Active-Active", "~Zero", "~Zero", "Both regions serve traffic (global LB)"],
]
for i, h in enumerate(headers2):
    left = Inches(0.35 + i * 3.2)
    add_rect(s, left, Inches(1.4), Inches(3.1), Inches(0.5), NAVY)
    tb(s, left + Inches(0.08), Inches(1.48), Inches(2.95), Inches(0.35),
       [(h, 12, True, WHITE)], align=PP_ALIGN.CENTER)
for r, row in enumerate(rows2):
    for c, cell in enumerate(row):
        left = Inches(0.35 + c * 3.2)
        top = Inches(1.95 + r * 0.8)
        bg = WHITE if r % 2 == 0 else SOFT
        add_rect(s, left, top, Inches(3.1), Inches(0.75), bg)
        color = [SKY, TEAL, CORAL, GOLD, GREEN][r] if c == 0 else INK
        tb(s, left + Inches(0.08), top + Inches(0.2), Inches(2.95), Inches(0.4),
           [(cell, 11, c == 0, color)], align=PP_ALIGN.CENTER)

tb(s, Inches(0.5), Inches(6.2), Inches(12.3), Inches(0.5),
   [("Rule: lower RTO/RPO ⇒ climb the ladder ⇒ higher ongoing cost.", 13, True, MUTED)])


# ═══════════════════════════════════════════════════════════════════════════
# 13. Strategy deep dive 1 — Backup & Restore / Pilot
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Strategy Deep Dive — Backup & Restore · Pilot Light", "Cost-efficient tiers for many university / SMB apps")
footer(s, p())

card(s, Inches(0.5), Inches(1.4), Inches(6.0), Inches(4.5), "Backup & Restore", [
    "How it works",
    "• Periodic backups to Cloud Storage / snapshots",
    "• On disaster: provision infra + restore data",
    "",
    "Pros: cheapest; simple to understand",
    "Cons: longest RTO; restore must be practiced",
    "",
    "Good for: blogs, internal tools, labs,",
    "non-critical reporting systems",
], SKY)

card(s, Inches(6.8), Inches(1.4), Inches(6.0), Inches(4.5), "Pilot Light", [
    "How it works",
    "• Tiny core always running in DR region",
    "  (e.g. DB replica + minimal services)",
    "• On disaster: scale up compute fleets",
    "",
    "Pros: faster than cold restore",
    "Cons: some always-on cost + runbook skill",
    "",
    "Good for: business apps that can wait",
    "tens of minutes but not hours",
], TEAL)


# ═══════════════════════════════════════════════════════════════════════════
# 14. Strategy deep dive 2 — Warm / Hot / Active-Active
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Strategy Deep Dive — Warm · Hot · Active-Active", "When downtime is measured in money, brand, or safety")
footer(s, p())

card(s, Inches(0.4), Inches(1.4), Inches(4.1), Inches(4.5), "Warm Standby", [
    "Scaled-down full copy",
    "always running in DR site.",
    "",
    "Flip DNS / LB → scale up.",
    "",
    "RTO/RPO: minutes.",
    "Cost: medium.",
], SKY)
card(s, Inches(4.65), Inches(1.4), Inches(4.1), Inches(4.5), "Hot Standby", [
    "Near full-size duplicate",
    "environment is ready.",
    "",
    "Failover is almost immediate.",
    "",
    "RTO/RPO: near zero.",
    "Cost: high (2× many resources).",
], CORAL)
card(s, Inches(8.9), Inches(1.4), Inches(4.0), Inches(4.5), "Active-Active", [
    "Multiple regions serve",
    "live traffic together.",
    "",
    "Global HTTP(S) LB + CDN.",
    "One region dies → others continue.",
    "",
    "RTO/RPO: ~zero.",
    "Hardest (data consistency!).",
], TEAL)


# ═══════════════════════════════════════════════════════════════════════════
# 15. Choosing strategy by business
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Choosing a Strategy — Business First", "Same cloud tools · different RTO/RPO budgets")
footer(s, p())

choices = [
    ("Campus ERP / results site", "Warm or Hot during exam week; Backup & Restore in off-season", CORAL),
    ("Food delivery order API", "Hot / Active-Active — every minute = cancelled orders", SKY),
    ("Hospital patient records", "Hot + strict RPO; often multi-region with audited restores", TEAL),
    ("Batch analytics job", "Fault tolerance + retries; Backup & Restore is often enough", GOLD),
]
for i, (t, d, c) in enumerate(choices):
    top = Inches(1.4 + i * 1.25)
    add_rounded(s, Inches(0.5), top, Inches(12.3), Inches(1.1), WHITE)
    add_rect(s, Inches(0.5), top, Inches(4.2), Inches(1.1), NAVY)
    tb(s, Inches(0.7), top + Inches(0.35), Inches(3.8), Inches(0.45), [(t, 13, True, WHITE)])
    tb(s, Inches(5.0), top + Inches(0.35), Inches(7.5), Inches(0.5), [(d, 13, False, MUTED)])


# ═══════════════════════════════════════════════════════════════════════════
# 16. Cloud-based backup solutions overview
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Cloud-Based Backup Solutions & Snapshots", "GCP building blocks you will see in labs & exams")
footer(s, p())

sols = [
    ("Compute Engine\nSnapshots", "Disk point-in-time copies\nIncremental & compressed\nRestore → new disk / VM", SKY),
    ("Cloud Storage", "Object versioning\nCross / dual-region buckets\nLifecycle → Coldline/Archive", TEAL),
    ("Cloud SQL", "Automated daily backups\nPoint-in-time recovery\nTight RPO for databases", CORAL),
    ("GKE / Apps", "Control plane etcd (Google)\nSnapshot Persistent Volumes\nMulti-cluster failover", GOLD),
]
for i, (t, body, c) in enumerate(sols):
    left = Inches(0.4 + i * 3.2)
    add_rounded(s, left, Inches(1.45), Inches(3.05), Inches(4.4), WHITE)
    add_rect(s, left, Inches(1.45), Inches(3.05), Inches(1.1), c)
    tb(s, left + Inches(0.15), Inches(1.55), Inches(2.75), Inches(0.9),
       [(t, 14, True, WHITE)], align=PP_ALIGN.CENTER)
    lines = body.split("\n")
    tb(s, left + Inches(0.2), Inches(2.9), Inches(2.7), Inches(2.5),
       [(ln, 13, False, MUTED) for ln in lines], align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════
# 17. Snapshots deep dive
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Compute Engine Snapshots — How Recovery Works", "Capture → store → rebuild")
footer(s, p())

card(s, Inches(0.5), Inches(1.4), Inches(6.0), Inches(4.5), "What a snapshot does", [
    "• Point-in-time copy of a persistent disk",
    "• Stored durably (Cloud Storage backed)",
    "• Incremental after the first snapshot",
    "• Create new disk / VM from snapshot",
    "",
    "Command shape:",
    "gcloud compute disks snapshot DISK \\",
    "  --snapshot-names=NAME --zone=ZONE",
], SKY)

card(s, Inches(6.8), Inches(1.4), Inches(6.0), Inches(4.5), "Recovery playbook", [
    "1. Identify last known-good snapshot",
    "2. Create disk from snapshot",
    "3. Create VM / attach disk",
    "4. Validate app & data",
    "5. Point LB / DNS to recovered stack",
    "",
    "Practice this BEFORE the outage.",
    "Measure actual restore time vs RTO.",
], TEAL)


# ═══════════════════════════════════════════════════════════════════════════
# 18. Cloud Storage & SQL backups
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Cloud Storage Versioning & Database PITR", "Protect objects and transactional data")
footer(s, p())

card(s, Inches(0.5), Inches(1.4), Inches(6.0), Inches(4.5), "Cloud Storage", [
    "• Object versioning → undelete / prior versions",
    "• Dual-region / multi-region → survive region loss",
    "• Lifecycle policies → cheaper cold storage",
    "",
    "Real life: media archives, ML datasets,",
    "backup dumps, Terraform state (carefully!).",
], SKY)

card(s, Inches(6.8), Inches(1.4), Inches(6.0), Inches(4.5), "Managed databases", [
    "Cloud SQL:",
    "• Automated daily backups",
    "• Point-in-Time Recovery (PITR)",
    "  → restore to e.g. 1:30 PM after 2 PM corruption",
    "",
    "Firestore / Datastore: export to GCS",
    "App state on GKE: snapshot PVs + app backups",
], CORAL)

example_banner(s, Inches(6.15),
               "A bad ALTER TABLE at 2:00 PM is a classic exam story — PITR is how you restore without yesterday’s full dump only.")


# ═══════════════════════════════════════════════════════════════════════════
# 19. Case studies intro
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Case Studies — DR Implementation", "From lab patterns to industry disasters")
footer(s, p())

cases = [
    ("01", "Fancy Store (Unit Lab)", "Single VM → MIG + LB + autohealing"),
    ("02", "Global E-commerce", "Multi-region + global LB + CDN"),
    ("03", "Data-Critical App", "Cloud SQL PITR for minute-level RPO"),
    ("04", "Batch Analytics", "Spot VMs + retries + state in GCS"),
    ("05", "Hotstar / Streaming", "Elasticity + CDN under traffic spikes"),
    ("06", "Banking / UPI", "Active-active mindset, strict RTO/RPO"),
]
for i, (num, title, sub) in enumerate(cases):
    col = i % 3
    row = i // 3
    left = Inches(0.5 + col * 4.2)
    top = Inches(1.5 + row * 2.4)
    add_rounded(s, left, top, Inches(3.95), Inches(2.15), WHITE)
    add_rect(s, left, top, Inches(3.95), Inches(0.55), [SKY, TEAL, CORAL, GOLD, SKY, TEAL][i])
    tb(s, left + Inches(0.2), top + Inches(0.12), Inches(3.5), Inches(0.4),
       [(num + "  " + title, 13, True, WHITE)], align=PP_ALIGN.CENTER)
    tb(s, left + Inches(0.25), top + Inches(0.95), Inches(3.5), Inches(0.8),
       [(sub, 13, False, MUTED)], align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════
# 20. Case 1 Fancy Store
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Case 1 — Fancy Store on Compute Engine (Lab)", "Problem → Solution → DR level")
footer(s, p())

card(s, Inches(0.5), Inches(1.4), Inches(4.0), Inches(4.5), "Problem", [
    "Single VM failure takes the",
    "whole store offline.",
    "",
    "Classic SPOF.",
    "No health checks,",
    "no automatic recovery.",
], CORAL)
card(s, Inches(4.7), Inches(1.4), Inches(4.0), Inches(4.5), "Solution", [
    "• Managed Instance Group",
    "• Health checks",
    "• Autohealing",
    "• HTTP(S) Load Balancer",
    "• Named ports for services",
    "",
    "Identical VMs from templates.",
], SKY)
card(s, Inches(8.9), Inches(1.4), Inches(4.0), Inches(4.5), "DR / HA level", [
    "HA within a zone.",
    "",
    "Extend to multi-zone MIG",
    "to survive zone failure.",
    "",
    "Still need backups for",
    "data / config recovery.",
], TEAL)


# ═══════════════════════════════════════════════════════════════════════════
# 21. Case 2 Global ecommerce
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Case 2 — Global E-commerce Regional Outage", "When one region is not enough")
footer(s, p())

card(s, Inches(0.5), Inches(1.4), Inches(6.0), Inches(3.6), "Problem", [
    "Regional outage affects ALL users.",
    "Revenue stops; brand damage is instant.",
    "Single-region HA cannot save you.",
], CORAL)
card(s, Inches(6.8), Inches(1.4), Inches(6.0), Inches(3.6), "Solution", [
    "• Multi-region deployment",
    "• Global HTTP(S) Load Balancer",
    "• Cloud CDN for static assets",
    "• Hot standby or Active-Active",
], SKY)
example_banner(s, Inches(5.3),
               "Amazon / Flipkart sale days and Shopify merchants design for regional failover so one datacenter wobble doesn’t cancel checkout.")


# ═══════════════════════════════════════════════════════════════════════════
# 22. Case 3 Data critical PITR
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Case 3 — Data-Critical Application (PITR)", "Corruption at 2:00 PM → restore to 1:30 PM")
footer(s, p())

steps = [
    ("1", "Incident", "Bad write / delete /\nransomware-like event"),
    ("2", "Decide RPO", "Business: lose ≤ 30 min\nof transactions"),
    ("3", "PITR restore", "Cloud SQL point-in-time\nto 1:30 PM state"),
    ("4", "Validate", "App checks, reconcile,\nreopen traffic"),
]
for i, (n, t, d) in enumerate(steps):
    left = Inches(0.45 + i * 3.2)
    add_rounded(s, left, Inches(1.55), Inches(3.05), Inches(3.2), WHITE)
    circ = s.shapes.add_shape(MSO_SHAPE.OVAL, left + Inches(1.15), Inches(1.8), Inches(0.7), Inches(0.7))
    circ.fill.solid()
    circ.fill.fore_color.rgb = [CORAL, SKY, TEAL, GOLD][i]
    circ.line.fill.background()
    tb(s, left + Inches(1.15), Inches(1.9), Inches(0.7), Inches(0.5),
       [(n, 16, True, WHITE)], align=PP_ALIGN.CENTER)
    tb(s, left + Inches(0.15), Inches(2.7), Inches(2.75), Inches(0.4),
       [(t, 14, True, NAVY)], align=PP_ALIGN.CENTER)
    tb(s, left + Inches(0.15), Inches(3.3), Inches(2.75), Inches(1.0),
       [(d, 12, False, MUTED)], align=PP_ALIGN.CENTER)

tb(s, Inches(0.5), Inches(5.2), Inches(12.3), Inches(1.2),
   [("DR level: Backup & Restore with a defined, tested RTO/RPO SLA — not “we hope the dump works.”", 14, True, NAVY),
    ("Real life: fintech ledgers, college marksheets, inventory DBs — PITR beats praying for yesterday’s .sql file.", 13, False, MUTED)])


# ═══════════════════════════════════════════════════════════════════════════
# 23. Case 4 Batch + industry
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Case 4 — Batch Analytics · Industry Stories", "Not every workload needs hot standby")
footer(s, p())

card(s, Inches(0.5), Inches(1.4), Inches(6.0), Inches(3.5), "Batch analytics (lab pattern)", [
    "Problem: job fails on one node",
    "Solution: Spot/Preemptible VMs + retry logic",
    "          Keep state in Cloud Storage",
    "DR level: fault tolerant, not tiny RTO",
    "",
    "Save money where downtime is acceptable.",
], TEAL)

card(s, Inches(6.8), Inches(1.4), Inches(6.0), Inches(3.5), "Industry analogues", [
    "• ML training on Spot → checkpoint to GCS",
    "• Nightly ETL: restart from last watermark",
    "• Log processing pipelines with replay",
    "",
    "Design for interruption, not immortality.",
], SKY)

example_banner(s, Inches(5.2),
               "Spotify Wrapped / recommendation batch jobs can retry; UPI settlement cannot. Same cloud — different DR tier.")


# ═══════════════════════════════════════════════════════════════════════════
# 24. Extra industry case studies
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "More Real-Life DR Stories", "Connect Unit II theory to products students use")
footer(s, p())

stories = [
    ("Hotstar / Disney+ (IPL)", "Traffic 50–100×. CDN + autoscale + multi-AZ. Goal: stay up for the last over, not just “have backups”."),
    ("PhonePe / Google Pay", "Payments need tight RTO/RPO. Multi-zone HA, strict failover drills, active capacity for salary-day spikes."),
    ("GitLab 2017 DB incident", "Bad delete + backup gaps = long outage. Lesson: test restores; multiple backup methods; clear RPO."),
    ("British Airways / airline OTAs", "Ops outages strand passengers. Multi-site thinking + runbooks matter more than pretty architecture slides."),
    ("IRCTC Tatkal window", "Thundering herd ≠ only HA. Queues, rate limits, graceful degradation are part of surviving “disasters of demand”."),
    ("Hospital / EHR systems", "Life-critical RTO. Hot/multi-site, audited backups, ransomware playbooks, offline contingency."),
]
for i, (t, d) in enumerate(stories):
    col = i % 2
    row = i // 2
    left = Inches(0.45 + col * 6.4)
    top = Inches(1.4 + row * 1.7)
    add_rounded(s, left, top, Inches(6.15), Inches(1.5), WHITE)
    add_rect(s, left, top, Inches(0.12), Inches(1.5), [SKY, TEAL, CORAL, GOLD, SKY, TEAL][i])
    tb(s, left + Inches(0.35), top + Inches(0.2), Inches(5.6), Inches(0.35), [(t, 13, True, NAVY)])
    tb(s, left + Inches(0.35), top + Inches(0.65), Inches(5.6), Inches(0.65), [(d, 12, False, MUTED)])


# ═══════════════════════════════════════════════════════════════════════════
# 25. DR checklist / best practices
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "DR Best Practices Checklist", "What production teams actually do")
footer(s, p())

checks = [
    "Write RTO/RPO per application — not one number for the whole college.",
    "Keep backups off the primary region (cross-region / dual-region).",
    "Automate snapshots + Cloud SQL backups; alert on backup failures.",
    "Run restore game days — measure real RTO, update runbooks.",
    "Combine HA (MIG/multi-zone) with DR (multi-region / backups).",
    "Use IaC so you can rebuild the stack when the console is on fire.",
    "Protect against ransomware: immutable / versioned backups + IAM least privilege.",
    "Document who declares a disaster and who executes failover.",
]
for i, b in enumerate(checks):
    top = Inches(1.35 + i * 0.65)
    add_rect(s, Inches(0.5), top, Inches(0.18), Inches(0.18), TEAL if i % 2 == 0 else SKY)
    tb(s, Inches(0.9), top - Inches(0.05), Inches(11.8), Inches(0.55), [(b, 13, False, INK)])


# ═══════════════════════════════════════════════════════════════════════════
# 26. Exam rapid revision
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Exam Rapid Revision — Unit II DR", "High-frequency points")
footer(s, p())

bullets = [
    "Backup = data copy · Disaster = unavailability event · DR = restore plan · Continuity = business keeps running.",
    "Backup types: Full, Incremental, Snapshot (GCE disks / GCS versioning).",
    "HA minimizes downtime via redundancy; Fault Tolerance continues despite failures; kill SPOFs.",
    "HA patterns: multi-instance, LB, health checks, autohealing, multi-zone, multi-region.",
    "RTO = max downtime · RPO = max data-loss window — never swap them.",
    "Strategy ladder: Backup & Restore → Pilot Light → Warm → Hot → Active-Active.",
    "GCP tools: GCE snapshots, GCS versioning/replication, Cloud SQL backups + PITR, GKE multi-cluster / PV snapshots.",
    "Case studies map: Fancy Store HA → Global multi-region → SQL PITR → Batch fault tolerance.",
]
for i, b in enumerate(bullets):
    top = Inches(1.35 + i * 0.65)
    add_rect(s, Inches(0.5), top, Inches(0.18), Inches(0.18), CORAL if i % 2 else SKY)
    tb(s, Inches(0.9), top - Inches(0.05), Inches(11.8), Inches(0.55), [(b, 12, False, INK)])


# ═══════════════════════════════════════════════════════════════════════════
# 27. Key takeaways
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Key Takeaways", "Design DR for people and money — not for slides")
footer(s, p())

takes = [
    ("Measure first", "Agree RTO & RPO with stakeholders, then pick the cheapest strategy that meets them."),
    ("HA ≠ full DR", "Multi-zone MIG is HA. Surviving a whole-region loss needs multi-region design + data protection."),
    ("Practice restores", "Untested backups are theater. Game days prove RTO; PITR drills prove RPO."),
    ("Match the workload", "UPI ≠ nightly ETL. Active-Active for payments; retries + GCS for batch."),
]
for i, (t, d) in enumerate(takes):
    top = Inches(1.45 + i * 1.25)
    add_rounded(s, Inches(0.5), top, Inches(12.3), Inches(1.1), WHITE)
    add_rect(s, Inches(0.5), top, Inches(0.12), Inches(1.1), [SKY, TEAL, CORAL, GOLD][i])
    tb(s, Inches(0.9), top + Inches(0.2), Inches(11.5), Inches(0.35), [(t, 16, True, NAVY)])
    tb(s, Inches(0.9), top + Inches(0.55), Inches(11.5), Inches(0.4), [(d, 13, False, MUTED)])


# ═══════════════════════════════════════════════════════════════════════════
# 28. Closing
# ═══════════════════════════════════════════════════════════════════════════
s = blank()
add_rect(s, Inches(0), Inches(0), SLIDE_W, SLIDE_H, NAVY)
add_rect(s, Inches(0), Inches(0), SLIDE_W, Inches(0.15), TEAL)
tb(s, Inches(0.8), Inches(2.2), Inches(11.5), Inches(0.8),
   [("Thank you", 44, True, WHITE)])
tb(s, Inches(0.8), Inches(3.2), Inches(11.5), Inches(0.5),
   [("Questions · Discussion · Lab mapping to DR tiers", 18, False, SKY_LIGHT)])
tb(s, Inches(0.8), Inches(4.3), Inches(11.5), Inches(1.2),
   [("Unit II focus: Disaster Recovery in Cloud", 14, False, SOFT),
    ("Backup & DR · HA & Fault Tolerance · RPO/RTO Strategies", 14, False, SOFT),
    ("Cloud Backups & Snapshots · Case Studies", 14, False, SOFT),
    ("Branch: cc-ii-ppt  ·  FIT & CS, Parul University", 13, False, MUTED)])
p()


out_dir = os.path.dirname(os.path.abspath(__file__))
out = os.path.join(out_dir, "CC-II-Unit-II-Disaster-Recovery-in-Cloud.pptx")
# Remove old networking-focused deck name if regenerating in place
old = os.path.join(out_dir, "CC-II-Unit-II-Networking-in-the-Cloud.pptx")
prs.save(out)
print(f"Saved {out}")
print(f"Slides: {len(prs.slides)}")
if os.path.exists(old):
    os.remove(old)
    print(f"Removed outdated {old}")
