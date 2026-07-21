#!/usr/bin/env python3
"""Generate Unit 4 teaching diagrams (PNG) and concept GIFs for the lecture deck."""
from PIL import Image, ImageDraw, ImageFont
import os

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "unit4-assets")
W, H = 1280, 720

NAVY = (10, 27, 46)
SKY = (26, 111, 181)
SKY_L = (61, 155, 224)
TEAL = (13, 165, 160)
CORAL = (232, 106, 76)
CLOUD = (244, 248, 252)
WHITE = (255, 255, 255)
SOFT = (232, 241, 248)
INK = (28, 43, 58)
MUTED = (90, 107, 124)
GOLD = (212, 160, 23)
GREEN = (31, 138, 95)
LIGHT_TEAL = (220, 245, 243)
LIGHT_SKY = (220, 236, 248)
LIGHT_CORAL = (255, 236, 230)
LIGHT_GOLD = (255, 246, 220)


def font(size, bold=False):
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    ]
    for p in candidates:
        if os.path.exists(p):
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


def rounded_rect(draw, xy, radius, fill, outline=None, width=2):
    draw.rounded_rectangle(xy, radius=radius, fill=fill, outline=outline, width=width)


def center_text(draw, text, cx, cy, f, fill=INK):
    bbox = draw.textbbox((0, 0), text, font=f)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text((cx - tw / 2, cy - th / 2), text, font=f, fill=fill)


def title_bar(draw, text, sub=None):
    draw.rectangle([0, 0, W, 70], fill=NAVY)
    draw.rectangle([0, 70, W, 78], fill=TEAL)
    draw.text((28, 18), text, font=font(28, True), fill=WHITE)
    if sub:
        draw.text((28, 48), sub, font=font(14), fill=SKY_L)


def save(img, name):
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, name)
    img.save(path)
    print("wrote", path)
    return path


def draw_pod(draw, x, y, label, color=GREEN, alive=True):
    fill = LIGHT_TEAL if alive else (240, 240, 240)
    outline = color if alive else MUTED
    rounded_rect(draw, [x, y, x + 100, y + 110], 10, fill, outline, 3)
    center_text(draw, "Pod", x + 50, y + 40, font(16, True), outline)
    center_text(draw, label, x + 50, y + 75, font(12), MUTED if alive else (180, 180, 180))


def make_architecture():
    img = Image.new("RGB", (W, H), CLOUD)
    d = ImageDraw.Draw(img)
    title_bar(d, "Kubernetes Architecture", "Control plane (brain) + worker nodes (muscle)")
    rounded_rect(d, [40, 110, 1240, 680], 18, WHITE, NAVY, 3)
    d.text((60, 125), "CLUSTER", font=font(16, True), fill=NAVY)
    rounded_rect(d, [70, 160, 1210, 340], 14, LIGHT_SKY, SKY, 3)
    d.text((90, 175), "CONTROL PLANE", font=font(18, True), fill=SKY)
    comps = [
        (120, "API Server", "Front door"),
        (390, "Scheduler", "Place Pods"),
        (660, "Controller Mgr", "Reconcile"),
        (940, "etcd", "State store"),
    ]
    for x, name, desc in comps:
        rounded_rect(d, [x, 215, x + 220, 310], 12, WHITE, SKY, 2)
        center_text(d, name, x + 110, 250, font(16, True), NAVY)
        center_text(d, desc, x + 110, 280, font(13), MUTED)
    d.polygon([(640, 350), (620, 380), (660, 380)], fill=TEAL)
    nodes = [(100, "Node 1", 2), (480, "Node 2", 2), (860, "Node 3", 1)]
    for x, name, pods in nodes:
        rounded_rect(d, [x, 410, x + 320, 650], 14, LIGHT_TEAL, TEAL, 3)
        center_text(d, name, x + 160, 435, font(16, True), TEAL)
        center_text(d, "kubelet · kube-proxy · runtime", x + 160, 465, font(12), MUTED)
        for i in range(pods):
            px = x + 40 + i * 130
            rounded_rect(d, [px, 500, px + 110, 610], 10, WHITE, GREEN, 2)
            center_text(d, "Pod", px + 55, 535, font(15, True), GREEN)
            center_text(d, "container", px + 55, 565, font(11), MUTED)
    save(img, "01-architecture.png")


def make_desired():
    img = Image.new("RGB", (W, H), CLOUD)
    d = ImageDraw.Draw(img)
    title_bar(d, "Desired State vs Actual State", "Controllers reconcile until they match")
    rounded_rect(d, [60, 120, 560, 420], 16, LIGHT_TEAL, TEAL, 3)
    center_text(d, "YOU DECLARE", 310, 160, font(20, True), TEAL)
    center_text(d, "Desired state", 310, 210, font(16, True), NAVY)
    for i, line in enumerate(["replicas: 3", "image: hello-app:v2", "stored in Git / YAML"]):
        center_text(d, line, 310, 270 + i * 40, font(18), INK)
    rounded_rect(d, [720, 120, 1220, 420], 16, LIGHT_CORAL, CORAL, 3)
    center_text(d, "CLUSTER REALITY", 970, 160, font(20, True), CORAL)
    center_text(d, "Actual state", 970, 210, font(16, True), NAVY)
    for i, line in enumerate(["2 Pods running", "1 Pod crashed", "needs healing"]):
        center_text(d, line, 970, 270 + i * 40, font(18), INK)
    d.polygon([(580, 250), (700, 250), (700, 230), (740, 260), (700, 290), (700, 270), (580, 270)], fill=GOLD)
    center_text(d, "Controller", 640, 200, font(14, True), GOLD)
    rounded_rect(d, [200, 480, 1080, 650], 14, WHITE, NAVY, 2)
    center_text(d, "Self-healing: create / replace Pods until actual == desired", 640, 540, font(20, True), NAVY)
    center_text(d, "Exam phrase: Desired state reconciled by controllers", 640, 590, font(16), MUTED)
    save(img, "02-desired-vs-actual.png")


def make_pipeline():
    img = Image.new("RGB", (W, H), CLOUD)
    d = ImageDraw.Draw(img)
    title_bar(d, "Deploy Pipeline on Kubernetes", "Practical 4 path")
    steps = [
        ("Write app", "Dockerfile", SKY),
        ("Build", "docker build", TEAL),
        ("Push", "Artifact Registry", GOLD),
        ("Deploy", "Deployment", CORAL),
        ("Expose", "Service / Ingress", GREEN),
        ("Verify", "pods · logs · IP", SKY),
    ]
    for i, (title, sub, color) in enumerate(steps):
        x = 40 + i * 205
        rounded_rect(d, [x, 220, x + 180, 420], 14, WHITE, color, 3)
        d.ellipse([x + 60, 250, x + 120, 310], fill=color)
        center_text(d, str(i + 1), x + 90, 280, font(22, True), WHITE)
        center_text(d, title, x + 90, 350, font(16, True), NAVY)
        center_text(d, sub, x + 90, 385, font(13), MUTED)
        if i < 5:
            d.polygon([(x + 185, 310), (x + 200, 320), (x + 185, 330)], fill=MUTED)
    rounded_rect(d, [80, 500, 1200, 640], 14, LIGHT_SKY, SKY, 2)
    center_text(d, "Image → Registry → Deployment → Service → Verify → (optional HPA)", 640, 550, font(20, True), NAVY)
    center_text(d, "Prefer kubectl apply -f manifests stored in Git", 640, 595, font(16), MUTED)
    save(img, "03-deploy-pipeline.png")


def make_services():
    img = Image.new("RGB", (W, H), CLOUD)
    d = ImageDraw.Draw(img)
    title_bar(d, "Kubernetes Service Types", "Memorise reachability + typical use")
    types = [
        ("ClusterIP", "Internal only", "Backend microservices", SKY, LIGHT_SKY),
        ("NodePort", "Node IP + high port", "Labs / simple expose", TEAL, LIGHT_TEAL),
        ("LoadBalancer", "External cloud IP", "Public web / APIs", CORAL, LIGHT_CORAL),
        ("ExternalName", "DNS CNAME out", "Point outside cluster", GOLD, LIGHT_GOLD),
    ]
    for i, (name, reach, use, color, bg) in enumerate(types):
        x = 40 + i * 310
        rounded_rect(d, [x, 140, x + 290, 560], 16, bg, color, 3)
        rounded_rect(d, [x, 140, x + 290, 220], 16, color)
        d.rectangle([x, 200, x + 290, 220], fill=color)
        center_text(d, name, x + 145, 180, font(20, True), WHITE)
        center_text(d, "Reachability", x + 145, 280, font(14, True), MUTED)
        center_text(d, reach, x + 145, 320, font(16, True), NAVY)
        center_text(d, "Typical use", x + 145, 400, font(14, True), MUTED)
        center_text(d, use, x + 145, 450, font(15, True), INK)
        d.ellipse([x + 115, 500, x + 175, 540], fill=color)
    rounded_rect(d, [80, 590, 1200, 680], 12, WHITE, NAVY, 2)
    center_text(d, "Must-say: Pod IPs are ephemeral — Services give a stable VIP / DNS name", 640, 635, font(18, True), NAVY)
    save(img, "04-service-types.png")


def make_ingress():
    img = Image.new("RGB", (W, H), CLOUD)
    d = ImageDraw.Draw(img)
    title_bar(d, "Ingress HTTP(S) Routing", "One entry · many services")
    rounded_rect(d, [40, 300, 200, 420], 12, LIGHT_GOLD, GOLD, 2)
    center_text(d, "Users", 120, 340, font(18, True), NAVY)
    center_text(d, "HTTPS", 120, 380, font(14), MUTED)
    d.polygon([(210, 350), (280, 360), (210, 370)], fill=MUTED)
    rounded_rect(d, [290, 250, 560, 470], 14, LIGHT_CORAL, CORAL, 3)
    center_text(d, "INGRESS", 425, 290, font(20, True), CORAL)
    center_text(d, "TLS terminate", 425, 340, font(14), INK)
    center_text(d, "/orders", 425, 380, font(16, True), NAVY)
    center_text(d, "/pay", 425, 420, font(16, True), NAVY)
    svcs = [(620, "orders-service", "ClusterIP", TEAL), (900, "payments-service", "ClusterIP", SKY)]
    for x, name, kind, color in svcs:
        rounded_rect(d, [x, 200, x + 240, 320], 12, WHITE, color, 3)
        center_text(d, name, x + 120, 245, font(15, True), NAVY)
        center_text(d, kind, x + 120, 285, font(13), MUTED)
        for j in range(2):
            px = x + 30 + j * 100
            rounded_rect(d, [px, 360, px + 80, 460], 8, LIGHT_TEAL, GREEN, 2)
            center_text(d, "Pod", px + 40, 410, font(14, True), GREEN)
    d.line([(560, 340), (620, 260)], fill=MUTED, width=3)
    d.line([(560, 400), (900, 260)], fill=MUTED, width=3)
    rounded_rect(d, [80, 520, 1200, 660], 12, WHITE, NAVY, 2)
    center_text(d, "Exam contrast: LoadBalancer ≈ one LB per service · Ingress = shared HTTP routing", 640, 570, font(17, True), NAVY)
    center_text(d, "Ideal for microservices APIs behind one hostname", 640, 615, font(15), MUTED)
    save(img, "05-ingress.png")


def make_gke():
    img = Image.new("RGB", (W, H), CLOUD)
    d = ImageDraw.Draw(img)
    title_bar(d, "GKE Autopilot vs Standard", "Exam comparison table as visuals")
    rounded_rect(d, [60, 120, 600, 650], 16, LIGHT_TEAL, GREEN, 3)
    rounded_rect(d, [60, 120, 600, 200], 16, GREEN)
    d.rectangle([60, 180, 600, 200], fill=GREEN)
    center_text(d, "Autopilot", 330, 160, font(26, True), WHITE)
    for i, r in enumerate([
        "Google manages nodes",
        "Pay for Pod resource requests",
        "Safer opinionated defaults",
        "Best for most labs & many apps",
        "create-auto …",
    ]):
        center_text(d, r, 330, 260 + i * 70, font(18), NAVY)
    rounded_rect(d, [680, 120, 1220, 650], 16, LIGHT_GOLD, GOLD, 3)
    rounded_rect(d, [680, 120, 1220, 200], 16, GOLD)
    d.rectangle([680, 180, 1220, 200], fill=GOLD)
    center_text(d, "Standard", 950, 160, font(26, True), WHITE)
    for i, r in enumerate([
        "You manage node pools",
        "Pay for node VMs",
        "GPUs / custom machines",
        "Best for special hardware",
        "clusters create …",
    ]):
        center_text(d, r, 950, 260 + i * 70, font(18), NAVY)
    save(img, "06-autopilot-vs-standard.png")


def gif_rolling_update():
    frames = []
    stages = [
        [("v1", True), ("v1", True), ("v1", True)],
        [("v2", True), ("v1", True), ("v1", True)],
        [("v2", True), ("v2", True), ("v1", True)],
        [("v2", True), ("v2", True), ("v2", True)],
    ]
    captions = [
        "Before: all Pods on hello-app:v1",
        "Rolling: replace one Pod with v2",
        "Continue: Service only hits Ready Pods",
        "Done: all Pods on hello-app:v2 (near zero downtime)",
    ]
    for stage, caption in zip(stages, captions):
        img = Image.new("RGB", (W, H), CLOUD)
        d = ImageDraw.Draw(img)
        title_bar(d, "Rolling Update (animated)", "Gradually replace Pods — Service stays up")
        rounded_rect(d, [80, 140, 1200, 280], 12, WHITE, SKY, 2)
        center_text(d, "Deployment hello-app  ·  replicas = 3", 640, 185, font(20, True), NAVY)
        center_text(d, caption, 640, 235, font(16), MUTED)
        for i, (lab, alive) in enumerate(stage):
            color = TEAL if lab == "v2" else CORAL
            draw_pod(d, 280 + i * 240, 360, lab, color=color, alive=alive)
        rounded_rect(d, [200, 540, 1080, 640], 12, LIGHT_SKY, SKY, 2)
        center_text(d, "Service → traffic only to Ready Pods", 640, 590, font(18, True), SKY)
        frames.append(img)
    frames = frames + [frames[-1], frames[-1]]
    path = os.path.join(OUT, "07-rolling-update.gif")
    frames[0].save(path, save_all=True, append_images=frames[1:], duration=900, loop=0)
    print("wrote", path)


def gif_hpa():
    frames = []
    counts = [1, 2, 3, 5, 4, 2]
    captions = [
        "Low load · HPA keeps min replicas",
        "CPU rises · HPA adds a Pod",
        "Load climbing · more replicas",
        "Peak traffic · scale to max",
        "Load easing · scale down begins",
        "Back to quieter traffic",
    ]
    for n, caption in zip(counts, captions):
        img = Image.new("RGB", (W, H), CLOUD)
        d = ImageDraw.Draw(img)
        title_bar(d, "HPA — Horizontal Pod Autoscaler", "Scale Pods based on metrics (CPU)")
        rounded_rect(d, [80, 130, 1200, 230], 12, WHITE, GOLD, 2)
        center_text(d, caption, 640, 160, font(18, True), NAVY)
        d.rectangle([120, 185, 1160, 210], fill=SOFT)
        fill_w = int(120 + (n / 5) * 1040)
        d.rectangle([120, 185, fill_w, 210], fill=CORAL if n >= 4 else TEAL)
        center_text(d, f"CPU signal  ·  replicas = {n}  (min=1 max=5)", 640, 250, font(15), MUTED)
        total_w = n * 120 + (n - 1) * 30
        start = (W - total_w) // 2
        for i in range(n):
            draw_pod(d, start + i * 150, 320, f"r{i+1}", color=GREEN)
        rounded_rect(d, [150, 500, 1130, 640], 12, LIGHT_TEAL, TEAL, 2)
        center_text(d, "kubectl autoscale deployment hello-app --cpu-percent=80 --min=1 --max=5", 640, 545, font(16, True), NAVY)
        center_text(d, "Analogy: MIG scales VMs · HPA scales Pods", 640, 595, font(15), MUTED)
        frames.append(img)
    frames = frames + [frames[-1]]
    path = os.path.join(OUT, "08-hpa-scale.gif")
    frames[0].save(path, save_all=True, append_images=frames[1:], duration=850, loop=0)
    print("wrote", path)


def gif_self_heal():
    frames = []
    stages = [
        ([True, True, True], "All 3 Pods healthy — matches desired replicas"),
        ([True, False, True], "Middle Pod crashes — actual ≠ desired"),
        ([True, False, True], "Controller detects failure…"),
        ([True, True, True], "Replacement Pod created — self-healed"),
    ]
    for pods, caption in stages:
        img = Image.new("RGB", (W, H), CLOUD)
        d = ImageDraw.Draw(img)
        title_bar(d, "Self-Healing Controllers", "Desired replicas stay true")
        rounded_rect(d, [100, 140, 1180, 250], 12, WHITE, TEAL, 2)
        center_text(d, "Desired: Deployment replicas = 3", 640, 175, font(20, True), NAVY)
        center_text(d, caption, 640, 220, font(16), MUTED)
        for i, alive in enumerate(pods):
            draw_pod(d, 300 + i * 230, 340, "ok" if alive else "dead",
                     color=GREEN if alive else CORAL, alive=alive)
        rounded_rect(d, [180, 520, 1100, 640], 12, LIGHT_CORAL if not all(pods) else LIGHT_TEAL,
                     CORAL if not all(pods) else TEAL, 2)
        msg = ("Controller Manager reconciles → create replacement Pod"
               if not all(pods) else "Actual state matches desired state again")
        center_text(d, msg, 640, 580, font(17, True), NAVY)
        frames.append(img)
    frames = frames + [frames[-1], frames[-1]]
    path = os.path.join(OUT, "09-self-healing.gif")
    frames[0].save(path, save_all=True, append_images=frames[1:], duration=900, loop=0)
    print("wrote", path)


def main():
    os.makedirs(OUT, exist_ok=True)
    make_architecture()
    make_desired()
    make_pipeline()
    make_services()
    make_ingress()
    make_gke()
    gif_rolling_update()
    gif_hpa()
    gif_self_heal()
    print(f"Assets ready in {OUT}")


if __name__ == "__main__":
    main()
