"""تقسيم ملاحظات المغذي إلى «زونات» عمل متوازنة جغرافيًا.

الهدف: توزيع الملاحظات على فرق ميدانية بحيث يحمل كل فريق عددًا متقاربًا من
الملاحظات، وتكون ملاحظات الفريق الواحد متجاورة على الأرض قدر الإمكان.

الخوارزمية: k-means مع قيد سعة (balanced k-means). في كل دورة نحسب مركز كل
زون، ثم نرتّب كل الأزواج (ملاحظة، زون) حسب المسافة تصاعديًا ونوزّع الملاحظات
بالأقرب فالأقرب مع منع أي زون من تجاوز سعته. هذا يمنع النتيجة المعتادة
لـ k-means العادي: زون بخمس ملاحظات وآخر بخمسين.
"""
from __future__ import annotations

import math
import random
from typing import Any, Dict, List, Optional, Tuple

# ألوان الزونات: كل لون مقرون بأيقونة جاهزة من مكتبة قوقل نفسها.
#
# «خرائطي» (My Maps) تتجاهل وسم <color> الذي يلوّن أيقونة بيضاء — وهي الطريقة
# التي تعمل في Google Earth — لكنها تحتفظ برابط الأيقونة كما هو. لذلك نعطي كل
# زون أيقونة ملوّنة مسبقًا من روابط قوقل القياسية، فتظهر الزونات بألوانها فور
# الاستيراد بلا تلوين يدوي. ولون الموقع مطابق للون الأيقونة حتى تتفق الخريطتان.
ZONE_PALETTE = [
    ("#DB4436", "red", "أحمر"),
    ("#4186F0", "blu", "أزرق"),
    ("#0F9D58", "grn", "أخضر"),
    ("#F4B400", "ylw", "أصفر"),
    ("#A23BC6", "purple", "بنفسجي"),
    ("#FF9900", "orange", "برتقالي"),
    ("#62AFF0", "ltblu", "سماوي"),
    ("#E9548D", "pink", "وردي"),
]

ZONE_COLORS = [c for c, _, _ in ZONE_PALETTE]

MAX_ZONES = 8

# قاعدة روابط أيقونات قوقل القياسية (مستقرة منذ سنين وتعمل داخل خرائطي)
ICON_BASE = "https://maps.google.com/mapfiles/kml/paddle/"


def color_for(i: int) -> str:
    return ZONE_PALETTE[i % len(ZONE_PALETTE)][0]


def color_name(i: int) -> str:
    """اسم اللون بالعربية — يُذكر مع رقم الزون حتى يميّزه الفني بلونه."""
    return ZONE_PALETTE[i % len(ZONE_PALETTE)][2]


def icon_for(i: int) -> str:
    """رابط أيقونة قوقل الملوّنة لهذا الزون."""
    return f"{ICON_BASE}{ZONE_PALETTE[i % len(ZONE_PALETTE)][1]}-circle.png"


def kml_color(hex_color: str, alpha: str = "ff") -> str:
    """KML يكتب اللون بترتيب معكوس: aabbggrr بدل rrggbb."""
    h = hex_color.lstrip("#")
    return f"{alpha}{h[4:6]}{h[2:4]}{h[0:2]}".lower()


# ------------------------------------------------------------------ الإحداثيات

def _as_float(v) -> Optional[float]:
    try:
        f = float(str(v).strip())
    except (TypeError, ValueError):
        return None
    return f if math.isfinite(f) else None


def extract_points(records: List[Dict[str, str]]) -> Tuple[List[Dict[str, Any]], int]:
    """يستخرج الملاحظات التي لها إحداثيات صالحة. يرجع (النقاط، عدد المتجاهَل).

    بعض الصفوف تأتي بإحداثيات فارغة أو مقلوبة (خط الطول مكان خط العرض)؛
    السعودية بين خطي عرض 16-33 وطول 34-56 تقريبًا، فنصحّح المقلوب ونتجاهل
    ما هو خارج المدى تمامًا."""
    points, skipped = [], 0
    for rec in records:
        lat, lon = _as_float(rec.get("lat")), _as_float(rec.get("lon"))
        if lat is None or lon is None:
            skipped += 1
            continue
        if not (-90 <= lat <= 90 and -180 <= lon <= 180):
            skipped += 1
            continue
        # تصحيح الانقلاب: خط عرض 42 مستحيل في السعودية بينما خط الطول 42 شائع
        if not (16 <= lat <= 33) and (16 <= lon <= 33) and (34 <= lat <= 56):
            lat, lon = lon, lat
        if lat == 0 and lon == 0:
            skipped += 1
            continue
        points.append({
            "note_id": rec.get("note_id", ""),
            "note": rec.get("note", ""),
            "office": rec.get("office", ""),
            "photo": rec.get("photo1", ""),
            "contractor": rec.get("contractor", ""),
            "lat": lat,
            "lon": lon,
        })
    return points, skipped


def _project(points: List[Dict[str, Any]]) -> List[Tuple[float, float]]:
    """إسقاط تقريبي إلى أمتار حتى تكون المسافات صحيحة (درجة الطول تقصر
    كلما ابتعدنا عن خط الاستواء، فلا يصح حساب المسافة على الدرجات مباشرة)."""
    if not points:
        return []
    mean_lat = sum(p["lat"] for p in points) / len(points)
    kx = 111320.0 * math.cos(math.radians(mean_lat))
    ky = 110540.0
    return [(p["lon"] * kx, p["lat"] * ky) for p in points]


# ------------------------------------------------------- التقسيم المتوازن

def _kmeans_pp_init(xy: List[Tuple[float, float]], k: int, rng: random.Random):
    centers = [xy[rng.randrange(len(xy))]]
    while len(centers) < k:
        d2 = []
        for p in xy:
            best = min((p[0] - c[0]) ** 2 + (p[1] - c[1]) ** 2 for c in centers)
            d2.append(best)
        total = sum(d2)
        if total <= 0:
            centers.append(xy[rng.randrange(len(xy))])
            continue
        r = rng.random() * total
        acc = 0.0
        for p, w in zip(xy, d2):
            acc += w
            if acc >= r:
                centers.append(p)
                break
    return centers


def resolve_capacities(n: int, k: int, slack: float,
                       caps: Optional[Dict[int, int]] = None) -> List[int]:
    """سعة كل زون. `caps` أعداد يحددها المستخدم لزونات بعينها (1-based)؛
    الباقي يتقاسم المتبقي بالتساوي. لو نقص المجموع عن عدد الملاحظات نوزّع
    الفرق على الزونات غير المحددة، ولو لم توجد رفعنا السعات بالتناسب."""
    base = int(math.ceil(n / k * (1.0 + max(0.0, slack))))
    out = [None] * k
    if caps:
        for z, v in caps.items():
            if 1 <= z <= k and v and v > 0:
                out[z - 1] = min(int(v), n)

    fixed = sum(v for v in out if v is not None)
    free = [i for i, v in enumerate(out) if v is None]
    remaining = max(0, n - fixed)
    if free:
        share = int(math.ceil(remaining / len(free) * (1.0 + max(0.0, slack)))) if remaining else 1
        for i in free:
            out[i] = max(1, share)
    elif fixed < n:
        scale = n / fixed if fixed else 1
        out = [max(1, int(math.ceil(v * scale))) for v in out]

    if not caps:
        out = [base] * k
    return [min(n, max(1, v)) for v in out]


def balanced_clusters(points: List[Dict[str, Any]], k: int,
                      iterations: int = 30, seed: int = 7,
                      slack: float = 0.0,
                      caps: Optional[Dict[int, int]] = None) -> List[int]:
    """يرجع رقم الزون لكل نقطة (0..k-1).

    `slack` هو التباين المسموح في أحجام الزونات: صفر يعني تساويًا صارمًا في
    العدد، وهو ما يجبر النقاط البعيدة على الانضمام لزون بعيد عنها لمجرد ملء
    الحصة فيتمدد الزون جغرافيًا. رفع التباين يسمح لكل نقطة بالذهاب لأقرب
    مركز فتتقلّص المسافات داخل الزون على حساب تفاوت الأعداد."""
    n = len(points)
    if n == 0:
        return []
    k = max(1, min(k, n, MAX_ZONES))
    if k == 1:
        return [0] * n

    xy = _project(points)
    rng = random.Random(seed)
    centers = _kmeans_pp_init(xy, k, rng)

    caps_list = resolve_capacities(n, k, slack, caps)
    assign = [0] * n
    for _ in range(iterations):
        pairs = []
        for i, p in enumerate(xy):
            for c, cen in enumerate(centers):
                dx, dy = p[0] - cen[0], p[1] - cen[1]
                pairs.append((dx * dx + dy * dy, i, c))
        pairs.sort()

        new_assign = [-1] * n
        load = [0] * k
        placed = 0
        for _d, i, c in pairs:
            if new_assign[i] != -1 or load[c] >= caps_list[c]:
                continue
            new_assign[i] = c
            load[c] += 1
            placed += 1
            if placed == n:
                break
        # احتياط: أي نقطة لم تجد مكانًا (نادر) تذهب لأقل الزونات حملًا
        for i in range(n):
            if new_assign[i] == -1:
                c = min(range(k), key=lambda z: load[z])
                new_assign[i] = c
                load[c] += 1

        new_centers = []
        for c in range(k):
            members = [xy[i] for i in range(n) if new_assign[i] == c]
            if members:
                new_centers.append((sum(m[0] for m in members) / len(members),
                                    sum(m[1] for m in members) / len(members)))
            else:
                new_centers.append(centers[c])

        stable = new_assign == assign
        assign, centers = new_assign, new_centers
        if stable:
            break
    return assign


# ------------------------------------------------------------ حدود الزون

def convex_hull(pts: List[Tuple[float, float]]) -> List[Tuple[float, float]]:
    """مضلّع محدّب يحيط بنقاط الزون (خوارزمية Andrew)."""
    pts = sorted(set(pts))
    if len(pts) <= 2:
        return pts

    def cross(o, a, b):
        return (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0])

    lower = []
    for p in pts:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], p) <= 0:
            lower.pop()
        lower.append(p)
    upper = []
    for p in reversed(pts):
        while len(upper) >= 2 and cross(upper[-2], upper[-1], p) <= 0:
            upper.pop()
        upper.append(p)
    return lower[:-1] + upper[:-1]


def haversine_km(a: Tuple[float, float], b: Tuple[float, float]) -> float:
    lat1, lon1, lat2, lon2 = map(math.radians, (a[0], a[1], b[0], b[1]))
    h = (math.sin((lat2 - lat1) / 2) ** 2
         + math.cos(lat1) * math.cos(lat2) * math.sin((lon2 - lon1) / 2) ** 2)
    return 2 * 6371.0 * math.asin(min(1.0, math.sqrt(h)))


def order_route(members: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], float]:
    """يرتّب مواقع الزون كمسار قيادة قصير: نبدأ من الأقرب للشمال الغربي ثم
    الأقرب فالأقرب، ثم نحسّن الترتيب بتبديلات 2-opt. يرجع (الترتيب، الطول كم).

    الهدف تقليل زمن التنقل داخل الزون، لا مجرد تجميع النقاط."""
    n = len(members)
    if n <= 2:
        d = haversine_km((members[0]["lat"], members[0]["lon"]),
                         (members[-1]["lat"], members[-1]["lon"])) if n == 2 else 0.0
        return list(members), round(d, 1)

    pts = [(m["lat"], m["lon"]) for m in members]

    def dist(i, j):
        return haversine_km(pts[i], pts[j])

    start = max(range(n), key=lambda i: pts[i][0] - pts[i][1])   # أقصى الشمال الغربي
    order = [start]
    unused = set(range(n)) - {start}
    while unused:
        last = order[-1]
        nxt = min(unused, key=lambda j: dist(last, j))
        order.append(nxt)
        unused.discard(nxt)

    # 2-opt: عكس مقاطع من المسار متى قصّر الطول الكلي
    improved = True
    rounds = 0
    while improved and rounds < 12:
        improved = False
        rounds += 1
        for i in range(n - 2):
            for j in range(i + 2, n):
                a, b = order[i], order[i + 1]
                c, d2 = order[j], order[(j + 1) % n]
                if (j + 1) % n == i:
                    continue
                before = dist(a, b) + dist(c, d2)
                after = dist(a, c) + dist(b, d2)
                if after + 1e-9 < before:
                    order[i + 1:j + 1] = reversed(order[i + 1:j + 1])
                    improved = True

    total = sum(dist(order[i], order[i + 1]) for i in range(n - 1))
    return [members[i] for i in order], round(total, 1)


def route_links(ordered: List[Dict[str, Any]], per_leg: int = 10) -> List[Dict[str, Any]]:
    """روابط خرائط قوقل للمسار. تُقسَّم إلى مقاطع لأن الرابط الواحد لا يحتمل
    محطات كثيرة (الصيغة الرسمية تسمح بثلاث محطات فقط على الجوال)، وكل مقطع
    يبدأ من آخر موقع في سابقه فلا ينقطع المسار."""
    links = []
    i = 0
    n = len(ordered)
    while i < n:
        leg = ordered[i:i + per_leg]
        if i > 0:
            leg = [ordered[i - 1]] + leg          # وصل المقطع بسابقه
        path = "/".join(f'{p["lat"]:.6f},{p["lon"]:.6f}' for p in leg)
        links.append({
            "url": "https://www.google.com/maps/dir/" + path,
            "stops": len(leg),
            "index": len(links) + 1,
        })
        i += per_leg
    return links


ISOLATED_KM = 3.0


def mark_isolation(points: List[Dict[str, Any]]) -> int:
    """يحسب لكل ملاحظة بعدها عن أقرب ملاحظة أخرى، ويميّز المعزولة منها.

    الملاحظة المعزولة هي ما يمدّد الزون جغرافيًا مهما حسّنّا التقسيم: لا
    خوارزمية تستطيع تقريب نقطة تبعد عشرة كيلومترات عن الجميع، والمفيد أن
    يراها المخطّط ليقرر بشأنها."""
    n = len(points)
    if n < 2:
        for p in points:
            p["iso_km"] = 0.0
            p["isolated"] = False
        return 0
    xy = _project(points)
    count = 0
    for i, p in enumerate(points):
        best = None
        xi, yi = xy[i]
        for j in range(n):
            if j == i:
                continue
            dx, dy = xi - xy[j][0], yi - xy[j][1]
            d2 = dx * dx + dy * dy
            if best is None or d2 < best:
                best = d2
        km = round((best ** 0.5) / 1000.0, 2) if best else 0.0
        p["iso_km"] = km
        p["isolated"] = km >= ISOLATED_KM
        if p["isolated"]:
            count += 1
    return count


def build_zones(points: List[Dict[str, Any]], k: int, slack: float = 0.0,
                overrides: Optional[Dict[str, int]] = None,
                caps: Optional[Dict[int, int]] = None,
                names: Optional[Dict[str, str]] = None) -> List[Dict[str, Any]]:
    """يجمع نتيجة التقسيم: لكل زون نقاطه ولونه واسم مقاوله وحدوده ومساره."""
    assign = balanced_clusters(points, k, slack=slack, caps=caps)

    # نقل يدوي من المستخدم: يغلب على نتيجة الخوارزمية
    if overrides:
        k_max = max(1, min(k, len(points), MAX_ZONES))
        for i, p in enumerate(points):
            z = overrides.get(str(p.get("note_id", "")))
            if z is not None and 1 <= z <= k_max:
                assign[i] = z - 1

    k_eff = (max(assign) + 1) if assign else 0
    zones = []
    for c in range(k_eff):
        members = [p for p, a in zip(points, assign) if a == c]
        if not members:
            continue
        lat_c = sum(p["lat"] for p in members) / len(members)
        lon_c = sum(p["lon"] for p in members) / len(members)
        hull = convex_hull([(p["lon"], p["lat"]) for p in members])
        span = 0.0
        for i in range(len(hull)):
            for j in range(i + 1, len(hull)):
                span = max(span, haversine_km((hull[i][1], hull[i][0]),
                                              (hull[j][1], hull[j][0])))
        ordered, route_km = order_route(members)
        zones.append({
            "index": c + 1,
            "name": (names or {}).get(str(c + 1), "").strip(),
            "color": color_for(c),
            "icon": icon_for(c),
            "color_name": color_name(c),
            "count": len(members),
            "center": {"lat": lat_c, "lon": lon_c},
            "hull": [{"lat": y, "lon": x} for x, y in hull],
            "span_km": round(span, 1),
            "route_km": route_km,
            "links": route_links(ordered),
            "points": ordered,
        })
    zones.sort(key=lambda z: z["index"])
    return zones


# --------------------------------------------------------------------- KML

def _esc(s: str) -> str:
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def build_kml(zones: List[Dict[str, Any]], title: str) -> str:
    """ملف KML يفتح في Google Earth ويُستورد في Google My Maps بألوانه."""
    out = ['<?xml version="1.0" encoding="UTF-8"?>',
           '<kml xmlns="http://www.opengis.net/kml/2.2">',
           f"<Document><name>{_esc(title)}</name>"]

    for z in zones:
        c = kml_color(z["color"])
        # أيقونة ملوّنة جاهزة لكل زون بدل تلوين أيقونة بيضاء: «خرائطي» تتجاهل
        # وسم <color> لكنها تحتفظ برابط الأيقونة، فتظهر الألوان فور الاستيراد
        out.append(f'<Style id="pin{z["index"]}">'
                   f'<IconStyle><scale>1.1</scale>'
                   f'<Icon><href>{z["icon"]}</href></Icon>'
                   f'<hotSpot x="0.5" y="0" xunits="fraction" yunits="fraction"/>'
                   f'</IconStyle>'
                   f'<LabelStyle><scale>0.8</scale></LabelStyle></Style>')
        out.append(f'<Style id="area{z["index"]}">'
                   f'<LineStyle><color>{c}</color><width>3</width></LineStyle>'
                   f'<PolyStyle><color>{kml_color(z["color"], "33")}</color></PolyStyle></Style>')

    for z in zones:
        title = f'زون {z["index"]}'
        if z.get("name"):
            title += f' — {_esc(z["name"])}'
        out.append(f'<Folder><name>{title} ({z["count"]} ملاحظة)</name>')
        if len(z["hull"]) >= 3:
            ring = " ".join(f'{p["lon"]},{p["lat"]},0' for p in z["hull"])
            first = z["hull"][0]
            ring += f' {first["lon"]},{first["lat"]},0'
            out.append(f'<Placemark><name>حدود زون {z["index"]}</name>'
                       f'<styleUrl>#area{z["index"]}</styleUrl>'
                       f'<Polygon><outerBoundaryIs><LinearRing>'
                       f'<coordinates>{ring}</coordinates>'
                       f'</LinearRing></outerBoundaryIs></Polygon></Placemark>')
        zone_label = f'زون {z["index"]}' + (f' — {_esc(z["name"])}' if z.get("name") else "")
        for order_no, p in enumerate(z["points"], 1):
            # صورة "قبل" داخل الوصف: يفتحها الفني من الخريطة فيرى الزاوية التي
            # صُوّرت منها الملاحظة قبل وصوله للموقع
            img = ""
            photo = (p.get("photo") or "").strip()
            if photo.startswith("http"):
                img = (f'<br><a href="{_esc(photo)}">'
                       f'<img src="{_esc(photo)}" width="260"></a>')
            desc = (f'<![CDATA[<b>ملاحظة {_esc(p["note_id"])}</b><br>'
                    f'{_esc(p["note"])}<br>المكتب: {_esc(p["office"])}<br>'
                    f'{zone_label} · الترتيب {order_no}{img}]]>')
            out.append(f'<Placemark><name>{_esc(p["note_id"]) or str(order_no)}</name>'
                       f'<description>{desc}</description>'
                       f'<styleUrl>#pin{z["index"]}</styleUrl>'
                       f'<Point><coordinates>{p["lon"]},{p["lat"]},0</coordinates></Point></Placemark>')
        out.append("</Folder>")

    out.append("</Document></kml>")
    return "\n".join(out)
