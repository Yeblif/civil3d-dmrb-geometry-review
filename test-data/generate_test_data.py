"""Generate synthetic Civil 3D-style test data for the geometry reviewer.

Creates a fictitious single-carriageway alignment with geometry deliberately
chosen to sit around CD 109 limits for a 70 kph design speed, so the review
has something to find:

  DEMO-MAINLINE  Line 200 m -> clothoid 60 m -> arc R255 (120 m)
                 -> clothoid 60 m -> Line 250 m      (total 690 m)
  Vertical       +2.0% -> crest at 300, L=68 m (K=17, on the R255 arc)
                 -> -2.0% -> sag at 520, L=60 m (K=20) -> +1.0%
  Superelevation 5.0% on the R255 curve (run-off on the clothoids)
  Design speed   70 kph from chainage 0

No real project, location or person is represented. Coordinates are an
arbitrary local grid.

Usage:  python3 generate_test_data.py   (writes files next to this script)
"""
import math
import os

HERE = os.path.dirname(os.path.abspath(__file__))
ALIGN = "DEMO-MAINLINE"
LAYOUT = "DEMO-MAINLINE - Design FG"
EXISTING = "DEMO-MAINLINE - Existing Ground"
DRAWING = "DEMO-Highway-Model"

E0, N0 = 1000.0, 2000.0          # arbitrary local grid origin
HEADING0 = 30.0                  # degrees, counter-clockwise from east
R = 255.0
LS = 60.0
ARC = 120.0
LINE1, LINE2 = 200.0, 250.0


def f(x, nd=12):
    return f"{x:.{nd}f}"


def pt(e, n):
    # LandXML points from Civil 3D are written "northing easting".
    return f"{f(n)} {f(e)}"


def clothoid_local(L, A2, n=4000):
    """Integrate a clothoid from zero curvature; returns x, y (left positive)."""
    x = y = 0.0
    ds = L / n
    for i in range(n):
        s = (i + 0.5) * ds
        th = s * s / (2.0 * A2)
        x += math.cos(th) * ds
        y += math.sin(th) * ds
    return x, y


def build_horizontal():
    out = []
    h = math.radians(HEADING0)
    e, n = E0, N0
    sta = 0.0

    def move(dx, dy, heading):
        # dx along heading, dy to the left of heading
        return (e + dx * math.cos(heading) - dy * math.sin(heading),
                n + dx * math.sin(heading) + dy * math.cos(heading))

    # Line 1
    e1, n1 = move(LINE1, 0.0, h)
    out.append(("Line", dict(dir=f(math.degrees(h)), length=f(LINE1)), [("Start", pt(e, n)), ("End", pt(e1, n1))]))
    e, n, sta = e1, n1, sta + LINE1

    # Entry clothoid, curving right (cw): local y is negative.
    A2 = R * LS
    x, y = clothoid_local(LS, A2)
    y = -y
    theta = LS / (2.0 * R)
    tan_short = abs(y) / math.sin(theta)
    tan_long = x - abs(y) / math.tan(theta)
    pi_e, pi_n = move(tan_long, 0.0, h)
    e2, n2 = move(x, y, h)
    out.append(("Spiral", dict(length=f(LS), radiusEnd=f(R), radiusStart="INF", rot="cw", spiType="clothoid",
                               theta=f(math.degrees(theta)), totalY=f(abs(y)), totalX=f(x),
                               tanLong=f(tan_long), tanShort=f(tan_short)),
                [("Start", pt(e, n)), ("PI", pt(pi_e, pi_n)), ("End", pt(e2, n2))]))
    e, n, sta = e2, n2, sta + LS
    h -= theta

    # Arc, curving right.
    delta = ARC / R
    ce, cn = move(0.0, -R, h)
    h_end = h - delta
    e3 = ce + R * math.cos(h_end + math.pi / 2)
    n3 = cn + R * math.sin(h_end + math.pi / 2)
    tangent = R * math.tan(delta / 2)
    pe, pn = move(tangent, 0.0, h)
    chord = 2 * R * math.sin(delta / 2)
    out.append(("Curve", dict(rot="cw", chord=f(chord), crvType="arc", delta=f(math.degrees(delta)),
                              dirEnd=f(math.degrees(h_end)), dirStart=f(math.degrees(h)),
                              external=f(R / math.cos(delta / 2) - R), length=f(ARC),
                              midOrd=f(R - R * math.cos(delta / 2)), radius=f(R), tangent=f(tangent)),
                [("Start", pt(e, n)), ("Center", pt(ce, cn)), ("End", pt(e3, n3)), ("PI", pt(pe, pn))]))
    e, n, sta = e3, n3, sta + ARC
    h = h_end

    # Exit clothoid: mirror of the entry clothoid, built from its far end.
    h_out = h - theta
    # End point found by running the standard clothoid backwards from the
    # tangent-out direction.
    x, y = clothoid_local(LS, A2)
    # In the reversed frame (heading h_out + pi), the spiral curves left.
    rev = h_out + math.pi
    # Start of exit spiral = end point E4 + local(x, y) in reversed frame,
    # so E4 = start - local(x, y) in reversed frame.
    dx = x * math.cos(rev) - y * math.sin(rev)
    dy = x * math.sin(rev) + y * math.cos(rev)
    e4, n4 = e - dx, n - dy
    pi_e = e4 + tan_long * math.cos(rev)
    pi_n = n4 + tan_long * math.sin(rev)
    out.append(("Spiral", dict(length=f(LS), radiusEnd="INF", radiusStart=f(R), rot="cw", spiType="clothoid",
                               theta=f(math.degrees(theta)), totalY=f(abs(y)), totalX=f(x),
                               tanLong=f(tan_long), tanShort=f(tan_short)),
                [("Start", pt(e, n)), ("PI", pt(pi_e, pi_n)), ("End", pt(e4, n4))]))
    e, n, sta = e4, n4, sta + LS
    h = h_out

    # Line 2
    e5 = e + LINE2 * math.cos(h)
    n5 = n + LINE2 * math.sin(h)
    out.append(("Line", dict(dir=f(math.degrees(h)), length=f(LINE2)), [("Start", pt(e, n)), ("End", pt(e5, n5))]))
    sta += LINE2
    return out, sta


def build_profile():
    # (station, elevation, curve length or None)
    pvis = [(0.0, 100.0, None), (300.0, 106.0, 68.0), (520.0, 101.6, 60.0), (690.0, 103.3, None)]
    lines = []
    for s, z, L in pvis:
        if L is None:
            lines.append(f"\t\t\t\t\t<PVI>{f(s)} {f(z)}</PVI>")
        else:
            lines.append(f'\t\t\t\t\t<ParaCurve length="{f(L)}">{f(s)} {f(z)}</ParaCurve>')
    return "\n".join(lines)


def build_existing(total):
    pts = []
    s = 0.0
    while s <= total + 1e-9:
        z = 99.0 + 3.0 * math.sin(s / 120.0) + 0.004 * s
        pts.append(f"{f(s, 3)} {f(z, 3)}")
        s += 30.0
    return " ".join(pts)


def write_landxml():
    geom, total = build_horizontal()
    el = []
    for tag, attrs, children in geom:
        a = " ".join(f'{k}="{v}"' for k, v in attrs.items())
        kids = "".join(f"\n\t\t\t\t\t<{c}>{v}</{c}>" for c, v in children)
        el.append(f"\t\t\t\t<{tag} {a}>{kids}\n\t\t\t\t</{tag}>")
    s_end1 = LINE1 + LS
    s_start2 = LINE1 + LS + ARC
    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<LandXML xmlns="http://www.landxml.org/schema/LandXML-1.2" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.landxml.org/schema/LandXML-1.2 http://www.landxml.org/schema/LandXML-1.2/LandXML-1.2.xsd" date="2026-01-01" time="00:00:00" version="1.2" language="English" readOnly="false">
\t<Units>
\t\t<Metric areaUnit="squareMeter" linearUnit="meter" volumeUnit="cubicMeter" temperatureUnit="celsius" pressureUnit="milliBars" diameterUnit="millimeter" angularUnit="decimal degrees" directionUnit="decimal degrees"/>
\t</Units>
\t<Project name="{DRAWING}.dwg"/>
\t<Application name="Autodesk Civil 3D" desc="Civil 3D" manufacturer="Autodesk, Inc." version="2026" manufacturerURL="www.autodesk.com/civil" timeStamp="2026-01-01T00:00:00">
\t</Application>
\t<Alignments name="">
\t\t<Alignment name="{ALIGN}" length="{f(total)}" staStart="0." desc="Synthetic test alignment">
\t\t\t<CoordGeom>
{chr(10).join(el)}
\t\t\t</CoordGeom>
\t\t\t<Profile name="{ALIGN}">
\t\t\t\t<ProfSurf name="{EXISTING}" state="existing">
\t\t\t\t\t<PntList2D>{build_existing(total)}</PntList2D>
\t\t\t\t</ProfSurf>
\t\t\t\t<ProfAlign name="{LAYOUT}">
{build_profile()}
\t\t\t\t</ProfAlign>
\t\t\t</Profile>
\t\t\t<Superelevation staStart="{f(s_end1)}" staEnd="{f(s_start2)}">
\t\t\t\t<BeginRunoffSta>{f(LINE1)}</BeginRunoffSta>
\t\t\t\t<FullSuperSta>{f(s_end1)}</FullSuperSta>
\t\t\t\t<FullSuperelev>5.</FullSuperelev>
\t\t\t\t<RunoffSta>{f(s_start2)}</RunoffSta>
\t\t\t\t<StartofRunoutSta>{f(s_start2 + LS)}</StartofRunoutSta>
\t\t\t</Superelevation>
\t\t\t<Feature name="SpeedStation" code="0" source="Autodesk Civil 3D">
\t\t\t\t<Property label="station" value="0.00000000"/>
\t\t\t\t<Property label="speed" value="70.00000000"/>
\t\t\t</Feature>
\t\t</Alignment>
\t</Alignments>
</LandXML>
"""
    path = os.path.join(HERE, "demo_mainline.xml")
    with open(path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(xml)
    return path, total


def write_super_csv(total):
    """Same columns and formatting as the DMRBEXPORTSUPER exporter."""
    header = ["Exported At", "Drawing", "Alignment", "Superelevation Region", "Station", "Description",
              "Smoothing Curve Length", "Left Outside Shoulder", "Left Outside Lane", "Left Inside Lane",
              "Left Inside Shoulder", "Right Outside Shoulder", "Right Outside Lane", "Right Inside Lane",
              "Right Inside Shoulder"]
    ts = "2026-01-01 00:00:00 +00:00"
    nc = -2.5   # normal crown
    full = 5.0  # full superelevation, curve to the right: left lane rises
    # Curving right: left lane goes from -2.5% to +5.0%, right lane from -2.5% to -5.0%.
    # Constant rate of 2.5% per 30 m: level crown at the tangent point,
    # full superelevation at the end of the 60 m clothoid.
    s0 = LINE1
    s_full = LINE1 + LS
    s_end_full = LINE1 + LS + ARC
    s_tan = s_end_full + LS
    rows = [
        (0.0, "", "Begin Alignment", nc, nc),
        (s0 - 30.0, "Curve.1", "End Normal Crown", nc, nc),
        (s0, "Curve.1", "Level Crown", 0.0, nc),
        (s0 + 30.0, "Curve.1", "Reverse Crown", -nc, nc),
        (s_full, "Curve.1", "Begin Full Super", full, -full),
        (s_end_full, "Curve.1", "End Full Super", full, -full),
        (s_tan - 30.0, "Curve.1", "Reverse Crown", -nc, nc),
        (s_tan, "Curve.1", "Level Crown", 0.0, nc),
        (s_tan + 30.0, "Curve.1", "Begin Normal Crown", nc, nc),
        (total, "", "End Alignment", nc, nc),
    ]

    def pct(v):
        return f"{v:.3f}%"

    lines = ["﻿" + ",".join(header)]
    for sta, region, desc, left, right in rows:
        lines.append(",".join([ts, DRAWING, ALIGN, region, f"{sta:.3f}m", desc, "0.000",
                               "", pct(left), "", "", "", pct(right), "", ""]))
    path = os.path.join(HERE, "demo_mainline_superelevation.csv")
    with open(path, "w", encoding="utf-8", newline="\r\n") as fh:
        fh.write("\n".join(lines) + "\n")
    return path


if __name__ == "__main__":
    xml_path, total = write_landxml()
    csv_path = write_super_csv(total)
    print(f"Wrote {xml_path} ({total:.3f} m)")
    print(f"Wrote {csv_path}")
