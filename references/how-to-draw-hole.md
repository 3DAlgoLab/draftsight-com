# Hole Profile & Top View

## M20TAPDP45

For a **M20 tapped hole, depth 45 mm**, the 2D CAD representation depends on whether you're drawing:

1. **Top View (plan view)**
2. **Section View / Side View** (most common for manufacturing drawings)

### 1\. Understand the callout

`M20 TAP DP 45`

Means:

- **M20** = Metric internal thread, nominal diameter 20 mm
- **TAP** = Tapped (threaded) hole
- **DP 45** = Thread depth = 45 mm

For a standard M20 coarse thread:

- Major diameter = 20 mm
- Tap drill diameter ≈ 17.5 mm

---

### 2\. Side / Section View Drawing

Draw the hole centerline first.

#### Threaded portion

According to ISO drafting practice:

- Major diameter line (M20) = thick line
- Minor diameter line (~17.5 mm) = thin line
- Depth = 45 mm

Example:
```
      ┌──────────────────── Surface
      │
      │  ┃            ┃   ← Major dia Ø20 (thick)
      │  │            │
      │  │            │
      │  │            │
45 mm │  │            │
      │  │            │
      │  │            │
      │  ┃            ┃
      │    \        /
      │      \____/      Drill point
      │
```

In AutoCAD:

1. Draw two vertical lines spaced **20 mm** apart.
2. Draw two thin lines spaced **17.5 mm** apart.
3. Extend threaded length = **45 mm**.
4. Add drill-point angle (typically 118° or 120°) at bottom if showing the drilled hole.

---

### 3\. Top View Drawing

For internal threads:

Draw two concentric circles:

- Outer circle = Ø20 (thick)
- Inner circle = Ø17.5 (thin)

⊙

Where:

- Large circle = thread major diameter
- Small circle = thread minor diameter

Add center marks.

---

### 4\. Typical Drawing Note

Instead of dimensioning all thread geometry, usually place a leader note:

M20 × 2.5 TAP
DP 45

or

M20 TAP
45 DEEP

or per ISO:
M20 × 2.5 - 6H
45 DP

---

### 5\. AutoCAD Construction Method

Commands:

L (draw centerline)

OFFSET 10 (create Ø20 boundaries)
OFFSET 8.75 (create Ø17.5 boundaries)
TRIM
DIMLINEAR
MLEADER
Then add the note:
M20 TAP DP 45

For a production drawing, the preferred method is usually **a hole symbol callout with a leader** rather than drawing every thread detail:

⌀
M20 × 2.5 TAP
45 DP

This is the standard style most machinists and CNC programmers expect to see.

---

### 6. Standard tap dimensions (ISO 261 / ISO 724)

Source: web search 2026-10-27 - Allied Machine ISO metric tap drill chart,
boltlab.io tapping atlas (ISO 2306/261/262/724/965-1), Slugger Tools
ISO 261/724 chart. Three sources agree on every value below.

`Tap drill` is the pre-tap hole diameter = the hole wall diameter to draw.
The ISO basic *minor diameter* (D1) is slightly smaller; the tap drill
gives ~75% thread engagement and is what drawings dimension.

#### Coarse pitch series (the default)

| Thread | Pitch mm | Major D mm | Minor D1 mm | Tap drill mm |
|--------|---------|-----------|-------------|--------------|
| M1.6   | 0.35    | 1.6       | 1.171       | 1.25         |
| M2     | 0.4     | 2.0       | 1.509       | 1.60         |
| M2.5   | 0.45    | 2.5       | 1.948       | 2.05         |
| M3     | 0.5     | 3.0       | 2.387       | 2.50         |
| M4     | 0.7     | 4.0       | 3.141       | 3.30         |
| M5     | 0.8     | 5.0       | 4.019       | 4.20         |
| M6     | 1.0     | 6.0       | 4.773       | 5.00         |
| M8     | 1.25    | 8.0       | 6.466       | 6.80         |
| M10    | 1.5     | 10.0      | 8.160       | 8.50         |
| M12    | 1.75    | 12.0      | 9.853       | 10.20        |
| M14    | 2.0     | 14.0      | 11.546      | 12.00        |
| M16    | 2.0     | 16.0      | 13.546      | 14.00        |
| M18    | 2.5     | 18.0      | 15.294      | 15.50        |
| M20    | 2.5     | 20.0      | 17.294      | 17.50        |
| M22    | 2.5     | 22.0      | 19.294      | 19.50        |
| M24    | 3.0     | 24.0      | 20.752      | 21.00        |
| M27    | 3.0     | 27.0      | 23.752      | 24.00        |
| M30    | 3.5     | 30.0      | 26.211      | 26.50        |
| M33    | 3.5     | 33.0      | 29.211      | 29.50        |
| M36    | 4.0     | 36.0      | 31.670      | 32.00        |

#### Common fine pitch series (write the pitch, e.g. M20 x 2)

| Thread  | Pitch mm | Tap drill mm |
|---------|---------|--------------|
| M6 x 0.75  | 0.75  | 5.25 |
| M8 x 1.0   | 1.0   | 7.00 |
| M10 x 1.25 | 1.25  | 8.80 |
| M12 x 1.25 | 1.25  | 10.80 |
| M16 x 1.5  | 1.5   | 14.50 |
| M18 x 1.5  | 1.5   | 16.50 |
| M20 x 1.5  | 1.5   | 18.50 |
| M24 x 2.0  | 2.0   | 22.00 |
| M27 x 2.0  | 2.0   | 25.00 |
| M30 x 2.0  | 2.0   | 28.00 |
| M33 x 2.0  | 2.0   | 31.00 |
| M36 x 3.0  | 3.0   | 33.00 |

#### Depth conventions for the section

- **118-degree drill point** below the thread depth; point cone length
  = (tap drill / 2) / tan(59 degrees). M20: 5.47 mm; M30: 8.28 mm.
- **Drilled depth** = thread depth + 2 to 3 pitches so the 118-degree
  point clears. M20DP45: 50 to 52.5 mm; M30DP45: 52 to 55.5 mm. The drawn family instead
  rounds the cone to a whole millimetre and adds exactly that - M30DP45 = 45 + 8 = 53,
  M30DP55 = 55 + 8 = 63, M20DP45 = 45 + 5 = 50 - which still lands inside the band.
- Fine-pitch callout must state the pitch (M20 x 2); coarse is implied.

---

### 7. Drawing it in one call

`draftsight_com.tapped_hole(sheet, "M30", 55)` draws the representation above - section plus top
view, dimensions and callouts - and returns the numbers it used. `tapped_hole_spec(size, depth)`
is the same output as a plain spec dict, computed with no COM, so the arithmetic is reviewable and
testable. The table below is measured out of the drawn family (`M30DP45`, `M30DP55`), in mm.

| Item | Rule | M30 x 3.5 | M20 x 2.5 |
|---|---|---|---|
| Tap drill | major - pitch | 26.5 | 17.5 |
| Point cone | (tap drill / 2) / tan 59°, rounded to the mm | 8 | 5 |
| Drilled depth | thread depth + cone | 55 + 8 = 63 | 45 + 5 = 50 |
| Section axis | `view_offset` (75) left of the top view | -75 | -75 |
| Vertical centring | top face at +drilled/2, drill point at -drilled/2 | +/-31.5 | +/-25 |
| Thread-depth dim | 32 right of the section axis | 55 | 45 |
| Drilled-depth dim | 30 left of the section axis | 63 | 50 |
| Tap-drill dim | 9.25 below the top-view circle edge | Ø26.5 | Ø17.5 |
| Major arc, top view | 90° to 360° - a quarter missing from the upper right | r15 | r10 |
| View titles | centred 26 below the drill point | | |

Layers: `OUTLINE` 7 (minor lines, point cone, top-view circle), `MAJOR` 8 (major lines and arc),
`THREAD` 2 (ANSI31 at scale 0.8, which renders 3.57 mm pitch, stopping at thread depth so the cone
stays clear), `DIM` 3, `CENTER` 1 with the `CENTER` linetype, `TEXT` 7. Sizes are
`draftsight_com.HOUSE_STYLE`, which is what those DWGs answer through `house_style()`.

Two departures from ISO, both inherited from the existing set - change them on purpose, not by
accident:

- **Top-view line weight is swapped.** ISO draws the major diameter as the thick 3/4 arc and the
  minor as a thin full circle; the family draws the minor circle thick (`OUTLINE`) and the major
  arc thin (`MAJOR`). In the *section* the family matches ISO: the hole wall (minor) is what the
  cutting plane passes through, so it is the thick line.
- **`M20DP45.dwg` is a hand-drawn predecessor, not a template.** Its titles sit off-centre, its
  callout reads `M20 x 2.5 TAP, DP45`, and it carries no drilled-depth dimension. The two M30
  drawings define the convention, and the builder reproduces them line for line.

