---
name: html-motion-to-pptx
description: Converts HTML web animations (CSS keyframes, Canvas, WebGL, JS requestAnimationFrame, SVG) into native PowerPoint 2D vector shapes with pseudo-3D orbital motion paths, depth scaling, and multi-layer Z-ordering.
---

# HTML Motion to Native Vector 3D PPTX Skill

## Overview
This skill provides instructions, architectural principles, mathematical projection algorithms, and automated scripts to convert dynamic HTML/CSS/JavaScript web animations into native Microsoft PowerPoint (`.pptx`) presentations. 

Unlike traditional methods that capture static screenshots, pre-render bloated MP4 videos, or insert jerky animated GIFs, this skill creates **100% pure 2D vector geometry** inside PowerPoint that moves along **mathematical 3D perspective trajectories** with depth scaling (`ScaleEffect`) and multi-layer Z-ordering.

---

## Key Principles & Constraints
1. **Zero Embedded Video / GIF**:
   - Never embed MP4, WebM, or GIF files. Slides must contain only native vector shapes (`msoShapeOval`, `msoShapeRectangle`, Bezier curves, textboxes).
   - Results in lightweight presentations (~60 KB vs 50 MB+), crisp 4K/8K resolution, and zero XML corruption.
2. **Hardware-Accelerated 60 FPS**:
   - PowerPoint renders native vector animations with hardware GPU acceleration.
   - Perpetual closed loops are achieved via `RepeatCount = 1000` (or `msoAnimRepeatInfinite`).
3. **Perspective Projection (Pseudo-3D Illusion)**:
   - 3D space $(x, y, z)$ is mapped to 2D screen coordinates $(X, Y)$ and scale factor $S(z)$.
   - Objects closer to the camera ($+z$) scale up ($>100\%$) and move in front.
   - Objects further away ($-z$) scale down ($<100\%$) and move behind.
4. **Layout Separation (HUD Dashboard Architecture)**:
   - Left Half ($X: 40 - 400\text{ pt}$): High-contrast semi-transparent text cards and statistics.
   - Right Half ($X: 420 - 940\text{ pt}$): Celestial stage / motion viewport.
   - Maintain a $25-30\text{ pt}$ safe margin so moving objects never collide with or obscure text.
5. **Seamless Single-Click Navigation (Zero Snap-back)**:
   - When slides have looping animations, PowerPoint natively treats a normal mouse click as "stop animation" (causing shapes to jump back to origin), requiring a second click to change slides.
   - Solution: Place a 100% transparent overlay (`Fill.Transparency = 1.0`) with `ActionSettings(1).Action = 1` (`ppActionNextSlide`). This allows instant 1-click slide advance without interrupting or snapping the active orbital animations.

---

## Mathematical Formulation

### 1. 4-Arc Cubic Bezier Ellipse
PowerPoint motion paths are normalized coordinate strings: `M 0 0 C c1x c1y c2x c2y endX endY ... Z`.
To create a closed elliptical orbit with semi-major axis $a$, semi-minor axis $b$, and center $(cx, cy)$, divided into 4 quadrants ($q = 0, 1, 2, 3$) starting at angle $\theta_0$:

The standard Bezier magic constant for circular/elliptical arcs is:
$$\kappa = \frac{4}{3}(\sqrt{2} - 1) \approx 0.5522847498$$

For each quadrant $q \in \{0, 1, 2, 3\}$ from angle $\theta_A = \theta_0 + q \cdot \frac{\pi}{2}$ to $\theta_B = \theta_0 + (q+1) \cdot \frac{\pi}{2}$:
$$C_1 = \left(cx + a\cos\theta_A - \kappa a\sin\theta_A,\; cy + b\sin\theta_A + \kappa b\cos\theta_A\right)$$
$$C_2 = \left(cx + a\cos\theta_B + \kappa a\sin\theta_B,\; cy + b\sin\theta_B - \kappa b\cos\theta_B\right)$$
$$\text{End} = \left(cx + a\cos\theta_B,\; cy + b\sin\theta_B\right)$$

Normalize each coordinate relative to the initial starting point $(X_{\text{init}}, Y_{\text{init}})$ and slide dimensions ($W = 960, H = 540$):
$$\Delta x = \frac{x - X_{\text{init}}}{960.0}, \quad \Delta y = \frac{y - Y_{\text{init}}}{540.0}$$

### 2. Depth Scaling via AutoReverse
In PowerPoint COM / OpenXML:
- Add a `ScaleEffect` behavior to the animation sequence with `AutoReverse = -1` (True).
- Set `ScaleEffect.ByX = 125` and `ScaleEffect.ByY = 125` (125%).
- When synchronized with the orbit duration, the shape naturally expands to 125% when orbiting forward and shrinks to 80% when swinging around the back.

---

## 4-Step Conversion Workflow

```
[ Step 1: DOM & Entity Extraction ]
  Analyze HTML structure: titles, badges, metrics, colors (hex).
  Extract moving entities: radii, periods, colors, groupings.
                   │
                   ▼
[ Step 2: Kinematic Parameterization ]
  Map Web coordinate loops / CSS keyframes to orbital parameters:
  - Center (cx, cy)
  - Semi-major axis a, Semi-minor axis b
  - Orbital period T (seconds)
  - Starting phase angle theta0
                   │
                   ▼
[ Step 3: Vector Synthesis & Animation Pipeline ]
  Generate native PowerPoint COM script or OpenXML markup:
  - Slide.Background.Fill (Hex dark space #030712)
  - Orbit ellipses: msoShapeOval (Fill.Visible = 0, Line.Visible = -1)
  - Celestial bodies & satellites: grouped vector shapes
  - Closed 4-Bezier motion path: MotionEffect.Path
  - Depth pulse: ScaleEffect
                   │
                   ▼
[ Step 4: Verification & Export ]
  Export high-resolution slide PNGs.
  Audit bounding boxes and ensure zero overlap between HUD and motion stage.
```

---

## Automation Script Template (PowerShell COM)

```powershell
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = 1
$pres = $ppt.Presentations.Add()
$pres.PageSetup.SlideWidth = 960.0
$pres.PageSetup.SlideHeight = 540.0

$slide = $pres.Slides.Add(1, 12) # Blank slide
$slide.Background.Fill.Solid()
$slide.Background.Fill.ForeColor.RGB = 0x120703 # #030712 BGR

# Create Orbit Guide
$orb = $slide.Shapes.AddShape(9, $cx - $a, $cy - $b, $a * 2, $b * 2)
$orb.Fill.Visible = 0; $orb.Line.Visible = -1
$orb.Line.ForeColor.RGB = 0x6E4A35; $orb.Line.Weight = 1.0; $orb.Line.Transparency = 0.35

# Create Moving Planet Vector
$pSh = $slide.Shapes.AddShape(9, $initX - 6, $initY - 6, 12, 12)
$pSh.Fill.Solid(); $pSh.Fill.ForeColor.RGB = 0xF8BD38; $pSh.Line.Visible = 0

# Assign 4-Bezier Closed Orbit (86 = msoAnimEffectPathCircle, 2 = WithPrevious)
$eff = $slide.TimeLine.MainSequence.AddEffect($pSh, 86, 0, 2)
$eff.Timing.Duration = $duration
$eff.Timing.RepeatCount = 1000
$eff.Timing.SmoothStart = 0; $eff.Timing.SmoothEnd = 0
$eff.Behaviors.Item(1).MotionEffect.Path = $bezierPath # Must end with " E"

# Assign Depth Scaling (54 = msoAnimEffectGrowShrink)
$effScale = $slide.TimeLine.MainSequence.AddEffect($pSh, 54, 0, 2)
$effScale.Timing.Duration = $duration / 2.0
$effScale.Timing.RepeatCount = 1000
$effScale.Timing.AutoReverse = -1
$sBeh = $effScale.Behaviors.Add(3) # 3 = msoAnimTypeScale
$sBeh.ScaleEffect.ByX = 125
$sBeh.ScaleEffect.ByY = 125

$pres.SaveAs("presentation.pptx")
```

---

## When to Activate This Skill
- The user provides an interactive HTML page with animations (CSS, Canvas, WebGL, JS) and wants it turned into PowerPoint.
- The user requests 3D-like rotating/revolving objects in PowerPoint without using video files or GIFs.
- Building presentation decks that require perpetual orbital motion, interactive telemetry dashboards, or sci-fi UI aesthetics.
