---
name: web3d-motion-to-pptx
description: Converts Web 3D interactive applications (Three.js, WebGL, Canvas), CSS styles (glassmorphism, neon glow, gradients, soft edges), and JavaScript physics/kinematics into 100% native vector PowerPoint shapes and smooth perpetual timeline animations.
---

# Web3D Motion & CSS to PowerPoint (web3d-motion-to-pptx)

Transform interactive Web 3D applications, CSS visual effects, and JavaScript mathematical physics into 100% native vector PowerPoint (.pptx) presentations. Zero video files, zero animated GIFs — pure vector geometry and native animation timeline behaviors.

---

## Core Engine Architecture

```
┌────────────────────────────────────────────────────────┐
│             Web Source (HTML / CSS / JS)               │
└──────────────────────────┬─────────────────────────────┘
                           │
             ┌─────────────┴─────────────┐
             ▼                           ▼
┌───────────────────────────┐ ┌──────────────────────────┐
│   CSS Aesthetics Engine   │ │  JS Kinematics Engine    │
│  - Neon Glow (box-shadow) │ │  - Keplerian Orbits      │
│  - SoftEdge Atmospheric   │ │  - 4-Bezier Closed Loops │
│  - Glassmorphic Cards     │ │  - Constant Linear Speed │
│  - Multi-stop Gradients   │ │  - Procedural Particles  │
│  - Cyberpunk Typography   │ │  - Single-Click Advance  │
└─────────────┬─────────────┘ └──────────┬───────────────┘
              │                          │
              └─────────────┬────────────┘
                            ▼
┌────────────────────────────────────────────────────────┐
│            PowerPoint Native Vector Output             │
│        (Shape Properties + Motion Effect 86)           │
└────────────────────────────────────────────────────────┘
```

---

## 1. CSS Aesthetics to PowerPoint Vector Mapping

Modern web visual design relies on lighting, blur, shadows, and glassmorphism. PowerPoint COM and OpenXML support direct equivalents:

### A. Neon Glow & Light Emission (`box-shadow` / `filter: drop-shadow`)
In Web CSS:
```css
box-shadow: 0 0 25px #38bdf8, 0 0 50px rgba(56, 189, 248, 0.4);
```
In PowerPoint COM:
```powershell
$shape.Glow.Color.RGB = 0x24BFFB  # Neon Cyan (BGR)
$shape.Glow.Radius = 18           # Glow radius in points
$shape.Glow.Transparency = 0.35   # Semi-transparent aura
```

### B. Atmospheric Haze & Nebulae (`filter: blur()`)
In Web CSS:
```css
filter: blur(8px); opacity: 0.5;
```
In PowerPoint COM:
```powershell
$shape.SoftEdge.Type = 2          # 1=1pt, 2=2.5pt, 3=5pt, 4=10pt, 5=25pt, 6=50pt
$shape.Fill.Transparency = 0.50
```

### C. Glassmorphism HUD Panels (`backdrop-filter: blur()` & `rgba()`)
In Web CSS:
```css
background: rgba(15, 23, 42, 0.75);
backdrop-filter: blur(16px);
border: 1px solid rgba(56, 189, 248, 0.3);
```
In PowerPoint COM:
```powershell
$card.Fill.Solid()
$card.Fill.ForeColor.RGB = 0x1E120A       # Deep space navy-slate (#0A121E BGR)
$card.Fill.Transparency = 0.20           # 80% opacity glass effect
$card.Line.Visible = -1
$card.Line.ForeColor.RGB = 0x47301F       # Subtle frosted border
$card.Line.Weight = 0.8
$card.Shadow.Visible = -1                # Ambient drop shadow
$card.Shadow.Blur = 12
$card.Shadow.Transparency = 0.60
$card.Shadow.OffsetX = 0; $card.Shadow.OffsetY = 4
```

### D. Multi-Stop Radial & Linear Gradients
```powershell
# Radial solar sphere / core gradient:
$sun.Fill.TwoColorGradient(1, 1)         # 1 = msoGradientHorizontal
$sun.Fill.ForeColor.RGB = 0x0B9EF5       # Bright Gold/Orange (#F59E0B)
$sun.Fill.BackColor.RGB = 0x24BFFB       # Corona Yellow (#FBBF24)
```

---

## 2. JavaScript Kinematics & Physics Engine

### A. Keplerian Orbit Parametrization (4-Cubic Bezier Closed Circuit)
To map any elliptical orbit with semi-major axis $a$, semi-minor axis $b$, center $(c_x, c_y)$, and initial angle $\theta_0$:

Optimal cubic Bezier control constant:
$$\kappa = \frac{4}{3}(\sqrt{2} - 1) \approx 0.55228475$$

For each of the 4 quadrants ($q = 0, 1, 2, 3$):
$$\theta_A = \theta_0 + \frac{q\pi}{2}, \quad \theta_B = \theta_0 + \frac{(q+1)\pi}{2}$$

Control points relative to shape's initial position $(X_{\text{init}}, Y_{\text{init}})$ and slide dimensions ($W = 960, H = 540$):
$$C_{1x} = \frac{c_x + a\cos\theta_A - \kappa a\sin\theta_A - X_{\text{init}}}{960.0}, \quad C_{1y} = \frac{c_y + b\sin\theta_A + \kappa b\cos\theta_A - Y_{\text{init}}}{540.0}$$
$$C_{2x} = \frac{c_x + a\cos\theta_B + \kappa a\sin\theta_B - X_{\text{init}}}{960.0}, \quad C_{2y} = \frac{c_y + b\sin\theta_B - \kappa b\cos\theta_B - Y_{\text{init}}}{540.0}$$
$$E_x = \frac{c_x + a\cos\theta_B - X_{\text{init}}}{960.0}, \quad E_y = \frac{c_y + b\sin\theta_B - Y_{\text{init}}}{540.0}$$

VML String format:
```
M 0 0 C c1x c1y c2x c2y Ex Ey ... C ... 0.00000 0.00000 Z
```

### B. Golden Rules for Flawless PowerPoint Perpetual Motion
1. **Always terminate with `Z` (Closepath):** 
   - Never end motion path strings with `E` (open subpath). `Z` tells PowerPoint to treat the path as a mathematically closed loop with zero discontinuity.
2. **Eliminate Default Deceleration (`Accelerate = 0.0`, `Decelerate = 0.0`):**
   - PowerPoint defaults to `Accelerate = 0.5` and `Decelerate = 0.5`, causing objects to crawl to a dead stop at the end of each period before restarting. Setting both to `0.0` guarantees 100% constant, uniform orbital velocity.
3. **Avoid Scale Effect Conflicts on Moving Shapes:**
   - Do NOT attach concurrent `GrowShrink` (Effect 54) with `AutoReverse = -1` on shapes executing a `MotionPath` (Effect 86). Scaling reverses alter the shape's coordinate transformation matrix, causing visible position snaps. Reserve `GrowShrink` exclusively for static pulsing cores (e.g. the Sun).
4. **Set Perpetual Looping:**
   ```powershell
   $eff.Timing.RepeatCount = 9999
   $eff.Timing.RepeatDuration = 99999
   $eff.Timing.SmoothStart = 0
   $eff.Timing.SmoothEnd = 0
   $eff.Timing.BounceEnd = 0
   $eff.Timing.RewindAtEnd = 0
   ```
5. **Calibrate Harmonic Orbital Periods:**
   - Calibrate periods so that all bodies complete multiple 360° revolutions during typical slide presentation time (e.g. 1.8s to 8.2s).

### C. Single-Click Advance Architecture
In PowerPoint, active repeating animations intercept mouse clicks. To ensure users can advance slides with a single click anywhere without pausing the animation:
```powershell
# Full-Screen Transparent Overlay on the TOP layer
$overlay = $slide.Shapes.AddShape(1, 0, 0, 960, 540) # msoShapeRectangle
$overlay.Name = "ClickAdvanceOverlay"
$overlay.Fill.Solid()
$overlay.Fill.Transparency = 1.0                     # 100% invisible
$overlay.Line.Visible = 0
$overlay.ActionSettings.Item(1).Action = 1           # 1 = ppActionNextSlide
```

---

## 3. Automation Script Template (PowerShell COM)

```powershell
param([string]$pptxPath = "output_presentation.pptx")

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = 1
$pres = $ppt.Presentations.Add()
$pres.PageSetup.SlideWidth = 960.0
$pres.PageSetup.SlideHeight = 540.0

$slide = $pres.Slides.Add(1, 12) # Blank slide
$slide.Background.Fill.Solid()
$slide.Background.Fill.ForeColor.RGB = 0x120703 # #030712 deep space

# 1. Sun with CSS-inspired Corona Glow
$sun = $slide.Shapes.AddShape(9, 645 - 20, 270 - 20, 40, 40)
$sun.Fill.Solid(); $sun.Fill.ForeColor.RGB = 0x0B9EF5
$sun.Line.Visible = 0
$sun.Glow.Color.RGB = 0x24BFFB
$sun.Glow.Radius = 24
$sun.Glow.Transparency = 0.35

# 2. Moving Celestial Body
$initX = 645.0 + 150.0; $initY = 270.0
$planet = $slide.Shapes.AddShape(9, $initX - 8, $initY - 8, 16, 16)
$planet.Fill.Solid(); $planet.Fill.ForeColor.RGB = 0xF8BD38
$planet.Line.Visible = 0

# 3. Closed Bezier Motion Path (Pure Constant Velocity)
$path = "M 0 0 C 0.00000 0.06136 -0.06992 0.11111 -0.15625 0.11111 C -0.24258 0.11111 -0.31250 0.06136 -0.31250 0.00000 C -0.31250 -0.06136 -0.24258 -0.11111 -0.15625 -0.11111 C -0.06992 -0.11111 0.00000 -0.06136 0.00000 0.00000 Z"

$eff = $slide.TimeLine.MainSequence.AddEffect($planet, 86, 0, 2) # PathCircle, WithPrevious
$eff.Timing.Duration = 3.6
$eff.Timing.RepeatCount = 9999
$eff.Timing.RepeatDuration = 99999
$eff.Timing.SmoothStart = 0; $eff.Timing.SmoothEnd = 0
$eff.Timing.Accelerate = 0.0; $eff.Timing.Decelerate = 0.0
$eff.Timing.BounceEnd = 0; $eff.Timing.RewindAtEnd = 0
$eff.Behaviors.Item(1).MotionEffect.Path = $path

# 4. Single-Click Advance Overlay
$overlay = $slide.Shapes.AddShape(1, 0, 0, 960, 540)
$overlay.Fill.Solid(); $overlay.Fill.Transparency = 1.0; $overlay.Line.Visible = 0
$overlay.ActionSettings.Item(1).Action = 1

$pres.SaveAs($pptxPath)
$pres.Close()
$ppt.Quit()
```

---

## 4. Verification Checklist

- [ ] Motion path strings terminate with `Z`, not `E`.
- [ ] `Timing.Accelerate = 0.0` and `Timing.Decelerate = 0.0` set on all looping motions.
- [ ] No concurrent `GrowShrink` (Effect 54) on moving shapes executing a `MotionPath` (Effect 86).
- [ ] Full-screen transparent overlay with `Action = ppActionNextSlide (1)` present on top layer.
- [ ] Key visual elements leverage `Shape.Glow`, `Shape.SoftEdge`, and `Shape.Shadow` for modern aesthetics.
- [ ] Presentation tested via Slide Show (`F5`): continuous uninterrupted motion and 1-click slide transition verified.
