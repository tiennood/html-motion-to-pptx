---
name: web3d-motion-to-pptx
description: Converts Web 3D interactive applications (Three.js, WebGL, Canvas), CSS styles (glassmorphism, neon glow, gradients, soft edges), and JavaScript physics/kinematics into 100% native vector PowerPoint shapes and smooth perpetual timeline animations.
---

# Web3D Motion & CSS to Native Vector 3D PPTX (web3d-motion-to-pptx)

A master-grade Agent Skill and automation pipeline that converts **interactive Web 3D applications (Three.js, WebGL, HTML5 Canvas), modern CSS aesthetics (neon glow, glassmorphism, soft edges, cyber gradients), and JavaScript mathematical kinematics** into **100% native vector PowerPoint presentations (.pptx)** with silky-smooth perpetual timeline animations.

---

## 🌟 Core Philosophy: Zero Video, Zero GIF, 60 FPS Native Vector

Traditional approaches to transferring web animations into slide decks rely on screen recordings (MP4/WebM) or animated GIFs. These introduce critical flaws:
- **Bloated files** (50MB–200MB+ for short loops).
- **Severe pixelation & 256-color banding** on high-resolution 4K/retina displays.
- **Zero editability** (text, colors, numbers, and trajectories are permanently baked).
- **High latency & playback stutter** during live presentations.

**The `web3d-motion-to-pptx` paradigm**:
- **0 Bytes of Video / GIF**: Uses native PowerPoint vector shapes (`msoShapeOval`, `msoShapeRectangle`, vector groups) and hardware-accelerated SMIL animation channels.
- **Ultra-lightweight**: Complete 6-slide presentations with hundreds of continuous animated elements weigh under **500 KB** (or 3–4 MB with embedded 8K photorealistic planetary textures).
- **Infinite Resolution**: Crisp vector curves render cleanly at 4K, 8K, and beyond.
- **Hardware-Accelerated 60 FPS**: Powered directly by the native Microsoft Office graphics and animation engine.
- **Full Slide Editability**: Every text label, number, color, speed, and orbit is directly editable by the user inside standard PowerPoint.

---

## 🚨 MANDATORY PRINCIPLE: LUÔN TỐI ĐA HÓA CHO 1 SLIDE (ALWAYS MAXIMIZE A SINGLE SLIDE)

> **CHỈ THỊ BẮT BUỘC**: Tuyệt đối **không** chia nhỏ nội dung thành nhiều slide tĩnh, thụ động nếu một slide đơn có thể đảm nhiệm toàn bộ trải nghiệm tương tác! Hãy luôn đẩy 1 slide lên giới hạn năng lực tối đa của nó.

### Bản Chất: 1 Slide = 1 Ứng Dụng Tương Tác Độc Lập Hoàn Chỉnh (Cockpit Application)
Thay vì chuyển trang liên tục, hãy cấu trúc slide theo kiến trúc 4 lớp tương tác đồng thời (**4-Layer Interactive Architecture**):

1. **Lớp 1: Động cơ chuyển động liên tục 60 FPS (Perpetual Background Engine)**:
   - Toàn bộ các quỹ đạo thiên thể Keplerian Bézier, vành đai tiểu hành tinh, hiệu ứng hạt, lõi sao dao động chạy liên tục ở chế độ `RepeatCount = 9999`, `Accelerate = 0`, `Decelerate = 0`, kết thúc bằng `Z`.
2. **Lớp 2: Tài nguyên 3D quang học siêu thực (Photorealistic 3D Spatial Canvas)**:
   - Các asset 3D 8K (Mặt Trời, Trái Đất, Mặt Trăng, Sao Hỏa, Sao Mộc, Sao Thổ, Voyager 1) hòa trộn vô hình trên nền đen `#000000`, tạo độ sâu trường ảnh vũ trụ vô tận.
3. **Lớp 3: Bảng điều khiển tương tác (Interactive Cockpit Toolbar & Hotspot Triggers)**:
   - Các nút bấm HUD hoặc chính các thiên thể đang bay trên quỹ đạo được gán làm vật thể kích hoạt (Trigger Shapes) thông qua `InteractiveSequences`.
4. **Lớp 4: Máy trạng thái đóng/mở Drawer hai chiều (Bi-Directional Drawer State Machine)**:
   - **Kích hoạt Mở**: Khi nhấp vào nút HUD hoặc thiên thể ➔ Bảng thông số chi tiết (Drawer Card) trượt mượt mà từ cạnh phải vào màn hình:
     ```powershell
     $seqOpen = $slide.TimeLine.InteractiveSequences.Add(-1)
     $effOpen = $seqOpen.AddEffect($drawerGroup, 2, 0, 4) # 2 = Fly, 4 = OnShapeClick
     $effOpen.Timing.TriggerShape = $triggerButton
     $effOpen.EffectParameters.Direction = 3 # From Right
     ```
   - **Kích hoạt Đóng**: Trên mỗi Drawer có nút `✖ ĐÓNG` với hiệu ứng thoát:
     ```powershell
     $seqClose = $slide.TimeLine.InteractiveSequences.Add(-1)
     $effClose = $seqClose.AddEffect($drawerGroup, 2, 0, 4)
     $effClose.Timing.TriggerShape = $closeButton
     $effClose.Exit = -1 # msoTrue (Exit Effect)
     $effClose.EffectParameters.Direction = 4 # To Right
     ```
   - **Chống đè chữ**: Thẻ Drawer sử dụng nền xanh đen sâu thẳm đậm đặc (`Transparency = 0.0`) với viền neon phát sáng để tránh hoàn toàn hiện tượng chồng chữ khi có nhiều drawer.

Tham khảo kịch bản mẫu đầy đủ tại: `scripts/build_max_single_slide.ps1`.

---

## 🏗️ Architecture Overview

```
┌────────────────────────────────────────────────────────────────────────┐
│                   Web 3D Source (HTML / CSS / JS)                      │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │
           ┌────────────────────────┼────────────────────────┐
           ▼                        ▼                        ▼
┌──────────────────────┐ ┌──────────────────────┐ ┌──────────────────────┐
│  CSS Visual Engine   │ │ JS Kinematics Engine │ │ Photorealistic 3D    │
│  - Neon Glow         │ │  - Keplerian Orbits  │ │ Asset Pipeline       │
│  - SoftEdge Blur     │ │  - 4-Bezier Curves   │ │  - 8K NASA / CGI     │
│  - Glassmorphism     │ │  - Uniform Speed     │ │  - Pure Black Blend  │
│  - Radial Gradients  │ │  - Sub-orbits (Moon) │ │  - 3D Bevel Shading  │
│  - Minimalist Type   │ │  - Click Advance     │ │  - Zero Edge Artifact│
└──────────┬───────────┘ └──────────┬───────────┘ └──────────┬───────────┘
           │                        │                        │
           └────────────────────────┼────────────────────────┘
                                    ▼
┌────────────────────────────────────────────────────────────────────────┐
│                PowerPoint Native Vector Output Engine                  │
│       (Shapes + 3D Lighting + SMIL Motion Path 86 + Slide Master)      │
└────────────────────────────────────────────────────────────────────────┘
```

---

## 1. Photorealistic 3D Asset Pipeline & Seamless Space Blending

### The Flat Primitive Problem
Standard geometric vector primitives (`msoShapeOval`, `msoShapeRectangle`) look like basic school clipart when used to represent complex 3D objects (planets, spacecraft, intricate machinery). 

### The Solution: Deep Cosmic Black (`#000000`) Seamless Asset Integration
1. **Curate/Generate Ultra-High-Resolution 3D Renders**: Use 8K/4K photorealistic CGI or NASA space photography (e.g. turbulent solar plasma flares, Earth's atmospheric glow with swirling cloud weather systems, Mars crater geology, Saturn's 3D tilted ice rings).
2. **Isolate on Pure Black Background (`#000000`)**: Ensure the asset's outer boundaries fade to absolute black.
3. **Pure Black Canvas Foundation**: Set both the slide background and the base canvas rectangle to pure black:
   ```powershell
   $slide.Background.Fill.Solid()
   $slide.Background.Fill.ForeColor.RGB = 0x000000
   
   $bg = $slide.Shapes.AddShape(1, 0, 0, 960, 540) # msoShapeRectangle
   $bg.Fill.Solid()
   $bg.Fill.ForeColor.RGB = 0x000000
   $bg.Line.Visible = 0
   $bg.ZOrder(1) # Send to back
   ```
4. **Seamless Dissolve**: When placed on the slide, the black boundaries of the 3D assets dissolve 100% invisibly into deep space. There are **zero rectangular bounding box artifacts**, achieving cinematic depth without fragile PNG alpha masks.
5. **Full Animation Support**: `AddPicture` shapes in PowerPoint fully support native animation timelines, including Bézier motion paths (`Effect 86`), spinning, scaling, and entrance transitions.

### PowerPoint COM 3D Properties: Critical Rules & Gotchas
- **Spherical Beveling**: To give flat vector circles genuine spherical depth and light shading:
  ```powershell
  $planet.ThreeD.BevelTopType = 6  # 6 = msoBevelCircle
  $planet.ThreeD.BevelTopInset = $radius / 2.0
  $planet.ThreeD.BevelTopDepth = $radius / 2.0
  ```
- **⚠️ Runtime Exception Warning**: Calling `$shape.ThreeD.Material` throws a `PropertyNotFound` or COM method exception. Never set `.Material` directly. Use `.PresetMaterial` or rely on the bevel geometry and lighting.

---

## 2. CSS Visual Effects to Native Vector Style Mapping

Modern web apps use lighting, blur, semi-transparency, and glassmorphism. Translate them directly to native PowerPoint shape properties:

| Web CSS Property | PowerPoint COM Equivalent | Code Implementation |
|---|---|---|
| **Neon Light Glow** (`box-shadow: 0 0 25px #38bdf8`) | `Shape.Glow` | `$s.Glow.Color.RGB = 0x24BFFB`<br>`$s.Glow.Radius = 18`<br>`$s.Glow.Transparency = 0.35` |
| **Atmospheric Blur** (`filter: blur(8px)`) | `Shape.SoftEdge` | `$s.SoftEdge.Type = 2` (1=1pt, 2=2.5pt, 3=5pt, 4=10pt, 5=25pt, 6=50pt)<br>`$s.Fill.Transparency = 0.50` |
| **Glassmorphism Card** (`backdrop-filter: blur()`, `rgba()`) | `Shape.Fill` + `Shape.Line` + `Shape.Shadow` | `$c.Fill.Solid()`<br>`$c.Fill.ForeColor.RGB = 0x1E120A` (Slate Navy BGR)<br>`$c.Fill.Transparency = 0.20`<br>`$c.Line.ForeColor.RGB = 0x47301F`<br>`$c.Shadow.Blur = 12; $c.Shadow.Transparency = 0.60` |
| **Radial Core Gradient** (`radial-gradient(...)`) | `Shape.Fill.TwoColorGradient` | `$s.Fill.TwoColorGradient(1, 1)`<br>`$s.Fill.ForeColor.RGB = 0x0B9EF5`<br>`$s.Fill.BackColor.RGB = 0x24BFFB` |
| **Starfield Particles** (`canvas.drawCircle`) | Procedural Vector Circles | Loop 60–100 micro-circles (`size: 0.8–2.0pt`), randomized pastel cosmic colors, zero border. |

---

## 3. JavaScript Kinematics & Keplerian Orbital Math

### 4-Quadrant Cubic Bézier Parametrization
PowerPoint motion paths (`msoAnimEffectPathRight` or `Effect 86`) accept normalized VML string paths relative to the shape's initial position $(X_{\text{init}}, Y_{\text{init}})$ and slide canvas dimensions ($W = 960, H = 540$).

To represent an ellipse with semi-major axis $a$, semi-minor axis $b$, center $(c_x, c_y)$, and starting angle $\theta_0$:

1. **Optimal Bezier Distance Constant**:
   $$\kappa = \frac{4}{3}(\sqrt{2} - 1) \approx 0.5522847498$$
2. **For Each Quadrant $q \in \{0, 1, 2, 3\}$**:
   $$\theta_A = \theta_0 + \frac{q\pi}{2}, \quad \theta_B = \theta_0 + \frac{(q+1)\pi}{2}$$
   Control point 1:
   $$C_{1x} = \frac{c_x + a\cos\theta_A - \kappa a\sin\theta_A - X_{\text{init}}}{W}, \quad C_{1y} = \frac{c_y + b\sin\theta_A + \kappa b\cos\theta_A - Y_{\text{init}}}{H}$$
   Control point 2:
   $$C_{2x} = \frac{c_x + a\cos\theta_B + \kappa a\sin\theta_B - X_{\text{init}}}{W}, \quad C_{2y} = \frac{c_y + b\sin\theta_B - \kappa b\cos\theta_B - Y_{\text{init}}}{H}$$
   Endpoint:
   $$E_x = \frac{c_x + a\cos\theta_B - X_{\text{init}}}{W}, \quad E_y = \frac{c_y + b\sin\theta_B - Y_{\text{init}}}{H}$$
3. **Format as VML String**:
   ```
   M 0 0 C c1x c1y c2x c2y Ex Ey ... C ... 0.00000 0.00000 Z
   ```

### ⚡ Golden Rules for 100% Smooth Perpetual Motion
1. **Always Terminate with `Z` (Closepath)**:
   - Ending with `E` leaves the path open, causing PowerPoint to reset coordinate origins or produce micro-stutters. `Z` instructs the SMIL engine to treat the trajectory as a closed continuous loop.
2. **Eliminate Default Deceleration (`Accelerate = 0.0`, `Decelerate = 0.0`)**:
   - PowerPoint defaults to 50% ease-in and 50% ease-out, making celestial bodies slow down and stall at each loop boundary. Explicitly setting both to `0.0` ensures **pure constant linear velocity**.
3. **Do NOT Mix `GrowShrink` AutoReverse on Moving Shapes**:
   - Applying `GrowShrink` (Effect 54) with `AutoReverse = -1` while an object is traversing a Motion Path (Effect 86) causes coordinate matrix transformation conflicts, resulting in visible shape snaps. Reserve scaling exclusively for stationary pulsing objects (e.g. the Sun).
4. **Hierarchical / Sub-Orbits (e.g. Moon around Earth)**:
   - For secondary bodies, apply the Keplerian path centered on the host planet's offset, or in zoomed comparison slides, anchor an elliptical orbit track around the primary hero body.

---

## 4. Color Science: Web Hex (`#RRGGBB`) to Windows COM BGR Integer

Windows COM `.Color.RGB` expects a 32-bit integer encoded in **little-endian BGR** (`(Blue << 16) | (Green << 8) | Red`). Directly passing standard RGB hex (`0x38BDF8`) swaps Red and Blue, turning bright cyan into orange.

Always use the standardized converter function:
```powershell
function Color-Hex([string]$hex) {
    $hex = $hex.TrimStart('#')
    $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return [int]($r + ($g * 256) + ($b * 65536))
}
```

---

## 5. Interaction & Navigation Model: Single-Click Advance Overlay

### The Problem
When a PowerPoint slide contains perpetual looping animations (`RepeatCount = 9999`), clicking anywhere on the screen during a slide show often targets the animating elements or timeline events, causing mouse clicks to be "swallowed" rather than advancing to the next slide.

### The Solution: Full-Screen Action Overlay
Place a 100% transparent vector rectangle covering the entire canvas on the top layer:
```powershell
$clickOverlay = $slide.Shapes.AddShape(1, 0, 0, 960, 540) # msoShapeRectangle
$clickOverlay.Fill.Transparency = 1.0                    # Completely invisible
$clickOverlay.Line.Visible = 0
$clickOverlay.ActionSettings.Item(1).Action = 1           # 1 = ppActionNextSlide

# Ensure slide show transition advances on click
$slide.SlideShowTransition.AdvanceOnClick = -1
$slide.SlideShowTransition.Duration = 0.8
```
This guarantees that **a single left-click anywhere on the screen instantly transitions to the next slide**, while all animations continue running smoothly in the background.

---

## 6. Web 3D vs PowerPoint Compatibility Matrix

| Web 3D / CSS Feature | PowerPoint Support | Strategy / Workaround |
|---|---|---|
| **Keplerian Orbits / Splines** | **100% Native** | Pre-compute 4-quadrant cubic Bézier curves with constant velocity (`Accelerate = 0`). |
| **CSS Glow & Soft Edges** | **100% Native** | Use `Shape.Glow` and `Shape.SoftEdge`. |
| **Glassmorphism Panels** | **100% Native** | Slate-navy fill + 80% opacity + 0.8pt border + subtle drop shadow. |
| **Photorealistic 3D Textures** | **100% Native** | NASA/CGI 8K renders blended onto `#000000` deep cosmic black canvas. |
| **Native 3D Models (.glb / .gltf)** | **100% Native** | PowerPoint 365/2019+ natively imports 3D models with 360° turntable rotation. |
| **Custom GLSL Shaders** | ❌ Not Supported | Pre-render shader effects into clean 3D graphic assets or procedural vector gradients. |
| **Mouse Orbit Camera Controls** | ❌ Not Supported | Emulate dynamic camera angles across sequential slides using smooth Morph transitions. |
| **Dynamic N-body Physics** | ❌ Not Supported | Pre-calculate deterministic trajectories into Bézier animation paths. |

---

## 7. Maximum Capabilities of a Single Slide

A single PowerPoint slide can function as a comprehensive interactive application:

1. **Concurrent Motion Channels**: Supports **100+ simultaneous independent animation timelines** running at 60 FPS without frame drops.
2. **Interactive Trigger State Machines (`Animation.TriggerShape`)**:
   - Shapes can act as interactive buttons without changing slides.
   - Clicking Planet A triggers an entrance animation for Drawer A, plays an audio briefing, and zooms in on telemetry data.
   - Clicking a "Close" button reverses the timeline.
3. **True 3D Model Manipulation (`.glb` / `.gltf`)**:
   - Full 3-axis rotation ($X, Y, Z$), camera zoom, and Turntable animation (`msoAnimEffectTurntable`).
4. **Synchronized Multimedia**: Audio voiceovers and sound effects synchronized to animation bookmarks.
5. **VBA Scripting & Event Loops**: For enterprise workstations, embedded VBA macros allow real-time sliders, math calculators, and state storage.

---

## 8. Minimalist Cinematic Design vs HUD Clutter

When transforming an interactive Web 3D page into a presentation, **avoid cluttering slides with simulated 2D HUD boxes, radar frames, and floating pill buttons** unless specifically requested.

**The Golden Aesthetic Standard**:
- **Clean Cosmic Canvas**: Give 75% of the slide space to the 3D celestial bodies and their glowing orbits.
- **Cinematic Typography**:
  - `Step Tag`: Subtle, uppercase tracking with glowing dot (e.g. `TRANG 01 / 06 • TOÀN CẢNH VŨ TRỤ`, Font size: 8–9pt, Amber/Cyan).
  - `Slide Title`: Bold, crisp heading (`Montserrat`, `Orbitron`, or `Outfit`, 20–24pt, Pure White).
  - `Description`: High-contrast, clean summary text (9–10pt, Light Slate `#94A3B8`).
- **Focus on Motion & Realism**: Let the high-res 3D assets, planetary rings, orbiting moons, and asteroid belts captivate the audience.

---

## 9. Production PowerShell COM Script Template

```powershell
param(
    [string]$pptxPath = (Join-Path (Get-Location) "web3d_output.pptx")
)

# 1. Close any running PowerPoint instances safely
Stop-Process -Name POWERPNT -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 500

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = 1
$pres = $ppt.Presentations.Add()
$pres.PageSetup.SlideWidth = 960.0
$pres.PageSetup.SlideHeight = 540.0

$W = 960.0; $H = 540.0
$ci = [System.Globalization.CultureInfo]::InvariantCulture
$kappa = (4.0 / 3.0) * ([Math]::Sqrt(2.0) - 1.0) # 0.5522847498

function Color-Hex([string]$hex) {
    $hex = $hex.TrimStart('#')
    $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return [int]($r + ($g * 256) + ($b * 65536))
}

function Get-BezierEllipsePath($cx, $cy, $a, $b, $initX, $initY, $theta0) {
    $quarters = @()
    for ($q = 0; $q -lt 4; $q++) {
        $thA = $theta0 + ($q * [Math]::PI / 2.0)
        $thB = $theta0 + (($q + 1) * [Math]::PI / 2.0)

        $c1x = ($cx + $a * [Math]::Cos($thA) - $kappa * $a * [Math]::Sin($thA) - $initX) / $W
        $c1y = ($cy + $b * [Math]::Sin($thA) + $kappa * $b * [Math]::Cos($thA) - $initY) / $H
        $c2x = ($cx + $a * [Math]::Cos($thB) + $kappa * $a * [Math]::Sin($thB) - $initX) / $W
        $c2y = ($cy + $b * [Math]::Sin($thB) - $kappa * $b * [Math]::Cos($thB) - $initY) / $H
        $endX = ($cx + $a * [Math]::Cos($thB) - $initX) / $W
        $endY = ($cy + $b * [Math]::Sin($thB) - $initY) / $H

        $quarters += [string]::Format($ci, "C {0:F5} {1:F5} {2:F5} {3:F5} {4:F5} {5:F5}", $c1x, $c1y, $c2x, $c2y, $endX, $endY)
    }
    return "M 0 0 " + ($quarters -join " ") + " Z"
}

# Create Slide
$slide = $pres.Slides.Add(1, 12) # Blank
$slide.Background.Fill.Solid()
$slide.Background.Fill.ForeColor.RGB = 0x000000 # Pure cosmic black

# Canvas Base
$bg = $slide.Shapes.AddShape(1, 0, 0, $W, $H)
$bg.Fill.Solid(); $bg.Fill.ForeColor.RGB = 0x000000; $bg.Line.Visible = 0

# Central Star
$cx = 580.0; $cy = 280.0
$sun = $slide.Shapes.AddShape(9, $cx - 24, $cy - 24, 48, 48)
$sun.Fill.Solid(); $sun.Fill.ForeColor.RGB = Color-Hex "F59E0B"
$sun.Line.Visible = 0
$sun.Glow.Color.RGB = Color-Hex "FBBF24"
$sun.Glow.Radius = 24; $sun.Glow.Transparency = 0.3

# Orbiting Planet
$a = 160.0; $b = 70.0; $r = 14.0; $th0 = 0.8
$initX = $cx + $a * [Math]::Cos($th0)
$initY = $cy + $b * [Math]::Sin($th0)
$planet = $slide.Shapes.AddShape(9, $initX - ($r / 2), $initY - ($r / 2), $r, $r)
$planet.Fill.Solid(); $planet.Fill.ForeColor.RGB = Color-Hex "38BDF8"
$planet.Line.Visible = 0
$planet.ThreeD.BevelTopType = 6
$planet.ThreeD.BevelTopInset = $r / 2; $planet.ThreeD.BevelTopDepth = $r / 2

# Orbit Track
$orb = $slide.Shapes.AddShape(9, $cx - $a, $cy - $b, $a * 2, $b * 2)
$orb.Fill.Visible = 0; $orb.Line.Visible = -1
$orb.Line.ForeColor.RGB = Color-Hex "38BDF8"; $orb.Line.Transparency = 0.82

# Motion Animation
$eff = $slide.TimeLine.MainSequence.AddEffect($planet, 86, 0, 2)
$eff.Timing.Duration = 18.0
$eff.Timing.RepeatCount = 9999
$eff.Timing.RepeatDuration = 99999
$eff.Timing.SmoothStart = 0; $eff.Timing.SmoothEnd = 0
$eff.Timing.Accelerate = 0.0; $eff.Timing.Decelerate = 0.0
$eff.Behaviors.Item(1).MotionEffect.Path = Get-BezierEllipsePath $cx $cy $a $b $initX $initY $th0

# Single-Click Advance Overlay
$overlay = $slide.Shapes.AddShape(1, 0, 0, $W, $H)
$overlay.Fill.Transparency = 1.0; $overlay.Line.Visible = 0
$overlay.ActionSettings.Item(1).Action = 1
$slide.SlideShowTransition.AdvanceOnClick = -1

$pres.SaveAs($pptxPath)
$pres.Close()
$ppt.Quit()
```

---

## 10. Verification & Quality Checklist

Before delivering presentations converted with this skill, verify:
- [ ] **Motion path strings terminate with `Z`** (closed loop), never `E`.
- [ ] **`Timing.Accelerate = 0.0` and `Timing.Decelerate = 0.0`** applied to all continuous orbital paths.
- [ ] **No `GrowShrink` (Effect 54) AutoReverse attached to shapes traversing a Motion Path (Effect 86)**.
- [ ] **All color values mapped through `Color-Hex`** (ensuring correct BGR translation in COM).
- [ ] **Slide background and backing canvas set to `0x000000`** to seamlessly dissolve 3D asset boundaries into space.
- [ ] **No calls to `$shape.ThreeD.Material`** (use `.PresetMaterial` or rely on `.BevelTopType = 6`).
- [ ] **Single-click advance overlay (`ppActionNextSlide`)** placed on the top layer of each slide.
- [ ] **Tested in Slide Show mode (`F5`)**: Continuous motion is uninterrupted, 60 FPS, and single-click advances slides instantly.
