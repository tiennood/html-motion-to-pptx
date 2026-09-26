# Web3D Motion & CSS to Native Vector 3D PPTX (Agent Skill)

[![Engine](https://img.shields.io/badge/PowerPoint-SMIL%20Hardware%20Accelerated-orange.svg)]()
[![Type](https://img.shields.io/badge/Vectors-100%25%20Native-blue.svg)]()
[![Video/GIF](https://img.shields.io/badge/Video%20Embeds-0%20Bytes-success.svg)]()
[![Assets](https://img.shields.io/badge/3D%20Assets-8K%20CGI%20%2F%20NASA-purple.svg)]()
[![License](https://img.shields.io/badge/License-MIT-green.svg)]()

An advanced Agent Skill & automation pipeline that transforms **Web 3D interactive applications (Three.js, WebGL, Canvas), CSS visual styles (glassmorphism, neon glow, gradients, soft edges), and JavaScript physics/kinematics** into **100% native Microsoft PowerPoint vector shapes, 8K photorealistic 3D assets, and smooth perpetual timeline animations**.

---

## 🌟 Why This Exists: The End of Video & GIF Embeds

Converting web animations into presentation slides traditionally suffered from severe limitations:
- **Video Embeds (MP4 / WebM)**: File sizes exceed 50–200MB, playback stutter is common, and you cannot edit text, colors, or numbers.
- **Animated GIFs**: Low 256-color depth, jagged dithered edges, heavy pixelation on modern 4K/8K projectors.
- **Static Screenshots**: Completely loses all animation, dynamic lighting, and presentation impact.

### The Web3D-to-PPTX Paradigm:
| Feature | Video Embeds (MP4) | Animated GIFs | **Web3D-to-PPTX (This Skill)** |
|---|---|---|---|
| **File Size** | 50 MB – 250 MB | 20 MB – 80 MB | **< 500 KB** (or ~3.5 MB with 8K 3D assets) |
| **Resolution** | Fixed raster (1080p) | 256 colors pixelated | **Infinite Vector Crispness (4K / 8K)** |
| **Frame Rate** | Locked 30/60 FPS | Stuttery 15–25 FPS | **Smooth 60 FPS Native SMIL Hardware Acceleration** |
| **Editability** | ❌ 0% (Baked pixels) | ❌ 0% (Baked pixels) | **100% Editable Shapes, Speeds, Texts & Labels** |
| **3D Realism** | High (static video) | Low | **8K CGI/NASA Renders with Seamless Black Blending** |
| **Interactive Control** | Play / Pause only | None | **Single-Click Slide Advance & Trigger State Machines** |

---

## 🏗️ Core Architecture & Pipeline

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

## 📐 Mathematical Formulation & Orbital Kinematics

### 1. 4-Quadrant Cubic Bézier Elliptical Trajectory
PowerPoint's animation engine (`Effect 86` / `msoAnimEffectPathRight`) requires normalized VML coordinates. A mathematically perfect, continuous elliptical orbit is generated using 4 cubic Bézier splines governed by the optimal constant:
$$\kappa = \frac{4}{3}(\sqrt{2} - 1) \approx 0.5522847498$$

For semi-major axis $a$, semi-minor axis $b$, center $(c_x, c_y)$, initial position $(X_{\text{init}}, Y_{\text{init}})$, and quadrant starting angle $\theta_A$:
$$C_{1x} = \frac{c_x + a\cos\theta_A - \kappa a\sin\theta_A - X_{\text{init}}}{W}, \quad C_{1y} = \frac{c_y + b\sin\theta_A + \kappa b\cos\theta_A - Y_{\text{init}}}{H}$$
$$C_{2x} = \frac{c_x + a\cos\theta_B + \kappa a\sin\theta_B - X_{\text{init}}}{W}, \quad C_{2y} = \frac{c_y + b\sin\theta_B - \kappa b\cos\theta_B - Y_{\text{init}}}{H}$$
$$E_x = \frac{c_x + a\cos\theta_B - X_{\text{init}}}{W}, \quad E_y = \frac{c_y + b\sin\theta_B - Y_{\text{init}}}{H}$$

### 2. Perpetual Uniform Velocity (`Accelerate = 0`, `Decelerate = 0`)
PowerPoint defaults to 50% ease-in and 50% ease-out acceleration, causing orbiting planets to decelerate and stall at each loop boundary.
By setting:
```powershell
$eff.Timing.Accelerate = 0.0
$eff.Timing.Decelerate = 0.0
$eff.Timing.SmoothStart = 0
$eff.Timing.SmoothEnd = 0
```
the objects maintain **100% constant, natural Keplerian orbital velocity**.

### 3. Closed Loop Terminator (`Z`)
All motion path strings end with `Z` (closepath) rather than `E`, instructing PowerPoint to close the curve seamlessly with zero jump or coordinate snap.

---

## 🎨 Photorealistic 3D Assets & Seamless Cosmic Blending

Instead of primitive flat geometric circles, this skill introduces a **Photorealistic 3D Hybrid Architecture**:
1. **8K Studio Renders**: Planets, moons, spacecraft (Voyager 1), and stars generated from NASA/CGI imagery with real craters, atmospheric scattering, and cloud systems.
2. **Seamless Deep Black Integration**: By isolating assets on pure `#000000` black and setting the slide canvas background to `0x000000`, the asset boundaries dissolve completely into deep space.
3. **Zero Bounding Box Artifacts**: No rectangular frames, no alpha channel fringing, achieving cinematic IMAX-grade depth.
4. **Hardware Animated**: `AddPicture` shapes in PowerPoint fully support native Bézier motion paths, scale transforms, and transitions.

---

## ⚡ Maximum Single-Slide Capabilities (Triết Lý Tối Đa Hóa 1 Slide)

A single PowerPoint slide is **never** limited to passive bullet points; it can be engineered as a complete, autonomous, interactive Web 3D application:
1. **The "Single-Slide As An App" Paradigm**:
   - Instead of fragmenting a user journey across multiple boring slides, pack the entire interactive experience into **1 single master cockpit slide**.
2. **4-Layer Interactive Architecture**:
   - **Layer 1 (Perpetual Motion)**: 100+ concurrent continuous Keplerian orbits, asteroid particles, pulsating stars running smoothly at 60 FPS (`RepeatCount = 9999`, `Accelerate = 0`, `Decelerate = 0`, `Z` closepath).
   - **Layer 2 (8K Photorealistic Canvas)**: High-resolution celestial bodies isolated on `#000000` deep cosmic black with spherical bevel depth.
   - **Layer 3 (Interactive HUD Trigger Toolbar)**: Real buttons or celestial bodies mapped to `TimeLine.InteractiveSequences` (`msoAnimTriggerOnShapeClick`).
   - **Layer 4 (Bi-Directional State Machine Drawers)**:
     - Click HUD button ➔ Detailed telemetry card slides in from right (`msoAnimEffectFly`).
     - Click `✖ ĐÓNG` ➔ Card slides out to right (`Exit = -1`), returning to the pure cosmic view without changing slides.
     - 100% opaque deep-space navy cards (`Transparency = 0.0`) with glowing neon borders prevent stacked text bleed-through.

---

## 📁 Repository Structure

```
web3d-motion-to-pptx/
├── SKILL.md                          # Comprehensive Agent Skill specification
├── README.md                         # Architecture, mathematics, and documentation
├── scripts/
│   ├── build_max_single_slide.ps1    # Flagship: 1-Slide interactive cockpit app with triggers & drawers
│   ├── build_pure_web3d_deck.ps1     # Production script: 6-slide cinematic photorealistic deck
│   ├── build_motion_pptx.ps1         # Spec-driven vector motion generator
│   └── bezier_math.py                # Python Bezier ellipse & kinematics solver
├── assets/
│   └── planets/                      # 8K Photorealistic 3D planetary assets
│       ├── sun.jpg                   # Solar plasma & corona
│       ├── earth.jpg                 # Continents, clouds & atmospheric halo
│       ├── moon.jpg                  # High-res Apollo craters
│       ├── mars.jpg                  # Olympus Mons & red oxide terrain
│       ├── jupiter.jpg               # Atmospheric cloud whorls & Great Red Spot
│       ├── saturn.jpg                # 3D tilted ice rings & Cassini division
│       └── voyager.jpg               # Voyager 1 interstellar deep space probe
└── examples/
    └── solar_system_spec.json        # 6-slide JSON specification schema
```

---

## 🚀 Quick Start & Usage

### 1. As an Antigravity Agent Skill
Invoke the skill directly in your prompt:
> *"Sử dụng skill `web3d-motion-to-pptx`, hãy phân tích web 3D trong `solar_3d/` và tạo bản trình chiếu PowerPoint vector 3D tương ứng."*

### 2. Run Direct via PowerShell COM
To build the **Flagship 1-Slide Maximum Interactive Cockpit** (Perpetual 60 FPS orbits + trigger buttons + bi-directional inspection drawers):
```powershell
powershell -ExecutionPolicy Bypass -File "./scripts/build_max_single_slide.ps1" -pptxPath "./solar_system_max_single_slide.pptx"
```

To build the **6-Slide Cinematic Storytelling Deck**:
```powershell
powershell -ExecutionPolicy Bypass -File "./scripts/build_pure_web3d_deck.ps1" -pptxPath "./solar_system_3d_vector.pptx"
```

To build a custom deck from a JSON specification:
```powershell
powershell -ExecutionPolicy Bypass -File "./scripts/build_motion_pptx.ps1" -specPath "./examples/solar_system_spec.json" -outPath "./output.pptx"
```

---

## 📜 License
MIT License. Created by Tein.
