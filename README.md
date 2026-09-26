# Web3D Motion & CSS to Native Vector 3D PPTX (Agent Skill)

An advanced Agent Skill & automation pipeline that transforms **Web 3D interactive applications (Three.js, WebGL, Canvas), CSS visual styles (glassmorphism, neon glow, gradients, soft edges), and JavaScript physics/kinematics** into **100% native Microsoft PowerPoint vector shapes and smooth perpetual timeline animations**.

---

## 🌟 Why This Exists
Converting web animations to PowerPoint traditionally suffered from major limitations:
- **Video embeds (MP4 / WebM)**: Huge file sizes (50MB+), choppy playback, cannot edit text or shapes.
- **Animated GIFs**: Low color depth (256 colors), jagged edges, heavy pixelation on 4K projectors.
- **Static Screenshots**: Loses all animation and interactivity completely.

**This skill introduces a new paradigm**:
- **0 Bytes of Video / GIF**: 100% native vector shapes (`msoShapeOval`, `msoShapeRectangle`, grouped vectors).
- **Sub-100KB presentation decks**: Fast loading, instant sharing via email.
- **Continuous 60 FPS**: Rendered via PowerPoint's native hardware-accelerated SMIL animation engine.
- **CSS Aesthetics**: Neon Glow (`Shape.Glow`), Glassmorphism transparency & ambient shadows (`Shape.Shadow`), Soft atmospheric edges (`Shape.SoftEdge`), and Cyberpunk typography.
- **Constant Linear Velocity**: Eliminates default 50% ease-in/ease-out hãm tốc by setting `Accelerate = 0` and `Decelerate = 0`.
- **Closed Circuit Orbits**: Terminated with `Z` for seamless, perpetual 360° loops without jumping or stopping.
- **1-Click Slide Advance**: Full-screen transparent overlay ensures 1 click immediately transitions to the next slide.
- **Editable & Crisp**: Every card, statistic, planet, and trajectory can be edited directly inside PowerPoint.

---

## 📐 Mathematical Formulation

### 1. 4-Arc Cubic Bezier Closed Elliptical Trajectory
PowerPoint's animation path syntax (`msoAnimEffectPath`) requires normalized coordinates. A smooth, mathematically closed ellipse is generated using 4 cubic Bezier curves with the constant:
$$\kappa = \frac{4}{3}(\sqrt{2} - 1) \approx 0.5522847498$$

Always terminating with `Z` for continuous seamless repetition.

### 2. Linear Velocity Tuning
By setting `Timing.Accelerate = 0.0` and `Timing.Decelerate = 0.0`, shapes orbit with uniform celestial speed instead of stalling at the loop boundary.

---

## 📁 Repository Structure
```
web3d-motion-to-pptx/
├── SKILL.md                          # Antigravity Agent Skill definition
├── README.md                         # Documentation & Architecture
├── scripts/
│   ├── bezier_math.py                # Python Bezier ellipse & projection calculations
│   └── build_motion_pptx.ps1         # PowerShell COM deck generator
└── examples/
    └── solar_system_spec.json        # 6-slide celestial system specification
```

---

## 🚀 Installation & Usage

### As an Antigravity Agent Skill
Place this folder into `.agents/skills/web3d-motion-to-pptx/` in your workspace root, or `~/.gemini/config/skills/web3d-motion-to-pptx/` for global availability.

Activate the skill by asking:
> *"Sử dụng skill web3d-motion-to-pptx, hãy phân tích các chuyển động và CSS trong file index.html và chuyển đổi thành slide PowerPoint vector 3D tương ứng."*

### Manual Execution (PowerShell COM)
```powershell
powershell -ExecutionPolicy Bypass -File "./scripts/build_motion_pptx.ps1" -specPath "./examples/solar_system_spec.json" -outPath "./output.pptx"
```

---

## 📜 License
MIT License. Created by Tein.
