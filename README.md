# HTML Motion to Native Vector 3D PPTX (Agent Skill)

An advanced Agent Skill & automation pipeline that transforms dynamic web animations (HTML5, CSS keyframes, Canvas, WebGL, SVG) into **native Microsoft PowerPoint 2D vector shapes with 3D perspective orbital motion paths, depth scaling, and multi-layer Z-ordering**.

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
- **Editable & Crisp**: Every card, statistic, planet, and trajectory can be edited directly inside PowerPoint.

---

## 📐 Mathematical Formulation

### 1. 4-Arc Cubic Bezier Closed Elliptical Trajectory
PowerPoint's animation path syntax (`msoAnimEffectPath`) requires normalized coordinates. A smooth, mathematically closed ellipse is generated using 4 cubic Bezier curves with the magic constant:
$$\kappa = \frac{4}{3}(\sqrt{2} - 1) \approx 0.5522847498$$

### 2. Pseudo-3D Depth Scaling
By attaching an `msoAnimEffectScale` with `AutoReverse = -1` (True) synchronized with the orbital period, objects scale from $80\%$ (at the far side behind the focal point) to $125\%$ (in the foreground), creating a realistic sense of 3D perspective depth.

---

## 📁 Repository Structure
```
html-motion-to-pptx/
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
Place this folder into `.agents/skills/html-motion-to-pptx/` in your workspace root, or `~/.gemini/config/skills/html-motion-to-pptx/` for global availability.

Activate the skill by asking:
> *"Convert the animations in `index.html` into a native 2D vector 3D PPTX presentation."*

### Manual Execution (PowerShell COM)
```powershell
powershell -ExecutionPolicy Bypass -File "./scripts/build_motion_pptx.ps1" -specPath "./examples/solar_system_spec.json" -outPath "./output.pptx"
```

---

## 📜 License
MIT License. Created by Tein.
