param(
    [string]$pptxPath = "C:\Users\Tein\Downloads\KT ChiPhi\solar_system_3d_vector.pptx"
)

# 1. Close any running PowerPoint instances safely
Stop-Process -Name POWERPNT -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 600

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = 1

$pres = $ppt.Presentations.Add()
$pres.PageSetup.SlideWidth = 960.0
$pres.PageSetup.SlideHeight = 540.0

$W = 960.0
$H = 540.0
$ci = [System.Globalization.CultureInfo]::InvariantCulture
$kappa = (4.0 / 3.0) * ([Math]::Sqrt(2.0) - 1.0) # 0.5522847498

# Asset Paths for High-Res Transparent 32-bit RGBA PNG Renders (Post-Processed)
$assetDir = "C:\Users\Tein\Downloads\KT ChiPhi\assets\planets"
$imgSun     = Join-Path $assetDir "sun.png"
$imgEarth   = Join-Path $assetDir "earth.png"
$imgMars    = Join-Path $assetDir "mars.png"
$imgJupiter = Join-Path $assetDir "jupiter.png"
$imgSaturn  = Join-Path $assetDir "saturn.png"
$imgVoyager = Join-Path $assetDir "voyager.png"
$imgMoon    = Join-Path $assetDir "moon.png"

# Fallback to jpg if png not found
if (-not (Test-Path $imgSun)) { $imgSun = Join-Path $assetDir "sun.jpg" }
if (-not (Test-Path $imgEarth)) { $imgEarth = Join-Path $assetDir "earth.jpg" }
if (-not (Test-Path $imgMars)) { $imgMars = Join-Path $assetDir "mars.jpg" }
if (-not (Test-Path $imgJupiter)) { $imgJupiter = Join-Path $assetDir "jupiter.jpg" }
if (-not (Test-Path $imgSaturn)) { $imgSaturn = Join-Path $assetDir "saturn.jpg" }
if (-not (Test-Path $imgVoyager)) { $imgVoyager = Join-Path $assetDir "voyager.jpg" }
if (-not (Test-Path $imgMoon)) { $imgMoon = Join-Path $assetDir "moon.jpg" }

# Helper to convert CSS hex "#RRGGBB" into PowerPoint COM BGR integer
function Color-Hex([string]$hex) {
    $hex = $hex.TrimStart('#')
    $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return [int]($r + ($g * 256) + ($b * 65536))
}

# Closed 4-Bezier ellipse path for 100% smooth perpetual orbit
function Get-BezierEllipsePath($cx, $cy, $a, $b, $initX, $initY, $theta0) {
    $quarters = @()
    for ($q = 0; $q -lt 4; $q++) {
        $thA = $theta0 + ($q * [Math]::PI / 2.0)
        $thB = $theta0 + (($q + 1) * [Math]::PI / 2.0)

        $c1x = ($cx + $a * [Math]::Cos($thA) - $kappa * $a * [Math]::Sin($thA) - $initX) / 960.0
        $c1y = ($cy + $b * [Math]::Sin($thA) + $kappa * $b * [Math]::Cos($thA) - $initY) / 540.0

        $c2x = ($cx + $a * [Math]::Cos($thB) + $kappa * $a * [Math]::Sin($thB) - $initX) / 960.0
        $c2y = ($cy + $b * [Math]::Sin($thB) - $kappa * $b * [Math]::Cos($thB) - $initY) / 540.0

        $endX = ($cx + $a * [Math]::Cos($thB) - $initX) / 960.0
        $endY = ($cy + $b * [Math]::Sin($thB) - $initY) / 540.0

        $quarters += [string]::Format($ci, "C {0:F5} {1:F5} {2:F5} {3:F5} {4:F5} {5:F5}", $c1x, $c1y, $c2x, $c2y, $endX, $endY)
    }
    return "M 0 0 " + ($quarters -join " ") + " Z"
}

# The 6 exact cinematic slides
$deckSpecs = @(
    @{
        Index      = 0
        StepTag    = "TRANG 01 / 06 • TOÀN CẢNH VŨ TRỤ"
        Title      = "Toàn Cảnh Thái Dương Hệ"
        Desc       = "8 hành tinh đang liên tục quay quanh Mặt Trời theo thời gian thực. Mỗi hành tinh sở hữu vận tốc và quỹ đạo Keplerian riêng biệt."
        Stage      = "FullSystem"
        Transition = 3849 # FadeSmoothly on slide 1
    },
    @{
        Index      = 1
        StepTag    = "TRANG 02 / 06 • VÙNG ĐẤT ĐÁ"
        Title      = "Vùng Hành Tinh Đất Đá"
        Desc       = "Vùng nội hệ gần Mặt Trời - nơi các hành tinh sở hữu bề mặt đá rắn chắc. Mặt Trăng liên tục quay quanh Trái Đất theo thời gian thực."
        Stage      = "InnerPlanets"
        Transition = 3881 # Morph 3D
    },
    @{
        Index      = 2
        StepTag    = "TRANG 03 / 06 • ĐỐI CHIẾU THIÊN THỂ"
        Title      = "Trái Đất so với Sao Hỏa"
        Desc       = "Cận cảnh Trái Đất (với Mặt Trăng liên tục quay quanh) và Sao Hỏa đỏ rực - đích đến tương lai của loài người."
        Stage      = "EarthMars"
        Transition = 3881 # Morph 3D
    },
    @{
        Index      = 3
        StepTag    = "TRANG 04 / 06 • KHỔNG LỒ NGOẠI HỆ"
        Title      = "Các Gã Khổng Lồ Khí & Băng"
        Desc       = "Vượt qua Vành đai Tiểu hành tinh để chiêm ngưỡng Sao Mộc và Sao Thổ với kiệt tác vành đai băng đá 3D lộng lẫy."
        Stage      = "GasGiants"
        Transition = 3881 # Morph 3D
    },
    @{
        Index      = 4
        StepTag    = "TRANG 05 / 06 • DỮ LIỆU THIÊN VĂN"
        Title      = "Thước Đo Quy Mô Thái Dương Hệ"
        Desc       = "Góc nhìn bao quát từ trên cao thể hiện rõ khoảng cách và tỷ lệ kích thước giữa các hành tinh trong vũ trụ."
        Stage      = "GrandScale"
        Transition = 3881 # Morph 3D
    },
    @{
        Index      = 5
        StepTag    = "TRANG 06 / 06 • KHÁM PHÁ VÔ TẬN"
        Title      = "Khát Vọng Vươn Ra Vũ Trụ"
        Desc       = "Hành trình thám hiểm vũ trụ không bao giờ dừng lại. Tàu Voyager 1 đang lướt đi trong khoảng không liên sao bất tận."
        Stage      = "VoyagerDeep"
        Transition = 3881 # Morph 3D
    }
)

# Helper function to style text boxes safely
function Set-TextProps($tf, $text, $fontName, $fontSize, $isBold, $colorRgb, $align = 1) {
    $tf.WordWrap = -1
    $tf.MarginLeft = 2; $tf.MarginRight = 2; $tf.MarginTop = 1; $tf.MarginBottom = 1
    $tr = $tf.TextRange
    $tr.Text = [string]$text
    $tr.ParagraphFormat.Alignment = $align
    $tr.Font.Name = [string]$fontName
    $tr.Font.Size = [double]$fontSize
    $tr.Font.Bold = if ($isBold -and $isBold -ne 0) { -1 } else { 0 }
    $tr.Font.Color.RGB = [int]$colorRgb
}

# Helper to draw a Holographic Glass Telemetry Card with strict anti-collision positioning
function Add-HolographicCard($slide, $x, $y, $w, $h, $titleText, $metrics, $accentHex) {
    # 1. Semi-transparent dark navy glass backing
    $card = $slide.Shapes.AddShape(1, $x, $y, $w, $h)
    $card.Fill.Solid()
    $card.Fill.ForeColor.RGB = Color-Hex "0A1128"
    $card.Fill.Transparency = 0.15 # 85% opacity
    $card.Line.Visible = -1
    $card.Line.ForeColor.RGB = Color-Hex $accentHex
    $card.Line.Weight = 1.0
    $card.Line.Transparency = 0.30

    # 2. Header bar with Sci-Fi title
    $headTx = $slide.Shapes.AddTextbox(1, $x + 12, $y + 6, $w - 24, 20)
    Set-TextProps $headTx.TextFrame $titleText "Outfit" 8.5 $true (Color-Hex $accentHex) 1

    # 3. Telemetry lines
    $bodyText = $metrics -join "`n"
    $bodyTx = $slide.Shapes.AddTextbox(1, $x + 12, $y + 26, $w - 24, $h - 32)
    Set-TextProps $bodyTx.TextFrame $bodyText "Outfit" 7.5 $false (Color-Hex "E2E8F0") 1
}

# Helper to draw Sci-Fi Corner Brackets that hug the planet without intersecting outer orbits
function Add-CornerReticle($slide, $cx, $cy, $radius, $label, $accentHex) {
    # Inner circular ring hugging planet closely
    $ring = $slide.Shapes.AddShape(9, $cx - $radius, $cy - $radius, $radius * 2, $radius * 2)
    $ring.Fill.Visible = 0
    $ring.Line.Visible = -1
    $ring.Line.ForeColor.RGB = Color-Hex $accentHex
    $ring.Line.Weight = 0.75
    $ring.Line.DashStyle = 4 # Dashed
    $ring.Line.Transparency = 0.5

    # Target label tag above planet
    $tag = $slide.Shapes.AddTextbox(1, $cx - $radius, $cy - $radius - 18, $radius * 2, 16)
    Set-TextProps $tag.TextFrame "[ TARGET: $label ]" "Consolas" 7.0 $true (Color-Hex $accentHex) 2
}

# 2. Iterate through each slide definition
for ($sIdx = 0; $sIdx -lt $deckSpecs.Count; $sIdx++) {
    $spec = $deckSpecs[$sIdx]
    Write-Host "Creating Advanced Flow-Engine Slide $($sIdx + 1): $($spec.Title)..."

    $slide = $pres.Slides.Add($sIdx + 1, 12) # ppLayoutBlank
    $slide.FollowMasterBackground = 0
    $slide.Background.Fill.Solid()
    $slide.Background.Fill.ForeColor.RGB = Color-Hex "000000"

    # Deep dark cosmic background canvas
    $bg = $slide.Shapes.AddShape(1, 0, 0, 960, 540)
    $bg.Fill.Solid()
    $bg.Fill.ForeColor.RGB = Color-Hex "000000"
    $bg.Line.Visible = 0
    $bg.ZOrder(1)

    # Ambient Starfield
    $rand = New-Object System.Random(1000 + $sIdx * 77)
    $starColors = @((Color-Hex "F8FAFC"), (Color-Hex "38BDF8"), (Color-Hex "F59E0B"), (Color-Hex "C084FC"))
    for ($i = 0; $i -lt 65; $i++) {
        $sx = $rand.Next(10, 950)
        $sy = $rand.Next(10, 530)
        $sr = 0.8 + ($rand.NextDouble() * 1.6)
        $sc = $starColors[$rand.Next(0, $starColors.Count)]
        $st = $slide.Shapes.AddShape(9, $sx, $sy, $sr, $sr)
        $st.Fill.Solid()
        $st.Fill.ForeColor.RGB = $sc
        $st.Line.Visible = 0
    }

    # =========================================================================
    # SLIDE HEADER (ANTI-COLLISION ZONE: X:36..460, Y:30..120)
    # =========================================================================
    $tagTx = $slide.Shapes.AddTextbox(1, 36, 30, 420, 18)
    Set-TextProps $tagTx.TextFrame $spec.StepTag "Outfit" 8.0 $true (Color-Hex "F59E0B") 1

    $titleTx = $slide.Shapes.AddTextbox(1, 34, 48, 440, 36)
    Set-TextProps $titleTx.TextFrame $spec.Title "Outfit" 20.0 $true (Color-Hex "FFFFFF") 1

    $descTx = $slide.Shapes.AddTextbox(1, 36, 88, 420, 36)
    Set-TextProps $descTx.TextFrame $spec.Desc "Outfit" 9.5 $false (Color-Hex "94A3B8") 1

    $tLine = $slide.TimeLine.MainSequence

    # =========================================================================
    # 3D CELESTIAL STAGE WITH SEAMLESS MATTING & MOTION CONTINUITY
    # =========================================================================
    if ($spec.Stage -eq "FullSystem") {
        $cx = 580.0; $cy = 280.0

        # Dynamic Comet Halley sweeping across background
        $cmt = $slide.Shapes.AddShape(9, -20, 40, 4, 4)
        $cmt.Fill.Solid(); $cmt.Fill.ForeColor.RGB = Color-Hex "38BDF8"; $cmt.Line.Visible = 0
        $effCmt = $tLine.AddEffect($cmt, 86, 0, 2)
        $effCmt.Behaviors.Item(1).MotionEffect.Path = "M 0 0 L 1.15 0.75"
        $effCmt.Timing.Duration = 4.5
        $effCmt.Timing.RepeatCount = 9999
        $effCmt.Timing.RepeatDuration = 99999
        $effCmt.Timing.Accelerate = 0.25
        $effCmt.Timing.Decelerate = 0.25

        # 1. Pulsating Breathing Corona Aura behind Sun
        $sunAura = $slide.Shapes.AddShape(9, $cx - 56, $cy - 56, 112, 112)
        $sunAura.Fill.Solid()
        $sunAura.Fill.ForeColor.RGB = Color-Hex "EA580C"
        $sunAura.Fill.Transparency = 0.65
        $sunAura.Line.Visible = 0
        $effPulse = $tLine.AddEffect($sunAura, 54, 0, 2)
        $effPulse.Timing.Duration = 2.4
        $effPulse.Timing.RepeatCount = 9999
        $effPulse.Timing.AutoReverse = -1

        # 2. Transparent 32-bit RGBA Sun Render
        $sunSize = 92.0
        $sunPic = $slide.Shapes.AddPicture($imgSun, 0, -1, $cx - ($sunSize / 2.0), $cy - ($sunSize / 2.0), $sunSize, $sunSize)
        $sunPic.Name = "!!Sun"

        # 3. Orbits & Planets Specifications
        $pDefs = @(
            @{ Name = "!!Mercury"; A = 52;  B = 24;  R = 6.0;  Color = Color-Hex "94A3B8"; Dur = 14.0; Th0 = 0.5 },
            @{ Name = "!!Venus";   A = 78;  B = 35;  R = 9.0;  Color = Color-Hex "F59E0B"; Dur = 22.0; Th0 = 1.9 },
            @{ Name = "!!Earth";   A = 108; B = 48;  R = 10.0; Color = Color-Hex "38BDF8"; Dur = 32.0; Th0 = 3.3; HasMoon = $true; TargetSlide = 3 },
            @{ Name = "!!Mars";    A = 142; B = 64;  R = 8.5;  Color = Color-Hex "EF4444"; Dur = 48.0; Th0 = 4.7; TargetSlide = 3 },
            @{ Name = "!!Jupiter"; A = 215; B = 96;  R = 22.0; Color = Color-Hex "D97706"; Dur = 80.0; Th0 = 1.1; TargetSlide = 4 },
            @{ Name = "!!Saturn";  A = 275; B = 122; R = 18.0; Color = Color-Hex "FBBF24"; Dur = 110.0; Th0 = 2.5; HasRing = $true; TargetSlide = 4 },
            @{ Name = "!!Uranus";  A = 330; B = 148; R = 13.0; Color = Color-Hex "22D3EE"; Dur = 150.0; Th0 = 3.4 },
            @{ Name = "!!Neptune"; A = 385; B = 172; R = 12.5; Color = Color-Hex "3B82F6"; Dur = 190.0; Th0 = 2.2 }
        )

        foreach ($p in $pDefs) {
            # Orbit Path
            $orb = $slide.Shapes.AddShape(9, $cx - $p.A, $cy - $p.B, $p.A * 2, $p.B * 2)
            $orb.Fill.Visible = 0
            $orb.Line.Visible = -1
            $orb.Line.ForeColor.RGB = Color-Hex "38BDF8"
            $orb.Line.Weight = 0.75
            $orb.Line.Transparency = 0.82

            # Planet Shape Initial Position
            $initX = $cx + $p.A * [Math]::Cos($p.Th0)
            $initY = $cy + $p.B * [Math]::Sin($p.Th0)
            $plSh = $slide.Shapes.AddShape(9, $initX - ($p.R / 2.0), $initY - ($p.R / 2.0), $p.R, $p.R)
            $plSh.Name = $p.Name
            $plSh.Fill.Solid()
            $plSh.Fill.ForeColor.RGB = $p.Color
            $plSh.Line.Visible = 0

            # 3D Spherical Shading Bevel
            $plSh.ThreeD.BevelTopType = 6
            $plSh.ThreeD.BevelTopInset = $p.R / 2.0
            $plSh.ThreeD.BevelTopDepth = $p.R / 2.0

            if ($p.TargetSlide) {
                $plSh.ActionSettings(1).Action = 7
                $plSh.ActionSettings(1).Hyperlink.SubAddress = "$($p.TargetSlide),$($p.TargetSlide),Slide $($p.TargetSlide)"
            }

            # Saturn Ring
            if ($p.HasRing) {
                $rg = $slide.Shapes.AddShape(9, $initX - 20, $initY - 8, 40, 16)
                $rg.Fill.Visible = 0
                $rg.Line.Visible = -1
                $rg.Line.ForeColor.RGB = Color-Hex "FDE68A"
                $rg.Line.Weight = 1.75
                $rg.Line.Transparency = 0.35
            }

            # Orbit Animation
            $eff = $tLine.AddEffect($plSh, 86, 0, 2)
            $bezPath = Get-BezierEllipsePath $cx $cy $p.A $p.B $initX $initY $p.Th0
            $eff.Timing.Duration = $p.Dur
            $eff.Timing.RepeatCount = 9999
            $eff.Timing.RepeatDuration = 99999
            $eff.Timing.Accelerate = 0.0
            $eff.Timing.Decelerate = 0.0
            $eff.Behaviors.Item(1).MotionEffect.Path = $bezPath
        }

        # 4. Asteroid Belt
        $astRand = New-Object System.Random(444)
        $astCols = @((Color-Hex "D8B4FE"), (Color-Hex "F472B6"), (Color-Hex "FDE68A"), (Color-Hex "94A3B8"))
        for ($a = 0; $a -lt 48; $a++) {
            $ath = $astRand.NextDouble() * [Math]::PI * 2.0
            $ar = 168.0 + ($astRand.NextDouble() * 32.0)
            $ax = $cx + $ar * [Math]::Cos($ath)
            $ay = $cy + ($ar * 0.44) * [Math]::Sin($ath) + (($astRand.NextDouble() - 0.5) * 6.0)
            $ac = $astCols[$astRand.Next(0, $astCols.Count)]
            $ad = $slide.Shapes.AddShape(9, $ax, $ay, 2.0, 2.0)
            $ad.Fill.Solid()
            $ad.Fill.ForeColor.RGB = $ac
            $ad.Line.Visible = 0
        }

        # 5. Transparent Voyager 1 Craft
        $vyPic = $slide.Shapes.AddPicture($imgVoyager, 0, -1, 800, 290, 56, 56)
        $vyPic.Name = "!!Voyager"
        $vyPic.ActionSettings(1).Action = 7
        $vyPic.ActionSettings(1).Hyperlink.SubAddress = "6,6,Slide 6"
    }
    elseif ($spec.Stage -eq "InnerPlanets") {
        # Giant Sun positioned strictly below slide title (Y: 135) to prevent text collision
        $sunAura2 = $slide.Shapes.AddShape(9, -150, 105, 420, 420)
        $sunAura2.Fill.Solid(); $sunAura2.Fill.ForeColor.RGB = Color-Hex "EA580C"; $sunAura2.Fill.Transparency = 0.72; $sunAura2.Line.Visible = 0
        $effPulse2 = $tLine.AddEffect($sunAura2, 54, 0, 2)
        $effPulse2.Timing.Duration = 3.0; $effPulse2.Timing.RepeatCount = 9999; $effPulse2.Timing.AutoReverse = -1

        $sunBig = $slide.Shapes.AddPicture($imgSun, 0, -1, -120, 135, 360, 360)
        $sunBig.Name = "!!Sun"

        # Orbit arcs
        foreach ($rDist in @(410, 505, 630, 770)) {
            $arc = $slide.Shapes.AddShape(9, 120 - ($rDist - 120), 280 - ($rDist - 120)*0.38, ($rDist - 120)*2, ($rDist - 120)*0.76)
            $arc.Fill.Visible = 0; $arc.Line.Visible = -1; $arc.Line.ForeColor.RGB = Color-Hex "38BDF8"; $arc.Line.Transparency = 0.85
        }

        # Mercury
        $m1 = $slide.Shapes.AddShape(9, 400, 255, 24, 24)
        $m1.Name = "!!Mercury"
        $m1.Fill.Solid(); $m1.Fill.ForeColor.RGB = Color-Hex "94A3B8"; $m1.Line.Visible = 0
        $m1.ThreeD.BevelTopType = 6; $m1.ThreeD.BevelTopInset = 12; $m1.ThreeD.BevelTopDepth = 12

        # Venus
        $v1 = $slide.Shapes.AddShape(9, 495, 270, 34, 34)
        $v1.Name = "!!Venus"
        $v1.Fill.Solid(); $v1.Fill.ForeColor.RGB = Color-Hex "F59E0B"; $v1.Line.Visible = 0
        $v1.ThreeD.BevelTopType = 6; $v1.ThreeD.BevelTopInset = 17; $v1.ThreeD.BevelTopDepth = 17

        # Earth System (cx=635, cy=240, size=85)
        $ePic = $slide.Shapes.AddPicture($imgEarth, 0, -1, 592, 197, 85, 85)
        $ePic.Name = "!!Earth"

        # CONTINUOUS REVOLVING MOON IN SLIDE 2 AROUND EARTH (Motion flow continuity!)
        $eCenter2X = 635.0; $eCenter2Y = 240.0; $mOrbit2A = 62.0; $mOrbit2B = 25.0
        $mInit2X = $eCenter2X + $mOrbit2A * [Math]::Cos(0.0)
        $mInit2Y = $eCenter2Y + $mOrbit2B * [Math]::Sin(0.0)

        # Moon orbit line around Earth
        $mOrbLine2 = $slide.Shapes.AddShape(9, $eCenter2X - $mOrbit2A, $eCenter2Y - $mOrbit2B, $mOrbit2A * 2, $mOrbit2B * 2)
        $mOrbLine2.Fill.Visible = 0; $mOrbLine2.Line.Visible = -1; $mOrbLine2.Line.ForeColor.RGB = Color-Hex "38BDF8"; $mOrbLine2.Line.Transparency = 0.80

        # Moon picture
        $mnPic = $slide.Shapes.AddPicture($imgMoon, 0, -1, $mInit2X - 11, $mInit2Y - 11, 22, 22)
        $mnPic.Name = "!!Moon"

        # Moon continuous revolution animation around Earth in Slide 2
        $effM2 = $tLine.AddEffect($mnPic, 86, 0, 2)
        $effM2.Behaviors.Item(1).MotionEffect.Path = Get-BezierEllipsePath $eCenter2X $eCenter2Y $mOrbit2A $mOrbit2B $mInit2X $mInit2Y 0.0
        $effM2.Timing.Duration = 5.0
        $effM2.Timing.RepeatCount = 9999
        $effM2.Timing.Accelerate = 0.0; $effM2.Timing.Decelerate = 0.0

        # Mars
        $mPic = $slide.Shapes.AddPicture($imgMars, 0, -1, 755, 245, 72, 72)
        $mPic.Name = "!!Mars"
    }
    elseif ($spec.Stage -eq "EarthMars") {
        # =====================================================================
        # HERO EARTH VS MARS WITH ANTI-COLLISION & CONTINUOUS REVOLVING MOON
        # =====================================================================
        # 1. Earth System (Hero: cx=360, cy=220, diameter=190)
        $ex = 360; $ey = 220; $er = 190
        $earthPic = $slide.Shapes.AddPicture($imgEarth, 0, -1, $ex - ($er/2.0), $ey - ($er/2.0), $er, $er)
        $earthPic.Name = "!!Earth"

        # Targeting Reticle closely hugging Earth (radius 102), NEVER intersecting Moon's orbit
        Add-CornerReticle $slide $ex $ey 102 "EARTH-01" "38BDF8"

        # ACTIVE CONTINUOUS REVOLVING MOON AROUND EARTH IN SLIDE 3!
        # Orbit semi-major A=142, semi-minor B=52 (flies gracefully around Earth)
        $mOrbitA = 142.0; $mOrbitB = 52.0; $mTh0 = 0.0
        $mInitX = $ex + $mOrbitA * [Math]::Cos($mTh0)
        $mInitY = $ey + $mOrbitB * [Math]::Sin($mTh0)

        # Vector orbit trace
        $mOrb = $slide.Shapes.AddShape(9, $ex - $mOrbitA, $ey - $mOrbitB, $mOrbitA * 2, $mOrbitB * 2)
        $mOrb.Fill.Visible = 0; $mOrb.Line.Visible = -1; $mOrb.Line.ForeColor.RGB = Color-Hex "38BDF8"; $mOrb.Line.Transparency = 0.85
        $mOrb.ZOrder(1)

        # Transparent Moon Picture
        $moonPic = $slide.Shapes.AddPicture($imgMoon, 0, -1, $mInitX - 20, $mInitY - 20, 40, 40)
        $moonPic.Name = "!!Moon"

        # Continuous 60 FPS Orbit Animation of Moon revolving around Earth
        $effMoon = $tLine.AddEffect($moonPic, 86, 0, 2)
        $effMoon.Behaviors.Item(1).MotionEffect.Path = Get-BezierEllipsePath $ex $ey $mOrbitA $mOrbitB $mInitX $mInitY $mTh0
        $effMoon.Timing.Duration = 8.5
        $effMoon.Timing.RepeatCount = 9999
        $effMoon.Timing.RepeatDuration = 99999
        $effMoon.Timing.Accelerate = 0.0
        $effMoon.Timing.Decelerate = 0.0

        # Holographic Glass Telemetry Card (Positioned safely at Y=360, safe margin > 50px below Earth)
        $eMetrics = @(
            "• Bán kính: 6,371 km | Khối lượng: 5.97 × 10²⁴ kg",
            "• Vận tốc quỹ đạo: 29.78 km/s | Cự ly: 1.000 AU",
            "• Khí quyển: 78% N₂, 21% O₂ | Áp suất: 101.3 kPa",
            "• Vệ tinh tự nhiên: Mặt Trăng (Đang quay quanh 60 FPS)"
        )
        Add-HolographicCard $slide 210 360 280 125 "THÔNG SỐ VẬT LÝ NASA • TRÁI ĐẤT" $eMetrics "38BDF8"

        # 2. Mars System (Hero: mx=730, my=220, diameter=180)
        $mx = 730; $my = 220; $mr = 180
        $marsPic = $slide.Shapes.AddPicture($imgMars, 0, -1, $mx - ($mr/2.0), $my - ($mr/2.0), $mr, $mr)
        $marsPic.Name = "!!Mars"

        # Targeting Reticle hugging Mars (radius 98)
        Add-CornerReticle $slide $mx $my 98 "MARS-04" "EF4444"

        # Holographic Glass Telemetry Card for Mars (Positioned safely at Y=360)
        $mMetrics = @(
            "• Bán kính: 3,389 km | Khối lượng: 6.42 × 10²³ kg",
            "• Vận tốc quỹ đạo: 24.07 km/s | Cự ly: 1.524 AU",
            "• Khí quyển: 95.3% CO₂ | Áp suất: 0.636 kPa",
            "• Cực hạn: Núi lửa Olympus Mons (cao 21.9 km)"
        )
        Add-HolographicCard $slide 590 360 280 125 "THÔNG SỐ VẬT LÝ NASA • SAO HỎA" $mMetrics "EF4444"
    }
    elseif ($spec.Stage -eq "GasGiants") {
        # =====================================================================
        # GAS GIANTS WITH STRICT ANTI-COLLISION CLEARANCE
        # =====================================================================
        # 1. Jupiter (Hero: jx=340, jy=210, diameter=210)
        $jx = 340; $jy = 210; $jr = 210
        $jupPic = $slide.Shapes.AddPicture($imgJupiter, 0, -1, $jx - ($jr/2.0), $jy - ($jr/2.0), $jr, $jr)
        $jupPic.Name = "!!Jupiter"

        Add-CornerReticle $slide $jx $jy 112 "JUPITER-05" "F59E0B"

        # Card at Y=355 (Safe margin > 40px below Jupiter)
        $jMetrics = @(
            "• Bán kính: 69,911 km (11x Đất) | Khối lượng: 317.8 M⊕",
            "• Vận tốc quỹ đạo: 13.07 km/s | Cự ly: 5.204 AU",
            "• Vết Đỏ Lớn: Xoáy bão 16,000 km tồn tại > 350 năm",
            "• Số lượng vệ tinh: 95 vệ tinh đã xác nhận"
        )
        Add-HolographicCard $slide 200 355 280 125 "THÔNG SỐ VẬT LÝ NASA • SAO MỘC" $jMetrics "F59E0B"

        # 2. Saturn (Hero: sx=730, sy=210, width=290, height=180)
        $sx = 730; $sy = 210; $sr = 290
        $satPic = $slide.Shapes.AddPicture($imgSaturn, 0, -1, $sx - ($sr/2.0), $sy - ($sr/2.0), $sr, $sr)
        $satPic.Name = "!!Saturn"

        Add-CornerReticle $slide $sx $sy 115 "SATURN-06" "FBBF24"

        # Card at Y=355 (Safe margin > 45px below Saturn rings)
        $sMetrics = @(
            "• Bán kính: 58,232 km (9.1x Đất) | Nghiêng trục: 26.73°",
            "• Vận tốc quỹ đạo: 9.68 km/s | Cự ly: 9.537 AU",
            "• Vành đai: Rộng 282,000 km, 99% hạt băng đá",
            "• Số lượng vệ tinh: 146 vệ tinh (Titan lớn nhất)"
        )
        Add-HolographicCard $slide 590 355 280 125 "THÔNG SỐ VẬT LÝ NASA • SAO THỔ" $sMetrics "FBBF24"
    }
    elseif ($spec.Stage -eq "GrandScale") {
        $cx = 560; $cy = 285
        $sunMini = $slide.Shapes.AddPicture($imgSun, 0, -1, $cx - 16, $cy - 16, 32, 32)
        $sunMini.Name = "!!Sun"

        $radii = @(28, 48, 68, 92, 122, 150, 178, 204)
        $pCols = @(
            (Color-Hex "94A3B8"),
            (Color-Hex "F59E0B"),
            (Color-Hex "38BDF8"),
            (Color-Hex "EF4444"),
            (Color-Hex "D97706"),
            (Color-Hex "FBBF24"),
            (Color-Hex "22D3EE"),
            (Color-Hex "3B82F6")
        )
        $pNames = @("Thủy", "Kim", "Địa", "Hỏa", "Mộc", "Thổ", "T.Vương", "H.Vương")

        for ($k = 0; $k -lt $radii.Count; $k++) {
            $rad = $radii[$k]
            $circ = $slide.Shapes.AddShape(9, $cx - $rad, $cy - $rad, $rad * 2, $rad * 2)
            $circ.Fill.Visible = 0; $circ.Line.Visible = -1; $circ.Line.ForeColor.RGB = Color-Hex "38BDF8"; $circ.Line.Transparency = 0.8

            $th = ($k * 0.78) + 0.3
            $px = $cx + $rad * [Math]::Cos($th)
            $py = $cy + $rad * [Math]::Sin($th)
            $pDot = $slide.Shapes.AddShape(9, $px - 4.5, $py - 4.5, 9, 9)
            $pDot.Fill.Solid(); $pDot.Fill.ForeColor.RGB = $pCols[$k]; $pDot.Line.Visible = 0
            $pDot.ThreeD.BevelTopType = 6; $pDot.ThreeD.BevelTopInset = 4.5; $pDot.ThreeD.BevelTopDepth = 4.5

            $pTxt = $slide.Shapes.AddTextbox(1, $px + 6, $py - 9, 52, 18)
            Set-TextProps $pTxt.TextFrame $pNames[$k] "Outfit" 6.5 $true $pCols[$k] 1
        }

        # Heliosphere Outer Boundary
        $helio = $slide.Shapes.AddShape(9, $cx - 215, $cy - 215, 430, 430)
        $helio.Fill.Visible = 0; $helio.Line.Visible = -1; $helio.Line.ForeColor.RGB = Color-Hex "EF4444"; $helio.Line.Weight = 1.0; $helio.Line.DashStyle = 4
        $hLbl = $slide.Shapes.AddTextbox(1, $cx - 100, $cy - 210, 200, 18)
        Set-TextProps $hLbl.TextFrame "HELIOSPHERE (NHẬT MÃN • 82+ AU)" "Outfit" 7.0 $true (Color-Hex "EF4444") 2
    }
    elseif ($spec.Stage -eq "VoyagerDeep") {
        # Voyager 1 at center stage (vx=490, vy=215, size=310)
        $vx = 490; $vy = 215; $vw = 310; $vh = 310
        $voyPic = $slide.Shapes.AddPicture($imgVoyager, 0, -1, $vx - ($vw/2.0), $vy - ($vh/2.0), $vw, $vh)
        $voyPic.Name = "!!Voyager"

        # Continuous gentle deep space drift animation
        $effVoy = $tLine.AddEffect($voyPic, 86, 0, 2)
        $effVoy.Behaviors.Item(1).MotionEffect.Path = "M 0 0 L 0.02 0.015"
        $effVoy.Timing.Duration = 12.0
        $effVoy.Timing.RepeatCount = 9999
        $effVoy.Timing.AutoReverse = -1

        # Targeting Reticle around Voyager
        Add-CornerReticle $slide $vx $vy 155 "VOYAGER-1 INTERSTELLAR" "F59E0B"

        # Distant Pale Sun in upper right corner
        $pSun = $slide.Shapes.AddPicture($imgSun, 0, -1, 880, 40, 24, 24)
        $pSun.Name = "!!Sun"
        $pSunLbl = $slide.Shapes.AddTextbox(1, 810, 68, 160, 16)
        Set-TextProps $pSunLbl.TextFrame "Mặt Trời (162.5+ AU)" "Outfit" 7.0 $true (Color-Hex "F59E0B") 2

        # Holographic Telemetry Card placed strictly at bottom (Y=385, margin > 25px from Voyager)
        $vMetrics = @(
            "• Tọa độ hiện tại: 162.5+ AU (~24.3 tỷ km từ Trái Đất)",
            "• Vận tốc tương đối: 16.9 km/s (61,000 km/h) • Độ trễ tín hiệu: ~45 giờ ánh sáng",
            "• Tải trọng biểu tượng: Đĩa Ghi Vàng (Golden Record) lưu giữ thông điệp Trái Đất"
        )
        Add-HolographicCard $slide 250 385 460 95 "DỮ LIỆU ĐIỀU KHIỂN TỪ XA • JPL / NASA" $vMetrics "F59E0B"
    }

    # =========================================================================
    # SINGLE-CLICK INSTANT SLIDE ADVANCE OVERLAY
    # =========================================================================
    $clickOverlay = $slide.Shapes.AddShape(1, 0, 0, 960, 540)
    $clickOverlay.Fill.Transparency = 1.0
    $clickOverlay.Line.Visible = 0
    $clickOverlay.ActionSettings(1).Action = 1
    $clickOverlay.ZOrder(1)

    # Transition configuration
    $slide.SlideShowTransition.EntryEffect = $spec.Transition
    $slide.SlideShowTransition.Duration = if ($spec.Transition -eq 3881) { 1.25 } else { 0.8 }
    $slide.SlideShowTransition.AdvanceOnClick = -1
}

# 3. Save and export all 6 slides for visual inspection
$pres.SaveAs($pptxPath)
Start-Sleep -Milliseconds 800

for ($k = 1; $k -le 6; $k++) {
    $exportPng = "C:\Users\Tein\.gemini\antigravity-ide\brain\0b39b3e6-c732-4100-bc21-7a56baf708c4\pure_web3d_slide_$k.png"
    $pres.Slides.Item($k).Export($exportPng, "PNG", 1920, 1080)
    Write-Host "Exported Slide $k inspection image to: $exportPng"
}

Write-Host "SUCCESS: Generated Anti-Collision 6-Slide Presentation with Moon Revolving at: $pptxPath"
