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

# Asset Paths for High-Res Photorealistic 3D Renders
$assetDir = "C:\Users\Tein\Downloads\KT ChiPhi\assets\planets"
$imgSun     = Join-Path $assetDir "sun.jpg"
$imgEarth   = Join-Path $assetDir "earth.jpg"
$imgMars    = Join-Path $assetDir "mars.jpg"
$imgJupiter = Join-Path $assetDir "jupiter.jpg"
$imgSaturn  = Join-Path $assetDir "saturn.jpg"
$imgVoyager = Join-Path $assetDir "voyager.jpg"
$imgMoon    = Join-Path $assetDir "moon.jpg"

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
        Desc       = "Vùng nội hệ gần Mặt Trời - nơi các hành tinh sở hữu bề mặt đá rắn chắc, mật độ vật chất cao và lõi kim loại nặng."
        Stage      = "InnerPlanets"
        Transition = 3881 # Morph 3D
    },
    @{
        Index      = 2
        StepTag    = "TRANG 03 / 06 • ĐỐI CHIẾU THIÊN THỂ"
        Title      = "Trái Đất so với Sao Hỏa"
        Desc       = "Quan sát cận cảnh Trái Đất (với Mặt Trăng quay quanh) và Sao Hỏa đỏ rực - đích đến tương lai của loài người."
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
        Desc       = "Hành trình thám hiểm vũ trụ không bao giờ dừng lại. Loài người đang từng bước trở thành nền văn minh đa hành tinh."
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

# Helper to draw a Holographic Glass Telemetry Card
function Add-HolographicCard($slide, $x, $y, $w, $h, $titleText, $metrics, $accentHex) {
    # 1. Semi-transparent dark navy glass backing
    $card = $slide.Shapes.AddShape(1, $x, $y, $w, $h)
    $card.Fill.Solid()
    $card.Fill.ForeColor.RGB = Color-Hex "0A1128" # Deep midnight glass
    $card.Fill.Transparency = 0.20 # 80% opacity
    $card.Line.Visible = -1
    $card.Line.ForeColor.RGB = Color-Hex $accentHex
    $card.Line.Weight = 1.0
    $card.Line.Transparency = 0.35

    # 2. Header bar with Sci-Fi title
    $headTx = $slide.Shapes.AddTextbox(1, $x + 10, $y + 6, $w - 20, 20)
    Set-TextProps $headTx.TextFrame $titleText "Outfit" 8.0 $true (Color-Hex $accentHex) 1

    # 3. Telemetry lines
    $bodyText = $metrics -join "`n"
    $bodyTx = $slide.Shapes.AddTextbox(1, $x + 10, $y + 24, $w - 20, $h - 28)
    Set-TextProps $bodyTx.TextFrame $bodyText "Outfit" 7.5 $false (Color-Hex "E2E8F0") 1
}

# Helper to draw a Sci-Fi Reticle ring
function Add-TargetReticle($slide, $cx, $cy, $radius, $label, $accentHex) {
    # Inner dashed ring
    $ring = $slide.Shapes.AddShape(9, $cx - $radius, $cy - $radius, $radius * 2, $radius * 2)
    $ring.Fill.Visible = 0
    $ring.Line.Visible = -1
    $ring.Line.ForeColor.RGB = Color-Hex $accentHex
    $ring.Line.Weight = 1.0
    $ring.Line.DashStyle = 4 # Dashed
    $ring.Line.Transparency = 0.4

    # Target label tag
    $tag = $slide.Shapes.AddTextbox(1, $cx - $radius, $cy - $radius - 14, $radius * 2, 14)
    Set-TextProps $tag.TextFrame "[ TARGET: $label ]" "Consolas" 6.5 $true (Color-Hex $accentHex) 2
}

# 2. Iterate through each slide definition
for ($sIdx = 0; $sIdx -lt $deckSpecs.Count; $sIdx++) {
    $spec = $deckSpecs[$sIdx]
    Write-Host "Creating Advanced Cinematic 3D Slide $($sIdx + 1): $($spec.Title)..."

    $slide = $pres.Slides.Add($sIdx + 1, 12) # ppLayoutBlank
    $slide.FollowMasterBackground = 0
    $slide.Background.Fill.Solid()
    $slide.Background.Fill.ForeColor.RGB = Color-Hex "000000"

    # Backing canvas shape to guarantee deep dark cosmic background
    $bg = $slide.Shapes.AddShape(1, 0, 0, 960, 540)
    $bg.Fill.Solid()
    $bg.Fill.ForeColor.RGB = Color-Hex "000000"
    $bg.Line.Visible = 0
    $bg.ZOrder(1)

    # Ambient Starfield (65 vector stars)
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
    # SLIDE HEADER
    # =========================================================================
    $tagTx = $slide.Shapes.AddTextbox(1, 36, 32, 400, 20)
    Set-TextProps $tagTx.TextFrame $spec.StepTag "Outfit" 8.0 $true (Color-Hex "F59E0B") 1

    $titleTx = $slide.Shapes.AddTextbox(1, 34, 52, 440, 38)
    Set-TextProps $titleTx.TextFrame $spec.Title "Outfit" 20.0 $true (Color-Hex "FFFFFF") 1

    $descTx = $slide.Shapes.AddTextbox(1, 36, 94, 420, 52)
    Set-TextProps $descTx.TextFrame $spec.Desc "Outfit" 9.5 $false (Color-Hex "94A3B8") 1

    $tLine = $slide.TimeLine.MainSequence

    # =========================================================================
    # 3D CELESTIAL STAGE
    # =========================================================================
    if ($spec.Stage -eq "FullSystem") {
        $cx = 580.0; $cy = 280.0

        # Dynamic Comet Halley sweeping across cosmic space
        $cmt = $slide.Shapes.AddShape(9, -20, 40, 4, 4)
        $cmt.Fill.Solid(); $cmt.Fill.ForeColor.RGB = Color-Hex "38BDF8"; $cmt.Line.Visible = 0
        $effCmt = $tLine.AddEffect($cmt, 86, 0, 2)
        $effCmt.Behaviors.Item(1).MotionEffect.Path = "M 0 0 L 1.15 0.75"
        $effCmt.Timing.Duration = 4.5
        $effCmt.Timing.RepeatCount = 9999
        $effCmt.Timing.RepeatDuration = 99999
        $effCmt.Timing.Accelerate = 0.25
        $effCmt.Timing.Decelerate = 0.25

        # 1. Pulsating Breathing Corona Aura behind the Sun
        $sunAura = $slide.Shapes.AddShape(9, $cx - 58, $cy - 58, 116, 116)
        $sunAura.Fill.Solid()
        $sunAura.Fill.ForeColor.RGB = Color-Hex "EA580C"
        $sunAura.Fill.Transparency = 0.65
        $sunAura.Line.Visible = 0
        $effPulse = $tLine.AddEffect($sunAura, 54, 0, 2) # GrowShrink
        $effPulse.Timing.Duration = 2.4
        $effPulse.Timing.RepeatCount = 9999
        $effPulse.Timing.AutoReverse = -1 # Smooth breathing expansion & contraction

        # 2. Photorealistic 3D Sun (Tagged for Morph 3D Tracking)
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
            $plSh.ThreeD.BevelTopType = 6 # msoBevelCircle
            $plSh.ThreeD.BevelTopInset = $p.R / 2.0
            $plSh.ThreeD.BevelTopDepth = $p.R / 2.0

            # Interactive Hyperlink if target slide exists
            if ($p.TargetSlide) {
                $plSh.ActionSettings(1).Action = 7 # ppActionHyperlink
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

            # Continuous Keplerian Orbit Animation
            $eff = $tLine.AddEffect($plSh, 86, 0, 2)
            $bezPath = Get-BezierEllipsePath $cx $cy $p.A $p.B $initX $initY $p.Th0
            $eff.Timing.Duration = $p.Dur
            $eff.Timing.RepeatCount = 9999
            $eff.Timing.RepeatDuration = 99999
            $eff.Timing.Accelerate = 0.0
            $eff.Timing.Decelerate = 0.0
            $eff.Behaviors.Item(1).MotionEffect.Path = $bezPath
        }

        # 4. Asteroid Belt (48 particles)
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

        # 5. Photorealistic 3D Voyager 1 Craft
        $vyPic = $slide.Shapes.AddPicture($imgVoyager, 0, -1, 800, 290, 56, 56)
        $vyPic.Name = "!!Voyager"
        $vyPic.ActionSettings(1).Action = 7
        $vyPic.ActionSettings(1).Hyperlink.SubAddress = "6,6,Slide 6"
    }
    elseif ($spec.Stage -eq "InnerPlanets") {
        # Giant Sun with Breathing Plasma Aura
        $sunAura2 = $slide.Shapes.AddShape(9, -150, 80, 450, 450)
        $sunAura2.Fill.Solid(); $sunAura2.Fill.ForeColor.RGB = Color-Hex "EA580C"; $sunAura2.Fill.Transparency = 0.72; $sunAura2.Line.Visible = 0
        $effPulse2 = $tLine.AddEffect($sunAura2, 54, 0, 2)
        $effPulse2.Timing.Duration = 3.0; $effPulse2.Timing.RepeatCount = 9999; $effPulse2.Timing.AutoReverse = -1

        $sunBig = $slide.Shapes.AddPicture($imgSun, 0, -1, -120, 110, 390, 390)
        $sunBig.Name = "!!Sun"

        # Orbit arcs
        foreach ($rDist in @(430, 525, 645, 775)) {
            $arc = $slide.Shapes.AddShape(9, 120 - ($rDist - 120), 270 - ($rDist - 120)*0.38, ($rDist - 120)*2, ($rDist - 120)*0.76)
            $arc.Fill.Visible = 0; $arc.Line.Visible = -1; $arc.Line.ForeColor.RGB = Color-Hex "38BDF8"; $arc.Line.Transparency = 0.85
        }

        # Mercury
        $m1 = $slide.Shapes.AddShape(9, 418, 244, 24, 24)
        $m1.Name = "!!Mercury"
        $m1.Fill.Solid(); $m1.Fill.ForeColor.RGB = Color-Hex "94A3B8"; $m1.Line.Visible = 0
        $m1.ThreeD.BevelTopType = 6; $m1.ThreeD.BevelTopInset = 12; $m1.ThreeD.BevelTopDepth = 12

        # Venus
        $v1 = $slide.Shapes.AddShape(9, 508, 264, 34, 34)
        $v1.Name = "!!Venus"
        $v1.Fill.Solid(); $v1.Fill.ForeColor.RGB = Color-Hex "F59E0B"; $v1.Line.Visible = 0
        $v1.ThreeD.BevelTopType = 6; $v1.ThreeD.BevelTopInset = 17; $v1.ThreeD.BevelTopDepth = 17

        # Earth & Moon
        $ePic = $slide.Shapes.AddPicture($imgEarth, 0, -1, 605, 195, 90, 90)
        $ePic.Name = "!!Earth"

        $mnPic = $slide.Shapes.AddPicture($imgMoon, 0, -1, 688, 180, 26, 26)
        $mnPic.Name = "!!Moon"

        # Mars
        $mPic = $slide.Shapes.AddPicture($imgMars, 0, -1, 745, 245, 74, 74)
        $mPic.Name = "!!Mars"
    }
    elseif ($spec.Stage -eq "EarthMars") {
        # 1. Earth System (Hero)
        $ex = 380; $ey = 255; $er = 210
        $earthPic = $slide.Shapes.AddPicture($imgEarth, 0, -1, $ex - ($er/2.0), $ey - ($er/2.0), $er, $er)
        $earthPic.Name = "!!Earth"

        # Moon in orbit
        $moonPic = $slide.Shapes.AddPicture($imgMoon, 0, -1, $ex + 105, $ey - 95, 48, 48)
        $moonPic.Name = "!!Moon"
        $mOrb = $slide.Shapes.AddShape(9, $ex - 135, $ey - 65, 270, 130)
        $mOrb.Fill.Visible = 0; $mOrb.Line.Visible = -1; $mOrb.Line.ForeColor.RGB = Color-Hex "38BDF8"; $mOrb.Line.Transparency = 0.82; $mOrb.ZOrder(1)

        # Targeting Reticle for Earth
        Add-TargetReticle $slide $ex $ey ($er/2.0 + 16) "EARTH-01" "38BDF8"

        # Holographic Glass Telemetry Card (NASA Earth Specs)
        $eMetrics = @(
            "• Bán kính: 6,371 km | Khối lượng: 5.97 × 10²⁴ kg",
            "• Vận tốc quỹ đạo: 29.78 km/s | Cự ly: 1.000 AU",
            "• Khí quyển: 78% N₂, 21% O₂ | Áp suất: 101.3 kPa",
            "• Vệ tinh: 1 (Mặt Trăng • Chu kỳ 27.3 ngày)"
        )
        Add-HolographicCard $slide ($ex - 130) ($ey + 125) 260 76 "THÔNG SỐ VẬT LÝ NASA • TRÁI ĐẤT" $eMetrics "38BDF8"

        # 2. Mars System (Hero)
        $mx = 725; $my = 255; $mr = 190
        $marsPic = $slide.Shapes.AddPicture($imgMars, 0, -1, $mx - ($mr/2.0), $my - ($mr/2.0), $mr, $mr)
        $marsPic.Name = "!!Mars"

        # Targeting Reticle for Mars
        Add-TargetReticle $slide $mx $my ($mr/2.0 + 16) "MARS-04" "EF4444"

        # Holographic Glass Telemetry Card (NASA Mars Specs)
        $mMetrics = @(
            "• Bán kính: 3,389 km | Khối lượng: 6.42 × 10²³ kg",
            "• Vận tốc quỹ đạo: 24.07 km/s | Cự ly: 1.524 AU",
            "• Khí quyển: 95.3% CO₂ | Áp suất: 0.636 kPa",
            "• Cực hạn: Núi Olympus Mons (cao 21.9 km)"
        )
        Add-HolographicCard $slide ($mx - 130) ($my + 125) 260 76 "THÔNG SỐ VẬT LÝ NASA • SAO HỎA" $mMetrics "EF4444"
    }
    elseif ($spec.Stage -eq "GasGiants") {
        # 1. Jupiter (Photorealistic 3D Hero)
        $jx = 370; $jy = 255; $jr = 240
        $jupPic = $slide.Shapes.AddPicture($imgJupiter, 0, -1, $jx - ($jr/2.0), $jy - ($jr/2.0), $jr, $jr)
        $jupPic.Name = "!!Jupiter"

        Add-TargetReticle $slide $jx $jy ($jr/2.0 + 16) "JUPITER-05" "F59E0B"

        $jMetrics = @(
            "• Bán kính: 69,911 km (11x Đất) | Khối lượng: 317.8 M⊕",
            "• Vận tốc quỹ đạo: 13.07 km/s | Cự ly: 5.204 AU",
            "• Vết Đỏ Lớn: Xoáy bão 16,000 km tồn tại > 350 năm",
            "• Số lượng vệ tinh: 95 vệ tinh đã xác nhận"
        )
        Add-HolographicCard $slide ($jx - 135) ($jy + 135) 270 76 "THÔNG SỐ VẬT LÝ NASA • SAO MỘC" $jMetrics "F59E0B"

        # 2. Saturn with 3D Rings
        $sx = 735; $sy = 255; $sr = 310
        $satPic = $slide.Shapes.AddPicture($imgSaturn, 0, -1, $sx - ($sr/2.0), $sy - ($sr/2.0), $sr, $sr)
        $satPic.Name = "!!Saturn"

        Add-TargetReticle $slide $sx $sy ($sr/2.0 + 14) "SATURN-06" "FBBF24"

        $sMetrics = @(
            "• Bán kính: 58,232 km (9.1x Đất) | Nghiêng trục: 26.73°",
            "• Vận tốc quỹ đạo: 9.68 km/s | Cự ly: 9.537 AU",
            "• Vành đai: Rộng 282,000 km, 99% hạt băng đá",
            "• Số lượng vệ tinh: 146 vệ tinh (Titan lớn nhất)"
        )
        Add-HolographicCard $slide ($sx - 135) ($jy + 135) 270 76 "THÔNG SỐ VẬT LÝ NASA • SAO THỔ" $sMetrics "FBBF24"
    }
    elseif ($spec.Stage -eq "GrandScale") {
        $cx = 560; $cy = 290
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
        $vx = 520; $vy = 245; $vw = 360; $vh = 360
        $voyPic = $slide.Shapes.AddPicture($imgVoyager, 0, -1, $vx - ($vw/2.0), $vy - ($vh/2.0), $vw, $vh)
        $voyPic.Name = "!!Voyager"

        # Targeting Reticle around Voyager
        Add-TargetReticle $slide $vx $vy 175 "VOYAGER-1 INTERSTELLAR" "F59E0B"

        # Distant Pale Sun
        $pSun = $slide.Shapes.AddPicture($imgSun, 0, -1, 875, 45, 24, 24)
        $pSun.Name = "!!Sun"
        $pSunLbl = $slide.Shapes.AddTextbox(1, 805, 74, 160, 16)
        Set-TextProps $pSunLbl.TextFrame "Mặt Trời (162.5+ AU)" "Outfit" 7.0 $true (Color-Hex "F59E0B") 2

        # Holographic Telemetry Card for Voyager 1
        $vMetrics = @(
            "• Tọa độ hiện tại: 162.5+ AU (~24.3 tỷ km từ Trái Đất)",
            "• Vận tốc tương đối: 16.9 km/s (61,000 km/h)",
            "• Độ trễ tín hiệu 2 chiều: ~45 giờ ánh sáng",
            "• Tải trọng biểu tượng: Đĩa Ghi Vàng (Golden Record) lưu giữ thông điệp Trái Đất"
        )
        Add-HolographicCard $slide 330 440 460 76 "DỮ LIỆU ĐIỀU KHIỂN TỪ XA • JPL / NASA" $vMetrics "F59E0B"
    }

    # =========================================================================
    # SINGLE-CLICK INSTANT SLIDE ADVANCE OVERLAY
    # =========================================================================
    $clickOverlay = $slide.Shapes.AddShape(1, 0, 0, 960, 540)
    $clickOverlay.Fill.Transparency = 1.0 # 100% invisible
    $clickOverlay.Line.Visible = 0
    $clickOverlay.ActionSettings(1).Action = 1 # ppActionNextSlide
    $clickOverlay.ZOrder(1) # Send behind interactive shapes

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

Write-Host "SUCCESS: Generated Advanced Cinematic Presentation with Morph 3D & Hologram HUD at: $pptxPath"
