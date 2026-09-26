param(
    [string]$pptxPath = (Join-Path (Get-Location) "solar_system_3d_vector.pptx")
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

# Asset Paths for High-Res Photorealistic 3D Renders (Portable resolution)
$assetDir = Join-Path $PSScriptRoot "..\assets\planets"
if (-not (Test-Path $assetDir)) {
    $assetDir = "C:\Users\Tein\Downloads\KT ChiPhi\assets\planets"
}
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

# The 6 exact slides matching solar_3d/index.html & js/main.js
$deckSpecs = @(
    @{
        Index      = 0
        StepTag    = "TRANG 01 / 06 • TOÀN CẢNH VŨ TRỤ"
        Title      = "Toàn Cảnh Thái Dương Hệ"
        Desc       = "8 hành tinh đang liên tục quay quanh Mặt Trời theo thời gian thực. Mỗi hành tinh sở hữu vận tốc và quỹ đạo Keplerian riêng biệt."
        Stage      = "FullSystem"
        Transition = 3849 # FadeSmoothly
    },
    @{
        Index      = 1
        StepTag    = "TRANG 02 / 06 • VÙNG ĐẤT ĐÁ"
        Title      = "Vùng Hành Tinh Đất Đá"
        Desc       = "Vùng nội hệ gần Mặt Trời - nơi các hành tinh sở hữu bề mặt đá rắn chắc, mật độ vật chất cao và lõi kim loại nặng."
        Stage      = "InnerPlanets"
        Transition = 3853 # PushLeft
    },
    @{
        Index      = 2
        StepTag    = "TRANG 03 / 06 • ĐỐI CHIẾU THIÊN THỂ"
        Title      = "Trái Đất so với Sao Hỏa"
        Desc       = "Quan sát cận cảnh Trái Đất (với Mặt Trăng quay quanh) và Sao Hỏa đỏ rực - đích đến tương lai của loài người."
        Stage      = "EarthMars"
        Transition = 3585 # SplitHorizontalOut
    },
    @{
        Index      = 3
        StepTag    = "TRANG 04 / 06 • KHỔNG LỒ NGOẠI HỆ"
        Title      = "Các Gã Khổng Lồ Khí & Băng"
        Desc       = "Vượt qua Vành đai Tiểu hành tinh để chiêm ngưỡng Sao Mộc và Sao Thổ với kiệt tác vành đai băng đá 3D lộng lẫy."
        Stage      = "GasGiants"
        Transition = 1537 # Dissolve
    },
    @{
        Index      = 4
        StepTag    = "TRANG 05 / 06 • DỮ LIỆU THIÊN VĂN"
        Title      = "Thước Đo Quy Mô Thái Dương Hệ"
        Desc       = "Góc nhìn bao quát từ trên cao thể hiện rõ khoảng cách và tỷ lệ kích thước giữa các hành tinh trong vũ trụ."
        Stage      = "GrandScale"
        Transition = 3849 # FadeSmoothly
    },
    @{
        Index      = 5
        StepTag    = "TRANG 06 / 06 • KHÁM PHÁ VÔ TẬN"
        Title      = "Khát Vọng Vươn Ra Vũ Trụ"
        Desc       = "Hành trình thám hiểm vũ trụ không bao giờ dừng lại. Loài người đang từng bước trở thành nền văn minh đa hành tinh."
        Stage      = "VoyagerDeep"
        Transition = 3853 # PushLeft
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

# 2. Iterate through each slide definition
for ($sIdx = 0; $sIdx -lt $deckSpecs.Count; $sIdx++) {
    $spec = $deckSpecs[$sIdx]
    Write-Host "Creating Pure Photorealistic 3D Slide $($sIdx + 1): $($spec.Title)..."

    $slide = $pres.Slides.Add($sIdx + 1, 12) # ppLayoutBlank
    $slide.FollowMasterBackground = 0
    $slide.Background.Fill.Solid()
    $slide.Background.Fill.ForeColor.RGB = Color-Hex "000000" # Pure cosmic black for seamless asset blend

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
    # SLIDE HEADER (Clean, minimalist typography without any HUD clutter)
    # =========================================================================
    $tagTx = $slide.Shapes.AddTextbox(1, 36, 32, 400, 20)
    Set-TextProps $tagTx.TextFrame $spec.StepTag "Outfit" 8.0 $true (Color-Hex "F59E0B") 1

    $titleTx = $slide.Shapes.AddTextbox(1, 34, 52, 440, 38)
    Set-TextProps $titleTx.TextFrame $spec.Title "Outfit" 20.0 $true (Color-Hex "FFFFFF") 1

    $descTx = $slide.Shapes.AddTextbox(1, 36, 94, 420, 52)
    Set-TextProps $descTx.TextFrame $spec.Desc "Outfit" 9.5 $false (Color-Hex "94A3B8") 1

    # =========================================================================
    # 3D CELESTIAL STAGE (PHOTOREALISTIC 3D ASSETS + KEPLERIAN MOTION)
    # =========================================================================
    $tLine = $slide.TimeLine.MainSequence

    if ($spec.Stage -eq "FullSystem") {
        # Full 8 Planets View (Clean, pure cosmic view)
        $cx = 580.0; $cy = 280.0

        # 1. The Photorealistic 3D Sun
        $sunSize = 92.0
        $sunPic = $slide.Shapes.AddPicture($imgSun, 0, -1, $cx - ($sunSize / 2.0), $cy - ($sunSize / 2.0), $sunSize, $sunSize)

        # 2. Orbits & Planets Specifications (Non-overlapping positions)
        $pDefs = @(
            @{ Name = "Sao Thủy"; A = 52;  B = 24;  R = 6.0;  Color = Color-Hex "94A3B8"; Dur = 14.0; Th0 = 0.5 },
            @{ Name = "Sao Kim";   A = 78;  B = 35;  R = 9.0;  Color = Color-Hex "F59E0B"; Dur = 22.0; Th0 = 1.9 },
            @{ Name = "Trái Đất"; A = 108; B = 48;  R = 10.0; Color = Color-Hex "38BDF8"; Dur = 32.0; Th0 = 3.3; HasMoon = $true },
            @{ Name = "Sao Hỏa";  A = 142; B = 64;  R = 8.5;  Color = Color-Hex "EF4444"; Dur = 48.0; Th0 = 4.7 },
            @{ Name = "Sao Mộc";  A = 215; B = 96;  R = 22.0; Color = Color-Hex "D97706"; Dur = 80.0; Th0 = 1.1 },
            @{ Name = "Sao Thổ";  A = 275; B = 122; R = 18.0; Color = Color-Hex "FBBF24"; Dur = 110.0; Th0 = 2.5; HasRing = $true },
            @{ Name = "Sao Thiên Vương"; A = 330; B = 148; R = 13.0; Color = Color-Hex "22D3EE"; Dur = 150.0; Th0 = 3.4 },
            @{ Name = "Sao Hải Vương";  A = 385; B = 172; R = 12.5; Color = Color-Hex "3B82F6"; Dur = 190.0; Th0 = 2.2 }
        )

        foreach ($p in $pDefs) {
            # Orbit Path
            $orb = $slide.Shapes.AddShape(9, $cx - $p.A, $cy - $p.B, $p.A * 2, $p.B * 2)
            $orb.Fill.Visible = 0
            $orb.Line.Visible = -1
            $orb.Line.ForeColor.RGB = Color-Hex "38BDF8"
            $orb.Line.Weight = 0.75
            $orb.Line.Transparency = 0.82

            # Planet Shape Initial Position with 3D Spherical Bevel
            $initX = $cx + $p.A * [Math]::Cos($p.Th0)
            $initY = $cy + $p.B * [Math]::Sin($p.Th0)
            $plSh = $slide.Shapes.AddShape(9, $initX - ($p.R / 2.0), $initY - ($p.R / 2.0), $p.R, $p.R)
            $plSh.Fill.Solid()
            $plSh.Fill.ForeColor.RGB = $p.Color
            $plSh.Line.Visible = 0

            # 3D Spherical Shading Bevel
            $plSh.ThreeD.BevelTopType = 6 # msoBevelCircle
            $plSh.ThreeD.BevelTopInset = $p.R / 2.0
            $plSh.ThreeD.BevelTopDepth = $p.R / 2.0

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
            $eff = $tLine.AddEffect($plSh, 86, 0, 2) # 86 = msoAnimEffectPathRight
            $bezPath = Get-BezierEllipsePath $cx $cy $p.A $p.B $initX $initY $p.Th0
            $eff.Timing.Duration = $p.Dur
            $eff.Timing.RepeatCount = 9999
            $eff.Timing.RepeatDuration = 99999
            $eff.Timing.Accelerate = 0.0
            $eff.Timing.Decelerate = 0.0
            $eff.Behaviors.Item(1).MotionEffect.Path = $bezPath
        }

        # 3. Asteroid Belt (48 particles between Mars and Jupiter)
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

        # 4. Photorealistic 3D Voyager 1 Craft cruising in outer space
        $vyPic = $slide.Shapes.AddPicture($imgVoyager, 0, -1, 800, 290, 56, 56)
    }
    elseif ($spec.Stage -eq "InnerPlanets") {
        # Zoomed Inner Rocky System with Photorealistic 3D Sun, Earth, Moon & Mars
        # 1. Giant Photorealistic 3D Sun (positioned safely below slide header)
        $sunBig = $slide.Shapes.AddPicture($imgSun, 0, -1, -120, 110, 390, 390)

        # 2. Orbits curve previews
        foreach ($rDist in @(430, 525, 645, 775)) {
            $arc = $slide.Shapes.AddShape(9, 120 - ($rDist - 120), 270 - ($rDist - 120)*0.38, ($rDist - 120)*2, ($rDist - 120)*0.76)
            $arc.Fill.Visible = 0; $arc.Line.Visible = -1; $arc.Line.ForeColor.RGB = Color-Hex "38BDF8"; $arc.Line.Transparency = 0.85
        }

        # Mercury (Smooth sphere)
        $m1 = $slide.Shapes.AddShape(9, 418, 244, 24, 24)
        $m1.Fill.Solid(); $m1.Fill.ForeColor.RGB = Color-Hex "94A3B8"; $m1.Line.Visible = 0

        # Venus (Smooth sphere)
        $v1 = $slide.Shapes.AddShape(9, 508, 264, 34, 34)
        $v1.Fill.Solid(); $v1.Fill.ForeColor.RGB = Color-Hex "F59E0B"; $v1.Line.Visible = 0

        # Earth (Photorealistic 3D Render)
        $ePic = $slide.Shapes.AddPicture($imgEarth, 0, -1, 605, 195, 90, 90)

        # Moon (Photorealistic 3D Render)
        $mnPic = $slide.Shapes.AddPicture($imgMoon, 0, -1, 688, 180, 26, 26)

        # Mars (Photorealistic 3D Render)
        $mPic = $slide.Shapes.AddPicture($imgMars, 0, -1, 745, 245, 74, 74)
    }
    elseif ($spec.Stage -eq "EarthMars") {
        # Hero Earth vs Mars Close-Up Comparison (Photorealistic 3D CGI)
        # 1. Earth System (Left Hero)
        $ex = 425; $ey = 270; $er = 230
        $earthPic = $slide.Shapes.AddPicture($imgEarth, 0, -1, $ex - ($er/2.0), $ey - ($er/2.0), $er, $er)

        # Moon in orbit around Earth
        $moonPic = $slide.Shapes.AddPicture($imgMoon, 0, -1, $ex + 115, $ey - 100, 52, 52)
        $mOrb = $slide.Shapes.AddShape(9, $ex - 145, $ey - 70, 290, 140)
        $mOrb.Fill.Visible = 0; $mOrb.Line.Visible = -1; $mOrb.Line.ForeColor.RGB = Color-Hex "38BDF8"; $mOrb.Line.Transparency = 0.8
        $mOrb.ZOrder(1)

        $eLbl = $slide.Shapes.AddTextbox(1, $ex - 120, $ey + 130, 240, 26)
        Set-TextProps $eLbl.TextFrame "TRÁI ĐẤT • 1.00 AU • CÁI NÔI SỰ SỐNG" "Outfit" 10.0 $true (Color-Hex "38BDF8") 2

        # 2. Mars System (Right Hero)
        $mx = 745; $my = 270; $mr = 205
        $marsPic = $slide.Shapes.AddPicture($imgMars, 0, -1, $mx - ($mr/2.0), $my - ($mr/2.0), $mr, $mr)

        $mLbl = $slide.Shapes.AddTextbox(1, $mx - 120, $my + 130, 240, 26)
        Set-TextProps $mLbl.TextFrame "SAO HỎA • 1.52 AU • HÀNH TINH ĐỎ" "Outfit" 10.0 $true (Color-Hex "EF4444") 2
    }
    elseif ($spec.Stage -eq "GasGiants") {
        # Gas Giants: Jupiter & Saturn with Photorealistic 3D Rings
        # 1. Jupiter (Photorealistic 3D Render)
        $jx = 420; $jy = 270; $jr = 255
        $jupPic = $slide.Shapes.AddPicture($imgJupiter, 0, -1, $jx - ($jr/2.0), $jy - ($jr/2.0), $jr, $jr)

        $jLbl = $slide.Shapes.AddTextbox(1, $jx - 130, $jy + 140, 260, 26)
        Set-TextProps $jLbl.TextFrame "SAO MỘC • ĐỆ NHẤT THÁI DƯƠNG HỆ" "Outfit" 10.0 $true (Color-Hex "F59E0B") 2

        # 2. Saturn (Photorealistic 3D Render with 3D Rings)
        $sx = 745; $sy = 270; $sr = 325
        $satPic = $slide.Shapes.AddPicture($imgSaturn, 0, -1, $sx - ($sr/2.0), $sy - ($sr/2.0), $sr, $sr)

        $sLbl = $slide.Shapes.AddTextbox(1, $sx - 140, $sy + 140, 280, 26)
        Set-TextProps $sLbl.TextFrame "SAO THỔ • KIỆT TÁC VÀNH ĐAI BĂNG ĐÁ" "Outfit" 10.0 $true (Color-Hex "FBBF24") 2
    }
    elseif ($spec.Stage -eq "GrandScale") {
        # Top-Down Perspective of the Entire Solar System
        $cx = 560; $cy = 290
        # Mini 3D Sun
        $sunMini = $slide.Shapes.AddPicture($imgSun, 0, -1, $cx - 16, $cy - 16, 32, 32)

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

        # Heliosphere Outer Boundary (kept nicely inside slide)
        $helio = $slide.Shapes.AddShape(9, $cx - 215, $cy - 215, 430, 430)
        $helio.Fill.Visible = 0; $helio.Line.Visible = -1; $helio.Line.ForeColor.RGB = Color-Hex "EF4444"; $helio.Line.Weight = 1.0; $helio.Line.DashStyle = 4
        $hLbl = $slide.Shapes.AddTextbox(1, $cx - 100, $cy - 210, 200, 18)
        Set-TextProps $hLbl.TextFrame "HELIOSPHERE (NHẬT MÃN • 82+ AU)" "Outfit" 7.0 $true (Color-Hex "EF4444") 2
    }
    elseif ($spec.Stage -eq "VoyagerDeep") {
        # Voyager 1 at the Outer Rim of interstellar space (Photorealistic 3D Spacecraft)
        $vx = 550; $vy = 255; $vw = 380; $vh = 380
        $voyPic = $slide.Shapes.AddPicture($imgVoyager, 0, -1, $vx - ($vw/2.0), $vy - ($vh/2.0), $vw, $vh)

        # Distant Pale Sun
        $pSun = $slide.Shapes.AddPicture($imgSun, 0, -1, 875, 45, 24, 24)
        $pSunLbl = $slide.Shapes.AddTextbox(1, 805, 74, 160, 16)
        Set-TextProps $pSunLbl.TextFrame "Mặt Trời (162.5+ AU)" "Outfit" 7.0 $true (Color-Hex "F59E0B") 2

        $vTitle = $slide.Shapes.AddTextbox(1, 460, 465, 450, 30)
        Set-TextProps $vTitle.TextFrame "VOYAGER 1 • SỨ MỆNH BỜ VỰC KHÔNG GIAN LIÊN SAO" "Outfit" 11.0 $true (Color-Hex "F59E0B") 1
    }

    # =========================================================================
    # SINGLE-CLICK INSTANT SLIDE ADVANCE OVERLAY
    # =========================================================================
    $clickOverlay = $slide.Shapes.AddShape(1, 0, 0, 960, 540)
    $clickOverlay.Fill.Transparency = 1.0 # 100% invisible
    $clickOverlay.Line.Visible = 0
    $clickOverlay.ActionSettings(1).Action = 1 # ppActionNextSlide

    # Transition
    $slide.SlideShowTransition.EntryEffect = $spec.Transition
    $slide.SlideShowTransition.Duration = 0.8
    $slide.SlideShowTransition.AdvanceOnClick = -1
}

# 3. Save and export all 6 slides for instant visual inspection
$pres.SaveAs($pptxPath)
Start-Sleep -Milliseconds 800

$exportDir = Join-Path $PSScriptRoot "..\export"
if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Force -Path $exportDir | Out-Null }

for ($k = 1; $k -le 6; $k++) {
    $exportPng = Join-Path $exportDir "slide_$k.png"
    $pres.Slides.Item($k).Export($exportPng, "PNG", 1920, 1080)
    Write-Host "Exported Slide $k inspection image to: $exportPng"
}

Write-Host "SUCCESS: Generated 100% Pure Photorealistic 3D Presentation at: $pptxPath"
