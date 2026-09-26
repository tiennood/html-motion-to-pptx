param(
    [string]$pptxPath = (Join-Path (Get-Location) "solar_system_max_single_slide.pptx")
)

# 1. Close any running PowerPoint instances safely
Stop-Process -Name POWERPNT -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 600

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = 1

$pres = $ppt.Presentations.Add()
$W = 960.0; $H = 540.0
$pres.PageSetup.SlideWidth = $W
$pres.PageSetup.SlideHeight = $H

$ci = [System.Globalization.CultureInfo]::InvariantCulture
$kappa = (4.0 / 3.0) * ([Math]::Sqrt(2.0) - 1.0) # 0.5522847498

# Asset Paths for High-Res Photorealistic 3D Renders
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

function Color-Hex([string]$hex) {
    $hex = $hex.TrimStart('#')
    $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return [int]($r + ($g * 256) + ($b * 65536))
}

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

Write-Host "Creating Single-Slide Maximum Interactive Solar 3D Application..."

$slide = $pres.Slides.Add(1, 12) # Blank
$slide.FollowMasterBackground = 0
$slide.Background.Fill.Solid()
$slide.Background.Fill.ForeColor.RGB = Color-Hex "000000"

# Pure Cosmic Black Base Canvas
$bg = $slide.Shapes.AddShape(1, 0, 0, $W, $H)
$bg.Fill.Solid(); $bg.Fill.ForeColor.RGB = Color-Hex "000000"; $bg.Line.Visible = 0
$bg.ZOrder(1)

# Ambient Starfield (80 vector stars)
$rand = New-Object System.Random(2026)
$starColors = @((Color-Hex "F8FAFC"), (Color-Hex "38BDF8"), (Color-Hex "F59E0B"), (Color-Hex "C084FC"))
for ($i = 0; $i -lt 80; $i++) {
    $sx = $rand.Next(10, 950); $sy = $rand.Next(10, 530)
    $sr = 0.8 + ($rand.NextDouble() * 1.8)
    $sc = $starColors[$rand.Next(0, $starColors.Count)]
    $st = $slide.Shapes.AddShape(9, $sx, $sy, $sr, $sr)
    $st.Fill.Solid(); $st.Fill.ForeColor.RGB = $sc; $st.Line.Visible = 0
}

# =========================================================================
# HEADER & COCKPIT CONTROL BUTTONS (INTERACTIVE TRIGGERS)
# =========================================================================
$headTag = $slide.Shapes.AddTextbox(1, 32, 22, 280, 18)
Set-TextProps $headTag.TextFrame "HỆ THỐNG QUAN SÁT THÁI DƯƠNG HỆ 3D • ĐỈNH CAO 1 SLIDE" "Outfit" 8.0 $true (Color-Hex "F59E0B") 1

$headTitle = $slide.Shapes.AddTextbox(1, 30, 40, 360, 32)
Set-TextProps $headTitle.TextFrame "Bản Đồ Thiên Thể Tương Tác" "Outfit" 18.0 $true (Color-Hex "FFFFFF") 1

$headSub = $slide.Shapes.AddTextbox(1, 32, 72, 340, 24)
Set-TextProps $headSub.TextFrame "Nhấp vào các nút thiên thể để kích hoạt bảng điều khiển chi tiết trực tiếp." "Outfit" 9.0 $false (Color-Hex "94A3B8") 1

# Interactive Control Toolbar (Right-aligned Buttons)
$btnY = 26.0; $btnH = 26.0; $btnGap = 6.0
$buttons = @(
    @{ Id = "Sun";   Text = "☀️ Mặt Trời";       W = 90.0;  Color = Color-Hex "F59E0B" },
    @{ Id = "Earth"; Text = "🌍 Trái Đất";       W = 95.0;  Color = Color-Hex "38BDF8" },
    @{ Id = "Mars";  Text = "🔴 Sao Hỏa";        W = 90.0;  Color = Color-Hex "EF4444" },
    @{ Id = "Gas";   Text = "🪐 Sao Mộc & Thổ";  W = 120.0; Color = Color-Hex "FBBF24" },
    @{ Id = "Voyager"; Text = "🚀 Voyager 1";    W = 100.0; Color = Color-Hex "A855F7" }
)

$curBtnX = 425.0
$triggerBtnShapes = @{}
foreach ($b in $buttons) {
    $btnSh = $slide.Shapes.AddShape(1, $curBtnX, $btnY, $b.W, $btnH)
    $btnSh.Fill.Solid(); $btnSh.Fill.ForeColor.RGB = Color-Hex "0F172A"
    $btnSh.Line.Visible = -1; $btnSh.Line.ForeColor.RGB = $b.Color; $btnSh.Line.Weight = 1.0
    $btnSh.Shadow.Blur = 6; $btnSh.Shadow.Transparency = 0.7
    Set-TextProps $btnSh.TextFrame $b.Text "Outfit" 8.5 $true $b.Color 2
    $triggerBtnShapes[$b.Id] = $btnSh
    $curBtnX += $b.W + $btnGap
}

# =========================================================================
# CENTRAL 3D CELESTIAL STAGE (8 PLANETS RUNNING PERPETUALLY AT 60 FPS)
# =========================================================================
$cx = 530.0; $cy = 300.0
$tLine = $slide.TimeLine.MainSequence

# 1. 8K Photorealistic 3D Sun
$sunSize = 88.0
$sunPic = $slide.Shapes.AddPicture($imgSun, 0, -1, $cx - ($sunSize / 2.0), $cy - ($sunSize / 2.0), $sunSize, $sunSize)

# 2. Planetary Orbits & Animated Spheres
$pDefs = @(
    @{ Name = "Sao Thủy"; A = 52;  B = 24;  R = 6.0;  Color = Color-Hex "94A3B8"; Dur = 12.0; Th0 = 0.5 },
    @{ Name = "Sao Kim";   A = 78;  B = 35;  R = 9.0;  Color = Color-Hex "F59E0B"; Dur = 20.0; Th0 = 1.9 },
    @{ Name = "Trái Đất"; A = 110; B = 48;  R = 10.0; Color = Color-Hex "38BDF8"; Dur = 30.0; Th0 = 3.3; HasMoon = $true },
    @{ Name = "Sao Hỏa";  A = 145; B = 64;  R = 8.5;  Color = Color-Hex "EF4444"; Dur = 45.0; Th0 = 4.7 },
    @{ Name = "Sao Mộc";  A = 215; B = 96;  R = 22.0; Color = Color-Hex "D97706"; Dur = 75.0; Th0 = 1.1 },
    @{ Name = "Sao Thổ";  A = 275; B = 122; R = 18.0; Color = Color-Hex "FBBF24"; Dur = 105.0; Th0 = 2.5; HasRing = $true },
    @{ Name = "Sao Thiên Vương"; A = 330; B = 148; R = 13.0; Color = Color-Hex "22D3EE"; Dur = 145.0; Th0 = 3.4 },
    @{ Name = "Sao Hải Vương";  A = 385; B = 172; R = 12.5; Color = Color-Hex "3B82F6"; Dur = 185.0; Th0 = 2.2 }
)

foreach ($p in $pDefs) {
    # Orbit Path
    $orb = $slide.Shapes.AddShape(9, $cx - $p.A, $cy - $p.B, $p.A * 2, $p.B * 2)
    $orb.Fill.Visible = 0; $orb.Line.Visible = -1
    $orb.Line.ForeColor.RGB = Color-Hex "38BDF8"; $orb.Line.Weight = 0.75; $orb.Line.Transparency = 0.82

    # Planet Initial Position
    $initX = $cx + $p.A * [Math]::Cos($p.Th0)
    $initY = $cy + $p.B * [Math]::Sin($p.Th0)
    $plSh = $slide.Shapes.AddShape(9, $initX - ($p.R / 2.0), $initY - ($p.R / 2.0), $p.R, $p.R)
    $plSh.Fill.Solid(); $plSh.Fill.ForeColor.RGB = $p.Color; $plSh.Line.Visible = 0
    $plSh.ThreeD.BevelTopType = 6; $plSh.ThreeD.BevelTopInset = $p.R / 2.0; $plSh.ThreeD.BevelTopDepth = $p.R / 2.0

    # Saturn Ring
    if ($p.HasRing) {
        $rg = $slide.Shapes.AddShape(9, $initX - 20, $initY - 8, 40, 16)
        $rg.Fill.Visible = 0; $rg.Line.Visible = -1
        $rg.Line.ForeColor.RGB = Color-Hex "FDE68A"; $rg.Line.Weight = 1.75; $rg.Line.Transparency = 0.35
    }

    # Motion Path Animation (Uniform Velocity, Closed Loop)
    $eff = $tLine.AddEffect($plSh, 86, 0, 2)
    $eff.Timing.Duration = $p.Dur
    $eff.Timing.RepeatCount = 9999
    $eff.Timing.RepeatDuration = 99999
    $eff.Timing.Accelerate = 0.0; $eff.Timing.Decelerate = 0.0
    $eff.Timing.SmoothStart = 0; $eff.Timing.SmoothEnd = 0
    $eff.Behaviors.Item(1).MotionEffect.Path = Get-BezierEllipsePath $cx $cy $p.A $p.B $initX $initY $p.Th0
}

# Asteroid Belt (48 particles)
$astRand = New-Object System.Random(999)
$astCols = @((Color-Hex "D8B4FE"), (Color-Hex "F472B6"), (Color-Hex "FDE68A"), (Color-Hex "94A3B8"))
for ($a = 0; $a -lt 48; $a++) {
    $ath = $astRand.NextDouble() * [Math]::PI * 2.0
    $ar = 168.0 + ($astRand.NextDouble() * 32.0)
    $ax = $cx + $ar * [Math]::Cos($ath)
    $ay = $cy + ($ar * 0.44) * [Math]::Sin($ath) + (($astRand.NextDouble() - 0.5) * 6.0)
    $ac = $astCols[$astRand.Next(0, $astCols.Count)]
    $ad = $slide.Shapes.AddShape(9, $ax, $ay, 2.0, 2.0)
    $ad.Fill.Solid(); $ad.Fill.ForeColor.RGB = $ac; $ad.Line.Visible = 0
}

# Voyager 1 distant probe icon
$vyPic = $slide.Shapes.AddPicture($imgVoyager, 0, -1, 885, 435, 48, 48)

# =========================================================================
# INTERACTIVE DRAWERS & STATE MACHINES (TRIGGERED ON CLICK)
# =========================================================================
# Drawer dimensions
$drwW = 340.0; $drwH = 430.0; $drwX = 600.0; $drwY = 75.0

function Create-Drawer($id, $title, $subTitle, $themeHex, $imgFile, $statsList, $descText) {
    $groupShapes = @()

    # Backdrop Glassmorphic Card (100% Solid Deep Space Navy to prevent bleed-through)
    $dCard = $slide.Shapes.AddShape(1, $drwX, $drwY, $drwW, $drwH)
    $dCard.Fill.Solid(); $dCard.Fill.ForeColor.RGB = Color-Hex "080E1A"; $dCard.Fill.Transparency = 0.0
    $dCard.Line.Visible = -1; $dCard.Line.ForeColor.RGB = Color-Hex $themeHex; $dCard.Line.Weight = 1.5
    $dCard.Shadow.Blur = 24; $dCard.Shadow.Transparency = 0.4
    $groupShapes += $dCard

    # Header Tag
    $dTag = $slide.Shapes.AddTextbox(1, $drwX + 16, $drwY + 14, 230, 18)
    Set-TextProps $dTag.TextFrame $subTitle "Outfit" 8.0 $true (Color-Hex $themeHex) 1
    $groupShapes += $dTag

    # Header Title
    $dTtl = $slide.Shapes.AddTextbox(1, $drwX + 14, $drwY + 30, 230, 28)
    Set-TextProps $dTtl.TextFrame $title "Outfit" 16.0 $true (Color-Hex "FFFFFF") 1
    $groupShapes += $dTtl

    # Close Button ("✖ ĐÓNG")
    $dClose = $slide.Shapes.AddShape(1, $drwX + $drwW - 74, $drwY + 16, 58, 22)
    $dClose.Fill.Solid(); $dClose.Fill.ForeColor.RGB = Color-Hex "1E293B"
    $dClose.Line.Visible = -1; $dClose.Line.ForeColor.RGB = Color-Hex "EF4444"; $dClose.Line.Weight = 0.8
    Set-TextProps $dClose.TextFrame "✖ ĐÓNG" "Outfit" 7.5 $true (Color-Hex "EF4444") 2
    $groupShapes += $dClose

    # 3D Visual Asset Preview inside Drawer
    $dImg = $slide.Shapes.AddPicture($imgFile, 0, -1, $drwX + 20, $drwY + 68, 110, 110)
    $groupShapes += $dImg

    # Quick Telemetry Specs
    $specY = $drwY + 70
    for ($k = 0; $k -lt $statsList.Count; $k++) {
        $st = $statsList[$k]
        $sBox = $slide.Shapes.AddTextbox(1, $drwX + 140, $specY + ($k * 26), 185, 24)
        Set-TextProps $sBox.TextFrame "$($st.Key): $($st.Val)" "Outfit" 8.0 $false (Color-Hex "CBD5E1") 1
        $groupShapes += $sBox
    }

    # Description Paragraph
    $dDesc = $slide.Shapes.AddTextbox(1, $drwX + 16, $drwY + 190, $drwW - 32, 220)
    Set-TextProps $dDesc.TextFrame $descText "Outfit" 9.0 $false (Color-Hex "94A3B8") 1
    $groupShapes += $dDesc

    # Group all drawer elements so they animate as a single cohesive unit
    $shapeNames = $groupShapes | ForEach-Object { $_.Name }
    $drawerGroup = $slide.Shapes.Range($shapeNames).Group()

    # Position off-screen right initially or attach trigger entrance
    # Create Interactive Trigger Sequences
    # 1. Trigger Open when corresponding Button is clicked:
    $seqOpen = $slide.TimeLine.InteractiveSequences.Add(-1)
    $effOpen = $seqOpen.AddEffect($drawerGroup, 2, 0, 4) # 2 = msoAnimEffectFly, 4 = msoAnimTriggerOnShapeClick
    $effOpen.Timing.TriggerShape = $triggerBtnShapes[$id]
    $effOpen.Timing.Duration = 0.4
    $effOpen.EffectParameters.Direction = 3 # 3 = From Right

    # 2. Trigger Close when the Close Button is clicked:
    $seqClose = $slide.TimeLine.InteractiveSequences.Add(-1)
    $effClose = $seqClose.AddEffect($drawerGroup, 2, 0, 4) # Fly
    $effClose.Timing.TriggerShape = $dClose
    $effClose.Exit = -1 # msoTrue (Exit effect!)
    $effClose.Timing.Duration = 0.35
    $effClose.EffectParameters.Direction = 4 # 4 = To Right

    return $drawerGroup
}

# 1. Sun Drawer
Create-Drawer "Sun" "Mặt Trời" "TRUNG TÂM NĂNG LƯỢNG THÁI DƯƠNG" "F59E0B" $imgSun @(
    @{ Key = "Đường kính"; Val = "1,392,700 km" },
    @{ Key = "Nhiệt độ lõi"; Val = "15,000,000 °C" },
    @{ Key = "Khối lượng"; Val = "333,000 x Trái Đất" },
    @{ Key = "Quang phổ"; Val = "G2V (sao lùn vàng)" }
) "Mặt Trời chiếm 99.86% tổng khối lượng toàn bộ Hệ Mặt Trời. Quá trình nhiệt hạch hydro thành heli tại lõi sinh ra nguồn năng lượng vô tận nuôi dưỡng mọi sự sống và điều khiển quỹ đạo của toàn bộ các thiên thể."

# 2. Earth Drawer
Create-Drawer "Earth" "Trái Đất & Mặt Trăng" "CÁI NÔI CỦA SỰ SỐNG" "38BDF8" $imgEarth @(
    @{ Key = "Khoảng cách"; Val = "1.00 AU (149.6M km)" },
    @{ Key = "Đường kính"; Val = "12,742 km" },
    @{ Key = "Chu kỳ quỹ đạo"; Val = "365.25 ngày" },
    @{ Key = "Vệ tinh"; Val = "Mặt Trăng (384,400 km)" }
) "Hành tinh duy nhất được biết đến sở hữu nước ở dạng lỏng trên bề mặt, bầu khí quyển giàu oxy bảo vệ sinh quyển khỏi bức xạ vũ trụ, và sự hiện diện của Mặt Trăng giúp ổn định độ nghiêng trục tự quay."

# 3. Mars Drawer
Create-Drawer "Mars" "Sao Hỏa" "HÀNH TINH ĐỎ - ĐÍCH ĐẾN TƯƠNG LAI" "EF4444" $imgMars @(
    @{ Key = "Khoảng cách"; Val = "1.52 AU (227.9M km)" },
    @{ Key = "Đường kính"; Val = "6,779 km (0.53 Trái Đất)" },
    @{ Key = "Nhiệt độ TB"; Val = "-63 °C" },
    @{ Key = "Đặc điểm"; Val = "Đỉnh Olympus Mons 21.9 km" }
) "Sở hữu hệ thống hẻm núi Valles Marineris dài nhất và ngọn núi lửa Olympus Mons cao nhất Hệ Mặt Trời. Sao Hỏa là mục tiêu trọng tâm trong các sứ mệnh tìm kiếm dấu vết sự sống và thuộc địa hóa không gian của nhân loại."

# 4. Gas Giants Drawer
Create-Drawer "Gas" "Sao Mộc & Sao Thổ" "CÁC GÃ KHỔNG LỒ NGOẠI HỆ" "FBBF24" $imgJupiter @(
    @{ Key = "Sao Mộc"; Val = "Đường kính 139,820 km" },
    @{ Key = "Vết Đỏ Lớn"; Val = "Cơn bão khổng lồ 350+ năm" },
    @{ Key = "Sao Thổ"; Val = "Vành đai băng đá 282,000 km" },
    @{ Key = "Số vệ tinh"; Val = "Hơn 240 vệ tinh tự nhiên" }
) "Sao Mộc đóng vai trò như 'máy hút bụi không gian' bảo vệ Trái Đất khỏi các vụ va chạm thiên thạch. Sao Thổ là kiệt tác thiên văn với hệ thống vành đai băng đá phản xạ ánh sáng rực rỡ."

# 5. Voyager Drawer
Create-Drawer "Voyager" "Voyager 1" "SỨ MỆNH BỜ VỰC LIÊN SAO" "A855F7" $imgVoyager @(
    @{ Key = "Phóng năm"; Val = "1977 (NASA)" },
    @{ Key = "Khoảng cách"; Val = "162.5+ AU (24 tỷ km)" },
    @{ Key = "Vị trí"; Val = "Không gian liên sao" },
    @{ Key = "Thông điệp"; Val = "Đĩa ghi vàng Golden Record" }
) "Vật thể nhân tạo bay xa nhất trong lịch sử loài người. Tàu Voyager 1 đã vượt qua ranh giới Nhật mãn (Heliosphere) và đang đơn độc di chuyển trong khoảng không gian sâu thẳm giữa các vì sao."

# Save presentation
$pres.SaveAs($pptxPath)
Start-Sleep -Milliseconds 800

# Export inspection image
$exportDir = Join-Path $PSScriptRoot "..\export"
if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Force -Path $exportDir | Out-Null }
$exportPng = Join-Path $exportDir "max_single_slide_cockpit.png"
$pres.Slides.Item(1).Export($exportPng, "PNG", 1920, 1080)
Write-Host "Exported Single-Slide Maximum preview to: $exportPng"

Write-Host "SUCCESS: Generated 100% Maximum Single-Slide Interactive Presentation at: $pptxPath"
