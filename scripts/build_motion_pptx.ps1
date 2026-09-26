param(
    [string]$specPath = "$PSScriptRoot\..\examples\solar_system_spec.json",
    [string]$outPath = "presentation_motion.pptx"
)

$ErrorActionPreference = "Stop"

# Load JSON Specification
if (-not (Test-Path $specPath)) {
    Write-Error "Spec file not found at: $specPath"
    exit 1
}

$spec = Get-Content -Raw -Path $specPath -Encoding UTF8 | ConvertFrom-Json

# Launch PowerPoint COM
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = 1
$pres = $ppt.Presentations.Add()

$W = if ($spec.layout.slideWidth) { [double]$spec.layout.slideWidth } else { 960.0 }
$H = if ($spec.layout.slideHeight) { [double]$spec.layout.slideHeight } else { 540.0 }

$pres.PageSetup.SlideWidth = $W
$pres.PageSetup.SlideHeight = $H

$ci = [System.Globalization.CultureInfo]::InvariantCulture
$kappa = (4.0 / 3.0) * ([Math]::Sqrt(2.0) - 1.0)

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

Write-Host "Building presentation from $specPath..."

foreach ($sData in $spec.slides) {
    $slide = $pres.Slides.Add($pres.Slides.Count + 1, 12) # Blank
    $slide.Background.Fill.Solid()
    $slide.Background.Fill.ForeColor.RGB = [int]$spec.theme.backgroundColor

    $pLeft = [double]$spec.layout.hudLeft
    $pWidth = [double]$spec.layout.hudWidth

    # Badge
    $badge = $slide.Shapes.AddShape(1, $pLeft, 32, 215, 22)
    $badge.Fill.Solid(); $badge.Fill.ForeColor.RGB = 0x2A1A0F; $badge.Fill.Transparency = 0.2
    $badge.Line.Visible = -1; $badge.Line.ForeColor.RGB = 0x6E4A35; $badge.Line.Weight = 1.0
    $badge.TextFrame.TextRange.Text = $sData.step
    $badge.TextFrame.TextRange.Font.Name = $spec.theme.primaryFont
    $badge.TextFrame.TextRange.Font.Size = 9
    $badge.TextFrame.TextRange.Font.Bold = -1
    $badge.TextFrame.TextRange.Font.Color.RGB = 0xF8BD38

    # Title
    $titleBox = $slide.Shapes.AddTextbox(1, $pLeft, 56, 385, 34)
    $titleBox.TextFrame.TextRange.Text = $sData.title
    $titleBox.TextFrame.TextRange.Font.Name = $spec.theme.primaryFont
    $titleBox.TextFrame.TextRange.Font.Size = 19.5
    $titleBox.TextFrame.TextRange.Font.Bold = -1
    $titleBox.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF

    # Desc
    $descBox = $slide.Shapes.AddTextbox(1, $pLeft, 92, $pWidth, 46)
    $descBox.TextFrame.TextRange.Text = $sData.desc
    $descBox.TextFrame.TextRange.Font.Name = $spec.theme.primaryFont
    $descBox.TextFrame.TextRange.Font.Size = 11.0
    $descBox.TextFrame.TextRange.Font.Color.RGB = 0xB8A394
    $descBox.TextFrame.WordWrap = -1

    # Cards
    $numCards = $sData.cards.Count
    $startY = 145.0
    $cardSpacing = 8.0
    $cardHeight = 108.0

    for ($cIdx = 0; $cIdx -lt $numCards; $cIdx++) {
        $c = $sData.cards[$cIdx]
        $cY = $startY + ($cIdx * ($cardHeight + $cardSpacing))

        $cBg = $slide.Shapes.AddShape(1, $pLeft, $cY, $pWidth, $cardHeight)
        $cBg.Fill.Solid(); $cBg.Fill.ForeColor.RGB = 0x1A110A; $cBg.Fill.Transparency = 0.08
        $cBg.Line.Visible = -1; $cBg.Line.ForeColor.RGB = [int]$c.color; $cBg.Line.Weight = 1.0

        $cTag = $slide.Shapes.AddShape(1, $pLeft + 10, $cY + 8, 115, 18)
        $cTag.Fill.Solid(); $cTag.Fill.ForeColor.RGB = 0x2A1A0F; $cTag.Fill.Transparency = 0.2
        $cTag.Line.Visible = 0
        $cTag.TextFrame.TextRange.Text = $c.tag
        $cTag.TextFrame.TextRange.Font.Size = 7.5
        $cTag.TextFrame.TextRange.Font.Bold = -1
        $cTag.TextFrame.TextRange.Font.Color.RGB = [int]$c.color

        $cName = $slide.Shapes.AddTextbox(1, $pLeft + 130, $cY + 7, $pWidth - 138, 20)
        $cName.TextFrame.TextRange.Text = "$($c.name) • $($c.stat)"
        $cName.TextFrame.TextRange.Font.Size = 9.5
        $cName.TextFrame.TextRange.Font.Bold = -1
        $cName.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF

        $cDet = $slide.Shapes.AddTextbox(1, $pLeft + 10, $cY + 28, $pWidth - 18, $cardHeight - 34)
        $cDet.TextFrame.TextRange.Text = $c.detail
        $cDet.TextFrame.TextRange.Font.Size = 10.0
        $cDet.TextFrame.TextRange.Font.Color.RGB = 0xCBD5E1
        $cDet.TextFrame.WordWrap = -1
    }

    # Orbits & Moving Entities
    $cx = [double]$spec.layout.celestialCenter.x
    $cy = [double]$spec.layout.celestialCenter.y

    # Central Sun
    $sun = $slide.Shapes.AddShape(9, $cx - 19, $cy - 19, 38, 38)
    $sun.Fill.Solid(); $sun.Fill.ForeColor.RGB = 0x0B9EF5; $sun.Line.Visible = 0

    foreach ($orb in $sData.orbits) {
        $a = [double]$orb.a; $b = [double]$orb.b; $diam = [double]$orb.size
        $color = [int]$orb.color; $dur = [double]$orb.duration
        $deg = [double]$orb.startAngle; $th0 = ($deg * [Math]::PI) / 180.0

        # Orbit track
        $orbLine = $slide.Shapes.AddShape(9, $cx - $a, $cy - $b, $a * 2, $b * 2)
        $orbLine.Fill.Visible = 0; $orbLine.Line.Visible = -1
        $orbLine.Line.ForeColor.RGB = 0x6E4A35; $orbLine.Line.Weight = 1.0; $orbLine.Line.Transparency = 0.35

        # Moving planet
        $initX = $cx + $a * [Math]::Cos($th0)
        $initY = $cy + $b * [Math]::Sin($th0)
        $pSh = $slide.Shapes.AddShape(9, $initX - ($diam / 2), $initY - ($diam / 2), $diam, $diam)
        $pSh.Fill.Solid(); $pSh.Fill.ForeColor.RGB = $color; $pSh.Line.Visible = 0

        # Animation Path
        $path = Get-BezierEllipsePath $cx $cy $a $b $initX $initY $th0
        $eff = $slide.TimeLine.MainSequence.AddEffect($pSh, 1, 0, 2)
        $eff.Timing.Duration = $dur
        $eff.Timing.RepeatCount = 1000
        $eff.Timing.SmoothStart = 0; $eff.Timing.SmoothEnd = 0
        $eff.Behaviors.Add(1).MotionEffect.Path = $path

        # Depth Scaling
        $effS = $slide.TimeLine.MainSequence.AddEffect($pSh, 1, 0, 2)
        $effS.Timing.Duration = $dur / 2.0
        $effS.Timing.RepeatCount = 1000
        $effS.Timing.AutoReverse = -1
        $sBeh = $effS.Behaviors.Add(3)
        $sBeh.ScaleEffect.ByX = 125
        $sBeh.ScaleEffect.ByY = 125
    }
}

$finalOut = (Resolve-Path .).Path + "\" + $outPath
$pres.SaveAs($finalOut)
Write-Host "Presentation generated successfully at: $finalOut"
