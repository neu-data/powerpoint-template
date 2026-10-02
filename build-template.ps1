# Builds the Neudata PowerPoint template (.potx) with a real slide master and layouts,
# plus an example deck showing every layout.
# Requires Microsoft PowerPoint (Windows) and Python 3.
# Run:  powershell -ExecutionPolicy Bypass -File build-template.ps1

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$logo = Join-Path $here "neudata-logo.png"
$outPotx = Join-Path $here "Neudata-Template.potx"
$outPptx = Join-Path $here "Neudata-Template-example.pptx"
$LOGO_RATIO = 750 / 520            # width / height of neudata-logo.png

function RGB([string]$hex) {
  $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
  $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
  $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
  return $r + 256 * $g + 65536 * $b
}

$TEAL = RGB "0D7377"; $BLUE = RGB "0070C0"; $GREY = RGB "8A99A3"; $INK = RGB "1A1A1A"; $WHITE = RGB "FFFFFF"
$HEAD = "Century Gothic"; $BODY = "Century Gothic"; $FOOT = "Verdana"
$TAGLINE = "Neudata  |  #ClearDataClearImpact"

# Placeholder type ids
$PH_TITLE = 1; $PH_BODY = 2; $PH_CTITLE = 3; $PH_SUBTITLE = 4; $PH_OBJECT = 7
$PH_SLIDENUM = 13; $PH_FOOTER = 15; $PH_DATE = 16

$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Add(0)
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540
$master = $pres.SlideMaster
# Theme colours and fonts are written into the saved file by set-theme.py
# (PowerPoint COM does not expose the theme on every Office install).

# ---- Helpers ------------------------------------------------------------------------
function Shapes-Of($container) { $list = @(); for ($i = 1; $i -le $container.Shapes.Count; $i++) { $list += $container.Shapes.Item($i) }; return $list }
function Ph($container, [int[]]$types) {
  foreach ($s in (Shapes-Of $container)) { if ($s.Type -eq 14 -and $types -contains $s.PlaceholderFormat.Type) { return $s } }
  return $null
}
function Phs($container, [int[]]$types) {
  $out = @(); foreach ($s in (Shapes-Of $container)) { if ($s.Type -eq 14 -and $types -contains $s.PlaceholderFormat.Type) { $out += $s } }; return $out
}
function RemovePlaceholders($container, [int[]]$types) {
  for ($i = $container.Shapes.Count; $i -ge 1; $i--) {
    $s = $container.Shapes.Item($i)
    if ($s.Type -eq 14 -and $types -contains $s.PlaceholderFormat.Type) { $s.Delete() }
  }
}
function Place($shape, $l, $t, $w, $h) { $shape.Left = $l; $shape.Top = $t; $shape.Width = $w; $shape.Height = $h }
function Layout($name) {
  for ($i = 1; $i -le $master.CustomLayouts.Count; $i++) { $l = $master.CustomLayouts.Item($i); if ($l.Name -eq $name) { return $l } }
  return $null
}
function AddLogo($container, $left, $top, $height, $name) {
  $p = $container.Shapes.AddPicture($logo, 0, -1, $left, $top, $height * $LOGO_RATIO, $height); $p.Name = $name; return $p
}
function AddTagline($container) {
  $t = $container.Shapes.AddTextbox(1, 280, 505, 400, 24); $t.Name = "Neudata footer"
  $r = $t.TextFrame.TextRange; $r.Text = $TAGLINE
  $r.Font.Name = $FOOT; $r.Font.Size = 11; $r.Font.Color.RGB = $TEAL; $r.ParagraphFormat.Alignment = 2
}
function StyleTitle($shape, $size) {
  $r = $shape.TextFrame.TextRange
  $r.Font.Name = $HEAD; $r.Font.Size = $size; $r.Font.Bold = -1; $r.Font.Color.RGB = $TEAL
  $r.ParagraphFormat.Alignment = 1; $shape.TextFrame.VerticalAnchor = 4
}
function StyleSlideNumber($shape) {
  Place $shape 860 505 75 24
  $r = $shape.TextFrame.TextRange; $r.Font.Name = $FOOT; $r.Font.Size = 10; $r.Font.Color.RGB = $GREY
  $r.ParagraphFormat.Alignment = 3
}
# Standard content chrome: title above a teal rule, small logo, tagline, slide number
function ContentChrome($layout) {
  RemovePlaceholders $layout @($PH_DATE, $PH_FOOTER)
  $t = Ph $layout @($PH_TITLE); if ($t) { Place $t 30 18 900 58; StyleTitle $t 30 }
  $n = Ph $layout @($PH_SLIDENUM); if ($n) { StyleSlideNumber $n }
  $rule = $layout.Shapes.AddLine(28, 82, 932, 82); $rule.Line.ForeColor.RGB = $TEAL; $rule.Line.Weight = 1.25; $rule.Name = "Title rule"
  AddLogo $layout 12 494 40 "Neudata logo" | Out-Null
  AddTagline $layout
}
function StyleBody($shape) {
  $r = $shape.TextFrame.TextRange
  $r.Font.Name = $BODY; $r.Font.Color.RGB = $INK; $r.Font.Size = 20
  $r.ParagraphFormat.Bullet.Font.Color.RGB = $TEAL
}

# ---- Slide master: text styles only (decoration lives on the layouts) ----------------
$master.Background.Fill.Solid(); $master.Background.Fill.ForeColor.RGB = $WHITE
RemovePlaceholders $master @($PH_DATE, $PH_FOOTER)
$mt = Ph $master @($PH_TITLE); Place $mt 30 18 900 58; StyleTitle $mt 30
$mb = Ph $master @($PH_BODY); Place $mb 34 100 892 370; StyleBody $mb
$mn = Ph $master @($PH_SLIDENUM); StyleSlideNumber $mn

# ---- Layouts ---------------------------------------------------------------------------
foreach ($n in @("Comparison", "Content with Caption", "Picture with Caption", "Title and Vertical Text",
                 "Vertical Title and Text", "Content", "Last Slide")) {
  $x = Layout $n; if ($x) { $x.Delete() }
}

# Title Slide: centred title and presenter, large logo bottom-left, date bottom-right
$ts = Layout "Title Slide"
$ts.DisplayMasterShapes = 0
RemovePlaceholders $ts @($PH_FOOTER, $PH_SLIDENUM)
$t = Ph $ts @($PH_CTITLE); Place $t 80 115 800 105; StyleTitle $t 44; $t.TextFrame.TextRange.ParagraphFormat.Alignment = 2
$s = Ph $ts @($PH_SUBTITLE); Place $s 180 232 600 90
$r = $s.TextFrame.TextRange; $r.Font.Size = 20; $r.Font.Color.RGB = $INK; $r.ParagraphFormat.Alignment = 2
$d = Ph $ts @($PH_DATE)
if ($d) { Place $d 640 438 280 30; $r = $d.TextFrame.TextRange; $r.Font.Name = $BODY; $r.Font.Size = 14; $r.Font.Bold = -1; $r.Font.Color.RGB = $INK; $r.ParagraphFormat.Alignment = 3 }
AddLogo $ts 40 345 165 "Neudata logo large" | Out-Null

# Section Header: logo, vertical teal rule, section title and number
$sh = Layout "Section Header"
$sh.DisplayMasterShapes = 0
RemovePlaceholders $sh @($PH_DATE, $PH_FOOTER)
$t = Ph $sh @($PH_TITLE); Place $t 390 185 540 85; StyleTitle $t 40
$b = Ph $sh @($PH_BODY); Place $b 390 275 540 45
$r = $b.TextFrame.TextRange; $r.Font.Size = 18; $r.Font.Color.RGB = $GREY; $r.ParagraphFormat.Bullet.Visible = 0
$n = Ph $sh @($PH_SLIDENUM); if ($n) { StyleSlideNumber $n }
AddLogo $sh 60 190 140 "Neudata logo" | Out-Null
$vr = $sh.Shapes.AddLine(360, 175, 360, 345); $vr.Line.ForeColor.RGB = $TEAL; $vr.Line.Weight = 2; $vr.Name = "Section rule"
AddTagline $sh

# Content layouts share the standard chrome
$tc = Layout "Title and Content"; ContentChrome $tc
$b = Ph $tc @($PH_BODY, $PH_OBJECT); Place $b 34 100 892 370

$two = Layout "Two Content"; ContentChrome $two
$cols = Phs $two @($PH_BODY, $PH_OBJECT)
if ($cols.Count -ge 2) { Place $cols[0] 34 100 436 370; Place $cols[1] 490 100 436 370 }

$to = Layout "Title Only"; ContentChrome $to
$bl = Layout "Blank"; RemovePlaceholders $bl @($PH_DATE, $PH_FOOTER)
$n = Ph $bl @($PH_SLIDENUM); if ($n) { StyleSlideNumber $n }
AddLogo $bl 12 494 40 "Neudata logo" | Out-Null; AddTagline $bl

# Title and Notes: title plus a blue note line (as in the original template)
$nl = $master.CustomLayouts.Add($tc.Index + 1); $nl.Name = "Title and Notes"
RemovePlaceholders $nl @($PH_BODY, $PH_OBJECT)
ContentChrome $nl
$nb = $nl.Shapes.AddPlaceholder($PH_BODY, 34, 100, 892, 370)
$r = $nb.TextFrame.TextRange; $r.Font.Size = 28; $r.Font.Color.RGB = $BLUE; $r.ParagraphFormat.Bullet.Visible = 0

# Closing - Copyright
$cl = $master.CustomLayouts.Add($master.CustomLayouts.Count + 1); $cl.Name = "Closing - Copyright"
RemovePlaceholders $cl @($PH_TITLE, $PH_DATE, $PH_FOOTER)
$cl.DisplayMasterShapes = 0
$n = Ph $cl @($PH_SLIDENUM); if ($n) { StyleSlideNumber $n }
AddLogo $cl (480 - 190 * $LOGO_RATIO / 2) 25 190 "Neudata logo large" | Out-Null   # centred
$ch = $cl.Shapes.AddTextbox(1, 150, 232, 700, 30); $ch.Name = "Copyright heading"
$r = $ch.TextFrame.TextRange; $r.Text = ("Copyright " + [char]0x00A9 + " Neudata, " + (Get-Date).Year)
$r.Font.Name = $HEAD; $r.Font.Size = 20; $r.Font.Bold = -1; $r.Font.Color.RGB = $BLUE
$q1 = [string][char]0x201C; $q2 = [string][char]0x201D
$legal = ("All property rights and copyright are reserved.`r`r" +
  "This presentation contains proprietary data, information, and materials developed by Neudata for dedicated use only " +
  "and may not be communicated, copied, reproduced, distributed, published, or cited, in whole or in part, without the " +
  "prior written consent of Neudata. Where explicit written permission has been granted, appropriate attribution must be " +
  "included, followed by " + $q1 + "by courtesy of Neudata" + $q2 + ".`r`r" +
  "Any unauthorized use or infringement of these materials may give rise to legal action and claims for damages, without " +
  "prejudice to any other rights of Neudata, including rights relating to patents, trademarks, or other forms of " +
  "intellectual property protection.")
$cb = $cl.Shapes.AddTextbox(1, 150, 268, 700, 215); $cb.Name = "Copyright notice"
$cb.TextFrame.WordWrap = -1
$r = $cb.TextFrame.TextRange; $r.Text = $legal; $r.Font.Name = "Consolas"; $r.Font.Size = 11; $r.Font.Color.RGB = $INK
AddTagline $cl

# ---- Example slides (one per layout) ------------------------------------------------------
function AddSlide($layoutName) {
  $lay = Layout $layoutName
  if ($null -eq $lay) { throw "Layout not found: $layoutName" }
  $sl = $pres.Slides.AddSlide($pres.Slides.Count + 1, $lay)
  $sl.HeadersFooters.SlideNumber.Visible = -1
  return $sl
}

$s1 = AddSlide "Title Slide"
(Ph $s1 @($PH_CTITLE)).TextFrame.TextRange.Text = "Presentation Title"
(Ph $s1 @($PH_SUBTITLE)).TextFrame.TextRange.Text = "Presenter Name, Role`rOrganisation"
$s1.HeadersFooters.DateAndTime.Visible = -1
$s1.HeadersFooters.DateAndTime.UseFormat = 0
$s1.HeadersFooters.DateAndTime.Text = "Date: DD-MM-YYYY"
$s1.HeadersFooters.SlideNumber.Visible = 0

$s2 = AddSlide "Section Header"
(Ph $s2 @($PH_TITLE)).TextFrame.TextRange.Text = "Section title"
(Ph $s2 @($PH_BODY)).TextFrame.TextRange.Text = "Section 1"

$s3 = AddSlide "Title and Content"
(Ph $s3 @($PH_TITLE)).TextFrame.TextRange.Text = "Title and content"
(Ph $s3 @($PH_BODY, $PH_OBJECT)).TextFrame.TextRange.Text = "First point`rSecond point`rThird point"

$s4 = AddSlide "Two Content"
(Ph $s4 @($PH_TITLE)).TextFrame.TextRange.Text = "Two columns"
$c = Phs $s4 @($PH_BODY, $PH_OBJECT)
$c[0].TextFrame.TextRange.Text = "Left column point`rAnother point"
$c[1].TextFrame.TextRange.Text = "Right column point`rAnother point"

$s5 = AddSlide "Title and Notes"
(Ph $s5 @($PH_TITLE)).TextFrame.TextRange.Text = "Title"
(Ph $s5 @($PH_BODY)).TextFrame.TextRange.Text = "Notes"

$s6 = AddSlide "Title Only"
(Ph $s6 @($PH_TITLE)).TextFrame.TextRange.Text = "Title only - for charts, tables and diagrams"

$s7 = AddSlide "Closing - Copyright"

# ---- Save -------------------------------------------------------------------------------------
$pres.SaveAs($outPptx, 24)   # ppSaveAsOpenXMLPresentation (example deck)
# The template itself starts empty: master and layouts only, plus a title slide to begin with
for ($i = $pres.Slides.Count; $i -ge 2; $i--) { $pres.Slides.Item($i).Delete() }
$pres.SaveAs($outPotx, 26)   # ppSaveAsOpenXMLTemplate
$pres.Close()
$pp.Quit()

python (Join-Path $here "set-theme.py") $outPotx $outPptx
"Saved: $outPotx"
"Saved: $outPptx"
