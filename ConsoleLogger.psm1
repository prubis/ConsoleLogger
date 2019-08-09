## Shamelessly stolen and slightly adapted from https://devblogs.microsoft.com/powershell/colorized-capture-of-console-screen-in-html-and-rtf/

if ($host.Name -ne 'ConsoleHost') {
  write-host -ForegroundColor Red "This script runs only in the console host. You cannot run this script in $($host.Name)."
  exit -1
}

# The Windows PowerShell console host redefines DarkYellow and DarkMagenta colors and uses them as defaults.
# The redefined colors do not correspond to the color names used in HTML, so they need to be mapped to digital color codes.
#

function ConvertTo-HtmlColor ($color) {

  if ($color -eq "DarkYellow") { $color = "#eeedf0" }
  if ($color -eq "DarkMagenta") { $color = "#012456" }
  if ($color -eq "Green") { $color = "#00ff00"}

  return $color
}

# Create an HTML span from text using the named console colors.
#

function New-HtmlSpan ($text, $forecolor = "DarkYellow", $backcolor = "DarkMagenta") {

  $forecolor = ConvertTo-HtmlColor $forecolor
  $backcolor = ConvertTo-HtmlColor $backcolor

  # You can also add font-weight:bold tag here if you want a bold font in output.

  return "<span style='font-family:Courier New;color:$forecolor;background:$backcolor'>$text</span>"
}

Function Save-ConsoleAsHtml ($TargetFile) {
    # Initialize the HTML string builder.

    $htmlBuilder = new-object System.Text.StringBuilder

    $htmlBuilder.Append("<pre style='MARGIN: 0in 10pt 0in;line-height:normal';font-size:10pt>") | Out-Null

    # Grab the console screen buffer contents using the Host console API.

    $bufferWidth = $host.ui.rawui.BufferSize.Width
    $bufferHeight = $host.ui.rawui.CursorPosition.Y

    $rec = new-object System.Management.Automation.Host.Rectangle 0,0,($bufferWidth - 1),$bufferHeight

    $buffer = $host.ui.rawui.GetBufferContents($rec)

    # Iterate through the lines in the console buffer.

    for($i = 0; $i -lt $bufferHeight; $i++) {

      $spanBuilder = new-object system.text.stringbuilder

      # Track the colors to identify spans of text with the same formatting.

      $currentForegroundColor = $buffer[$i, 0].Foregroundcolor
      $currentBackgroundColor = $buffer[$i, 0].Backgroundcolor

      for($j = 0; $j -lt $bufferWidth; $j++) {

        $cell = $buffer[$i,$j]

        # If the colors change, generate an HTML span and append it to the HTML string builder.

        if (($cell.ForegroundColor -ne $currentForegroundColor) -or ($cell.BackgroundColor -ne $currentBackgroundColor)) {
        
          $spanHtml = New-HtmlSpan $spanBuilder.ToString() $currentForegroundColor $currentBackgroundColor
          $htmlBuilder.Append($spanHtml) | Out-Null

          # Reset the span builder and colors.

          $spanBuilder = new-object system.text.stringbuilder

          $currentForegroundColor = $cell.Foregroundcolor
          $currentBackgroundColor = $cell.Backgroundcolor

        }

        # Substitute characters which have special meaning in HTML.

        switch ($cell.Character)  {
          '>'     { $htmlChar = '&gt;' }
          '<'     { $htmlChar = '&lt;' }
          '&'     { $htmlChar = '&amp;' }
          default {
            $htmlChar = $cell.Character
          }
        }

        $spanBuilder.Append($htmlChar) | Out-Null
      }

      $spanHtml = New-HtmlSpan $spanBuilder.ToString() $currentForegroundColor $currentBackgroundColor
      $htmlBuilder.Append($spanHtml) | Out-Null
      
      $htmlBuilder.Append("&nbsp;<br>") | Out-Null
    }

    # Append HTML ending tag.

    $htmlBuilder.Append("</pre>") | Out-Null

    $htmlBuilder.ToString() | Out-File $TargetFile
}


Function Copy-ConsoleAsHtml ($NumberOfLines) {
    # If we haven't specified how many lines to copy, copy them all
    If ($NumberOfLines -eq $null){ $NumberOfLines = $host.ui.rawui.CursorPosition.Y}

    # Initialize the HTML string builder.

    $htmlBuilder = new-object System.Text.StringBuilder

    $htmlBuilder.Append("<pre style='MARGIN: 0in 10pt 0in;line-height:normal';font-size:10pt>") | Out-Null

    # Grab the console screen buffer contents using the Host console API.

    $bufferWidth = $host.ui.rawui.BufferSize.Width
    $bufferHeight = $host.ui.rawui.CursorPosition.Y
    
    $bufferStartHeight = $host.ui.rawui.CursorPosition.Y - $NumberOfLines

    $rec = new-object System.Management.Automation.Host.Rectangle 0,$bufferStartHeight,($bufferWidth - 1),$bufferHeight

    $buffer = $host.ui.rawui.GetBufferContents($rec)

    # Iterate through the lines in the console buffer.

    for($i = 0; $i -lt $NumberOfLines; $i++) {

      $spanBuilder = new-object system.text.stringbuilder

      # Track the colors to identify spans of text with the same formatting.

      $currentForegroundColor = $buffer[$i, 0].Foregroundcolor
      $currentBackgroundColor = $buffer[$i, 0].Backgroundcolor

      for($j = 0; $j -lt $bufferWidth; $j++) {

        $cell = $buffer[$i,$j]

        # If the colors change, generate an HTML span and append it to the HTML string builder.

        if (($cell.ForegroundColor -ne $currentForegroundColor) -or ($cell.BackgroundColor -ne $currentBackgroundColor)) {
        
          $spanHtml = New-HtmlSpan $spanBuilder.ToString() $currentForegroundColor $currentBackgroundColor
          $htmlBuilder.Append($spanHtml) | Out-Null

          # Reset the span builder and colors.

          $spanBuilder = new-object system.text.stringbuilder

          $currentForegroundColor = $cell.Foregroundcolor
          $currentBackgroundColor = $cell.Backgroundcolor

        }

        # Substitute characters which have special meaning in HTML.

        switch ($cell.Character)  {
          '>'     { $htmlChar = '&gt;' }
          '<'     { $htmlChar = '&lt;' }
          '&'     { $htmlChar = '&amp;' }
          default {
            $htmlChar = $cell.Character
          }
        }

        $spanBuilder.Append($htmlChar) | Out-Null
      }

      $spanHtml = New-HtmlSpan $spanBuilder.ToString() $currentForegroundColor $currentBackgroundColor
      $htmlBuilder.Append($spanHtml) | Out-Null
      
      $htmlBuilder.Append("&nbsp;<br>") | Out-Null
    }

    # Append HTML ending tag.

    $htmlBuilder.Append("</pre>") | Out-Null

    $htmlBuilder.ToString() | Set-Clipboard -AsHtml
}
