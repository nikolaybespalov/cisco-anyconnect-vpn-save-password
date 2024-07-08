Dim USER
Dim PASSWORD

USER = "SET_YOUR_USERNAME_HERE"
PASSWORD = "SET_YOUR_PASSWORD_HERE"  'if there is a double quote replace with " & CHR(34) & "

'encode password string compatible with sendkey
sPasswordToSend = ""
For ii = 1 To Len( PASSWORD)
  sChar = Mid( PASSWORD, ii, 1)
  Select Case sChar
  Case "{", "}", "(", ")", "[", "]", "^", "%", "+", "~"
    sPasswordToSend = sPasswordToSend & "{" & sChar & "}"
  Case """"
    sPasswordToSend = sPasswordToSend & sChar & " " 'adding a space after the doublequote, guess there are more char which need this if you have US kayboard
  Case Else
    sPasswordToSend = sPasswordToSend & sChar
  End Select
Next


Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.Run """%PROGRAMFILES(x86)%\Cisco\Cisco Secure Client\UI\csc_ui.exe"""

WScript.Sleep 1000

WshShell.AppActivate "Cisco Secure Client"

WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{ENTER}"

WScript.Sleep 1000
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{ENTER}"

WScript.Sleep 2000
WshShell.SendKeys PASSWORD
WshShell.SendKeys "+{TAB}"
WshShell.SendKeys USER
WshShell.SendKeys "{ENTER}"
WScript.Sleep 100

Function ClipBoard(input)
  If IsNull(input) Then
    ClipBoard = CreateObject("HTMLFile").parentWindow.clipboardData.getData("Text")
    If IsNull(ClipBoard) Then ClipBoard = ""
  Else
    CreateObject("WScript.Shell").Run _
      "mshta.exe javascript:eval(""document.parentWindow.clipboardData.setData('text','" _
      & Replace(Replace(Replace(input, "'", "\\u0027"), """","\\u0022"),Chr(13),"\\r\\n") & "');window.close()"")", _
      0,True
  End If
End Function

ClipBoard(PASSWORD)
