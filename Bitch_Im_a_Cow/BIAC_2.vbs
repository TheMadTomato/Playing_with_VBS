Set objShell = CreateObject("WScript.Shell")

' ASCII art of a cow saying "Bitch, I'm a Cow"
cowsay = " _______"
cowsay = cowsay & vbCrLf & "< Bitch, I'm a Cow >"
cowsay = cowsay & vbCrLf & " -------"
cowsay = cowsay & vbCrLf & "        \   ^__^"
cowsay = cowsay & vbCrLf & "         \  (oo)\_______"
cowsay = cowsay & vbCrLf & "            (__)\       )\/\"
cowsay = cowsay & vbCrLf & "                ||----w |"
cowsay = cowsay & vbCrLf & "                ||     ||"

' Display the cow saying "Bitch, I'm a Cow"
MsgBox cowsay, vbInformation, "Cowsay"

' Loop to prevent closing the script by clicking exit
Do While True
    WScript.Sleep 1000 ' Sleep for 1 second
Loop
