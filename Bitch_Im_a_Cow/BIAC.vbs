Set objShell = WScript.CreateObject("WScript.Shell")

' Create a simple window
Set objIE = CreateObject("InternetExplorer.Application")
objIE.Visible = True
objIE.ToolBar = 0
objIE.StatusBar = 0
objIE.AddressBar = 0
objIE.MenuBar = 0
objIE.Width = 400
objIE.Height = 200
objIE.Navigate "about:blank"

' Display a message in the window
objIE.Document.Body.InnerHTML = "<h1>Unkillable Window</h1>" & _
                                "<p>This window will attempt to prevent being closed. " & _
                                "Terminate it from Task Manager if needed.</p>"

' Loop to prevent closing the window by clicking exit
Do While objIE.Visible
    WScript.Sleep 1000 ' Sleep for 1 second
Loop
