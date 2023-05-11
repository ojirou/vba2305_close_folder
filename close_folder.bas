Attribute VB_Name = "Module1"
Sub close_folder()
    Dim Sh As Object
    Dim w As Object
    Set Sh = CreateObject("Shell.Application")
    For Each w In Sh.Windows
        If (InStr(TypeName(w.document), "IShellFolderView") > 0) Then
            w.Quit
        End If
    Next w
End Sub
