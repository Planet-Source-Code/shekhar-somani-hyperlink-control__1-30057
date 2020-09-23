Attribute VB_Name = "modShell"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function OpenFile(FilePath, OwnerHWnd As Long, StartupDirectory As String, nShowCmd As Long) As Long
    OpenFile = ShellExecute(OwnerHWnd, "Open", FilePath, vbNullString, StartupDirectory, nShowCmd)
End Function
