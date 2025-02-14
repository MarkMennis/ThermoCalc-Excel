Attribute VB_Name = "Module"
Function Scheil(Optional VersionID)
    Print #2, "@@"
    If IsMissing(VersionID) Then
        VersionID = ""
    Else
        VersionID = "_" & VersionID
    End If
    Print #2, "@@ Scheil" & VersionID
    Print #2, "@@"
    Print #2, ""
    Print #2, "GO SCHEIL"
End Function
Function Database()
    Print #2, "GOTO_MODULE DA"
End Function

