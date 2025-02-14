Attribute VB_Name = "Scheil"
Function GlobMin(Active As Boolean, Optional ReadOnly As Boolean)
    Dim Switch As String
    If Active Then
        Switch = " YES"
    Else
        Switch = " NO"
    End If
    If Not IsMissing(ReadOnly) Then
        If Not ReadOnly Then Print #2, "GLOBAL" & Switch
    End If
    GlobMin = "GLOBAL" & Switch
End Function
Function TempStep(Optional ReadOnly As Boolean)
    Dim Step As Double
    Step = GetAttribute("Scheil", "Temp Step")
    TempStep = ""
    If Step > 0 Then
        If Not IsMissing(ReadOnly) Then
            If Not ReadOnly Then Print #2, "TEMPERATURE_STEP" & " " & CStr(Step)
        End If
        TempStep = "TEMPERATURE_STEP" & " " & CStr(Step)
    End If
End Function
Function EvaluateSP(Active As Boolean, Optional ReadOnly As Boolean, Optional FileID)
    Print #2, "EVALUATE_SEGREGATION_PROFILE"
    If Active Then
        Print #2, "YES"
        If IsMissing(FileID) Then
            FileID = ""
        Else
            FileID = "_" & FileID
        End If
        Print #2, GetAttribute("Dictra", "Grid Points") & " " & GetAttribute("Scheil", "Segregation Profile") & FileID & ".TXT" & " " & "YES"
    Else
        Print #2, "NO"
    End If
End Function
Function Start(Optional ReadOnly As Boolean)
    If Not IsMissing(ReadOnly) Then
        If Not ReadOnly Then Print #2, "START_WIZARD"
    End If
    Start = "START_WIZARD"
End Function
Function DefineSystem(Optional SeqComposition)
    ' Select Database
    Database.DTB
    ' Enter Composition
    result = Database.Elements(1, True)
    Print #2, "YES"
    If IsMissing(SeqComposition) Then
        result = Database.Composition(2)
    Else
        Print #2, SeqComposition
    End If
    Print #2, ""
    ' Enter Starting Temperature
    Print #2, CStr(GetAttribute("Scheil", "Temperature"))
    ' Enter Phases
    Print #2, "*"
    Database.Phases
    Print #2, UCase(CStr(GetAttribute("Scheil", "Retain All Phases")))
    ' Miscellaneous
    Print #2, UCase(CStr(GetAttribute("Scheil", "Miscibility Gap Check")))
    Print #2, "NONE" 'Fast Diffusing Elements (No input support yet)
End Function
Function Plots()

End Function
