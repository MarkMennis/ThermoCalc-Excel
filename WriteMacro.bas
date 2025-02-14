Attribute VB_Name = "WriteMacro"
Sub ScheilMacro()
    ScheilMacroWirte (True)
End Sub
Function ScheilMacroWirte(Optional WriteSequence As Boolean)
    Dim filename As String
    If IsMissing(WriteSequence) Then
        WriteSequence = True
    End If
    
    If WriteSequence Then
        ScheilFile = ThisWorkbook.Path & "\Scheil (Sequence).txt"
        DictraRefFile = ThisWorkbook.Path & "\Dictra Segregation Profiles Ref.txt"
        Open DictraRefFile For Output As #20
    Else
        ScheilFile = ThisWorkbook.Path & "\Scheil.txt"
    End If
    
    Open ScheilFile For Output As #2
    
    If Not WriteSequence Then
        Module.Scheil
        Scheil.GlobMin (True)
        Scheil.TempStep
        Scheil.EvaluateSP (True)
        Scheil.Start
        Scheil.DefineSystem
    Else
        Dim SeqMatrix()
        Dim Hub As Worksheet
        Set Hub = ThisWorkbook.Sheets("Hub")
        
        AttIndex = AttributePosition("System", "Element")
        ScanCol = AttIndex(1) - 2
        IncludeCol = ScanCol + 5
        StartRow = AttIndex(0) + 2
        EndRow = AttIndex(0) + 13
        
        For i = StartRow To EndRow
            If UCase(Hub.Cells(i, IncludeCol).value) = "YES" Then
                n = n + 1
            End If
        Next i
        
        ReDim SeqMatrix(1 To n, 1 To 4)
        
        For i = StartRow To EndRow
            If UCase(Hub.Cells(i, IncludeCol).value) = "YES" Then
                j = j + 1
                SeqMatrix(j, 1) = UCase(Hub.Cells(i, ScanCol).value)
                SeqMatrix(j, 2) = Hub.Cells(i, ScanCol + 2).value
                SeqMatrix(j, 3) = Hub.Cells(i, ScanCol + 3).value
                SeqMatrix(j, 4) = Hub.Cells(i, ScanCol + 4).value
            End If
        Next i
        
        Dim currentCombo() As Variant
        ReDim currentCombo(1 To n)
        
        ' Generate combinations
        GenerateCombinationFromMatrix SeqMatrix, n, 1, currentCombo, 2, 20
        
    End If
    
    Close #2
    If WriteSequence Then Close #20
End Function
Sub GenerateCombinationFromMatrix(ByRef paramMatrix As Variant, ByVal n As Integer, ByVal index As Integer, ByRef currentCombo() As Variant, ByVal SchFileNum As Integer, ByVal DicFileNum As Integer)
    Dim value As Variant
    Dim lowerBound As Double, upperBound As Double, stepVal As Double
    
    GBLMin = Scheil.GlobMin(True, True)
    TMPStep = Scheil.TempStep(True)
    SCHStart = Scheil.Start(True)
    SegrName = GetAttribute("Scheil", "Segregation Profile")
    
    If index > n Then
        Dim line As String, i As Integer
        Dim firstElement As Boolean
        firstElement = True
        line = ""
        ' Loop through each list's value in the combination.
        For i = 1 To n
            If currentCombo(i) <> 0 Then
                If Not firstElement Then line = line & " "
                If Not firstElement Then Version = Version & "_"
                line = line & paramMatrix(i, 1) & " " & currentCombo(i)
                Version = Version & paramMatrix(i, 1) & "_" & currentCombo(i)
                firstElement = False
            End If
        Next i
        ' Only print the line if at least one non-zero value exists.
        If line <> "" Then
            Module.Scheil (Version)
            Print #SchFileNum, GBLMin
            Print #SchFileNum, TMPStep
            result = Scheil.EvaluateSP(True, , Version)
            Print #SchFileNum, SCHStart
            Scheil.DefineSystem (line)
            'Print #SchFileNum, line
            'Print #SchFileNum, Version
            
            Print #DicFileNum, SegrName & Version & ".TXT"
        End If
    Else
        lowerBound = paramMatrix(index, 2)
        upperBound = paramMatrix(index, 3)
        stepVal = paramMatrix(index, 4)
        For value = lowerBound To upperBound Step stepVal
            currentCombo(index) = value
            GenerateCombinationFromMatrix paramMatrix, n, index + 1, currentCombo, SchFileNum, DicFileNum
        Next value
    End If
End Sub
