Attribute VB_Name = "Database"
Function Append(Database As String)
    Print #2, "APPEND_DATABASE" & Database
End Function
Function SwitchDatabase(Database As String)
    Print #2, "SWITCH_DATABASE" & Database
End Function
Function DTB()
    Print #2, UCase(GetAttribute("System", "Database"))
End Function
Function DatabaseMob(Optional CheckOnly As Boolean)
    If GetAttribute("System", "Mobility Database") <> "" Then
        If IsMissing(CheckOnly) Then
            If Not CheckOnly Then
                Print #2, UCase(GetAttribute("System", "Mobility Database"))
            End If
        End If
        DatabaseMob = True
    Else
        DatabaseMob = False
    End If
End Function
Function Elements(Optional index As Integer, Optional SingleValue As Boolean)
    Dim SelectElement As Boolean
    Dim Hub As Worksheet
    Set Hub = ThisWorkbook.Sheets("Hub")
    
    SelectElement = False
    If Not IsMissing(SingleValue) Then
        If SingleValue And Not IsMissing(index) Then
            SelectElement = True
        End If
    End If
    
    If Not IsMissing(index) Then
        ID = index
    Else
        ID = 1
    End If
    
    AttIndex = AttributePosition("System", "Element")
    ScanCol = AttIndex(1) - 2
    StartRow = AttIndex(0) + ID
    EndRow = AttIndex(0) + 13
    
    If SelectElement Then
        Compos = UCase(Hub.Cells(StartRow, ScanCol).value)
    Else
        For i = StartRow To EndRow
            Element = Hub.Cells(i, ScanCol).value
            Include = Hub.Cells(i, ScanCol + 5).value
            If UCase(Element) <> "NONE" And UCase(Include) = "YES" Then
                Compos = Compos & UCase(Element) & " "
            End If
        Next i
    End If
    
    Print #2, Compos
End Function
Function Composition(Optional index As Integer, Optional SingleValue As Boolean)
    Dim SelectElement As Boolean
    Dim Hub As Worksheet
    Set Hub = ThisWorkbook.Sheets("Hub")
    
    SelectElement = False
    If Not IsMissing(SingleValue) Then
        If SingleValue And Not IsMissing(index) Then
            SelectElement = True
        End If
    End If
    
    If Not IsMissing(index) Then
        ID = index
    Else
        ID = 2
    End If
    
    AttIndex = AttributePosition("System", "Element")
    ScanCol = AttIndex(1) - 2
    StartRow = AttIndex(0) + ID
    EndRow = AttIndex(0) + 13
    
    If SelectElement Then
        Element = Hub.Cells(StartRow, ScanCol).value
        Amount = Hub.Cells(StartRow, ScanCol + 1).value
        Compos = UCase(Element) & " " & CStr(Amount)
    Else
        For i = StartRow To EndRow
            Element = Hub.Cells(i, ScanCol).value
            Amount = Hub.Cells(i, ScanCol + 1).value
            Include = Hub.Cells(i, ScanCol + 5).value
            If UCase(Element) <> "NONE" And UCase(Include) = "YES" And Amount > 0 Then
                Compos = Compos & UCase(Element) & " " & CStr(Amount) & " "
            End If
        Next i
    End If
    
    Print #2, Compos
End Function
Function Phases(Optional index As Integer, Optional SingleValue As Boolean)
    Dim SelectElement As Boolean
    Dim Hub As Worksheet
    Set Hub = ThisWorkbook.Sheets("Hub")
    
    SelectElement = False
    If Not IsMissing(SingleValue) Then
        If SingleValue And Not IsMissing(index) Then
            SelectElement = True
        End If
    End If
    
    If Not IsMissing(index) Then
        ID = index
    Else
        ID = 1
    End If
    
    AttIndex = AttributePosition("System", "Phase")
    ScanCol = AttIndex(1) - 2
    StartRow = AttIndex(0) + ID
    EndRow = AttIndex(0) + 13
    
    If SelectElement Then
        Element = Hub.Cells(StartRow, ScanCol).value
        Ph = UCase(Element) & " " & CStr(Amount)
    Else
        For i = StartRow To EndRow
            Element = Hub.Cells(i, ScanCol).value
            Include = Hub.Cells(i, ScanCol + 2).value
            If UCase(Element) <> "NONE" And UCase(Include) = "YES" Then
                Ph = Ph & UCase(Element) & " "
            End If
        Next i
    End If
    
    Print #2, Ph
End Function
Function DefineElements()
    Print #2, "DEFINE_ELEMENTS"
    Database.Elements
End Function
Function RejectPhases(Phases)
    Print #2, "REJ PH " & Phases
End Function
Function RestorePhases()
    Print #2, "RESTORE PHASES"
    Database.Phases
End Function
Function GetData()
    Print #2, "GET_DATA"
End Function
Function Define()
    '===================================
    '====        System Data        ====
    '===================================
    Module.Database
    Database.SwitchDatabase
    Database.DefineSystem
    Database.RejectPhases ("*")
    Database.RestorePhases
    Database.GetData
    
    '===================================
    '====        Mobily Data        ====
    '===================================
    If System.DatabaseMob(True) Then
        Database.Append
        Database.DatabaseMob
        Database.DefineSystem
        Database.RejectPhases ("*")
        Database.RestorePhases
        Database.GetData
    End If
End Function

