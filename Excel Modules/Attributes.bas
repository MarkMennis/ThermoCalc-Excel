Attribute VB_Name = "Attributes"
Function GetAttribute(Category, feature)
    Dim Hub As Worksheet
    Dim index()
    Set Hub = ThisWorkbook.Sheets("Hub")
    
    index() = AttributePosition(Category, feature)
    GetAttribute = Hub.Cells(index(0), index(1))
End Function
Function TransferAttribute(Category, feature, value)
    Dim Hub As Worksheet
    Dim index()
    Set Hub = ThisWorkbook.Sheets("Hub")
    
    index() = AttributePosition(Category, feature)
    Hub.Cells(index(0), index(1)) = value
End Function
Function AttributePosition(Category, feature) As Variant
    Dim Hub As Worksheet
    Dim Col As Integer
    Dim Search As Boolean
    Dim LastRow As Integer
    
    Set Hub = ThisWorkbook.Sheets("Hub")
    
    ' Get Column of Category
    Search = True
    Do While Search
        Col = Col + 1
        If UCase(Hub.Cells(1, Col).value) = UCase(Category) Or Col = 100 Then
            Search = False
        End If
    Loop
    
    ' Get last search row
    LastRow = Hub.Cells(Hub.Rows.Count, Col).End(xlUp).Row
    
    ' Search for Value
    For i = 2 To LastRow
        prop = Hub.Cells(i, Col).value
        If UCase(prop) = UCase(feature) Then
            AttributePosition = Array(i, Col + 2)
            Exit Function
        End If
    Next i
End Function
