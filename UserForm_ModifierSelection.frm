VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_ModifierSelection 
   Caption         =   "Saisissez les modificateurs"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4905
   OleObjectBlob   =   "UserForm_ModifierSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_ModifierSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim wsSel As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    
    Set wsSel = ThisWorkbook.Sheets("Sélection")
    lastRow = wsSel.Cells(wsSel.Rows.Count, 1).End(xlUp).Row
    
    Set rng = wsSel.Range("A2:A" & lastRow)
    For Each cell In rng
        Me.cmbCode.AddItem cell.Value
    Next cell
End Sub


Private Sub cmbCode_Change()
    Dim wsSel As Worksheet
    Dim lastRow As Long
    Dim found As Range
    Dim modList As String
    Dim modifiers() As String
    Dim i As Integer
    
    Set wsSel = ThisWorkbook.Sheets("Sélection")
    lastRow = wsSel.Cells(wsSel.Rows.Count, 1).End(xlUp).Row
    
    Set found = wsSel.Range("A2:A" & lastRow).Find(Me.cmbCode.Value, LookAt:=xlWhole)
    
    If Not found Is Nothing Then
        modList = wsSel.Cells(found.Row, 3).Value
        
        modList = Replace(modList, "[", "")
        modList = Replace(modList, "]", "")
        
        Me.lstModifiers.Clear
        If modList <> "" Then
            modifiers = Split(modList, ",")
            For i = LBound(modifiers) To UBound(modifiers)
                Me.lstModifiers.AddItem Trim(modifiers(i))
            Next i
        End If
    End If
End Sub


Private Sub btnApply_Click()
    Dim wsSel As Worksheet
    Dim found As Range
    Dim selectedModifiers As String
    Dim basePrice As Double
    Dim modifiedPrice As Double
    Dim i As Integer
    
    Set wsSel = ThisWorkbook.Sheets("Sélection")
    If Me.cmbCode.Value = "" Then
        MsgBox "Veuillez sélectionner un code.", vbExclamation
        Exit Sub
    End If
    
    Set found = wsSel.Range("A:A").Find(Me.cmbCode.Value, LookAt:=xlWhole)
    If found Is Nothing Then Exit Sub
    
    basePrice = wsSel.Cells(found.Row, 4).Value
    
    For i = 0 To Me.lstModifiers.ListCount - 1
        If Me.lstModifiers.Selected(i) Then
            If selectedModifiers = "" Then
                selectedModifiers = Me.lstModifiers.List(i)
            Else
                selectedModifiers = selectedModifiers & "," & Me.lstModifiers.List(i)
            End If
        End If
    Next i
    
    If selectedModifiers = "" Then
        selectedModifiers = wsSel.Cells(found.Row, 3).Value
    End If
    
    modifiedPrice = ExtractModPrice(selectedModifiers, basePrice)
    
    wsSel.Cells(found.Row, 3).Value = selectedModifiers
    wsSel.Cells(found.Row, 5).Value = modifiedPrice
    
    Call SortPrixModifie

    Unload Me
End Sub

Function ExtractModPrice(modifiers As String, basePrice As Double) As Double
    Dim ws As Worksheet, lastRow As Long, i As Integer
    Dim modList As Variant, modCode As String, ModPrice As Variant
    Dim ModTotalPrice As Double
    
    Set ws = ThisWorkbook.Sheets("Modifiers")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    modifiers = Replace(Replace(Replace(modifiers, "[", ""), "]", ""), " ", "")
    modList = Split(modifiers, ",")
    
    ModTotalPrice = 0
    For i = LBound(modList) To UBound(modList)
        modCode = Trim(modList(i))
        On Error Resume Next
        ModPrice = Application.WorksheetFunction.VLookup(modCode, ws.Range("A:C"), 3, False)
        On Error GoTo 0
        
        If Not IsError(ModPrice) And ModPrice <> "" Then
            If InStr(1, ModPrice, "%") > 0 Then
                ModTotalPrice = ModTotalPrice + (basePrice * CDbl(Replace(ModPrice, "%", "")) / 100)
            ElseIf IsNumeric(ModPrice) Then
                ModTotalPrice = ModTotalPrice + CDbl(ModPrice)
            End If
        End If
    Next i
    
    ExtractModPrice = basePrice + ModTotalPrice
End Function

Sub SortPrixModifie()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sélection")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ws.Range("A1:E" & lastRow).Sort Key1:=ws.Range("E1"), Order1:=xlDescending, Header:=xlYes
End Sub
