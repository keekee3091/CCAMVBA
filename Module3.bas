Attribute VB_Name = "Module3"
Sub RechercheOptimiseeAvecFiltrageParCode()
    Dim wsCCAM As Worksheet, wsModifiers As Worksheet, wsResultats As Worksheet
    Dim lastRowCCAM As Long
    Dim i As Long, j As Long, keyword As String
    Dim resultats As Object
    Dim ccamData As Variant
    Dim code As String, intitulé As String, prixPrincipal As Double
    Dim matchFound As Boolean
    Dim keywords() As String
    Dim item As Variant
    Dim p As Long
    Dim codeFiltre As String
    Dim regex As Object

    Set wsCCAM = ThisWorkbook.Sheets("CCAM")

    On Error Resume Next
    Set wsResultats = ThisWorkbook.Sheets("Résultats")
    If wsResultats Is Nothing Then
        Set wsResultats = ThisWorkbook.Sheets.Add
        wsResultats.Name = "Résultats"
    End If
    On Error GoTo 0

    wsResultats.Cells.ClearContents
    
    Dim delCbx As OLEObject
    For Each delCbx In wsResultats.OLEObjects
        If TypeName(delCbx.Object) = "CheckBox" Then
            delCbx.Delete
        End If
    Next delCbx
    
    wsResultats.Range("A1:E1").Value = Array("Code", "Intitulé", "Modificateurs", "Prix Principal", "Prix Modifié")

    keyword = Trim(InputBox("Entrez un mot-clé pour la recherche :"))
    If keyword = "" Then Exit Sub

    Set resultats = CreateObject("Scripting.Dictionary")
    lastRowCCAM = wsCCAM.Cells(wsCCAM.Rows.Count, 1).End(xlUp).Row

    keywords = Split(keyword, " ")

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^[A-Z]{4}\d{3}$"
    regex.IgnoreCase = True
    regex.Global = False

    ccamData = wsCCAM.Range("A2:G" & lastRowCCAM).Value

    For i = 1 To UBound(ccamData, 1)

        code = ccamData(i, 1)
        intitulé = ccamData(i, 2)
        prixPrincipal = GetNumericValue(ccamData(i, 3))

        If InStr(1, code, "[", vbTextCompare) > 0 Or InStr(1, code, "]", vbTextCompare) > 0 Then GoTo NextIteration
        If Not regex.Test(code) Then GoTo NextIteration

        matchFound = True
        For j = LBound(keywords) To UBound(keywords)
            If InStr(1, LCase(intitulé), LCase(keywords(j)), vbTextCompare) = 0 Then
                matchFound = False
                Exit For
            End If
        Next j


        If matchFound Then
            codeFiltre = Left(code, 4)
            Exit For
        End If
NextIteration:
    Next i

   Dim modifiers As String, Price As Double, ModPrice As Double

    If codeFiltre <> "" Then
        For i = 1 To UBound(ccamData, 1)
            If Left(ccamData(i, 1), 4) = codeFiltre And Not resultats.Exists(ccamData(i, 1)) And regex.Test(ccamData(i, 1)) Then
                Price = GetNumericValue(ccamData(i, 3))
                modifiers = ExtractModifiers(wsCCAM, i + 1)
                ModPrice = IIf(modifiers <> "", ExtractModPrice(modifiers, Price), Price)
                resultats.Add ccamData(i, 1), Array(ccamData(i, 1), ccamData(i, 2), modifiers, Price, ModPrice)
            End If
        Next i
    End If
    
    For i = 1 To UBound(ccamData, 1)
        matchFound = True
        For j = LBound(keywords) To UBound(keywords)
            If InStr(1, LCase(ccamData(i, 2)), LCase(keywords(j)), vbTextCompare) = 0 Then
                matchFound = False
                Exit For
            End If
        Next j
        
        If matchFound And Not resultats.Exists(ccamData(i, 1)) And regex.Test(ccamData(i, 1)) Then
            Price = GetNumericValue(ccamData(i, 3))
            modifiers = ExtractModifiers(wsCCAM, i + 1)
            ModPrice = IIf(modifiers <> "", ExtractModPrice(modifiers, Price), Price)
            resultats.Add ccamData(i, 1), Array(ccamData(i, 1), ccamData(i, 2), modifiers, Price, ModPrice)
        End If
    Next i


If resultats.Count > 0 Then
    Dim sortedResults() As Variant
    ReDim sortedResults(1 To resultats.Count, 1 To 5)
    Dim index As Integer: index = 1

    For Each item In resultats.Items
        For j = 1 To 5
            sortedResults(index, j) = item(j - 1)
        Next j
        index = index + 1
    Next item

    wsResultats.Range("A2").Resize(UBound(sortedResults, 1), UBound(sortedResults, 2)).Value = sortedResults
    Call SortPrixModifie
    Call HighlightKeywords(wsResultats, keyword)
    Call AddCheckboxes

    Range("E:E").EntireColumn.Hidden = True
    
    Sheets("Résultats").Activate
Else
    MsgBox "Aucun résultat trouvé. Veuillez vérifier vos mots-clés."
End If
End Sub

Sub AddCheckboxes()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim cbx As OLEObject
    Dim btn As Shape
    Dim btn2 As Shape
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set ws = ThisWorkbook.Sheets("Résultats")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For Each cbx In ws.OLEObjects
        If TypeName(cbx.Object) = "CheckBox" Then cbx.Delete
    Next cbx

    For i = 2 To lastRow
        Set cbx = ws.OLEObjects.Add(ClassType:="Forms.CheckBox.1", _
            Left:=ws.Cells(i, 6).Left + 2, Top:=ws.Cells(i, 6).Top + 2, _
            Width:=15, Height:=15)
        cbx.Object.Caption = ""
        cbx.LinkedCell = ws.Cells(i, 6).Address
        cbx.Object.BackStyle = 0
        cbx.Object.Value = False
        cbx.Name = "CheckBox_" & i
    Next i

    For Each btn In ws.Shapes
        If btn.Type = msoFormControl Then btn.Delete
    Next btn
    
    For Each btn2 In ws.Shapes
        If btn2.Type = msoFormControl Then btn2.Delete
    Next btn2

    Set btn = ws.Shapes.AddFormControl(xlButtonControl, ws.Cells(1, 8).Left, ws.Cells(1, 8).Top, 150, 25)
    With btn
        .TextFrame.Characters.Text = "Copier les Sélections"
        .OnAction = "CopySelectedResults"
    End With
    
    Set btn2 = ws.Shapes.AddFormControl(xlButtonControl, ws.Cells(3, 8).Left, ws.Cells(3, 8).Top, 150, 25)
    With btn2
        .TextFrame.Characters.Text = "Rechercher les mots-clés"
        .OnAction = "RechercheOptimiseeAvecFiltrageParCode"
    End With

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


Sub CopySelectedResults()
    Dim wsRes As Worksheet, wsSel As Worksheet
    Dim lastRowRes As Long, lastRowSel As Long
    Dim i As Long, found As Range
    
    Set wsRes = ThisWorkbook.Sheets("Résultats")
    
    On Error Resume Next
    Set wsSel = ThisWorkbook.Sheets("Sélection")
    If wsSel Is Nothing Then
        Set wsSel = ThisWorkbook.Sheets.Add
        wsSel.Name = "Sélection"
        wsSel.Range("A1:E1").Value = Array("Code", "Intitulé", "Modificateurs", "Prix Principal", "Prix Modifié")
    End If
    On Error GoTo 0

    lastRowRes = wsRes.Cells(wsRes.Rows.Count, 1).End(xlUp).Row
    lastRowSel = wsSel.Cells(wsSel.Rows.Count, 1).End(xlUp).Row + 1
    
    For i = 2 To lastRowRes
        If wsRes.Cells(i, 6).Value = True Then
            Set found = wsSel.Range("A2:A" & wsSel.Cells(wsSel.Rows.Count, 1).End(xlUp).Row).Find(wsRes.Cells(i, 1).Value, LookAt:=xlWhole)
            
            If found Is Nothing Then
                wsSel.Range("A" & lastRowSel & ":E" & lastRowSel).Value = wsRes.Range("A" & i & ":E" & i).Value
                lastRowSel = lastRowSel + 1
            End If
        End If
    Next i
    
    For Each cbx In wsRes.OLEObjects
        If TypeName(cbx.Object) = "CheckBox" Then
            cbx.Object.Value = False
            wsRes.Range(cbx.LinkedCell).Value = False
        End If
    Next cbx
    
    Call AddButtonToSelectionSheet(wsSel)

End Sub

Sub AddButtonToSelectionSheet(wsSel As Worksheet)
    Dim btn As Shape, btn2 As Shape
    
    For Each btn In wsSel.Shapes
        If btn.Name = "ExecuteMacro" Then btn.Delete
    Next btn

    Set btn = wsSel.Shapes.AddFormControl(xlButtonControl, wsSel.Cells(1, 8).Left, wsSel.Cells(1, 8).Top, 150, 25)
    With btn
        .TextFrame.Characters.Text = "Filtrer les modificateurs"
        .OnAction = "ApplyModifiers"
        .Name = "ExecuteMacro"
    End With
    
    For Each btn2 In wsSel.Shapes
        If btn2.Name = "DeleteMacro" Then btn2.Delete
    Next btn2

    Set btn2 = wsSel.Shapes.AddFormControl(xlButtonControl, wsSel.Cells(3, 8).Left, wsSel.Cells(3, 8).Top, 150, 25)
    With btn2
        .TextFrame.Characters.Text = "Supprimer la feuille"
        .OnAction = "Supprimer_Feuille_Selection"
        .Name = "DeleteMacro"
    End With
End Sub

Sub Supprimer_Feuille_Selection()

    Dim wsSelection As Worksheet
    Set wsSelection = ThisWorkbook.Sheets("Sélection")
    
    wsSelection.Cells.ClearContents
    
    wsSelection.Range("A1:E1").Value = Array("Code", "Intitulé", "Prix Principal", "Modificateurs", "Prix Modifié")
End Sub

Sub ApplyModifiers()
    UserForm_ModifierSelection.Show
End Sub



Sub HighlightKeywords(ws As Worksheet, ByVal keywordString As String)
    Dim lastRow As Long, i As Long, j As Long, cellText As String
    Dim keywords() As String
    
    keywords = Split(keywordString, " ")

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    With ws.Range("B2:B" & lastRow).Font
        .ColorIndex = xlAutomatic
        .Bold = False
    End With

    For i = 2 To lastRow
        cellText = ws.Cells(i, 2).Value
        
        For j = LBound(keywords) To UBound(keywords)
            If InStr(1, LCase(cellText), LCase(keywords(j)), vbTextCompare) > 0 Then
                HighlightText ws.Cells(i, 2), keywords(j)
            End If
        Next j
    Next i
End Sub

Sub HighlightText(cell As Range, word As String)
    Dim startPos As Integer
    startPos = InStr(1, LCase(cell.Value), LCase(word), vbTextCompare)

    Do While startPos > 0
        With cell.Characters(startPos, Len(word)).Font
            .Color = RGB(255, 0, 0)
            .Bold = True
        End With
        startPos = InStr(startPos + Len(word), LCase(cell.Value), LCase(word), vbTextCompare)
    Loop
End Sub

Sub SortPrixModifie()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Résultats")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ws.Range("A1:E" & lastRow).Sort Key1:=ws.Range("D1"), Order1:=xlDescending, Header:=xlYes
End Sub

Function GetNumericValue(cellValue As Variant) As Double
    If IsNumeric(cellValue) Then
        GetNumericValue = CDbl(cellValue)
    Else
        GetNumericValue = 0
    End If
End Function

Function ExtractModifiers(ws As Worksheet, rowIndex As Long) As String
    Dim modValue As String
    Dim regex As Object
    Dim matches As Object
    Dim i As Integer
    Dim cleanModValue As String
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\[\s*[A-Z0-9](?:\s*,\s*[A-Z0-9])*\s*\]"
    regex.Global = True
    
    modValue = Trim(ws.Cells(rowIndex + 1, 1).Value & " " & ws.Cells(rowIndex + 1, 2).Value)

    cleanModValue = ""
    
    Set matches = regex.Execute(modValue)
    
    If matches.Count > 0 Then
        For i = 0 To matches.Count - 1
            cleanModValue = cleanModValue & matches(i).Value & " "
        Next i
    End If

    ExtractModifiers = Trim(cleanModValue)
End Function



Function ExtractModPrice(modifiers As String, basePrice As Double) As Double
    Dim ws As Worksheet, lastRow As Long, i As Integer
    Dim modList As Variant, modCode As String, ModPrice As Variant
    Dim ModTotalPrice As Double
    
    Set ws = ThisWorkbook.Sheets("Modifiers")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    modList = Split(Replace(Replace(modifiers, "[", ""), "]", ""), ", ")
    
    ModTotalPrice = 0
    For i = LBound(modList) To UBound(modList)
        modCode = Trim(modList(i))
        On Error Resume Next
        ModPrice = Application.WorksheetFunction.VLookup(modCode, ws.Range("A:C"), 3, False)
        On Error GoTo 0
        
            If InStr(1, ModPrice, "%") > 0 Then
                ModTotalPrice = ModTotalPrice + (basePrice * CDbl(Replace(ModPrice, "%", "")) / 100)
            Else
                ModTotalPrice = ModTotalPrice + CDbl(ModPrice)
            End If
    Next i
    
    ExtractModPrice = basePrice + ModTotalPrice
End Function
