Attribute VB_Name = "FunLibraryMod"
Const outCnt = 13
    
'ں��ޒ�`�̍쐬
'W�Ȍ�S�폜
'N�AO�AP�AQ��폜
'K��폜
Function delColumns() As Integer()

    Columns("W:AM").Select
    Selection.Delete Shift:=xlToLeft
    Columns("N:Q").Select
    Selection.Delete Shift:=xlToLeft
    Columns("K:K").Select
    Selection.Delete Shift:=xlToLeft
    
End Function

'�o�R���\��
Function getOutColList() As Integer()
    Dim outColList(outCnt) As Integer

    'XX�\��
     For i = 1 To outCnt
         Select Case i
             Case 1
                 outColList(i) = 1
             Case 2
                 outColList(i) = 2
             Case 3
                 outColList(i) = 3
             Case 4
                 outColList(i) = 4
             Case 13
                 outColList(i) = 13
             Case Else
                outColList(i) = outColList(i - 1) + 1
         End Select
     Next i
    
    getOutColList = outColList
    
End Function

'�����z�\��
Function getOutColList1() As Integer()
    Dim outColList(outCnt) As Integer

    '�����z�\��
    For i = 1 To outCnt
        Select Case i
            Case 9
                outColList(i) = 43
            Case Else
               outColList(i) = outColList(i - 1) + 1
        End Select
    Next i
    
    getOutColList1 = outColList
    
End Function
    
'����\��
Function getOutColList2() As Integer()
    Dim outColList(outCnt) As Integer

    '����\��
    For i = 1 To outCnt
        Select Case i
            Case 13
                outColList(i) = 62
            Case Else
               outColList(i) = outColList(i - 1) + 1
        End Select
    Next i
    
    getOutColList2 = outColList
    
End Function

'�o�R���\��
Function getOutColList3() As Integer()
    Dim outColList(outCnt) As Integer

    '����\��
    For i = 1 To outCnt
        Select Case i
            Case 7
                outColList(i) = 71
            Case 10
                outColList(i) = 85
            Case Else
               outColList(i) = outColList(i - 1) + 1
        End Select
    Next i
    
    getOutColList3 = outColList
    
End Function

'�݌ɊǗ�
Function getOutColList4() As Integer()
    Dim outColList(outCnt) As Integer

    '�݌�OPEN���ʕ\��
     For i = 1 To outCnt
         Select Case i
             Case 2
                 outColList(i) = wColNum("C")
             Case 3
                 outColList(i) = wColNum("F")
             Case 5
                 outColList(i) = wColNum("Q")
             Case 6
                 outColList(i) = wColNum("S")
             Case 9
                 outColList(i) = wColNum("Z")
             Case Else
                outColList(i) = outColList(i - 1) + 1
         End Select
     Next i
    
    getOutColList4 = outColList
    
End Function

Public Function wColNm(ColNum)
    'MsgBox wColNm(27) return AA
    wColNm = Split(Cells(1, ColNum).Address, "$")(1)
End Function

Public Function wColNum(ColNm)
    'MsgBox wColNum("D") return 4
    wColNum = Range(ColNm & 1).Column
End Function

Sub CrtDBAddress()

    Dim iptSRow, iptSCol As Long
    Dim iptSAddr, iptSSheet As String
    
    Dim outSRow, outSCol As Long
    Dim outERow, outECol As Long
    Dim outSAddr As String
    
    'ReDim outColList(outCnt) As Integer
    Dim outColList As Variant
    
    iptSRow = Selection.Row
    iptSCol = Selection.Column
    iptSAddr = Selection.Address
    iptSSheet = Selection.Worksheet.Name

    Select Case Cells(iptSRow, iptSCol)
       Case "�݌ɊǗ��e�[�u��"
          outColList = getOutColList4() '�݌ɊǗ�
       Case "���Ɋm�FHEAD TBL"
          'outColList = getOutColList3() '�o�R���\��
          'outColList = getOutColList1() '�����z�\��
          outColList = getOutColList2() '����\��
       Case Else
          outColList = getOutColList()
    End Select
    
    Dim Target As Range, SelectCell As Range
    On Error Resume Next
    Set SelectCell = Application.InputBox _
                                  ("�o�̓Z����I�����Ă�������", Type:=8)
    If Err.Number <> 0 Then Exit Sub    ''[�L�����Z��]�{�^�����N���b�N���ꂽ
    
    Application.ScreenUpdating = False
    
    For Each Target In SelectCell
        outSRow = Target.Row
        outSCol = Target.Column
        outSAddr = Target.Address
        outSSheet = Target.Parent.Name
        Exit For
    Next Target
    outERow = outSRow + 3
    outECol = outSCol + outCnt - 1

    Worksheets(outSSheet).Activate
    Range(Cells(outSRow, outSCol), Cells(outERow, outECol)).Select
    Selection.NumberFormatLocal = "G/�W��"
    
    'ں��ޖ�
    Worksheets(outSSheet).Cells(outSRow, outSCol).Value = "=" & iptSSheet & "!" & ColumnLetter(iptSCol) & iptSRow & ""
    Dim i As Integer
    For i = 0 To outCnt
        Worksheets(outSSheet).Cells(outSRow + 1, outSCol + i).Value = "=" & iptSSheet & "!" & ColumnLetter(iptSCol + outColList(i + 1) - 1) & iptSRow + 1 & ""
        Worksheets(outSSheet).Cells(outSRow + 2, outSCol + i).Value = "=" & iptSSheet & "!" & ColumnLetter(iptSCol + outColList(i + 1) - 1) & iptSRow + 2 & ""
        Worksheets(outSSheet).Cells(outSRow + 3, outSCol + i).Value = "=" & iptSSheet & "!" & ColumnLetter(iptSCol + outColList(i + 1) - 1) & iptSRow + 3 & ""
    Next i
    Range(Cells(outSRow, outSCol), Cells(outERow, outECol)).Select
    Selection.NumberFormatLocal = "@"
    Call setCellRecHeader(outSSheet, outSRow, outSCol, outERow, outECol)
    Worksheets(outSSheet).Cells(outSRow, outSCol).Select

    Application.ScreenUpdating = True

End Sub


Private Sub setCellRecHeader(outSSheet, sLine, sColumn, eLine, eColumn)

    'heade 3�s�Að��ٖ��A���فA���{�ꖼ
    Worksheets(outSSheet).Cells(sLine, sColumn).Select
    Selection.Font.Bold = True
    
    ActiveCell.Offset(1, 0).Select
    sLine = sLine + 1
    Call fmtCellLineStyle(sLine, sColumn, eLine, eColumn)
    
    '�ŏ��J�[�\����I��
    Range(Cells(sLine, sColumn), Cells(sLine, eColumn)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
    '�������ւP�ړ�
    sLine = sLine + 1
    Range(Cells(sLine, sColumn), Cells(sLine, eColumn)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    
    sLine = sLine - 1
    
End Sub

Public Sub fmtCellLineStyle(sLine, sColumn, eLine, eColumn)

    Range(Cells(sLine, sColumn), Cells(eLine, eColumn)).Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Sub


Public Function ColumnLetter(Column As Long) As String
    If Column < 1 Then Exit Function
    ColumnLetter = ColumnLetter(Int((Column - 1) / 26)) & Chr(((Column - 1) Mod 26) + Asc("A"))
End Function


Sub CrtDBAddress2()

Dim l As Long
Dim s As String
l = ActiveCell.Row
l = ActiveCell.Column
s = ActiveCell.Address

l = Selection.Row
l = Selection.Column
s = Selection.Address

End Sub


Sub chgPicture()

    Selection.ShapeRange.ScaleHeight 0.85, msoFalse, msoScaleFromTopLeft
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
    End With
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With


    'Selection.ShapeRange.LockAspectRatio = msoFalse
    'Selection.ShapeRange.ScaleWidth 0.9, msoFalse, msoScaleFromTopLeft
    'Selection.ShapeRange.PictureFormat.Crop.PictureWidth = 1378
    'Selection.ShapeRange.PictureFormat.Crop.PictureHeight = 458
    'Selection.ShapeRange.PictureFormat.Crop.PictureOffsetX = 0
    'Selection.ShapeRange.PictureFormat.Crop.PictureOffsetY = 0
    
End Sub

'�k�����Ȃ��A�g��������̂�
Sub chgPicture2()
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
    End With
End Sub


Sub chgPicture3(sh As Variant, lflg As Variant)

    Selection.ShapeRange.ScaleHeight sh, msoFalse, msoScaleFromTopLeft
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
    End With
    If lflg = "Y" Then
        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
        End With
    End If
    If lflg = "N" Then
        Selection.ShapeRange.Line.Visible = msoFalse
    End If
    
End Sub

Sub chgFont()

    'Selection.Font.Bold = True '����
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Font
        '.Color = -16776961 '�Ԏ�
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        '.Color = 65535  '���F
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Sub

Sub FormatDBTable()

    Range(Selection, Selection.End(xlToRight)).Select
    If ActiveCell.Offset(1, 0).Value <> "" Then
        Range(Selection, Selection.End(xlDown)).Select
    End If
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.NumberFormatLocal = "@"
    
    '�ŏ��J�[�\����I��
    ActiveCell.Offset(0, 0).Select
    If ActiveCell.Offset(0, 0).Value = "" Then
      Exit Sub
    End If
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
    '�������ւP�ړ�
    ActiveCell.Offset(1, 0).Select
    If ActiveCell.Offset(0, 0).Value = "" Then
      Exit Sub
    End If
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    
    '�������ւP�ړ�
    ActiveCell.Offset(1, 0).Select
    If ActiveCell.Offset(0, 0).Value = "" Then
      Exit Sub
    End If
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    
    Selection.End(xlUp).Select
    
End Sub



Sub FormatDBTable2()

    Range(Selection, Selection.End(xlToRight)).Select
    If ActiveCell.Offset(1, 0).Value <> "" Then
        Range(Selection, Selection.End(xlDown)).Select
    End If
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.NumberFormatLocal = "@"
    
    '�ŏ��J�[�\����I��
    ActiveCell.Offset(0, 0).Select
    If ActiveCell.Offset(0, 0).Value = "" Then
      Exit Sub
    End If
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
    '�������ւP�ړ�
    ActiveCell.Offset(1, 0).Select
    If ActiveCell.Offset(0, 0).Value = "" Then
      Exit Sub
    End If
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    
    Selection.End(xlUp).Select
    
End Sub


Sub FormatDBTable_CSE()

    If ActiveCell.Offset(0, 0).Value = "" Then
      Exit Sub
    End If
    If ActiveCell.Offset(0, 1).Value = "" And ActiveCell.Offset(1, 0).Value <> "" Then
      Selection.Font.Bold = True
      ActiveCell.Offset(1, 0).Select
    End If
    Range(Selection, Selection.End(xlToRight)).Select
    If ActiveCell.Offset(1, 0).Value <> "" Then
        Range(Selection, Selection.End(xlDown)).Select
    End If
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.NumberFormatLocal = "@"
    
    '�ŏ��J�[�\����I��
    ActiveCell.Offset(0, 0).Select
    If ActiveCell.Offset(0, 1).Value = "" Then
      Exit Sub
    End If
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    ActiveCell.Offset(1, 0).Select
    If ActiveCell.Offset(0, 1).Value = "" Then
      ActiveCell.Offset(-1, 0).Select
      Exit Sub
    End If
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    
    ActiveCell.Offset(-1, 0).Select
    
End Sub


