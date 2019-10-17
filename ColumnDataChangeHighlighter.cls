VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Sub PartsHighlight()
'This macro changes highlight color everytime number changes in a column

'Write the column number and maximum number of rows
PartNumber = 5

'Variable to keep track of the color change
Dim a As Integer
a = 0

i = 2
Do Until (Len(Cells(i, PartNumber)) = 0)
    If Cells(i, PartNumber) = Cells(i - 1, PartNumber) Then 'And (Len(Cells(i, PartNumber)) > 0) Then
        a = a + 2 'Part number did not change. don't switch color
    'ElseIf (Len(Cells(i + 1, PartNumber)) = 0) Then
        
    Else:
        a = a + 1 'Part number changed. Switch to new color
    End If

    
    If (WorksheetFunction.IsEven(a)) Then
        Cells(i, PartNumber).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.399975585192419
            .PatternTintAndShade = 0
        End With
    ElseIf Not (WorksheetFunction.IsEven(a)) Then
        Cells(i, PartNumber).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark2
            .TintAndShade = -0.499984740745262
            .PatternTintAndShade = 0
        End With
    End If
        
i = i + 1
Loop

End Sub

