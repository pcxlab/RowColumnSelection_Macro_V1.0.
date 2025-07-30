VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)

'Cells.Interior.ColorIndex = xlColorIndexNone
'Target.EntireColumn.Interior.ColorIndex = 36
'Target.EntireRow.Interior.ColorIndex = 44
'Target.Interior.Gradient = Default

'Target.Interior.ColorIndex = XlColorIndex = none

'End Sub


' Code begin
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)

    ' Remove previous formatting
    Cells.Interior.ColorIndex = xlColorIndexNone
    
    ' Highlight the entire column with a new color
    ' ColorIndex 36 is light yellow, here are some alternatives:
    ' 6 = Yellow, 15 = Grey, 34 = Light Turquoise, 38 = Light Green
    'Target.EntireColumn.Interior.ColorIndex = 50 ' Light Turquoise
    Target.EntireColumn.Interior.ColorIndex = 36 ' Light Turquoise
    ' Highlight the entire row with a different color
    ' ColorIndex 44 is light orange, here are some alternatives:
    ' 3 = Red, 50 = Light Blue, 35 = Light Green, 24 = Blue
    Target.EntireRow.Interior.ColorIndex = 34 ' Light Blue
    
    ' Optional: You can also set individual cell gradient (commented out for now)
    ' Target.Interior.Color = RGB(255, 255, 204) ' Light yellow RGB example
    ' Target.Interior.Gradient = Default

End Sub


