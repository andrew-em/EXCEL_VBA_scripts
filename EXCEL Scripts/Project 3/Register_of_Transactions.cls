VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "????????"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub ???????_2()

Dim sheet As Worksheet

For Each sheet In Worksheets

    sheet.Visible = True
    sheet.Cells(sheet.UsedRange.Rows.Count + 1, 8) = "??????? ??????"
    sheet.Cells(sheet.UsedRange.Rows.Count, 9).Formula = "=AVERAGE(I2:I" & sheet.UsedRange.Rows.Count - 1 & ")"
    sheet.Cells(sheet.UsedRange.Rows.Count, 9).HorizontalAlignment = xlCenter
            
    
Next sheet


End Sub
