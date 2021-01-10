Attribute VB_Name = "Module1"
Option Explicit
Dim i, nr, nc As Integer
' Open the form
Sub RunForm()
With frmOrganizer
      .StartUpPosition = 0
      .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
      .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
      .Show
End With

End Sub

' pre fill frmAddRecord form

Sub PreFillForm()

Dim arr() As Variant

' to populate the name for the label
Range("A1").Select
'nr = Application.WorksheetFunction.CountA(Rows(1))
nc = Range("A1").End(xlToRight).Column

ReDim arr(nc)
    
    For i = 1 To nc
        arr(i) = Cells(1, i)
    Next i

With frmAddRecord
    .txtInput1.Visible = False
    .txtInput2.Visible = False
    .txtInput3.Visible = False
    .txtInput4.Visible = False
    .txtInput5.Visible = False
    .txtInput6.Visible = False
    .txtInput7.Visible = False
    .txtInput8.Visible = False
    .txtInput9.Visible = False
    .txtInput10.Visible = False
    .txtInput11.Visible = False
    .txtInput12.Visible = False
End With

If nc >= 1 Then
    frmAddRecord.Label1 = arr(1)
    frmAddRecord.txtInput1.Visible = True
    frmAddRecord.Width = 250
End If

If nc >= 2 Then
    frmAddRecord.Label2 = arr(2)
    frmAddRecord.txtInput2.Visible = True
    frmAddRecord.Width = 250
End If

If nc >= 3 Then
    frmAddRecord.Label3 = arr(3)
    frmAddRecord.txtInput3.Visible = True
    frmAddRecord.Width = 250
End If
If nc >= 4 Then
    frmAddRecord.Label4 = arr(4)
    frmAddRecord.txtInput4.Visible = True
    frmAddRecord.Width = 250
End If
If nc >= 5 Then
    frmAddRecord.Label5 = arr(5)
    frmAddRecord.txtInput5.Visible = True
    frmAddRecord.Width = 250
End If
If nc >= 6 Then
    frmAddRecord.Label6 = arr(6)
    frmAddRecord.txtInput6.Visible = True
    frmAddRecord.Width = 250
End If
If nc >= 7 Then
    frmAddRecord.Label7 = arr(7)
    frmAddRecord.txtInput7.Visible = True
    frmAddRecord.Width = 500
End If

If nc >= 8 Then
    frmAddRecord.Label8 = arr(8)
    frmAddRecord.txtInput8.Visible = True
    frmAddRecord.Width = 500
End If
If nc >= 9 Then
    frmAddRecord.Label9 = arr(9)
    frmAddRecord.txtInput9.Visible = True
    frmAddRecord.Width = 500
End If
If nc >= 10 Then
    frmAddRecord.Label10 = arr(10)
    frmAddRecord.txtInput10.Visible = True
    frmAddRecord.Width = 500
End If
If nc >= 11 Then
    frmAddRecord.Label11 = arr(11)
    frmAddRecord.txtInput11.Visible = True
    frmAddRecord.Width = 500
End If
If nc >= 12 Then
    frmAddRecord.Label12 = arr(12)
    frmAddRecord.txtInput12.Visible = True
    frmAddRecord.Width = 500
End If


End Sub


Sub ListBoxPopulate()

nc = Range("A1").End(xlToRight).Column
    
    For i = 1 To nc
    frmDeleteCategory.lbxDeleteCategory.AddItem Cells(1, i)
    Next i
    
End Sub

Sub ListBoxPopulateforAdd()
nc = Range("A1").End(xlToRight).Column
    For i = 1 To nc
    frmAddCategory.lbxAddCategory.AddItem Cells(1, i)
    Next i
End Sub
