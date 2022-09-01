VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Tasks 
   Caption         =   "Tasks"
   ClientHeight    =   585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3870
   OleObjectBlob   =   "Tasks.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Tasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
If Me.txtCat.Value = "" Then
    MsgBox "Kindly add the category", vbInformation
    Exit Sub
End If

Dim cat As Worksheet
Set cat = ThisWorkbook.Sheets("Tasks")
Dim lr As Integer

If Application.WorksheetFunction.CountIf(cat.Range("B:B"), Me.txtCat.Value) > 0 Then
    MsgBox "The category exists", vbInformation
    Exit Sub
End If

lr = Application.WorksheetFunction.CountA(cat.Range("A:A"))



cat.Range("A" & lr + 1) = lr
cat.Range("B" & lr + 1) = Me.txtCat.Value

MsgBox "Category successfully added"

Me.txtCat.Value = ""
End Sub
