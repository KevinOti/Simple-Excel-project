VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Goals 
   Caption         =   "Goals"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4965
   OleObjectBlob   =   "Goals.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Goals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
'Call Calendar.select_d(Me.TextBox3)
Me.txtDate.Value = Calendar.select_d
End Sub

Private Sub CommandButton2_Click()
If Me.cmbCategory.Value = "" Then
    MsgBox "Please select the category", vbCritical
    Exit Sub
End If
If Me.cmbTasks.Value = "" Then
    MsgBox "Please select the task", vbCritical
    Exit Sub
End If
If Me.cmbStatus.Value = "" Then
    MsgBox "Please select the status", vbCritical
    Exit Sub
End If
If Me.cmbCategory.Value = "" Then
    MsgBox "Please select the category", vbCritical
    Exit Sub
End If
If Me.cmbTasks.Value = "" Then
    MsgBox "Please select the task", vbCritical
End If

If Me.cmbStatus.Value = "" Then
    MsgBox "please select the status", vbCritical
    Exit Sub
End If
Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Activities")
Dim lr As Integer
lr = Application.WorksheetFunction.CountA(sh.Range("A:A"))
sh.Range("A" & lr + 1).Value = lr
sh.Range("B" & lr + 1).Value = Me.cmbCategory.Value
sh.Range("C" & lr + 1).Value = Me.cmbTasks.Value
sh.Range("D" & lr + 1).Value = Me.cmbStatus.Value
sh.Range("E" & lr + 1).Value = Me.txtDate.Value

MsgBox "Successfully submitted", vbInformation

Me.cmbCategory.Value = ""
Me.cmbTasks.Value = ""
Me.cmbStatus.Value = ""

Call show_act
End Sub

Private Sub CommandButton4_Click()
Cat_Add.Show
End Sub

Private Sub CommandButton5_Click()
Tasks.Show
End Sub

Private Sub Label5_Click()

End Sub

Private Sub UserForm_Activate()
Me.txtDate.Value = Format(Date, "D-MMM-YYYY")
With Me.cmbStatus
    .AddItem "Yes"
    .AddItem "No"
End With

Call reveal_details
Call show_act
End Sub

Sub reveal_details()

Dim ts As Worksheet
Dim cs As Worksheet
Set ts = ThisWorkbook.Sheets("Tasks")
Set cs = ThisWorkbook.Sheets("Categories")
Dim i As Integer
For i = 2 To Application.WorksheetFunction.CountA(cs.Range("B:B"))
Me.cmbCategory.AddItem cs.Range("B" & i)
Next i

For i = 2 To Application.WorksheetFunction.CountA(ts.Range("B:B"))
Me.cmbTasks.AddItem ts.Range("B" & i)
Next i
End Sub

Sub show_act()
Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Activities")
Dim lr As Integer
lr = Application.WorksheetFunction.CountA(sh.Range("A:A"))
If lr = 1 Then lr = 2

With Me.ListBox1
    .ColumnHeads = True
    .ColumnWidths = "30, 65, 50, 50, 10, 50"
    .ColumnCount = 5
    .RowSource = "Activities!A2:H" & lr
End With
End Sub
