VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmThicknessQuery 
   Caption         =   "Lithology Thickness Query"
   ClientHeight    =   2100
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5088
   OleObjectBlob   =   "frmThicknessQuery.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmThicknessQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbxUOp_Click()
On Error GoTo ErrorHandler

'  If Len(cbxUOp.Text) > 0 Then
'    cbxUThickness.Enabled = True
'  Else
'    cbxUThickness.Enabled = False
'  End If

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmThicknessQuery.cbxUOp_Click"
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrorHandler

    Unload Me

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmThicknessQuery.cmdCancel_Click"
End Sub

Private Sub cmdFind_Click()
On Error GoTo ErrorHandler

  CodeUtils.LithologyThicknessQuery txtTop.Text, txtBottom.Text, cbxUOp.Text, cbxLOp.Text, cbxUThickness.Text, cbxLThickness.Text
  ThisDocument.CommandBars.Find(arcid.Query_ZoomToSelected, True).Execute

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmThicknessQuery.cmdFind_Click"
End Sub

Private Sub txtBottom_Change()
On Error GoTo ErrorHandler

  If Len(txtBottom.Text) > 0 Then
    cbxLOp.Enabled = True
    cbxLThickness.Enabled = True
    cmdFind.Enabled = True
  Else
    cbxLOp.Enabled = False
    cbxLThickness.Enabled = False
    cmdFind.Enabled = False
  End If

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmThicknessQuery.txtBottom_Change"
End Sub

Private Sub txtTop_Change()
On Error GoTo ErrorHandler

  If Len(txtTop.Text) > 0 Then
    cbxUOp.Enabled = True
    cbxUThickness.Enabled = True
    txtBottom.Enabled = True
  Else
    cbxUOp.Enabled = False
    cbxUThickness.Enabled = False
    txtBottom.Enabled = False
  End If

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmThicknessQuery.txtTop_Change"
End Sub

Private Sub UserForm_Click()
On Error GoTo ErrorHandler

  cmdFind.Enabled = False

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmThicknessQuery.UserForm_Click"
End Sub

Private Sub UserForm_Initialize()
On Error GoTo ErrorHandler

  cbxUOp.AddItem "="
  cbxUOp.AddItem "<"
  cbxUOp.AddItem ">"
  cbxUOp.AddItem ">="
  cbxUOp.AddItem "<="
  cbxUOp.AddItem "<>"
  cbxUOp.ListIndex = 0
  
  cbxLOp.AddItem "="
  cbxLOp.AddItem "<"
  cbxLOp.AddItem ">"
  cbxLOp.AddItem ">="
  cbxLOp.AddItem "<="
  cbxLOp.AddItem "<>"
  cbxLOp.ListIndex = 0
  
  cbxUThickness.AddItem "0"
  cbxUThickness.AddItem "10"
  cbxUThickness.AddItem "20"
  cbxUThickness.AddItem "30"
  cbxUThickness.AddItem "40"
  cbxUThickness.AddItem "50"
  cbxUThickness.AddItem "60"
  cbxUThickness.ListIndex = 0
  
  cbxLThickness.AddItem "0"
  cbxLThickness.AddItem "10"
  cbxLThickness.AddItem "20"
  cbxLThickness.AddItem "30"
  cbxLThickness.AddItem "40"
  cbxLThickness.AddItem "50"
  cbxLThickness.AddItem "60"
  cbxLThickness.ListIndex = 0

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmThicknessQuery.UserForm_Initialize"
End Sub
