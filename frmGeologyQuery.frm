VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGeologyQuery 
   Caption         =   "Lithology Query"
   ClientHeight    =   2100
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   2820
   OleObjectBlob   =   "frmGeologyQuery.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGeologyQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
On Error GoTo ErrorHandler

    Unload Me

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmGeologyQuery.cmdCancel_Click"
End Sub

Private Sub cmdFind_Click()
On Error GoTo ErrorHandler

    CodeUtils.GetSurfaceGeo txtTop.Text, txtBottom.Text
    ThisDocument.CommandBars.Find(arcid.Query_ZoomToSelected, True).Execute
    
Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmGeologyQuery.cmdFind_Click"
End Sub

Private Sub txtBottom_Change()
On Error GoTo ErrorHandler

  If Len(txtBottom.Text) > 0 Then
    cmdFind.Enabled = True
  Else
    cmdFind.Enabled = False
  End If

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmGeologyQuery.txtBottom_Change"
End Sub

Private Sub txtTop_Change()
On Error GoTo ErrorHandler

  If Len(txtTop.Text) > 0 Then
    txtBottom.Enabled = True
  Else
    txtBottom.Enabled = False
  End If

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmGeologyQuery.txtTop_Change"
End Sub

Private Sub UserForm_Click()
On Error GoTo ErrorHandler

  cmdFind.Enabled = False

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmGeologyQuery.UserForm_Click"
End Sub
