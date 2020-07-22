VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmList 
   Caption         =   "List Stack Lithologies"
   ClientHeight    =   5130
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6120
   OleObjectBlob   =   "frmList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
On Error GoTo ErrorHandler

   Unload Me

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmList.cmdClose_Click"
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo ErrorHandler

   Unload Me
   frmSurfaceGeo.Show

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmList.cmdUpdate_Click"
End Sub
