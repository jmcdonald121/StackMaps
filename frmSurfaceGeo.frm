VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSurfaceGeo 
   Caption         =   "Surface Geology Input"
   ClientHeight    =   3660
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4500
   OleObjectBlob   =   "frmSurfaceGeo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSurfaceGeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboLithology_Change()
On Error GoTo ErrorHandler

   If Me.cboLithology.TextLength < 1 Then
      Me.txtThickness.Enabled = False
      Me.cboModifier.Enabled = False
      Me.cmdCommit.Enabled = False
   Else
      Me.txtThickness.Enabled = True
      'Me.txtThickness.SetFocus
   End If

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmSurfaceGeo.cboLithology_Change"
End Sub

Private Sub cboModifier_Change()
On Error GoTo ErrorHandler

   If Me.cboModifier.TextLength < 1 Then
      Me.cmdCommit.Enabled = False
   Else
      Me.cmdCommit.Enabled = True
   End If

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmSurfaceGeo.cboModifier_Change"
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrorHandler
   
   If frmSurfaceGeo.Caption = "Surface Geology Updated" Then
      Dim pFeatureLayer As IFeatureLayer
      Dim pMxDoc As IMxDocument
      Dim pMap As IMap
      Set pMxDoc = ThisDocument
      Set pMap = pMxDoc.FocusMap
      Set pFeatureLayer = CodeUtils.FindLayer("Surface Geology", pMap)
      Dim pFCursor As IFeatureCursor
      Dim pFCC As IFeatureClass
      Set pFCC = pFeatureLayer.FeatureClass
      Dim pQueryFilter As IQueryFilter
      Set pQueryFilter = New QueryFilter
      pQueryFilter.WhereClause = "GEO_ID = '" & frmSurfaceGeo.txtGeoID & "'"
      Set pFCursor = pFCC.Search(pQueryFilter, True)
      Dim pFeature As IFeature
      Set pFeature = pFCursor.NextFeature
      
      If pFeature Is Nothing Then
        Exit Sub
      Else
        pFeature.Value(pFeature.Fields.FindField("Attribute")) = "Y"
        pFeature.Store
      End If
      pMxDoc.ActiveView.Refresh
   End If
   Unload frmSurfaceGeo

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmSurfaceGeo.cmdCancel_Click"
End Sub

Private Sub cmdCommit_Click()
On Error GoTo ErrorHandler

  Dim strGeoID As String
  Dim intLayer As Integer
  Dim strLithology As String
  Dim intThickness As Integer
  Dim strModifier As String
  
  strGeoID = frmSurfaceGeo.txtGeoID.Text
  intLayer = frmSurfaceGeo.txtLayer.Text
  strLithology = frmSurfaceGeo.cboLithology.Text
  strModifier = frmSurfaceGeo.cboModifier.Text
  
  If VarType(frmSurfaceGeo.txtThickness.Text) = vbNull Or frmSurfaceGeo.txtThickness.Text = "" Then
    MsgBox "Thickness is NOT POPULATED!" & vbCrLf & "Please add a thickness for this unit", vbCritical, "NO THICKNESS!!!"
    Exit Sub
  Else
    intThickness = frmSurfaceGeo.txtThickness.Text
  End If
  
  If strGeoID = "" Then
    MsgBox "GeoID is Not Populated."
    Exit Sub
  End If
  Debug.Print "Value of Thickness = " & intLayer
  
  If intLayer = 0 Then
    MsgBox "Layer is Not Populated."
    Exit Sub
  End If
  
  If strLithology = "" Then
    MsgBox "Lithology is Not Populated."
    Exit Sub
  End If
  
  If strModifier = "" Then
    MsgBox "Modifier is Not Populated."
    Exit Sub
  ElseIf strModifier = "none" Then
     strModifier = "n"
  End If
  
  'Update the 1-M lithology table
  CodeUtils.UpdateLithologyDB strGeoID, intLayer, strLithology, intThickness, strModifier

  If intLayer = 1 Then
    'Update the lithology field in the Surface_Geology polygon featureclass
    'Used to symbolize the surface or top-most lithology
    UpdateLithologyField strLithology

  End If
  intLayer = intLayer + 1
  txtLayer.Text = intLayer
  frmSurfaceGeo.Caption = "Surface Geology Updated"
  
Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmSurfaceGeo.cmdCommit_Click"
End Sub

Private Sub txtThickness_Change()
On Error GoTo ErrorHandler

   If Me.txtThickness.TextLength < 1 Then
      Me.cboModifier.Enabled = False
      Me.cmdCommit.Enabled = False
   Else
      Me.cboModifier.Enabled = True
   End If

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmSurfaceGeo.txtThickness_Change"
End Sub

Private Sub UserForm_Initialize()
On Error GoTo ErrorHandler

  frmSurfaceGeo.txtLayer.SetFocus
  frmSurfaceGeo.txtLayer.Text = 1
    
  frmSurfaceGeo.cboModifier.Enabled = False
  frmSurfaceGeo.cboLithology.AddItem ("w")
  frmSurfaceGeo.cboLithology.AddItem ("m")
  frmSurfaceGeo.cboLithology.AddItem ("a")
  frmSurfaceGeo.cboLithology.AddItem ("At")
  frmSurfaceGeo.cboLithology.AddItem ("O")
  frmSurfaceGeo.cboLithology.AddItem ("C")
  frmSurfaceGeo.cboLithology.AddItem ("L")
  frmSurfaceGeo.cboLithology.AddItem ("LC")
  frmSurfaceGeo.cboLithology.AddItem ("S")
  frmSurfaceGeo.cboLithology.AddItem ("SG")
  frmSurfaceGeo.cboLithology.AddItem ("IC")
  frmSurfaceGeo.cboLithology.AddItem ("CG")
  frmSurfaceGeo.cboLithology.AddItem ("T")
  frmSurfaceGeo.cboLithology.AddItem ("P")
  frmSurfaceGeo.cboLithology.AddItem ("Ss")
  frmSurfaceGeo.cboLithology.AddItem ("SSh")
  frmSurfaceGeo.cboLithology.AddItem ("Sh")

  frmSurfaceGeo.cboModifier.AddItem ("none")
  frmSurfaceGeo.cboModifier.AddItem ("()")
  frmSurfaceGeo.cboModifier.AddItem ("-")
  
  frmSurfaceGeo.cboLithology.ListIndex = 0
  frmSurfaceGeo.cboModifier.ListIndex = 0
  
  'frmSurfaceGeo.txtThickness.Text = 10
  
Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmSurfaceGeo.UserForm_Initialize"
End Sub
Public Sub UpdateLithologyField(strLithology As String)
'This procedure updates the lithology field in the Surface_Geology polygon featureclass
'This LITHOLOGY field value is used to symbolize the map based upon the surface or top-most
'lithology.

On Error GoTo ErrorHandler

      Dim pFeatureLayer As IFeatureLayer
      Dim pMxDoc As IMxDocument
      Dim pMap As IMap
      Set pMxDoc = ThisDocument
      Set pMap = pMxDoc.FocusMap
      Set pFeatureLayer = CodeUtils.FindLayer("Surface Geology", pMap)
      Dim pFCursor As IFeatureCursor
      Dim pFCC As IFeatureClass
      Set pFCC = pFeatureLayer.FeatureClass
      Dim pQueryFilter As IQueryFilter
      Set pQueryFilter = New QueryFilter
      pQueryFilter.WhereClause = "GEO_ID = '" & frmSurfaceGeo.txtGeoID & "'"
      Set pFCursor = pFCC.Search(pQueryFilter, True)
      Dim pFeature As IFeature
      Set pFeature = pFCursor.NextFeature
      
      If pFeature Is Nothing Then
        Exit Sub
      Else
        pFeature.Value(pFeature.Fields.FindField("LITHOLOGY")) = strLithology
        pFeature.Store
      End If
      pMxDoc.ActiveView.Refresh

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "frmSurfaceGeo.UpdateLithologyField"
End Sub
