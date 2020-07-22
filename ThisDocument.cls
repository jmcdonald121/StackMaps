VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private m_pEditor As IEditor
Private m_pApp As IApplication
Private m_pMxDoc As IMxDocument
Private m_pMap As IMap
Private m_GeoID As String
Private WithEvents m_pEditorEvents As Editor
Attribute m_pEditorEvents.VB_VarHelpID = -1
Private m_pUserPoint As IPoint
Private m_blnAnnotationLayerPresent As Boolean

Private Function AddLithologyAnno_CursorID() As Variant
On Error GoTo ErrorHandler

'As of ArcGIS 9.2 SP2, there is a bug with assigning a CursorID.  See the
'following URL for information: http://forums.esri.com/Thread.asp?c=93&f=993&t=210651#645216
'Jim McDonald, 20070501

'   Set AddLithologyAnno_CursorID = LoadPicture("C:\workspace\ogs\Cursors\M.cur ")
'   Set AddLithologyAnno_CursorID = LoadPicture("C:\workspace\ogs\Cursors\Identify_m.cur")
'   Set AddLithologyAnno_CursorID = LoadPicture("D:\Stack\OGS\Cursors\Add_anno.cur")
   
Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "AddLithologyAnno_CursorID"
End Function

Private Function AddLithologyAnno_Enabled() As Boolean
On Error GoTo ErrorHandler

  Dim blnEditState As Boolean
  blnEditState = CurrentEditState
  If blnEditState And m_blnAnnotationLayerPresent Then
     AddLithologyAnno_Enabled = True
  Else
     AddLithologyAnno_Enabled = False
  End If

Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "AddLithologyAnno_Enabled"
End Function

Private Sub AddLithologyAnno_MouseDown(ByVal button As Long, ByVal shift As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ErrorHandler

    Dim blnEdit As Boolean
    blnEdit = CurrentEditState
    ' Enabled Event should prevent this case but put in just in case.
    If Not blnEdit Then
     MsgBox "Requires an Edit Session.", vbCritical
     Exit Sub
    End If
    If button <> 1 Then Exit Sub
    
    Dim pFeatureLayer As IFeatureLayer
    Dim pPoint As IPoint
    Dim pActiveView As IActiveView
    
    ' Return User Point from Map
    Set pActiveView = m_pMap
    Set pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
    
    ' Find Surface Geology Layer
    Set pFeatureLayer = CodeUtils.FindLayer("Surface Geology", m_pMap)
    If pFeatureLayer Is Nothing Then
      MsgBox "Surface Geology Layer could NOT be located. Please add to Map.", vbCritical
      Exit Sub
    End If
    
    ' Locate Surface Geology Polygon based on User Point
    Dim pGeologyRow As IRow
    Set pGeologyRow = CodeUtils.IdentifyFeature(pFeatureLayer, pPoint)
    If pGeologyRow Is Nothing Then
       MsgBox "No Feature was located.", vbCritical
       Exit Sub
    End If
    
    ' Get Feature Object
    Dim pFeature As IFeature
    Set pFeature = pGeologyRow 'QI
    
    If pFeature Is Nothing Then
       MsgBox "nothing selected"
       Exit Sub
    End If
    
    ' Get some Values from the Polygon
    ' Exit Sub if Surface Geology Polygon has Null Value
    If IsNull(pFeature.Value(pFeature.Fields.FindField("LITHOLOGY"))) Then
      MsgBox "Feature has no attributes to create Annotation."
      Exit Sub
    End If
    
    Dim strLithology As String
    strLithology = pFeature.Value(pFeature.Fields.FindField("LITHOLOGY"))
    m_GeoID = pFeature.Value(pFeature.Fields.FindField("GEO_ID"))
    CodeUtils.FlashFeature pFeature, m_pMxDoc
    Dim pGeometry As IGeometry
    Set pGeometry = pFeature.ShapeCopy
    
    ' Get the Centroid of the Polygon
    If TypeOf pGeometry Is IPolygon Then
          Dim pPolygon As IPolygon
          Set pPolygon = pGeometry 'QI
          Dim pArea As IArea
          Set pArea = pPolygon 'QI
          Dim pCentroid As IPoint
          Set pCentroid = New Point
          pArea.QueryLabelPoint pCentroid
    Else
      Exit Sub
    End If
    
    'Flash Label Point Location
    ' Start Drawing on screen
    m_pMxDoc.ActiveView.ScreenDisplay.StartDrawing 0, esriNoScreenCache
    CodeUtils.FlashPoint m_pMxDoc.ActiveView.ScreenDisplay, pCentroid
    ' Finish drawing on screen
    m_pMxDoc.ActiveView.ScreenDisplay.FinishDrawing '<- Extremely Important
    
    ' Get Related Rows from Lithology Table (see CodeUtils Module)
    Dim pRelatedSet As ISet
    Set pRelatedSet = CodeUtils.RelatedRows(pFeatureLayer, pFeature)
     
    ' Make sure some Related Rows exist
    If pRelatedSet Is Nothing Then
      MsgBox "Can NOT locate any related records for Surface Geology Polygon.", vbCritical
      Exit Sub
    End If
    
    ' Add Values from each Row in the Related Set into a Collection so that the Records
    ' are sorted by Layer (i.e. Layer 1, Layer 2, etc.)
    Dim pRow As IRow
    Dim i As Integer
    Dim intCount As Integer
    Dim intLayer As Integer
    intCount = 1
    Dim pCollection As New Collection
    Dim strModifier As String
    Dim intThickness As Integer
    Dim strThickness As String
    For i = 0 To pRelatedSet.Count - 1
        pRelatedSet.Reset
        Set pRow = pRelatedSet.Next
        Do While Not pRow Is Nothing
          intLayer = pRow.Value(pRow.Fields.FindField("LAYER"))
          If intLayer = intCount Then
             strModifier = pRow.Value(pRow.Fields.FindField("MODIFIER"))
             intThickness = pRow.Value(pRow.Fields.FindField("THICKNESS"))
             If intThickness > 10 Then
                strThickness = Left(intThickness, 1)
             Else
                strThickness = ""
             End If
             Select Case strModifier
             Case "()"
               pCollection.Add "(" & pRow.Value(pRow.Fields.FindField("LITHOLOGY")) & strThickness & ")"
             Case "+", "-"
               pCollection.Add pRow.Value(pRow.Fields.FindField("LITHOLOGY")) & strThickness & strModifier
             Case Else
               pCollection.Add pRow.Value(pRow.Fields.FindField("LITHOLOGY")) & strThickness
             End Select
             Exit Do
          End If
          Set pRow = pRelatedSet.Next
        Loop
        intCount = intCount + 1
    Next i
        
'****************************************************
' Separation between Text Elements
'****************************************************
    Dim intSeparation As Integer
    intSeparation = 700
    
    Dim blnLastElement As Boolean
    blnLastElement = False

    Dim pGC As IGraphicsContainer
    Set pGC = m_pMap.ActiveGraphicsLayer 'QI
    
    Dim pGCSel As IGraphicsContainerSelect
    Set pGCSel = pGC 'QI
    pGCSel.UnselectAllElements
    
    Dim pElementCollection As IElementCollection
    Set pElementCollection = New ElementCollection
    
    Dim pTextPoint As IPoint
    Dim pTextElement As ITextElement
    Dim pElement As IElement
    Dim pTxtSym As ITextSymbol
    Dim pFont As IFontDisp
    
    ' Get the total Y Extent of Text Elements
    Dim dblSpacing As Double
    dblSpacing = (pCollection.Count - 1) * intSeparation
    
    ' Get the amount to move each Element
    Dim dblOffset As Double
    dblOffset = CodeUtils.ElementSpacing(pCollection.Count, intSeparation)
  
    For i = 1 To pCollection.Count
       If i = pCollection.Count Then 'Last element
          blnLastElement = False
       Else
          blnLastElement = True
       End If
       
       ' Create TextElement
        Set pTextElement = New TextElement
        
        ' Create TextSymbol and a Font w/wo underline
        Set pTxtSym = New TextSymbol
        Set pFont = CodeUtils.CreateFont(blnLastElement)
    
        'Set TextSymbol's Font property
        With pTxtSym
           .Font = pFont
           .HorizontalAlignment = esriTHACenter
           .VerticalAlignment = esriTVACenter
        End With
        
        'Set TextElement's Scale, Symbol, and Text Properties
        With pTextElement
           .ScaleText = True
           .Symbol = pTxtSym
           .Text = pCollection.item(i)
        End With
    
        Set pElement = pTextElement 'QI
        Set pTextPoint = New Point
        pTextPoint.X = pCentroid.X
        pTextPoint.Y = (pCentroid.Y + dblSpacing) - dblOffset
        pElement.Geometry = pTextPoint
        pElementCollection.Add pElement
        
        dblOffset = dblOffset + intSeparation
    Next i
        
    Dim pSelElement As IElement
    ' Call CurrentEditState Function to see if Edit Session is open
    If m_pEditor.EditState = esriStateEditing Then
      m_pEditor.StartOperation
      For i = 0 To pElementCollection.Count - 1
        pElementCollection.QueryItem i, pSelElement
        pGC.AddElement pSelElement, 0
        pGCSel.SelectElement pSelElement
      Next i
      m_pEditor.StopOperation "Add " & pFeature.Value(pFeature.Fields.FindField("GEO_ID")) & " Annotation"
    Else
      pGC.AddElement pElement, 0
    End If
    
   ' Make sure the user can see the text that has been placed
   ' If the Text is NOT in the visible extent of the Map
   ' Pan the map to that location.
   Dim pEnvelope As IEnvelope
   
   Dim pRelationalOperator As IRelationalOperator
   Set pRelationalOperator = pActiveView.Extent
   
   If Not pRelationalOperator.Contains(pCentroid) Then
      MsgBox "Text was placed outside the visible extent of the map, Panning to that location.", vbInformation
      Set pEnvelope = pActiveView.Extent
      pEnvelope.CenterAt pCentroid
      pActiveView.Extent = pEnvelope
   End If
   pActiveView.Refresh
   
Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "AddLithologyAnno_Select"
End Sub

Private Sub AddLithologyAnno_Select()
On Error GoTo ErrorHandler

    Set m_pMxDoc = ThisDocument 'QI
    Set m_pMap = m_pMxDoc.FocusMap

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "AddLithologyAnno_Select"
End Sub

Private Function AddLithologyAnno_ToolTip() As String
On Error GoTo ErrorHandler

  AddLithologyAnno_ToolTip = "Add Lithology Annotation"

Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "AddLithologyAnno_ToolTip"
End Function

Private Sub GeologyQuery_Click()
On Error GoTo ErrorHandler

  frmGeologyQuery.Show vbModal

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "GeologyQuery_Click"
End Sub

Private Function GeologySeedTool_Enabled() As Boolean
On Error GoTo ErrorHandler

  Dim pFLayer As IFeatureLayer
  Dim pMap As IMap
  Dim pMxDoc As IMxDocument
  Dim blnEditing As Boolean
  
  Set pMxDoc = ThisDocument
  Set pMap = pMxDoc.FocusMap
  Set pFLayer = CodeUtils.FindLayer("Surface Geology", pMap)
  blnEditing = CurrentEditState
  
  If pFLayer Is Nothing Or blnEditing Then
     GeologySeedTool_Enabled = False
  Else
     GeologySeedTool_Enabled = True
  End If

Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "GeologySeedTool_Enabled"
End Function

Private Sub GeologySeedTool_MouseDown(ByVal button As Long, ByVal shift As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ErrorHandler

    Dim pFeatureLayer As IFeatureLayer
    Dim pPoint As IPoint
    Dim pApp As IMxApplication
    Dim pActiveView As IActiveView
    
    Set pApp = Application 'QI
    Set m_pMxDoc = ThisDocument
    Set m_pMap = m_pMxDoc.FocusMap
    Set pActiveView = m_pMap
    Set pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
    
    Set pFeatureLayer = CodeUtils.FindLayer("Surface Geology", m_pMap)
        
    Dim pGeologyRow As IRow
    Set pGeologyRow = CodeUtils.IdentifyFeature(pFeatureLayer, pPoint)
    If pGeologyRow Is Nothing Then
       MsgBox "No Feature was located.", vbCritical
       Exit Sub
    End If
    
    Dim pFeature As IFeature
    Set pFeature = pGeologyRow 'QI
    'Set pFeature = pFeatureCursor.NextFeature
    
    If pFeature Is Nothing Then
       MsgBox "nothing selected"
       Exit Sub
    End If
    
    CodeUtils.FlashFeature pFeature, m_pMxDoc
    
    Dim pGeometry As IGeometry
    Set pGeometry = pFeature.Value(pFeature.Fields.FindField("SHAPE"))
    Dim pPolygon As IPolygon
    Set pPolygon = pGeometry  'QI
    
    CodeUtils.PlantSeed m_pMxDoc, m_pMap, pPolygon

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "GeologySeedTool_Enabled"
End Sub

Private Function GeologySeedTool_ToolTip() As String
On Error GoTo ErrorHandler

  GeologySeedTool_ToolTip = "Geology Seed Tool"
  
Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "GeologySeedTool_ToolTip"
End Function

Private Sub m_pEditorEvents_OnCreateFeature(ByVal obj As IObject)
On Error GoTo ErrorHandler

  Dim pFieldIndex As Long
  pFieldIndex = obj.Fields.FindField("GEO_ID")
  If pFieldIndex <> -1 Then
    obj.Value(pFieldIndex) = m_GeoID
    Application.StatusBar.Message(0) = m_GeoID & " has been added to Annotation Feature(s)"
  End If
  
Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "m_pEditorEvents_OnCreateFeature"
End Sub

Private Function MxDocument_OnContextMenu(ByVal X As Long, ByVal Y As Long) As Boolean
On Error GoTo ErrorHandler

  Dim pCommandItem As ICommandItem
  Set pCommandItem = Application.CurrentTool
  If pCommandItem.Name = "AddLithologyAnno" Then
    Set m_pMap = m_pMxDoc.FocusMap
    Dim pActiveView As IActiveView
    Set pActiveView = m_pMap
    
    Set m_pUserPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
    Dim pMenu As ICommandBar
    Set pMenu = CommandBars.Create("New", esriCmdBarTypeShortcutMenu)
    
    pMenu.CreateMacroItem "Search for Annotation", 8, "Project.ThisDocument.QueryLithologyAnno"
    pMenu.Add arcid.PanZoom_ZoomInFixed
    pMenu.Add arcid.PanZoom_ZoomOutFixed
    pMenu.Popup
    MxDocument_OnContextMenu = True
  Else
    MxDocument_OnContextMenu = False
  End If

Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "MxDocument_OnContextMenu"
End Function

Private Sub StartEditingAnnotation_Click()
On Error GoTo ErrorHandler

  Dim pID As UID
  Set pID = New UID
  pID = "esriEditor.Editor"
  Set m_pEditor = Application.FindExtensionByCLSID(pID)
  'Call StartEditing routine
  StartEditing m_pEditor

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "StartEditingAnnotation_Click"
End Sub

Public Sub StartEditing(pEditor As IEditor)
On Error GoTo ErrorHandler

  Dim pFDOGraphicsLayer As IFDOGraphicsLayer
  Dim pFeatureLayer As IFeatureLayer
  Dim pDataset As IDataset
  Dim iLayerCount As Integer
  Dim pEditLayers As IEditLayers
  
  Set m_pApp = Application
  Set m_pMxDoc = m_pApp.Document 'QI
  Set m_pMap = m_pMxDoc.FocusMap
  
  'If an edit session has already been started exit
  If Not pEditor.EditState = esriStateNotEditing Then
    ' Initialize Editor
    Set m_pEditorEvents = pEditor
    Exit Sub
  Else
    Set m_pEditorEvents = pEditor
  End If
  
  'Start editing the workspace of the first FDOGraphicsLayer you find
  For iLayerCount = 0 To m_pMap.LayerCount - 1
    If TypeOf m_pMap.Layer(iLayerCount) Is IFDOGraphicsLayer Then
       Set pFDOGraphicsLayer = m_pMap.Layer(iLayerCount)
       Set pFeatureLayer = pFDOGraphicsLayer 'QI
       Set pDataset = pFeatureLayer.FeatureClass
       pEditor.StartEditing pDataset.Workspace
       Exit For
    End If
  Next iLayerCount
  
  ' Only begin Editing if Lithology Annotation Layer is present.
  If Not pFDOGraphicsLayer Is Nothing Then
      m_blnAnnotationLayerPresent = True
      
      'Establish Center Snap Agent
      Dim pSnapAgent As ISnapAgent
      Dim pSnapEnv As ISnapEnvironment
      Dim pMySnapAgent As clsCenterSnap
      Dim pUID As New UID
      Dim Count As Integer
    
      pUID = "esriEditor.Editor"
    
      Set pSnapEnv = pEditor
      Set pEditLayers = pEditor
      Set pFeatureLayer = CodeUtils.FindLayer("Surface Geology", m_pMap)
      
      If Not pFeatureLayer Is Nothing Then
        pEditLayers.SetCurrentLayer pFeatureLayer, 0
      End If
      
      'Find custom snap agent and set is featureclass property
      For Count = 0 To pSnapEnv.SnapAgentCount - 1
        Set pSnapAgent = pSnapEnv.SnapAgent(Count)
        If pSnapAgent.Name = "Center Snap" Then
          Set pMySnapAgent = pSnapAgent
          pMySnapAgent.FeatureClass = pFeatureLayer.FeatureClass
        End If
      Next Count
      
      Set pFeatureLayer = CodeUtils.FindLayer("Annotation Leader Lines", m_pMap)
      
      If Not pFeatureLayer Is Nothing Then
        pEditLayers.SetCurrentLayer pFeatureLayer, 0
      End If

      Set m_pMap.ActiveGraphicsLayer = pFDOGraphicsLayer
  Else
    m_blnAnnotationLayerPresent = False
    MsgBox "Annotation Layer is NOT present.", vbCritical, "Warning - Missing Data"
  End If

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "StartEditing"
End Sub
Public Sub SetEditLayer(pMap As IMap, pEditor As IEditor)
On Error GoTo ErrorHandler

  Dim pFeatureLayer As IFeatureLayer
  Set pFeatureLayer = CodeUtils.FindLayer("Lithology Annotation", pMap)
  Dim pEditLayers As IEditLayers
  Set pEditLayers = pEditor 'QI
  
  If pEditLayers.IsEditable(pFeatureLayer) Then
    pEditLayers.SetCurrentLayer pFeatureLayer, 0
  End If
  
Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "SetEditLayer"
End Sub
Public Function CurrentEditState() As Boolean
On Error GoTo ErrorHandler

  Dim pID As UID
  Set pID = New UID
  pID = "esriEditor.Editor"
  Set m_pEditor = Application.FindExtensionByCLSID(pID)

  If m_pEditor.EditState = esriStateEditing Then
    CurrentEditState = True
  Else
    CurrentEditState = False
  End If

Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CurrentEditState"
End Function

Private Function StartEditingAnnotation_Enabled() As Boolean
On Error GoTo ErrorHandler

  Dim blnEditState As Boolean
  blnEditState = CurrentEditState
  If blnEditState Then
     StartEditingAnnotation_Enabled = False
  Else
     StartEditingAnnotation_Enabled = True
  End If

Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "StartEditingAnnotation_Enabled"
End Function

Private Function SurfaceGeologyTool_CursorID() As Variant
On Error GoTo ErrorHandler

'As of ArcGIS 9.2 SP2, there is a bug with assigning a CursorID.  See the
'following URL for information: http://forums.esri.com/Thread.asp?c=93&f=993&t=210651#645216
'Jim McDonald, 20070501

'  Set SurfaceGeologyTool_CursorID = LoadPicture("D:\Stack\OGS\Cursors\Surface_Geology.cur")

Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "SurfaceGeologyTool_CursorID"
End Function

'Private Function SurfaceGeologyTool_Enabled() As Boolean
'On Error GoTo ErrorHandler
'
'  Dim pFLayer As IFeatureLayer
'  Dim pMap As IMap
'  Dim pMxDoc As IMxDocument
'  Dim blnEditing As Boolean
'
'  Set pMxDoc = ThisDocument
'  Set pMap = pMxDoc.FocusMap
'  Set pFLayer = CodeUtils.FindLayer("Surface Geology", pMap)
'  blnEditing = CurrentEditState
'
'  If pFLayer Is Nothing Or blnEditing Then
'     SurfaceGeologyTool_Enabled = False
'  Else
'     SurfaceGeologyTool_Enabled = True
'  End If
'
'Exit Function 'Avoid ErrorHandler
'ErrorHandler:
'   MsgBox Err.Description, vbInformation, "SurfaceGeologyTool_Enabled"
'End Function

Private Sub SurfaceGeologyTool_MouseDown(ByVal button As Long, ByVal shift As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ErrorHandler

    Dim pFeatureLayer As IFeatureLayer
    Dim pPoint As IPoint
    Dim pApp As IMxApplication
    Dim pActiveView As IActiveView
    
    Set pApp = Application 'QI
    Set m_pMxDoc = ThisDocument
    Set m_pMap = m_pMxDoc.FocusMap
    Set pActiveView = m_pMap
    Set pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
    
    Set pFeatureLayer = CodeUtils.FindLayer("Surface Geology", m_pMap)
        
    Dim pGeologyRow As IRow
    Set pGeologyRow = CodeUtils.IdentifyFeature(pFeatureLayer, pPoint)
    If pGeologyRow Is Nothing Then
       MsgBox "No Feature was located.", vbCritical
       Exit Sub
    End If
    
    Dim pFeature As IFeature
    Set pFeature = pGeologyRow 'QI
    
    If pFeature Is Nothing Then
       MsgBox "nothing selected"
       Exit Sub
    End If
    
    CodeUtils.FlashFeature pFeature, m_pMxDoc
    
    Dim pGeometry As IGeometry
    Set pGeometry = pFeature.Value(pFeature.Fields.FindField("SHAPE"))
    Dim pPolygon As IPolygon
    Set pPolygon = pGeometry  'QI
    
    ' Find Geology Graphic object and change shape
    CodeUtils.PlantSeed m_pMxDoc, m_pMap, pPolygon
    
    Dim strGeoID As String
    strGeoID = pFeature.Value(pFeature.Fields.FindField("geo_id"))
    Application.StatusBar.Message(0) = strGeoID
    frmSurfaceGeo.txtGeoID.Text = strGeoID
    
    CodeUtils.CheckLithology m_pMap, strGeoID
'    Exit Sub

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "SurfaceGeologyTool_MouseDown"
End Sub

Private Function SurfaceGeologyTool_ToolTip() As String
On Error GoTo ErrorHandler

   SurfaceGeologyTool_ToolTip = "Input Surface Geology"

Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "SurfaceGeologyTool_ToolTip"
End Function
Public Sub UpdateSurfaceGeology()
On Error GoTo ErrorHandler

   Set m_pMxDoc = ThisDocument
   Set m_pMap = m_pMxDoc.FocusMap

       ' CodeUtils.CheckLithology strGeoID
       Dim pTableCollection As ITableCollection
       Set pTableCollection = m_pMap
       
       Dim pTable As ITable
       Set pTable = pTableCollection.Table(0)
        
       Dim pDataset As IDataset
       Set pDataset = pTable
        
       Dim pWorkspace As IWorkspace
       Set pWorkspace = pDataset.Workspace
       
       'Test to see if table is Lithology before continuing
       If pDataset.Name <> "lithology" Then
          MsgBox "Table could NOT be Located!" & vbCrLf & "Tool is unable to finish processing your request.", vbCritical, "Warning"
          Exit Sub
       Else
          Application.StatusBar.Message(0) = "Table was located."
       End If
    
       Dim pCursor As ICursor
       Dim pQueryFilter As IQueryFilter
       Set pQueryFilter = New QueryFilter
       pQueryFilter.WhereClause = "LAYER = " & "1"
       Set pCursor = pTable.Search(pQueryFilter, True)
    
       Dim pRow As IRow
       Set pRow = pCursor.NextRow
       
       Dim strGeoID As String
       Dim strLayer As String
       Dim strLithology As String
       Dim strThickness As String
       Dim strModifier As String
       Dim strDate As String
       Dim strUser As String
       
       Dim strString As String
       strString = ""
       
       Dim pFeatureLayer As IFeatureLayer
       Set pFeatureLayer = CodeUtils.FindLayer("Surface Geology", m_pMap)
       Dim pFCursor As IFeatureCursor
       Dim pFeature As IFeature
       Dim pUpdateQueryFilter As IQueryFilter
       Dim lngFieldIndex As Long
       
       If Not pRow Is Nothing Then
         Do While Not pRow Is Nothing
            strGeoID = pRow.Value(pRow.Fields.FindField("GEO_ID"))
            strLithology = pRow.Value(pRow.Fields.FindField("LITHOLOGY"))
            Set pUpdateQueryFilter = New QueryFilter
            pUpdateQueryFilter.WhereClause = "GEO_ID = '" & strGeoID & "'"
            Set pFCursor = pFeatureLayer.Search(pUpdateQueryFilter, False)
            Set pFeature = pFCursor.NextFeature
            If Not pFeature Is Nothing Then
                lngFieldIndex = pFeature.Fields.FindField("Lithology")
                pFeature.Value(lngFieldIndex) = strLithology
                pFeature.Store
            End If
            Set pRow = pCursor.NextRow
         Loop
        Exit Sub
       End If
    Exit Sub

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "UpdateSurfaceGeology"
End Sub

Private Sub ThicknessQuery_Click()
On Error GoTo ErrorHandler

 frmThicknessQuery.Show

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "ThicknessQuery_Click"
End Sub

Private Sub TransferFromSeed_Click()
On Error GoTo ErrorHandler

  ' Find Seed Graphic
  Set m_pMxDoc = Application.Document 'QI
  Set m_pMap = m_pMxDoc.FocusMap
  Dim pGeometry As IGeometry
  Set pGeometry = CodeUtils.FindGeologyGraphic
    
  ' Exit Sub if no Seed is located
  If pGeometry Is Nothing Then
    MsgBox "Please use the Geology Seed Tool to specify a Seed Polygon before using this Command.", vbCritical
    Exit Sub
  End If
  
  ' Get the Centroid of the Seed
  If TypeOf pGeometry Is IPolygon Then
        Dim pPolygon As IPolygon
        Set pPolygon = pGeometry 'QI
        Dim pArea As IArea
        Set pArea = pPolygon 'QI
        Dim pCentroid As IPoint
        Set pCentroid = New Point
        pArea.QueryLabelPoint pCentroid
  Else
    Exit Sub
  End If
  
  ' Locate the Seed Feature
  Dim pFeatureLayer As IFeatureLayer
  Set pFeatureLayer = CodeUtils.FindLayer("Surface Geology", m_pMap)
  Dim pRow As IRow
  Set pRow = CodeUtils.IdentifyFeature(pFeatureLayer, pCentroid)
  Dim pSeedFeature As IFeature
  Set pSeedFeature = pRow 'QI
  
  'Get Top Lithology and GEO_ID from Seed Feature
  Dim strPolygonLith As String
  strPolygonLith = pSeedFeature.Value(pSeedFeature.Fields.FindField("LITHOLOGY"))
  Dim strSeedGeoID As String
  strSeedGeoID = pSeedFeature.Value(pSeedFeature.Fields.FindField("GEO_ID"))
  
  CodeUtils.FlashFeature pSeedFeature, m_pMxDoc
  
  ' Delete existing Related Records for each selected feature
  Dim pFSel As IFeatureSelection
  Dim pSelection As ISelectionSet
  Dim pFCursor As IFeatureCursor
  Dim pSelFeature As IFeature
  Dim strSelGeoID As String
  Set pFSel = pFeatureLayer 'QI
  Set pSelection = pFSel.SelectionSet
  pSelection.Search Nothing, False, pFCursor
  Set pSelFeature = pFCursor.NextFeature
  Do While Not pSelFeature Is Nothing
    CodeUtils.FlashFeature pSelFeature, m_pMxDoc
    strSelGeoID = pSelFeature.Value(pSelFeature.Fields.FindField("GEO_ID"))
    'Transfer Lithology Value to selected features
    pSelFeature.Value(pSelFeature.Fields.FindField("LITHOLOGY")) = strPolygonLith
    pSelFeature.Value(pSelFeature.Fields.FindField("Attribute")) = "Y"
    pSelFeature.Store
    Dim pLithCursor As ICursor
    Set pLithCursor = CodeUtils.QueryLithologyDB(strSelGeoID)
    Dim pSelRow As IRow
    Set pSelRow = pLithCursor.NextRow
    If Not pSelRow Is Nothing Then
        Do While Not pSelRow Is Nothing
           pSelRow.Delete
           Set pSelRow = pLithCursor.NextRow
        Loop
    End If
    Set pSelFeature = pFCursor.NextFeature
  Loop
  
  ' Get a Cursor for Seed from Lithology table
  Dim pSeedCursor As ICursor
  Set pSeedCursor = CodeUtils.QueryLithologyDB(strSeedGeoID)
  
  ' Process each Lithology Row for Seed GeoID
  Dim pSeedRow As IRow
  Set pSeedRow = pSeedCursor.NextRow
  If pSeedRow Is Nothing Then
    MsgBox "Seed Feature has no related Records in Lithology table!", vbCritical
    Exit Sub
  End If

  Dim strLayer As String
  Dim strLithology As String
  Dim strThickness As String
  Dim strModifier As String
  Dim strDate As String
  Dim strUser As String
  
  Dim pTable As ITable
  Set pTable = CodeUtils.GetLithologyTable
  
  ' Get Lithology table so Rows can be added
  If pTable Is Nothing Then
    MsgBox "Lithology table could NOT be located!", vbCritical
    Exit Sub
  End If
  
  
  Dim pNewRow As IRow
    
  'Re-establish Selection Cursor
  pSelection.Search Nothing, False, pFCursor
  Set pSelFeature = pFCursor.NextFeature

  Do While Not pSelFeature Is Nothing
    ' Re-establish Cursor for Seed Rows in Lithology
    Set pSeedCursor = CodeUtils.QueryLithologyDB(strSeedGeoID)
    Set pSeedRow = pSeedCursor.NextRow
      Do While Not pSeedRow Is Nothing
         strLayer = pSeedRow.Value(pSeedRow.Fields.FindField("LAYER"))
         strLithology = pSeedRow.Value(pSeedRow.Fields.FindField("LITHOLOGY"))
         strThickness = pSeedRow.Value(pSeedRow.Fields.FindField("THICKNESS"))
         strModifier = pSeedRow.Value(pSeedRow.Fields.FindField("MODIFIER"))
         strDate = Date & " " & Time
         strUser = Environ("Username")
        
         Set pNewRow = pTable.CreateRow
         strSelGeoID = pSelFeature.Value(pSelFeature.Fields.FindField("GEO_ID"))
         pNewRow.Value(pNewRow.Fields.FindField("GEO_ID")) = strSelGeoID
         pNewRow.Value(pNewRow.Fields.FindField("LAYER")) = strLayer
         pNewRow.Value(pNewRow.Fields.FindField("LITHOLOGY")) = strLithology
         pNewRow.Value(pNewRow.Fields.FindField("THICKNESS")) = strThickness
         pNewRow.Value(pNewRow.Fields.FindField("MODIFIER")) = strModifier
         pNewRow.Value(pNewRow.Fields.FindField("CREATION_DATE")) = strDate
         pNewRow.Value(pNewRow.Fields.FindField("USER_NAME")) = strUser
         pNewRow.Store
         ' For each selected Feature Add Row
         Set pSeedRow = pSeedCursor.NextRow
       Loop
    Set pSelFeature = pFCursor.NextFeature
  Loop
  
  Dim pActiveView As IActiveView
  Set pActiveView = m_pMap
  pActiveView.Refresh

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "TransferFromSeed_Click"
End Sub

Private Function TransferFromSeed_Enabled() As Boolean
On Error GoTo ErrorHandler
  
  Dim blnEditing As Boolean
  blnEditing = CurrentEditState
  
  Set m_pMxDoc = Application.Document 'QI
  Set m_pMap = m_pMxDoc.FocusMap
  
  Dim pFeatureLayer As IFeatureLayer
  Set pFeatureLayer = CodeUtils.FindLayer("Surface Geology", m_pMap)
  
  If pFeatureLayer Is Nothing Then
    TransferFromSeed_Enabled = False
    Exit Function
  End If
  
  Dim pFSel As IFeatureSelection
  Set pFSel = pFeatureLayer 'QI
  
  If Not blnEditing And pFSel.SelectionSet.Count > 0 Then
     TransferFromSeed_Enabled = True
  Else
     TransferFromSeed_Enabled = False
  End If

Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "TransferFromSeed_Enabled"
End Function

Private Function TransferFromSeed_ToolTip() As String
On Error GoTo ErrorHandler

  TransferFromSeed_ToolTip = "Duplicate Lithology"

Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "TransferFromSeed_ToolTip"
End Function
Public Sub CreateSplineText()
On Error GoTo ErrorHandler

  ' Find Seed Graphic
  Set m_pMxDoc = Application.Document 'QI
  Set m_pMap = m_pMxDoc.FocusMap
  Dim pGeometry As IGeometry
  Set pGeometry = CodeUtils.FindGeologyGraphic
    
  ' Exit Sub if no Seed is located
  If pGeometry Is Nothing Then
    MsgBox "Please use the Geology Seed Tool to specify a Seed Polygon before using this Command.", vbCritical
    Exit Sub
  End If
  
  ' Get the Centroid of the Seed
  If TypeOf pGeometry Is IPolygon Then
        Dim pPolygon As IPolygon
        Set pPolygon = pGeometry 'QI
        Dim pArea As IArea
        Set pArea = pPolygon 'QI
        Dim pCentroid As IPoint
        Set pCentroid = New Point
        pArea.QueryLabelPoint pCentroid
  Else
    Exit Sub
  End If
  
  ' Locate the Seed Feature
  Dim pFeatureLayer As IFeatureLayer
  Set pFeatureLayer = CodeUtils.FindLayer("Surface Geology", m_pMap)
  Dim pRow As IRow
  Set pRow = CodeUtils.IdentifyFeature(pFeatureLayer, pCentroid)
  Dim pSeedFeature As IFeature
  Set pSeedFeature = pRow 'QI
  
  'Get Top Lithology and GEO_ID from Seed Feature
  Dim strPolygonLith As String
  strPolygonLith = pSeedFeature.Value(pSeedFeature.Fields.FindField("LITHOLOGY"))
  Dim strSeedGeoID As String
  strSeedGeoID = pSeedFeature.Value(pSeedFeature.Fields.FindField("GEO_ID"))
  
  CodeUtils.FlashFeature pSeedFeature, m_pMxDoc
  Exit Sub
  
  Dim pGraphicSelection As IGraphicsContainerSelect
  Set pGraphicSelection = m_pMxDoc.ActiveView
  
  If TypeOf pGraphicSelection.SelectedElement(0).Geometry Is ICurve Then
        
    Dim pElement As IElement
    Set pElement = New TextElement
    
    Dim pColor As IRgbColor
    Set pColor = New RgbColor
    pColor.RGB = 255
    
    Dim pFont As IFontDisp
    Set pFont = New stdole.StdFont
    pFont.Name = "AriaL"
    pFont.Size = 10
    Dim pTextSymbol As ISimpleTextSymbol
    Set pTextSymbol = New TextSymbol
    
    Dim pTextPath As ITextPath
    Set pTextPath = New SimpleTextPath
    With pTextPath
      Set .Geometry = pElement.Geometry
      .XOffset = 10
      .YOffset = 20
    End With
    With pTextSymbol
      .Font = pFont
      .Color = pColor
      .HorizontalAlignment = esriTHALeft
      .VerticalAlignment = esriTVABaseline
      .Clip = False
      .Size = 10
      .Text = "set on the TextSymbol"
      .XOffset = 0
      .YOffset = 0
    End With
    
    Set pTextSymbol.TextPath = pTextPath
    
    Dim pTextElement As ITextElement
    Set pTextElement = New TextElement
    With pTextElement
      .ScaleText = False
      .Symbol = pTextSymbol
      .Text = "Set on the Element"
    End With
    
    Set pElement = pTextElement
    
    Dim pGraphicsContainer As IGraphicsContainer
    Set pGraphicsContainer = m_pMxDoc.ActiveView
    pGraphicsContainer.AddElement pElement, 0
    m_pMxDoc.ActiveView.Refresh
  End If

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CreateSplineText"
End Sub

Public Sub QueryLithologyAnno()
On Error GoTo ErrorHandler

    Dim pFeatureLayer As IFeatureLayer
    Dim pPoint As IPoint
    Dim pActiveView As IActiveView
    Set pActiveView = m_pMxDoc.FocusMap
    
    ' Find Surface Geology Layer
    Set pFeatureLayer = CodeUtils.FindLayer("Surface Geology", m_pMap)
    If pFeatureLayer Is Nothing Then
      MsgBox "Surface Geology Layer could NOT be located. Please add to Map.", vbCritical
      Exit Sub
    End If
    
    Dim pFDOLayer As IFDOGraphicsLayer
    Dim i As Integer
    Dim pLayer As ILayer
    For i = 0 To m_pMxDoc.FocusMap.LayerCount - 1
        Set pLayer = m_pMxDoc.FocusMap.Layer(i)
        If TypeOf pLayer Is IFDOGraphicsLayer Then
           Set pFDOLayer = m_pMxDoc.FocusMap.Layer(i)
           Exit For
        End If
    Next i
        
    If Not pFDOLayer Is Nothing Then
        ' Locate Surface Geology Polygon based on User Point
        Dim pGeologyRow As IRow
        Set pGeologyRow = CodeUtils.IdentifyFeature(pFeatureLayer, m_pUserPoint)
        If pGeologyRow Is Nothing Then
           MsgBox "No Feature was located.", vbCritical
           Exit Sub
        End If
        
        ' Get Feature Object
        Dim pFeature As IFeature
        Set pFeature = pGeologyRow 'QI
        
        If pFeature Is Nothing Then
           MsgBox "nothing selected"
           Exit Sub
        End If
        
        ' Get some Values from the Polygon
        '
        ' Exit Sub if Surface Geology Polygon has Null Value
        If IsNull(pFeature.Value(pFeature.Fields.FindField("LITHOLOGY"))) Then
          MsgBox "Feature has no attributes to create Annotation."
          Exit Sub
        End If
        
        Dim strLithology As String
        strLithology = pFeature.Value(pFeature.Fields.FindField("LITHOLOGY"))
        m_GeoID = pFeature.Value(pFeature.Fields.FindField("GEO_ID"))
        CodeUtils.FlashFeature pFeature, m_pMxDoc
    
        Dim pFDOFeatureLayer As IFeatureLayer
        Set pFDOFeatureLayer = pFDOLayer 'QI
        
        Dim pFeatureSelection As IFeatureSelection
        Set pFeatureSelection = pFDOFeatureLayer 'QI
        pFeatureSelection.Clear
        
        Dim pQueryFilter As IQueryFilter
        Set pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = "GEO_ID = '" & m_GeoID & "'"
        pFeatureSelection.SelectFeatures pQueryFilter, esriSelectionResultNew, False
        pFeatureSelection.SelectionChanged
          
        If pFeatureSelection.SelectionSet.Count = 0 Then
            MsgBox "No Annotation exists for " & m_GeoID & ".", vbInformation
        Else
            pActiveView.PartialRefresh 4, Nothing, Nothing
        End If
    End If

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "QueryLithologyAnno"
End Sub
