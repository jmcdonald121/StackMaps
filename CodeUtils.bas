Attribute VB_Name = "CodeUtils"
Option Explicit
' The following variable requires you Add a Reference to Microsoft Runtime Scripting
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private m_pMxDoc As IMxDocument
Private m_pMap As IMap
Private m_pActiveView As IActiveView
Private m_pQFilter As IQueryFilter
Private m_pFClass As IFeatureClass
Private m_pFCursor As IFeatureCursor

Public Function FindLayer(pLayerName As String, pMap As IMap) As IFeatureLayer
'Finds a layer based on a name and then returns that layer as an IFeatureLayer
On Error GoTo ErrorHandler

  Dim iLoop As Integer
  Dim gLoop As Integer
  Dim pFLayer As IFeatureLayer
  Dim pCompositeLayer As ICompositeLayer
  Dim pLayerEnum As IEnumLayer
  Dim pLayer As ILayer
  For iLoop = 0 To pMap.LayerCount - 1
    If TypeOf pMap.Layer(iLoop) Is IFeatureLayer Then
      Set pFLayer = pMap.Layer(iLoop)
      If pFLayer.Name = pLayerName Then
        Set FindLayer = pFLayer
        Exit Function
      End If
    ElseIf TypeOf pMap.Layer(iLoop) Is IGroupLayer Then
        Set pCompositeLayer = pMap.Layer(iLoop)
        For gLoop = 0 To pCompositeLayer.Count - 1
         'Set pLayerEnum =
         Set pLayer = pCompositeLayer.Layer(gLoop)
          If TypeOf pLayer Is ILayer Then
            Set pFLayer = pLayer
            If pFLayer.Name = pLayerName Then
              Set FindLayer = pFLayer
              Exit Function
            End If
          End If
        Next gLoop
     End If
            
  Next iLoop
  
  Set FindLayer = Nothing
  Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CodeUtils.FindLayer"
End Function
Public Sub FlashFeature(pFeature As IFeature, pMxDoc As IMxDocument)
'Flash a Geometry temporarily on the Map
On Error GoTo ErrorHandler
  
  ' Start Drawing on screen
  pMxDoc.ActiveView.ScreenDisplay.StartDrawing 0, esriNoScreenCache
  
  ' Switch functions based on Geomtry type
  Select Case pFeature.Shape.GeometryType
    Case esriGeometryPolyline
      FlashLine pMxDoc.ActiveView.ScreenDisplay, pFeature.ShapeCopy
    Case esriGeometryPolygon
      FlashPolygon pMxDoc.ActiveView.ScreenDisplay, pFeature.ShapeCopy
    Case esriGeometryPoint
      FlashPoint pMxDoc.ActiveView.ScreenDisplay, pFeature.ShapeCopy
  End Select
  
  ' Finish drawing on screen
  pMxDoc.ActiveView.ScreenDisplay.FinishDrawing
  Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CodeUtils.FlashFeature"
End Sub
Public Sub FlashLine(pDisplay As IScreenDisplay, pGeometry As IGeometry)
'Flash a Line Geometry temporarily on the Map
On Error GoTo ErrorHandler
  Dim pLineSymbol As ISimpleLineSymbol
  Dim pSymbol As ISymbol
  Dim pRGBColor As IRgbColor
  
  Set pLineSymbol = New SimpleLineSymbol
  pLineSymbol.Width = 4
  
  Set pRGBColor = New RgbColor
  pRGBColor.Green = 128
  
  Set pSymbol = pLineSymbol
  pSymbol.ROP2 = esriROPNotXOrPen
  
  pDisplay.SetSymbol pLineSymbol
  pDisplay.DrawPolyline pGeometry
  Sleep 300
  pDisplay.DrawPolyline pGeometry
  
  Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CodeUtils.FlashLine"
End Sub
Public Sub FlashMultiPoint(pDisplay As IScreenDisplay, pGeometry As IGeometry)
'Flash a Mulitpoint Geometry temporarily on the Map
On Error GoTo ErrorHandler
  Dim pMarkerSymbol As ISimpleMarkerSymbol
  Dim pSymbol As ISymbol
  Dim pRGBColor As IRgbColor
  
  Set pMarkerSymbol = New SimpleMarkerSymbol
  pMarkerSymbol.Style = esriSMSCircle
  pMarkerSymbol.Size = 24
  Set pRGBColor = New RgbColor
  pRGBColor.Green = 128
  
  Set pSymbol = pMarkerSymbol
  pSymbol.ROP2 = esriROPNotXOrPen
  
  pDisplay.SetSymbol pMarkerSymbol
  pDisplay.DrawMultipoint pGeometry
  Sleep 1200
  pDisplay.DrawMultipoint pGeometry
  
  Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CodeUtils.FlashMultiPoint"
End Sub
Public Sub FlashPoint(pDisplay As IScreenDisplay, pGeometry As IGeometry)
'Flash a Mulitpoint Geometry temporarily on the Map
On Error GoTo ErrorHandler
  Dim pMarkerSymbol As ISimpleMarkerSymbol
  Dim pSymbol As ISymbol
  Dim pRGBColor As IRgbColor
  
  Set pMarkerSymbol = New SimpleMarkerSymbol
  pMarkerSymbol.Style = esriSMSCircle
  
  Set pRGBColor = New RgbColor
  pRGBColor.Green = 128
  
  Set pSymbol = pMarkerSymbol
  pSymbol.ROP2 = esriROPNotXOrPen
  
  pDisplay.SetSymbol pMarkerSymbol
  pDisplay.DrawPoint pGeometry
  Sleep 300
  pDisplay.DrawPoint pGeometry
  
  Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CodeUtils.FlashPoint"
End Sub
Public Sub FlashPolygon(pDisplay As IScreenDisplay, pGeometry As IGeometry)
'Flash a Polygon Geometry temporarily on the Map
On Error GoTo ErrorHandler
  Dim pFillSymbol As ISimpleFillSymbol
  Dim pSymbol As ISymbol
  Dim pRGBColor As IRgbColor
  
  Set pFillSymbol = New SimpleFillSymbol
  pFillSymbol.Outline = Nothing
  
  Set pRGBColor = New RgbColor
  pRGBColor.Green = 128
  
  Set pSymbol = pFillSymbol
  pSymbol.ROP2 = esriROPNotXOrPen
  
  pDisplay.SetSymbol pFillSymbol
  pDisplay.DrawPolygon pGeometry
  Sleep 300
  pDisplay.DrawPolygon pGeometry
  
  Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CodeUtils.FlashPolygon"
End Sub
Public Function IdentifyFeature(pFLayer As IFeatureLayer, pPoint As IPoint) As IRow
'Finds and returns an IRow based upon an input FeatureLayer and Point
On Error GoTo ErrorHandler
    Dim pIdentify As IIdentify
    Dim pArray As IArray
    Dim pGeometry As IGeometry
    Set pGeometry = pPoint 'QI
    Set pIdentify = pFLayer  'QI
    Set pArray = pIdentify.Identify(pGeometry)
    If pArray Is Nothing Then
      Set IdentifyFeature = Nothing
      Exit Function
    End If
     
    Dim i As Integer
    Dim pRowIDObj As IRowIdentifyObject
    Dim pRow As IRow
     
    For i = 0 To pArray.Count - 1
      If TypeOf pArray.Element(i) Is IRowIdentifyObject Then
        Set pRowIDObj = pArray.Element(i)
        Set IdentifyFeature = pRowIDObj.Row
      End If
    Next

    Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CodeUtils.IdentifyFeature"
End Function

Public Sub CheckLithology(pMap As IMap, strGeoID As String)
On Error GoTo ErrorHandler
   Dim pTableCollection As ITableCollection
   Set pTableCollection = pMap
   
   Dim pTable As ITable
   Set pTable = pTableCollection.Table(0)
    
   Dim pDataset As IDataset
   Set pDataset = pTable
    
   Dim pWorkspace As IWorkspace
   Set pWorkspace = pDataset.Workspace
   
   'Test to see if table is FAXLOG before continuing
   If pDataset.Name <> "lithology" Then
      MsgBox "Table could NOT be Located!" & vbCrLf & "Tool is unable to finish processing your request.", vbCritical, "Warning"
      Exit Sub
   Else
      Application.StatusBar.Message(0) = "Table was located."
   End If

   Dim pCursor As ICursor
   Dim pQueryFilter As IQueryFilter
   Set pQueryFilter = New QueryFilter
   pQueryFilter.WhereClause = "GEO_ID = '" & strGeoID & "'"
   Set pCursor = pTable.Search(pQueryFilter, True)

   Dim pRow As IRow
   Set pRow = pCursor.NextRow
   Dim strLayer As String
   Dim strLithology As String
   Dim strThickness As String
   Dim strModifier As String
   Dim strDate As String
   Dim strUser As String
   
   Dim strString As String
   strString = ""
   
   If Not pRow Is Nothing Then
      Do While Not pRow Is Nothing
         strGeoID = pRow.Value(pRow.Fields.FindField("GEO_ID"))
         strLayer = pRow.Value(pRow.Fields.FindField("LAYER"))
         strLithology = pRow.Value(pRow.Fields.FindField("LITHOLOGY"))
         strThickness = pRow.Value(pRow.Fields.FindField("THICKNESS"))
         strModifier = pRow.Value(pRow.Fields.FindField("MODIFIER"))
         strDate = pRow.Value(pRow.Fields.FindField("CREATION_DATE"))
         strUser = pRow.Value(pRow.Fields.FindField("USER_NAME"))
         strString = strString & "GeoID: " & strGeoID & vbCrLf & "Layer:" & strLayer & vbCrLf & "Lithology: " & strLithology & vbCrLf & "Thickness:" & strThickness & vbCrLf & "Modifier:" & strModifier & vbCrLf & "User:" & strUser & vbCrLf & "Date:" & strDate & vbCrLf & vbCrLf & vbCrLf
         frmList.txtRecords.Text = strString
         Set pRow = pCursor.NextRow
      Loop
      frmList.Show
   Else
     frmSurfaceGeo.Show
     Exit Sub
   End If

  Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CodeUtils.CheckLithology"
End Sub
Public Sub DrawPolygon(pMxDoc As IMxDocument, pPolyGeo As IPolygon, strName As String)
On Error GoTo ErrorHandler
   Dim pOutline As ISimpleLineSymbol
   Dim pColor As IRgbColor
   Dim pOLColor As IRgbColor
   Dim pSFillSymbol As ISimpleFillSymbol
   Dim pFillShapeElement As IFillShapeElement
   Dim pGCon As IGraphicsContainer
   Dim pElem As IElement
   Dim pElementProperties As IElementProperties
   Dim pPElem As IPolygonElement
   
   Set pColor = New RgbColor
   Set pOLColor = New RgbColor
   Set pOutline = New SimpleLineSymbol
   Set pSFillSymbol = New SimpleFillSymbol
   
   If strName = "Geology" Then
        With pColor
           .Red = 0
           .Green = 0
           .Blue = 255
        End With
        With pOLColor
           .Red = 0
           .Green = 255
           .Blue = 0
        End With
        With pOutline
           .Color = pOLColor
           .Style = esriSLSSolid
           .Width = 4
        End With
        With pSFillSymbol
           .Color = pColor
           .Outline = pOutline
           .Style = esriSFSForwardDiagonal
        End With
    Else
        With pColor
           .Red = 0
           .Green = 255
           .Blue = 0
        End With
        With pOLColor
           .Red = 0
           .Green = 255
           .Blue = 255
        End With
        With pOutline
           .Color = pOLColor
           .Style = esriSLSSolid
           .Width = 3
        End With
        With pSFillSymbol
           .Color = pColor
           .Outline = pOutline
           .Style = esriSFSHollow
        End With
    End If

    Set pMxDoc = ThisDocument
    Set pGCon = pMxDoc.ActiveView.GraphicsContainer
    
    Set pElem = New PolygonElement
    Set pFillShapeElement = pElem 'QI
    pElem.Geometry = pPolyGeo
    Set pPElem = pElem 'QI
    pFillShapeElement.Symbol = pSFillSymbol
    
    Set pElementProperties = pPElem
    pElementProperties.Name = strName
    pGCon.AddElement pPElem, 0
  Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CodeUtils.DrawPolygon"
End Sub
Public Function FindGeologyGraphic() As IGeometry
On Error GoTo ErrorHandler
    'Get ActiveView
    Set m_pMxDoc = ThisDocument
    Set m_pMap = m_pMxDoc.FocusMap

    ' Find Parcel Graphic object and change shape
    Dim pDGC As IGraphicsContainer
    Dim pEnumElem As IEnumElement
    Dim pElement As IElement
    Dim pElementProperties As IElementProperties
    Dim pDeleteElement As IElement
    
    'QI
    Set pDGC = m_pMap 'QI
    pDGC.Reset
    
    Set pElement = pDGC.Next
    If pElement Is Nothing Then
        ' Parcel Graphic does NOT exist
        Set FindGeologyGraphic = Nothing
        Exit Function
    End If
    Do While (Not pElement Is Nothing)
        If (TypeOf pElement Is IFillShapeElement) Then
           'QI
           Set pElementProperties = pElement
           If pElementProperties.Name = "Geology" Then
              ' Graphic has been located, set variable and exit function
              Set FindGeologyGraphic = pElement.Geometry
              Exit Function
           End If
        End If
        Set pElement = pDGC.Next
    Loop
    
    ' Graphics exist but not the Parcel Graphic
    Set FindGeologyGraphic = Nothing
  
Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CodeUtils.FindGeologyGraphic"
End Function
Public Sub PlantSeed(pMxDoc As IMxDocument, pMap As IMap, pPolygon As IGeometry)
On Error GoTo ErrorHandler
   ' Find Geology Graphic object and change shape
    Dim pDGC As IGraphicsContainer
    Dim pEnumElem As IEnumElement
    Dim pElement As IElement
    Dim pElementProperties As IElementProperties
    Dim pDeleteElement As IElement
    Dim pGeometry As IGeometry
    
    'QI
    Set pDGC = pMap
    pDGC.Reset
    
    Set pElement = pDGC.Next
    Dim blnBuffer As Boolean
    blnBuffer = False
    
    Set pGeometry = pPolygon 'QI
    'Update Geology Graphic on the Data View's GraphicsContainer
    If pElement Is Nothing Then
        Call CodeUtils.DrawPolygon(pMxDoc, pPolygon, "Geology")
    Else
        Do While (Not pElement Is Nothing)
            If (TypeOf pElement Is IFillShapeElement) Then
               'QI
               Set pElementProperties = pElement
               'Application.StatusBar.Message(0) = "Graphic Element " & pElementProperties.name
               
               'Use Select Case to update text by specific Name of Element
               Select Case pElementProperties.Name
                Case Is = "Geology"
                   pElement.Geometry = pGeometry
                Case Is = "Buffer"
                   blnBuffer = True
                   pDGC.DeleteElement pElement
                Case Else
                   Application.StatusBar.Message(0) = "Not a named Graphic element."
               End Select
            End If
            Set pElement = pDGC.Next
            Set pElementProperties = pElement
        Loop
     End If
   pMxDoc.ActiveView.PartialRefresh 8, Nothing, Nothing
    
Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CodeUtils.PlantSeed"
End Sub
Public Sub UpdateLithologyDB(strGeoID As String, intLayer As Integer, strLithology As String, intThickness As Integer, strModifier As String)
On Error GoTo ErrorHandler

   Set m_pMxDoc = ThisDocument
   
   Set m_pMap = m_pMxDoc.FocusMap
   
   Dim pTableCollection As ITableCollection
   Set pTableCollection = m_pMap
   
   Dim pTable As ITable
   Set pTable = pTableCollection.Table(0)
    
   Dim pDataset As IDataset
   Set pDataset = pTable
    
   Dim pWorkspace As IWorkspace
   Set pWorkspace = pDataset.Workspace
   
   'Test to see if table is FAXLOG before continuing
   If pDataset.Name <> "lithology" Then
      MsgBox "Table could NOT be Located!" & vbCrLf & "Tool is unable to finish processing your request.", vbCritical, "Warning"
      Exit Sub
   Else
      Application.StatusBar.Message(0) = "Table was located."
   End If
  
   'Create a new Row and update values
   Dim pRow As IRow
   Dim pCursor As ICursor
   Dim pQueryFilter As IQueryFilter
   Set pQueryFilter = New QueryFilter
   pQueryFilter.WhereClause = "GEO_ID = '" & strGeoID & "' AND LAYER = " & intLayer
   Set pCursor = pTable.Search(pQueryFilter, False)
   Set pRow = pCursor.NextRow
   If pRow Is Nothing Then
     Set pRow = pTable.CreateRow
   End If
    
    pRow.Value(pRow.Fields.FindField("GEO_ID")) = strGeoID
    pRow.Value(pRow.Fields.FindField("LAYER")) = intLayer
    pRow.Value(pRow.Fields.FindField("LITHOLOGY")) = strLithology
    pRow.Value(pRow.Fields.FindField("THICKNESS")) = intThickness
    pRow.Value(pRow.Fields.FindField("MODIFIER")) = strModifier
    pRow.Value(pRow.Fields.FindField("USER_NAME")) = Environ("UserName")
    pRow.Value(pRow.Fields.FindField("CREATION_DATE")) = Date & " " & Time
    pRow.Store
   
Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CodeUtils.UpdateLithologyDB"
End Sub

Public Function QueryLithologyDB(strGeoID As String) As ICursor
On Error GoTo ErrorHandler

   Set m_pMxDoc = ThisDocument
   
   Set m_pMap = m_pMxDoc.FocusMap
   
   Dim pTableCollection As ITableCollection
   Set pTableCollection = m_pMap
   
   Dim pTable As ITable
   Set pTable = pTableCollection.Table(0)
    
   Dim pDataset As IDataset
   Set pDataset = pTable
    
   Dim pWorkspace As IWorkspace
   Set pWorkspace = pDataset.Workspace
   
   'Test to see if table is FAXLOG before continuing
   If pDataset.Name <> "lithology" Then
      MsgBox "Table could NOT be Located!" & vbCrLf & "Tool is unable to finish processing your request.", vbCritical, "Warning"
      Exit Function
   Else
      Application.StatusBar.Message(0) = "Table was located."
   End If
  
   'Create a new Row and update values
   Dim pQueryFilter As IQueryFilter
   Set pQueryFilter = New QueryFilter
   pQueryFilter.WhereClause = "GEO_ID = '" & strGeoID & "'"
   Set QueryLithologyDB = pTable.Search(pQueryFilter, False)
   
   Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CodeUtils.UpdateLithologyDB"
End Function
Public Function GetLithologyTable() As ITable
On Error GoTo ErrorHandler

   Set m_pMxDoc = ThisDocument
   
   Set m_pMap = m_pMxDoc.FocusMap
   
   Dim pTableCollection As ITableCollection
   Set pTableCollection = m_pMap
   
   Dim pTable As ITable
   Set pTable = pTableCollection.Table(0)
    
   Dim pDataset As IDataset
   Set pDataset = pTable
    
   Dim pWorkspace As IWorkspace
   Set pWorkspace = pDataset.Workspace
   
   'Test to see if table is FAXLOG before continuing
   If pDataset.Name <> "lithology" Then
      Set GetLithologyTable = Nothing
      Exit Function
   Else
      Application.StatusBar.Message(0) = "Table was located."
   End If
  
   Set GetLithologyTable = pTable
   Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CodeUtils.UpdateLithologyDB"
End Function

Public Function RelatedRows(pFeatureLayer As IFeatureLayer, pFeature As IFeature) As ISet
On Error GoTo ErrorHandler
      ' get the relationship on the Surface_geology layer
    Dim pFeatureClass As IFeatureClass
    Set pFeatureClass = pFeatureLayer.FeatureClass
    Dim pObjectclass As IObjectClass
    Set pObjectclass = pFeatureClass  'QI
    
    Dim pRelationshipClass As IRelationshipClass
    Dim pEnumRelationshipClass As IEnumRelationshipClass
    Set pEnumRelationshipClass = pObjectclass.RelationshipClasses(esriRelRoleOrigin)
    pEnumRelationshipClass.Reset
    Set pRelationshipClass = pEnumRelationshipClass.Next
      ' find the relationship that has a destination class = lithology (tabular records)
    Do Until pRelationshipClass Is Nothing
      If UCase(pRelationshipClass.DestinationClass.AliasName) = UCase("lithology") Then
        Exit Do
      Else
        Set pRelationshipClass = pEnumRelationshipClass.Next  ' route
      End If
    Loop
    If pRelationshipClass Is Nothing Then
      MsgBox "Error finding correct relationship.", vbExclamation + vbOKOnly
      Exit Function
    End If
    
      ' find the related lithology data
      ' pRelatedSet = lithology records
    Dim pRelatedSet As ISet
    Set pRelatedSet = pRelationshipClass.GetObjectsRelatedToObject(pFeature)
    pRelatedSet.Reset
    
    If pRelatedSet.Count > 0 Then
      Set RelatedRows = pRelatedSet
    Else
      Set RelatedRows = Nothing
    End If
Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CodeUtils.RelatedRows"
End Function
Public Function CreateFont(blnUnderline As Boolean) As IFontDisp
On Error GoTo ErrorHandler

    'Set font Properties
    Dim pFont As IFontDisp
    Set pFont = New stdole.StdFont
    With pFont
      .Name = "Arial"
      .Size = 6
      .Bold = False
      .Italic = False
      .Underline = blnUnderline
      .Strikethrough = False
    End With
    Set CreateFont = pFont

Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CodeUtils.CreateFont"
End Function
Public Function ElementSpacing(lngElements As Long, intSeparation As Integer) As Double
On Error GoTo ErrorHandler

  ' Total number of spacings
  Dim intSpacings As Integer
  intSpacings = CInt(lngElements - 1)
  
  ' Total Y extent for spacings
  Dim intSpacing As Integer
  intSpacing = intSeparation * intSpacings
  
  ' Distance to move each point
'  Dim dblOffset As Double
'  dblOffset = (intSpacing / 2)
   
  ElementSpacing = (intSpacing / 2)
'  Dim dblStartPoint As Double
'  Dim pPoint As IPoint
'  Set pPoint = New Point
'  pPoint.PutCoords 500000, 500000
'
'  Dim item As Integer
'  ' For each item add the full extent to the y, then subtract half of it to center
'  For item = 1 To intCount
'     dblStartPoint = (pPoint.y + spacing) - dblOffset
'     Debug.Print dblStartPoint; Chr(39) & item
'     dblOffset = dblOffset + 700
'  Next item

Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CodeUtils.ElementSpacing"
End Function

Public Sub GetSurfaceGeo(strTop As String, strBottom As String)
On Error GoTo ErrorHandler

  Dim pMxDoc As IMxDocument
  Dim pMap As IMap
  Dim pAV As IActiveView
  Dim pGFLayer As IGeoFeatureLayer
  Dim pFSel As IFeatureSelection
  Dim pDataset As IDataset
  
  Dim pFWorkspace As IFeatureWorkspace
  Dim pSurfaceGeoFC As IFeatureClass
  Dim pQD As IQueryDef
  Dim pCursor As ICursor
  Dim pRow As IRow
  Dim i As Long
  Dim GeoIds() As String
  
  Set pMxDoc = ThisDocument
  Set pMap = pMxDoc.FocusMap
  Set pAV = pMap
  Set pGFLayer = FindLayer("Surface Geology", pMap)
  Set pFSel = pGFLayer
  Set pSurfaceGeoFC = pGFLayer.FeatureClass
  Set pDataset = pSurfaceGeoFC
  Set pFWorkspace = pDataset.Workspace
  
  Set pQD = pFWorkspace.CreateQueryDef
  pQD.SubFields = "l1.geo_id"
  pQD.Tables = "lithology as l1, lithology as l2"
  pQD.WhereClause = "l1.geo_id = l2.geo_id and l1.objectid <> l2.objectid " & _
                    "and l1.lithology = '" & strTop & "' and l2.lithology = '" & strBottom & "' " & _
                    "and l1.layer < l2.layer"
  Set pCursor = pQD.Evaluate
  
  i = 0
  Set pRow = pCursor.NextRow
  If pRow Is Nothing Then
    MsgBox "Lithology " & strTop & " does NOT occur above " & strBottom & ".", vbCritical
    pMap.ClearSelection
    pAV.PartialRefresh 4, Nothing, pAV.Extent
    Exit Sub
  End If
  
  Do Until pRow Is Nothing
    i = i + 1
    ReDim Preserve GeoIds(1 To i)
    GeoIds(i) = "'" & pRow.Value(0) & "'"
    Set pRow = pCursor.NextRow
  Loop
  
  Dim sGeoIds As String
  sGeoIds = Join(GeoIds, ",")
  
  Dim pQF As IQueryFilter
  Set pQF = New QueryFilter
  pQF.WhereClause = "GEO_ID in (" & sGeoIds & ")"
  pFSel.SelectFeatures pQF, esriSelectionResultNew, False
  Dim pSelEvents As ISelectionEvents
  Set pSelEvents = pMap
  pSelEvents.SelectionChanged
  
  pAV.Refresh

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CodeUtils.GetSurfaceGeo"
End Sub
Public Sub LithologyThicknessQuery(strTop As String, strBottom As String, strTopOp As String, strBotOp As String, intTopThickness As Integer, intBotThickness As Integer)
On Error GoTo ErrorHandler

  Dim pMxDoc As IMxDocument
  Dim pMap As IMap
  Dim pAV As IActiveView
  Dim pGFLayer As IGeoFeatureLayer
  Dim pFSel As IFeatureSelection
  Dim pDataset As IDataset
  
  Dim pFWorkspace As IFeatureWorkspace
  Dim pSurfaceGeoFC As IFeatureClass
  Dim pQD As IQueryDef
  Dim pCursor As ICursor
  Dim pRow As IRow
  Dim i As Long
  Dim GeoIds() As String
  
  Set pMxDoc = ThisDocument
  Set pMap = pMxDoc.FocusMap
  Set pAV = pMap
  Set pGFLayer = FindLayer("Surface Geology", pMap)
  Set pFSel = pGFLayer
  Set pSurfaceGeoFC = pGFLayer.FeatureClass
  Set pDataset = pSurfaceGeoFC
  Set pFWorkspace = pDataset.Workspace
  
  Set pQD = pFWorkspace.CreateQueryDef
  pQD.SubFields = "l1.geo_id"
  pQD.Tables = "lithology as l1, lithology as l2"
  pQD.WhereClause = "l1.geo_id = l2.geo_id and l1.objectid <> l2.objectid " & _
                    "and l1.lithology = '" & strTop & "' and l2.lithology = '" & strBottom & "' " & _
                    "and l1.thickness " & strTopOp & " " & intTopThickness & " and l2.thickness " & strBotOp & " " & intBotThickness
  Set pCursor = pQD.Evaluate
  
  i = 0
  Set pRow = pCursor.NextRow
  If pRow Is Nothing Then
    MsgBox "Lithology " & strTop & " with a thickness " & strTopOp & " " & intTopThickness & " does NOT occur above " & strBottom & " with a thickness " & strBotOp & " " & intBotThickness & ".", vbCritical
    pMap.ClearSelection
    pAV.PartialRefresh 4, Nothing, pAV.Extent
    Exit Sub
  End If
  
  Do Until pRow Is Nothing
    i = i + 1
    ReDim Preserve GeoIds(1 To i)
    GeoIds(i) = "'" & pRow.Value(0) & "'"
    Set pRow = pCursor.NextRow
  Loop
  
  If i = 0 Then Exit Sub
  
  Dim sGeoIds As String
  sGeoIds = Join(GeoIds, ",")
  
  Dim pQF As IQueryFilter
  Set pQF = New QueryFilter
  pQF.WhereClause = "GEO_ID in (" & sGeoIds & ")"
  pFSel.SelectFeatures pQF, esriSelectionResultNew, False
  Dim pSelEvents As ISelectionEvents
  Set pSelEvents = pMap
  pSelEvents.SelectionChanged
  
  pAV.Refresh

Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "CodeUtils.LithologyThicknessQuery"
End Sub


