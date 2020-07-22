Attribute VB_Name = "Utilities"
Option Explicit
' ESRI hopes that you will find these sample scripts and other items useful and that they will contribute to
' your success in using ArcInfo. Please note that these samples are not supported by ESRI. These samples
' are provided for non-commercial purposes only. Permission to use, copy, and distribute is hereby granted,
' provided there is no charge or fee for such copies. THESE SAMPLES ARE PROVIDED
' "AS IS", WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANT ABILITY AND
' FITNESS FOR A PARTICULAR PURPOSE. ESRI shall not be liable for any damages under any theory
' of law related to Licensee's use of these samples, even if ESRI is advised of the possibility of such damage.

' The following variable requires you Add a Reference to Microsoft Runtime Scripting
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private m_pMxDoc As IMxDocument
Private m_pActiveView As IActiveView
Private m_pQFilter As IQueryFilter
Private m_pFClass As IFeatureClass
Private m_pFCursor As IFeatureCursor
Public g_Stoptimer As Boolean

Public Function ZoomToFeature(pMap As IMap, pFLayer As IFeatureLayer, strFieldName As String, strFeature As String) As Envelope
On Error GoTo ErrorHandler
    If pMap Is Nothing Or pFLayer Is Nothing Or strFieldName = "" Or strFieldName = "" Then
       Set ZoomToFeature = Nothing
       Exit Function
    End If
    
    Set m_pQFilter = New QueryFilter
    m_pQFilter.WhereClause = strFieldName & " = '" & strFeature & "'"
    'Debug.Print strFieldName & " = '" & strFeature & "'"
    
    Set m_pFClass = pFLayer.FeatureClass
    Set m_pFCursor = m_pFClass.Search(m_pQFilter, True)
    
    Dim pFeature As IFeature
    Set pFeature = m_pFCursor.NextFeature
    
    If pFeature Is Nothing Then
       'Debug.Print "pfeature is nothing!"
       Set ZoomToFeature = Nothing
       Exit Function
    End If
    
'    Dim FieldIndex As Long
'    Dim pField as  iField
'    Dim pFields as  iFields
'    Set pFields = m_pFClass.Fields
'
'    For I = 0 To (pFields.FieldCount - 1)
'      Set pField = pFields.Field(I)
'      If (pField.Type = esriFieldTypeGeometry) Then
'        FieldIndex = m_pFClass.FindField(I)
'        Exit For
'      End If
'    Next I
'
'    MsgBox pField.name

    Dim pEnvelope As IEnvelope
    Dim pGeometry As IGeometry
'    Set pGeometry = pFeature.ShapeCopy
    
    'Set pEnvelope = New  Envelope
    Dim i As Integer
    i = 0
    Do Until pFeature Is Nothing
      Set pGeometry = pFeature.ShapeCopy
      If i > 0 Then
         pEnvelope.Union pGeometry.Envelope
      Else
         Set pEnvelope = pGeometry.Envelope
      End If
      Set pFeature = m_pFCursor.NextFeature
      i = i + 1
    Loop
    Set ZoomToFeature = pEnvelope
    
    Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "Utilities.ZoomToFeature"
End Function
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
   MsgBox Err.Description, vbInformation, "Utilities.IdentifyFeature"
End Function
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
      If UCase(pFLayer.Name) = UCase(pLayerName) Then
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
            If UCase(pFLayer.Name) = UCase(pLayerName) Then
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
   MsgBox Err.Description, vbInformation, "Utilities.FindLayer"
End Function
Public Function IdentifyValue(pFLayer As IFeatureLayer, strFieldName As String, pPoint As IPoint) As String
'Finds and returns a Value (string) based upon an input FeatureLayer, FieldName and Point
On Error GoTo ErrorHandler
    Dim pIdentify As IIdentify
    Dim pArray As IArray
    Dim pGeometry As IGeometry
    Set pGeometry = pPoint 'QI
    Set pIdentify = pFLayer  'QI
    Set pArray = pIdentify.Identify(pGeometry)
    If pArray Is Nothing Then
      IdentifyValue = ""
      Exit Function
    End If
     
    Dim i As Integer
    Dim pRowIDObj As IRowIdentifyObject
    Dim pRow As IRow
     
    For i = 0 To pArray.Count - 1
      If TypeOf pArray.Element(i) Is IRowIdentifyObject Then
        Set pRowIDObj = pArray.Element(i)
        Set pRow = pRowIDObj.Row
          IdentifyValue = pRow.Value(pRow.Fields.FindField(strFieldName))
      End If
    Next
  Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "Utilities.IdentifyValue"
End Function
Public Function FindLayer_old(pLayerName As String, pMap As IMap) As IFeatureLayer
'*******************
'   OBSOLETE
'*******************
On Error GoTo ErrorHandler
  ' finds a layer based on a name and then returns that layer as an IFeatureLayer

  Dim iLoop As Integer
  Dim pFLayer As IFeatureLayer

  For iLoop = 0 To pMap.LayerCount - 1
    If TypeOf pMap.Layer(iLoop) Is IFeatureLayer Then
      Set pFLayer = pMap.Layer(iLoop)
      If UCase(pFLayer.Name) = UCase(pLayerName) Then
        Set FindLayer = pFLayer
        Exit Function
      End If
    End If
  Next iLoop
  
  Set FindLayer = Nothing
  Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "Utilities.FindLayer_old"
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
   MsgBox Err.Description, vbInformation, "Utilities.FlashFeature"
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
   MsgBox Err.Description, vbInformation, "Utilities.FlashLine"
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
   MsgBox Err.Description, vbInformation, "Utilities.FlashMultiPoint"
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
   MsgBox Err.Description, vbInformation, "Utilities.FlashPoint"
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
   MsgBox Err.Description, vbInformation, "Utilities.FlashPolygon"
End Sub
Public Function FindNamedGraphic(strGraphicName As String) As IGeometry
'Finds a Graphic Element on the Map that has a specified name assigned (IElementProperties:Name)
On Error GoTo ErrorHandler
    
    Dim pMap As IMap
    Dim pDGC As IGraphicsContainer
    Dim pEnumElem As IEnumElement
    Dim pElement As IElement
    Dim pElementProperties As IElementProperties
    Dim pDeleteElement As IElement
    
    'Get ActiveView
    Set m_pMxDoc = ThisDocument
    Set pMap = m_pMxDoc.FocusMap

    ' Find Named Graphic object and change shape
    Set pDGC = pMap 'QI
    pDGC.Reset
    
    Set pElement = pDGC.Next
    If pElement Is Nothing Then
        ' Named Graphic does NOT exist
        Set FindNamedGraphic = Nothing
        Exit Function
    End If
    Do While (Not pElement Is Nothing)
        If (TypeOf pElement Is IFillShapeElement) Then
           'QI
           Set pElementProperties = pElement
           If pElementProperties.Name = strGraphicName Then
              ' Graphic has been located, set variable and exit function
              Set FindNamedGraphic = pElement.Geometry
              Exit Function
           End If
        End If
        Set pElement = pDGC.Next
    Loop
    
    ' Graphics exist but not the Named Graphic
    Set FindNamedGraphic = Nothing

  Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "Utilities.FindNamedGraphic"
End Function
Public Function TableSort(pFLayer As IFeatureLayer, strFieldName As String) As ICursor
'Sorts in Ascending order a FeatureLayer's table based upon a specified Field name
On Error GoTo ErrorHandler
  Dim pMxDoc As IMxDocument
  Dim pFC As IFeatureClass
  Dim pDSet As IDataset
  Dim pWksp As IFeatureWorkspace
  Dim pQDef As IQueryDef
  Dim pTS As ITableSort
  Dim pTC As ITrackCancel
  Dim pC As ICursor
  Dim pRow As IRow
  Dim iRowIndex As Long
  
  Set pMxDoc = ThisDocument
  Set pFC = pFLayer.FeatureClass
  Set pDSet = pFC
  Set pWksp = pDSet.Workspace
  'Set pQDef = pWksp.CreateQueryDef
  
  'pQDef.SubFields = "OBJECTID," & strFieldName
  'pQDef.Tables = "cagis.political_bnd"
  
  Set pTS = New TableSort
  With pTS
    Set .Table = pFC
    .Ascending(strFieldName) = True
    .Fields = strFieldName
  End With
  Set pTC = New CancelTracker
  pTS.Sort pTC
  
  Set TableSort = pTS.Rows
'  iRowIndex = pC.Fields.FindField(strFieldName)
'  Set pRow = pC.NextRow
'  Do While (Not pRow Is Nothing)
'    Debug.Print pRow.Value(0), pRow.Value(iRowIndex)
'    Set pRow = pC.NextRow
'  Loop
  
  Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "Utilities.TableSort"
End Function
Public Sub UpdateCountyComboBox(Optional strMap As String)
'Updates the SelectCounty ComboBox with all counties or just counties in a specified Region
On Error GoTo ErrorHandler
  Dim pMap As IMap
  Dim pLayer As ILayer
  Dim pFLayer As IFeatureLayer
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim pData As IDataStatistics
  Dim blnRWMap As Boolean
  
  If strMap = "" Then
    Set pMap = Utilities.GetMap("St")
    blnRWMap = False
  Else
    Set pMap = Utilities.GetMap("RW")
    blnRWMap = True
  End If
  
  Set pFLayer = Utilities.FindLayer("Counties", pMap)
  
  If Not pMap Is Nothing Then
     Application.StatusBar.Message(0) = pMap.Name
     If Not pFLayer Is Nothing Then
        Application.StatusBar.Message(0) = pFLayer.Name
        Set pFCursor = Utilities.TableSort(pFLayer, "COUNTY")
        If Not pFCursor Is Nothing Then
          Set pFeature = pFCursor.NextFeature
          ' Update SelectCounty ComboBox UIControl
           Project.ThisDocument.SelectCounty.RemoveAll
           Project.ThisDocument.SelectCounty.EditText = "Choose a county:"
           If Not pFeature Is Nothing Then
              Do While Not pFeature Is Nothing
                Project.ThisDocument.SelectCounty.AddItem pFeature.Value(pFeature.Fields.FindField("COUNTY"))
                Set pFeature = pFCursor.NextFeature
              Loop
           Else
             Exit Sub
           End If
        Else
           Exit Sub
        End If
     End If
     If blnRWMap Then
         Project.ThisDocument.SelectCounty.AddItem "See All Counties"
     End If
  End If
  Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "Utilities.UpdateCountyComboBox"
End Sub
Public Sub UpdateSiteComboBox()
'Updates the SelectSite ComboBox with only the available sites in the RWIS Sites Layer
On Error GoTo ErrorHandler
  Dim pMap As IMap
  Dim pLayer As ILayer
  Dim pFLayer As IFeatureLayer
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  
  Set pMap = Utilities.GetMap("RW")
  
  If Not pMap Is Nothing Then
     Application.StatusBar.Message(0) = pMap.Name
     Set pFLayer = Utilities.FindLayer("RWIS Sites", pMap)
     If Not pFLayer Is Nothing Then
        Application.StatusBar.Message(0) = pFLayer.Name
        Set pFCursor = Utilities.TableSort(pFLayer, "SITENAME")
        If Not pFCursor Is Nothing Then
          Set pFeature = pFCursor.NextFeature
          ' Update SelectSite ComboBox UIControl
           Project.ThisDocument.SelectSite.RemoveAll
           Project.ThisDocument.SelectSite.EditText = "Choose a site:"
           If Not pFeature Is Nothing Then
              Do While Not pFeature Is Nothing
                Project.ThisDocument.SelectSite.AddItem pFeature.Value(pFeature.Fields.FindField("SITENAME"))
                Set pFeature = pFCursor.NextFeature
              Loop
           Else
             Exit Sub
           End If
        Else
           Exit Sub
        End If
    Else
       MsgBox "Unable to locate Site Layer.", vbCritical, "Missing Data"
    End If
  End If
  
  Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "Utilities.UpdateSiteComboBox"
End Sub

Public Function GetMap(strMap As String) As IMap
'Gets a specified Map (DataFrame)
On Error GoTo ErrorHandler
  Dim mLoop As Integer
  Dim pMaps As IMaps
  Dim pMap As IMap
  
  Set m_pMxDoc = ThisDocument
  Set pMaps = m_pMxDoc.Maps
  
  For mLoop = 0 To pMaps.Count - 1
    Set pMap = pMaps.item(mLoop)
    If Left(pMap.Name, 2) = strMap Then
       Set GetMap = pMap
       Exit Function
    End If
  Next mLoop
  
  Set GetMap = Nothing
  
  Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "Utilities.GetMap"
End Function
Public Sub PopulateWithUniqueValues(pCBO As ComboBox, pFLayer As IFeatureLayer, strFieldName As String, strComboBox As String)
'Populates a specified ComboBox with a list of Unique values
On Error GoTo ErrorHandler

  Dim pFeatureCursor As IFeatureCursor
  Dim pQueryFilter As IQueryFilter
  Dim pEnumVariant As IEnumVariantSimple
  Dim pDataStats As IDataStatistics
  
  Set pQueryFilter = New QueryFilter
  pQueryFilter.WhereClause = strFieldName & " <> ''"
  Set pFeatureCursor = pFLayer.Search(pQueryFilter, False)
  Set pDataStats = New DataStatistics
  pDataStats.Field = strFieldName
  Set pDataStats.Cursor = pFeatureCursor
  pDataStats.SampleRate = -1
  pDataStats.SimpleStats = True
  Set pEnumVariant = pDataStats.UniqueValues
  pEnumVariant.Reset
  
  pCBO.RemoveAll
  Select Case strComboBox
  Case "Site"
    ' Update SelectSite ComboBox UIControl
     pCBO.EditText = "Choose a site:"
  Case "County"
    ' Update SelectCounty ComboBox UIControl
     pCBO.EditText = "Choose a county:"
     'Exit Sub
  End Select
  
  If pDataStats.UniqueValueCount > 0 Then
    Dim lNextUniqueVal As Long
    For lNextUniqueVal = 0 To pDataStats.UniqueValueCount - 1
      'Debug.Print pEnumVariant.Next
      pCBO.AddItem pEnumVariant.Next
    Next lNextUniqueVal
  End If
  
  Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "Utilities.PopulateWithUniqueValues"
End Sub
Public Sub UpdateComboBoxes()
'Updates the SelectSite and SelectCounty ComboBox(s)
'Issued following a change in Region
On Error GoTo ErrorHandler
  Utilities.UpdateCountyComboBox
  Utilities.UpdateSiteComboBox
  
  Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "Utilities.UpdateComboBoxes"
End Sub

Public Sub LocateCounty()
'***************************************************
'   OBSOLETE
'***************************************************
On Error GoTo ErrorHandler
  MsgBox g_intX & g_intY
  Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "Utilities.LocateCounty"
End Sub

Public Sub FlashSites()
'Turns on Flashing Alarms and System Beep
On Error GoTo ErrorHandler
  Dim pMap As IMap
  Dim pFLayer As IFeatureLayer
  Dim pMPoint As IMultipoint
  Dim pFCursor As IFeatureCursor
  Dim pQFilter As IQueryFilter
  Dim pFeature As IFeature
  Dim pGeometry As IGeometry
  Dim pGeo2 As IGeometry
  Dim pGeoColl As IGeometryCollection
  Dim pSpRef1 As ISpatialReference
  Dim pGCS As GeographicCoordinateSystem
  Dim psPRFC As SpatialReferenceEnvironment
  Dim pPCS As IProjectedCoordinateSystem
  Dim pSpRef2 As ISpatialReference
  Dim PauseTime, Start, Finish, TotalTime
  Dim i As Integer
  
  Set m_pMxDoc = ThisDocument
  Set pMap = m_pMxDoc.FocusMap
    
  Set pFLayer = pMap.Layer(0)
  
  Set pMPoint = New Multipoint

  Set pQFilter = New QueryFilter
  pQFilter.WhereClause = "ALARMSTATU = 1"
  Set pFCursor = pFLayer.Search(pQFilter, False)

  Set pFeature = pFCursor.NextFeature
  
    
  'Set Projections
  'Establish Input Spatial Reference System
  Set psPRFC = New SpatialReferenceEnvironment
  Set pGCS = psPRFC.CreateGeographicCoordinateSystem(esriSRGeoCS_NAD1983)
  Set pSpRef1 = pGCS 'QI
   
  'Establish Output Spatial Reference System
  Set pPCS = psPRFC.CreateProjectedCoordinateSystem(esriSRProjCS_NAD1983UTM_18N)
  Set pSpRef2 = pPCS 'QI
  Set pGeoColl = pMPoint
    
  Do While Not pFeature Is Nothing
    Set pGeometry = pFeature.ShapeCopy
    Set pGeo2 = Utilities.ProjectGeometry(pMap, pGeometry, pSpRef1, pSpRef2)
    pGeoColl.AddGeometry pGeo2
    Set pFeature = pFCursor.NextFeature
  Loop
  
  m_pMxDoc.ActiveView.ScreenDisplay.StartDrawing 0, esriNoScreenCache
  Utilities.FlashMultiPoint m_pMxDoc.ActiveView.ScreenDisplay, pMPoint
  ' Finish drawing on screen
  m_pMxDoc.ActiveView.ScreenDisplay.FinishDrawing
   
  g_Stoptimer = False
  PauseTime = 65   ' Set duration.
  Start = Timer    ' Set start time.
  Do While Timer < Start + PauseTime
      DoEvents     ' Yield to other processes.
      m_pMxDoc.ActiveView.ScreenDisplay.StartDrawing 0, esriNoScreenCache
      Utilities.FlashMultiPoint m_pMxDoc.ActiveView.ScreenDisplay, pMPoint
      ' Finish drawing on screen
      m_pMxDoc.ActiveView.ScreenDisplay.FinishDrawing
      For i = 1 To 3    ' Loop 3 times.
          Beep    ' Sound a tone.
      Next i
      g_Stoptimer = Utilities.g_Stoptimer
      If g_Stoptimer Then
         Exit Sub
      End If
  Loop
    Finish = Timer                ' Set end time.
    TotalTime = Finish - Start    ' Calculate total time.
  
  Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "Utilities.FlashSites"
End Sub
Public Function ProjectGeometry(pMap As IMap, pGeometry As IGeometry, pInSpatialReference As ISpatialReference, pOutSpatialReference As ISpatialReference) As IGeometry
'Projects Geometries based upon Input and Output Spatial References (Projections)
On Error GoTo ErrorHandler
  
  'Debug.Print pGeometry.Envelope.LowerLeft.X & " " & pGeometry.Envelope.LowerLeft.Y
        
  Dim pGeo As IGeometry
  ' Switch functions based on Geomtry type
  Select Case pGeometry.GeometryType
    Case esriGeometryPolyline
      Dim pPolyline As IPolyline
      Set pPolyline = pGeometry
      Set pGeo = pPolyline 'QI
      Dim pGeoPolyline As IPolyline
      Set pGeo.SpatialReference = pInSpatialReference
      pGeo.Project pOutSpatialReference
          
      Set pGeoPolyline = pGeo 'QI
      Set ProjectGeometry = pGeoPolyline
        
    Case esriGeometryPolygon
      Dim pPolygon As IPolygon
      Set pPolygon = pGeometry
      Set pGeo = pPolygon 'QI
      Dim pGeoPolygon As IPolygon
      Set pGeo.SpatialReference = pInSpatialReference
      pGeo.Project pOutSpatialReference
          
      Set pGeoPolygon = pGeo 'QI
      Set ProjectGeometry = pGeoPolygon
        
    Case esriGeometryPoint
      Dim pPoint As IPoint
      Set pPoint = pGeometry 'QI
      Set pGeo = pPoint 'QI
      Dim pGeoPoint As IPoint
      Set pGeo.SpatialReference = pInSpatialReference
      pGeo.Project pOutSpatialReference
          
      Set pGeoPoint = pGeo 'QI
      'Debug.Print pGeoPoint.Envelope.LowerLeft.X & " " & pGeoPoint.Envelope.LowerLeft.Y
      Set ProjectGeometry = pGeoPoint
        
    Case esriGeometryMultipoint
      Dim pMultipoint As IMultipoint
      Set pMultipoint = pGeometry 'QI
      Set pGeo = pMultipoint 'QI
      Dim pGeoMultipoint As IMultipoint
      Set pGeo.SpatialReference = pInSpatialReference
      pGeo.Project pOutSpatialReference
          
      Set pGeoMultipoint = pGeo 'QI
      Set ProjectGeometry = pGeoMultipoint
        
    Case esriGeometryEnvelope
      Dim pEnvelope As IEnvelope
      Set pEnvelope = pGeometry
      Set pGeo = pEnvelope 'QI
      Dim pGeoEnvelope As IEnvelope
      Set pGeo.SpatialReference = pInSpatialReference
      pGeo.Project pOutSpatialReference
  End Select
  Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "Utilities.ProjectGeometry"
End Function
Public Function StopTimer() As Boolean
On Error GoTo ErrorHandler
  StopTimer = g_Stoptimer
  Exit Function 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "Utilities.StopTimer"
End Function
Public Sub DisableAlarms()
'Disables Alarms
On Error GoTo ErrorHandler
  g_Stoptimer = True
   Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "Utilities.DisableAlarms"
End Sub

Public Sub LoadRWISLayers(strRegion As String)
'Loads RWIS Layers into Map following a selection from SelectRegion ComboBox
On Error GoTo ErrorHandler
  Dim pMap As IMap
  Dim pLayer As ILayer
  Dim pEnumLayer As IEnumLayer
  Dim intRegion As Integer
  Dim pArray, pName, n
  Dim strLayerDir As String
  Dim strLayerFile As String
  Dim pGxLayer As IGxLayer
  Dim pGxFile As IGxFile
  Dim i As Integer
  Dim pFLayer As IFeatureLayer
  
  Set m_pMxDoc = ThisDocument
  Set pMap = GetMap("RW")
  
  ' Remove existing layers
  If pMap.LayerCount > 0 Then
    Set pEnumLayer = pMap.Layers
    pEnumLayer.Reset
    
    Set pLayer = pEnumLayer.Next
    Do While Not pLayer Is Nothing
       pMap.DeleteLayer pLayer
       'MsgBox "Would have deleted " & pLayer.name
       Set pLayer = pEnumLayer.Next
    Loop
  End If
  
  ' Returns filename with specified extension. If more than one *.lyr
  ' file exists, the first file found is returned.
  strLayerDir = Environ("MAGISDIR")
  strLayerFile = Dir(strLayerDir & "\Layers\*.lyr")
  
  intRegion = 0
  If Left(strRegion, 2) = "Re" Then
     pName = Split(strRegion, "Region ", -1)
     Dim lngX As Long
     lngX = UBound(pName)
     n = pName(lngX)
     If IsNumeric(n) Then
        intRegion = n
        Do While strLayerFile <> ""   ' Start the loop
            ' Call Dir again without arguments to return the next *.lyr file in the same directory.
          'MsgBox Chr(34) & Left(strLayerFile, 2) & Chr(34) & vbCrLf & Chr(34) & strReg & Chr(34)
          If Left(strLayerFile, 2) = Left("r" & CStr(n), 2) Then
              Set pGxLayer = New GxLayer
              Set pGxFile = pGxLayer
              pGxFile.Path = strLayerDir & "\Layers\" & strLayerFile
          
              pMap.AddLayer pGxLayer.Layer
              'Debug.Print strLayerFile
          End If
          strLayerFile = Dir
        Loop
        'Clip the Map to the extent of the Region
        Utilities.ClipMapToRegion pMap, CInt(n)
     End If
  ElseIf Left(strRegion, 2) = "St" Then
  Else 'Do NOT know what they typed or selected
    Exit Sub
  End If
    
  ' Sort the Layers
  If pMap.LayerCount > 0 Then
     For i = 0 To pMap.LayerCount - 1
          Set pLayer = pMap.Layer(i)
          strLayerFile = Left(pLayer.Name, 4)
          Select Case strLayerFile
          Case "RWIS"
            pMap.MoveLayer pLayer, 0
          Case "Surf"
            pMap.MoveLayer pLayer, 1
          Case "Base"
            pMap.MoveLayer pLayer, 2
          End Select
     Next i
     'm_pMxDoc.UpdateContents
     'Zoom to extent of Counties Layer
     Set pFLayer = Utilities.FindLayer("Counties", pMap)
     If Not pFLayer Is Nothing Then
       Dim pGeoDataset As IGeoDataset
       Set pGeoDataset = pFLayer
       Set m_pActiveView = pMap
       Dim pEnvelope As IEnvelope
       Set pEnvelope = pGeoDataset.Extent
       pEnvelope.Expand 1.075, 1.075, True
       m_pActiveView.Extent = pEnvelope
       m_pActiveView.Refresh
     End If
     'Update the SelectCounty ComboBox to include Only the Counties in Selected Region
     Utilities.UpdateCountyComboBox "RWIS Map"
     Utilities.UpdateSiteComboBox
   Else
      MsgBox "No Layers were loaded.", vbCritical, "Warning"
   End If
   Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "Utilities.LoadRWISLayers"
End Sub
Public Sub TestRegion()
'  Dim strLayerDir As String
'  Dim strLayerFile As String
'  strLayerDir = Environ("MAGISDIR")
'
'  ' Load the Layers
'  strLayerFile = Dir(strLayerDir & "\Layers\*.lyr")
'  MsgBox strLayerFile
 Set m_pMxDoc = ThisDocument
 Utilities.ClipMapToRegion m_pMxDoc.FocusMap, 4
End Sub
Sub ClipMapWithGraphic()
On Error GoTo ErrorHandler
    Dim pMxDocument As IMxDocument
    Dim pMap As IMap
    Dim pViewManager As IViewManager
    Dim pEnumElement As IEnumElement
    Dim pElement As IElement
    Dim pActiveView As IActiveView
    Dim pBorder As IBorder

    Set pMxDocument = Application.Document
    Set pMap = pMxDocument.FocusMap
    Set pViewManager = pMap 'QI
    Set pEnumElement = pViewManager.ElementSelection
    pEnumElement.Reset
    Set pElement = pEnumElement.Next
    If pElement Is Nothing Then
        MsgBox "No element selected"
        Exit Sub
    End If
    
    If pElement.Geometry.GeometryType = esriGeometryPolygon Then
       Set pBorder = New SymbolBorder
       'Use default line symbol
       pMap.ClipGeometry = pElement.Geometry
       pMap.ClipBorder = pBorder
       Set pActiveView = pMap
       pActiveView.Refresh
    End If
    Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "Utilities.ClipMapWithGraphic"
End Sub
Public Sub ClipMapToRegion(pMap As IMap, n As Integer)
On Error GoTo ErrorHandler
  Dim pFLayer As IFeatureLayer
  Dim pGeometry As IGeometry
  Dim pFCursor As IFeatureCursor
  Dim pQFilter As IQueryFilter
  Dim pFeature As IFeature
  Dim pStateMap As IMap
  
  If Left(pMap.Name, 2) = "RW" Then
      Set m_pMxDoc = ThisDocument
      Set pStateMap = Utilities.GetMap("St")
      Set pFLayer = Utilities.FindLayer("Regions", pStateMap)
      Set pQFilter = New QueryFilter
      pQFilter.WhereClause = "REGION = " & n
      Set pFCursor = pFLayer.Search(pQFilter, False)
      Set pFeature = pFCursor.NextFeature
        
      Do While Not pFeature Is Nothing
        Set pGeometry = pFeature.ShapeCopy
        pMap.ClipGeometry = pGeometry
        'Debug.Print pFeature.value(pFeature.Fields.FindField("REGION"))
        Set pFeature = pFCursor.NextFeature
      Loop
      pMap.ClipGeometry = pGeometry
      Set m_pActiveView = pMap
      m_pActiveView.Refresh
  End If
    Exit Sub 'Avoid ErrorHandler
ErrorHandler:
   MsgBox Err.Description, vbInformation, "Utilities.ClipMapToRegion"
End Sub
Public Sub ReportAnnoInfo()
  Dim pApp As IApplication
  Set pApp = Application
  Dim pMxDoc As IMxDocument
  Set pMxDoc = pApp.Document
  
'     Dim pGC As IGraphicsContainer
'     Set pGC = pMxDoc.FocusMap 'QI
'
'     Dim pElement As IElement
'
'     pGC.Reset
'     Set pElement = pGC.Next
'     Do While Not pElement Is Nothing
'       If TypeOf pElement Is ITextElement Then
'         Dim pTE As ITextElement
'         Set pTE = pElement 'QI
'         MsgBox pTE.Text
'       End If
'       Set pElement = pGC.Next
'     Loop
  
  
  Dim pFDOLayer As IFDOGraphicsLayer
  Dim pLayer As ILayer
  Set pLayer = pMxDoc.FocusMap.Layer(0)
  If TypeOf pLayer Is IFDOGraphicsLayer Then
     Set pFDOLayer = pLayer
     MsgBox pLayer.Name
     Dim pGLayer As IGraphicsLayer
     Set pGLayer = pFDOLayer
     pGLayer.Activate pMxDoc.ActiveView.ScreenDisplay
     Dim pMap As IMap
     Set pMap = pMxDoc.FocusMap
     Set pMap.ActiveGraphicsLayer = pGLayer
     Dim pGC2 As IGraphicsContainer
     Set pGC2 = pGLayer 'QI
     Dim pElement As IElement
     pGC2.Reset
     Set pElement = pGC2.Next
     Do While Not pElement Is Nothing
       If TypeOf pElement Is ITextElement Then
         Dim pTE As ITextElement
         Set pTE = pElement 'QI
         MsgBox pTE.Text
       End If
       Set pElement = pGC2.Next
     Loop
  End If
End Sub

Public Sub AddGeologyAnno()
    ' set up fields for mapping
    Dim pOFlds As IFields
    Dim pIFlds As IFieldsSet
    pOFlds = pTargetFClass.FieldsSet
    pIFlds = pSrcFClass.Fields
    Dim pFld As IField
    Dim lSrcFlds() As Long
    Dim lTrgFlds() As Long
    Dim lFld As Long, lExFld As Long, i As Long, lFld2 As Long
    lExFld = 0
    For lFld = 0 To pIFlds.FieldCount - 1
      Set pFld = pIFlds.Field(lFld)
      If Not pFld.Type = esriFieldTypeOID And Not pFld.Type = esriFieldTypeGeometry And Not UCase(pFld.Name) = "ELEMENT" And _
        Not UCase(pFld.Name) = "ANNOTATIONCLASSID" And Not UCase(pFld.Name) = "ZORDER" And _
        pFld.Editable Then
        lExFld = lExFld + 1
      End If
    Next lFld
      
    ReDim lSrcFlds(lExFld) As Long
    ReDim lTrgFlds(lExFld) As Long
      
    i = 0
    For lFld = 0 To pIFlds.FieldCount - 1
      Set pFld = pIFlds.Field(lFld)
      If Not pFld.Type = esriFieldTypeOID And Not pFld.Type = esriFieldTypeGeometry And Not UCase(pFld.Name) = "ELEMENT" And _
      Not UCase(pFld.Name) = "ANNOTATIONCLASSID" And Not UCase(pFld.Name) = "ZORDER" And _
      pFld.Editable Then
        lSrcFlds(i) = lFld
        lTrgFlds(i) = pOFlds.FindField(pFld.Name)
        i = i + 1
      End If
    Next lFld
    Dim pICursor As IFeatureCursor
    Set pICursor = pSrcFClass.Search(Nothing, True)
      
    Dim pIFeat As IFeature
    Set pIFeat = pICursor.NextFeature
    Dim pGLF As IFDOGraphicsLayerFactory
    Set pGLF = New FDOGraphicsLayerFactory
      
    Set pDataset = pTargetFClass
    Dim pFDOGLayer As IFDOGraphicsLayer
    Set pFDOGLayer = pGLF.OpenGraphicsLayer(pDataset.Workspace, pTargetFClass.FeatureDataset, pDataset.Name)
      
    Dim pFDOACon As IFDOAttributeConversion
    Set pFDOACon = pFDOGLayer
      
    pFDOGLayer.BeginAddElements
    pFDOACon.SetupAttributeConversion2 lExFld, lSrcFlds, lTrgFlds
      
    While Not pIFeat Is Nothing
      Set pAnnoFeature = pIFeat
      Set pAClone = pAnnoFeature.Annotation
      Set pGSElement = pAClone.Clone
      pFDOGLayer.DoAddFeature pIFeat, pGSElement, 0
      Set pIFeat = pICursor.NextFeature
      m_frmProg.ProgressBar1.Value = m_frmProg.ProgressBar1.Value + 1
      m_frmProg.lblCurrent = m_frmProg.ProgressBar1.Value
      m_frmProg.lblCurrent.Refresh
    Wend
    pFDOGLayer.EndAddElements

End Sub
Public Sub ListAnnoFeatures()
  Dim pApp As IApplication
  Set pApp = Application
  
  Dim pMxDoc As IMxDocument
  Set pMxDoc = pApp.Document
  
  Dim pFDOLayer As IFDOGraphicsLayer
  Dim pLayer As ILayer
  
  Set pLayer = pMxDoc.FocusMap.Layer(0)
  
  If TypeOf pLayer Is IFDOGraphicsLayer Then
     Set pFDOLayer = pLayer
     Dim pFLayer As IFeatureLayer
     Set pFLayer = pLayer 'QI
     Dim pFCursor As IFeatureCursor
     Set pFCursor = pFLayer.Search(Nothing, True)
     Dim pFeature As IFeature
     Set pFeature = pFCursor.NextFeature
     If pFeature Is Nothing Then
        MsgBox "No features"
        Exit Sub
     End If
     Dim pAnnoFeature As IAnnotationFeature
     Do While Not pFeature Is Nothing
        Set pAnnoFeature = pFeature
        Dim pTextElement As ITextElement
        Set pTextElement = pAnnoFeature.Annotation
        Dim pSymbol As ITextSymbol
        Set pSymbol = pTextElement.Symbol
        MsgBox pTextElement.Text & " " & pSymbol.Font.Name
        Set pFeature = pFCursor.NextFeature
     Loop
  End If
End Sub

