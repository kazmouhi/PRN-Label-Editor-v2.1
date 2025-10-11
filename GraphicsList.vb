Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Data
Imports System.Diagnostics
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.Net
'Imports System.Linq
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ClipperLib
Imports Draw
Imports SVGLib

Namespace Draw

    ''' <summary>
    ''' List of graphic objects
    ''' </summary>
    <Serializable>
    Public Class GraphicsList

        Public ReadOnly _graphicsList As ArrayList
        Public ReadOnly _inMemoryList As ArrayList
        'Private ReadOnly _undoRedo As Command.UndoRedo

        Private _isCut As Boolean

        Public Sub New()
            _graphicsList = New ArrayList()
            _inMemoryList = New ArrayList()
            DrawObject.undoRedo = New Command.UndoRedo()
        End Sub


        ''' <summary>
        ''' Count and this [nIndex] allow to read all graphics objects
        ''' from GraphicsList in the loop.
        ''' </summary>
        Public ReadOnly Property ObjCount As Integer
            Get
                Return _graphicsList.Count
            End Get
        End Property

        ''' <summary>
        ''' Gets or Sets the description
        ''' </summary>
        Public Property Description As String

        ''' <summary>
        ''' SelectedCount and GetSelectedObject allow to read
        ''' selected objects in the loop
        ''' </summary>
        Public ReadOnly Property SelectionCount As Integer
            Get
                Dim n = 0
                For Each o As DrawObject In _graphicsList
                    If o.Selected Then n += 1
                Next
                Return n
            End Get
        End Property


        Default Public ReadOnly Property Item(ByVal index As Integer) As DrawObject
            Get
                If index < 0 OrElse index >= _graphicsList.Count Then Return Nothing
                Return _graphicsList(index)
            End Get
        End Property

        Private _JobSize As PointF
        Private _JobDepth As Single

        Private ColShapeOutline As Color = Color.FromArgb(0, 0, 0)
        Private ColDragHandles As Color = Color.FromArgb(192, 192, 255)
        Private ColPointHandle As Color = Color.FromArgb(0, 0, 0)
        Private ColSelectedPointhandle As Color = Color.FromArgb(220, 20, 60)
        Private ColB1Handle As Color = Color.FromArgb(220, 20, 60)
        Private ColB2Handle As Color = Color.FromArgb(220, 20, 60)
        Private ColRotateStart As Color = Color.FromArgb(220, 20, 60)
        Private ColSelectionRectangle As Color = Color.FromArgb(128, 128, 128)
        Private ColCrosshair As Color = Color.FromArgb(128, 128, 128)
        Private ColLeadIn As Color = Color.FromArgb(243, 146, 80)
        Private ColLeadOut As Color = Color.FromArgb(123, 176, 123)
        Private JobType As New DrawObject.JobTypes


        Private Property ToolData As New DataTable
        Private Property EndmillDiam As Single = 4


        Public Sub SetJobtype(JType As DrawObject.JobTypes)
            JobType = JType
            For Each o As DrawObject In _graphicsList
                o.JobType = JType
            Next
        End Sub

        Public TPShowing As Boolean = False

        Public Sub ToolpathShowing(IsToolPathShowing As Boolean)
            TPShowing = IsToolPathShowing
            For Each o As DrawObject In _graphicsList
                o.ToolPathShowing = IsToolPathShowing
            Next
        End Sub

        Public Sub SetExport(IsExport As Boolean)
            For Each o As DrawObject In _graphicsList
                o.SetExport = IsExport
            Next
        End Sub


        Public Sub SetJobDimensions(JobSize As PointF, JobDepth As Single)
            _JobSize = JobSize
            _JobDepth = JobDepth

            For Each shape As DrawObject In _graphicsList
                shape.StockDepth = _JobDepth
            Next
        End Sub

        Public Sub SetDefaultToolData(ToolData As String)
            Dim p As String = ToolData

        End Sub



        Public Sub Add(ByVal obj As DrawObject)
            ' insert to the top of z-order
            UpdateColours(ColShapeOutline, ColDragHandles, ColPointHandle, ColSelectedPointhandle, ColB1Handle, ColB1Handle, ColRotateStart, ColSelectionRectangle, ColCrosshair, ColLeadIn, ColLeadOut, obj)

            obj.JobType = JobType
            obj.ToolPathShowing = TPShowing

            If obj.Tag = "rect" Then
                obj.ReplacePointArray(obj.ConvertToPath)
            End If


            _graphicsList.Insert(0, obj)

            ' Convert to ArrayList for command
            Dim graphicsArrayList As New ArrayList()
            graphicsArrayList.AddRange(_graphicsList)

            Dim create = New Command.CreateCommand(obj, graphicsArrayList)
            DrawObject.undoRedo.AddCommand(create)

        End Sub

        ' **************   Read from SVG
        Public Sub AddFromSvg(ByVal ele As SvgElement)
            While ele IsNot Nothing
                Dim o As DrawObject = CreateDrawObject(ele)
                If o IsNot Nothing Then Me.Add(o)
                Dim child As SvgElement = ele.getChild()
                While child IsNot Nothing
                    AddFromSvg(child)
                    child = child.getNext()
                End While
                ele = ele.getNext()
            End While
        End Sub

        Public Function AreItemsInMemory() As Boolean
            Return _inMemoryList.Count > 0
        End Function

        ''' <summary>
        ''' Clear all objects in the list
        ''' </summary>
        ''' <returns>
        ''' true if at least one object is deleted
        ''' </returns>
        Public Function Clear() As Boolean
            Dim result = _graphicsList.Count > 0
            _graphicsList.Clear()
            Return result
        End Function

        Public Sub CutSelection()
            Dim i As Integer
            Dim n = _graphicsList.Count
            _inMemoryList.Clear()

            For i = n - 1 To 0 Step -1
                If CType(_graphicsList(i), DrawObject).Selected Then
                    _inMemoryList.Add(_graphicsList(i))
                End If
            Next
            _isCut = True

            ' Convert to ArrayList for command
            Dim graphicsArrayList As New ArrayList()
            graphicsArrayList.AddRange(_graphicsList)
            Dim memoryArrayList As New ArrayList()
            memoryArrayList.AddRange(_inMemoryList)

            Dim cmd = New Command.CutCommand(graphicsArrayList, memoryArrayList)
            cmd.Execute()
            DrawObject.undoRedo.AddCommand(cmd)
        End Sub

        Public Function CopySelection() As Boolean
            Dim result = False
            Dim n = _graphicsList.Count
            _inMemoryList.Clear()

            For i = n - 1 To 0 Step -1
                If CType(_graphicsList(i), DrawObject).Selected Then
                    _inMemoryList.Add(_graphicsList(i))
                    result = True
                    _isCut = False
                End If
            Next

            Return result
        End Function

        Public Sub PasteSelection()
            Dim n = _inMemoryList.Count

            UnselectAll()

            If n > 0 Then
                Dim tempList = New ArrayList()

                Dim i As Integer
                For i = n - 1 To 0 Step -1
                    tempList.Add(CType(_inMemoryList(CInt(i)), DrawObject).Clone())
                Next

                If _inMemoryList.Count > 0 Then
                    ' Convert to ArrayList for command
                    Dim graphicsArrayList As New ArrayList()
                    graphicsArrayList.AddRange(_graphicsList)

                    Dim cmd = New Command.PasteCommand(graphicsArrayList, tempList)
                    cmd.Execute()
                    DrawObject.undoRedo.AddCommand(cmd)

                    'If the items are cut, we will not delete it
                    If _isCut Then _inMemoryList.Clear()
                End If
            End If
        End Sub

        ''' <summary>
        ''' Delete selected items
        ''' </summary>
        ''' <returns>
        ''' true if at least one object is deleted
        ''' </returns>
        Public Function DeleteSelection(Optional ByVal IgnoreUndoRedo As Boolean = False) As Boolean
            ' Convert to ArrayList for command
            Dim graphicsArrayList As New ArrayList()
            graphicsArrayList.AddRange(_graphicsList)

            Dim cmd = New Command.DeleteCommand(graphicsArrayList)
            cmd.Execute()

            If IgnoreUndoRedo = False Then
                DrawObject.undoRedo.AddCommand(cmd)
            End If

            Return True
        End Function

        Public Sub Draw(ByVal g As Graphics)
            Dim n = _graphicsList.Count
            Dim o As DrawObject

            For i = n - 1 To 0 Step -1
                o = CType(_graphicsList(i), DrawObject)

                o.Draw(g)
                If o.NeedsConversion = True Then
                    ConvertToSelectedPath(True, o.SelectedHandle)
                End If

            Next
        End Sub

        Public Function GetAllSelected() As List(Of DrawObject)
            Dim selectionList = New List(Of DrawObject)()

            For Each o As DrawObject In _graphicsList
                If o.Selected Then selectionList.Add(o)
            Next
            Return selectionList
        End Function

        'Dave - get all shapes
        Public Function GetAllShapes() As List(Of DrawObject)
            Dim selectionList = New List(Of DrawObject)
            For i = 0 To _graphicsList.Count - 1
                Dim o As DrawObject = CType(_graphicsList.Item(i), DrawObject)
                selectionList.Add(o)
            Next
            Return selectionList
        End Function


        Public Function GetFirstSelected() As DrawObject
            For Each o As DrawObject In _graphicsList
                If o.Selected Then Return o
            Next
            Return Nothing
        End Function

        Public Function GetSelectedObject(ByVal index As Integer) As DrawObject
            Dim n = -1
            For Each o As DrawObject In _graphicsList
                If o.Selected Then
                    n += 1
                    If n = index Then Return o
                End If
            Next
            Return Nothing
        End Function

        Public Function GetXmlString(AddToolData As Boolean) As String

            Dim sXml = ""
            Dim n = _graphicsList.Count
            For i = n - 1 To 0 Step -1
                sXml += CType(_graphicsList(i), DrawObject).GetXmlStr(AddToolData)
            Next
            Return sXml
        End Function

        Public Function IsAnythingSelected() As Boolean
            For Each o As DrawObject In _graphicsList
                If o.Selected Then Return True
            Next

            Return False
        End Function

        Public Sub Move(ByVal movedItemsList As ArrayList, ByVal delta As PointF)
            Dim cmd = New Command.MoveCommand(movedItemsList, delta)
            DrawObject.undoRedo.AddCommand(cmd)
        End Sub

        Public Function MoveSelectionToBack() As Boolean
            Dim n = _graphicsList.Count
            Dim tempList = New ArrayList()

            For i = n - 1 To 0 Step -1
                If CType(_graphicsList(i), DrawObject).Selected = True Then
                    tempList.Add(_graphicsList(i))
                End If
            Next

            ' Convert to ArrayList for command
            Dim graphicsArrayList As New ArrayList()
            graphicsArrayList.AddRange(_graphicsList)

            Dim cmd = New Command.SendToBackCommand(graphicsArrayList, tempList)
            cmd.Execute()
            DrawObject.undoRedo.AddCommand(cmd)

            Return True
        End Function

        ''' <summary>
        ''' Move selected items to front (beginning of the list)
        ''' </summary>
        ''' <returns>
        ''' true if at least one object is moved
        ''' </returns>
        Public Function MoveSelectionToFront() As Boolean
            Dim tempList = New ArrayList()

            For Each o As DrawObject In _graphicsList
                If o.Selected Then
                    tempList.Add(o)
                End If
            Next

            If tempList.Count = 0 Then Return False

            ' Convert to ArrayList for command
            Dim graphicsArrayList As New ArrayList()
            graphicsArrayList.AddRange(_graphicsList)

            Dim cmd = New Command.BringToFrontCommand(graphicsArrayList, tempList)
            cmd.Execute()
            DrawObject.undoRedo.AddCommand(cmd)

            Return True
        End Function

        ''' <summary>
        ''' Shape Property Changed
        ''' </summary>
        Public Sub PropertyChanged(ByVal itemChanged As GridItem, ByVal oldVal As Object)
            Dim i As Integer
            Dim n = _graphicsList.Count
            Dim tempList = New ArrayList()

            For i = n - 1 To 0 Step -1
                If CType(_graphicsList(i), DrawObject).Selected Then
                    tempList.Add(_graphicsList(i))
                End If
            Next

            Dim cmd = New Command.PropertyChangeCommand(tempList, itemChanged, oldVal)

            DrawObject.undoRedo.AddCommand(cmd)
        End Sub

        Public Sub Redo()
            DrawObject.undoRedo.Redo()
        End Sub

        'Public Sub Resize(ByVal newscale As SizeF, ByVal oldscale As SizeF)
        ' For Each o As DrawObject In _graphicsList
        '  o.Resize(newscale, oldscale)
        ' Next
        'End Sub

        Public Sub ResizeCommand(ByVal obj As DrawObject, ByVal old As PointF, ByVal newP As PointF, ByVal handle As Integer)
            Dim cmd = New Command.ResizeCommand(obj, old, newP, handle)
            DrawObject.undoRedo.AddCommand(cmd)
        End Sub

        Public Sub SelectAll()
            For Each o As DrawObject In _graphicsList
                o.Selected = True
                o.SelectedHandle = 0

            Next
        End Sub

        Public Sub SelectInRectangle(ByVal rectangle As RectangleF)
            UnselectAll()

            For Each o As DrawObject In _graphicsList
                If o.IntersectsWith(dozoom(rectangle)) Then
                    o.Selected = True
                End If
            Next
        End Sub


        Public Function dozoom(ByVal r As RectangleF) As RectangleF
            Return New RectangleF(r.Location.X / DrawObject.Zoom, r.Location.Y / DrawObject.Zoom, r.Size.Width / DrawObject.Zoom, r.Size.Height / DrawObject.Zoom)
        End Function

        ' *************************************************
        Public Sub Undo()
            DrawObject.undoRedo.Undo()
        End Sub

        Public Sub UnselectAll(Optional ByVal IsolatedObject As DrawObject = Nothing)
            For Each o As DrawObject In _graphicsList
                o.Selected = False
                'o.EditPoints = False
                o.SelectedHandle = 0 'DAVE - Added to reset any selected handles when shape deselected
                'o.ResetRotateStartPoint()
            Next

        End Sub

        Private Function CreateDrawObject(ByVal svge As SvgElement) As DrawObject
            Dim o As DrawObject = Nothing
            Select Case svge.getElementType()
                Case SvgElement._SvgElementType.typeLine
                    o = New DrawLine(CType(svge, SvgLine))
                Case SvgElement._SvgElementType.typeRect
                    o = New DrawRectangle(CType(svge, SvgRect))
                Case SvgElement._SvgElementType.typeEllipse
                    o = New DrawEllipse(CType(svge, SvgEllipse))
                Case SvgElement._SvgElementType.typeCircle
                    o = New DrawCircle(CType(svge, SvgCircle))
                Case SvgElement._SvgElementType.typePolyline
                    o = New DrawPolyline(CType(svge, SvgPolyline))
                Case SvgElement._SvgElementType.typePolygon
                    o = New DrawPolygon(CType(svge, SvgPolygon))
                Case SvgElement._SvgElementType.typeGroup
                    o = CreateGroup(CType(svge, SvgGroup))
                Case SvgElement._SvgElementType.typePath
                    o = New DrawPath(CType(svge, SvgPath))
    'Case SvgElement._SvgElementType.typeImage
    ' o = New DrawImage(CType(svge, SvgImage))
    'Case SvgElement._SvgElementType.typeText
    ' o = New DrawText(CType(svge, SvgText))
                Case SvgElement._SvgElementType.typeDesc
                    Description = CType(svge, SvgDesc).Value
                Case Else
            End Select
            Return o
        End Function

        Private Function CreateGroup(ByVal svg As SvgGroup) As DrawObject
            Dim o As DrawObject = Nothing
            Dim child As SvgElement = svg.getChild()
            If child IsNot Nothing Then AddFromSvg(child)
            Return o
        End Function

        Public Sub FlipSelectedHorizontal()

            For Each o As DrawObject In _graphicsList
                If o.Selected = True Then
                    o.FlipHorizontal()
                End If

            Next

        End Sub



        Public Function ShapesOutsideCanvas(CanvasSize As PointF) As List(Of DrawObject)

            UnselectAll()

            Dim RectList As New List(Of DrawObject)


            For Each o As DrawObject In _graphicsList
                Dim TRect As RectangleF = o.dozoom(o.BoundingRectangle)
                With TRect
                    If .X < 0 Then
                        RectList.Add(o)
                        o.Selected = True
                    ElseIf .Y < 0 Then
                        RectList.Add(o)
                        o.Selected = True
                    ElseIf (.X + .Width) > CanvasSize.X Then
                        RectList.Add(o)
                        o.Selected = True
                    ElseIf (.Y + .Height) > CanvasSize.Y Then
                        RectList.Add(o)
                        o.Selected = True
                    End If
                End With
            Next

            Return RectList
        End Function



        Public Sub AlignV1(alignment As DrawObject.Alignment)


            Dim selectionList = New List(Of DrawObject)()
            Dim val As Single = Nothing
            For Each o As DrawObject In _graphicsList
                If o.Selected = True Then
                    Dim r As RectangleF = o.Rectangle
                    If val <> Nothing Then

                        Select Case alignment
                            Case DrawObject.Alignment.top
                                r.Y = val
                            Case DrawObject.Alignment.left
                                r.X = val
                            Case DrawObject.Alignment.bottom
                                r.Y = val - r.Height
                            Case DrawObject.Alignment.right
                                r.X = val - r.Width
                            Case DrawObject.Alignment.middle
                                r.Y = val - (r.Height / 2)
                            Case DrawObject.Alignment.center
                                r.X = val - (r.Width / 2)
                        End Select

                        o.setRectangle(r)
                    Else
                        Select Case alignment
                            Case DrawObject.Alignment.top
                                val = r.Y
                            Case DrawObject.Alignment.left
                                val = r.X
                            Case DrawObject.Alignment.bottom
                                val = r.Y + r.Height
                            Case DrawObject.Alignment.right
                                val = r.X + r.Width
                            Case DrawObject.Alignment.middle
                                val = o.Rectangle.Y + (r.Height / 2)
                            Case DrawObject.Alignment.center
                                val = o.Rectangle.X + (r.Width / 2)
                        End Select
                    End If
                End If
            Next
        End Sub


        Public Sub Align(alignment As DrawObject.Alignment)

            Dim selectionList = New List(Of DrawObject)()
            Dim cmds As New List(Of Command.ICommand)
            Dim val As Single = Nothing
            For Each o As DrawObject In _graphicsList
                If o.Selected = True Then
                    Dim r As RectangleF = o.BoundingRectangle
                    If val <> Nothing Then

                        Select Case alignment
                            Case DrawObject.Alignment.top
                                r.Y = val
                            Case DrawObject.Alignment.left
                                r.X = val
                            Case DrawObject.Alignment.bottom
                                r.Y = val - r.Height
                            Case DrawObject.Alignment.right
                                r.X = val - r.Width
                            Case DrawObject.Alignment.middle
                                r.Y = val - (r.Height / 2)
                            Case DrawObject.Alignment.center
                                r.X = val - (r.Width / 2)
                        End Select

                        Dim delta As PointF = o.MoveBoundingRect(r.Location)

                        'Dim delta As PointF = o.MoveBoundingRectOffset(r.Location)
                        Dim movelist As New ArrayList()
                        movelist.Add(o)
                        cmds.Add(New Command.MoveCommand(movelist, delta))
                        'o.Move(delta)

                        'done to ensure no change in size
                    Else
                        Select Case alignment
                            Case DrawObject.Alignment.top
                                val = r.Y
                            Case DrawObject.Alignment.left
                                val = r.X
                            Case DrawObject.Alignment.bottom
                                val = r.Y + r.Height
                            Case DrawObject.Alignment.right
                                val = r.X + r.Width
                            Case DrawObject.Alignment.middle
                                val = r.Y + (r.Height / 2)
                            Case DrawObject.Alignment.center
                                val = r.X + (r.Width / 2)
                        End Select
                    End If
                End If
            Next


            Dim cmd = New Command.Commands(cmds)
            DrawObject.undoRedo.AddCommand(cmd)

        End Sub


        Public Sub AlignTops()
            Align(DrawObject.Alignment.top)
        End Sub
        Public Sub AlignBottoms()
            Align(DrawObject.Alignment.bottom)
        End Sub
        Public Sub AlignVerticals()
            Align(DrawObject.Alignment.center)
        End Sub

        Public Sub AlignHorizontals()
            Align(DrawObject.Alignment.middle)
        End Sub

        Public Sub AlignLefts()
            Align(DrawObject.Alignment.left)
        End Sub

        Public Sub AlignRights()

            Align(DrawObject.Alignment.right)
        End Sub

        Public Sub SpaceEvenly(isHorizontal As Boolean)
            Dim selectionList As New List(Of RectangleF)
            Dim SelectedObjects As New List(Of DrawObject)

            For Each o As DrawObject In _graphicsList
                If o.Selected = True Then
                    selectionList.Add(o.Rectangle)
                    SelectedObjects.Add(o)
                End If
            Next

            If SelectedObjects.Count <= 0 Then Exit Sub

            Dim NewList As New List(Of RectangleF)

            If isHorizontal = True Then
                NewList = SpaceRectsVertically(selectionList)

            Else
                NewList = SpaceRectsHorizontally(selectionList)

            End If


            For i = 0 To SelectedObjects.Count - 1
                SelectedObjects(i).setRectangle(NewList(i))
            Next i
        End Sub

        Public Function SpaceRectsVertically(rectangles As List(Of RectangleF)) As List(Of RectangleF)
            rectangles = rectangles.OrderBy(Function(r) r.Top).ToList()

            Dim highestRect As RectangleF = rectangles.OrderByDescending(Function(r) r.Top).First()
            Dim lowestRect As RectangleF = rectangles.OrderBy(Function(r) r.Top).First()
            Dim distance As Single = DrawObject.center(highestRect).Y - DrawObject.center(lowestRect).Y
            Dim offset As Single = distance / (rectangles.Count - 1)
            For i As Integer = 1 To rectangles.Count - 1
                Dim pp = DrawObject.center(rectangles(i - 1)).Y
                Dim top As Single = pp + offset
                top -= rectangles(i).Height / 2
                rectangles(i) = New RectangleF(rectangles(i).Left, top, rectangles(i).Width, rectangles(i).Height)
            Next
            Return rectangles
        End Function

        Public Function SpaceRectsHorizontally(rectangles As List(Of RectangleF)) As List(Of RectangleF)
            rectangles = rectangles.OrderBy(Function(r) r.Left).ToList()

            Dim leftmostRect As RectangleF = rectangles.OrderBy(Function(r) r.Left).First()
            Dim rightmostRect As RectangleF = rectangles.OrderByDescending(Function(r) r.Right).First()
            Dim distance As Single = DrawObject.center(rightmostRect).X - DrawObject.center(leftmostRect).X
            Dim offset As Single = distance / (rectangles.Count - 1)
            For i As Integer = 1 To rectangles.Count - 1
                Dim pp = DrawObject.center(rectangles(i - 1)).X
                Dim left As Single = pp + offset
                left -= rectangles(i).Width / 2
                rectangles(i) = New RectangleF(left, rectangles(i).Top, rectangles(i).Width, rectangles(i).Height)
            Next
            Return rectangles
        End Function

        Public Sub ConvertToSelectedPath(Optional ByVal EditPoints As Boolean = False, Optional SelectedPoint As Integer = -1)
            'Exit Sub 'For testing
            Dim i As Integer

            Dim cutlist As New ArrayList
            Dim pastelist As New ArrayList
            Dim n = _graphicsList.Count
            For i = n - 1 To 0 Step -1
                If CType(_graphicsList(i), DrawObject).Selected = True Then
                    Dim PS As DrawObject = CType(_graphicsList(i), DrawObject)
                    Dim tmpAngle As Single = PS.CurrentAngle
                    PS.Rotate(0)
                    Dim pcs = PS.ConvertToPath()

                    If Not pcs Is Nothing Then
                        Dim path As New DrawPath(pcs)
                        path.TOOLDATA = PS.TOOLDATA
                        path.JobType = JobType
                        path.Selected = True
                        path.EditPoints = EditPoints
                        path.SelectedHandle = SelectedPoint
                        path.RotateStartPoint = PS.RotateStartPoint
                        path.Rotate(tmpAngle)
                        'path.CurrentAngle =

                        path.Name = PS.Name

                        pastelist.Add(path)
                        cutlist.Add(_graphicsList(i))
                        _graphicsList.RemoveAt(i)
                        _graphicsList.Insert(i, path)

                    End If
                End If
            Next

            If pastelist.Count > 0 Then
                Dim cmds As New List(Of Command.ICommand)
                cmds.Add(New Command.CutCommand(_graphicsList, cutlist))
                cmds.Add(New Command.PasteCommand(_graphicsList, pastelist))
                DrawObject.undoRedo.AddCommand(New Command.Commands(cmds))
            End If

        End Sub

        Public Sub ConvertAllToPath()
            'Exit Sub

            Dim i As Integer

            Dim cutlist As New ArrayList
            Dim pastelist As New ArrayList
            Dim n = _graphicsList.Count
            For i = n - 1 To 0 Step -1
                Dim PS As DrawObject = CType(_graphicsList(i), DrawObject)
                Dim tmpAngle As Single = PS.CurrentAngle
                PS.Rotate(0)
                Dim pcs = PS.ConvertToPath()

                If Not pcs Is Nothing Then
                    Dim path As New DrawPath(pcs)
                    path.TOOLDATA = PS.TOOLDATA
                    path.JobType = JobType
                    path.Selected = PS.Selected
                    path.EditPoints = PS.EditPoints
                    path.SelectedHandle = PS.SelectedHandle
                    path.RotateStartPoint = PS.RotateStartPoint
                    path.CurrentAngle = tmpAngle
                    path.Rotate(tmpAngle)


                    path.Name = PS.Name

                    pastelist.Add(path)
                    cutlist.Add(_graphicsList(i))
                    _graphicsList.RemoveAt(i)
                    _graphicsList.Insert(i, path)

                End If
            Next

            If pastelist.Count > 0 Then
                Dim cmds As New List(Of Command.ICommand)
                cmds.Add(New Command.CutCommand(_graphicsList, cutlist))
                cmds.Add(New Command.PasteCommand(_graphicsList, pastelist))
                DrawObject.undoRedo.AddCommand(New Command.Commands(cmds))
            End If

        End Sub

        Public Sub CreatePolygon(PolygonSize As Single, NumberOfSides As Integer)
            Dim PC As New List(Of PathCommands)
            PC = CreatePolygonPath(PolygonSize, NumberOfSides)
            Dim DP As New DrawPath(PC)
            DP.EditPoints = False
            DP.setSelectionRectangle()
            Add(DP)
        End Sub

        Public Function CreatePolygonPath(PolygonSize As Single, NumberOfSides As Integer) As List(Of PathCommands)
            Dim polygonCommands As New List(Of PathCommands)

            ' Center of the polygon
            Dim centerX As Single = PolygonSize / 2
            Dim centerY As Single = PolygonSize / 2

            ' Radius of the polygon
            Dim radius As Single = PolygonSize / 2

            ' Calculate each corner point and create the PathCommands
            For i As Integer = 0 To NumberOfSides - 1
                Dim angle As Double = i * 2 * Math.PI / NumberOfSides
                Dim x As Single = CSng(centerX + radius * Math.Cos(angle))
                Dim y As Single = CSng(centerY + radius * Math.Sin(angle))

                ' If this is the first point, we move to it. Otherwise, we draw a line to it.
                Dim commandType As Char = If(i = 0, "M"c, "L"c)

                ' Add the PathCommand to our list
                polygonCommands.Add(New PathCommands(New PointF(x, y), Nothing, Nothing, commandType))
            Next

            ' Close the polygon path by using 'Z' as the command type
            polygonCommands.Add(New PathCommands(polygonCommands(0).P, Nothing, Nothing, "Z"c))

            Return polygonCommands
        End Function


        Public Sub CreateGear(GearType As Integer, numTeeth As Integer, outerDiameter As Single, boreDiameter As Single, pitchDepth As Single, PressureAngle As Single)
            Dim PC As New List(Of PathCommands)

            Select Case GearType
                Case 0 'Simple
                    PC = CreateSimpleGearPath(numTeeth, outerDiameter, boreDiameter, pitchDepth)
                Case 1 'Involute
                    PC = CreateInvoluteGearPath(numTeeth, outerDiameter, boreDiameter, pitchDepth, PressureAngle)
                Case 2 'Sprocket

            End Select


            Dim DP As New DrawPath(PC)
            DP.setSelectionRectangle()

            Add(DP)
        End Sub

        '    numTeeth: The number Of teeth On the gear. This parameter determines the number Of points that will be generated For the gear shape.
        '    pitchDiameter: The diameter Of the pitch circle, which Is the circle that touches the tip Of the gear teeth. This parameter determines the overall size Of the gear.
        '    Moduleo: The Module of the gear, which Is the ratio Of the reference diameter (usually the pitch diameter) To the number Of teeth. This parameter determines the size Of the teeth Of the gear.
        '    pressureAngle: The angle between the line tangent To the pitch circle at the point Of contact between two gears And the line Of centers. This parameter determines the shape Of the teeth Of the gear.
        '    addendum: The height Of the tooth above the pitch circle. This parameter determines the height Of the teeth Of the gear.
        '    dedendum: The depth Of the tooth below the pitch circle. This parameter determines the depth Of the valley between the teeth Of the gear.

        Public Function CreateSimpleGearPath(numTeeth As Integer, outerDiameter As Single, boreDiameter As Single, pitchDepth As Single) As List(Of PathCommands)
            Dim pathCommands As New List(Of PathCommands)

            ' --- Ensure Outer Diameter is Exact ---
            Dim outerRadius As Single = outerDiameter / 2 ' Directly define the outer radius
            Dim modul As Single = outerDiameter / numTeeth ' Define module based on outer diameter
            Dim addendum As Single = modul ' Standard addendum
            Dim dedendum As Single = 1.25 * modul ' Standard dedendum

            ' Use pitchDepth to adjust inner radius depth
            Dim innerRadius As Single = outerRadius - pitchDepth
            Dim boreRadius As Single = boreDiameter / 2

            ' Center of gear (assuming 0,0)
            Dim centerX As Single = 0
            Dim centerY As Single = 0

            ' Temporary storage to track the widest points
            Dim minX As Single = Single.MaxValue
            Dim maxX As Single = Single.MinValue
            Dim minY As Single = Single.MaxValue
            Dim maxY As Single = Single.MinValue

            ' --- Create Bore Hole Using Bezier Approximation ---
            Dim kappa As Single = 0.5522848F
            Dim rx As Single = boreRadius
            Dim ry As Single = boreRadius

            Dim dx As Single = kappa * rx
            Dim dy As Single = kappa * ry

            Dim borePoints As New List(Of (PointF, PointF, PointF, PointF)) From {
                 (New PointF(centerX + rx, centerY), New PointF(centerX + rx, centerY - dy), New PointF(centerX + dx, centerY - ry), New PointF(centerX, centerY - ry)),
                 (New PointF(centerX, centerY - ry), New PointF(centerX - dx, centerY - ry), New PointF(centerX - rx, centerY - dy), New PointF(centerX - rx, centerY)),
                 (New PointF(centerX - rx, centerY), New PointF(centerX - rx, centerY + dy), New PointF(centerX - dx, centerY + ry), New PointF(centerX, centerY + ry)),
                 (New PointF(centerX, centerY + ry), New PointF(centerX + dx, centerY + ry), New PointF(centerX + rx, centerY + dy), New PointF(centerX + rx, centerY))
             }

            ' Start bore path
            pathCommands.Add(New PathCommands(borePoints(0).Item1, Nothing, Nothing, "M"c))

            ' Add bore curve segments
            For Each segment In borePoints
                pathCommands.Add(New PathCommands(segment.Item4, segment.Item2, segment.Item3, "C"c))
            Next

            ' Close the bore properly before moving to the gear
            pathCommands.Add(New PathCommands(borePoints(0).Item1, Nothing, Nothing, "L"c))
            pathCommands.Add(New PathCommands(borePoints(0).Item1, Nothing, Nothing, "Z"c))

            ' --- Move to the Correct Gear Start Position ---
            Dim firstAngle As Single = -Math.PI / 2 ' Start at 12 o’clock
            Dim startX As Single = centerX + outerRadius * Math.Cos(firstAngle)
            Dim startY As Single = centerY + outerRadius * Math.Sin(firstAngle)

            ' Move to gear start point
            pathCommands.Add(New PathCommands(New PointF(startX, startY), Nothing, Nothing, "M"c))

            ' Track the first gear point for proper closure
            Dim firstGearPoint As New PointF(startX, startY)

            ' --- Generate Gear Teeth ---
            Dim angleStep As Single = (2 * Math.PI) / numTeeth
            Dim gearPoints As New List(Of PointF)

            For i As Integer = 0 To numTeeth - 1
                Dim angle As Single = firstAngle + (i * angleStep)

                ' Tooth tip (outer radius)
                Dim toothX As Single = centerX + outerRadius * Math.Cos(angle)
                Dim toothY As Single = centerY + outerRadius * Math.Sin(angle)
                gearPoints.Add(New PointF(toothX, toothY))

                ' Update min/max for scaling correction
                minX = Math.Min(minX, toothX)
                maxX = Math.Max(maxX, toothX)
                minY = Math.Min(minY, toothY)
                maxY = Math.Max(maxY, toothY)

                ' Tooth valley (inner radius, adjusted with pitchDepth)
                Dim valleyAngle As Single = angle + (angleStep / 2)
                Dim valleyX As Single = centerX + innerRadius * Math.Cos(valleyAngle)
                Dim valleyY As Single = centerY + innerRadius * Math.Sin(valleyAngle)
                gearPoints.Add(New PointF(valleyX, valleyY))

                ' Update min/max for scaling correction
                minX = Math.Min(minX, valleyX)
                maxX = Math.Max(maxX, valleyX)
                minY = Math.Min(minY, valleyY)
                maxY = Math.Max(maxY, valleyY)
            Next

            ' --- Calculate Scaling Factor ---
            Dim actualWidth As Single = maxX - minX
            Dim actualHeight As Single = maxY - minY
            Dim actualDiameter As Single = Math.Max(actualWidth, actualHeight)
            Dim scaleFactor As Single = outerDiameter / actualDiameter

            ' --- Apply Scaling Correction ---
            Dim correctedGearPoints As New List(Of PointF)
            For Each pt In gearPoints
                Dim correctedX As Single = centerX + (pt.X - centerX) * scaleFactor
                Dim correctedY As Single = centerY + (pt.Y - centerY) * scaleFactor
                correctedGearPoints.Add(New PointF(correctedX, correctedY))
            Next

            ' --- Apply Corrected Points to Path ---
            pathCommands.Add(New PathCommands(correctedGearPoints(0), Nothing, Nothing, "M"c))
            For Each pt In correctedGearPoints.Skip(1)
                pathCommands.Add(New PathCommands(pt, Nothing, Nothing, "L"c))
            Next

            ' --- Close the Gear Path Properly ---
            pathCommands.Add(New PathCommands(correctedGearPoints(0), Nothing, Nothing, "L"c))
            pathCommands.Add(New PathCommands(correctedGearPoints(0), Nothing, Nothing, "Z"c))

            Return pathCommands
        End Function


        Public Function CreateInvoluteGearPath(numTeeth As Integer, outerDiameter As Single, boreDiameter As Single, pitchDepth As Single, pressureAngle As Single) As List(Of PathCommands)
            Dim pathCommands As New List(Of PathCommands)

            ' --- Define Gear Parameters ---
            Dim pitchRadius As Single = outerDiameter / 2
            Dim baseRadius As Single = pitchRadius * Math.Cos(pressureAngle * Math.PI / 180) ' Base circle for involute
            Dim addendum As Single = outerDiameter / numTeeth
            Dim dedendum As Single = 1.25 * addendum
            Dim outerRadius As Single = pitchRadius + addendum
            Dim boreRadius As Single = boreDiameter / 2

            ' Center of gear
            Dim centerX As Single = 0
            Dim centerY As Single = 0

            ' --- Create Bore Hole ---
            Dim kappa As Single = 0.5522848F
            Dim rx As Single = boreRadius
            Dim ry As Single = boreRadius
            Dim dx As Single = kappa * rx
            Dim dy As Single = kappa * ry

            Dim borePoints As New List(Of (PointF, PointF, PointF, PointF)) From {
                 (New PointF(centerX + rx, centerY), New PointF(centerX + rx, centerY - dy), New PointF(centerX + dx, centerY - ry), New PointF(centerX, centerY - ry)),
                 (New PointF(centerX, centerY - ry), New PointF(centerX - dx, centerY - ry), New PointF(centerX - rx, centerY - dy), New PointF(centerX - rx, centerY)),
                 (New PointF(centerX - rx, centerY), New PointF(centerX - rx, centerY + dy), New PointF(centerX - dx, centerY + ry), New PointF(centerX, centerY + ry)),
                 (New PointF(centerX, centerY + ry), New PointF(centerX + dx, centerY + ry), New PointF(centerX + rx, centerY + dy), New PointF(centerX + rx, centerY))
             }

            ' Start bore path
            pathCommands.Add(New PathCommands(borePoints(0).Item1, Nothing, Nothing, "M"c))
            For Each segment In borePoints
                pathCommands.Add(New PathCommands(segment.Item4, segment.Item2, segment.Item3, "C"c))
            Next
            pathCommands.Add(New PathCommands(borePoints(0).Item1, Nothing, Nothing, "Z"c))

            ' --- Generate Involute Gear Teeth with Correct Bézier Curves ---
            Dim angleStep As Single = (2 * Math.PI) / numTeeth
            Dim firstGearPoint As PointF = Nothing
            Dim isFirstPoint As Boolean = True

            For i As Integer = 0 To numTeeth - 1
                Dim toothAngle As Single = i * angleStep
                Dim involutePoints As New List(Of (PointF, PointF, PointF, PointF))

                ' Generate involute curve using Bézier approximation
                Dim t As Single = 0
                Dim stepSize As Single = 0.05 ' Adjust for smoothness
                Dim startPoint As PointF = Nothing

                While True
                    Dim x As Single = baseRadius * (Math.Cos(t) + t * Math.Sin(t))
                    Dim y As Single = baseRadius * (Math.Sin(t) - t * Math.Cos(t))
                    Dim distance As Single = Math.Sqrt(x * x + y * y)

                    If distance >= outerRadius Then Exit While ' Stop at tooth tip

                    ' Compute control points for Bézier curve
                    Dim p0 As New PointF(x, y) ' Start of involute
                    Dim p3 As New PointF(x * 1.02, y * 1.02) ' Tooth tip

                    ' Calculate Bézier control points
                    Dim tangentX As Single = -y ' Tangent direction
                    Dim tangentY As Single = x
                    Dim p1 As New PointF(x + tangentX * 0.3, y + tangentY * 0.3) ' Control point B1
                    Dim p2 As New PointF(x + tangentX * 0.6, y + tangentY * 0.6) ' Control point B2

                    ' Rotate each point to the correct tooth angle
                    Dim rotatedP0 As New PointF(centerX + (p0.X * Math.Cos(toothAngle) - p0.Y * Math.Sin(toothAngle)), centerY + (p0.X * Math.Sin(toothAngle) + p0.Y * Math.Cos(toothAngle)))
                    Dim rotatedP1 As New PointF(centerX + (p1.X * Math.Cos(toothAngle) - p1.Y * Math.Sin(toothAngle)), centerY + (p1.X * Math.Sin(toothAngle) + p1.Y * Math.Cos(toothAngle)))
                    Dim rotatedP2 As New PointF(centerX + (p2.X * Math.Cos(toothAngle) - p2.Y * Math.Sin(toothAngle)), centerY + (p2.X * Math.Sin(toothAngle) + p2.Y * Math.Cos(toothAngle)))
                    Dim rotatedP3 As New PointF(centerX + (p3.X * Math.Cos(toothAngle) - p3.Y * Math.Sin(toothAngle)), centerY + (p3.X * Math.Sin(toothAngle) + p3.Y * Math.Cos(toothAngle)))

                    involutePoints.Add((rotatedP0, rotatedP1, rotatedP2, rotatedP3))

                    If startPoint = Nothing Then
                        startPoint = rotatedP0
                    End If

                    t += stepSize
                End While

                ' Store first point
                If isFirstPoint Then
                    firstGearPoint = startPoint
                    pathCommands.Add(New PathCommands(firstGearPoint, Nothing, Nothing, "M"c))
                    isFirstPoint = False
                End If

                ' Add Bézier curves to path
                For Each segment In involutePoints
                    pathCommands.Add(New PathCommands(segment.Item4, segment.Item2, segment.Item3, "C"c))
                Next
            Next

            ' Close the gear path
            pathCommands.Add(New PathCommands(firstGearPoint, Nothing, Nothing, "Z"c))

            Return pathCommands
        End Function


        Public Sub RepeatShape(DR As DrawObject, ShapeCount As Integer, SizeIncrease As Single, ShapeType As String)
            Dim RectList As New List(Of RectangleF)
            RectList = GetRepeatedRectangles(DR.Rectangle, ShapeCount, SizeIncrease)

            Select Case ShapeType
                Case "Draw.DrawRectangle"
                    Dim DP As New DrawRectangle
                Case "Draw.DrawEllipse"
                    Dim DP As New DrawEllipse
                Case Else
                    'Cannot be a draw Path, line or polygon
            End Select

            For Each rect In RectList
                Select Case ShapeType
                    Case "Draw.DrawRectangle"
                        Dim DP As New DrawRectangle
                        DP.setRectangle(New RectangleF(rect.Location, rect.Size))
                        'DP.OriginalSelectionRect = New RectangleF(rect.Location, rect.Size)
                        Add(DP)
                    Case "Draw.DrawEllipse"
                        Dim DP As New DrawEllipse
                        DP.setRectangle(New RectangleF(rect.Location, rect.Size))
                        'DP.OriginalSelectionRect = New RectangleF(rect.Location, rect.Size)
                        Add(DP)
                    Case Else
                        'Cannot be a draw Path, line or polygon
                End Select

            Next

        End Sub


        Public Sub CreateFlangePath(centerX As Single, centerY As Single, outerDiameter As Single, boreDiameter As Single, boltHoleDiameter As Single, boltHoleCount As Integer, Optional boltCircleRadius As Single = -1)

            Dim flangePath As New List(Of PathCommands)

            Dim center As New PointF(centerX, centerY)

            ' If boltCircleRadius is negative, set it to the middle of bore and outer diameter
            If boltCircleRadius < 0 Then
                boltCircleRadius = (outerDiameter + boreDiameter) / 4
            End If

            ' Create the inner bore (center hole)
            flangePath.AddRange(CreateEllipsePath(center, boreDiameter / 2))

            ' Create bolt holes
            Dim angleIncrement As Double = (2 * Math.PI) / boltHoleCount
            Dim boltRadius As Single = boltHoleDiameter / 2

            For i As Integer = 0 To boltHoleCount - 1
                ' Calculate bolt hole position
                Dim angle As Double = i * angleIncrement
                Dim boltX As Single = centerX + boltCircleRadius * Math.Cos(angle)
                Dim boltY As Single = centerY + boltCircleRadius * Math.Sin(angle)

                ' Add bolt hole as an ellipse path
                flangePath.AddRange(CreateEllipsePath(New PointF(boltX, boltY), boltRadius))
            Next

            ' Create the outer flange perimeter
            flangePath.AddRange(CreateEllipsePath(center, outerDiameter / 2))

            Dim pcs As New DrawPath(flangePath)
            Add(pcs)

        End Sub

        'Used for CreateFlangePath()
        Public Function CreateEllipsePath(center As PointF, radius As Single) As List(Of PathCommands)
            Dim path As New List(Of PathCommands)

            ' Approximate the ellipse using four quadratic Bézier curves
            Dim kappa As Single = 0.552284749831 ' Factor for Bézier approximation of a circle

            Dim rx As Single = radius
            Dim ry As Single = radius

            ' Key points around the ellipse
            Dim p0 As New PointF(center.X - rx, center.Y) ' Left
            Dim p1 As New PointF(center.X, center.Y - ry) ' Top
            Dim p2 As New PointF(center.X + rx, center.Y) ' Right
            Dim p3 As New PointF(center.X, center.Y + ry) ' Bottom

            ' Control points
            Dim cp1 As New PointF(center.X - rx, center.Y - kappa * ry)
            Dim cp2 As New PointF(center.X - kappa * rx, center.Y - ry)
            Dim cp3 As New PointF(center.X + kappa * rx, center.Y - ry)
            Dim cp4 As New PointF(center.X + rx, center.Y - kappa * ry)
            Dim cp5 As New PointF(center.X + rx, center.Y + kappa * ry)
            Dim cp6 As New PointF(center.X + kappa * rx, center.Y + ry)
            Dim cp7 As New PointF(center.X - kappa * rx, center.Y + ry)
            Dim cp8 As New PointF(center.X - rx, center.Y + kappa * ry)

            ' Start the path
            path.Add(New PathCommands(p0, Nothing, Nothing, "M"c))

            ' Quadratic Bézier curves forming the ellipse
            path.Add(New PathCommands(p1, cp1, cp2, "C"c))
            path.Add(New PathCommands(p2, cp3, cp4, "C"c))
            path.Add(New PathCommands(p3, cp5, cp6, "C"c))
            path.Add(New PathCommands(p0, cp7, cp8, "C"c))

            ' Close the path
            path.Add(New PathCommands(p0, Nothing, Nothing, "Z"c))

            Return path
        End Function


        Public Sub RingShape(DR As DrawObject, ShapeCount As Integer, SizeIncrease As Single, ShapeType As String, ShapeSize As SizeF) 'DR = shape to ring around, shapecount is how many shapes, sizeincrease is how far new shapes will be from center and shaoetype had the option of rectangle or ellipse but e just want ellipse for this.

            Dim RectList As New List(Of RectangleF)
            RectList = CreateCircularRectangles(DR.Rectangle, ShapeCount, SizeIncrease, ShapeSize)

            Select Case ShapeType
                Case "Draw.DrawRectangle"
                    Dim DP As New DrawRectangle
                Case "Draw.DrawEllipse"
                    Dim DP As New DrawEllipse
                Case Else
                    'Cannot be a draw Path, line or polygon
            End Select


            For Each rect In RectList
                Select Case ShapeType
                    Case "Draw.DrawRectangle"
                        Dim DP As New DrawRectangle
                        DP.setRectangle(New RectangleF(rect.Location, rect.Size))
                        ' DP.OriginalSelectionRect = New RectangleF(rect.Location, rect.Size)
                        Add(DP)
                    Case "Draw.DrawEllipse"
                        Dim DP As New DrawEllipse
                        DP.setRectangle(New RectangleF(rect.Location, rect.Size))
                        ' DP.OriginalSelectionRect = New RectangleF(rect.Location, rect.Size)
                        Add(DP)
                    Case Else
                        'Cannot be a draw Path, line or polygon
                End Select

            Next

        End Sub

        'Used for RingShape()
        Public Function CreateCircularRectangles(centerRectangle As RectangleF, count As Integer, distance As Single, shapeSize As SizeF, Optional offset As PointF = Nothing) As List(Of RectangleF)
            If offset = Nothing Then
                offset = New PointF(0, 0)
            End If
            Dim rectangles As New List(Of RectangleF)()
            Dim angleIncrement As Double = (2 * Math.PI) / count
            For i As Integer = 0 To count - 1
                Dim x As Single = (centerRectangle.X + (centerRectangle.Width / 2)) + offset.X + (distance * Math.Cos(i * angleIncrement)) - shapeSize.Width / 2
                Dim y As Single = (centerRectangle.Y + (centerRectangle.Height / 2)) + offset.Y + (distance * Math.Sin(i * angleIncrement)) - shapeSize.Height / 2
                Dim newRectangle As New RectangleF(x, y, shapeSize.Width, shapeSize.Height)
                rectangles.Add(newRectangle)
            Next
            Return rectangles
        End Function

        Public Function EllipseToBezier(rectF As RectangleF, numPoints As Integer) As GraphicsPath
            Dim path As New GraphicsPath
            Dim stepSize As Single = CSng(2.0 * Math.PI / numPoints)
            Dim theta As Single = 0
            Dim x As Single, y As Single
            For i As Integer = 0 To numPoints - 1
                x = rectF.X + rectF.Width / 2 + (rectF.Width / 2) * Math.Cos(theta)
                y = rectF.Y + rectF.Height / 2 + (rectF.Height / 2) * Math.Sin(theta)
                If i = 0 Then
                    path.StartFigure()
                Else
                    path.AddLine(x, y, x, y)
                End If
                theta += stepSize
            Next
            path.CloseFigure()
            Return path
        End Function

        Public Function RoundedRectangleToBezier(rectF As RectangleF, radius As Integer) As GraphicsPath
            Dim path As New GraphicsPath
            path.StartFigure()
            path.AddArc(rectF.X, rectF.Y, radius, radius, 180, 90)
            path.AddLine(rectF.X + radius, rectF.Y, rectF.Right - radius, rectF.Y)
            path.AddArc(rectF.Right - radius - rectF.X, rectF.Y, radius, radius, 270, 90)
            path.AddLine(rectF.Right, rectF.Y + radius, rectF.Right, rectF.Bottom - radius)
            path.AddArc(rectF.Right - radius - rectF.X, rectF.Bottom - radius - rectF.Y, radius, radius, 0, 90)
            path.AddLine(rectF.Right - radius, rectF.Bottom, rectF.X + radius, rectF.Bottom)
            path.AddArc(rectF.X, rectF.Bottom - radius - rectF.Y, radius, radius, 90, 90)
            path.AddLine(rectF.X, rectF.Bottom - radius, rectF.X, rectF.Y + radius)
            path.CloseFigure()
            Return path
        End Function


        Public Sub ShapeCopier(Direction As Integer, SpaceFromEdge As Single) '0 = Up, 1 = Right, 2 = Down, 3 = Left

            Try

                Dim DOList As New List(Of DrawObject)

                For Each shape As DrawObject In _graphicsList
                    If shape.Selected = True Then

                        Dim OrigRect As RectangleF = shape.Rectangle

                        Dim NewDO As DrawObject = CType(shape.Clone, DrawObject)

                        Dim NewRect As RectangleF

                        Select Case Direction
                            Case 0 'Up
                                NewRect = New RectangleF(OrigRect.X, OrigRect.Y - (OrigRect.Height + SpaceFromEdge), OrigRect.Width, OrigRect.Height)
                            Case 1 'Right
                                NewRect = New RectangleF(OrigRect.X + (OrigRect.Width + SpaceFromEdge), OrigRect.Y, OrigRect.Width, OrigRect.Height)
                            Case 2 'Down
                                NewRect = New RectangleF(OrigRect.X, OrigRect.Y + (OrigRect.Height + SpaceFromEdge), OrigRect.Width, OrigRect.Height)
                            Case 3 'Left
                                NewRect = New RectangleF(OrigRect.X - (OrigRect.Width + SpaceFromEdge), OrigRect.Y, OrigRect.Width, OrigRect.Height)
                        End Select

                        NewDO.setRectangle(NewRect)

                        DOList.Add(NewDO)
                        shape.Selected = False

                    End If
                Next

                For Each DOval As DrawObject In DOList
                    Add(DOval)
                Next

            Catch ex As Exception
                ' Log the exception for debugging purposes
                MessageBox.Show($"Error in ShapeCopier: {ex.ToString()}")

                ' Optionally, inform the user that something went wrong
                MessageBox.Show("An unexpected error occurred while copying the shape.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Sub


        Function GetRepeatedRectangles(originalRectangle As RectangleF, repeatCount As Integer, sizeIncrease As Single) As List(Of RectangleF)
            Dim repeatedRectangles As New List(Of RectangleF)

            For i As Integer = 0 To repeatCount - 1
                Dim repeatedRectangle As New RectangleF(originalRectangle.X - (i + 1) * sizeIncrease, originalRectangle.Y - (i + 1) * sizeIncrease, originalRectangle.Width + (i + 1) * sizeIncrease * 2, originalRectangle.Height + (i + 1) * sizeIncrease * 2)
                repeatedRectangles.Add(repeatedRectangle)
            Next

            Return repeatedRectangles
        End Function

        Public Sub UpdateColours(ColOutline As Color, ColDragHandle As Color, ColPointhandle As Color, ColSelectedhandle As Color, ColB1 As Color, ColB2 As Color, ColRS As Color, ColSR As Color, CH As Color, ColLI As Color, ColLO As Color, Optional ByVal DrawObj As DrawObject = Nothing)
            Try
                Try
                    If DrawObj IsNot Nothing Then
                        DrawObj.JobType = JobType
                    End If
                Catch
                    DrawObj.JobType = DrawObject.JobTypes.Laser
                End Try


                If DrawObj IsNot Nothing Then
                    'Doing specific DrawObject likely as added new
                    DrawObj.ColShapeOutline = ColOutline
                    DrawObj.ColDragHandles = ColDragHandle
                    DrawObj.ColPointHandle = ColPointhandle
                    DrawObj.ColSelectedPointhandle = ColSelectedhandle
                    DrawObj.ColB1Handle = ColB1
                    DrawObj.ColB2Handle = ColB1
                    DrawObj.ColRotateStart = ColRS
                    DrawObj.ColSelectionRectangle = ColSR
                    DrawObj.ColCrosshair = ColCrosshair
                    DrawObj.ColLeadIn = ColLI
                    DrawObj.ColLeadOut = ColLO
                Else
                    'Likely called from setting change or start up, all shapes changing, record colors
                    Me.ColShapeOutline = ColOutline
                    Me.ColDragHandles = ColDragHandle
                    Me.ColPointHandle = ColPointhandle
                    Me.ColSelectedPointhandle = ColSelectedhandle
                    Me.ColB1Handle = ColB1
                    Me.ColB2Handle = ColB1
                    Me.ColRotateStart = ColRS
                    Me.ColSelectionRectangle = ColSR
                    Me.ColCrosshair = CH

                    For Each o As DrawObject In _graphicsList
                        o.ColShapeOutline = ColOutline
                        o.ColDragHandles = ColDragHandle
                        o.ColPointHandle = ColPointhandle
                        o.ColSelectedPointhandle = ColSelectedhandle
                        o.ColB1Handle = ColB1
                        o.ColB2Handle = ColB1
                        o.ColRotateStart = ColRS
                        o.ColSelectionRectangle = ColSR
                        o.ColCrosshair = CH
                        o.ColLeadIn = ColLI
                        o.ColLeadOut = ColLO
                    Next

                End If

            Catch
            End Try

        End Sub

#Region "NewMergeCode"

        '===============================================================
        '========== COMPLETE FIXED MERGING CODE WITH IMPROVED UNION ==
        '===============================================================

        ' Main entry point for merging operations
        Public Sub MergeSelected(MergeType As Integer)
            Dim ObjList As New List(Of DrawObject)
            Dim indicesToRemove As New List(Of Integer)

            ' Collect all selected objects and their indices
            For i As Integer = 0 To _graphicsList.Count - 1
                Dim o As DrawObject = CType(_graphicsList(i), DrawObject)
                If o.Selected Then
                    ObjList.Add(o.Clone)
                    indicesToRemove.Add(i)
                End If
            Next

            If ObjList.Count < 2 Then
                MessageBox.Show("Please select at least two objects to merge.", "Insufficient Selection", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' Perform the merge operation FIRST
            Dim Pc As List(Of PathCommands)

            Select Case MergeType
                Case 0 ' Union
                    Pc = MergePathsUnion(ObjList)
                Case 1 ' Punch Out
                    Pc = MergePathsPunchOut(ObjList)
                Case 2 ' Intersect
                    Pc = MergePathsIntersect(ObjList)
                Case Else
                    MessageBox.Show("Unknown merge type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
            End Select

            ' Check if we got a valid result
            If Pc Is Nothing OrElse Pc.Count = 0 Then
                MessageBox.Show("Merge operation resulted in an empty path.", "Merge Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Remove the original shapes (in reverse order to maintain indices)
            For i As Integer = indicesToRemove.Count - 1 To 0 Step -1
                _graphicsList.RemoveAt(indicesToRemove(i))
            Next

            ' Add the merged result back into the graphics list
            Dim newObject As New DrawPath(Pc)
            newObject.Selected = True ' Select the new merged object
            Add(newObject)
        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        Private Function IsClockwise(path As BezierPath) As Boolean
            Dim area# = 0
            For Each s In path.Segments
                area += (s.EndPoint.X - s.StartPoint.X) * (s.EndPoint.Y + s.StartPoint.Y)
            Next
            Return area > 0
        End Function

        ''' <summary>
        ''' Union of *all* shapes in the list (no limit on count)
        ''' </summary>

        Public Function MergePathsUnion(drawObjects As List(Of DrawObject)) As List(Of PathCommands)
            If drawObjects.Count = 0 Then Return New List(Of PathCommands)
            If drawObjects.Count = 1 Then Return drawObjects(0).PathCommands

            ' Start with the first object's path
            Dim resultPath As List(Of PathCommands) = New List(Of PathCommands)(drawObjects(0).PathCommands)

            ' Union with each additional path
            For i As Integer = 1 To drawObjects.Count - 1
                resultPath = UnionTwoPaths(resultPath, drawObjects(i).PathCommands)
            Next

            Return resultPath
        End Function

        ''' <summary>
        ''' Performs a Punch-Out (subtract all subsequent shapes from the first).
        ''' Works with any path type and preserves "around" edge traversal.
        ''' </summary>
        Public Function MergePathsPunchOut(drawObjects As List(Of DrawObject)) As List(Of PathCommands)
            If drawObjects.Count = 0 Then Return New List(Of PathCommands)
            If drawObjects.Count = 1 Then Return drawObjects(0).PathCommands

            '--- 1. Convert all shapes to Bezier paths
            Dim allBeziers As New List(Of List(Of BezierPath))
            For Each o In drawObjects
                allBeziers.Add(ConvertToBezierPaths(o.PathCommands))
            Next

            Dim basePaths As List(Of BezierPath) = allBeziers(0)

            '--- 2. Collect all intersections between base and each puncher
            Dim intersections As New List(Of PathIntersection)
            For i = 1 To allBeziers.Count - 1
                intersections.AddRange(FindAllPathIntersections(basePaths, allBeziers(i)))
            Next
            intersections = CleanupIntersections(intersections)

            '--- 3. Split base path at intersections
            Dim splitBase As List(Of BezierPath) = SplitPathAtIntersections(basePaths, intersections, True)

            '--- 4. Flag which segments are inside a puncher (to remove)
            For Each seg In splitBase.SelectMany(Function(p) p.Segments)
                Dim mid As PointF = seg.PointAt(0.5)
                Dim inside As Boolean = False
                For i = 1 To allBeziers.Count - 1
                    For Each puncher In allBeziers(i)
                        If puncher.IsClosed AndAlso ContainsPoint(puncher, mid) Then
                            inside = True
                            Exit For
                        End If
                    Next
                    If inside Then Exit For
                Next
                seg.Keep = Not inside
            Next

            '--- 5. Walk remaining boundary
            Dim keptSegs As List(Of BezierSegment) =
        splitBase.SelectMany(Function(p) p.Segments).Where(Function(s) s.Keep).ToList()
            Dim walked As List(Of BezierPath) = WalkOuterBoundary(keptSegs)

            '--- 6. Ensure correct winding: outer CCW, holes CW
            For Each path In walked
                Dim area = CalculatePathArea(path)
                If area < 0 Then
                    path.Segments.Reverse()
                End If
            Next

            '--- 7. Convert back to PathCommands
            Return ConvertToPathCommandsImproved(walked)
        End Function


        ''' <summary>
        ''' Performs an intersection operation on multiple paths
        ''' </summary>
        Public Function MergePathsIntersect(drawObjects As List(Of DrawObject)) As List(Of PathCommands)
        End Function

#Region “New Boundary-Tracing Core”

        ''' <summary>
        ''' Union: keep the OUTSIDE of all shapes
        ''' </summary>
        Private Function UnionTwoPaths(path1 As List(Of PathCommands), path2 As List(Of PathCommands)) As List(Of PathCommands)
            ' Convert to BezierPath for processing
            Dim bezierPath1 As List(Of BezierPath) = ConvertToBezierPaths(path1)
            Dim bezierPath2 As List(Of BezierPath) = ConvertToBezierPaths(path2)

            ' First check for no overlap case with improved bounding box test
            If Not PathsOverlap(bezierPath1, bezierPath2) Then
                ' No overlap - just combine the paths
                Return CombineDisjointPaths(path1, path2)
            End If

            ' Find all intersections between the paths
            Dim intersections As List(Of PathIntersection) = FindAllPathIntersections(bezierPath1, bezierPath2)

            ' If no intersections but bounding boxes overlap, one path may be inside the other
            If intersections.Count = 0 Then
                ' Check if one path contains the other completely
                If CompletelyContains(bezierPath1, bezierPath2) Then
                    ' Path1 contains path2 completely, so just return path1
                    Return path1
                ElseIf CompletelyContains(bezierPath2, bezierPath1) Then
                    ' Path2 contains path1 completely, so just return path2
                    Return path2
                End If

                ' Paths don't intersect and don't contain each other, just combine them
                Return CombineDisjointPaths(path1, path2)
            End If

            ' Clean up duplicate or very close intersections
            intersections = CleanupIntersections(intersections)

            ' Split both paths at intersection points
            Dim splitPath1 As List(Of BezierPath) = SplitPathAtIntersections(bezierPath1, intersections, True)
            Dim splitPath2 As List(Of BezierPath) = SplitPathAtIntersections(bezierPath2, intersections, False)

            ' Mark segments based on containment - FIXED LOGIC FOR UNION
            ' For union, we keep segments that are OUTSIDE the other path
            MarkPathSegmentsImproved(splitPath1, splitPath2, False)
            MarkPathSegmentsImproved(splitPath2, splitPath1, False)

            ' Combine the segments that should be kept
            Dim resultSegments As List(Of BezierSegment) = CombineSegmentsForOperation(splitPath1, splitPath2)

            ' Reconstruct continuous paths from segments
            Dim resultPaths As List(Of BezierPath) = ReconstructPathsImproved(resultSegments, intersections)

            ' Convert back to PathCommands with properly closed paths
            Return ConvertToPathCommandsImproved(resultPaths)
        End Function

        ''' <summary>
        ''' Extracts only the outer boundary from a set of paths (removes holes)
        ''' </summary>
        Private Function GetOuterBoundaryOnly(paths As List(Of BezierPath)) As List(Of PathCommands)
            If paths.Count = 0 Then Return New List(Of PathCommands)

            ' Find the path with the largest area (this is the outer boundary)
            Dim outerPath As BezierPath = Nothing
            Dim maxArea As Double = 0

            For Each path In paths
                If Not path.IsClosed Then Continue For

                Dim area = CalculatePathArea(path)

                ' The outer boundary has the largest positive area
                ' (counter-clockwise winding)
                If Math.Abs(area) > maxArea Then
                    maxArea = Math.Abs(area)
                    outerPath = path
                End If
            Next

            If outerPath Is Nothing Then
                ' If no closed path found, return all paths
                Return ConvertToPathCommandsImproved(paths)
            End If

            ' Return only the outer boundary
            Dim result As New List(Of BezierPath)
            result.Add(outerPath)

            Return ConvertToPathCommandsImproved(result)
        End Function

        ''' <summary>
        ''' Calculates the signed area of a closed path using the shoelace formula
        ''' Positive area = counter-clockwise, Negative area = clockwise
        ''' </summary>
        Private Function CalculatePathArea(path As BezierPath) As Double
            If Not path.IsClosed OrElse path.Segments.Count < 3 Then
                Return 0
            End If

            Dim area As Double = 0

            For Each segment In path.Segments
                ' For accurate area calculation of curves, we need to sample points
                Dim steps As Integer = If(segment.SegmentType = "C"c, 10, 1)

                For i As Integer = 0 To steps - 1
                    Dim t1 As Double = i / steps
                    Dim t2 As Double = (i + 1) / steps

                    Dim p1 = segment.PointAt(t1)
                    Dim p2 = segment.PointAt(t2)

                    ' Shoelace formula: area += (x1 * y2 - x2 * y1)
                    area += (p1.X * p2.Y - p2.X * p1.Y)
                Next
            Next

            Return area / 2.0
        End Function

#End Region


#Region "Path Conversion"

        ''' <summary>
        ''' Converts PathCommands list to BezierPath for processing - COMPLETE VERSION
        ''' </summary>
        Private Function ConvertToBezierPaths(pathCommands As List(Of PathCommands)) As List(Of BezierPath)
            Dim result As New List(Of BezierPath)()

            If pathCommands.Count = 0 Then
                MessageBox.Show("ConvertToBezierPaths: No path commands provided")
                Return result
            End If

            MessageBox.Show($"ConvertToBezierPaths: Processing {pathCommands.Count} commands")

            Dim currentPath As New BezierPath()
            Dim startPoint As PointF = Nothing
            Dim currentPoint As PointF = Nothing
            Dim pathIndex As Integer = 0

            For i As Integer = 0 To pathCommands.Count - 1
                Dim cmd As PathCommands = pathCommands(i)
                Dim cmdType As Char = GetCommandChar(cmd) ' Use the existing helper function

                Select Case cmdType
                    Case "M"c ' MoveTo - Start a new subpath
                        If currentPath.Segments.Count > 0 Then
                            ' Add the current path to the result if it has segments
                            currentPath.PathIndex = pathIndex
                            result.Add(currentPath)
                            currentPath = New BezierPath()
                            pathIndex += 1
                        End If

                        startPoint = cmd.P
                        currentPoint = cmd.P

                    Case "L"c ' LineTo
                        If currentPoint <> Nothing Then
                            Dim segment As New BezierSegment(currentPoint, cmd.P, Nothing, Nothing, "L"c)
                            segment.PathIndex = pathIndex
                            segment.SegmentIndex = currentPath.Segments.Count
                            currentPath.Segments.Add(segment)
                            currentPoint = cmd.P
                        End If

                    Case "C"c ' Cubic Bezier
                        If currentPoint <> Nothing Then
                            Dim segment As New BezierSegment(currentPoint, cmd.P, cmd.b1, cmd.b2, "C"c)
                            segment.PathIndex = pathIndex
                            segment.SegmentIndex = currentPath.Segments.Count
                            currentPath.Segments.Add(segment)
                            currentPoint = cmd.P
                        End If

                    Case "Z"c ' ClosePath
                        If currentPath.Segments.Count > 0 Then
                            currentPath.IsClosed = True
                            ' Add line back to start if needed
                            If PointDistance(currentPoint, startPoint) > 0.001 Then
                                Dim segment As New BezierSegment(currentPoint, startPoint, Nothing, Nothing, "L"c)
                                segment.PathIndex = pathIndex
                                segment.SegmentIndex = currentPath.Segments.Count
                                currentPath.Segments.Add(segment)
                            End If
                        End If
                End Select
            Next

            ' Add the last path if it has segments
            If currentPath.Segments.Count > 0 Then
                currentPath.PathIndex = pathIndex
                result.Add(currentPath)
            End If

            MessageBox.Show($"ConvertToBezierPaths: Created {result.Count} bezier paths")
            Return result
        End Function
#End Region

#Region "Helper Methods"


        ''' <summary>
        ''' Gets the bounding box for a segment
        ''' </summary>
        Private Function GetSegmentBoundingBox(segment As BezierSegment) As RectangleF
            Dim minX As Single = Math.Min(segment.StartPoint.X, segment.EndPoint.X)
            Dim minY As Single = Math.Min(segment.StartPoint.Y, segment.EndPoint.Y)
            Dim maxX As Single = Math.Max(segment.StartPoint.X, segment.EndPoint.X)
            Dim maxY As Single = Math.Max(segment.StartPoint.Y, segment.EndPoint.Y)

            If segment.SegmentType = "C"c Then
                ' Include control points for Bézier curves
                minX = Math.Min(minX, Math.Min(segment.Control1.X, segment.Control2.X))
                minY = Math.Min(minY, Math.Min(segment.Control1.Y, segment.Control2.Y))
                maxX = Math.Max(maxX, Math.Max(segment.Control1.X, segment.Control2.X))
                maxY = Math.Max(maxY, Math.Max(segment.Control1.Y, segment.Control2.Y))
            End If

            Return New RectangleF(minX, minY, maxX - minX, maxY - minY)
        End Function

        ''' <summary>
        ''' Checks if two sets of paths overlap
        ''' </summary>
        Private Function PathsOverlap(paths1 As List(Of BezierPath), paths2 As List(Of BezierPath)) As Boolean
            For Each path1 In paths1
                Dim bounds1 As RectangleF = path1.GetBoundingBox()

                For Each path2 In paths2
                    Dim bounds2 As RectangleF = path2.GetBoundingBox()

                    If bounds1.IntersectsWith(bounds2) Then
                        Return True
                    End If
                Next
            Next

            Return False
        End Function

        ''' <summary>
        ''' Checks if the first set of paths completely contains the second set
        ''' </summary>
        Private Function CompletelyContains(containerPaths As List(Of BezierPath), containedPaths As List(Of BezierPath)) As Boolean
            ' Check each path in the contained set
            For Each containedPath In containedPaths
                If Not containedPath.IsClosed Then Continue For

                ' Take several sample points from this path
                Dim sampleCount As Integer = Math.Max(4, containedPath.Segments.Count)
                Dim allPointsContained As Boolean = True

                ' Use segment midpoints as sample points
                For Each segment In containedPath.Segments
                    Dim samplePoint As PointF = segment.PointAt(0.5)
                    Dim isContained As Boolean = False

                    ' Check if any container path contains this point
                    For Each containerPath In containerPaths
                        If containerPath.IsClosed AndAlso ContainsPoint(containerPath, samplePoint) Then
                            isContained = True
                            Exit For
                        End If
                    Next

                    If Not isContained Then
                        allPointsContained = False
                        Exit For
                    End If
                Next

                If Not allPointsContained Then
                    Return False
                End If
            Next

            Return True
        End Function

        ''' <summary>
        ''' Combines two sets of paths that don't overlap
        ''' </summary>
        Private Function CombineDisjointPaths(path1 As List(Of PathCommands), path2 As List(Of PathCommands)) As List(Of PathCommands)
            Dim result As New List(Of PathCommands)(path1)
            result.AddRange(path2)
            Return result
        End Function

        ''' <summary>
        ''' Improved version of MarkPathSegments with better inside/outside detection
        ''' </summary>
        Private Sub MarkPathSegmentsImproved(pathsToMark As List(Of BezierPath), otherPaths As List(Of BezierPath), keepInsideSegments As Boolean)
            ' For each path to mark
            For Each path In pathsToMark
                ' For each segment
                For Each segment In path.Segments
                    ' Test multiple points along the segment for more accurate inside/outside detection
                    Dim testPoints As Integer = 3 ' Use multiple test points per segment
                    Dim insideCount As Integer = 0

                    For i As Integer = 1 To testPoints
                        Dim t As Double = i / (testPoints + 1) ' Evenly distribute test points along segment
                        Dim testPoint As PointF = segment.PointAt(t)

                        ' Check if this point is inside any other path
                        For Each otherPath In otherPaths
                            If otherPath.IsClosed AndAlso ContainsPoint(otherPath, testPoint) Then
                                insideCount += 1
                                Exit For
                            End If
                        Next
                    Next

                    ' Segment is considered inside if majority of test points are inside
                    segment.IsInside = (insideCount > testPoints / 2)

                    ' Set keep flag based on containment and operation
                    segment.Keep = If(keepInsideSegments, segment.IsInside, Not segment.IsInside)
                Next
            Next
        End Sub

        ''' <summary>
        ''' Improved version of ReconstructPaths with better path reconstruction logic
        ''' </summary>
        Private Function ReconstructPathsImproved(segments As List(Of BezierSegment), intersections As List(Of PathIntersection)) As List(Of BezierPath)
            Dim result As New List(Of BezierPath)()

            ' Exit early if no segments
            If segments.Count = 0 Then Return result

            ' Create a copy of segments we can modify
            Dim remainingSegments As New List(Of BezierSegment)(segments)

            ' Use a better endpoint tolerance for more accurate connection detection
            Const ENDPOINT_TOLERANCE As Double = 0.001

            ' Process until all segments are used
            While remainingSegments.Count > 0
                Dim currentPath As New BezierPath()

                ' Start with first available segment
                Dim currentSegment As BezierSegment = remainingSegments(0)
                remainingSegments.RemoveAt(0)
                currentPath.Segments.Add(currentSegment)

                Dim startPoint As PointF = currentSegment.StartPoint
                Dim currentEndPoint As PointF = currentSegment.EndPoint

                ' Loop safety counter
                Dim loopLimit As Integer = remainingSegments.Count + 10
                Dim loopCount As Integer = 0

                ' Flag indicating if the path is closed
                Dim pathClosed As Boolean = False

                ' Keep adding connected segments until path is closed or no more segments can be added
                While loopCount < loopLimit
                    loopCount += 1

                    ' Check if we can close the path
                    If PointDistance(currentEndPoint, startPoint) < ENDPOINT_TOLERANCE Then
                        pathClosed = True
                        Exit While
                    End If

                    ' Find next segment to add
                    Dim bestSegmentIndex As Integer = -1
                    Dim bestDistance As Double = ENDPOINT_TOLERANCE
                    Dim needToReverse As Boolean = False

                    For i As Integer = 0 To remainingSegments.Count - 1
                        Dim segment As BezierSegment = remainingSegments(i)

                        ' Check distance to start point of segment
                        Dim distToStart As Double = PointDistance(currentEndPoint, segment.StartPoint)
                        If distToStart < bestDistance Then
                            bestSegmentIndex = i
                            bestDistance = distToStart
                            needToReverse = False
                        End If

                        ' Check distance to end point of segment
                        Dim distToEnd As Double = PointDistance(currentEndPoint, segment.EndPoint)
                        If distToEnd < bestDistance Then
                            bestSegmentIndex = i
                            bestDistance = distToEnd
                            needToReverse = True
                        End If
                    Next

                    ' If found a connecting segment
                    If bestSegmentIndex >= 0 Then
                        Dim nextSegment As BezierSegment = remainingSegments(bestSegmentIndex)
                        remainingSegments.RemoveAt(bestSegmentIndex)

                        ' Reverse segment if needed
                        If needToReverse Then
                            nextSegment = ReverseSegment(nextSegment)
                        End If

                        ' Add to current path
                        currentPath.Segments.Add(nextSegment)
                        currentEndPoint = nextSegment.EndPoint
                    Else
                        ' No connecting segment found
                        Exit While
                    End If
                End While

                ' Set path closed property
                currentPath.IsClosed = pathClosed

                ' Only add non-empty paths
                If currentPath.Segments.Count > 0 Then
                    result.Add(currentPath)
                End If
            End While

            Return result
        End Function

        ''' <summary>
        ''' Tests if a point is inside a closed path
        ''' </summary>
        Private Function ContainsPoint(path As BezierPath, point As PointF) As Boolean
            If Not path.IsClosed Then Return False

            Dim crossings As Integer = 0

            For Each segment In path.Segments
                If segment.SegmentType = "L"c Then
                    ' Line segment ray casting
                    If (segment.StartPoint.Y <= point.Y And segment.EndPoint.Y > point.Y) Or
                    (segment.EndPoint.Y <= point.Y And segment.StartPoint.Y > point.Y) Then

                        Dim t As Double = (point.Y - segment.StartPoint.Y) / (segment.EndPoint.Y - segment.StartPoint.Y)
                        Dim x As Double = segment.StartPoint.X + t * (segment.EndPoint.X - segment.StartPoint.X)

                        If x > point.X Then
                            crossings += 1
                        End If
                    End If
                ElseIf segment.SegmentType = "C"c Then
                    ' For Bézier curves, use more accurate subdivision
                    Dim steps As Integer = 20 ' Use more steps for better accuracy
                    Dim prevPoint As PointF = segment.StartPoint

                    For i As Integer = 1 To steps
                        Dim t As Double = i / steps
                        Dim currPoint As PointF = segment.PointAt(t)

                        If (prevPoint.Y <= point.Y And currPoint.Y > point.Y) Or
                        (currPoint.Y <= point.Y And prevPoint.Y > point.Y) Then

                            ' Only divide if difference is significant
                            If Math.Abs(currPoint.Y - prevPoint.Y) > 0.00001 Then
                                Dim s As Double = (point.Y - prevPoint.Y) / (currPoint.Y - prevPoint.Y)
                                Dim x As Double = prevPoint.X + s * (currPoint.X - prevPoint.X)

                                If x > point.X Then
                                    crossings += 1
                                End If
                            End If
                        End If

                        prevPoint = currPoint
                    Next
                End If
            Next

            ' Odd number of crossings means point is inside
            Return (crossings Mod 2) = 1
        End Function



        ''' <summary>
        ''' Calculates the Euclidean distance between two points
        ''' </summary>
        Private Function PointDistance(p1 As PointF, p2 As PointF) As Double
            Dim dx As Double = p2.X - p1.X
            Dim dy As Double = p2.Y - p1.Y
            Return Math.Sqrt(dx * dx + dy * dy)
        End Function

        ''' <summary>
        ''' Reverses a segment (swaps start and end points and adjusts control points)
        ''' </summary>
        Private Function ReverseSegment(segment As BezierSegment) As BezierSegment
            Dim reversed As New BezierSegment()

            ' Swap start and end points
            reversed.StartPoint = segment.EndPoint
            reversed.EndPoint = segment.StartPoint

            ' For Bézier curves, swap and adjust control points
            If segment.SegmentType = "C"c Then
                ' For cubic Bézier curves, the control points need to be swapped and reversed
                reversed.Control1 = segment.Control2
                reversed.Control2 = segment.Control1
                reversed.SegmentType = "C"c
            Else
                ' For line segments, no control points to adjust
                reversed.SegmentType = "L"c
            End If

            ' Copy other properties
            reversed.PathIndex = segment.PathIndex
            reversed.SegmentIndex = segment.SegmentIndex
            reversed.Keep = segment.Keep
            reversed.IsInside = segment.IsInside

            ' Swap T1 and T2 values
            reversed.T1 = segment.T2
            reversed.T2 = segment.T1

            Return reversed
        End Function

        ''' <summary>
        ''' Removes duplicate intersection points that are very close to each other
        ''' </summary>
        Private Function CleanupIntersections(intersections As List(Of PathIntersection)) As List(Of PathIntersection)
            Dim result As New List(Of PathIntersection)()

            For i As Integer = 0 To intersections.Count - 1
                Dim isDuplicate As Boolean = False

                ' Check if this intersection is very close to any already added intersection
                For j As Integer = 0 To result.Count - 1
                    Dim dx As Double = intersections(i).IntersectionPoint.X - result(j).IntersectionPoint.X
                    Dim dy As Double = intersections(i).IntersectionPoint.Y - result(j).IntersectionPoint.Y
                    Dim distSquared As Double = dx * dx + dy * dy

                    If distSquared < 0.0001 Then ' Small threshold for duplicates
                        isDuplicate = True
                        Exit For
                    End If
                Next

                If Not isDuplicate Then
                    result.Add(intersections(i))
                End If
            Next

            Return result
        End Function

        ''' <summary>
        ''' Combines segments from paths for the specified boolean operation
        ''' </summary>
        Private Function CombineSegmentsForOperation(paths1 As List(Of BezierPath), paths2 As List(Of BezierPath)) As List(Of BezierSegment)
            Dim result As New List(Of BezierSegment)()

            ' Add segments from first set of paths
            For Each path In paths1
                For Each segment In path.Segments
                    If segment.Keep Then
                        result.Add(segment.Clone())
                    End If
                Next
            Next

            ' Add segments from second set of paths if provided
            If paths2 IsNot Nothing Then
                For Each path In paths2
                    For Each segment In path.Segments
                        If segment.Keep Then
                            result.Add(segment.Clone())
                        End If
                    Next
                Next
            End If

            Return result
        End Function

        ''' <summary>
        ''' Converts BezierPaths back to PathCommands
        ''' </summary>
        Private Function ConvertToPathCommandsImproved(bezierPaths As List(Of BezierPath)) As List(Of PathCommands)
            Dim result As New List(Of PathCommands)()

            For Each path In bezierPaths
                If path.Segments.Count = 0 Then Continue For

                ' Start with a MoveTo command
                Dim firstSegment As BezierSegment = path.Segments(0)
                result.Add(New PathCommands(firstSegment.StartPoint, Nothing, Nothing, "M"c))

                ' Add each segment
                For Each segment In path.Segments
                    If segment.SegmentType = "L"c Then
                        result.Add(New PathCommands(segment.EndPoint, Nothing, Nothing, "L"c))
                    ElseIf segment.SegmentType = "C"c Then
                        result.Add(New PathCommands(segment.EndPoint, segment.Control1, segment.Control2, "C"c))
                    End If
                Next

                ' Close the path if it should be closed
                If path.IsClosed Then
                    result.Add(New PathCommands(firstSegment.StartPoint, Nothing, Nothing, "Z"c))
                End If
            Next

            Return result
        End Function


        ''' <summary>
        ''' Reverses a complete path (for creating holes with proper winding order)
        ''' </summary>
        Private Function ReversePath(pathCommands As List(Of PathCommands)) As List(Of PathCommands)
            If pathCommands.Count = 0 Then Return New List(Of PathCommands)()

            Dim result As New List(Of PathCommands)()
            Dim points As New List(Of PointF)()

            ' Collect all points from the path
            For Each cmd In pathCommands
                If cmd.Pc = "M"c Or cmd.Pc = "L"c Or cmd.Pc = "C"c Then
                    points.Add(cmd.P)
                End If
            Next

            If points.Count = 0 Then Return New List(Of PathCommands)()

            ' Start with MoveTo at the last point
            result.Add(New PathCommands(points(points.Count - 1), Nothing, Nothing, "M"c))

            ' Add segments in reverse order
            For i As Integer = pathCommands.Count - 1 To 0 Step -1
                Dim cmd = pathCommands(i)

                If i = 0 Then Continue For ' Skip the first MoveTo

                Dim prevCmd = pathCommands(i - 1)

                If cmd.Pc = "L"c Then
                    result.Add(New PathCommands(prevCmd.P, Nothing, Nothing, "L"c))
                ElseIf cmd.Pc = "C"c Then
                    ' Reverse Bézier curve - swap control points
                    result.Add(New PathCommands(prevCmd.P, cmd.b2, cmd.b1, "C"c))
                ElseIf cmd.Pc = "Z"c Then
                    ' Don't add Z yet, will add at the end
                End If
            Next

            ' Close the path
            result.Add(New PathCommands(result(0).P, Nothing, Nothing, "Z"c))

            Return result
        End Function
        Private Function ExtractOuterBoundaryDirect(drawObjects As List(Of DrawObject)) As List(Of PathCommands)
            ' Simple approach: find the bounding rectangle of all objects and create a path around it
            If drawObjects.Count = 0 Then Return New List(Of PathCommands)

            ' Calculate overall bounds
            Dim minX As Single = Single.MaxValue
            Dim minY As Single = Single.MaxValue
            Dim maxX As Single = Single.MinValue
            Dim maxY As Single = Single.MinValue

            For Each obj In drawObjects
                Dim bounds = obj.BoundingRectangle
                minX = Math.Min(minX, bounds.Left)
                minY = Math.Min(minY, bounds.Top)
                maxX = Math.Max(maxX, bounds.Right)
                maxY = Math.Max(maxY, bounds.Bottom)
            Next

            ' Create a simple rectangle path around all objects
            Dim path As New List(Of PathCommands)
            path.Add(New PathCommands(New PointF(minX, minY), Nothing, Nothing, "M"c))
            path.Add(New PathCommands(New PointF(maxX, minY), Nothing, Nothing, "L"c))
            path.Add(New PathCommands(New PointF(maxX, maxY), Nothing, Nothing, "L"c))
            path.Add(New PathCommands(New PointF(minX, maxY), Nothing, Nothing, "L"c))
            path.Add(New PathCommands(New PointF(minX, minY), Nothing, Nothing, "Z"c))

            Return path
        End Function
        ''' <summary>
        ''' Converts DrawObject.PathCommands into a GraphicsPath
        ''' </summary>
        Private Function ConvertToGraphicsPath(obj As DrawObject) As GraphicsPath
            Dim gp As New GraphicsPath()

            Try
                Dim cmds As List(Of PathCommands) = obj.PathCommands
                If cmds Is Nothing OrElse cmds.Count = 0 Then Return gp

                Dim startPoint As PointF = PointF.Empty
                Dim currentPoint As PointF = PointF.Empty
                Dim figureStarted As Boolean = False

                For Each cmd In cmds
                    Dim commandType As Char = GetCommandChar(cmd)

                    Select Case commandType
                        Case "M"c
                            If figureStarted Then
                                gp.StartFigure()
                            End If
                            startPoint = cmd.P
                            currentPoint = cmd.P
                            figureStarted = True

                        Case "L"c
                            If figureStarted Then
                                gp.AddLine(currentPoint, cmd.P)
                                currentPoint = cmd.P
                            End If

                        Case "C"c
                            If figureStarted AndAlso Not cmd.b1.IsEmpty AndAlso Not cmd.b2.IsEmpty Then
                                gp.AddBezier(currentPoint, cmd.b1, cmd.b2, cmd.P)
                                currentPoint = cmd.P
                            ElseIf figureStarted Then
                                gp.AddLine(currentPoint, cmd.P)
                                currentPoint = cmd.P
                            End If

                        Case "Z"c
                            If figureStarted Then
                                gp.CloseFigure()
                                figureStarted = False
                            End If
                    End Select
                Next

            Catch ex As Exception
                MessageBox.Show($"ConvertToGraphicsPath Error: {ex.Message}")
            End Try

            Return gp
        End Function

        Private Function ConvertRegionToPathCommands(region As Region) As List(Of PathCommands)
            Dim pathCommands As New List(Of PathCommands)()

            Try
                ' Get the bounds of the region
                Using boundsPath As New GraphicsPath()
                    ' Get the region bounds as rectangles
                    Dim matrix As New Drawing2D.Matrix()
                    Dim rectangles() As RectangleF = region.GetRegionScans(matrix)

                    If rectangles IsNot Nothing AndAlso rectangles.Length > 0 Then
                        ' Create a path from the union of all rectangles
                        For Each rect In rectangles
                            boundsPath.AddRectangle(rect)
                        Next

                        ' Convert GraphicsPath to PathCommands
                        If boundsPath.PointCount > 0 Then
                            Dim points As PointF() = boundsPath.PathPoints
                            Dim types As Byte() = boundsPath.PathTypes

                            For i As Integer = 0 To points.Length - 1
                                If (types(i) And CByte(PathPointType.Start)) <> 0 Then
                                    ' Start new figure
                                    pathCommands.Add(New PathCommands(points(i), Nothing, Nothing, "M"c))
                                ElseIf (types(i) And CByte(PathPointType.Bezier)) <> 0 Then
                                    ' Bezier curve - get next two control points
                                    If i + 2 < points.Length Then
                                        pathCommands.Add(New PathCommands(points(i + 2), points(i), points(i + 1), "C"c))
                                        i += 2
                                    End If
                                Else
                                    ' Line segment
                                    pathCommands.Add(New PathCommands(points(i), Nothing, Nothing, "L"c))
                                End If
                            Next

                            ' Close the path
                            If pathCommands.Count > 0 Then
                                pathCommands.Add(New PathCommands(pathCommands(0).P, Nothing, Nothing, "Z"c))
                            End If
                        End If
                    End If
                End Using

            Catch ex As Exception
                MessageBox.Show($"ConvertRegionToPathCommands Error: {ex.Message}")
            End Try

            Return pathCommands
        End Function


        ''' <summary>
        ''' Determines the command character property name dynamically.
        ''' </summary>
        Private Function GetCommandChar(cmd As Object) As Char
            Try
                Dim t = cmd.GetType()
                If t.GetProperty("Cmd") IsNot Nothing Then
                    Return CChar(t.GetProperty("Cmd").GetValue(cmd, Nothing))
                ElseIf t.GetProperty("Command") IsNot Nothing Then
                    Return CChar(t.GetProperty("Command").GetValue(cmd, Nothing))
                ElseIf t.GetProperty("Type") IsNot Nothing Then
                    Return CChar(t.GetProperty("Type").GetValue(cmd, Nothing))
                End If
            Catch
            End Try
            Return "L"c
        End Function

        ''' <summary>
        ''' Marks segments that are on the outer boundary of the union
        ''' A segment is on the outer boundary if it's NOT inside any other shape
        ''' </summary>
        Private Sub MarkOuterBoundarySegments(splitPaths As List(Of BezierPath), originalPaths As List(Of BezierPath))
            For Each path In splitPaths
                For Each segment In path.Segments
                    ' Test the midpoint of this segment
                    Dim testPoint = segment.PointAt(0.5)

                    ' Count how many shapes contain this point
                    Dim containmentCount As Integer = 0

                    For Each originalPath In originalPaths
                        If originalPath.IsClosed AndAlso ContainsPoint(originalPath, testPoint) Then
                            containmentCount += 1
                        End If
                    Next

                    ' Segment is on outer boundary if it's inside exactly ONE shape
                    ' (or zero if it's on a non-overlapping part)
                    ' If inside 2+ shapes, it's an internal edge we don't want
                    segment.Keep = (containmentCount <= 1)
                Next
            Next
        End Sub

        ''' <summary>
        ''' Walks connected outer boundary segments to form complete paths
        ''' Uses angle-based selection to prefer continuing along the boundary
        ''' </summary>
        Private Function WalkOuterBoundary(segments As List(Of BezierSegment)) As List(Of BezierPath)
            Dim result As New List(Of BezierPath)
            If segments.Count = 0 Then Return result

            ' Build endpoint lookup map with tolerance
            Const TOLERANCE As Double = 0.01
            Dim endpointMap As New Dictionary(Of String, List(Of BezierSegment))

            For Each seg In segments
                If Not seg.Keep Then Continue For

                Dim startKey = PointToKeyRounded(seg.StartPoint, TOLERANCE)
                Dim endKey = PointToKeyRounded(seg.EndPoint, TOLERANCE)

                If Not endpointMap.ContainsKey(endKey) Then
                    endpointMap(endKey) = New List(Of BezierSegment)
                End If
                endpointMap(endKey).Add(seg)
            Next

            Dim used As New HashSet(Of BezierSegment)
            Dim keptSegments = segments.Where(Function(s) s.Keep).ToList()

            ' Keep walking until all segments are used
            While used.Count < keptSegments.Count
                Dim seed = keptSegments.FirstOrDefault(Function(s) Not used.Contains(s))
                If seed Is Nothing Then Exit While

                Dim path As New BezierPath()
                Dim currentSeg = seed
                Dim startPoint = seed.StartPoint

                Dim maxSteps As Integer = keptSegments.Count + 10
                Dim stepCount As Integer = 0

                Do While stepCount < maxSteps
                    ' Add current segment
                    path.Segments.Add(DirectCast(currentSeg.Clone(), BezierSegment))
                    used.Add(currentSeg)
                    stepCount += 1

                    Dim currentEnd = currentSeg.EndPoint

                    ' Check if we closed the loop
                    If PointDistance(currentEnd, startPoint) < TOLERANCE AndAlso path.Segments.Count > 1 Then
                        path.IsClosed = True
                        Exit Do
                    End If

                    ' Find next segment using angle-based selection (prefer rightmost turn = outer boundary)
                    Dim endKey = PointToKeyRounded(currentEnd, TOLERANCE)
                    Dim candidates As List(Of BezierSegment) = Nothing

                    If Not endpointMap.TryGetValue(endKey, candidates) Then Exit Do

                    ' Filter out used segments
                    candidates = candidates.Where(Function(c) Not used.Contains(c)).ToList()
                    If candidates.Count = 0 Then Exit Do

                    ' Calculate incoming direction vector
                    Dim incomingDir = GetSegmentExitDirection(currentSeg)

                    ' Find segment that makes the rightmost turn (for outer boundary)
                    Dim bestSeg As BezierSegment = Nothing
                    Dim bestAngle As Double = -Math.PI * 2 ' Start with very negative

                    For Each cand In candidates
                        ' Check both forward and reverse directions
                        For Each reversed In {False, True}
                            Dim testSeg = If(reversed, ReverseSegment(cand), cand)

                            ' Check if this segment starts at our current end point
                            If PointDistance(testSeg.StartPoint, currentEnd) > TOLERANCE Then Continue For

                            Dim outgoingDir = GetSegmentEntryDirection(testSeg)

                            ' Calculate turn angle (positive = left turn, negative = right turn)
                            Dim angle = CalculateSignedAngle(incomingDir, outgoingDir)

                            ' For outer boundary, prefer rightmost turn (most negative angle)
                            ' But also accept left turns if no right turn available
                            If bestSeg Is Nothing OrElse angle > bestAngle Then
                                bestAngle = angle
                                bestSeg = testSeg
                            End If
                        Next
                    Next

                    If bestSeg Is Nothing Then Exit Do
                    currentSeg = bestSeg
                Loop

                If path.Segments.Count > 0 Then
                    result.Add(path)
                End If
            End While

            Return result
        End Function

        ''' <summary>
        ''' Gets the direction vector at the exit point of a segment
        ''' </summary>
        Private Function GetSegmentExitDirection(segment As BezierSegment) As PointF
            If segment.SegmentType = "L"c Then
                ' For lines, direction is simply end - start
                Return New PointF(
            segment.EndPoint.X - segment.StartPoint.X,
            segment.EndPoint.Y - segment.StartPoint.Y)
            ElseIf segment.SegmentType = "C"c Then
                ' For curves, use the direction from last control point to end
                Return New PointF(
            segment.EndPoint.X - segment.Control2.X,
            segment.EndPoint.Y - segment.Control2.Y)
            Else
                Return New PointF(
            segment.EndPoint.X - segment.StartPoint.X,
            segment.EndPoint.Y - segment.StartPoint.Y)
            End If
        End Function

        ''' <summary>
        ''' Gets the direction vector at the entry point of a segment
        ''' </summary>
        Private Function GetSegmentEntryDirection(segment As BezierSegment) As PointF
            If segment.SegmentType = "L"c Then
                Return New PointF(
            segment.EndPoint.X - segment.StartPoint.X,
            segment.EndPoint.Y - segment.StartPoint.Y)
            ElseIf segment.SegmentType = "C"c Then
                ' For curves, use the direction from start to first control point
                Return New PointF(
            segment.Control1.X - segment.StartPoint.X,
            segment.Control1.Y - segment.StartPoint.Y)
            Else
                Return New PointF(
            segment.EndPoint.X - segment.StartPoint.X,
            segment.EndPoint.Y - segment.StartPoint.Y)
            End If
        End Function

        ''' <summary>
        ''' Calculates signed angle between two direction vectors
        ''' Positive = counter-clockwise (left turn), Negative = clockwise (right turn)
        ''' </summary>
        Private Function CalculateSignedAngle(dir1 As PointF, dir2 As PointF) As Double
            ' Normalize vectors
            Dim len1 = Math.Sqrt(dir1.X * dir1.X + dir1.Y * dir1.Y)
            Dim len2 = Math.Sqrt(dir2.X * dir2.X + dir2.Y * dir2.Y)

            If len1 < 0.0001 OrElse len2 < 0.0001 Then Return 0

            Dim n1 = New PointF(dir1.X / len1, dir1.Y / len1)
            Dim n2 = New PointF(dir2.X / len2, dir2.Y / len2)

            ' Cross product (z-component) gives signed area
            Dim cross = n1.X * n2.Y - n1.Y * n2.X

            ' Dot product gives angle
            Dim dot = n1.X * n2.X + n1.Y * n2.Y

            ' atan2 gives signed angle
            Return Math.Atan2(cross, dot)
        End Function

        ''' <summary>
        ''' Rounds point coordinates and creates a lookup key
        ''' </summary>
        Private Function PointToKeyRounded(p As PointF, tolerance As Double) As String
            ' Round to tolerance grid
            Dim x = Math.Round(p.X / tolerance) * tolerance
            Dim y = Math.Round(p.Y / tolerance) * tolerance
            Return $"{x:F3},{y:F3}"
        End Function



#End Region

#Region "Intersection Detection"

        ''' <summary>
        ''' Finds intersection between two line segments
        ''' </summary>
        Private Function FindLineLineIntersection(a1 As PointF, a2 As PointF, b1 As PointF, b2 As PointF) As Tuple(Of PointF, Double, Double)
            ' Line A represented as a1 + t1(a2 - a1)
            ' Line B represented as b1 + t2(b2 - b1)

            Dim a1x As Double = a1.X
            Dim a1y As Double = a1.Y
            Dim a2x As Double = a2.X
            Dim a2y As Double = a2.Y
            Dim b1x As Double = b1.X
            Dim b1y As Double = b1.Y
            Dim b2x As Double = b2.X
            Dim b2y As Double = b2.Y

            ' Calculate direction vectors
            Dim vax As Double = a2x - a1x
            Dim vay As Double = a2y - a1y
            Dim vbx As Double = b2x - b1x
            Dim vby As Double = b2y - b1y

            ' Calculate the cross product determinant
            Dim denominator As Double = vax * vby - vay * vbx

            ' If denominator is zero, lines are parallel
            If Math.Abs(denominator) < 0.0001 Then Return Nothing

            ' Calculate difference in start points
            Dim dx As Double = b1x - a1x
            Dim dy As Double = b1y - a1y

            ' Calculate parameters
            Dim t1 As Double = (dx * vby - dy * vbx) / denominator
            Dim t2 As Double = (vax * dy - vay * dx) / denominator

            ' Check if intersection is within both line segments
            If t1 >= 0 And t1 <= 1 And t2 >= 0 And t2 <= 1 Then
                ' Calculate intersection point
                Dim x As Single = a1x + t1 * vax
                Dim y As Single = a1y + t1 * vay

                Return New Tuple(Of PointF, Double, Double)(New PointF(x, y), t1, t2)
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' Finds intersections between a line segment and a cubic Bézier curve
        ''' </summary>
        Private Function FindLineCurveIntersections(lineStart As PointF, lineEnd As PointF,
                                             p0 As PointF, p1 As PointF, p2 As PointF, p3 As PointF) As List(Of Tuple(Of PointF, Double, Double))
            Dim intersections As New List(Of Tuple(Of PointF, Double, Double))()

            ' First check if we can use simple subdivision approach
            Dim depth As Integer = 0
            Dim maxDepth As Integer = 8 ' Maximum recursion depth

            FindLineCurveIntersectionsRecursive(lineStart, lineEnd, p0, p1, p2, p3, 0.0, 1.0, depth, maxDepth, intersections)

            Return intersections
        End Function

        ''' <summary>
        ''' Recursively finds intersections between a line and a curve using subdivision
        ''' </summary>
        Private Sub FindLineCurveIntersectionsRecursive(lineStart As PointF, lineEnd As PointF,
                                                 p0 As PointF, p1 As PointF, p2 As PointF, p3 As PointF,
                                                 tMin As Double, tMax As Double, depth As Integer, maxDepth As Integer,
                                                 ByRef intersections As List(Of Tuple(Of PointF, Double, Double)))
            ' Get curve bounding box
            Dim minX As Single = Math.Min(Math.Min(p0.X, p1.X), Math.Min(p2.X, p3.X))
            Dim minY As Single = Math.Min(Math.Min(p0.Y, p1.Y), Math.Min(p2.Y, p3.Y))
            Dim maxX As Single = Math.Max(Math.Max(p0.X, p1.X), Math.Max(p2.X, p3.X))
            Dim maxY As Single = Math.Max(Math.Max(p0.Y, p1.Y), Math.Max(p2.Y, p3.Y))

            Dim curveBounds As New RectangleF(minX, minY, maxX - minX, maxY - minY)

            ' Get line bounding box
            Dim lineMinX As Single = Math.Min(lineStart.X, lineEnd.X)
            Dim lineMinY As Single = Math.Min(lineStart.Y, lineEnd.Y)
            Dim lineMaxX As Single = Math.Max(lineStart.X, lineEnd.X)
            Dim lineMaxY As Single = Math.Max(lineStart.Y, lineEnd.Y)

            Dim lineBounds As New RectangleF(lineMinX, lineMinY, lineMaxX - lineMinX, lineMaxY - lineMinY)

            ' Check if bounding boxes intersect
            If Not curveBounds.IntersectsWith(lineBounds) Then Return

            ' If we're at maximum depth or the curve is small enough, check line intersection
            If depth >= maxDepth OrElse (maxX - minX < 1.0 And maxY - minY < 1.0) Then
                ' Use line-line approximation
                Dim intersection As Tuple(Of PointF, Double, Double) = FindLineLineIntersection(
                lineStart, lineEnd, p0, p3)

                If intersection IsNot Nothing Then
                    ' Calculate t-value on the curve
                    Dim tCurve As Double = tMin + (tMax - tMin) * intersection.Item2
                    intersections.Add(New Tuple(Of PointF, Double, Double)(intersection.Item1, intersection.Item2, tCurve))
                End If
                Return
            End If

            ' Subdivide the curve
            Dim tMid As Double = (tMin + tMax) / 2

            ' Calculate mid-point using de Casteljau's algorithm
            Dim p01 As PointF = Midpoint(p0, p1)
            Dim p12 As PointF = Midpoint(p1, p2)
            Dim p23 As PointF = Midpoint(p2, p3)

            Dim p012 As PointF = Midpoint(p01, p12)
            Dim p123 As PointF = Midpoint(p12, p23)

            Dim p0123 As PointF = Midpoint(p012, p123)

            ' Recursively find intersections in each half
            FindLineCurveIntersectionsRecursive(lineStart, lineEnd, p0, p01, p012, p0123, tMin, tMid, depth + 1, maxDepth, intersections)
            FindLineCurveIntersectionsRecursive(lineStart, lineEnd, p0123, p123, p23, p3, tMid, tMax, depth + 1, maxDepth, intersections)
        End Sub

        ''' <summary>
        ''' Finds intersections between two cubic Bézier curves
        ''' </summary>
        Private Function FindCurveCurveIntersections(a0 As PointF, a1 As PointF, a2 As PointF, a3 As PointF,
                                               b0 As PointF, b1 As PointF, b2 As PointF, b3 As PointF) As List(Of Tuple(Of PointF, Double, Double))
            Dim intersections As New List(Of Tuple(Of PointF, Double, Double))()

            ' First check if bounding boxes overlap
            Dim minAx As Single = Math.Min(Math.Min(a0.X, a1.X), Math.Min(a2.X, a3.X))
            Dim minAy As Single = Math.Min(Math.Min(a0.Y, a1.Y), Math.Min(a2.Y, a3.Y))
            Dim maxAx As Single = Math.Max(Math.Max(a0.X, a1.X), Math.Max(a2.X, a3.X))
            Dim maxAy As Single = Math.Max(Math.Max(a0.Y, a1.Y), Math.Max(a2.Y, a3.Y))

            Dim minBx As Single = Math.Min(Math.Min(b0.X, b1.X), Math.Min(b2.X, b3.X))
            Dim minBy As Single = Math.Min(Math.Min(b0.Y, b1.Y), Math.Min(b2.Y, b3.Y))
            Dim maxBx As Single = Math.Max(Math.Max(b0.X, b1.X), Math.Max(b2.X, b3.X))
            Dim maxBy As Single = Math.Max(Math.Max(b0.Y, b1.Y), Math.Max(b2.Y, b3.Y))

            Dim boundsA As New RectangleF(minAx, minAy, maxAx - minAx, maxAy - minAy)
            Dim boundsB As New RectangleF(minBx, minBy, maxBx - minBx, maxBy - minBy)

            If Not boundsA.IntersectsWith(boundsB) Then Return intersections

            ' Use recursive subdivision approach
            Dim depth As Integer = 0
            Dim maxDepth As Integer = 10  ' Increased depth for better accuracy

            FindCurveCurveIntersectionsRecursive(a0, a1, a2, a3, 0.0, 1.0, b0, b1, b2, b3, 0.0, 1.0, depth, maxDepth, intersections)

            Return intersections
        End Function

        ''' <summary>
        ''' Recursively finds intersections between two curves using subdivision
        ''' </summary>
        Private Sub FindCurveCurveIntersectionsRecursive(a0 As PointF, a1 As PointF, a2 As PointF, a3 As PointF, tMinA As Double, tMaxA As Double,
                                                  b0 As PointF, b1 As PointF, b2 As PointF, b3 As PointF, tMinB As Double, tMaxB As Double,
                                                  depth As Integer, maxDepth As Integer,
                                                  ByRef intersections As List(Of Tuple(Of PointF, Double, Double)))
            ' Get bounding boxes
            Dim minAx As Single = Math.Min(Math.Min(a0.X, a1.X), Math.Min(a2.X, a3.X))
            Dim minAy As Single = Math.Min(Math.Min(a0.Y, a1.Y), Math.Min(a2.Y, a3.Y))
            Dim maxAx As Single = Math.Max(Math.Max(a0.X, a1.X), Math.Max(a2.X, a3.X))
            Dim maxAy As Single = Math.Max(Math.Max(a0.Y, a1.Y), Math.Max(a2.Y, a3.Y))

            Dim minBx As Single = Math.Min(Math.Min(b0.X, b1.X), Math.Min(b2.X, b3.X))
            Dim minBy As Single = Math.Min(Math.Min(b0.Y, b1.Y), Math.Min(b2.Y, b3.Y))
            Dim maxBx As Single = Math.Max(Math.Max(b0.X, b1.X), Math.Max(b2.X, b3.X))
            Dim maxBy As Single = Math.Max(Math.Max(b0.Y, b1.Y), Math.Max(b2.Y, b3.Y))

            Dim boundsA As New RectangleF(minAx, minAy, maxAx - minAx, maxAy - minAy)
            Dim boundsB As New RectangleF(minBx, minBy, maxBx - minBx, maxBy - minBy)

            ' Early exit if bounding boxes don't intersect
            If Not boundsA.IntersectsWith(boundsB) Then Return

            ' If we're at maximum depth or the curves are small enough, check as line segments
            If depth >= maxDepth OrElse ((maxAx - minAx < 0.5 And maxAy - minAy < 0.5) And (maxBx - minBx < 0.5 And maxBy - minBy < 0.5)) Then
                ' Approximate with line segments
                Dim intersection As Tuple(Of PointF, Double, Double) = FindLineLineIntersection(a0, a3, b0, b3)

                If intersection IsNot Nothing Then
                    ' Calculate t-values on the original curves
                    Dim tA As Double = tMinA + (tMaxA - tMinA) * intersection.Item2
                    Dim tB As Double = tMinB + (tMaxB - tMinB) * intersection.Item3

                    ' Add to results
                    intersections.Add(New Tuple(Of PointF, Double, Double)(intersection.Item1, tA, tB))
                End If
                Return
            End If

            ' Determine which curve to split (the one with larger bounding box)
            Dim areaA As Double = boundsA.Width * boundsA.Height
            Dim areaB As Double = boundsB.Width * boundsB.Height

            If areaA > areaB Then
                ' Split curve A
                Dim tMidA As Double = (tMinA + tMaxA) / 2

                ' Calculate split using de Casteljau's algorithm
                Dim a01 As PointF = Midpoint(a0, a1)
                Dim a12 As PointF = Midpoint(a1, a2)
                Dim a23 As PointF = Midpoint(a2, a3)

                Dim a012 As PointF = Midpoint(a01, a12)
                Dim a123 As PointF = Midpoint(a12, a23)

                Dim a0123 As PointF = Midpoint(a012, a123)

                ' Recursively check both halves
                FindCurveCurveIntersectionsRecursive(a0, a01, a012, a0123, tMinA, tMidA, b0, b1, b2, b3, tMinB, tMaxB, depth + 1, maxDepth, intersections)
                FindCurveCurveIntersectionsRecursive(a0123, a123, a23, a3, tMidA, tMaxA, b0, b1, b2, b3, tMinB, tMaxB, depth + 1, maxDepth, intersections)
            Else
                ' Split curve B
                Dim tMidB As Double = (tMinB + tMaxB) / 2

                ' Calculate split using de Casteljau's algorithm
                Dim b01 As PointF = Midpoint(b0, b1)
                Dim b12 As PointF = Midpoint(b1, b2)
                Dim b23 As PointF = Midpoint(b2, b3)

                Dim b012 As PointF = Midpoint(b01, b12)
                Dim b123 As PointF = Midpoint(b12, b23)

                Dim b0123 As PointF = Midpoint(b012, b123)

                ' Recursively check both halves
                FindCurveCurveIntersectionsRecursive(a0, a1, a2, a3, tMinA, tMaxA, b0, b01, b012, b0123, tMinB, tMidB, depth + 1, maxDepth, intersections)
                FindCurveCurveIntersectionsRecursive(a0, a1, a2, a3, tMinA, tMaxA, b0123, b123, b23, b3, tMidB, tMaxB, depth + 1, maxDepth, intersections)
            End If
        End Sub

        ''' <summary>
        ''' Calculates the midpoint between two points
        ''' </summary>
        Private Function Midpoint(p1 As PointF, p2 As PointF) As PointF
            Return New PointF((p1.X + p2.X) / 2, (p1.Y + p2.Y) / 2)
        End Function

#End Region

#Region "Path Manipulation"

        ''' <summary>
        ''' Splits paths at intersection points
        ''' </summary>
        Private Function SplitPathAtIntersections(paths As List(Of BezierPath), intersections As List(Of PathIntersection), isFirstPath As Boolean) As List(Of BezierPath)
            Dim result As New List(Of BezierPath)()

            ' Clone the original paths first
            For Each path In paths
                Dim clonedPath As New BezierPath()
                clonedPath.PathIndex = path.PathIndex
                clonedPath.IsClosed = path.IsClosed

                ' Deep copy all segments
                clonedPath.Segments = New List(Of BezierSegment)()
                For Each segment In path.Segments
                    clonedPath.Segments.Add(segment.Clone())
                Next

                result.Add(clonedPath)
            Next

            ' Process each path and split at intersection points
            For i As Integer = 0 To result.Count - 1
                Dim currentPath As BezierPath = result(i)

                ' Get all intersections affecting this path
                Dim pathIntersections As New List(Of PathIntersection)()
                For Each intersection In intersections
                    If (isFirstPath And intersection.Path1Index = currentPath.PathIndex) Or
                    (Not isFirstPath And intersection.Path2Index = currentPath.PathIndex) Then
                        pathIntersections.Add(intersection)
                    End If
                Next

                ' Sort intersections by segment index and t-parameter
                pathIntersections.Sort(Function(a, b)
                                           Dim segIndexA As Integer = If(isFirstPath, a.Segment1Index, a.Segment2Index)
                                           Dim segIndexB As Integer = If(isFirstPath, b.Segment1Index, b.Segment2Index)

                                           If segIndexA = segIndexB Then
                                               Dim tA As Double = If(isFirstPath, a.T1, a.T2)
                                               Dim tB As Double = If(isFirstPath, b.T1, b.T2)
                                               Return tA.CompareTo(tB)
                                           End If

                                           Return segIndexA.CompareTo(segIndexB)
                                       End Function)

                ' Now split each segment at intersection points
                Dim newSegments As New List(Of BezierSegment)()

                ' Organize intersections by segment
                Dim intersectionsBySegment As New Dictionary(Of Integer, List(Of Tuple(Of PathIntersection, Double)))()

                For Each intersection In pathIntersections
                    Dim segIdx As Integer = If(isFirstPath, intersection.Segment1Index, intersection.Segment2Index)
                    Dim t As Double = If(isFirstPath, intersection.T1, intersection.T2)

                    If Not intersectionsBySegment.ContainsKey(segIdx) Then
                        intersectionsBySegment(segIdx) = New List(Of Tuple(Of PathIntersection, Double))()
                    End If

                    intersectionsBySegment(segIdx).Add(New Tuple(Of PathIntersection, Double)(intersection, t))
                Next

                ' Process each segment
                For segIdx As Integer = 0 To currentPath.Segments.Count - 1
                    Dim segment As BezierSegment = currentPath.Segments(segIdx)

                    If intersectionsBySegment.ContainsKey(segIdx) Then
                        ' Sort t-values in ascending order
                        Dim segmentIntersections As List(Of Tuple(Of PathIntersection, Double)) = intersectionsBySegment(segIdx)
                        segmentIntersections.Sort(Function(a, b) a.Item2.CompareTo(b.Item2))

                        ' Process each split point
                        Dim currentSegment As BezierSegment = segment.Clone()
                        Dim lastT As Double = 0.0

                        For Each Itema In segmentIntersections
                            Dim intersection As PathIntersection = Itema.Item1
                            Dim t As Double = Itema.Item2

                            ' Adjust t relative to current segment
                            Dim relativeT As Double = (t - lastT) / (1.0 - lastT)

                            ' Avoid splitting at endpoints or with very small segments
                            If relativeT <= 0.001 Or relativeT >= 0.999 Then Continue For

                            ' Split the segment
                            Dim splitResult As Tuple(Of BezierSegment, BezierSegment) = SplitSegment(currentSegment, relativeT)

                            ' Add first half to result
                            newSegments.Add(splitResult.Item1)

                            ' Update for next iteration
                            currentSegment = splitResult.Item2
                            lastT = t
                        Next

                        ' Add the last segment
                        newSegments.Add(currentSegment)
                    Else
                        ' No intersections for this segment, just add it as is
                        newSegments.Add(segment.Clone())
                    End If
                Next

                ' Replace segments in the path
                currentPath.Segments = newSegments
            Next

            Return result
        End Function

        ''' <summary>
        ''' Splits a segment at parameter t (0.0 to 1.0)
        ''' </summary>
        Private Function SplitSegment(segment As BezierSegment, t As Double) As Tuple(Of BezierSegment, BezierSegment)
            If t <= 0.0 Or t >= 1.0 Then
                Throw New ArgumentException("Split parameter must be between 0.0 and 1.0 exclusive")
            End If

            If segment.SegmentType = "L"c Then
                ' Split a line segment at parameter t
                Dim midPoint As PointF = segment.PointAt(t)

                Dim segment1 As New BezierSegment(segment.StartPoint, midPoint, Nothing, Nothing, "L"c)
                Dim segment2 As New BezierSegment(midPoint, segment.EndPoint, Nothing, Nothing, "L"c)

                ' Set path indices and relative t-values
                segment1.PathIndex = segment.PathIndex
                segment2.PathIndex = segment.PathIndex
                segment1.SegmentIndex = segment.SegmentIndex
                segment2.SegmentIndex = segment.SegmentIndex
                segment1.T1 = segment.T1
                segment1.T2 = segment.T1 + (segment.T2 - segment.T1) * t
                segment2.T1 = segment.T1 + (segment.T2 - segment.T1) * t
                segment2.T2 = segment.T2

                Return New Tuple(Of BezierSegment, BezierSegment)(segment1, segment2)
            ElseIf segment.SegmentType = "C"c Then
                ' Split a cubic Bézier curve at parameter t using de Casteljau's algorithm
                Dim p0 As PointF = segment.StartPoint
                Dim p1 As PointF = segment.Control1
                Dim p2 As PointF = segment.Control2
                Dim p3 As PointF = segment.EndPoint

                ' Calculate new control points
                Dim p01 As PointF = New PointF((1 - t) * p0.X + t * p1.X, (1 - t) * p0.Y + t * p1.Y)
                Dim p12 As PointF = New PointF((1 - t) * p1.X + t * p2.X, (1 - t) * p1.Y + t * p2.Y)
                Dim p23 As PointF = New PointF((1 - t) * p2.X + t * p3.X, (1 - t) * p2.Y + t * p3.Y)

                Dim p012 As PointF = New PointF((1 - t) * p01.X + t * p12.X, (1 - t) * p01.Y + t * p12.Y)
                Dim p123 As PointF = New PointF((1 - t) * p12.X + t * p23.X, (1 - t) * p12.Y + t * p23.Y)

                Dim p0123 As PointF = New PointF((1 - t) * p012.X + t * p123.X, (1 - t) * p012.Y + t * p123.Y)

                ' First half of the curve
                Dim segment1 As New BezierSegment(p0, p0123, p01, p012, "C"c)

                ' Second half of the curve
                Dim segment2 As New BezierSegment(p0123, p3, p123, p23, "C"c)

                ' Set path indices and relative t-values
                segment1.PathIndex = segment.PathIndex
                segment2.PathIndex = segment.PathIndex
                segment1.SegmentIndex = segment.SegmentIndex
                segment2.SegmentIndex = segment.SegmentIndex
                segment1.T1 = segment.T1
                segment1.T2 = segment.T1 + (segment.T2 - segment.T1) * t
                segment2.T1 = segment.T1 + (segment.T2 - segment.T1) * t
                segment2.T2 = segment.T2

                Return New Tuple(Of BezierSegment, BezierSegment)(segment1, segment2)
            Else
                ' For other types, treat as a line segment
                Dim midPoint As PointF = segment.PointAt(t)

                Dim segment1 As New BezierSegment(segment.StartPoint, midPoint, Nothing, Nothing, "L"c)
                Dim segment2 As New BezierSegment(midPoint, segment.EndPoint, Nothing, Nothing, "L"c)

                segment1.PathIndex = segment.PathIndex
                segment2.PathIndex = segment.PathIndex
                segment1.SegmentIndex = segment.SegmentIndex
                segment2.SegmentIndex = segment.SegmentIndex
                segment1.T1 = segment.T1
                segment1.T2 = segment.T1 + (segment.T2 - segment.T1) * t
                segment2.T1 = segment.T1 + (segment.T2 - segment.T1) * t
                segment2.T2 = segment.T2

                Return New Tuple(Of BezierSegment, BezierSegment)(segment1, segment2)
            End If
        End Function

        ''' <summary>
        ''' Finds all intersections between two sets of paths
        ''' </summary>
        Private Function FindAllPathIntersections(paths1 As List(Of BezierPath), paths2 As List(Of BezierPath)) As List(Of PathIntersection)
            Dim intersections As New List(Of PathIntersection)()

            ' First do a quick bounding box test to eliminate non-intersecting paths
            For Each path1 In paths1
                Dim bounds1 As RectangleF = path1.GetBoundingBox()

                For Each path2 In paths2
                    Dim bounds2 As RectangleF = path2.GetBoundingBox()

                    ' Only test segment intersections if bounding boxes overlap
                    If bounds1.IntersectsWith(bounds2) Then
                        For i As Integer = 0 To path1.Segments.Count - 1
                            Dim seg1 As BezierSegment = path1.Segments(i)

                            For j As Integer = 0 To path2.Segments.Count - 1
                                Dim seg2 As BezierSegment = path2.Segments(j)

                                ' Find intersections between these segments
                                Dim segIntersections As List(Of Tuple(Of PointF, Double, Double)) = FindSegmentIntersections(seg1, seg2)

                                For Each intersection In segIntersections
                                    Dim point As PointF = intersection.Item1
                                    Dim t1 As Double = intersection.Item2
                                    Dim t2 As Double = intersection.Item3

                                    intersections.Add(New PathIntersection(
                                     point,
                                     path1.PathIndex,
                                     path2.PathIndex,
                                     i,
                                     j,
                                     t1,
                                     t2))
                                Next
                            Next
                        Next
                    End If
                Next
            Next

            Return intersections
        End Function

        ''' <summary>
        ''' Finds intersections between two Bézier segments
        ''' </summary>
        Private Function FindSegmentIntersections(segment1 As BezierSegment, segment2 As BezierSegment) As List(Of Tuple(Of PointF, Double, Double))
            Dim result As New List(Of Tuple(Of PointF, Double, Double))()

            ' Quick bounding box test first
            Dim box1 As RectangleF = GetSegmentBoundingBox(segment1)
            Dim box2 As RectangleF = GetSegmentBoundingBox(segment2)

            If Not box1.IntersectsWith(box2) Then Return result

            ' Different types of intersection tests based on segment types
            If segment1.SegmentType = "L"c And segment2.SegmentType = "L"c Then
                ' Line-line intersection
                Dim intersection As Tuple(Of PointF, Double, Double) = FindLineLineIntersection(
                 segment1.StartPoint, segment1.EndPoint,
                 segment2.StartPoint, segment2.EndPoint)

                If intersection IsNot Nothing Then
                    result.Add(intersection)
                End If
            ElseIf segment1.SegmentType = "L"c And segment2.SegmentType = "C"c Then
                ' Line-curve intersection
                result.AddRange(FindLineCurveIntersections(
                 segment1.StartPoint, segment1.EndPoint,
                 segment2.StartPoint, segment2.Control1, segment2.Control2, segment2.EndPoint))
            ElseIf segment1.SegmentType = "C"c And segment2.SegmentType = "L"c Then
                ' Curve-line intersection (swap order)
                Dim swappedResults = FindLineCurveIntersections(
                 segment2.StartPoint, segment2.EndPoint,
                 segment1.StartPoint, segment1.Control1, segment1.Control2, segment1.EndPoint)

                ' Swap t1 and t2 in the results
                For Each Itema In swappedResults
                    result.Add(New Tuple(Of PointF, Double, Double)(Itema.Item1, Itema.Item3, Itema.Item2))
                Next
            ElseIf segment1.SegmentType = "C"c And segment2.SegmentType = "C"c Then
                ' Curve-curve intersection
                result.AddRange(FindCurveCurveIntersections(
                 segment1.StartPoint, segment1.Control1, segment1.Control2, segment1.EndPoint,
                 segment2.StartPoint, segment2.Control1, segment2.Control2, segment2.EndPoint))
            End If

            Return result
        End Function

#End Region


#Region "Path Classes"

        ''' <summary>
        ''' Represents a Bézier curve segment with all control points
        ''' </summary>
        Private Class BezierSegment
            Implements ICloneable

            Public StartPoint As PointF
            Public EndPoint As PointF
            Public Control1 As PointF
            Public Control2 As PointF
            Public SegmentType As Char
            Public Keep As Boolean = True
            Public IsInside As Boolean = False
            Public PathIndex As Integer
            Public SegmentIndex As Integer
            Public T1 As Double = 0.0
            Public T2 As Double = 1.0

            Public Sub New()
            End Sub

            Public Sub New(startPt As PointF, endPt As PointF, ctrl1 As PointF, ctrl2 As PointF, type As Char)
                StartPoint = startPt
                EndPoint = endPt
                Control1 = ctrl1
                Control2 = ctrl2
                SegmentType = type
            End Sub

            ' ✅ Implement ICloneable.Clone properly
            Public Function Clone() As Object Implements ICloneable.Clone
                Return Me.MemberwiseClone()
            End Function

            ' Calculate point at parameter t (0.0 to 1.0)
            Public Function PointAt(t As Double) As PointF
                If SegmentType = "L"c Then
                    Return New PointF(
                StartPoint.X + t * (EndPoint.X - StartPoint.X),
                StartPoint.Y + t * (EndPoint.Y - StartPoint.Y))
                ElseIf SegmentType = "C"c Then
                    Dim mt As Double = 1.0 - t
                    Dim x As Single = (mt * mt * mt) * StartPoint.X +
                              3 * (mt * mt) * t * Control1.X +
                              3 * mt * (t * t) * Control2.X +
                              (t * t * t) * EndPoint.X
                    Dim y As Single = (mt * mt * mt) * StartPoint.Y +
                              3 * (mt * mt) * t * Control1.Y +
                              3 * mt * (t * t) * Control2.Y +
                              (t * t * t) * EndPoint.Y
                    Return New PointF(x, y)
                Else
                    Return New PointF(
                StartPoint.X + t * (EndPoint.X - StartPoint.X),
                StartPoint.Y + t * (EndPoint.Y - StartPoint.Y))
                End If
            End Function
        End Class

        ''' <summary>
        ''' Represents a complete path made up of Bézier segments
        ''' </summary>
        Private Class BezierPath
            Public Segments As List(Of BezierSegment)
            Public IsClosed As Boolean
            Public PathIndex As Integer

            Public Sub New()
                Segments = New List(Of BezierSegment)()
                IsClosed = False
            End Sub

            Public Sub New(segments As List(Of BezierSegment), isClosed As Boolean)
                Me.Segments = segments
                Me.IsClosed = isClosed
            End Sub

            ' Returns a bounding box for quick intersection tests
            Public Function GetBoundingBox() As RectangleF
                If Segments.Count = 0 Then
                    Return New RectangleF(0, 0, 0, 0)
                End If

                Dim minX As Single = Single.MaxValue
                Dim minY As Single = Single.MaxValue
                Dim maxX As Single = Single.MinValue
                Dim maxY As Single = Single.MinValue

                For Each segment In Segments
                    ' Check start and end points
                    minX = Math.Min(minX, segment.StartPoint.X)
                    minY = Math.Min(minY, segment.StartPoint.Y)
                    maxX = Math.Max(maxX, segment.StartPoint.X)
                    maxY = Math.Max(maxY, segment.StartPoint.Y)

                    minX = Math.Min(minX, segment.EndPoint.X)
                    minY = Math.Min(minY, segment.EndPoint.Y)
                    maxX = Math.Max(maxX, segment.EndPoint.X)
                    maxY = Math.Max(maxY, segment.EndPoint.Y)

                    ' For Bézier curves, check control points as well
                    If segment.SegmentType = "C"c Then
                        minX = Math.Min(minX, segment.Control1.X)
                        minY = Math.Min(minY, segment.Control1.Y)
                        maxX = Math.Max(maxX, segment.Control1.X)
                        maxY = Math.Max(maxY, segment.Control1.Y)

                        minX = Math.Min(minX, segment.Control2.X)
                        minY = Math.Min(minY, segment.Control2.Y)
                        maxX = Math.Max(maxX, segment.Control2.X)
                        maxY = Math.Max(maxY, segment.Control2.Y)
                    End If
                Next

                Return New RectangleF(minX, minY, maxX - minX, maxY - minY)
            End Function
        End Class

        ''' <summary>
        ''' Represents an intersection between two path segments
        ''' </summary>
        Private Class PathIntersection
            Public IntersectionPoint As PointF
            Public Path1Index As Integer
            Public Path2Index As Integer
            Public Segment1Index As Integer
            Public Segment2Index As Integer
            Public T1 As Double ' Parameter on first segment
            Public T2 As Double ' Parameter on second segment

            Public Sub New()
            End Sub

            Public Sub New(point As PointF, p1Idx As Integer, p2Idx As Integer, s1Idx As Integer, s2Idx As Integer, t1Param As Double, t2Param As Double)
                IntersectionPoint = point
                Path1Index = p1Idx
                Path2Index = p2Idx
                Segment1Index = s1Idx
                Segment2Index = s2Idx
                T1 = t1Param
                T2 = t2Param
            End Sub
        End Class

#End Region

        '===============================================================
        '========== END COMPLETE FIXED MERGING CODE ===================
        '===============================================================


#End Region



    End Class

End Namespace
