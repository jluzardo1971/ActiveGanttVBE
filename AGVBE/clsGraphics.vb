Imports System.Windows.Media.Imaging
Imports System.Diagnostics.Debug

Friend Class clsGraphics

    Private Structure T_PRECT
        Public lLeft As Integer
        Public lTop As Integer
        Public lRight As Integer
        Public lBottom As Integer
    End Structure

    Private mp_udtPreviousClipRegion As T_PRECT

    Private Const SizeOfARGB As Integer = 4

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_bNeedsRendering As Boolean

    Private mp_lFocusLeft As Integer
    Private mp_lFocusTop As Integer
    Private mp_lFocusRight As Integer
    Private mp_lFocusBottom As Integer

    Private mp_aPixels As Integer()
    Friend mp_lClipX1 As Integer
    Friend mp_lClipX2 As Integer
    Friend mp_lClipY1 As Integer
    Friend mp_lClipY2 As Integer

    Private mp_oSelectionLine As Line
    Private mp_oSelectionRectangle As Rectangle

    Private mp_lSelectionRectangleIndex As Integer = -1
    Private mp_lSelectionLineIndex As Integer = -1

    Private mp_bCheckClip As Boolean

    Friend mp_oTextFinalLayout As Rect

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_bNeedsRendering = True
        mp_oSelectionLine = New Line()
        mp_oSelectionRectangle = New Rectangle()
    End Sub

    Private ReadOnly Property GetBitmap() As WriteableBitmap
        Get
            Return mp_oControl.mp_oBitmap
        End Get
    End Property

    Private Sub SetPixel(ByVal lIndex As Integer, ByVal lColor As Integer)
        If mp_bCheckClip = True Then
            Dim lY As Integer
            Dim lX As Integer
            lY = System.Convert.ToInt32(System.Math.Floor(lIndex / Width))
            If (lY) < mp_lClipY1 Then
                Return
            End If
            If (lY) > mp_lClipY2 Then
                Return
            End If
            lX = lIndex - (lY * Width)
            If (lX + 1) <= mp_lClipX1 Or lX >= mp_lClipX2 Then
                Return
            End If
        End If
        If lIndex > (mp_aPixels.Length - 1) Then
            Return
        End If
        mp_aPixels(lIndex) = lColor
    End Sub

    Public ReadOnly Property Width() As Integer
        Get
            If GetBitmap Is Nothing Then
                Return 0
            Else
                Return GetBitmap.PixelWidth
            End If
        End Get
    End Property

    Public ReadOnly Property Height() As Integer
        Get
            If GetBitmap Is Nothing Then
                Return 0
            Else
                Return GetBitmap.PixelHeight
            End If
        End Get
    End Property

    Friend Property f_FocusLeft() As Integer
        Get
            Return mp_lFocusLeft
        End Get
        Set(ByVal Value As Integer)
            mp_lFocusLeft = Value
        End Set
    End Property

    Friend Property f_FocusTop() As Integer
        Get
            Return mp_lFocusTop
        End Get
        Set(ByVal Value As Integer)
            mp_lFocusTop = Value
        End Set
    End Property

    Friend Property f_FocusRight() As Integer
        Get
            Return mp_lFocusRight
        End Get
        Set(ByVal Value As Integer)
            mp_lFocusRight = Value
        End Set
    End Property

    Friend Property f_FocusBottom() As Integer
        Get
            Return mp_lFocusBottom
        End Get
        Set(ByVal Value As Integer)
            mp_lFocusBottom = Value
        End Set
    End Property

    Public Sub StartDrawing()
        If mp_oControl.mp_oBitmap Is Nothing Then
            mp_oControl.mp_oBitmap = New WriteableBitmap(System.Convert.ToInt32(mp_oControl.oCanvas.ActualWidth), System.Convert.ToInt32(mp_oControl.oCanvas.ActualHeight))
            mp_oControl.oImage.Source = mp_oControl.mp_oBitmap
        End If
        If (mp_oControl.mp_oBitmap.PixelWidth <> mp_oControl.oCanvas.ActualWidth) Or (mp_oControl.mp_oBitmap.PixelHeight <> mp_oControl.oCanvas.ActualHeight) Then
            mp_oControl.mp_oBitmap = New WriteableBitmap(System.Convert.ToInt32(mp_oControl.oCanvas.ActualWidth), System.Convert.ToInt32(mp_oControl.oCanvas.ActualHeight))
            mp_oControl.oImage.Source = mp_oControl.mp_oBitmap
        End If
        mp_lClipX1 = 0
        mp_lClipY1 = 0
        mp_lClipX2 = System.Convert.ToInt32(mp_oControl.oCanvas.ActualWidth)
        mp_lClipY2 = System.Convert.ToInt32(mp_oControl.oCanvas.ActualHeight)


    End Sub

    Public Sub TerminateDrawing()
        mp_bNeedsRendering = False
        GetBitmap.Invalidate()
    End Sub

    Public Property NeedsRendering() As Boolean
        Get
            Return mp_bNeedsRendering
        End Get
        Set(ByVal Value As Boolean)
            mp_bNeedsRendering = Value
        End Set
    End Property

    Public Sub Clear(ByVal oColor As Color)
        Dim iColor As Integer = SWM_Color_To_Int32(oColor)
        mp_aPixels = GetBitmap.Pixels
        Dim iPixelCount As Integer = mp_aPixels.Length
        For i As Integer = 0 To iPixelCount - 1
            SetPixel(i, iColor)
        Next
    End Sub

    Private Function SWM_Color_To_Int32(ByVal oColor As Color) As Integer
        Dim iReturn As Integer
        Dim iA As Integer = oColor.A
        Dim iR As Integer = oColor.R
        Dim iG As Integer = oColor.G
        Dim iB As Integer = oColor.B
        iReturn = (iA << 24) Or (iR << 16) Or (iG << 8) Or iB
        Return iReturn
    End Function

    Private Sub mp_DrawLineDDA(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, oColor As Color)
        Dim lColor As Integer = SWM_Color_To_Int32(oColor)
        Dim w As Integer = GetBitmap.PixelWidth
        mp_aPixels = GetBitmap.Pixels
        Dim dx As Integer = x2 - x1
        Dim dy As Integer = y2 - y1
        Dim len As Integer = If(dy >= 0, dy, -dy)
        Dim lenx As Integer = If(dx >= 0, dx, -dx)
        If lenx > len Then
            len = lenx
        End If
        If len <> 0 Then
            Dim incx As Single = dx / CSng(len)
            Dim incy As Single = dy / CSng(len)
            Dim x As Single = x1
            Dim y As Single = y1
            For i As Integer = 0 To len - 1
                SetPixel(CInt(y) * w + CInt(x), lColor)
                x += incx
                y += incy
            Next
        End If
    End Sub

    Private Sub mp_FillRectangle(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, oColor As Color)
        Dim lColor As Integer = SWM_Color_To_Int32(oColor)
        Dim w As Integer = GetBitmap.PixelWidth
        Dim h As Integer = GetBitmap.PixelHeight
        mp_aPixels = GetBitmap.Pixels

        x2 = x2 + 1
        y2 = y2 + 1

        ' Check boundaries
        If x1 < 0 Then
            x1 = 0
        End If
        If y1 < 0 Then
            y1 = 0
        End If
        If x2 < 0 Then
            x2 = 0
        End If
        If y2 < 0 Then
            y2 = 0
        End If
        If x1 >= w Then
            x1 = w - 1
        End If
        If y1 >= h Then
            y1 = h - 1
        End If
        If x2 >= w Then
            x2 = w - 1
        End If
        If y2 >= h Then
            y2 = h - 1
        End If


        ' Fill first line
        Dim startY As Integer = y1 * w
        Dim startYPlusX1 As Integer = startY + x1
        Dim endOffset As Integer = startY + x2
        For x As Integer = startYPlusX1 To endOffset - 1
            SetPixel(x, lColor)
        Next

        ' Copy first line
        Dim len As Integer = (x2 - x1) * SizeOfARGB
        Dim srcOffsetBytes As Integer = startYPlusX1 * SizeOfARGB
        Dim offset2 As Integer = y2 * w + x1
        Dim y As Integer = startYPlusX1 + w
        While y < offset2
            Buffer.BlockCopy(mp_aPixels, srcOffsetBytes, mp_aPixels, y * SizeOfARGB, len)
            y += w
        End While
    End Sub

    Private Sub mp_DrawRectangle(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, oColor As Color)
        Dim lColor As Integer = SWM_Color_To_Int32(oColor)
        Dim w As Integer = GetBitmap.PixelWidth
        mp_aPixels = GetBitmap.Pixels
        Dim startY As Integer = y1 * w
        Dim endY As Integer = y2 * w
        Dim offset2 As Integer = endY + x1
        Dim endOffset As Integer = startY + x2
        Dim startYPlusX1 As Integer = startY + x1
        For x As Integer = startYPlusX1 To endOffset
            SetPixel(x, lColor)
            SetPixel(offset2, lColor)
            offset2 += 1
        Next
        endOffset = startYPlusX1 + w
        offset2 -= w
        Dim y As Integer = startY + x2 + w
        While y < offset2
            SetPixel(y, lColor)
            SetPixel(endOffset, lColor)
            endOffset += w
            y += w
        End While
    End Sub

    Private Sub mp_DrawPolyline(ByVal points As Point(), oColor As Color)
        Dim x1 As Integer = System.Convert.ToInt32(points(0).X)
        Dim y1 As Integer = System.Convert.ToInt32(points(0).Y)
        Dim x2 As Integer, y2 As Integer
        For i As Integer = 1 To points.Length - 1
            x2 = System.Convert.ToInt32(points(i).X)
            y2 = System.Convert.ToInt32(points(i).Y)
            mp_DrawLine(x1, y1, x2, y2, oColor)
            x1 = x2
            y1 = y2
        Next
    End Sub

    Private Sub mp_DrawEllipseCentered(xc As Integer, yc As Integer, xr As Integer, yr As Integer, oColor As Color)
        Dim lColor As Integer = SWM_Color_To_Int32(oColor)
        Dim w As Integer = GetBitmap.PixelWidth
        mp_aPixels = GetBitmap.Pixels

        ' Init vars
        Dim uh As Integer, lh As Integer
        Dim x As Integer = xr
        Dim y As Integer = 0
        Dim xrSqTwo As Integer = (xr * xr) << 1
        Dim yrSqTwo As Integer = (yr * yr) << 1
        Dim xChg As Integer = yr * yr * (1 - (xr << 1))
        Dim yChg As Integer = xr * xr
        Dim err As Integer = 0
        Dim xStopping As Integer = yrSqTwo * xr
        Dim yStopping As Integer = 0

        ' Draw first set of points counter clockwise where tangent line slope > -1.
        While xStopping >= yStopping
            ' Draw 4 quadrant points at once
            uh = (yc + y) * w
            ' Upper half
            lh = (yc - y) * w
            ' Lower half
            SetPixel(xc + x + uh, lColor)
            ' Quadrant I
            SetPixel(xc - x + uh, lColor)
            ' Quadrant II
            SetPixel(xc - x + lh, lColor)
            ' Quadrant III
            SetPixel(xc + x + lh, lColor)
            ' Quadrant IV
            y += 1
            yStopping += xrSqTwo
            err += yChg
            yChg += xrSqTwo
            If (xChg + (err << 1)) > 0 Then
                x -= 1
                xStopping -= yrSqTwo
                err += xChg
                xChg += yrSqTwo
            End If
        End While

        ' ReInit vars
        x = 0
        y = yr
        uh = (yc + y) * w
        ' Upper half
        lh = (yc - y) * w
        ' Lower half
        xChg = yr * yr
        yChg = xr * xr * (1 - (yr << 1))
        err = 0
        xStopping = 0
        yStopping = xrSqTwo * yr

        ' Draw second set of points clockwise where tangent line slope < -1.
        While xStopping <= yStopping
            ' Draw 4 quadrant points at once
            SetPixel(xc + x + uh, lColor)
            ' Quadrant I
            SetPixel(xc - x + uh, lColor)
            ' Quadrant II
            SetPixel(xc - x + lh, lColor)
            ' Quadrant III
            SetPixel(xc + x + lh, lColor)
            ' Quadrant IV
            x += 1
            xStopping += yrSqTwo
            err += xChg
            xChg += yrSqTwo
            If (yChg + (err << 1)) > 0 Then
                y -= 1
                uh = (yc + y) * w
                ' Upper half
                lh = (yc - y) * w
                ' Lower half
                yStopping -= xrSqTwo
                err += yChg
                yChg += xrSqTwo
            End If
        End While
    End Sub

    Private Sub mp_FillEllipseCentered(xc As Integer, yc As Integer, xw As Integer, yw As Integer, oColor As Color)
        Dim lColor As Integer = SWM_Color_To_Int32(oColor)
        mp_aPixels = GetBitmap.Pixels
        Dim w As Integer = GetBitmap.PixelWidth
        Dim h As Integer = GetBitmap.PixelHeight
        Dim xr As Integer = System.Convert.ToInt32(xw / 2)
        Dim yr As Integer = System.Convert.ToInt32(yw / 2)

        ' Init vars
        Dim uh As Integer, lh As Integer, uy As Integer, ly As Integer, lx As Integer, rx As Integer
        Dim x As Integer = xr
        Dim y As Integer = 0
        Dim xrSqTwo As Integer = (xr * xr) << 1
        Dim yrSqTwo As Integer = (yr * yr) << 1
        Dim xChg As Integer = yr * yr * (1 - (xr << 1))
        Dim yChg As Integer = xr * xr
        Dim err As Integer = 0
        Dim xStopping As Integer = yrSqTwo * xr
        Dim yStopping As Integer = 0

        ' Draw first set of points counter clockwise where tangent line slope > -1.
        While xStopping >= yStopping
            ' Draw 4 quadrant points at once
            uy = yc + y
            ' Upper half
            ly = yc - y
            ' Lower half
            If uy < 0 Then
                uy = 0
            End If
            ' Clip
            If uy >= h Then
                uy = h - 1
            End If
            ' ...
            If ly < 0 Then
                ly = 0
            End If
            If ly >= h Then
                ly = h - 1
            End If
            uh = uy * w
            ' Upper half
            lh = ly * w
            ' Lower half
            rx = xc + x
            lx = xc - x
            If rx < 0 Then
                rx = 0
            End If
            ' Clip
            If rx >= w Then
                rx = w - 1
            End If
            ' ...
            If lx < 0 Then
                lx = 0
            End If
            If lx >= w Then
                lx = w - 1
            End If

            ' Draw line
            For i As Integer = lx To rx
                SetPixel(i + uh, lColor)
                ' Quadrant II to I (Actually two octants)
                ' Quadrant III to IV
                SetPixel(i + lh, lColor)
            Next

            y += 1
            yStopping += xrSqTwo
            err += yChg
            yChg += xrSqTwo
            If (xChg + (err << 1)) > 0 Then
                x -= 1
                xStopping -= yrSqTwo
                err += xChg
                xChg += yrSqTwo
            End If
        End While

        ' ReInit vars
        x = 0
        y = yr
        uy = yc + y
        ' Upper half
        ly = yc - y
        ' Lower half
        If uy < 0 Then
            uy = 0
        End If
        ' Clip
        If uy >= h Then
            uy = h - 1
        End If
        ' ...
        If ly < 0 Then
            ly = 0
        End If
        If ly >= h Then
            ly = h - 1
        End If
        uh = uy * w
        ' Upper half
        lh = ly * w
        ' Lower half
        xChg = yr * yr
        yChg = xr * xr * (1 - (yr << 1))
        err = 0
        xStopping = 0
        yStopping = xrSqTwo * yr

        ' Draw second set of points clockwise where tangent line slope < -1.
        While xStopping <= yStopping
            ' Draw 4 quadrant points at once
            rx = xc + x
            lx = xc - x
            If rx < 0 Then
                rx = 0
            End If
            ' Clip
            If rx >= w Then
                rx = w - 1
            End If
            ' ...
            If lx < 0 Then
                lx = 0
            End If
            If lx >= w Then
                lx = w - 1
            End If

            ' Draw line
            For i As Integer = lx To rx
                SetPixel(i + uh, lColor)
                ' Quadrant II to I (Actually two octants)
                ' Quadrant III to IV
                SetPixel(i + lh, lColor)
            Next

            x += 1
            xStopping += yrSqTwo
            err += xChg
            xChg += yrSqTwo
            If (yChg + (err << 1)) > 0 Then
                y -= 1
                uy = yc + y
                ' Upper half
                ly = yc - y
                ' Lower half
                If uy < 0 Then
                    uy = 0
                End If
                ' Clip
                If uy >= h Then
                    uy = h - 1
                End If
                ' ...
                If ly < 0 Then
                    ly = 0
                End If
                If ly >= h Then
                    ly = h - 1
                End If
                uh = uy * w
                ' Upper half
                lh = ly * w
                ' Lower half
                yStopping -= xrSqTwo
                err += yChg
                yChg += xrSqTwo
            End If
        End While
    End Sub

    Private Sub mp_FillPolygon(points As Point(), oColor As Color)
        Dim lColor As Integer = SWM_Color_To_Int32(oColor)
        Dim w As Integer = GetBitmap.PixelWidth
        Dim h As Integer = GetBitmap.PixelHeight
        mp_aPixels = GetBitmap.Pixels
        Dim pn As Integer = points.Length
        Dim pnh As Integer = points.Length >> 1
        Dim intersectionsX As Integer() = New Integer(pnh) {}
        Dim yMin As Integer = h
        Dim yMax As Integer = 0
        For i As Integer = 0 To pn - 1
            Dim py As Integer = System.Convert.ToInt32(points(i).Y)
            If py < yMin Then
                yMin = py
            End If
            If py > yMax Then
                yMax = py
            End If
        Next
        If yMin < 0 Then
            yMin = 0
        End If
        If yMax >= h Then
            yMax = h - 1
        End If
        For y As Integer = yMin To yMax
            Dim vxi As Single = System.Convert.ToSingle(points(0).X)
            Dim vyi As Single = System.Convert.ToSingle(points(0).Y)
            ' Based on http://alienryderflex.com/polygon_fill/
            Dim intersectionCount As Integer = 0
            For i As Integer = 1 To pn - 1
                ' Next point x, y
                Dim vxj As Single = System.Convert.ToSingle(points(i).X)
                Dim vyj As Single = System.Convert.ToSingle(points(i).Y)

                ' Is the scanline between the two points
                If vyi < y AndAlso vyj >= y OrElse vyj < y AndAlso vyi >= y Then
                    intersectionsX(intersectionCount) = CInt((vxi + (y - vyi) / (vyj - vyi) * (vxj - vxi)))
                    intersectionCount = intersectionCount + 1
                End If
                vxi = vxj
                vyi = vyj
            Next

            ' Sort the intersections from left to right using Insertion sort 
            ' It's faster than Array.Sort for this small data set
            Dim t As Integer, j As Integer
            For i As Integer = 1 To intersectionCount - 1
                t = intersectionsX(i)
                j = i
                While j > 0 AndAlso intersectionsX(j - 1) > t
                    intersectionsX(j) = intersectionsX(j - 1)
                    j = j - 1
                End While
                intersectionsX(j) = t
            Next

            ' Fill the pixels between the intersections
            For i As Integer = 0 To intersectionCount - 1
                Dim x0 As Integer = intersectionsX(i)
                Dim x1 As Integer = intersectionsX(i + 1)

                ' Check boundary
                If x1 > 0 AndAlso x0 < w Then
                    If x0 < 0 Then
                        x0 = 0
                    End If
                    If x1 >= w Then
                        x1 = w - 1
                    End If

                    ' Fill the pixels
                    For x As Integer = x0 To x1
                        SetPixel(y * w + x, lColor)
                    Next
                End If
            Next
        Next
    End Sub

    Private Sub mp_DrawLine(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, oColor As Color)
        Dim lColor As Integer = SWM_Color_To_Int32(oColor)
        Dim w As Integer = GetBitmap.PixelWidth
        mp_aPixels = GetBitmap.Pixels
        Dim dx As Integer = x2 - x1
        Dim dy As Integer = y2 - y1
        Const PRECISION_SHIFT As Integer = 8
        Const PRECISION_VALUE As Integer = 1 << PRECISION_SHIFT
        Dim lenX As Integer, lenY As Integer
        Dim incy1 As Integer
        If dy >= 0 Then
            incy1 = PRECISION_VALUE
            lenY = dy
        Else
            incy1 = -PRECISION_VALUE
            lenY = -dy
        End If
        Dim incx1 As Integer
        If dx >= 0 Then
            incx1 = 1
            lenX = dx
        Else
            incx1 = -1
            lenX = -dx
        End If
        If lenX > lenY Then
            Dim incy As Integer = (dy << PRECISION_SHIFT) \ lenX
            Dim y As Integer = y1 << PRECISION_SHIFT
            For i As Integer = 0 To lenX - 1
                SetPixel((y >> PRECISION_SHIFT) * w + x1, lColor)
                x1 += incx1
                y += incy
            Next
        Else
            If lenY = 0 Then
                Return
            End If
            Dim incx As Integer = (dx << PRECISION_SHIFT) \ lenY
            Dim lIndex As Integer = (x1 + y1 * w) << PRECISION_SHIFT
            Dim inc As Integer = incy1 * w + incx
            Dim lArrayPosIndex As Integer = 0
            Dim lArrayLength As Integer = mp_aPixels.Length
            For i As Integer = 0 To lenY - 1
                lArrayPosIndex = (lIndex >> PRECISION_SHIFT)
                If lArrayPosIndex < lArrayLength And lArrayPosIndex >= 0 Then
                    SetPixel(lArrayPosIndex, lColor)
                End If
                '// Their Code
                'pixels(index >> PRECISION_SHIFT) = lColor
                lIndex += inc
            Next
        End If
    End Sub

    '// Public Functions

    Public Function PolygonInsideCanvas(ByVal oPoints As Point()) As Boolean
        Dim lLength = oPoints.Length
        Dim i As Integer
        Dim bReturn As Boolean = False
        Dim lPointsOutside As Integer = 0
        For i = 0 To lLength - 1
            If oPoints(i).X > mp_lClipX1 And oPoints(i).X < mp_lClipX2 Then
                If oPoints(i).Y > mp_lClipY1 And oPoints(i).Y < mp_lClipY2 Then
                    bReturn = True
                Else
                    lPointsOutside = lPointsOutside + 1
                End If
            Else
                lPointsOutside = lPointsOutside + 1
            End If
        Next
        If lPointsOutside > 0 Then
            mp_bCheckClip = True
        Else
            mp_bCheckClip = False
        End If
        Return bReturn
    End Function

    Public Sub DrawPolygon(ByVal oColor As Color, ByRef oPoints() As Point)
        mp_DrawPolyline(oPoints, oColor)
        GetBitmap.Invalidate()
    End Sub

    Public Sub DrawEdge(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal clrBackColor As Color, ByVal v_yButtonStyle As GRE_BUTTONSTYLE, ByVal v_lEdgeType As GRE_EDGETYPE, ByVal v_bFilled As Boolean, ByVal oStyle As clsStyle)
        Dim lExteriorLeftTopColor As Color
        Dim lInteriorLeftTopColor As Color
        Dim lExteriorRightBottomColor As Color
        Dim lInteriorRightBottomColor As Color
        If v_yButtonStyle = GRE_BUTTONSTYLE.BT_NORMALWINDOWS Then
            Select Case v_lEdgeType
                Case GRE_EDGETYPE.ET_RAISED
                    If oStyle Is Nothing Then
                        lExteriorLeftTopColor = Color.FromArgb(255, 240, 240, 240)
                        lInteriorLeftTopColor = Color.FromArgb(255, 192, 192, 192)
                        lInteriorRightBottomColor = Colors.Gray
                        lExteriorRightBottomColor = Color.FromArgb(255, 64, 64, 64)
                    Else
                        lExteriorLeftTopColor = oStyle.ButtonBorderStyle.RaisedExteriorLeftTopColor
                        lInteriorLeftTopColor = oStyle.ButtonBorderStyle.RaisedInteriorLeftTopColor
                        lInteriorRightBottomColor = oStyle.ButtonBorderStyle.RaisedInteriorRightBottomColor
                        lExteriorRightBottomColor = oStyle.ButtonBorderStyle.RaisedExteriorRightBottomColor
                    End If
                Case GRE_EDGETYPE.ET_SUNKEN
                    If oStyle Is Nothing Then
                        lExteriorLeftTopColor = Colors.Gray
                        lInteriorLeftTopColor = Color.FromArgb(255, 64, 64, 64)
                        lInteriorRightBottomColor = Color.FromArgb(255, 192, 192, 192)
                        lExteriorRightBottomColor = Color.FromArgb(255, 240, 240, 240)
                    Else
                        lExteriorLeftTopColor = oStyle.ButtonBorderStyle.SunkenExteriorLeftTopColor
                        lInteriorLeftTopColor = oStyle.ButtonBorderStyle.SunkenInteriorLeftTopColor
                        lInteriorRightBottomColor = oStyle.ButtonBorderStyle.SunkenInteriorRightBottomColor
                        lExteriorRightBottomColor = oStyle.ButtonBorderStyle.SunkenExteriorRightBottomColor
                    End If
            End Select
            '// Exterior Left
            DrawLine(v_X1, v_Y1, v_X1, v_Y2, GRE_LINETYPE.LT_NORMAL, lExteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Exterior Top
            DrawLine(v_X1, v_Y1, v_X2, v_Y1, GRE_LINETYPE.LT_NORMAL, lExteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Exterior Right
            DrawLine(v_X2, v_Y2, v_X2, v_Y1, GRE_LINETYPE.LT_NORMAL, lExteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Exterior Bottom
            DrawLine(v_X1, v_Y2, v_X2, v_Y2, GRE_LINETYPE.LT_NORMAL, lExteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Interior Left
            DrawLine(v_X1 + 1, v_Y1 + 1, v_X1 + 1, v_Y2 - 1, GRE_LINETYPE.LT_NORMAL, lInteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Interior Top
            DrawLine(v_X1 + 1, v_Y1 + 1, v_X2 - 1, v_Y1 + 1, GRE_LINETYPE.LT_NORMAL, lInteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Interior Right
            DrawLine(v_X2 - 1, v_Y2 - 1, v_X2 - 1, v_Y1 + 1, GRE_LINETYPE.LT_NORMAL, lInteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Interior Bottom
            DrawLine(v_X1 + 1, v_Y2 - 1, v_X2 - 1, v_Y2 - 1, GRE_LINETYPE.LT_NORMAL, lInteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            If v_bFilled = True Then
                DrawLine(v_X1 + 2, v_Y1 + 2, v_X2 - 2, v_Y2 - 2, GRE_LINETYPE.LT_FILLED, clrBackColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            End If
        Else
            Select Case v_lEdgeType
                Case GRE_EDGETYPE.ET_RAISED
                    If oStyle Is Nothing Then
                        lExteriorLeftTopColor = Colors.White
                        lExteriorRightBottomColor = Color.FromArgb(255, 64, 64, 64)
                    Else
                        lExteriorLeftTopColor = oStyle.ButtonBorderStyle.RaisedExteriorLeftTopColor
                        lExteriorRightBottomColor = oStyle.ButtonBorderStyle.RaisedExteriorRightBottomColor
                    End If
                Case GRE_EDGETYPE.ET_SUNKEN
                    If oStyle Is Nothing Then
                        lExteriorLeftTopColor = Colors.Gray
                        lExteriorRightBottomColor = Color.FromArgb(255, 255, 255, 255)
                    Else
                        lExteriorLeftTopColor = oStyle.ButtonBorderStyle.SunkenExteriorLeftTopColor
                        lExteriorRightBottomColor = oStyle.ButtonBorderStyle.SunkenExteriorRightBottomColor
                    End If
            End Select
            DrawLine(v_X1, v_Y1, v_X2, v_Y1, GRE_LINETYPE.LT_NORMAL, lExteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            DrawLine(v_X1, v_Y1, v_X1, v_Y2, GRE_LINETYPE.LT_NORMAL, lExteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            DrawLine(v_X1, v_Y2, v_X2, v_Y2, GRE_LINETYPE.LT_NORMAL, lExteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            DrawLine(v_X2, v_Y2, v_X2, v_Y1 - 1, GRE_LINETYPE.LT_NORMAL, lExteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            If v_bFilled = True Then
                DrawLine(v_X1 + 1, v_Y1 + 1, v_X2 - 1, v_Y2 - 1, GRE_LINETYPE.LT_FILLED, clrBackColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            End If
        End If
    End Sub

    Public Sub DrawLine(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_yStyle As GRE_LINETYPE, ByVal oColor As Color, ByVal v_lDrawStyle As GRE_LINEDRAWSTYLE)
        DrawLine(v_X1, v_Y1, v_X2, v_Y2, v_yStyle, oColor, v_lDrawStyle, 1, True)
    End Sub

    Public Sub DrawLine(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_yStyle As GRE_LINETYPE, ByVal oColor As Color, ByVal v_lDrawStyle As GRE_LINEDRAWSTYLE, ByVal v_lWidth As Integer)
        DrawLine(v_X1, v_Y1, v_X2, v_Y2, v_yStyle, oColor, v_lDrawStyle, v_lWidth, True)
    End Sub

    Public Sub DrawLine(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_yStyle As GRE_LINETYPE, ByVal oColor As Color, ByVal v_lDrawStyle As GRE_LINEDRAWSTYLE, ByVal v_lWidth As Integer, ByVal v_bCreatePens As Boolean)
        If (v_X1 < mp_lClipX1 And v_X2 < mp_lClipX1) Or (v_X1 > mp_lClipX2 And v_X2 > mp_lClipX2) Then
            Return
        End If
        If (v_Y1 < mp_lClipY1 And v_Y2 < mp_lClipY1) Or (v_Y1 > mp_lClipY2 And v_Y2 > mp_lClipY2) Then
            Return
        End If
        CorrectRectCoords(v_X1, v_Y1, v_X2, v_Y2)

        'Select Case v_lDrawStyle
        '    Case GRE_LINEDRAWSTYLE.LDS_SOLID
        '        mp_ucPen.DashStyle = Drawing.Drawing2D.DashStyle.Solid
        '    Case GRE_LINEDRAWSTYLE.LDS_DOT
        '        mp_ucPen.DashStyle = Drawing.Drawing2D.DashStyle.Dot
        'End Select
        Select Case v_yStyle
            Case GRE_LINETYPE.LT_NORMAL
                If v_Y1 = v_Y2 Then
                    v_X2 = v_X2 + 1
                ElseIf v_X1 = v_X2 Then
                    v_Y2 = v_Y2 + 1
                End If
                mp_DrawLineDDA(v_X1, v_Y1, v_X2, v_Y2, oColor)
            Case GRE_LINETYPE.LT_BORDER
                DrawLine(v_X1, v_Y1, v_X2, v_Y1, GRE_LINETYPE.LT_NORMAL, oColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(v_X1, v_Y2, v_X2, v_Y2, GRE_LINETYPE.LT_NORMAL, oColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(v_X1, v_Y1, v_X1, v_Y2, GRE_LINETYPE.LT_NORMAL, oColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(v_X2, v_Y1, v_X2, v_Y2, GRE_LINETYPE.LT_NORMAL, oColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            Case GRE_LINETYPE.LT_FILLED
                mp_FillRectangle(v_X1, v_Y1, v_X2, v_Y2, oColor)
        End Select
        GetBitmap.Invalidate()
    End Sub

    Private Sub mp_InsideClipRegion(ByVal X As Integer, ByVal Y As Integer)
        If X > mp_lClipX1 And X < mp_lClipX2 And Y > mp_lClipY1 And Y < mp_lClipY2 Then
            mp_bCheckClip = False
        Else
            mp_bCheckClip = True
        End If
    End Sub

    Public Sub DrawFigure(ByVal X As Integer, ByVal Y As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal yFigureType As GRE_FIGURETYPE, ByVal oBorderColor As Color, ByVal oFillColor As Color, ByVal yBorderStyle As GRE_LINEDRAWSTYLE)
        mp_InsideClipRegion(X, Y)
        If dx Mod 2 <> 0 Then
            dx = dx + 1
            dy = dy + 1
        End If
        Select Case yFigureType
            Case GRE_FIGURETYPE.FT_PROJECTUP
                Dim Points(5) As Point
                Points(0).X = X
                Points(0).Y = Y
                Points(1).X = X + dx / 2
                Points(1).Y = Y + dy / 2
                Points(2).X = X + dx / 2
                Points(2).Y = Y + dy
                Points(3).X = X - dx / 2
                Points(3).Y = Y + dy
                Points(4).X = X - dx / 2
                Points(4).Y = Y + dy / 2
                Points(5).X = Points(0).X
                Points(5).Y = Points(0).Y
                mp_DrawFigureAux(oFillColor, oBorderColor, Points)
            Case GRE_FIGURETYPE.FT_PROJECTDOWN
                Dim Points(5) As Point
                Points(0).X = X + dx / 2
                Points(0).Y = Y
                Points(1).X = X + dx / 2
                Points(1).Y = Y + dy / 2
                Points(2).X = X
                Points(2).Y = Y + dy
                Points(3).X = X - dx / 2
                Points(3).Y = Y + dy / 2
                Points(4).X = X - dx / 2
                Points(4).Y = Y
                Points(5).X = Points(0).X
                Points(5).Y = Points(0).Y
                mp_DrawFigureAux(oFillColor, oBorderColor, Points)
            Case GRE_FIGURETYPE.FT_DIAMOND
                Dim Points(4) As Point
                Points(0).X = X
                Points(0).Y = Y
                Points(1).X = X + dx / 2
                Points(1).Y = Y + dy / 2
                Points(2).X = X
                Points(2).Y = Y + dy
                Points(3).X = X - dx / 2
                Points(3).Y = Y + dy / 2
                Points(4).X = Points(0).X
                Points(4).Y = Points(0).Y
                mp_DrawFigureAux(oFillColor, oBorderColor, Points)
            Case GRE_FIGURETYPE.FT_CIRCLEDIAMOND
                Dim Points(4) As Point
                Points(0).X = X
                Points(0).Y = Y + dy / 4
                Points(1).X = X + dx / 4
                Points(1).Y = Y + dy / 2
                Points(2).X = X
                Points(2).Y = Y + (3 * dy) / 4
                Points(3).X = X - dx / 4
                Points(3).Y = Y + dy / 2
                Points(4).X = Points(0).X
                Points(4).Y = Points(0).Y
                mp_DrawEllipseCentered(X, System.Convert.ToInt32(Y + dy / 2), System.Convert.ToInt32(dx / 2), System.Convert.ToInt32(dy / 2), oFillColor)
                mp_DrawFigureAux(oFillColor, oBorderColor, Points)
            Case GRE_FIGURETYPE.FT_TRIANGLEUP
                Dim Points(3) As Point
                Points(0).X = X
                Points(0).Y = Y
                Points(1).X = X + dx / 2
                Points(1).Y = Y + dy
                Points(2).X = X - dx / 2
                Points(2).Y = Y + dy
                Points(3).X = Points(0).X
                Points(3).Y = Points(0).Y
                mp_DrawFigureAux(oFillColor, oBorderColor, Points)
            Case GRE_FIGURETYPE.FT_TRIANGLEDOWN
                Dim Points(3) As Point
                Points(0).X = X + dx / 2
                Points(0).Y = Y
                Points(1).X = X - dx / 2
                Points(1).Y = Y
                Points(2).X = X
                Points(2).Y = Y + dy
                Points(3).X = Points(0).X
                Points(3).Y = Points(0).Y
                mp_DrawFigureAux(oFillColor, oBorderColor, Points)
            Case GRE_FIGURETYPE.FT_TRIANGLERIGHT
                Dim Points(3) As Point
                Points(0).X = X
                Points(0).Y = Y
                Points(1).X = X
                Points(1).Y = Y + dy
                Points(2).X = X + dx
                Points(2).Y = Y + dy / 2
                Points(3).X = Points(0).X
                Points(3).Y = Points(0).Y
                mp_DrawFigureAux(oFillColor, oBorderColor, Points)
            Case GRE_FIGURETYPE.FT_TRIANGLELEFT
                Dim Points(3) As Point
                Points(0).X = X
                Points(0).Y = Y
                Points(1).X = X
                Points(1).Y = Y + dy
                Points(2).X = X - dx
                Points(2).Y = Y + dy / 2
                Points(3).X = Points(0).X
                Points(3).Y = Points(0).Y
                mp_DrawFigureAux(oFillColor, oBorderColor, Points)
            Case GRE_FIGURETYPE.FT_CIRCLETRIANGLEUP
                Dim Points(3) As Point
                Points(0).X = X
                Points(0).Y = Y + dy / 4
                Points(1).X = X + dx / 4
                Points(1).Y = Y + (3 * dy) / 4
                Points(2).X = X - dx / 4
                Points(2).Y = Y + (3 * dy) / 4
                Points(3).X = Points(0).X
                Points(3).Y = Points(0).Y
                mp_DrawEllipseCentered(X, System.Convert.ToInt32(Y + dy / 2), System.Convert.ToInt32(dx / 2), System.Convert.ToInt32(dy / 2), oFillColor)
                mp_DrawFigureAux(oFillColor, oBorderColor, Points)
            Case GRE_FIGURETYPE.FT_CIRCLETRIANGLEDOWN
                Dim Points(3) As Point
                Points(0).X = X - dx / 4
                Points(0).Y = Y + dy / 4
                Points(1).X = X + dx / 4
                Points(1).Y = Y + dy / 4
                Points(2).X = X
                Points(2).Y = Y + (3 * dy) / 4
                Points(3).X = Points(0).X
                Points(3).Y = Points(0).Y
                mp_DrawEllipseCentered(X, System.Convert.ToInt32(Y + dy / 2), System.Convert.ToInt32(dx / 2), System.Convert.ToInt32(dy / 2), oFillColor)
                mp_DrawFigureAux(oFillColor, oBorderColor, Points)
            Case GRE_FIGURETYPE.FT_ARROWUP
                Dim Points(7) As Point
                Points(0).X = X
                Points(0).Y = Y
                Points(1).X = X + dx / 2
                Points(1).Y = Y + dy / 2
                Points(2).X = X + dx / 4
                Points(2).Y = Y + dy / 2
                Points(3).X = X + dx / 4
                Points(3).Y = Y + dy
                Points(4).X = X - dx / 4
                Points(4).Y = Y + dy
                Points(5).X = X - dx / 4
                Points(5).Y = Y + dy / 2
                Points(6).X = X - dx / 2
                Points(6).Y = Y + dy / 2
                Points(7).X = Points(0).X
                Points(7).Y = Points(0).Y
                mp_DrawFigureAux(oFillColor, oBorderColor, Points)
            Case GRE_FIGURETYPE.FT_ARROWDOWN
                Dim Points(7) As Point
                Points(0).X = X - dx / 4
                Points(0).Y = Y
                Points(1).X = X + dx / 4
                Points(1).Y = Y
                Points(2).X = X + dx / 4
                Points(2).Y = Y + dy / 2
                Points(3).X = X + dx / 2
                Points(3).Y = Y + dy / 2
                Points(4).X = X
                Points(4).Y = Y + dy
                Points(5).X = X - dx / 2
                Points(5).Y = Y + dy / 2
                Points(6).X = X - dx / 4
                Points(6).Y = Y + dy / 2
                Points(7).X = Points(0).X
                Points(7).Y = Points(0).Y
                mp_DrawFigureAux(oFillColor, oBorderColor, Points)
            Case GRE_FIGURETYPE.FT_CIRCLEARROWUP
                Dim Points(7) As Point
                Points(0).X = X
                Points(0).Y = Y + dy / 4
                Points(1).X = X + dx / 4
                Points(1).Y = Y + dy / 2
                Points(2).X = X + dx / 8
                Points(2).Y = Y + dy / 2
                Points(3).X = X + dx / 8
                Points(3).Y = Y + (3 * dy) / 4
                Points(4).X = X - dx / 8
                Points(4).Y = Y + (3 * dy) / 4
                Points(5).X = X - dx / 8
                Points(5).Y = Y + dy / 2
                Points(6).X = X - dx / 4
                Points(6).Y = Y + dy / 2
                Points(7).X = Points(0).X
                Points(7).Y = Points(0).Y
                mp_DrawEllipseCentered(X, System.Convert.ToInt32(Y + dy / 2), System.Convert.ToInt32(dx / 2), System.Convert.ToInt32(dy / 2), oFillColor)
                mp_DrawFigureAux(oFillColor, oBorderColor, Points)
            Case GRE_FIGURETYPE.FT_CIRCLEARROWDOWN
                Dim Points(7) As Point
                Points(0).X = X - dx / 8
                Points(0).Y = Y + dy / 4
                Points(1).X = X + dx / 8
                Points(1).Y = Y + dy / 4
                Points(2).X = X + dx / 8
                Points(2).Y = Y + dy / 2
                Points(3).X = X + dx / 4
                Points(3).Y = Y + dy / 2
                Points(4).X = X
                Points(4).Y = Y + (3 * dy) / 4
                Points(5).X = X - dx / 4
                Points(5).Y = Y + dy / 2
                Points(6).X = X - dx / 8
                Points(6).Y = Y + dy / 2
                Points(7).X = Points(0).X
                Points(7).Y = Points(0).Y
                mp_DrawEllipseCentered(X, System.Convert.ToInt32(Y + dy / 2), System.Convert.ToInt32(dx / 2), System.Convert.ToInt32(dy / 2), oFillColor)
                mp_DrawFigureAux(oFillColor, oBorderColor, Points)
            Case GRE_FIGURETYPE.FT_SMALLPROJECTUP
                Dim Points(5) As Point
                Points(0).X = X
                Points(0).Y = Y + dy / 2
                Points(1).X = X + dx / 4
                Points(1).Y = Y + (3 * dy) / 4
                Points(2).X = X + dx / 4
                Points(2).Y = Y + dy
                Points(3).X = X - dx / 4
                Points(3).Y = Y + dy
                Points(4).X = X - dx / 4
                Points(4).Y = Y + (3 * dy) / 4
                Points(5).X = Points(0).X
                Points(5).Y = Points(0).Y
                mp_DrawFigureAux(oFillColor, oBorderColor, Points)
            Case GRE_FIGURETYPE.FT_SMALLPROJECTDOWN
                Dim Points(5) As Point
                Points(0).X = X + dx / 4
                Points(0).Y = Y
                Points(1).X = X + dx / 4
                Points(1).Y = Y + dy / 4
                Points(2).X = X
                Points(2).Y = Y + dy / 2
                Points(3).X = X - dx / 4
                Points(3).Y = Y + dy / 4
                Points(4).X = X - dx / 4
                Points(4).Y = Y
                Points(5).X = Points(0).X
                Points(5).Y = Points(0).Y
                mp_DrawFigureAux(oFillColor, oBorderColor, Points)
            Case GRE_FIGURETYPE.FT_RECTANGLE
                Dim Points(4) As Point
                Points(0).X = X - dx / 8
                Points(0).Y = Y
                Points(1).X = X + dx / 8
                Points(1).Y = Y
                Points(2).X = X + dx / 8
                Points(2).Y = Y + dy
                Points(3).X = X - dx / 8
                Points(3).Y = Y + dy
                Points(4).X = Points(0).X
                Points(4).Y = Points(0).Y
                mp_DrawFigureAux(oFillColor, oBorderColor, Points)
            Case GRE_FIGURETYPE.FT_SQUARE
                Dim Points(4) As Point
                Points(0).X = X - dx / 4
                Points(0).Y = Y + dx / 4
                Points(1).X = X + dx / 4
                Points(1).Y = Y + dx / 4
                Points(2).X = X + dx / 4
                Points(2).Y = Y + (3 * dy) / 4
                Points(3).X = X - dx / 4
                Points(3).Y = Y + (3 * dy) / 4
                Points(4).X = Points(0).X
                Points(4).Y = Points(0).Y
                mp_DrawFigureAux(oFillColor, oBorderColor, Points)
            Case GRE_FIGURETYPE.FT_CIRCLE
                mp_FillEllipseCentered(X, System.Convert.ToInt32(Y + dy / 2), dx - 1, dy - 1, oFillColor)
            Case Else
                Return
        End Select

    End Sub

    Private Sub mp_DrawFigureAux(ByVal oFillColor As Color, ByVal oBorderColor As Color, ByVal oPoints() As Point)
        If PolygonInsideCanvas(oPoints) = True Then
            mp_FillPolygon(oPoints, oFillColor)
            mp_DrawPolyline(oPoints, oBorderColor)
        End If
    End Sub

    Public Sub DrawPattern(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal oColor As Color, ByVal v_lDrawStyle As GRE_PATTERN, ByVal v_iPatternFactor As Integer)
        Dim tmp As Integer
        Dim c As Integer
        Dim c1 As Integer
        Dim c2 As Integer
        Dim i1 As Integer
        Dim j1 As Integer
        Dim i2 As Integer
        Dim j2 As Integer
        If v_X1 > v_X2 Then
            tmp = v_X1
            v_X1 = v_X2
            v_X2 = tmp
        End If
        If v_Y1 > v_Y2 Then
            tmp = v_Y1
            v_Y1 = v_Y2
            v_Y2 = tmp
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_HORIZONTALLINE Or v_lDrawStyle = GRE_PATTERN.FP_CROSS Then
            For j1 = (v_Y1 + v_iPatternFactor) To v_Y2 Step v_iPatternFactor
                DrawLine(v_X1, j1, v_X2, j1, GRE_LINETYPE.LT_NORMAL, oColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            Next j1
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_VERTICALLINE Or v_lDrawStyle = GRE_PATTERN.FP_CROSS Then
            For j1 = (v_X1 + v_iPatternFactor) To v_X2 Step v_iPatternFactor
                DrawLine(j1, v_Y1, j1, v_Y2, GRE_LINETYPE.LT_NORMAL, oColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            Next j1
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_UPWARDDIAGONAL Or v_lDrawStyle = GRE_PATTERN.FP_DIAGONALCROSS Then
            c1 = System.Convert.ToInt32((v_Y1 + v_X1) / v_iPatternFactor + 1)
            c2 = System.Convert.ToInt32((v_Y2 + v_X2) / v_iPatternFactor)
            For c = c1 To c2
                i1 = v_X1
                i2 = v_X2
                j1 = c * v_iPatternFactor - i1
                j2 = c * v_iPatternFactor - i2
                If j2 < v_Y1 Then
                    i2 = c * v_iPatternFactor - v_Y1
                    j2 = c * v_iPatternFactor - i2
                End If
                If j1 > v_Y2 Then
                    i1 = c * v_iPatternFactor - v_Y2
                    j1 = c * v_iPatternFactor - i1
                End If
                DrawLine(i1, j1, i2, j2, GRE_LINETYPE.LT_NORMAL, oColor, GRE_LINEDRAWSTYLE.LDS_SOLID, 1, False)
            Next c
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_DOWNWARDDIAGONAL Or v_lDrawStyle = GRE_PATTERN.FP_DIAGONALCROSS Then
            c1 = System.Convert.ToInt32((v_Y1 - v_X2) / v_iPatternFactor + 1)
            c2 = System.Convert.ToInt32((v_Y2 - v_X1) / v_iPatternFactor)
            For c = c1 To c2
                i1 = v_X1
                i2 = v_X2
                j1 = i1 + c * v_iPatternFactor
                j2 = i2 + c * v_iPatternFactor
                If j1 < v_Y1 Then
                    i1 = v_Y1 - c * v_iPatternFactor
                    j1 = i1 + c * v_iPatternFactor
                End If
                If j2 > v_Y2 Then
                    i2 = v_Y2 - c * v_iPatternFactor
                    j2 = i2 + c * v_iPatternFactor
                End If
                DrawLine(i1, j1, i2, j2, GRE_LINETYPE.LT_NORMAL, oColor, GRE_LINEDRAWSTYLE.LDS_SOLID, 1, False)
            Next c
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_LIGHT Then
            For j1 = (v_Y1 + 1) To (v_Y2 - 1)
                If j1 Mod 2 = 0 Then
                    For j2 = (v_X1 + 1) To (v_X2 - 1) Step 4
                        DrawLine(j2, j1, j2 + 1, j1, GRE_LINETYPE.LT_NORMAL, oColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    Next j2
                Else
                    For j2 = (v_X1 + 3) To (v_X2 - 1) Step 4
                        DrawLine(j2, j1, j2 + 1, j1, GRE_LINETYPE.LT_NORMAL, oColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    Next j2
                End If
            Next j1
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_MEDIUM Then
            For j1 = (v_Y1 + 1) To (v_Y2 - 1)
                If j1 Mod 2 = 0 Then
                    For j2 = (v_X1 + 1) To (v_X2 - 1) Step 2
                        DrawLine(j2, j1, j2 + 1, j1, GRE_LINETYPE.LT_NORMAL, oColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    Next j2
                Else
                    For j2 = (v_X1 + 2) To (v_X2 - 1) Step 2
                        DrawLine(j2, j1, j2 + 1, j1, GRE_LINETYPE.LT_NORMAL, oColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    Next j2
                End If
            Next j1
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_DARK Then
            For j1 = (v_Y1 + 1) To (v_Y2 - 1)
                If j1 Mod 2 = 0 Then
                    For j2 = (v_X1 + 1) To (v_X2 - 1) Step 4
                        If j2 + 3 < v_X2 Then
                            DrawLine(j2, j1, j2 + 3, j1, GRE_LINETYPE.LT_NORMAL, oColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                        Else
                            DrawLine(j2, j1, v_X2, j1, GRE_LINETYPE.LT_NORMAL, oColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                        End If
                    Next j2
                Else
                    DrawLine(v_X1, j1, v_X1 + 2, j1, GRE_LINETYPE.LT_NORMAL, oColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    For j2 = (v_X1 + 3) To (v_X2 - 1) Step 4
                        If j2 + 3 < v_X2 Then
                            DrawLine(j2, j1, j2 + 3, j1, GRE_LINETYPE.LT_NORMAL, oColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                        Else
                            DrawLine(j2, j1, v_X2, j1, GRE_LINETYPE.LT_NORMAL, oColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                        End If
                    Next j2
                End If
            Next j1
        End If
    End Sub

    Public Sub DrawPolyLine(ByVal oColor As Color, ByVal v_lWidth As Integer, ByVal v_lDrawStyle As GRE_LINEDRAWSTYLE, ByRef r_oPoints() As Point, ByVal v_Len As Integer)

    End Sub

    Public Sub DrawTextEx(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_sParam As String, ByVal v_lFlags As clsTextFlags, ByVal oColor As Color, ByVal v_oFont As Font, Optional ByVal v_bClip As Boolean = True)
        If (v_X1 < mp_lClipX1 And v_X2 < mp_lClipX1) Or (v_X1 > mp_lClipX2 And v_X2 > mp_lClipX2) Then
            Return
        End If
        If (v_Y1 < mp_lClipY1 And v_Y2 < mp_lClipY1) Or (v_Y1 > mp_lClipY2 And v_Y2 > mp_lClipY2) Then
            Return
        End If
        Dim oTextBlock As New TextBlock
        oTextBlock.FontFamily = New FontFamily(v_oFont.FamilyName)
        oTextBlock.Text = v_sParam
        oTextBlock.FontSize = v_oFont.Size
        oTextBlock.Foreground = New SolidColorBrush(oColor)
        oTextBlock.FontWeight = v_oFont.FontWeight
        oTextBlock.Measure(New Size(v_X2 - v_X1, v_Y2 - v_Y1))
        Dim X As Integer = 0
        Dim Y As Integer = 0
        Select Case v_lFlags.HorizontalAlignment
            Case GRE_HORIZONTALALIGNMENT.HAL_LEFT
                X = System.Convert.ToInt32(v_X1)
            Case GRE_HORIZONTALALIGNMENT.HAL_CENTER
                X = System.Convert.ToInt32(((v_X2 - v_X1) - oTextBlock.ActualWidth) / 2) + v_X1
            Case GRE_HORIZONTALALIGNMENT.HAL_RIGHT
                X = System.Convert.ToInt32(v_X2 - oTextBlock.ActualWidth)
        End Select
        Select Case v_lFlags.VerticalAlignment
            Case GRE_VERTICALALIGNMENT.VAL_TOP
                Y = System.Convert.ToInt32(v_Y1)
            Case GRE_VERTICALALIGNMENT.VAL_CENTER
                Y = System.Convert.ToInt32(((v_Y2 - v_Y1) - oTextBlock.ActualHeight) / 2) + v_Y1
            Case GRE_VERTICALALIGNMENT.VAL_BOTTOM
                Y = System.Convert.ToInt32(v_Y2 - oTextBlock.ActualHeight)
        End Select

        Dim oTransform As New TranslateTransform()
        oTransform.X = X
        oTransform.Y = Y

        GetBitmap.Render(oTextBlock, oTransform)
        If v_sParam.Length > 0 Then
            mp_oTextFinalLayout.X = X
            mp_oTextFinalLayout.Y = Y
            mp_oTextFinalLayout.Width = oTextBlock.ActualWidth
            mp_oTextFinalLayout.Height = oTextBlock.ActualHeight
        End If
        GetBitmap.Invalidate()

    End Sub

    Public Sub DrawAlignedText(ByVal v_lLeft As Integer, ByVal v_lTop As Integer, ByVal v_lRight As Integer, ByVal v_lBottom As Integer, ByVal v_sParam As String, ByVal v_yHPos As GRE_HORIZONTALALIGNMENT, ByVal v_yVPos As GRE_VERTICALALIGNMENT, ByVal oColor As Color, ByVal v_oFont As Font)
        DrawAlignedText(v_lLeft, v_lTop, v_lRight, v_lBottom, v_sParam, v_yHPos, v_yVPos, oColor, v_oFont, True)
    End Sub

    Public Sub DrawAlignedText(ByVal v_lLeft As Integer, ByVal v_lTop As Integer, ByVal v_lRight As Integer, ByVal v_lBottom As Integer, ByVal v_sParam As String, ByVal v_yHPos As GRE_HORIZONTALALIGNMENT, ByVal v_yVPos As GRE_VERTICALALIGNMENT, ByVal oColor As Color, ByVal v_oFont As Font, ByVal v_bClip As Boolean)
        If (v_lLeft < mp_lClipX1 And v_lRight < mp_lClipX1) Or (v_lLeft > mp_lClipX2 And v_lRight > mp_lClipX2) Then
            Return
        End If
        If (v_lTop < mp_lClipY1 And v_lBottom < mp_lClipY1) Or (v_lTop > mp_lClipY2 And v_lBottom > mp_lClipY2) Then
            Return
        End If
        Dim oTextBlock As New TextBlock
        oTextBlock.FontFamily = New FontFamily(v_oFont.FamilyName)
        oTextBlock.Text = v_sParam
        oTextBlock.FontSize = v_oFont.Size
        oTextBlock.Foreground = New SolidColorBrush(oColor)
        oTextBlock.FontWeight = v_oFont.FontWeight
        If v_lRight - v_lLeft <= 0 Then
            Return
        End If
        If v_lBottom - v_lTop <= 0 Then
            Return
        End If
        oTextBlock.Measure(New Size(v_lRight - v_lLeft, v_lBottom - v_lTop))

        Dim X As Integer = 0
        Dim Y As Integer = 0
        Select Case v_yHPos
            Case GRE_HORIZONTALALIGNMENT.HAL_LEFT
                X = System.Convert.ToInt32(v_lLeft)
            Case GRE_HORIZONTALALIGNMENT.HAL_CENTER
                X = System.Convert.ToInt32(((v_lRight - v_lLeft) - oTextBlock.ActualWidth) / 2) + v_lLeft
            Case GRE_HORIZONTALALIGNMENT.HAL_RIGHT
                X = System.Convert.ToInt32(v_lRight - oTextBlock.ActualWidth)
        End Select
        Select Case v_yVPos
            Case GRE_VERTICALALIGNMENT.VAL_TOP
                Y = System.Convert.ToInt32(v_lTop)
            Case GRE_VERTICALALIGNMENT.VAL_CENTER
                Y = System.Convert.ToInt32(((v_lBottom - v_lTop) - oTextBlock.ActualHeight) / 2) + v_lTop
            Case GRE_VERTICALALIGNMENT.VAL_BOTTOM
                Y = System.Convert.ToInt32(v_lBottom - oTextBlock.ActualHeight)
        End Select

        Dim lClipX1 As Integer
        Dim lClipX2 As Integer
        Dim lClipY1 As Integer
        Dim lClipY2 As Integer
        If (X - mp_lClipX1) > 0 Then
            lClipX1 = 0
        Else
            lClipX1 = mp_lClipX1 - X
        End If
        If ((X + oTextBlock.ActualWidth) - mp_lClipX2) > 0 Then
            lClipX2 = System.Convert.ToInt32((X + oTextBlock.ActualWidth) - mp_lClipX2)
        Else
            lClipX2 = 0
        End If

        If (Y - mp_lClipY1) > 0 Then
            lClipY1 = 0
        Else
            lClipY1 = mp_lClipY1 - Y
        End If
        If ((Y + oTextBlock.ActualHeight) - mp_lClipY2) > 0 Then
            lClipY2 = System.Convert.ToInt32((Y + oTextBlock.ActualHeight) - mp_lClipY2)
        Else
            lClipY2 = 0
        End If
        If (X + oTextBlock.ActualWidth) < mp_lClipX1 Or X > mp_lClipX2 Then
            Return
        End If
        If (Y + oTextBlock.ActualHeight) < mp_lClipY1 Or Y > mp_lClipY2 Then
            Return
        End If


        Dim oClipRectangle As New RectangleGeometry()
        If oTextBlock.ActualWidth - lClipX2 <= 0 Then
            Return
        End If
        If oTextBlock.ActualHeight - lClipY2 <= 0 Then
            Return
        End If
        oClipRectangle.Rect = New System.Windows.Rect(lClipX1, lClipY1, oTextBlock.ActualWidth - lClipX2, oTextBlock.ActualHeight - lClipY2)
        oTextBlock.Clip = oClipRectangle

        Dim oTransform As New TranslateTransform()
        oTransform.X = X
        oTransform.Y = Y

        GetBitmap.Render(oTextBlock, oTransform)
        If v_sParam.Length > 0 Then
            mp_oTextFinalLayout.X = X
            mp_oTextFinalLayout.Y = Y
            mp_oTextFinalLayout.Width = oTextBlock.ActualWidth
            mp_oTextFinalLayout.Height = oTextBlock.ActualHeight
        End If
        GetBitmap.Invalidate()
    End Sub

    Public Sub CorrectRectCoords(ByRef X1 As Integer, ByRef Y1 As Integer, ByRef X2 As Integer, ByRef Y2 As Integer)
        Dim iBuff As Integer = 0
        If (X2 - X1) < 0 Then
            iBuff = X1
            X1 = X2
            X2 = iBuff
        End If
        If (Y2 - Y1) < 0 Then
            iBuff = Y1
            Y1 = Y2
            Y2 = iBuff
        End If
        Dim bStraightLine_H As Boolean = False
        Dim bStraightLine_V As Boolean = False
        Dim bFill As Boolean = False
        Dim bCheckClip As Boolean = False

        If (Y1 = Y2) Then
            bStraightLine_H = True
        ElseIf (X1 = X2) Then
            bStraightLine_V = True
        Else
            bFill = True
        End If
        If bStraightLine_H = True Or bFill = True Then
            If X1 < mp_lClipX1 Then
                If System.Math.Abs(mp_lClipX1 - X1) > 10 Then
                    X1 = mp_lClipX1 - 1
                End If
                bCheckClip = True
            End If
            If X2 > mp_lClipX2 Then
                If System.Math.Abs(X2 - mp_lClipX2) > 10 Then
                    X2 = mp_lClipX2 + 1
                End If
                bCheckClip = True
            End If
        End If
        If bStraightLine_V = True Or bFill = True Then
            If Y1 < mp_lClipY1 Then
                If System.Math.Abs(mp_lClipY1 - Y1) > 10 Then
                    Y1 = mp_lClipY1 - 1
                End If
                bCheckClip = True
            End If
            If Y2 > mp_lClipY2 Then
                If System.Math.Abs(Y2 - mp_lClipY2) > 10 Then
                    Y2 = mp_lClipY2 + 1
                End If
                bCheckClip = True
            End If
        End If
        mp_bCheckClip = bCheckClip
    End Sub

    Public Sub ClipRegion(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_bStore As Boolean)
        'Dim iBuff As Integer = 0
        'If (v_X2 - v_X1) < 0 Then
        '    iBuff = v_X1
        '    v_X1 = v_X2
        '    v_X2 = iBuff
        'End If
        'If (v_Y2 - v_Y1) < 0 Then
        '    iBuff = v_Y1
        '    v_Y1 = v_Y2
        '    v_Y2 = iBuff
        'End If
        'If Not v_X1 < v_X2 Then
        '    ClearClipRegion()
        '    Return
        'End If
        'If Not v_Y1 < v_Y2 Then
        '    ClearClipRegion()
        '    Return
        'End If
        mp_lClipX1 = v_X1
        mp_lClipY1 = v_Y1
        mp_lClipX2 = v_X2
        mp_lClipY2 = v_Y2
        If v_bStore = True Then
            mp_udtPreviousClipRegion.lLeft = v_X1
            mp_udtPreviousClipRegion.lRight = v_X2
            mp_udtPreviousClipRegion.lTop = v_Y1
            mp_udtPreviousClipRegion.lBottom = v_Y2
        End If
    End Sub

    Public Sub RestorePreviousClipRegion()
        ClipRegion(mp_udtPreviousClipRegion.lLeft, mp_udtPreviousClipRegion.lTop, mp_udtPreviousClipRegion.lRight, mp_udtPreviousClipRegion.lBottom, False)
    End Sub

    Public Sub ClearClipRegion()
        mp_lClipX1 = 0
        mp_lClipY1 = 0
        mp_lClipX2 = Width
        mp_lClipY2 = Height
    End Sub

    Public Sub TileImageHorizontal(ByVal ImageHandle As Image, ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_bTransparent As Boolean)
        Dim X As Integer
        Dim lImageWidth As Integer
        Dim lImageHeight As Integer
        lImageHeight = System.Convert.ToInt32(ImageHandle.Height)
        lImageWidth = System.Convert.ToInt32(ImageHandle.Width)
        Do While X < (v_X2 - v_X1)
            If (X + lImageWidth) > (v_X2 - v_X1) Then
                PaintImage(ImageHandle, v_X2 - lImageWidth, v_Y1, v_X2, v_Y1 + lImageHeight, 0, 0, v_bTransparent)
            Else
                PaintImage(ImageHandle, v_X1 + X, v_Y1, v_X1 + X + lImageWidth, v_Y1 + lImageHeight, 0, 0, v_bTransparent)
            End If
            X = X + lImageWidth
        Loop
    End Sub

    Public Sub PaintImage(ByVal oImage As Image, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal xOrigin As Integer, ByVal yOrigin As Integer, ByVal bUseMask As Boolean)

        Dim lClipX1 As Integer
        Dim lClipX2 As Integer
        Dim lClipY1 As Integer
        Dim lClipY2 As Integer

        If (mp_lClipX1 - X1) > 0 Then
            lClipX1 = mp_lClipX1 - X1
        Else
            lClipX1 = 0
        End If
        If (mp_lClipX2 - X2) > 0 Then
            lClipX2 = 0
        Else
            lClipX2 = X2 - mp_lClipX2
        End If

        If (mp_lClipY1 - Y1) > 0 Then
            lClipY1 = mp_lClipY1 - Y1
        Else
            lClipY1 = 0
        End If
        If (mp_lClipY2 - Y2) > 0 Then
            lClipY2 = 0
        Else
            lClipY2 = Y2 - mp_lClipY2
        End If



        Dim oClipRectangle As New RectangleGeometry()
        If oImage.Width - lClipX2 <= 0 Then
            Return
        End If
        If oImage.Height - lClipY2 <= 0 Then
            Return
        End If
        oClipRectangle.Rect = New System.Windows.Rect(lClipX1, lClipY1, oImage.Width - lClipX2, oImage.Height - lClipY2)
        oImage.Clip = oClipRectangle

        Dim oTransform As New TranslateTransform()
        oTransform.X = X1
        oTransform.Y = Y1
        'oImage.Measure(New Size(X2 - X1, Y2 - Y1))
        'oImage.Arrange(New Rect(0, 0, X2 - X1, Y2 - Y1))
        GetBitmap.Render(oImage, oTransform)
        GetBitmap.Invalidate()
    End Sub

    Public Sub DrawImage(ByRef v_oImage As Image, ByRef v_yHorizontalAlignment As GRE_HORIZONTALALIGNMENT, ByRef v_yVerticalAlignment As GRE_VERTICALALIGNMENT, ByVal v_lImageXMargin As Integer, ByVal v_lImageYMargin As Integer, ByRef v_lLeft As Integer, ByRef v_lRight As Integer, ByRef v_lTop As Integer, ByRef v_lBottom As Integer, ByVal v_bTransparent As Boolean)
        Dim bDrawImage As Boolean
        Dim bHorizontalSmall As Boolean
        Dim bVerticalSmall As Boolean
        Dim XOrigin As Integer
        Dim YOrigin As Integer
        Dim xDest As Integer
        Dim yDest As Integer
        Dim lxWidth As Integer
        Dim lyHeight As Integer
        Dim lImageHeight As Integer
        Dim lImageWidth As Integer
        If (v_oImage Is Nothing) Then
            Return
        End If
        lImageHeight = System.Convert.ToInt32(v_oImage.Height)
        lImageWidth = System.Convert.ToInt32(v_oImage.Width)
        If v_yHorizontalAlignment = GRE_HORIZONTALALIGNMENT.HAL_CENTER Then
            v_lImageXMargin = 0
        End If
        If v_yVerticalAlignment = GRE_VERTICALALIGNMENT.VAL_CENTER Then
            v_lImageYMargin = 0
        End If
        bDrawImage = True
        If (v_lRight - v_lLeft) < (lImageWidth + v_lImageXMargin) Then
            lxWidth = v_lRight - v_lLeft - v_lImageXMargin
            If lxWidth <= 0 Then bDrawImage = False
            bHorizontalSmall = True
        Else
            lxWidth = lImageWidth
            bHorizontalSmall = False
        End If
        If (v_lBottom - v_lTop) < (lImageHeight + v_lImageYMargin) Then
            lyHeight = v_lBottom - v_lTop - v_lImageYMargin
            If lyHeight <= 0 Then bDrawImage = False
            bVerticalSmall = True
        Else
            lyHeight = lImageHeight
            bVerticalSmall = False
        End If
        If bHorizontalSmall = False Then
            Select Case v_yHorizontalAlignment
                Case GRE_HORIZONTALALIGNMENT.HAL_LEFT
                    xDest = v_lLeft + v_lImageXMargin
                Case GRE_HORIZONTALALIGNMENT.HAL_CENTER
                    xDest = System.Convert.ToInt32(((v_lRight - v_lLeft) - lImageWidth) / 2 + v_lLeft)
                Case GRE_HORIZONTALALIGNMENT.HAL_RIGHT
                    xDest = v_lRight - lImageWidth - v_lImageXMargin
            End Select
            XOrigin = 0
        Else
            Select Case v_yHorizontalAlignment
                Case GRE_HORIZONTALALIGNMENT.HAL_LEFT
                    XOrigin = 0
                    xDest = v_lLeft + v_lImageXMargin
                Case GRE_HORIZONTALALIGNMENT.HAL_CENTER
                    XOrigin = System.Convert.ToInt32((lImageWidth - lxWidth) / 2)
                    xDest = v_lLeft
                Case GRE_HORIZONTALALIGNMENT.HAL_RIGHT
                    XOrigin = lImageWidth - lxWidth
                    xDest = v_lRight - lxWidth - v_lImageXMargin
            End Select
        End If
        If bVerticalSmall = False Then
            Select Case v_yVerticalAlignment
                Case GRE_VERTICALALIGNMENT.VAL_TOP
                    yDest = v_lTop + v_lImageYMargin
                Case GRE_VERTICALALIGNMENT.VAL_CENTER
                    yDest = System.Convert.ToInt32(((v_lBottom - v_lTop) - lImageHeight) / 2 + v_lTop)
                Case GRE_VERTICALALIGNMENT.VAL_BOTTOM
                    yDest = v_lBottom - lImageHeight - v_lImageYMargin
            End Select
            YOrigin = 0
        Else
            Select Case v_yVerticalAlignment
                Case GRE_VERTICALALIGNMENT.VAL_TOP
                    YOrigin = 0
                    yDest = v_lTop + v_lImageYMargin
                Case GRE_VERTICALALIGNMENT.VAL_CENTER
                    YOrigin = System.Convert.ToInt32((lImageHeight - lyHeight) / 2)
                    yDest = v_lTop
                Case GRE_VERTICALALIGNMENT.VAL_BOTTOM
                    YOrigin = lImageHeight - lyHeight
                    yDest = v_lBottom - lyHeight - v_lImageYMargin
            End Select
        End If
        If bDrawImage = True Then
            PaintImage(v_oImage, xDest, yDest, xDest + lxWidth, yDest + lyHeight, XOrigin, YOrigin, v_bTransparent)
        End If
    End Sub

    Public Sub DrawFocusRectangle(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer)
        DrawLine(v_X1, v_Y1, v_X2, v_Y2, GRE_LINETYPE.LT_BORDER, mp_oControl.SelectionColor, GRE_LINEDRAWSTYLE.LDS_DOT)
    End Sub

    Public Sub GradientFill(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal clrStartColor As Color, ByVal clrEndColor As Color, ByVal iGradientType As GRE_GRADIENTFILLMODE)
        If (v_X2 - v_X1) <= 0 Then
            Return
        End If
        If (v_Y2 - v_Y1) <= 0 Then
            Return
        End If
        If (v_X1 < mp_lClipX1 And v_X2 < mp_lClipX1) Or (v_X1 > mp_lClipX2 And v_X2 > mp_lClipX2) Then
            Return
        End If
        If (v_Y1 < mp_lClipY1 And v_Y2 < mp_lClipY1) Or (v_Y1 > mp_lClipY2 And v_Y2 > mp_lClipY2) Then
            Return
        End If

        Dim oBrush As LinearGradientBrush = Nothing
        Dim oGradientStopCollection As New GradientStopCollection()
        Dim oGradientStopStart As New GradientStop
        Dim oGradientStopEnd As New GradientStop
        Dim oRectangle As New Rectangle
        oGradientStopStart.Color = clrStartColor
        oGradientStopStart.Offset = 0
        oGradientStopEnd.Color = clrEndColor
        oGradientStopEnd.Offset = 1
        oGradientStopCollection.Add(oGradientStopEnd)
        oGradientStopCollection.Add(oGradientStopStart)

        If (iGradientType = GRE_GRADIENTFILLMODE.GDT_VERTICAL) Then
            oBrush = New LinearGradientBrush(oGradientStopCollection, 90.0)
        ElseIf (iGradientType = GRE_GRADIENTFILLMODE.GDT_HORIZONTAL) Then
            oBrush = New LinearGradientBrush(oGradientStopCollection, 0.0)
        End If

        oRectangle.Width = v_X2 - v_X1 + 1
        oRectangle.Height = v_Y2 - v_Y1 + 1
        oRectangle.Fill = oBrush

        Dim lClipX1 As Integer
        Dim lClipX2 As Integer
        Dim lClipY1 As Integer
        Dim lClipY2 As Integer

        If (mp_lClipX1 - v_X1) > 0 Then
            lClipX1 = mp_lClipX1 - v_X1
        Else
            lClipX1 = 0
        End If
        If (mp_lClipX2 - v_X2) > 0 Then
            lClipX2 = 0
        Else
            lClipX2 = v_X2 - mp_lClipX2
        End If

        If (mp_lClipY1 - v_Y1) > 0 Then
            lClipY1 = mp_lClipY1 - v_Y1
        Else
            lClipY1 = 0
        End If
        If (mp_lClipY2 - v_Y2) > 0 Then
            lClipY2 = 0
        Else
            lClipY2 = v_Y2 - mp_lClipY2
        End If

        Dim oClipRectangle As New RectangleGeometry()
        If oRectangle.Width - lClipX2 <= 0 Then
            Return
        End If
        If oRectangle.Height - lClipY2 <= 0 Then
            Return
        End If
        oClipRectangle.Rect = New System.Windows.Rect(lClipX1, lClipY1, oRectangle.Width - lClipX2, oRectangle.Height - lClipY2)
        oRectangle.Clip = oClipRectangle

        Dim oTransform As New TranslateTransform()
        oTransform.X = v_X1
        oTransform.Y = v_Y1

        GetBitmap.Render(oRectangle, oTransform)
        GetBitmap.Invalidate()
    End Sub

    Public Sub HatchFill(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal oForeColor As Color, ByVal oBackColor As Color, ByVal yHatchStyle As GRE_HATCHSTYLE)
        CorrectRectCoords(v_X1, v_Y1, v_X2, v_Y2)
        Dim lBackColor As Integer = SWM_Color_To_Int32(oBackColor)
        Dim lForeColor As Integer = SWM_Color_To_Int32(oForeColor)
        mp_aPixels = GetBitmap.Pixels
        Dim iPixelCount As Integer = mp_aPixels.Length
        Dim lWidth As Integer = GetBitmap.PixelWidth
        Dim lHeight As Integer = GetBitmap.PixelHeight
        Dim y As Integer
        Dim x As Integer
        Dim iy As Integer
        Dim ix As Integer
        Dim xOffset As Integer = 0
        Dim xDistance As Integer = 0
        Dim yDistance As Integer = 0
        Dim xStart As Integer = 0
        Dim xEnd As Integer = 0
        Dim yEnd As Integer = 0
        Dim aRect As Integer(,)

        If v_X1 > lWidth Then
            Return
        End If
        If v_Y1 > lHeight Then
            Return
        End If
        If v_X1 < 0 Then
            v_X1 = 0
        End If
        If v_X2 > lWidth Then
            v_X2 = lWidth
        End If

        If v_Y1 < 0 Then
            v_Y1 = 0
        End If
        If v_Y2 > lHeight Then
            v_Y2 = lHeight
        End If
        If yHatchStyle = GRE_HATCHSTYLE.HS_PERCENT60 Or yHatchStyle = GRE_HATCHSTYLE.HS_PERCENT70 Or yHatchStyle = GRE_HATCHSTYLE.HS_PERCENT75 Or yHatchStyle = GRE_HATCHSTYLE.HS_PERCENT80 Or yHatchStyle = GRE_HATCHSTYLE.HS_PERCENT90 Then
            For y = v_Y1 To v_Y2
                For x = v_X1 To v_X2
                    SetPixel((y * lWidth) + x, lForeColor)
                Next
            Next
        Else
            For y = v_Y1 To v_Y2
                For x = v_X1 To v_X2
                    SetPixel((y * lWidth) + x, lBackColor)
                Next
            Next
        End If



        Select Case yHatchStyle
            Case GRE_HATCHSTYLE.HS_HORIZONTAL
                yDistance = 8
                For y = v_Y1 To v_Y2 Step yDistance
                    For x = v_X1 To v_X2
                        SetPixel((y * lWidth) + x, lForeColor)
                    Next
                Next
            Case GRE_HATCHSTYLE.HS_VERTICAL
                xDistance = 8
                For y = v_Y1 To v_Y2
                    For x = v_X1 To v_X2 Step xDistance
                        SetPixel((y * lWidth) + x, lForeColor)
                    Next
                Next
            Case GRE_HATCHSTYLE.HS_FORWARDDIAGONAL
                xDistance = 8
                For y = v_Y1 To v_Y2
                    xStart = (y * lWidth) + v_X1
                    xEnd = (y * lWidth) + v_X2
                    For x = (xStart - xDistance) To (xEnd + xDistance) Step xDistance
                        If (x + xOffset) >= xStart And (x + xOffset) <= xEnd Then
                            SetPixel(x + xOffset, lForeColor)
                        End If
                    Next
                    xOffset = xOffset + 1
                    If xOffset > (xDistance - 1) Then
                        xOffset = 0
                    End If
                Next
            Case GRE_HATCHSTYLE.HS_BACKWARDDIAGONAL
                xDistance = 8
                For y = v_Y1 To v_Y2
                    xStart = (y * lWidth) + v_X1
                    xEnd = (y * lWidth) + v_X2
                    For x = (xStart - xDistance) To (xEnd + xDistance) Step xDistance
                        If (x - xOffset) >= xStart And (x - xOffset) <= xEnd Then
                            SetPixel(x - xOffset, lForeColor)
                        End If
                    Next
                    xOffset = xOffset + 1
                    If xOffset > (xDistance - 1) Then
                        xOffset = 0
                    End If
                Next
            Case GRE_HATCHSTYLE.HS_LARGEGRID
                xDistance = 8
                yDistance = 8
                For y = (v_Y1) To v_Y2 Step yDistance
                    For x = v_X1 To v_X2
                        SetPixel((y * lWidth) + x, lForeColor)
                    Next
                Next
                For y = v_Y1 To v_Y2
                    For x = (v_X1) To v_X2 Step xDistance
                        SetPixel((y * lWidth) + x, lForeColor)
                    Next
                Next
            Case GRE_HATCHSTYLE.HS_DIAGONALCROSS
                xDistance = 8
                For y = v_Y1 To v_Y2
                    xStart = (y * lWidth) + v_X1
                    xEnd = (y * lWidth) + v_X2
                    For x = (xStart - xDistance) To (xEnd + xDistance) Step xDistance
                        If (x + xOffset) >= xStart And (x + xOffset) <= xEnd Then
                            SetPixel(x + xOffset, lForeColor)
                        End If
                    Next
                    xOffset = xOffset + 1
                    If xOffset > (xDistance - 1) Then
                        xOffset = 0
                    End If
                Next
                For y = v_Y1 To v_Y2
                    xStart = (y * lWidth) + v_X1
                    xEnd = (y * lWidth) + v_X2
                    For x = (xStart - xDistance) To (xEnd + xDistance) Step xDistance
                        If (x - xOffset) >= xStart And (x - xOffset) <= xEnd Then
                            SetPixel(x - xOffset, lForeColor)
                        End If
                    Next
                    xOffset = xOffset + 1
                    If xOffset > (xDistance - 1) Then
                        xOffset = 0
                    End If
                Next
            Case GRE_HATCHSTYLE.HS_PERCENT05
                aRect = New Integer(,) { _
                {0, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {1, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_PERCENT10
                aRect = New Integer(,) { _
                {0, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {1, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {1, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_PERCENT20
                aRect = New Integer(,) { _
                {1, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 1, 0, 0, 0, 1, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {1, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 1, 0, 0, 0, 1, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_PERCENT25
                aRect = New Integer(,) { _
                {1, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 1, 0, 0, 0, 1, 0}, _
                {1, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 1, 0, 0, 0, 1, 0}, _
                {1, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 1, 0, 0, 0, 1, 0}, _
                {1, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 1, 0, 0, 0, 1, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_PERCENT30
                aRect = New Integer(,) { _
                {1, 0, 1, 0, 1, 0, 1, 0}, _
                {0, 1, 0, 0, 0, 1, 0, 0}, _
                {1, 0, 1, 0, 1, 0, 1, 0}, _
                {0, 0, 0, 1, 0, 0, 0, 1}, _
                {1, 0, 1, 0, 1, 0, 1, 0}, _
                {0, 1, 0, 0, 0, 1, 0, 0}, _
                {1, 0, 1, 0, 1, 0, 1, 0}, _
                {0, 0, 0, 1, 0, 0, 0, 1} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_PERCENT40
                aRect = New Integer(,) { _
                {1, 0, 1, 0, 1, 0, 1, 0}, _
                {0, 0, 0, 1, 0, 1, 0, 1}, _
                {1, 0, 1, 0, 1, 0, 1, 0}, _
                {0, 1, 0, 1, 0, 1, 0, 1}, _
                {1, 0, 1, 0, 1, 0, 1, 0}, _
                {0, 1, 0, 1, 0, 0, 0, 1}, _
                {1, 0, 1, 0, 1, 0, 1, 0}, _
                {0, 1, 0, 1, 0, 1, 0, 1} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_PERCENT50
                aRect = New Integer(,) { _
                {1, 0, 1, 0, 1, 0, 1, 0}, _
                {0, 1, 0, 1, 0, 1, 0, 1}, _
                {1, 0, 1, 0, 1, 0, 1, 0}, _
                {0, 1, 0, 1, 0, 1, 0, 1}, _
                {1, 0, 1, 0, 1, 0, 1, 0}, _
                {0, 1, 0, 1, 0, 1, 0, 1}, _
                {1, 0, 1, 0, 1, 0, 1, 0}, _
                {0, 1, 0, 1, 0, 1, 0, 1} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_PERCENT60
                aRect = New Integer(,) { _
                {0, 1, 0, 0, 0, 1, 0, 0}, _
                {1, 0, 1, 0, 1, 0, 1, 0}, _
                {0, 0, 0, 1, 0, 0, 0, 1}, _
                {1, 0, 1, 0, 1, 0, 1, 0}, _
                {0, 1, 0, 0, 0, 1, 0, 0}, _
                {1, 0, 1, 0, 1, 0, 1, 0}, _
                {0, 0, 0, 1, 0, 0, 0, 1}, _
                {1, 0, 1, 0, 1, 0, 1, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lBackColor)

            Case GRE_HATCHSTYLE.HS_PERCENT70
                aRect = New Integer(,) { _
                {1, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 1, 0, 0, 0, 1, 0}, _
                {1, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 1, 0, 0, 0, 1, 0}, _
                {1, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 1, 0, 0, 0, 1, 0}, _
                {1, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 1, 0, 0, 0, 1, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lBackColor)

            Case GRE_HATCHSTYLE.HS_PERCENT75
                aRect = New Integer(,) { _
                {0, 0, 1, 0, 0, 0, 1, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {1, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 1, 0, 0, 0, 1, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {1, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lBackColor)

            Case GRE_HATCHSTYLE.HS_PERCENT80
                aRect = New Integer(,) { _
                {0, 0, 0, 1, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 1}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 1, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 1}, _
                {0, 0, 0, 0, 0, 0, 0, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lBackColor)

            Case GRE_HATCHSTYLE.HS_PERCENT90
                aRect = New Integer(,) { _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {1, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lBackColor)

            Case GRE_HATCHSTYLE.HS_LIGHTDOWNWARDDIAGONAL
                xDistance = 4
                For y = v_Y1 To v_Y2
                    xStart = (y * lWidth) + v_X1
                    xEnd = (y * lWidth) + v_X2
                    For x = (xStart - xDistance) To (xEnd + xDistance) Step xDistance
                        If (x + xOffset) >= xStart And (x + xOffset) <= xEnd Then
                            SetPixel(x + xOffset, lForeColor)
                        End If
                    Next
                    xOffset = xOffset + 1
                    If xOffset > (xDistance - 1) Then
                        xOffset = 0
                    End If
                Next
            Case GRE_HATCHSTYLE.HS_LIGHTUPWARDDIAGONAL
                xDistance = 4
                For y = v_Y1 To v_Y2
                    xStart = (y * lWidth) + v_X1
                    xEnd = (y * lWidth) + v_X2
                    For x = (xStart - xDistance) To (xEnd + xDistance) Step xDistance
                        If (x - xOffset) >= xStart And (x - xOffset) <= xEnd Then
                            SetPixel(x - xOffset, lForeColor)
                        End If
                    Next
                    xOffset = xOffset + 1
                    If xOffset > (xDistance - 1) Then
                        xOffset = 0
                    End If
                Next
            Case GRE_HATCHSTYLE.HS_DARKDOWNWARDDIAGONAL
                xDistance = 4
                For y = v_Y1 To v_Y2
                    xStart = (y * lWidth) + v_X1
                    xEnd = (y * lWidth) + v_X2
                    For x = (xStart - xDistance) To (xEnd + xDistance) Step xDistance
                        If (x + xOffset) >= xStart And (x + xOffset) <= xEnd Then
                            SetPixel(x + xOffset, lForeColor)
                        End If
                        If (x + xOffset + 1) >= xStart And (x + xOffset + 1) <= xEnd Then
                            SetPixel(x + xOffset + 1, lForeColor)
                        End If
                    Next
                    xOffset = xOffset + 1
                    If xOffset > (xDistance - 1) Then
                        xOffset = 0
                    End If
                Next




            Case GRE_HATCHSTYLE.HS_DARKUPWARDDIAGONAL
                xDistance = 4
                For y = v_Y1 To v_Y2
                    xStart = (y * lWidth) + v_X1
                    xEnd = (y * lWidth) + v_X2
                    For x = (xStart - xDistance) To (xEnd + xDistance) Step xDistance
                        If (x - xOffset) >= xStart And (x - xOffset) <= xEnd Then
                            SetPixel(x - xOffset, lForeColor)
                        End If
                        If (x - xOffset + 1) >= xStart And (x - xOffset + 1) <= xEnd Then
                            SetPixel(x - xOffset + 1, lForeColor)
                        End If
                    Next
                    xOffset = xOffset + 1
                    If xOffset > (xDistance - 1) Then
                        xOffset = 0
                    End If
                Next
            Case GRE_HATCHSTYLE.HS_WIDEDOWNWARDDIAGONAL
                xDistance = 8
                For y = v_Y1 To v_Y2
                    xStart = (y * lWidth) + v_X1
                    xEnd = (y * lWidth) + v_X2
                    For x = (xStart - xDistance) To (xEnd + xDistance) Step xDistance
                        If (x + xOffset) >= xStart And (x + xOffset) <= xEnd Then
                            SetPixel(x + xOffset, lForeColor)
                        End If
                        If (x + xOffset + 1) >= xStart And (x + xOffset + 1) <= xEnd Then
                            SetPixel(x + xOffset + 1, lForeColor)
                        End If
                        If (x + xOffset + 2) >= xStart And (x + xOffset + 2) <= xEnd Then
                            SetPixel(x + xOffset + 2, lForeColor)
                        End If
                    Next
                    xOffset = xOffset + 1
                    If xOffset > (xDistance - 1) Then
                        xOffset = 0
                    End If
                Next
            Case GRE_HATCHSTYLE.HS_WIDEUPWARDDIAGONAL
                xDistance = 8
                For y = v_Y1 To v_Y2
                    xStart = (y * lWidth) + v_X1
                    xEnd = (y * lWidth) + v_X2
                    For x = (xStart - xDistance) To (xEnd + xDistance) Step xDistance
                        If (x - xOffset) >= xStart And (x - xOffset) <= xEnd Then
                            SetPixel(x - xOffset, lForeColor)
                        End If
                        If (x - xOffset + 1) >= xStart And (x - xOffset + 1) <= xEnd Then
                            SetPixel(x - xOffset + 1, lForeColor)
                        End If
                        If (x - xOffset + 2) >= xStart And (x - xOffset + 2) <= xEnd Then
                            SetPixel(x - xOffset + 2, lForeColor)
                        End If
                    Next
                    xOffset = xOffset + 1
                    If xOffset > (xDistance - 1) Then
                        xOffset = 0
                    End If
                Next
            Case GRE_HATCHSTYLE.HS_LIGHTVERTICAL
                xDistance = 4
                For y = v_Y1 To v_Y2
                    For x = v_X1 To v_X2 Step xDistance
                        SetPixel((y * lWidth) + x, lForeColor)
                    Next
                Next
            Case GRE_HATCHSTYLE.HS_LIGHTHORIZONTAL
                yDistance = 4
                For y = v_Y1 To v_Y2 Step yDistance
                    For x = v_X1 To v_X2
                        SetPixel((y * lWidth) + x, lForeColor)
                    Next
                Next
            Case GRE_HATCHSTYLE.HS_NARROWVERTICAL
                xDistance = 2
                For y = v_Y1 To v_Y2
                    For x = v_X1 To v_X2 Step xDistance
                        SetPixel((y * lWidth) + x, lForeColor)
                    Next
                Next
            Case GRE_HATCHSTYLE.HS_NARROWHORIZONTAL
                yDistance = 2
                For y = v_Y1 To v_Y2 Step yDistance
                    For x = v_X1 To v_X2
                        SetPixel((y * lWidth) + x, lForeColor)
                    Next
                Next
            Case GRE_HATCHSTYLE.HS_DARKVERTICAL
                xDistance = 4
                For y = v_Y1 To v_Y2
                    xEnd = (y * lWidth) + v_X2
                    For x = v_X1 To v_X2 Step xDistance
                        SetPixel((y * lWidth) + x, lForeColor)
                        If ((y * lWidth) + x + 1) <= xEnd Then
                            SetPixel((y * lWidth) + x + 1, lForeColor)
                        End If
                    Next
                Next
            Case GRE_HATCHSTYLE.HS_DARKHORIZONTAL
                yDistance = 4
                yEnd = (v_Y2 * lWidth) + v_X2
                For y = v_Y1 To v_Y2 Step yDistance
                    For x = v_X1 To v_X2
                        SetPixel((y * lWidth) + x, lForeColor)
                        If (((y + 1) * lWidth) + x) <= yEnd Then
                            SetPixel(((y + 1) * lWidth) + x, lForeColor)
                        End If
                    Next
                Next
            Case GRE_HATCHSTYLE.HS_DASHEDDOWNWARDDIAGONAL
                aRect = New Integer(,) { _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {1, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 1, 0, 0, 0, 1, 0, 0}, _
                {0, 0, 1, 0, 0, 0, 1, 0}, _
                {0, 0, 0, 1, 0, 0, 0, 1}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_DASHEDUPWARDDIAGONAL
                aRect = New Integer(,) { _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 1, 0, 0, 0, 1}, _
                {0, 0, 1, 0, 0, 0, 1, 0}, _
                {0, 1, 0, 0, 0, 1, 0, 0}, _
                {1, 0, 0, 0, 1, 0, 0, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_DASHEDHORIZONTAL
                aRect = New Integer(,) { _
                {0, 0, 0, 0, 1, 1, 1, 1}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {1, 1, 1, 1, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_DASHEDVERTICAL
                aRect = New Integer(,) { _
                {1, 0, 0, 0, 0, 0, 0, 0}, _
                {1, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 0, 0, 1, 0, 0, 0}, _
                {1, 0, 0, 0, 0, 0, 0, 0}, _
                {1, 0, 0, 0, 0, 0, 0, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_SMALLCONFETTI
                aRect = New Integer(,) { _
                {1, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 1, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 1, 0}, _
                {0, 0, 0, 1, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 1}, _
                {0, 0, 1, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 1, 0, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_LARGECONFETTI
                aRect = New Integer(,) { _
                {0, 0, 0, 0, 1, 1, 0, 0}, _
                {1, 0, 0, 0, 1, 1, 0, 1}, _
                {1, 0, 1, 1, 0, 0, 0, 1}, _
                {0, 0, 1, 1, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 1, 1}, _
                {0, 0, 0, 1, 1, 0, 1, 1}, _
                {1, 1, 0, 1, 1, 0, 0, 0}, _
                {1, 1, 0, 0, 0, 0, 0, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_ZIGZAG
                aRect = New Integer(,) { _
                {1, 0, 0, 0, 0, 0, 0, 1}, _
                {0, 1, 0, 0, 0, 0, 1, 0}, _
                {0, 0, 1, 0, 0, 1, 0, 0}, _
                {0, 0, 0, 1, 1, 0, 0, 0}, _
                {1, 0, 0, 0, 0, 0, 0, 1}, _
                {0, 1, 0, 0, 0, 0, 1, 0}, _
                {0, 0, 1, 0, 0, 1, 0, 0}, _
                {0, 0, 0, 1, 1, 0, 0, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_WAVE
                aRect = New Integer(,) { _
                {0, 0, 1, 0, 0, 1, 0, 1}, _
                {1, 1, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 1, 1, 0, 0, 0}, _
                {0, 0, 1, 0, 0, 1, 0, 1}, _
                {1, 1, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 1, 1, 0, 0, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_DIAGONALBRICK
                aRect = New Integer(,) { _
                {0, 0, 0, 0, 0, 0, 0, 1}, _
                {0, 0, 0, 0, 0, 0, 1, 0}, _
                {0, 0, 0, 0, 0, 1, 0, 0}, _
                {0, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 0, 1, 1, 0, 0, 0}, _
                {0, 0, 1, 0, 0, 1, 0, 0}, _
                {0, 1, 0, 0, 0, 0, 1, 0}, _
                {1, 0, 0, 0, 0, 0, 0, 1} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_HORIZONTALBRICK
                aRect = New Integer(,) { _
                {0, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 0, 0, 1, 0, 0, 0}, _
                {1, 1, 1, 1, 1, 1, 1, 1}, _
                {1, 0, 0, 0, 0, 0, 0, 0}, _
                {1, 0, 0, 0, 0, 0, 0, 0}, _
                {1, 0, 0, 0, 0, 0, 0, 0}, _
                {1, 1, 1, 1, 1, 1, 1, 1}, _
                {0, 0, 0, 0, 1, 0, 0, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_WEAVE
                aRect = New Integer(,) { _
                {1, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 1, 0, 1, 0, 1, 0, 0}, _
                {0, 0, 1, 0, 0, 0, 1, 0}, _
                {0, 1, 0, 0, 0, 1, 0, 1}, _
                {1, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 0, 1, 0, 1, 0, 0}, _
                {0, 0, 1, 0, 0, 0, 1, 0}, _
                {0, 1, 0, 1, 0, 0, 0, 1} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_PLAID
                aRect = New Integer(,) { _
                {1, 1, 1, 1, 0, 0, 0, 0}, _
                {1, 1, 1, 1, 0, 0, 0, 0}, _
                {1, 0, 1, 0, 1, 0, 1, 0}, _
                {0, 1, 0, 1, 0, 1, 0, 1}, _
                {1, 0, 1, 0, 1, 0, 1, 0}, _
                {0, 1, 0, 1, 0, 1, 0, 1}, _
                {1, 1, 1, 1, 0, 0, 0, 0}, _
                {1, 1, 1, 1, 0, 0, 0, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_DIVOT
                aRect = New Integer(,) { _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {1, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 1}, _
                {1, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 1, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 0, 1, 0, 0, 0, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_DOTTEDGRID
                aRect = New Integer(,) { _
                {1, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {1, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {1, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {1, 0, 1, 0, 1, 0, 1, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)



            Case GRE_HATCHSTYLE.HS_DOTTEDDIAMOND
                aRect = New Integer(,) { _
                {1, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 1, 0, 0, 0, 1, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 1, 0, 0, 0, 1, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_SHINGLE
                aRect = New Integer(,) { _
                {0, 0, 0, 0, 0, 0, 0, 1}, _
                {0, 0, 0, 0, 0, 0, 0, 1}, _
                {0, 0, 0, 0, 0, 0, 1, 1}, _
                {1, 0, 0, 0, 0, 1, 0, 0}, _
                {0, 1, 0, 0, 1, 0, 0, 0}, _
                {0, 0, 1, 1, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 1, 1, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 1, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)


            Case GRE_HATCHSTYLE.HS_TRELLIS
                aRect = New Integer(,) { _
                {1, 1, 1, 1, 1, 1, 1, 1}, _
                {0, 1, 1, 0, 0, 1, 1, 0}, _
                {1, 1, 1, 1, 1, 1, 1, 1}, _
                {1, 0, 0, 1, 1, 0, 0, 1}, _
                {1, 1, 1, 1, 1, 1, 1, 1}, _
                {0, 1, 1, 0, 0, 1, 1, 0}, _
                {1, 1, 1, 1, 1, 1, 1, 1}, _
                {1, 0, 0, 1, 1, 0, 0, 1} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_SPHERE
                aRect = New Integer(,) { _
                {1, 0, 0, 0, 1, 1, 1, 1}, _
                {1, 0, 0, 0, 1, 1, 1, 1}, _
                {0, 1, 1, 1, 0, 1, 1, 1}, _
                {1, 0, 0, 1, 1, 0, 0, 0}, _
                {1, 1, 1, 1, 1, 0, 0, 0}, _
                {1, 1, 1, 1, 1, 0, 0, 0}, _
                {0, 1, 1, 1, 0, 1, 1, 1}, _
                {1, 0, 0, 0, 1, 0, 0, 1} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)

            Case GRE_HATCHSTYLE.HS_SMALLGRID
                xDistance = 4
                yDistance = 4
                For y = (v_Y1) To v_Y2 Step yDistance
                    For x = v_X1 To v_X2
                        SetPixel((y * lWidth) + x, lForeColor)
                    Next
                Next
                For y = v_Y1 To v_Y2
                    For x = (v_X1) To v_X2 Step xDistance
                        SetPixel((y * lWidth) + x, lForeColor)
                    Next
                Next
            Case GRE_HATCHSTYLE.HS_SMALLCHECKERBOARD
                xDistance = 4
                For y = v_Y1 To v_Y2
                    xStart = (y * lWidth) + v_X1
                    xEnd = (y * lWidth) + v_X2
                    For x = (xStart - xDistance) To (xEnd + xDistance) Step xDistance
                        If (x + xOffset) >= xStart And (x + xOffset) <= xEnd Then
                            SetPixel(x + xOffset, lForeColor)
                        End If
                    Next
                    xOffset = xOffset + 1
                    If xOffset > (xDistance - 1) Then
                        xOffset = 0
                    End If
                Next
                For y = v_Y1 To v_Y2
                    xStart = (y * lWidth) + v_X1
                    xEnd = (y * lWidth) + v_X2
                    For x = (xStart - xDistance) To (xEnd + xDistance) Step xDistance
                        If (x - xOffset) >= xStart And (x - xOffset) <= xEnd Then
                            SetPixel(x - xOffset, lForeColor)
                        End If
                    Next
                    xOffset = xOffset + 1
                    If xOffset > (xDistance - 1) Then
                        xOffset = 0
                    End If
                Next
            Case GRE_HATCHSTYLE.HS_LARGECHECKERBOARD
                xDistance = 8
                yDistance = 8
                yEnd = (v_Y2 * lWidth) + v_X2
                For y = v_Y1 To v_Y2 Step yDistance
                    xStart = (y * lWidth) + v_X1
                    xEnd = (y * lWidth) + v_X2
                    For x = v_X1 To v_X2 Step xDistance
                        For iy = 0 To 3
                            If (y + iy) <= v_Y2 Then
                                For ix = 0 To 3
                                    If ((y * lWidth) + x + ix) >= xStart And ((y * lWidth) + x + ix) <= xEnd Then
                                        SetPixel(((y + iy) * lWidth) + x + ix, lForeColor)
                                    End If
                                Next
                            End If
                        Next
                        For iy = 4 To 7
                            If (y + iy) <= v_Y2 Then
                                For ix = 4 To 7
                                    If ((y * lWidth) + x + ix) >= xStart And ((y * lWidth) + x + ix) <= xEnd Then
                                        SetPixel(((y + iy) * lWidth) + x + ix, lForeColor)
                                    End If
                                Next
                            End If
                        Next
                    Next
                Next
            Case GRE_HATCHSTYLE.HS_OUTLINEDDIAMOND
                xDistance = 7
                For y = v_Y1 To v_Y2
                    xStart = (y * lWidth) + v_X1
                    xEnd = (y * lWidth) + v_X2
                    For x = (xStart - xDistance) To (xEnd + xDistance) Step xDistance
                        If (x + xOffset) >= xStart And (x + xOffset) <= xEnd Then
                            SetPixel(x + xOffset, lForeColor)
                        End If
                    Next
                    xOffset = xOffset + 1
                    If xOffset > (xDistance - 1) Then
                        xOffset = 0
                    End If
                Next
                For y = v_Y1 To v_Y2
                    xStart = (y * lWidth) + v_X1
                    xEnd = (y * lWidth) + v_X2
                    For x = (xStart - xDistance) To (xEnd + xDistance) Step xDistance
                        If (x - xOffset) >= xStart And (x - xOffset) <= xEnd Then
                            SetPixel(x - xOffset, lForeColor)
                        End If
                    Next
                    xOffset = xOffset + 1
                    If xOffset > (xDistance - 1) Then
                        xOffset = 0
                    End If
                Next

            Case GRE_HATCHSTYLE.HS_SOLIDDIAMOND
                aRect = New Integer(,) { _
                {0, 1, 1, 1, 1, 1, 0, 0}, _
                {0, 0, 1, 1, 1, 0, 0, 0}, _
                {0, 0, 0, 1, 0, 0, 0, 0}, _
                {0, 0, 0, 0, 0, 0, 0, 0}, _
                {0, 0, 0, 1, 0, 0, 0, 0}, _
                {0, 0, 1, 1, 1, 0, 0, 0}, _
                {0, 1, 1, 1, 1, 1, 0, 0}, _
                {1, 1, 1, 1, 1, 1, 1, 0} _
                }
                mp_DrawRaster(v_X1, v_Y1, v_X2, v_Y2, lWidth, aRect, lForeColor)
        End Select
        GetBitmap.Invalidate()
    End Sub

    Private Sub mp_DrawRaster(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_x2 As Integer, ByVal v_Y2 As Integer, ByVal lWidth As Integer, ByVal aRect As Integer(,), ByVal lForeColor As Integer)
        Dim xDistance As Integer
        Dim yDistance As Integer
        Dim yEnd As Integer
        Dim xStart As Integer
        Dim xEnd As Integer
        Dim lVisible As Integer
        Dim ix As Integer
        Dim iy As Integer
        xDistance = UBound(aRect, 2) + 1
        yDistance = UBound(aRect, 1) + 1
        yEnd = (v_Y2 * lWidth) + v_x2
        For y = v_Y1 To v_Y2 Step yDistance
            xStart = (y * lWidth) + v_X1
            xEnd = (y * lWidth) + v_x2
            For x = v_X1 To v_x2 Step xDistance
                For ix = 0 To xDistance - 1
                    For iy = 0 To yDistance - 1
                        lVisible = aRect(iy, ix)
                        If lVisible = 1 Then
                            If (y + iy) <= v_Y2 Then
                                If ((y * lWidth) + x + ix) >= xStart And ((y * lWidth) + x + ix) <= xEnd Then
                                    SetPixel(((y + iy) * lWidth) + x + ix, lForeColor)
                                End If
                            End If
                        End If
                    Next
                Next
            Next
        Next
    End Sub

    Public Sub ResetFocusRectangle()

    End Sub

    Public Sub DrawReversibleFrameEx()
        DrawReversibleFrame(mp_lFocusLeft, mp_lFocusTop, mp_lFocusRight, mp_lFocusBottom)
    End Sub

    Public Sub DrawReversibleLine(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer)
        If mp_lSelectionLineIndex = -1 Then
            mp_oSelectionLine.X1 = v_X1
            mp_oSelectionLine.X2 = v_X2
            mp_oSelectionLine.Y1 = v_Y1
            mp_oSelectionLine.Y2 = v_Y2
            mp_oSelectionLine.Stroke = New SolidColorBrush(mp_oControl.SelectionColor)
            mp_oSelectionLine.StrokeThickness = 1
            mp_oSelectionLine.IsHitTestVisible = False
            mp_oControl.oCanvas.Children.Add(mp_oSelectionLine)
            mp_lSelectionLineIndex = mp_oControl.oCanvas.Children.Count() - 1
        Else
            mp_oSelectionLine.X1 = v_X1
            mp_oSelectionLine.X2 = v_X2
            mp_oSelectionLine.Y1 = v_Y1
            mp_oSelectionLine.Y2 = v_Y2
        End If
    End Sub

    Public Sub EraseReversibleLines()
        If mp_lSelectionLineIndex > -1 Then
            mp_oControl.oCanvas.Children.Remove(mp_oSelectionLine)
            mp_lSelectionLineIndex = -1
        End If
    End Sub

    Public Sub DrawReversibleFrame(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer)
        If ((v_X2 - v_X1 + 1) <= 1) Then
            Return
        End If
        If ((v_Y2 - v_Y1 + 1) <= 1) Then
            Return
        End If
        If mp_lSelectionRectangleIndex = -1 Then
            mp_oSelectionRectangle.Width = v_X2 - v_X1 + 1
            mp_oSelectionRectangle.Height = v_Y2 - v_Y1 + 1
            mp_oSelectionRectangle.SetValue(Canvas.LeftProperty, CDbl(v_X1))
            mp_oSelectionRectangle.SetValue(Canvas.TopProperty, CDbl(v_Y1))
            mp_oSelectionRectangle.Stroke = New SolidColorBrush(mp_oControl.SelectionColor)
            mp_oSelectionRectangle.StrokeThickness = 1
            mp_oSelectionRectangle.Name = "SelectionRectangle"
            mp_oSelectionRectangle.IsHitTestVisible = False
            mp_oControl.oCanvas.Children.Add(mp_oSelectionRectangle)
            mp_lSelectionRectangleIndex = mp_oControl.oCanvas.Children.Count() - 1
        Else
            mp_oSelectionRectangle.Width = v_X2 - v_X1 + 1
            mp_oSelectionRectangle.Height = v_Y2 - v_Y1 + 1
            mp_oSelectionRectangle.SetValue(Canvas.LeftProperty, CDbl(v_X1))
            mp_oSelectionRectangle.SetValue(Canvas.TopProperty, CDbl(v_Y1))
        End If
    End Sub

    Public Sub EraseReversibleFrames()
        If mp_lSelectionRectangleIndex > -1 Then
            mp_oControl.oCanvas.Children.Remove(mp_oSelectionRectangle)
            mp_lSelectionRectangleIndex = -1
        End If
    End Sub

    Friend Sub mp_DrawItemI(ByRef oTask As clsTask, ByVal sStyleIndex As String, ByVal Selected As Boolean, ByRef v_oStyle As clsStyle)
        Dim oStyle As clsStyle
        Dim oMilestoneStyle As clsMilestoneStyle
        If (v_oStyle Is Nothing) Then
            If mp_oControl.StrLib.StrIsNumeric(sStyleIndex) Then
                If mp_oControl.StrLib.StrCLng(sStyleIndex) < 0 Or mp_oControl.StrLib.StrCLng(sStyleIndex) > mp_oControl.Styles.Count Then
                    mp_oControl.mp_ErrorReport(SYS_ERRORS.STYLE_INVALID_INDEX, "Style object element not found when preparing to draw, invalid index", "mp_DrawItemI")
                    Return
                End If
            Else
                If mp_oControl.Styles.oCollection.m_bDoesKeyExist(sStyleIndex) = False Then
                    mp_oControl.mp_ErrorReport(SYS_ERRORS.STYLE_INVALID_KEY, "Style object element not found when preparing to draw, invalid key (""" & sStyleIndex & """)", "mp_DrawItemI")
                    Return
                End If
            End If
            oStyle = mp_oControl.Styles.FItem(sStyleIndex)
        Else
            oStyle = v_oStyle
        End If
        Select Case oStyle.Appearance
            Case E_STYLEAPPEARANCE.SA_FLAT
                oMilestoneStyle = oStyle.MilestoneStyle
                DrawFigure(mp_oControl.MathLib.GetXCoordinateFromDate(oTask.StartDate), oTask.Top, oTask.Bottom - oTask.Top, oTask.Bottom - oTask.Top, oMilestoneStyle.ShapeIndex, oMilestoneStyle.BorderColor, oMilestoneStyle.FillColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            Case E_STYLEAPPEARANCE.SA_GRAPHICAL
                If oStyle.MilestoneStyle.Image Is Nothing Then

                Else
                    DrawImage(oStyle.MilestoneStyle.Image, oStyle.ImageAlignmentHorizontal, oStyle.ImageAlignmentVertical, oStyle.ImageXMargin, oStyle.ImageYMargin, oTask.Left, oTask.Right, oTask.Top, oTask.Bottom, oStyle.UseMask)
                End If
            Case Else
                oMilestoneStyle = oStyle.MilestoneStyle
                DrawFigure(mp_oControl.MathLib.GetXCoordinateFromDate(oTask.StartDate), oTask.Top, oTask.Bottom - oTask.Top, oTask.Bottom - oTask.Top, oMilestoneStyle.ShapeIndex, oMilestoneStyle.BorderColor, oMilestoneStyle.FillColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
        End Select
        mp_DrawItemText(oTask.Left, oTask.Top, oTask.Right, oTask.Bottom, oTask.LeftTrim, oTask.RightTrim, oStyle, oTask.Text)
        If oStyle.SelectionRectangleStyle.Visible = True And Selected Then
            If oStyle.SelectionRectangleStyle.Mode = E_SELECTIONRECTANGLEMODE.SRM_DOTTED Then
                DrawFocusRectangle(oTask.Left, oTask.Top, oTask.Right, oTask.Bottom)
            ElseIf oStyle.SelectionRectangleStyle.Mode = E_SELECTIONRECTANGLEMODE.SRM_COLOR Then
                DrawLine(oTask.Left, oTask.Top, oTask.Right, oTask.Bottom, GRE_LINETYPE.LT_BORDER, oStyle.SelectionRectangleStyle.Color, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.SelectionRectangleStyle.BorderWidth)
            End If
        End If
    End Sub

    Friend Sub mp_DrawItemEx(ByVal v_lLeft As Integer, ByVal v_lTop As Integer, ByVal v_lRight As Integer, ByVal v_lBottom As Integer, ByVal sText As String, ByVal v_bIsSelected As Boolean, ByRef v_oImage As Image, ByVal v_lLeftTrim As Integer, ByVal v_lRightTrim As Integer, ByRef v_oStyle As clsStyle, ByVal clrBackColor As Color, ByVal clrForeColor As Color, ByVal clrStartGradientColor As Color, ByVal clrEndGradientColor As Color, ByVal clrHatchBackColor As Color, ByVal clrHatchForeColor As Color)
        Dim oStyle As clsStyle
        Dim oTaskStyle As clsTaskStyle
        If (v_oStyle Is Nothing) Then
            mp_oControl.mp_ErrorReport(SYS_ERRORS.STYLE_NULL, "Style object is null when preparing to draw.", "mp_DrawItemEx")
            Return
        Else
            oStyle = v_oStyle
        End If
        oTaskStyle = oStyle.TaskStyle
        Select Case oStyle.Appearance
            Case E_STYLEAPPEARANCE.SA_RAISED
                DrawEdge(v_lLeft, v_lTop, v_lRight, v_lBottom, clrBackColor, oStyle.ButtonStyle, GRE_EDGETYPE.ET_RAISED, True, v_oStyle)
            Case E_STYLEAPPEARANCE.SA_SUNKEN
                DrawEdge(v_lLeft, v_lTop, v_lRight, v_lBottom, clrBackColor, oStyle.ButtonStyle, GRE_EDGETYPE.ET_SUNKEN, True, v_oStyle)
            Case E_STYLEAPPEARANCE.SA_FLAT
                Dim lTop As Integer
                Dim lBottom As Integer
                lTop = v_lTop
                lBottom = v_lBottom
                Select Case oStyle.FillMode
                    Case GRE_FILLMODE.FM_COMPLETELYFILLED
                    Case GRE_FILLMODE.FM_UPPERHALFFILLED
                        lBottom = System.Convert.ToInt32(v_lTop + ((v_lBottom - v_lTop) / 2))
                    Case GRE_FILLMODE.FM_LOWERHALFFILLED
                        lTop = System.Convert.ToInt32(v_lBottom - ((v_lBottom - v_lTop) / 2))
                End Select
                If (oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID) Then
                    DrawLine(v_lLeft, lTop, v_lRight, lBottom, GRE_LINETYPE.LT_FILLED, clrBackColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                ElseIf (oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT) Then
                    GradientFill(v_lLeft, lTop, v_lRight, lBottom, clrStartGradientColor, clrEndGradientColor, oStyle.GradientFillMode)
                ElseIf (oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_PATTERN) Then
                    DrawPattern(v_lLeft, lTop, v_lRight, lBottom, clrBackColor, oStyle.Pattern, oStyle.PatternFactor)
                ElseIf (oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_HATCH) Then
                    HatchFill(v_lLeft, lTop, v_lRight, lBottom, clrHatchForeColor, clrHatchBackColor, oStyle.HatchStyle)
                End If
                If oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE Then
                    DrawLine(v_lLeft, lTop, v_lRight, lBottom, GRE_LINETYPE.LT_BORDER, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.BorderWidth)
                ElseIf oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM Then
                    If oStyle.CustomBorderStyle.Left = True Then
                        DrawLine(v_lLeft, lTop, v_lLeft, lBottom, GRE_LINETYPE.LT_NORMAL, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.BorderWidth)
                    End If
                    If oStyle.CustomBorderStyle.Top = True Then
                        DrawLine(v_lLeft, lTop, v_lRight, lTop, GRE_LINETYPE.LT_NORMAL, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.BorderWidth)
                    End If
                    If oStyle.CustomBorderStyle.Right = True Then
                        DrawLine(v_lRight, lTop, v_lRight, lBottom, GRE_LINETYPE.LT_NORMAL, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.BorderWidth)
                    End If
                    If oStyle.CustomBorderStyle.Bottom = True Then
                        DrawLine(v_lLeft, lBottom, v_lRight, lBottom, GRE_LINETYPE.LT_NORMAL, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.BorderWidth)
                    End If
                End If
                DrawFigure(v_lRight, v_lTop, v_lBottom - v_lTop, v_lBottom - v_lTop, oTaskStyle.EndShapeIndex, oTaskStyle.EndBorderColor, oTaskStyle.EndFillColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawFigure(v_lLeft, v_lTop, v_lBottom - v_lTop, v_lBottom - v_lTop, oTaskStyle.StartShapeIndex, oTaskStyle.StartBorderColor, oTaskStyle.StartFillColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            Case E_STYLEAPPEARANCE.SA_CELL
                DrawLine(v_lLeft, v_lTop, v_lRight, v_lBottom, GRE_LINETYPE.LT_FILLED, clrBackColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(v_lLeft, v_lBottom, v_lRight, v_lBottom, GRE_LINETYPE.LT_NORMAL, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.BorderWidth)
            Case E_STYLEAPPEARANCE.SA_GRAPHICAL
                If oTaskStyle.MiddleImage Is Nothing Or oTaskStyle.StartImage Is Nothing Or oTaskStyle.EndImage Is Nothing Then

                Else
                    Dim lImageHeight As Integer
                    Dim lImageWidth As Integer
                    lImageHeight = System.Convert.ToInt32(oTaskStyle.MiddleImage.Height)
                    lImageWidth = System.Convert.ToInt32(oTaskStyle.MiddleImage.Width)
                    TileImageHorizontal(oTaskStyle.MiddleImage, v_lLeft, v_lTop, v_lRight, v_lBottom, oStyle.UseMask)
                    '// Exit if the start and end sections don't fit
                    If (v_lRight - v_lLeft) > (lImageWidth * 2) Then
                        '// Left Section
                        PaintImage(oTaskStyle.StartImage, v_lLeft, v_lTop, v_lLeft + lImageWidth, v_lTop + lImageHeight, 0, 0, oStyle.UseMask)
                        '// Right Section
                        PaintImage(oTaskStyle.EndImage, v_lRight - lImageWidth, v_lTop, v_lRight, v_lTop + lImageHeight, 0, 0, oStyle.UseMask)
                    End If
                End If
        End Select
        If Not (v_oImage Is Nothing) Then
            DrawImage(v_oImage, oStyle.ImageAlignmentHorizontal, oStyle.ImageAlignmentVertical, oStyle.ImageXMargin, oStyle.ImageYMargin, v_lLeft, v_lRight, v_lTop, v_lBottom, oStyle.UseMask)
        End If
        mp_DrawItemText(v_lLeft, v_lTop, v_lRight, v_lBottom, v_lLeftTrim, v_lRightTrim, oStyle, sText)
        If oStyle.SelectionRectangleStyle.Visible = True And v_bIsSelected Then
            mp_DrawSelectionRectangle(v_lLeft, v_lTop, v_lRight, v_lBottom, oStyle)
        End If
    End Sub

    Friend Sub mp_DrawSelectionRectangle(ByVal v_lLeft As Integer, ByVal v_lTop As Integer, ByVal v_lRight As Integer, ByVal v_lBottom As Integer, ByVal oStyle As clsStyle)
        If oStyle.SelectionRectangleStyle.Mode = E_SELECTIONRECTANGLEMODE.SRM_DOTTED Then
            DrawFocusRectangle(v_lLeft + oStyle.SelectionRectangleStyle.OffsetLeft, v_lTop + oStyle.SelectionRectangleStyle.OffsetTop, v_lRight - oStyle.SelectionRectangleStyle.OffsetRight, v_lBottom - oStyle.SelectionRectangleStyle.OffsetBottom)
        ElseIf oStyle.SelectionRectangleStyle.Mode = E_SELECTIONRECTANGLEMODE.SRM_COLOR Then
            DrawLine(v_lLeft + oStyle.SelectionRectangleStyle.OffsetLeft, v_lTop + oStyle.SelectionRectangleStyle.OffsetTop, v_lRight - oStyle.SelectionRectangleStyle.OffsetRight, v_lBottom - oStyle.SelectionRectangleStyle.OffsetBottom, GRE_LINETYPE.LT_BORDER, oStyle.SelectionRectangleStyle.Color, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.SelectionRectangleStyle.BorderWidth)
        End If
    End Sub

    Friend Sub mp_DrawItem(ByVal v_lLeft As Integer, ByVal v_lTop As Integer, ByVal v_lRight As Integer, ByVal v_lBottom As Integer, ByVal sStyleIndex As String, ByVal sText As String, ByVal v_bIsSelected As Boolean, ByRef v_oImage As Image, ByVal v_lLeftTrim As Integer, ByVal v_lRightTrim As Integer, ByRef v_oStyle As clsStyle)
        Dim oStyle As clsStyle
        If (v_oStyle Is Nothing) Then
            If mp_oControl.StrLib.StrIsNumeric(sStyleIndex) Then
                If mp_oControl.StrLib.StrCLng(sStyleIndex) < 0 Or mp_oControl.StrLib.StrCLng(sStyleIndex) > mp_oControl.Styles.Count Then
                    mp_oControl.mp_ErrorReport(SYS_ERRORS.STYLE_INVALID_INDEX, "Style object element not found when preparing to draw, invalid index", "mp_DrawItem")
                    Return
                End If
            Else
                If mp_oControl.Styles.oCollection.m_bDoesKeyExist(sStyleIndex) = False Then
                    mp_oControl.mp_ErrorReport(SYS_ERRORS.STYLE_INVALID_KEY, "Style object element not found when preparing to draw, invalid key (""" & sStyleIndex & """)", "mp_DrawItem")
                    Return
                End If
            End If
            oStyle = mp_oControl.Styles.FItem(sStyleIndex)
        Else
            oStyle = v_oStyle
        End If
        mp_DrawItemEx(v_lLeft, v_lTop, v_lRight, v_lBottom, sText, v_bIsSelected, v_oImage, v_lLeftTrim, v_lRightTrim, oStyle, oStyle.BackColor, oStyle.ForeColor, oStyle.StartGradientColor, oStyle.EndGradientColor, oStyle.HatchBackColor, oStyle.HatchForeColor)
    End Sub

    Private Sub mp_DrawItemText(ByVal v_lLeft As Integer, ByVal v_lTop As Integer, ByVal v_lRight As Integer, ByVal v_lBottom As Integer, ByVal v_lLeftTrim As Integer, ByVal v_lRightTrim As Integer, ByRef oStyle As clsStyle, ByVal sText As String)
        Dim lTextLeft As Integer
        Dim lTextRight As Integer
        Dim lTextTop As Integer
        Dim lTextBottom As Integer
        If oStyle.TextVisible = False Then
            Return
        End If
        If sText = "" Then
            Return
        End If
        Select Case oStyle.TextPlacement
            Case E_TEXTPLACEMENT.SCP_OBJECTEXTENTSPLACEMENT
                If (oStyle.DrawTextInVisibleArea = False) Then
                    lTextLeft = v_lLeft
                    lTextRight = v_lRight
                Else
                    lTextLeft = v_lLeftTrim
                    lTextRight = v_lRightTrim
                End If
                lTextTop = v_lTop
                lTextBottom = v_lBottom
                If oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT Then
                    lTextLeft = v_lLeft + oStyle.TextXMargin
                End If
                If oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT Then
                    lTextRight = v_lRight - oStyle.TextXMargin
                End If
                If oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_TOP Then
                    lTextTop = v_lTop + oStyle.TextYMargin
                End If
                If oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM Then
                    lTextBottom = v_lBottom - oStyle.TextYMargin
                End If
                DrawAlignedText(lTextLeft, lTextTop, lTextRight, lTextBottom, sText, oStyle.TextAlignmentHorizontal, oStyle.TextAlignmentVertical, oStyle.ForeColor, oStyle.Font, oStyle.ClipText)
            Case E_TEXTPLACEMENT.SCP_OFFSETPLACEMENT
                DrawTextEx(v_lLeft + oStyle.TextFlags.OffsetLeft, v_lTop + oStyle.TextFlags.OffsetTop, v_lRight - oStyle.TextFlags.OffsetRight, v_lBottom - oStyle.TextFlags.OffsetBottom, sText, oStyle.TextFlags, oStyle.ForeColor, oStyle.Font, oStyle.ClipText)
            Case E_TEXTPLACEMENT.SCP_EXTERIORPLACEMENT
                If oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT Then
                    lTextLeft = v_lLeft - mp_oControl.mp_lStrWidth(sText, oStyle.Font) - oStyle.TextXMargin
                    lTextRight = v_lLeft - oStyle.TextXMargin + 1
                End If
                If oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT Then
                    lTextLeft = v_lRight + oStyle.TextXMargin
                    lTextRight = v_lRight + mp_oControl.mp_lStrWidth(sText, oStyle.Font) + oStyle.TextXMargin + 1
                End If
                If oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_CENTER Then
                    lTextLeft = v_lLeft
                    lTextRight = v_lRight + 1
                End If
                If oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_TOP Then
                    lTextTop = v_lTop - mp_oControl.mp_lStrHeight(sText, oStyle.Font) - oStyle.TextYMargin
                    lTextBottom = v_lTop - oStyle.TextYMargin + 1
                End If
                If oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM Then
                    lTextTop = v_lBottom + oStyle.TextYMargin
                    lTextBottom = v_lBottom + mp_oControl.mp_lStrHeight(sText, oStyle.Font) + oStyle.TextYMargin + 1
                End If
                If oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_CENTER Then
                    lTextTop = v_lTop
                    lTextBottom = v_lBottom + 1
                End If
                DrawAlignedText(lTextLeft, lTextTop, lTextRight, lTextBottom, sText, GRE_HORIZONTALALIGNMENT.HAL_LEFT, GRE_VERTICALALIGNMENT.VAL_TOP, oStyle.ForeColor, oStyle.Font, oStyle.ClipText)
        End Select
    End Sub

    Public Sub DrawButton(ByVal oRect As Rect, ByVal state As E_SCROLLBUTTONSTATE)
        Dim clrLightGray As Color = Color.FromArgb(255, 192, 192, 192)
        Dim clrMediumGray As Color = Color.FromArgb(255, 128, 128, 128)
        Dim clrDarkGray As Color = Color.FromArgb(255, 64, 64, 64)
        DrawLine(System.Convert.ToInt32(oRect.X + 1), System.Convert.ToInt32(oRect.Y + 1), System.Convert.ToInt32(oRect.X + oRect.Width - 3), System.Convert.ToInt32(oRect.Y + oRect.Height - 3), GRE_LINETYPE.LT_FILLED, clrLightGray, GRE_LINEDRAWSTYLE.LDS_SOLID)

        DrawLine(System.Convert.ToInt32(oRect.X), System.Convert.ToInt32(oRect.Y), System.Convert.ToInt32(oRect.X + oRect.Width - 2), System.Convert.ToInt32(oRect.Y), GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
        DrawLine(System.Convert.ToInt32(oRect.X), System.Convert.ToInt32(oRect.Y), System.Convert.ToInt32(oRect.X), System.Convert.ToInt32(oRect.Y + oRect.Height - 2), GRE_LINETYPE.LT_NORMAL, Colors.White, GRE_LINEDRAWSTYLE.LDS_SOLID)
        DrawLine(System.Convert.ToInt32(oRect.X), System.Convert.ToInt32(oRect.Y + oRect.Height - 1), System.Convert.ToInt32(oRect.X + oRect.Width - 1), System.Convert.ToInt32(oRect.Y + oRect.Height - 1), GRE_LINETYPE.LT_NORMAL, clrDarkGray, GRE_LINEDRAWSTYLE.LDS_SOLID)
        DrawLine(System.Convert.ToInt32(oRect.X + oRect.Width - 1), System.Convert.ToInt32(oRect.Y), System.Convert.ToInt32(oRect.X + oRect.Width - 1), System.Convert.ToInt32(oRect.Y + oRect.Height - 1), GRE_LINETYPE.LT_NORMAL, clrDarkGray, GRE_LINEDRAWSTYLE.LDS_SOLID)

        DrawLine(System.Convert.ToInt32(oRect.X + 1), System.Convert.ToInt32(oRect.Y + oRect.Height - 2), System.Convert.ToInt32(oRect.X + oRect.Width - 2), System.Convert.ToInt32(oRect.Y + oRect.Height - 2), GRE_LINETYPE.LT_NORMAL, clrMediumGray, GRE_LINEDRAWSTYLE.LDS_SOLID)
        DrawLine(System.Convert.ToInt32(oRect.X + oRect.Width - 2), System.Convert.ToInt32(oRect.Y + 1), System.Convert.ToInt32(oRect.X + oRect.Width - 2), System.Convert.ToInt32(oRect.Y + oRect.Height - 2), GRE_LINETYPE.LT_NORMAL, clrMediumGray, GRE_LINEDRAWSTYLE.LDS_SOLID)
    End Sub

    Friend Sub DrawScrollButton(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal width As Integer, ByVal height As Integer, ByVal button As E_SCROLLBUTTON, ByVal state As E_SCROLLBUTTONSTATE)
        Dim aRect As Integer(,)
        Dim lWidth As Integer = GetBitmap.PixelWidth
        mp_aPixels = GetBitmap.Pixels


        Select Case button
            Case E_SCROLLBUTTON.SB_RIGHT
                Select Case state
                    Case E_SCROLLBUTTONSTATE.BS_NORMAL
                        aRect = New Integer(,) { _
                        {4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 2}, _
                        {4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 2, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 2, 2, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 2, 2, 2, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 2, 2, 2, 2, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 2, 2, 2, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 2, 2, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 2, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 2}, _
                        {2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2} _
                        }
                        mp_DrawRasterButton(X1, Y1, X1 + width, Y1 + height, lWidth, aRect)
                    Case E_SCROLLBUTTONSTATE.BS_PUSHED
                        aRect = New Integer(,) { _
                        {3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 1}, _
                        {3, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 2, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 2, 2, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 2, 2, 2, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 2, 2, 2, 2, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 2, 2, 2, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 2, 2, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 2, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1} _
                        }
                        mp_DrawRasterButton(X1, Y1, X1 + width, Y1 + height, lWidth, aRect)
                    Case E_SCROLLBUTTONSTATE.BS_INACTIVE
                        aRect = New Integer(,) { _
                        {4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 2}, _
                        {4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 3, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 3, 3, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 3, 3, 3, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 3, 3, 3, 3, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 3, 3, 3, 1, 1, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 3, 3, 1, 1, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 3, 1, 1, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 1, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 2}, _
                        {2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2} _
                        }
                        mp_DrawRasterButton(X1, Y1, X1 + width, Y1 + height, lWidth, aRect)
                End Select

            Case E_SCROLLBUTTON.SB_LEFT
                Select Case state
                    Case E_SCROLLBUTTONSTATE.BS_NORMAL
                        aRect = New Integer(,) { _
                        {4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 2}, _
                        {4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 2, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 2, 2, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 2, 2, 2, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 2, 2, 2, 2, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 2, 2, 2, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 2, 2, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 2, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 2}, _
                        {2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2} _
                        }
                        mp_DrawRasterButton(X1, Y1, X1 + width, Y1 + height, lWidth, aRect)
                    Case E_SCROLLBUTTONSTATE.BS_PUSHED
                        aRect = New Integer(,) { _
                        {3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 1}, _
                        {3, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 2, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 2, 2, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 2, 2, 2, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 2, 2, 2, 2, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 2, 2, 2, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 2, 2, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 2, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1} _
                        }
                        mp_DrawRasterButton(X1, Y1, X1 + width, Y1 + height, lWidth, aRect)
                    Case E_SCROLLBUTTONSTATE.BS_INACTIVE
                        aRect = New Integer(,) { _
                        {4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 2}, _
                        {4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 3, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 3, 3, 1, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 3, 3, 3, 1, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 3, 3, 3, 3, 1, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 3, 3, 3, 1, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 3, 3, 1, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 3, 1, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 1, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 2}, _
                        {2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2} _
                        }
                        mp_DrawRasterButton(X1, Y1, X1 + width, Y1 + height, lWidth, aRect)
                End Select

            Case E_SCROLLBUTTON.SB_UP
                Select Case state
                    Case E_SCROLLBUTTONSTATE.BS_NORMAL
                        aRect = New Integer(,) { _
                        {4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 2}, _
                        {4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 2, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 2, 2, 2, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 2, 2, 2, 2, 2, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 2, 2, 2, 2, 2, 2, 2, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 2}, _
                        {2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2} _
                        }
                        mp_DrawRasterButton(X1, Y1, X1 + width, Y1 + height, lWidth, aRect)
                    Case E_SCROLLBUTTONSTATE.BS_PUSHED
                        aRect = New Integer(,) { _
                        {3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 1}, _
                        {3, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 2, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 2, 2, 2, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 2, 2, 2, 2, 2, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 2, 2, 2, 2, 2, 2, 2, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1} _
                        }
                        mp_DrawRasterButton(X1, Y1, X1 + width, Y1 + height, lWidth, aRect)
                    Case E_SCROLLBUTTONSTATE.BS_INACTIVE
                        aRect = New Integer(,) { _
                        {4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 2}, _
                        {4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 3, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 3, 3, 3, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 3, 3, 3, 3, 3, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 3, 3, 3, 3, 3, 3, 3, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 1, 1, 1, 1, 1, 1, 1, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 2}, _
                        {2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2} _
                        }
                        mp_DrawRasterButton(X1, Y1, X1 + width, Y1 + height, lWidth, aRect)
                End Select

            Case E_SCROLLBUTTON.SB_DOWN
                Select Case state
                    Case E_SCROLLBUTTONSTATE.BS_NORMAL
                        aRect = New Integer(,) { _
                        {4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 2}, _
                        {4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 2, 2, 2, 2, 2, 2, 2, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 2, 2, 2, 2, 2, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 2, 2, 2, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 2, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 2}, _
                        {2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2} _
                        }
                        mp_DrawRasterButton(X1, Y1, X1 + width, Y1 + height, lWidth, aRect)
                    Case E_SCROLLBUTTONSTATE.BS_PUSHED
                        aRect = New Integer(,) { _
                        {3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 1}, _
                        {3, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 2, 2, 2, 2, 2, 2, 2, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 2, 2, 2, 2, 2, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 2, 2, 2, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 2, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {3, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 1}, _
                        {1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1} _
                        }
                        mp_DrawRasterButton(X1, Y1, X1 + width, Y1 + height, lWidth, aRect)
                    Case E_SCROLLBUTTONSTATE.BS_INACTIVE
                        aRect = New Integer(,) { _
                        {4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 2}, _
                        {4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 3, 3, 3, 3, 3, 3, 3, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 3, 3, 3, 3, 3, 1, 1, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 3, 3, 3, 1, 1, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 3, 1, 1, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 1, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 3, 2}, _
                        {4, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 2}, _
                        {2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2} _
                        }
                        mp_DrawRasterButton(X1, Y1, X1 + width, Y1 + height, lWidth, aRect)
                End Select

        End Select


    End Sub

    Private Sub mp_DrawRasterButton(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_x2 As Integer, ByVal v_Y2 As Integer, ByVal lWidth As Integer, ByVal aRect As Integer(,))
        Dim xDistance As Integer
        Dim yDistance As Integer
        Dim yEnd As Integer
        Dim xStart As Integer
        Dim xEnd As Integer
        Dim lColorIndex As Integer
        Dim ix As Integer
        Dim iy As Integer
        Dim lWhite As Integer = SWM_Color_To_Int32(Color.FromArgb(255, 255, 255, 255))
        Dim lBlack As Integer = SWM_Color_To_Int32(Color.FromArgb(255, 0, 0, 0))
        Dim lDarkGray As Integer = SWM_Color_To_Int32(Color.FromArgb(255, 128, 128, 128))
        Dim lLightGray As Integer = SWM_Color_To_Int32(Color.FromArgb(255, 192, 192, 192))
        Dim lColor As Integer
        xDistance = UBound(aRect, 2) + 1
        yDistance = UBound(aRect, 1) + 1
        yEnd = (v_Y2 * lWidth) + v_x2
        For y = v_Y1 To v_Y2 Step yDistance + 1
            xStart = (y * lWidth) + v_X1
            xEnd = (y * lWidth) + v_x2
            For x = v_X1 To v_x2 Step xDistance + 1
                For ix = 0 To xDistance - 1
                    For iy = 0 To yDistance - 1
                        lColorIndex = aRect(iy, ix)
                        If (y + iy) <= v_Y2 Then
                            If ((y * lWidth) + x + ix) >= xStart And ((y * lWidth) + x + ix) <= xEnd Then
                                lColor = 0
                                If lColorIndex = 1 Then
                                    lColor = lWhite
                                ElseIf lColorIndex = 2 Then
                                    lColor = lBlack
                                ElseIf lColorIndex = 3 Then
                                    lColor = lDarkGray
                                ElseIf lColorIndex = 4 Then
                                    lColor = lLightGray
                                End If
                                If lColor <> 0 Then
                                    SetPixel(((y + iy) * lWidth) + x + ix, lColor)
                                End If
                            End If
                        End If
                    Next
                Next
            Next
        Next
    End Sub

    Friend Sub DrawPoint(ByVal X As Integer, ByVal Y As Integer, ByVal clrColor As Color)
        mp_DrawLine(X, Y, X + 1, Y + 1, clrColor)
    End Sub

    Friend Sub mp_DrawArrow(ByVal v_X As Integer, ByVal v_Y As Integer, ByVal v_ArrowDirection As GRE_ARROWDIRECTION, ByVal v_ArrowSize As Integer, ByVal v_lColor As Color)
        Dim i As Integer
        Select Case v_ArrowDirection
            Case GRE_ARROWDIRECTION.AWD_LEFT
                DrawPoint(v_X, v_Y, v_lColor)
                For i = 1 To v_ArrowSize
                    v_X = v_X + 1
                    DrawLine(v_X, v_Y - i, v_X, v_Y + i, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                Next
            Case GRE_ARROWDIRECTION.AWD_RIGHT
                DrawPoint(v_X, v_Y, v_lColor)
                For i = 1 To v_ArrowSize
                    v_X = v_X - 1
                    DrawLine(v_X, v_Y - i, v_X, v_Y + i, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                Next
            Case GRE_ARROWDIRECTION.AWD_UP
                DrawPoint(v_X, v_Y, v_lColor)
                For i = 1 To v_ArrowSize
                    v_Y = v_Y + 1
                    DrawLine(v_X - i, v_Y, v_X + i, v_Y, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                Next
            Case GRE_ARROWDIRECTION.AWD_DOWN
                DrawPoint(v_X, v_Y, v_lColor)
                For i = 1 To v_ArrowSize
                    v_Y = v_Y - 1
                    DrawLine(v_X - i, v_Y, v_X + i, v_Y, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                Next
        End Select
    End Sub

End Class
