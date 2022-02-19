Option Explicit On

Public Class clsDrawing

    Private mp_oControl As ActiveGanttVBECtl

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
    End Sub

    
    'Public Function GraphicsInfo() As Graphics
    '    Return mp_oControl.clsG.oGraphics()
    'End Function

    
    Public Sub DrawLine(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal LineColor As Color, ByVal LineStyle As GRE_LINEDRAWSTYLE, ByVal LineWidth As Integer)
        mp_oControl.clsG.DrawLine(X1, Y1, X2, Y2, GRE_LINETYPE.LT_NORMAL, LineColor, LineStyle, LineWidth, True)
    End Sub

    
    Public Sub DrawBorder(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal LineColor As Color, ByVal LineStyle As GRE_LINEDRAWSTYLE, ByVal LineWidth As Integer)
        mp_oControl.clsG.DrawLine(X1, Y1, X2, Y2, GRE_LINETYPE.LT_BORDER, LineColor, LineStyle, LineWidth, True)
    End Sub

    
    Public Sub DrawRectangle(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal LineColor As Color, ByVal LineStyle As GRE_LINEDRAWSTYLE, ByVal LineWidth As Integer)
        mp_oControl.clsG.DrawLine(X1, Y1, X2, Y2, GRE_LINETYPE.LT_FILLED, LineColor, LineStyle, LineWidth, True)
    End Sub

    
    Public Sub DrawText(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal Text As String, ByVal TextFlags As clsTextFlags, ByVal TextColor As Color, ByVal TextFont As Font)
        mp_oControl.clsG.DrawTextEx(X1, Y1, X2, Y2, Text, TextFlags, TextColor, TextFont)
    End Sub


    Public Sub DrawAlignedText(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal Text As String, ByVal HorizontalAlignment As GRE_HORIZONTALALIGNMENT, ByVal VerticalAlignment As GRE_VERTICALALIGNMENT, ByVal TextColor As Color, ByVal TextFont As Font)
        mp_oControl.clsG.DrawAlignedText(X1, Y1, X2, Y2, Text, HorizontalAlignment, VerticalAlignment, TextColor, TextFont)
    End Sub

    
    Public Sub PaintImage(ByVal Image As Image, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
        mp_oControl.clsG.PaintImage(Image, X1, Y1, X2, Y2, 0, 0, True)
    End Sub

    'Public Sub ClearClipRegion()
    '    mp_oControl.clsG.ClearClipRegion()
    'End Sub

End Class
