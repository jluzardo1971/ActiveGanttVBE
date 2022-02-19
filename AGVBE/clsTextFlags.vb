'// ----------------------------------------------------------------------------------------
'//                              COPYRIGHT NOTICE
'// ----------------------------------------------------------------------------------------
'//
'// The Source Code Store LLC
'// ACTIVEGANTT SCHEDULER COMPONENT FOR VB.Net
'// Silverlight Control
'// Copyright (c) 2002-2009 The Source Code Store LLC
'//
'// All Rights Reserved. No parts of this file may be reproduced or transmitted in any
'// form or by any means without the written permission of the author.
'// ----------------------------------------------------------------------------------------
Option Explicit On 

Public Class clsTextFlags

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_yVerticalAlignment As GRE_VERTICALALIGNMENT
    Private mp_yHorizontalAlignment As GRE_HORIZONTALALIGNMENT
    Private mp_bWordWrap As Boolean
    Private mp_bRightToLeft As Boolean
    Private mp_lOffsetBottom As Integer
    Private mp_lOffsetLeft As Integer
    Private mp_lOffsetRight As Integer
    Private mp_lOffsetTop As Integer

    Public Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_yVerticalAlignment = GRE_VERTICALALIGNMENT.VAL_TOP
        mp_yHorizontalAlignment = GRE_HORIZONTALALIGNMENT.HAL_LEFT
        mp_bWordWrap = False
        mp_bRightToLeft = False
        mp_lOffsetBottom = 0
        mp_lOffsetLeft = 0
        mp_lOffsetRight = 0
        mp_lOffsetTop = 0
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Property VerticalAlignment() As GRE_VERTICALALIGNMENT
        Get
            Return mp_yVerticalAlignment
        End Get
        Set(ByVal Value As GRE_VERTICALALIGNMENT)
            mp_yVerticalAlignment = Value
        End Set
    End Property

    Public Property HorizontalAlignment() As GRE_HORIZONTALALIGNMENT
        Get
            Return mp_yHorizontalAlignment
        End Get
        Set(ByVal Value As GRE_HORIZONTALALIGNMENT)
            mp_yHorizontalAlignment = Value
        End Set
    End Property

    Public Property WordWrap() As Boolean
        Get
            Return mp_bWordWrap
        End Get
        Set(ByVal Value As Boolean)
            mp_bWordWrap = Value
        End Set
    End Property

    Public Property RightToLeft() As Boolean
        Get
            Return mp_bRightToLeft
        End Get
        Set(ByVal Value As Boolean)
            mp_bRightToLeft = Value
        End Set
    End Property

    Public Property OffsetBottom() As Integer
        Get
            Return mp_lOffsetBottom
        End Get
        Set(ByVal Value As Integer)
            mp_lOffsetBottom = Value
        End Set
    End Property

    Public Property OffsetLeft() As Integer
        Get
            Return mp_lOffsetLeft
        End Get
        Set(ByVal Value As Integer)
            mp_lOffsetLeft = Value
        End Set
    End Property

    Public Property OffsetRight() As Integer
        Get
            Return mp_lOffsetRight
        End Get
        Set(ByVal Value As Integer)
            mp_lOffsetRight = Value
        End Set
    End Property

    Public Property OffsetTop() As Integer
        Get
            Return mp_lOffsetTop
        End Get
        Set(ByVal Value As Integer)
            mp_lOffsetTop = Value
        End Set
    End Property

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "TextFlags")
        oXML.InitializeWriter()
        oXML.WriteProperty("HorizontalAlignment", mp_yHorizontalAlignment)
        oXML.WriteProperty("OffsetBottom", mp_lOffsetBottom)
        oXML.WriteProperty("OffsetLeft", mp_lOffsetLeft)
        oXML.WriteProperty("OffsetRight", mp_lOffsetRight)
        oXML.WriteProperty("OffsetTop", mp_lOffsetTop)
        oXML.WriteProperty("RightToLeft", mp_bRightToLeft)
        oXML.WriteProperty("VerticalAlignment", mp_yVerticalAlignment)
        oXML.WriteProperty("WordWrap", mp_bWordWrap)
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "TextFlags")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("HorizontalAlignment", mp_yHorizontalAlignment)
        oXML.ReadProperty("OffsetBottom", mp_lOffsetBottom)
        oXML.ReadProperty("OffsetLeft", mp_lOffsetLeft)
        oXML.ReadProperty("OffsetRight", mp_lOffsetRight)
        oXML.ReadProperty("OffsetTop", mp_lOffsetTop)
        oXML.ReadProperty("RightToLeft", mp_bRightToLeft)
        oXML.ReadProperty("VerticalAlignment", mp_yVerticalAlignment)
        oXML.ReadProperty("WordWrap", mp_bWordWrap)
    End Sub


End Class


