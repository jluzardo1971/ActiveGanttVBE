Option Explicit On

Public Class clsNode

    Private mp_oControl As ActiveGanttVBECtl
    Friend mp_bExpanded As Boolean
    Private mp_sTag As String
    Private mp_oImage As Image
    Private mp_oExpandedImage As Image
    Private mp_oSelectedImage As Image
    Private mp_bChecked As Boolean
    Private mp_lDepth As Integer
    Private mp_oRow As clsRow
    Private mp_bCheckBoxVisible As Boolean
    Private mp_bImageVisible As Boolean
    Friend mp_oParent As clsNode
    Private mp_sExpandedImageTag As String
    Private mp_sSelectedImageTag As String
    Private mp_sImageTag As String
    Private mp_sStyleIndex As String
    Private mp_oStyle As clsStyle
    Friend mp_lTextLeft As Double
    Friend mp_lTextTop As Double
    Friend mp_lTextRight As Double
    Friend mp_lTextBottom As Double
    Private mp_bAllowTextEdit As Boolean

    Friend Sub New(ByVal Value As ActiveGanttVBECtl, ByVal oRow As clsRow)
        mp_oControl = Value
        mp_oRow = oRow
        mp_lDepth = 0
        mp_bExpanded = True
        mp_oImage = Nothing
        mp_oExpandedImage = Nothing
        mp_oSelectedImage = Nothing
        mp_bChecked = False
        mp_bCheckBoxVisible = False
        mp_bImageVisible = False
        mp_sTag = ""
        mp_sExpandedImageTag = ""
        mp_sSelectedImageTag = ""
        mp_sImageTag = ""
        mp_sStyleIndex = "DS_NODE"
        mp_oStyle = mp_oControl.Styles.FItem("DS_NODE")
        mp_bAllowTextEdit = False
    End Sub

    Public Property AllowTextEdit() As Boolean
        Get
            Return mp_bAllowTextEdit
        End Get
        Set(ByVal value As Boolean)
            mp_bAllowTextEdit = value
        End Set
    End Property

    Public ReadOnly Property Row() As clsRow
        Get
            Return mp_oRow
        End Get
    End Property

    
    Public Property CheckBoxVisible() As Boolean
        Get
            If mp_oControl.Treeview.CheckBoxes = True Then
                Return mp_bCheckBoxVisible
            Else
                Return False
            End If
        End Get
        Set(ByVal Value As Boolean)
            mp_bCheckBoxVisible = Value
        End Set
    End Property

    
    Public Property ImageVisible() As Boolean
        Get
            If mp_oControl.Treeview.Images = True Then
                If mp_oImage Is Nothing Then
                    Return False
                Else
                    Return mp_bImageVisible
                End If
            Else
                Return False
            End If
        End Get
        Set(ByVal Value As Boolean)
            mp_bImageVisible = Value
        End Set
    End Property

    
    Public ReadOnly Property Left() As Integer
        Get
            If Hidden = True Then
                Return 0
            Else
                Return mp_oControl.Treeview.Left + mp_oControl.Treeview.Indentation + (Depth * mp_oControl.Treeview.Indentation)
            End If
        End Get
    End Property

    
    Public ReadOnly Property Top() As Integer
        Get
            Return mp_oRow.Top
        End Get
    End Property

    
    Public ReadOnly Property Bottom() As Integer
        Get
            Return mp_oRow.Bottom
        End Get
    End Property

    
    Public Function NextSibling() As clsNode
        Dim lIndex As Integer = Index
        Dim oNode As clsNode = Nothing
        Dim oRow As clsRow = Nothing
        If (lIndex + 1) <= mp_oControl.Rows.Count Then
            For lIndex = (Index + 1) To mp_oControl.Rows.Count
                oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex), clsRow)
                oNode = oRow.Node
                If oNode.Depth = Depth Then
                    Return oNode
                ElseIf oNode.Depth = Depth - 1 Then
                    Return Nothing
                End If
            Next
            Return Nothing
        Else
            Return Nothing 'This Node is the Last Node
        End If
    End Function

    
    Public Function PreviousSibling() As clsNode
        Dim lIndex As Integer = Index
        Dim oNode As clsNode = Nothing
        Dim oRow As clsRow = Nothing
        If lIndex > 1 Then
            For lIndex = (Index - 1) To 1 Step -1
                oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex), clsRow)
                oNode = oRow.Node
                If oNode.Depth = Depth Then
                    Return oNode
                ElseIf oNode.Depth = Depth - 1 Then
                    Return Nothing
                End If
            Next
            Return Nothing
        Else
            Return Nothing 'This node is the First Node
        End If
    End Function

    
    Public Function IsFirstSibling() As Boolean
        Dim oNode As clsNode = Nothing
        oNode = PreviousSibling()
        If oNode Is Nothing Then
            Return True
        Else
            Return False
        End If
    End Function

    
    Public Function FirstSibling() As clsNode
        Dim oNode As clsNode = Nothing
        Dim oReturnNode As clsNode = Nothing
        oReturnNode = Me
        oNode = PreviousSibling()
        Do While Not oNode Is Nothing
            oReturnNode = oNode
            oNode = oReturnNode.PreviousSibling()
        Loop
        Return oReturnNode
    End Function

    
    Public Function IsLastSibling() As Boolean
        Dim oNode As clsNode = Nothing
        oNode = NextSibling()
        If oNode Is Nothing Then
            Return True
        Else
            Return False
        End If
    End Function

    
    Public Function LastSibling() As clsNode
        Dim oNode As clsNode = Nothing
        Dim oReturnNode As clsNode = Nothing
        oReturnNode = Me
        oNode = NextSibling()
        Do While Not oNode Is Nothing
            oReturnNode = oNode
            oNode = oReturnNode.NextSibling()
        Loop
        Return oReturnNode
    End Function

    
    Public Function Child() As clsNode
        Dim lIndex As Integer = Index + 1
        Dim oNode As clsNode = Nothing
        Dim oRow As clsRow = Nothing
        If lIndex <= mp_oControl.Rows.Count Then
            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex), clsRow)
            oNode = oRow.Node
            If oNode.Depth = (Depth + 1) Then
                Return oNode
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function

    
    Public Function Parent() As clsNode
        Return mp_oParent
    End Function

    
    Public Function IsRoot() As Boolean
        If Parent() Is Nothing Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function FullPath() As String
        Dim oNode As clsNode = Nothing
        Dim sReturn As String = ""
        oNode = Parent()
        Do While Not oNode Is Nothing
            sReturn = oNode.Text & mp_oControl.Treeview.PathSeparator & sReturn
            oNode = oNode.Parent()
        Loop
        Return sReturn & Text
    End Function

    
    Public Property Depth() As Integer
        Get
            Return mp_lDepth
        End Get
        Set(ByVal Value As Integer)
            mp_lDepth = Value
        End Set
    End Property
    
    Public Function Children() As Integer
        Dim lIndex As Integer
        Dim oNode As clsNode = Nothing
        Dim oRow As clsRow = Nothing
        Dim lReturn As Integer
        For lIndex = (mp_oRow.Index + 1) To mp_oControl.Rows.Count
            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex), clsRow)
            oNode = oRow.Node
            If Depth < oNode.Depth Then
                If Depth + 1 = oNode.Depth Then
                    lReturn = lReturn + 1
                End If
            Else
                Exit For
            End If
        Next lIndex
        Return lReturn
    End Function

    
    Public Property Expanded() As Boolean
        Get
            If Children() > 0 Then
                Return mp_bExpanded
            Else
                Return True
            End If
        End Get
        Set(ByVal Value As Boolean)
            mp_bExpanded = Value
            mp_oControl.VerticalScrollBar.Update()
        End Set
    End Property

    
    Public Property Selected() As Boolean
        Get
            Return (Index = mp_oControl.SelectedRowIndex)
        End Get
        Set(ByVal Value As Boolean)
            If Value = True Then
                mp_oControl.SelectedRowIndex = Index
            Else
                If mp_oControl.SelectedRowIndex = Index Then
                    mp_oControl.SelectedRowIndex = 0
                End If
            End If
        End Set
    End Property

    
    Public Property ExpandedImage() As Image
        Get
            Return mp_oExpandedImage
        End Get
        Set(ByVal Value As Image)
            mp_oExpandedImage = Value
        End Set
    End Property

    
    Public Property SelectedImage() As Image
        Get
            Return mp_oSelectedImage
        End Get
        Set(ByVal Value As Image)
            mp_oSelectedImage = Value
        End Set
    End Property

    
    Public Property Image() As Image
        Get
            Return mp_oImage
        End Get
        Set(ByVal Value As Image)
            mp_oImage = Value
        End Set
    End Property

    
    Public Property Tag() As String
        Get
            Return mp_sTag
        End Get
        Set(ByVal Value As String)
            mp_sTag = Value
        End Set
    End Property

    
    Public Property Text() As String
        Get
            Return mp_oRow.Text
        End Get
        Set(ByVal Value As String)
            mp_oRow.Text = Value
        End Set
    End Property

    
    Public Property Checked() As Boolean
        Get
            Return mp_bChecked
        End Get
        Set(ByVal Value As Boolean)
            mp_bChecked = Value
        End Set
    End Property

    
    Public Property Height() As Integer
        Get
            Return mp_oRow.Height
        End Get
        Set(ByVal Value As Integer)
            mp_oRow.Height = Value
        End Set
    End Property

    Friend ReadOnly Property YCenter() As Integer
        Get
            Return System.Convert.ToInt32(Top + (Height / 2))
        End Get
    End Property

    Friend ReadOnly Property mt_TextLeft() As Integer
        Get
            If CheckBoxVisible = False And ImageVisible = False Then
                Return Left + 10
            ElseIf CheckBoxVisible = True And ImageVisible = True Then
                Return System.Convert.ToInt32(Left + 28 + mp_oImage.Width + 5)
            ElseIf CheckBoxVisible = True Then
                Return Left + 28
            ElseIf ImageVisible = True Then
                Return System.Convert.ToInt32(Left + 10 + mp_oImage.Width + 5)
            End If
            Return 0
        End Get
    End Property

    Friend ReadOnly Property mt_TextTop() As Integer
        Get
            Return YCenter - (mp_oControl.mp_lStrHeight(Text, mp_oStyle.Font)) + 3
        End Get
    End Property

    Friend ReadOnly Property mt_TextRight() As Integer
        Get
            Return mt_TextLeft + (mp_oControl.mp_lStrWidth(Text, mp_oStyle.Font)) + 10
        End Get
    End Property

    Friend ReadOnly Property mt_TextBottom() As Integer
        Get
            Return YCenter + (mp_oControl.mp_lStrHeight(Text, mp_oStyle.Font)) - 3
        End Get
    End Property

    Friend ReadOnly Property CheckBoxLeft() As Integer
        Get
            If CheckBoxVisible = True And mp_oControl.Treeview.CheckBoxes = True Then
                Return Left + 14
            Else
                Return 0
            End If
        End Get
    End Property

    Friend ReadOnly Property ImageLeft() As Integer
        Get
            If ImageVisible = True And mp_oControl.Treeview.Images = True Then
                Return System.Convert.ToInt32(mt_TextLeft - 3 - mp_oImage.Width)
            Else
                Return 0
            End If
        End Get
    End Property

    Friend ReadOnly Property ImageTop() As Integer
        Get
            If ImageVisible = True And mp_oControl.Treeview.Images = True Then
                Return System.Convert.ToInt32(YCenter - (mp_oImage.Height / 2))
            Else
                Return 0
            End If
        End Get
    End Property

    Friend ReadOnly Property ImageRight() As Integer
        Get
            If ImageVisible = True And mp_oControl.Treeview.Images = True Then
                Return System.Convert.ToInt32(ImageLeft + mp_oImage.Width)
            Else
                Return 0
            End If
        End Get
    End Property

    Friend ReadOnly Property ImageBottom() As Integer
        Get
            If ImageVisible = True And mp_oControl.Treeview.Images = True Then
                Return System.Convert.ToInt32(ImageTop + mp_oImage.Height)
            Else
                Return 0
            End If
        End Get
    End Property

    Friend ReadOnly Property Index() As Integer
        Get
            Return mp_oRow.Index
        End Get
    End Property

    
    Public ReadOnly Property Hidden() As Boolean
        Get
            If mp_lDepth = 0 Then
                Return False
            End If
            Dim bHidden As Boolean
            bHidden = False
            bHidden = RecurseHidden(Me, bHidden)
            Return bHidden
        End Get
    End Property

    Private Function RecurseHidden(ByRef oNode As clsNode, ByVal bHidden As Boolean) As Boolean
        oNode = oNode.Parent()
        If Not oNode Is Nothing Then
            If oNode.mp_bExpanded = False Then
                bHidden = True
            End If
            bHidden = RecurseHidden(oNode, bHidden)
        End If
        Return bHidden
    End Function

    Public Property ExpandedImageTag() As String
        Get
            Return mp_sExpandedImageTag
        End Get
        Set(ByVal Value As String)
            mp_sExpandedImageTag = Value
        End Set
    End Property

    Public Property SelectedImageTag() As String
        Get
            Return mp_sSelectedImageTag
        End Get
        Set(ByVal Value As String)
            mp_sSelectedImageTag = Value
        End Set
    End Property

    Public Property ImageTag() As String
        Get
            Return mp_sImageTag
        End Get
        Set(ByVal Value As String)
            mp_sImageTag = Value
        End Set
    End Property

    Public Property StyleIndex() As String
        Get
            If mp_sStyleIndex = "DS_NODE" Then
                Return ""
            Else
                Return mp_sStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            If Value.Length = 0 Then Value = "DS_NODE"
            mp_sStyleIndex = Value
            mp_oStyle = mp_oControl.Styles.FItem(Value)
        End Set
    End Property

    Public ReadOnly Property Style() As clsStyle
        Get
            Return mp_oStyle
        End Get
    End Property

    
    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "Node")
        oXML.InitializeWriter()
        oXML.WriteProperty("CheckBoxVisible", mp_bCheckBoxVisible)
        oXML.WriteProperty("Checked", mp_bChecked)
        oXML.WriteProperty("Depth", mp_lDepth)
        oXML.WriteProperty("Expanded", mp_bExpanded)
        oXML.WriteProperty("ExpandedImage", mp_oExpandedImage)
        oXML.WriteProperty("Image", mp_oImage)
        oXML.WriteProperty("ImageVisible", mp_bImageVisible)
        oXML.WriteProperty("SelectedImage", mp_oSelectedImage)
        oXML.WriteProperty("Tag", mp_sTag)
        oXML.WriteProperty("ExpandedImageTag", mp_sExpandedImageTag)
        oXML.WriteProperty("SelectedImageTag", mp_sSelectedImageTag)
        oXML.WriteProperty("ImageTag", mp_sImageTag)
        oXML.WriteProperty("StyleIndex", mp_sStyleIndex)
        oXML.WriteProperty("AllowTextEdit", mp_bAllowTextEdit)
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "Node")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("CheckBoxVisible", mp_bCheckBoxVisible)
        oXML.ReadProperty("Checked", mp_bChecked)
        oXML.ReadProperty("Depth", mp_lDepth)
        oXML.ReadProperty("Expanded", mp_bExpanded)
        oXML.ReadProperty("ExpandedImage", mp_oExpandedImage)
        oXML.ReadProperty("Image", mp_oImage)
        oXML.ReadProperty("ImageVisible", mp_bImageVisible)
        oXML.ReadProperty("SelectedImage", mp_oSelectedImage)
        oXML.ReadProperty("Tag", mp_sTag)
        oXML.ReadProperty("ExpandedImageTag", mp_sExpandedImageTag)
        oXML.ReadProperty("SelectedImageTag", mp_sSelectedImageTag)
        oXML.ReadProperty("ImageTag", mp_sImageTag)
        oXML.ReadProperty("StyleIndex", mp_sStyleIndex)
        oXML.ReadProperty("AllowTextEdit", mp_bAllowTextEdit)
        StyleIndex = mp_sStyleIndex
    End Sub

End Class
