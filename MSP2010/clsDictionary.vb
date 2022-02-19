Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections.ObjectModel

Friend Class clsDictionary
    Inherits KeyedCollection(Of String, clsValuePair)

    Private mp_lKey As Integer = 1

    Public Sub New()
        MyBase.New()
    End Sub

    Public Overloads Sub Add(ByVal Value As Integer, ByVal Key As String)
        Dim oValuePair As New clsValuePair(Key, Value)
        MyBase.Add(oValuePair)
    End Sub

    Public Overloads Sub Add(ByVal Value As String)
        Dim oValuePair As New clsValuePair(mp_lKey.ToString(), System.Convert.ToInt32(Value))
        MyBase.Add(oValuePair)
        mp_lKey = mp_lKey + 1
    End Sub

    Public Overloads Function Contains(ByVal Key As String) As Boolean
        Return MyBase.Contains(Key)
    End Function

    Public ReadOnly Property StrItem(ByVal Index As Integer) As String
        Get
            Return MyBase.Item(Index).IntValue.ToString()
        End Get
    End Property

    Default Public Overloads Property Item(ByVal Key As String) As Integer
        Get
            Return MyBase.Item(Key).IntValue
        End Get
        Set(ByVal Value As Integer)
            MyBase.Item(Key).IntValue = Value
        End Set
    End Property

    Protected Overrides Function GetKeyForItem(ByVal item As clsValuePair) As String
        Return item.Key
    End Function

End Class

Friend Class clsValuePair
    Private mp_sKey As String
    Private mp_lIntValue As Integer

    Public Sub New(ByVal sKey As String, ByVal lIntValue As Integer)
        mp_sKey = sKey
        mp_lIntValue = lIntValue
    End Sub

    Public Property Key() As String
        Get
            Return mp_sKey
        End Get
        Set(ByVal value As String)
            mp_sKey = value
        End Set
    End Property

    Public Property IntValue() As Integer
        Get
            Return mp_lIntValue
        End Get
        Set(ByVal value As Integer)
            mp_lIntValue = value
        End Set
    End Property
End Class


