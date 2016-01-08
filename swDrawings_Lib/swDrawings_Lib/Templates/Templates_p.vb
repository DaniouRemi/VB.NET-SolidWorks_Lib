Public Class Templates_p

    Private _Name As String
    Private _Adress As String
    Private _Size As String
    Private _Orientation As String
    Private _ScaleMode As List(Of ScaleMode)

    Public Property Name As String
        Set(value As String)
            _Name = value
        End Set
        Get
            Return _Name
        End Get
    End Property
    Public Property Adress As String
        Set(value As String)
            _Adress = value
        End Set
        Get
            Return _Adress
        End Get
    End Property
    Public Property Orientation As String
        Set(value As String)
            _Orientation = value
        End Set
        Get
            Return _Orientation
        End Get
    End Property
    Public Property Size As String
        Set(value As String)
            _Size = value
        End Set
        Get
            Return _Size
        End Get
    End Property
    Public Property ScaleMode As List(Of ScaleMode)
        Set(value As List(Of ScaleMode))
            _ScaleMode = value
        End Set
        Get
            Return _ScaleMode
        End Get
    End Property

End Class

Public Class ScaleMode

    Private _PartSize As Double
    Private _ViewZone As Double

    Public Property PartSize As Double
        Set(value As Double)
            _PartSize = value
        End Set
        Get
            Return _PartSize
        End Get
    End Property
    Public Property ViewZone As Double
        Set(value As Double)
            _ViewZone = value
        End Set
        Get
            Return _ViewZone
        End Get
    End Property

End Class

