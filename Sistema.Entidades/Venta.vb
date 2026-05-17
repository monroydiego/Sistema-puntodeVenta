Public Class Venta
    Private _IdVenta           As Integer
    Private _IdCliente         As Integer
    Private _IdUsuario         As Integer
    Private _IdTipoComprobante As Integer
    Private _NumComprobante    As String
    Private _FechaVenta        As Date
    Private _Impuesto          As Decimal
    Private _TotalVenta        As Decimal
    Private _Estado            As String

    Public Property IdVenta As Integer
        Get
            Return _IdVenta
        End Get
        Set(value As Integer)
            _IdVenta = value
        End Set
    End Property

    Public Property IdCliente As Integer
        Get
            Return _IdCliente
        End Get
        Set(value As Integer)
            _IdCliente = value
        End Set
    End Property

    Public Property IdUsuario As Integer
        Get
            Return _IdUsuario
        End Get
        Set(value As Integer)
            _IdUsuario = value
        End Set
    End Property

    Public Property IdTipoComprobante As Integer
        Get
            Return _IdTipoComprobante
        End Get
        Set(value As Integer)
            _IdTipoComprobante = value
        End Set
    End Property

    Public Property NumComprobante As String
        Get
            Return _NumComprobante
        End Get
        Set(value As String)
            _NumComprobante = value
        End Set
    End Property

    Public Property FechaVenta As Date
        Get
            Return _FechaVenta
        End Get
        Set(value As Date)
            _FechaVenta = value
        End Set
    End Property

    Public Property Impuesto As Decimal
        Get
            Return _Impuesto
        End Get
        Set(value As Decimal)
            _Impuesto = value
        End Set
    End Property

    Public Property TotalVenta As Decimal
        Get
            Return _TotalVenta
        End Get
        Set(value As Decimal)
            _TotalVenta = value
        End Set
    End Property

    Public Property Estado As String
        Get
            Return _Estado
        End Get
        Set(value As String)
            _Estado = value
        End Set
    End Property
End Class
