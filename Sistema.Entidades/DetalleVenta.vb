Public Class DetalleVenta
    Private _IdDetalleVenta As Integer
    Private _IdVenta As Integer
    Private _IdArticulo As Integer
    Private _Cantidad As Integer
    Private _Precio As Decimal
    Private _Descuento As Decimal
    Private _Subtotal As Decimal

    Public Property IdDetalleVenta As Integer
        Get
            Return _IdDetalleVenta
        End Get
        Set(value As Integer)
            _IdDetalleVenta = value
        End Set
    End Property

    Public Property IdVenta As Integer
        Get
            Return _IdVenta
        End Get
        Set(value As Integer)
            _IdVenta = value
        End Set
    End Property

    Public Property IdArticulo As Integer
        Get
            Return _IdArticulo
        End Get
        Set(value As Integer)
            _IdArticulo = value
        End Set
    End Property

    Public Property Cantidad As Integer
        Get
            Return _Cantidad
        End Get
        Set(value As Integer)
            _Cantidad = value
        End Set
    End Property

    Public Property Precio As Decimal
        Get
            Return _Precio
        End Get
        Set(value As Decimal)
            _Precio = value
        End Set
    End Property

    Public Property Descuento As Decimal
        Get
            Return _Descuento
        End Get
        Set(value As Decimal)
            _Descuento = value
        End Set
    End Property

    Public Property Subtotal As Decimal
        Get
            Return _Subtotal
        End Get
        Set(value As Decimal)
            _Subtotal = value
        End Set
    End Property
End Class
