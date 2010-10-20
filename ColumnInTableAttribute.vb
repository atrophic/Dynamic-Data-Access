<AttributeUsage(AttributeTargets.Property)> _
Public Class ColumnInTableAttribute
    Inherits Attribute

    Private _ColumnName As String
    Public ReadOnly Property ColumnName() As String
        Get
            Return _ColumnName
        End Get
    End Property

    Public Sub New(ByVal columnName As String)
        _ColumnName = columnName
    End Sub

End Class
