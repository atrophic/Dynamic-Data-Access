Imports System.Data
Imports System.Runtime.CompilerServices

Public Module DataRecordExtensions
    <Extension()> _
    Public Function HasColumn(ByVal dr As IDataRecord, ByVal columnName As String) As Boolean
        For i As Int32 = 0 To dr.FieldCount - 1
            If dr.GetName(i).Equals(columnName, StringComparison.InvariantCultureIgnoreCase) Then
                Return True
            End If
        Next

        Return False
    End Function
End Module
