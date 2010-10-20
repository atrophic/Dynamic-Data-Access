Imports System.Data.SqlClient
Imports System.Reflection
Imports System.ComponentModel
Imports System.Data
Imports System.Linq

Public Class DynamicDataAccess

    Public Shared DefaultDatabase As String = ConfigurationManager.ConnectionStrings("DefaultDatabase").ToString()
    Public Delegate Function DataMapper(Of T As New)(ByVal reader As IDataRecord) As T

    ''' <summary>
    ''' Uses reflection and the ColumnInTable attribute (if set) on each property to fill a data object using a IDataRecord object.
    ''' </summary>
    ''' <typeparam name="T">The Type of object you want returned.  Must implement the default New constructor.</typeparam>
    ''' <param name="dataRecord">The data record with which to fill the data object.</param>
    ''' <returns>A new object of type T with properties set to the data in dataRecord, if the property's ColumnInTable attribute or Name matches a column name.</returns>
    ''' <remarks>
    '''   This matches the DataMapper delegate, and is used by default for methods that accept such delegate if one is not otherwise provided.
    '''   Providing data mapping methods is recommended for high use / complex objects as they will outperform this reflection-based method.
    ''' </remarks>
    Public Shared Function DefaultDataMapper(Of T As New)(ByVal dataRecord As IDataRecord) As T
        If dataRecord Is Nothing Then Return Nothing

        Dim objType As Type = GetType(T)
        Dim result As New T

        For Each prop As PropertyInfo In objType.GetProperties()
            Dim nameInDB As String

            ' see if the ColumnInTable attribute is set, use it if it is or use the property's name if it isn't
            Dim attr As ColumnInTableAttribute = prop.GetCustomAttributes(GetType(ColumnInTableAttribute), False).Cast(Of ColumnInTableAttribute).SingleOrDefault

            ' Get the name as it appears in the database either from ColumnInTable attribute, or from the property name itself
            If Not String.IsNullOrEmpty(attr.ColumnName) Then
                nameInDB = attr.ColumnName
            Else
                nameInDB = prop.Name
            End If

            ' if the name we have matches a column in the database, let's try setting it
            ' this won't work for custom types unless the custom type has a TypeConverter defined for it
            If dataRecord.HasColumn(nameInDB) Then
                Dim converter As TypeConverter = TypeDescriptor.GetConverter(prop.PropertyType)
                If dataRecord.Item(nameInDB).GetType.Equals(prop.PropertyType) _
                   OrElse converter.CanConvertFrom(dataRecord.Item(nameInDB).GetType) Then

                    prop.SetValue(result, dataRecord.Item(nameInDB), Nothing)
                End If
            End If
        Next

        Return result
    End Function

    ''' <summary>
    ''' Executes a stored procedure with passed parameters without regard for its results.
    ''' </summary>
    ''' <param name="storedProc">Name of the stored procedure to execute</param>
    ''' <param name="params">A series of key/value pairs to pass as parameters.  For example: "@username", "steve", "@dynamicValue", objDynamic</param>
    Public Shared Sub Execute(ByVal storedProc As String, ByVal ParamArray params As Object())
        Using db As New SqlConnection(DefaultDatabase)
            db.Open()
            Using cmd As New SqlCommand(storedProc, db)
                cmd.CommandType = Data.CommandType.StoredProcedure

                If params.Length Mod 2 <> 0 Then
                    Throw New ArgumentException("must contain an even number of parameters", "params")
                End If

                For i As Int32 = 0 To params.Length - 1 Step 2
                    cmd.Parameters.AddWithValue(params(i), params(i + 1))
                Next

                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Executes a stored procedure with passed parameters and returns its results as a DataSet.
    ''' </summary>
    ''' <param name="storedProc">Name of the stored procedure to execute</param>
    ''' <param name="params">A series of key/value pairs to pass as parameters.  For example: "@username", "steve", "@dynamicValue", objDynamic</param>
    ''' <returns>DataSet resulting from the stored procedure.</returns>
    Public Shared Function GetDataSet(ByVal storedProc As String, ByVal ParamArray params As Object()) As DataSet
        Using db As New SqlConnection(DefaultDatabase)
            db.Open()
            Using cmd As New SqlCommand(storedProc, db)
                cmd.CommandType = Data.CommandType.StoredProcedure

                If params.Length Mod 2 <> 0 Then
                    Throw New ArgumentException("must contain an even number of parameters", "params")
                End If

                For i As Int32 = 0 To params.Length - 1 Step 2
                    cmd.Parameters.AddWithValue(params(i), params(i + 1))
                Next

                Dim da As New SqlDataAdapter(cmd)
                Dim result As New DataSet
                da.Fill(result)

                Return result
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Executes a stored procedure with passed parameters and returns its results as a single DataTable.
    ''' </summary>
    ''' <param name="storedProc">Name of the stored procedure to execute</param>
    ''' <param name="params">A series of key/value pairs to pass as parameters.  For example: "@username", "steve", "@dynamicValue", objDynamic</param>
    ''' <returns>DataTable resulting from the stored procedure.</returns>
    ''' <exception cref="NotSupportedException">Thrown if the DataSet contains more than one Table.</exception>
    Public Shared Function GetDataTable(ByVal storedProc As String, ByVal ParamArray params As Object()) As DataTable
        Dim ds As DataSet = GetDataSet(storedProc, params)
        If ds.Tables.Count.Equals(0) Then Return Nothing
        If ds.Tables.Count > 1 Then Throw New NotSupportedException("DataSet contains multiple tables, don't know which one to return.")
        Return ds.Tables(0)
    End Function

    ''' <summary>
    ''' Executes a stored procedure with passed parameters and returns its results as a List of the requested object.
    ''' </summary>
    ''' <param name="storedProc">Name of the stored procedure to execute</param>
    ''' <param name="params">A series of key/value pairs to pass as parameters.  For example: "@username", "steve", "@dynamicValue", objDynamic</param>
    ''' <returns>List of type T using <see cref="DefaultDataMapper">DefaultDataMapper</see> to map the data.</returns>
    Public Shared Function GetMultple(Of T As New)(ByVal storedProc As String, ByVal ParamArray params As Object()) As List(Of T)
        Return GetMultple(Of T)(storedProc, AddressOf DefaultDataMapper(Of T), params)
    End Function

    ''' <summary>
    ''' Executes a stored procedure with passed parameters and returns its results as a List of the requested object.
    ''' </summary>
    ''' <param name="storedProc">Name of the stored procedure to execute</param>
    ''' <param name="dataMapper">The delegate function used to map the stored procedure results to the data object</param>
    ''' <param name="params">A series of key/value pairs to pass as parameters.  For example: "@username", "steve", "@dynamicValue", objDynamic</param>
    ''' <returns>List of type T using <paramref name="dataMapper">dataMapper</paramref> to map the data.</returns>
    Public Shared Function GetMultple(Of T As New)(ByVal storedProc As String, ByVal dataMapper As DataMapper(Of T), ByVal ParamArray params As Object()) As List(Of T)
        Dim results As New List(Of T)

        Using db As New SqlConnection(DefaultDatabase)
            db.Open()
            Using cmd As New SqlCommand(storedProc, db)
                cmd.CommandType = Data.CommandType.StoredProcedure

                If params.Length Mod 2 <> 0 Then
                    Throw New ArgumentException("must contain an even number of parameters", "params")
                End If

                For i As Int32 = 0 To params.Length - 1 Step 2
                    cmd.Parameters.AddWithValue(params(i), params(i + 1))
                Next

                Dim reader As SqlDataReader = cmd.ExecuteReader

                While reader.Read
                    results.Add(dataMapper.Invoke(reader))
                End While
            End Using
        End Using

        If results.Count.Equals(0) Then Return Nothing

        Return results
    End Function

    ''' <summary>
    ''' Executes a stored procedure with passed parameters and returns its scalar result as an object of type T.
    ''' </summary>
    ''' <param name="storedProc">Name of the stored procedure to execute</param>
    ''' <param name="params">A series of key/value pairs to pass as parameters.  For example: "@username", "steve", "@dynamicValue", objDynamic</param>
    ''' <returns>Single object of type T DirectCasted from ExecuteScalar.</returns>
    ''' <exception cref="InvalidCastException">The type of the object returned by the stored procedure must be DirectCastable to type T an exception will be thrown.</exception>
    Public Shared Function GetScalar(Of T As New)(ByVal storedProc As String, ByVal ParamArray params As Object()) As T
        Using db As New SqlConnection(DefaultDatabase)
            db.Open()
            Using cmd As New SqlCommand(storedProc, db)
                cmd.CommandType = Data.CommandType.StoredProcedure

                If params.Length Mod 2 <> 0 Then
                    Throw New ArgumentException("must contain an even number of parameters", "params")
                End If

                For i As Int32 = 0 To params.Length - 1 Step 2
                    cmd.Parameters.AddWithValue(params(i), params(i + 1))
                Next

                Return DirectCast(cmd.ExecuteScalar, T)
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Executes a stored procedure with passed parameters and returns its results as a single object of the requested type.
    ''' </summary>
    ''' <param name="storedProc">Name of the stored procedure to execute</param>
    ''' <param name="params">A series of key/value pairs to pass as parameters.  For example: "@username", "steve", "@dynamicValue", objDynamic</param>
    ''' <returns>Object of type T using <see cref="DefaultDataMapper">DefaultDataMapper</see> to map the data.</returns>
    Public Shared Function GetSingle(Of T As New)(ByVal storedProc As String, ByVal ParamArray params As Object()) As T
        Return GetSingle(Of T)(storedProc, AddressOf DefaultDataMapper(Of T), params)
    End Function

    ''' <summary>
    ''' Executes a stored procedure with passed parameters and returns its results as a single object of the requested type.
    ''' </summary>
    ''' <param name="storedProc">Name of the stored procedure to execute</param>
    ''' <param name="dataMapper">The delegate function used to map the stored procedure results to the data object</param>
    ''' <param name="params">A series of key/value pairs to pass as parameters.  For example: "@username", "steve", "@dynamicValue", objDynamic</param>
    ''' <returns>Object of type T using <paramref name="dataMapper">dataMapper</paramref> to map the data.</returns>
    Public Shared Function GetSingle(Of T As New)(ByVal storedProc As String, ByVal dataMapper As DataMapper(Of T), ByVal ParamArray params As Object()) As T
        Dim results As List(Of T) = GetMultple(storedProc, dataMapper, params)

        If results Is Nothing OrElse results.Count = 0 Then Return Nothing

        If results.Count = 1 Then Return results.Item(0)

        Throw New NotSupportedException("Multiple records were returned in a call to GetSingle.  Use GetMultiple instead.")
    End Function
End Class
