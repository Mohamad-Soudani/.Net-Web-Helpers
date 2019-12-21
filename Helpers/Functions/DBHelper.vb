Imports System.Data.SqlClient
Imports System.IO
Imports System.Reflection

Namespace Helpers
    Public Module DBHelper

        Public DefaultConnectionString As String = "Data Source=.;Initial Catalog=catalog_name; User Id=sa;Password=123456;"
        Public Connection As SqlClient.SqlConnection
        Public tran As SqlClient.SqlTransaction

        Sub New()
            Connection = EstablishConnectionWithDataBase()
        End Sub

        Sub WindowsAuthLogIn(Optional connection As SqlClient.SqlConnection = Nothing)
            connection = If(connection IsNot Nothing, connection, DBHelper.Connection)
            If File.Exists("ImpersonateFile.txt") Then
                If connection.State <> ConnectionState.Open Then
                    connection.Open()
                    Dim command As New SqlClient.SqlCommand
                    command.CommandType = CommandType.Text
                    command.Connection = connection
                    command.CommandText = "Execute AS Login=''"
                    command.ExecuteNonQuery()
                Else
                    If connection.State = ConnectionState.Closed Then
                        connection.Open()
                    End If
                End If
            End If
        End Sub

        Sub ExecuteNonQuery(ByVal query As String, Optional ByVal connection As SqlClient.SqlConnection = Nothing, Optional ByVal param As Dictionary(Of String, Object) = Nothing, Optional withTrans As Boolean = True, Optional timeOut As Integer = 30)
            'Try
            If connection Is Nothing Then
                connection = DBHelper.Connection
            End If
            Dim command As New SqlClient.SqlCommand
            If connection.State = ConnectionState.Closed Then
                connection.Open()
            End If
            If (withTrans) Then
                tran = connection.BeginTransaction(IsolationLevel.Serializable)
            Else
            End If

            command.CommandTimeout = timeOut

            command = DoExecuteNonQuery(query, param, connection, command)
            If (withTrans) Then
                tran.Commit()
            End If
            'Catch ex As Exception
            '    If connection.State = ConnectionState.Open Then
            '        'tran.Rollback()
            '        connection.Close()
            '    End If
            'End Try
        End Sub

        Sub ExecuteMultipleNonQuery(ByVal queryList As String(), Optional ByVal param As Dictionary(Of String, Object) = Nothing, Optional ByVal connection As SqlClient.SqlConnection = Nothing, Optional commit As Boolean = True)
            Try
                If connection Is Nothing Then
                    connection = DBHelper.Connection
                End If
                Dim command As New SqlClient.SqlCommand
                If (connection.State.Equals(ConnectionState.Closed)) Then
                    connection.Open()
                End If
                tran = connection.BeginTransaction(IsolationLevel.Serializable)
                For Each query As String In queryList
                    DoExecuteNonQuery(query, param, connection, command)
                Next
                If commit Then
                    tran.Commit()
                End If
            Catch ex As Exception
                'MessageBox.Show(ex.Message)
                tran.Rollback()
            End Try
        End Sub


        Function DoExecuteNonQuery(query As String, param As Dictionary(Of String, Object), connection As SqlConnection, command As SqlCommand, Optional withTrans As Boolean = True) As SqlCommand
            command.Connection = connection
            If withTrans Then
                command.Transaction = tran
            End If
            command.CommandText = query
            If connection.State.Equals(ConnectionState.Closed) Then
                connection.Open()
            End If
            If param IsNot Nothing Then
                For Each keyValuePair As KeyValuePair(Of String, Object) In param
                    command.Parameters.AddWithValue(keyValuePair.Key, keyValuePair.Value)
                Next
            End If
            command.ExecuteNonQuery()
            Return command
        End Function
        Function Get_Data(ByVal query As String, Optional ByVal connection As SqlClient.SqlConnection = Nothing, Optional timeOut As Integer = 30) As DataTable
            Return DoGetData(query, connection, timeOut)
        End Function

        Function ExecuteScalar(ByVal query As String, Optional ByVal connection As SqlClient.SqlConnection = Nothing) As Integer
            If connection Is Nothing Then
                connection = DBHelper.Connection
            End If

            If connection.State.Equals(ConnectionState.Closed) Then
                connection.Open()
            End If
            Dim command As New SqlClient.SqlCommand
            command.Connection = connection
            command.CommandText = query
            Dim result = command.ExecuteScalar()
            Return If(IsDBNull(result), 0, Convert.ToInt32(result))

        End Function

        Function Get_Stored_Proceedure_Data(ByVal query As String, Optional ByVal params As Dictionary(Of String, Object) = Nothing, Optional ByVal connection As SqlClient.SqlConnection = Nothing, Optional timeOut As Boolean = True) As DataTable
            If connection Is Nothing Then
                connection = DBHelper.Connection
            End If
            connection.Close()
            connection.Open()
            Dim command As SqlCommand = Prepare_StoredProceedure_Command(query, params, connection, timeOut)
            Dim adapter As New SqlDataAdapter(command)
            Dim DS As DataTable = New DataTable()
            adapter.Fill(DS)
            Return DS
            'MessageBox.Show(ex.Message)
            Return Nothing
        End Function

        Sub Execute_Stored_Proceedure(ByVal name As String, Optional ByVal params As Dictionary(Of String, Object) = Nothing, Optional ByVal connection As SqlClient.SqlConnection = Nothing, Optional timeOut As Boolean = True)
            Try
                Dim command As SqlCommand = Prepare_StoredProceedure_Command(name, params, connection, timeOut)
                command.ExecuteNonQuery()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(ex)
            End Try
        End Sub
        Private Function Prepare_StoredProceedure_Command(query As String, params As Dictionary(Of String, Object), connection As SqlConnection, timeOut As Boolean) As SqlCommand
            If connection Is Nothing Then
                connection = DBHelper.Connection
            End If
            Dim command As New SqlCommand
            command.Connection = connection
            command.CommandText = query
            command.CommandType = CommandType.StoredProcedure
            If timeOut Then
                command.CommandTimeout = 30
            Else
                command.CommandTimeout = 0
            End If
            If params IsNot Nothing Then

                For Each keyValuePair As KeyValuePair(Of String, Object) In params
                    command.Parameters.AddWithValue(keyValuePair.Key, keyValuePair.Value)
                Next
            End If
            Return command
        End Function

        Function List(Of T As {Class, New})(Optional ByVal query As String = Nothing, Optional ByVal where As String = Nothing, Optional ByVal connection As SqlConnection = Nothing, Optional LoadProxy As Boolean = False, Optional timeOut As Integer = 30, Optional tableName As String = "") As List(Of T)
            where = If(Not String.IsNullOrEmpty(where), "where " + where, "")
            Dim table As String = If(String.IsNullOrEmpty(tableName), GetType(T).Name, tableName)
            table = FixTableNameWithAttributes(Of T)(table)
            query = If(Not String.IsNullOrEmpty(query), query, "Select * from " + table + " " + where)
            Dim data As DataTable = DoGetData(query, connection, timeOut)
            Dim dataList As List(Of T) = data.ToList(Of T)
            If LoadProxy AndAlso GetType(IProxy).IsAssignableFrom(GetType(T)) Then
                For Each item As T In dataList
                    Dim proxyItem As IProxy = item
                    'proxyItem.LoadProxyObject(proxyItem)
                Next
            End If
            Return dataList
        End Function

        Function Delete(Of T As {Class, New})(Optional where As String = "", Optional ByVal connection As SqlConnection = Nothing, Optional commitTrans As Boolean = True) As Boolean
            Try
                where = If(Not String.IsNullOrEmpty(where), "where " + where, "")
                Dim query As String = "Delete from " + GetType(T).Name + " " + where
                ExecuteNonQuery(query, connection, withTrans:=commitTrans)
                Return True
            Catch ex As Exception
                'MessageBox.Show(ex.Message)
            End Try
        End Function


        Sub Delete(Of T As {Class, New})(ByVal obj As T, ByVal Optional TableName As String = Nothing, Optional ByVal connection As SqlConnection = Nothing, Optional commitTrans As Boolean = True)
            Dim properties = GetType(T).GetProperties()
            Dim primaryKeys As List(Of PrimaryKey) = CollectPrimaryKeys(obj, properties)
            TableName = If(TableName, GetType(T).Name)
            TableName = FixTableNameWithAttributes(Of T)(TableName)
            Dim whereCondition = String.Join(" And ", primaryKeys.Select(Function(k) k.Name + " = '" + k.Value.Trim + "'"))
            Dim query As String = "Delete from " + TableName + " where " + whereCondition
            ExecuteNonQuery(query, connection, withTrans:=commitTrans)
        End Sub

        Sub DeleteAll(Of T As {Class, New})(ByVal list As IEnumerable(Of T), ByVal Optional TableName As String = Nothing, Optional ByVal connection As SqlConnection = Nothing, Optional commitTrans As Boolean = True)
            For Each item As T In list
                Delete(item, connection:=connection, commitTrans:=commitTrans)
            Next
        End Sub

        Public Function CollectPrimaryKeys(Of T As {Class, New})(ByVal obj As T, ByVal properties As PropertyInfo()) As List(Of PrimaryKey)
            Dim primaryKeys As List(Of PrimaryKey) = New List(Of PrimaryKey)()
            For Each prop As PropertyInfo In properties
                Dim isKey = CheckForPrimaryKey(prop)
                If isKey Then
                    Dim valueString = ""
                    Dim value = prop.GetValue(obj, Nothing)
                    Dim dummy As Integer = 0
                    valueString = GenerateValueString(dummy, Nothing, value)
                    primaryKeys.Add(New PrimaryKey(prop.Name, valueString))
                End If
            Next
            Return primaryKeys
        End Function

        Private Function DoGetData(query As String, connection As SqlConnection, Optional timeOut As Integer = 30) As DataTable
            If connection Is Nothing Then
                connection = DBHelper.Connection
            End If
            Dim command As New SqlCommand
            With command
                .CommandText = query
                .Transaction = tran
                .Connection = connection
            End With

            command.CommandTimeout = timeOut

            Dim adapter As SqlDataAdapter = New SqlDataAdapter(command)
            Dim DtFill As New DataTable
            DtFill.Clear()
            adapter.Fill(DtFill)
            Return DtFill
        End Function

        Sub InsertOrUpdate(Of T As {Class, New})(ByVal obj As T, ByVal Optional TableName As String = Nothing, Optional ByVal connection As SqlConnection = Nothing, Optional commitTrans As Boolean = True)
            Try
                Dim propertiesString As Object = Nothing
                Dim mappedValues As Dictionary(Of String, Object) = Nothing
                Dim param As Dictionary(Of String, Object) = Nothing
                Dim table As String = Nothing
                Dim query As String = ""
                Dim whereCondition As String = ""
                Dim itemExists As Boolean = False
                PrepareDatabaseTransactionData(obj, TableName, propertiesString, mappedValues, table, whereCondition, itemExists, param, connection:=connection)
                If itemExists Then
                    query = GenerateUpdateQuery(whereCondition, mappedValues, table)
                Else
                    query = GenerateInsertQuery(mappedValues, table)
                End If
                ExecuteNonQuery(query, param:=param, connection:=connection, withTrans:=commitTrans)
            Catch ex As Exception
                'MessageDialoges.ErrorMsg(ex.Message)
            End Try
        End Sub

        Sub InsertOrUpdateAll(Of T As {Class, New})(ByVal list As IEnumerable(Of T), ByVal Optional TableName As String = Nothing, Optional ByVal connection As SqlConnection = Nothing, Optional commitTrans As Boolean = True)
            For Each item As T In list
                InsertOrUpdate(item, commitTrans:=commitTrans)
            Next
        End Sub

        Sub Update(Of T As {Class, New})(ByVal obj As T, ByVal Optional TableName As String = Nothing, Optional ByVal connection As SqlConnection = Nothing, Optional commitTrans As Boolean = True)
            Try
                Dim propertiesString As Object = Nothing
                Dim mappedValues As Dictionary(Of String, Object) = Nothing
                Dim param As Dictionary(Of String, Object) = Nothing
                Dim table As String = Nothing
                Dim query As String = ""
                Dim whereCondition As String = ""
                Dim itemExists As Boolean = False
                PrepareDatabaseTransactionData(obj, TableName, propertiesString, mappedValues, table, whereCondition, itemExists, param, connection:=connection)
                If itemExists Then
                    Dim ignoredProperties As List(Of PropertyInfo) = GetType(T).GetProperties().ToList().Where(Function(x) IsUpdateIgnored(x)).ToList()
                    For Each prop As PropertyInfo In ignoredProperties
                        If mappedValues.ContainsKey(prop.Name) Then
                            mappedValues.Remove(prop.Name)
                        End If
                    Next
                    query = GenerateUpdateQuery(whereCondition, mappedValues, table)
                End If
                ExecuteNonQuery(query, param:=param, connection:=connection, withTrans:=commitTrans)
            Catch ex As Exception
                'MessageDialoges.ErrorMsg(ex.Message)
                System.Diagnostics.Debug.WriteLine(ex)
            End Try
        End Sub

        Function GenerateUpdateQuery(Of T As {Class, New})(ByVal obj As T, ByVal Optional TableName As String = Nothing, Optional ByVal connection As SqlConnection = Nothing) As String
            Try
                Dim propertiesString As Object = Nothing
                Dim mappedValues As Dictionary(Of String, Object) = Nothing
                Dim param As Dictionary(Of String, Object) = Nothing
                Dim table As String = Nothing
                Dim query As String = ""
                Dim whereCondition As String = ""
                Dim itemExists As Boolean = False
                PrepareDatabaseTransactionData(obj, TableName, propertiesString, mappedValues, table, whereCondition, itemExists, param, connection:=connection)
                If itemExists Then
                    query = GenerateUpdateQuery(whereCondition, mappedValues, table)
                End If
                Return GenerateUpdateQuery(whereCondition, mappedValues, table)
            Catch ex As Exception
                'MessageDialoges.ErrorMsg(ex.Message)
            End Try
        End Function

        Sub Insert(Of T As {Class, New})(ByVal obj As T, ByVal Optional TableName As String = Nothing, Optional ByVal connection As SqlConnection = Nothing, Optional commitTrans As Boolean = True, Optional timeOut As Integer = 30)
            Dim propertiesString As Object = Nothing
            Dim mappedValues As Dictionary(Of String, Object) = Nothing
            Dim param As Dictionary(Of String, Object) = Nothing
            Dim table As String = Nothing
            Dim query As String = ""
            Dim whereCondition As String = ""
            Dim itemExists As Boolean = False
            PrepareDatabaseTransactionData(obj, TableName, propertiesString, mappedValues, table, whereCondition, itemExists, param, connection:=connection, checkExisting:=False)
            If mappedValues.Count > 0 Then
                query = GenerateInsertQuery(mappedValues, table)
                ExecuteNonQuery(query, param:=param, connection:=connection, withTrans:=commitTrans, timeOut:=timeOut)
            End If
        End Sub

        Sub InsertAll(Of T As {Class, New})(ByVal list As IEnumerable(Of T), ByVal Optional TableName As String = Nothing, Optional ByVal connection As SqlConnection = Nothing, Optional commitTrans As Boolean = True, Optional timeOut As Integer = 30)
            For Each item As T In list
                Insert(item, connection:=connection, commitTrans:=commitTrans, timeOut:=timeOut)
            Next
        End Sub

        <Obsolete>
        Sub InsertAll_SameTransaction(Of T As {Class, New})(ByVal list As IEnumerable(Of T), ByVal Optional TableName As String = Nothing, Optional ByVal connection As SqlConnection = Nothing, Optional commitTrans As Boolean = True)
            InsertAll(list, TableName, connection, False)
            'If commitTrans Then
            '    tran.Commit()
            '    'MessageDialoges.ErrorMsg(ex.Message)
            'End If
        End Sub

        Sub UpdateAll(Of T As {Class, New})(ByVal list As IEnumerable(Of T), ByVal Optional TableName As String = Nothing, Optional ByVal connection As SqlConnection = Nothing, Optional commitTrans As Boolean = True)
            For Each item As T In list
                Update(item, commitTrans:=commitTrans)
            Next
        End Sub


        Function GetItem(Of T As {Class, New})(id As String, Optional ByVal Connection As SqlClient.SqlConnection = Nothing) As T
            Return GetItem(Of T)({id}, connection:=Connection)
        End Function

        Function GetItem(Of T As {Class, New})(ids As String(), Optional ByVal connection As SqlConnection = Nothing) As T
            If connection Is Nothing Then
                connection = DBHelper.Connection
            End If
            Dim properties As PropertyInfo() = GetType(T).GetProperties()
            Dim primaryKeys As New List(Of PrimaryKey)
            Dim key As String = ""
            Dim foreignAtts As New Dictionary(Of PropertyInfo, ForeignObjectAttribute)
            Dim proxyAtts As New List(Of PropertyInfo)
            Try
                For Each prop As PropertyInfo In properties
                    If IsIgnored(prop) Then
                        CheckForForeignObjects(prop, foreignAtts)
                        CheckForProxyObjects(prop, proxyAtts)
                    End If
                    If CheckForPrimaryKey(prop) Then
                        Dim keyIndex As Integer = GetPrimaryKeyIndex(prop)
                        If keyIndex > -1 Then
                            Dim idValue As String = ids(keyIndex)
                            primaryKeys.Add(New PrimaryKey(prop.Name, idValue))
                        End If
                    End If
                Next
                Dim where As String = GeneratePrimaryKeysString(primaryKeys)
                Dim item As T = List(Of T)(where:=where, connection:=connection).SingleOrDefault()
                If item IsNot Nothing Then
                    SetProxyobjects(item, proxyAtts, connection:=connection)
                    GetForeignLists(item, foreignAtts, connection:=connection)
                End If
                Return item
            Catch ex As Exception
                'MessageBox.Show(ex.Message)
                Return Nothing
            End Try
        End Function

        Private Function GeneratePrimaryKeysString(primaryKeys As List(Of PrimaryKey)) As String
            Return String.Join(" AND ", primaryKeys.Select(Function(k) k.Name & " = '" + k.Value.Trim + "'"))
        End Function

        Private Sub SetProxyobjects(Of T As {Class, New})(item As T, proxyAtts As List(Of PropertyInfo), connection As SqlConnection)
            If proxyAtts.Count > 0 Then
                For Each prop As PropertyInfo In proxyAtts
                    Dim p As T = ConvertToType(item, item.GetType())
                    prop.SetValue(item, p, Nothing)
                Next
            End If
        End Sub

        Sub CheckForProxyObjects(prop As PropertyInfo, proxyAtts As List(Of PropertyInfo))
            Dim att As ProxyObjectAttribute
            If prop.IsDefined(GetType(ProxyObjectAttribute), False) Then
                att = CType(Attribute.GetCustomAttribute(prop, GetType(ProxyObjectAttribute)), ProxyObjectAttribute)
                proxyAtts.Add(prop)
            End If
        End Sub

        Sub GetForeignLists(Of T As {Class, New})(item As T, foreignAtts As Dictionary(Of PropertyInfo, ForeignObjectAttribute), Optional ByVal connection As SqlConnection = Nothing)
            Try
                If foreignAtts.Keys.Count > 0 Then
                    For Each keyvalue As KeyValuePair(Of PropertyInfo, ForeignObjectAttribute) In foreignAtts
                        Dim prop As PropertyInfo = keyvalue.Key
                        Dim foreignatt As ForeignObjectAttribute = keyvalue.Value
                        Dim key As String = item.GetType().GetProperty(foreignatt.PrimaryKey).GetValue(item, Nothing).ToString().ToDBString()
                        Dim query As String = " Select * from " + foreignatt.TableName + " where " + foreignatt.ForeignKey + " = " + key
                        Dim data As DataTable = Get_Data(query, connection:=connection)
                        Dim foreignObj = Activator.CreateInstance(prop.PropertyType)
                        Dim lst = data.ToList(foreignatt.Type).ToArray().ToList()
                        Dim o As Object = Nothing
                        If foreignatt.SingleObject Then
                            o = lst.SingleOrDefault()
                        ElseIf lst.Count > 0 Then
                            o = lst
                        End If
                        Dim f = ConvertToType(o, prop.PropertyType)
                        prop.SetValue(item, f, Nothing)
                    Next
                End If
            Catch ex As Exception
                'MessageDialoges.WarningMsg(ex.Message)
            End Try
        End Sub
        Private Sub CheckForForeignObjects(prop As PropertyInfo, ByRef foreignAtts As Dictionary(Of PropertyInfo, ForeignObjectAttribute))
            Dim att As ForeignObjectAttribute
            If prop.IsDefined(GetType(ForeignObjectAttribute), False) Then
                att = CType(Attribute.GetCustomAttribute(prop, GetType(ForeignObjectAttribute)), ForeignObjectAttribute)
                foreignAtts.Add(prop, att)
            End If
        End Sub

        Private Sub PrepareDatabaseTransactionData(Of T As {Class, New})(obj As T, TableName As String, ByRef propertiesString As Object, ByRef mappedValues As Dictionary(Of String, Object), ByRef table As String, ByRef whereCondition As String, ByRef itemExists As Boolean, param As Dictionary(Of String, Object), Optional ByVal connection As SqlConnection = Nothing, Optional checkExisting As Boolean = True)
            Dim properties As PropertyInfo() = GetType(T).GetProperties()
            Dim ProxyProp As PropertyInfo = HasProxy(properties)
            Dim primaryKeys As List(Of PrimaryKey) = FilterProperties(properties)
            propertiesString = properties.ToDataBaseCommaSeparatedString("Name").Replace("'", "")
            Dim attachmentIndex = -1
            mappedValues = New Dictionary(Of String, Object)()
            attachmentIndex = MapValuesToProperties(obj, properties, attachmentIndex, param, mappedValues, primaryKeys)
            table = If(TableName, GetType(T).Name)
            table = FixTableNameWithAttributes(Of T)(table)
            GetProxyPrimaryKeys(obj, properties, ProxyProp, primaryKeys)
            If (checkExisting) Then
                whereCondition = String.Join(" AND ", primaryKeys.Select(Function(k) k.Name & " = '" + k.Value.Trim + "'"))
                itemExists = CheckIfExistsInLocalDB(obj, whereCondition, table, connection)
            End If
        End Sub

        Private Function GetProxyPrimaryKeys(Of T As {Class, New})(obj As T, properties As PropertyInfo(), ProxyProp As PropertyInfo, primaryKeys As List(Of PrimaryKey)) As Object
            If ProxyProp IsNot Nothing Then
                If ProxyProp.PropertyType = GetType(T) Then
                    Dim proxyObject As T = ProxyProp.GetValue(obj, Nothing)
                    If proxyObject IsNot Nothing Then
                        primaryKeys.Clear()
                        For Each prop As PropertyInfo In properties
                            Dim isKey = CheckForPrimaryKey(prop)
                            If isKey Then
                                Dim valueString = ""
                                Dim value = prop.GetValue(proxyObject, Nothing)
                                valueString = GenerateValueString(0, Nothing, value)
                                primaryKeys.Add(New PrimaryKey(prop.Name, valueString))
                            End If
                        Next
                    End If
                End If
            End If
        End Function

        Private Function HasProxy(properties As PropertyInfo()) As PropertyInfo
            Return properties.Where(Function(x) x.IsDefined(GetType(ProxyObjectAttribute), False)).FirstOrDefault()
        End Function

        Function FixTableNameWithAttributes(Of T As {Class, New})(ByRef tableName As String) As String
            Dim objType As Type = GetType(T)
            Return FixTableNameWithAttributes(tableName, objType)
        End Function

        Public Function FixTableNameWithAttributes(tableName As String, objType As Type) As String
            If objType.IsDefined(GetType(ReplaceUnderscoreWithHyphenAttribute), False) Then
                Dim att As ReplaceUnderscoreWithHyphenAttribute = System.Attribute.GetCustomAttributes(objType)(0)
                ReplaceHyphensAtPositions(tableName, att.HyphenLocations)
            ElseIf objType.IsDefined(GetType(TableNameAttribute), False) Then
                Dim attr As TableNameAttribute = Attribute.GetCustomAttributes(objType)(0)
                tableName = attr.TableName
            End If
            tableName = "[" + tableName + "]"
            Return tableName
        End Function

        Private Sub ReplaceHyphensAtPositions(ByRef tableName As String, hyphenLocations() As Integer)
            If hyphenLocations.Count() > 0 Then
                Dim hyphenOrder As Integer = 0
                Dim charList As New List(Of Char)
                For Each c As Char In tableName
                    If c = "_" Then
                        hyphenOrder = hyphenOrder + 1
                        If hyphenLocations.Contains(hyphenOrder) Then
                            c = "-"
                        End If
                    End If
                    charList.Add(c)
                Next
                tableName = String.Join("", charList)
            Else
                tableName = tableName.Replace("_", "-")
            End If
        End Sub

        Private Function GenerateUpdateQuery(ByVal whereCondition As String, ByVal mappedValues As Dictionary(Of String, Object), ByVal tableName As String) As String

            Dim setValues = String.Join(",", mappedValues.ToList().Select(Function(kv) kv.Key + " = " + kv.Value.ToString().SingleQuotes())).RemoveAttachmentPrints()
            Dim query As String = " Update " + tableName + " set " + setValues + " where " + whereCondition
            Return query
        End Function

        Private Function MapValuesToProperties(Of T)(ByVal obj As T, ByVal properties As PropertyInfo(), ByVal attachmentIndex As Integer, ByVal param As Dictionary(Of String, Object), ByVal mappedValues As Dictionary(Of String, Object), primaryKeys As List(Of PrimaryKey)) As Integer
            For Each prop As PropertyInfo In properties
                Dim valueString = ""
                Dim value = prop.GetValue(obj, Nothing)
                valueString = GenerateValueString(attachmentIndex, param, value)
                If Not String.IsNullOrEmpty(valueString) Then
                    mappedValues.Add(prop.Name, valueString.Trim)
                    Dim key As PrimaryKey = primaryKeys.Where(Function(x) x.Name = prop.Name).SingleOrDefault()
                    If key IsNot Nothing Then
                        key.Value = mappedValues(prop.Name)
                    End If
                End If
            Next
            Return attachmentIndex
        End Function

        Private Function CheckIfExistsInLocalDB(Of T As {Class, New})(obj As T, ByVal whereCondition As String, ByVal tableName As String, Optional ByVal connection As SqlConnection = Nothing) As Boolean
            Dim query = " Select * from " + tableName + " where " + whereCondition
            Dim result As List(Of T) = List(Of T)(query, connection:=connection)
            Return If(result IsNot Nothing, result.Count > 0, False)
        End Function

        Private Function GenerateInsertQuery(mappedValues As Dictionary(Of String, Object), ByVal tableName As String) As String
            Dim valueString As String = mappedValues.Values.ToDataBaseCommaSeparatedString(addPrefix:=True)
            Dim propertiesString = mappedValues.Keys.ToCommaSeparatedString()
            Return " Insert Into " + tableName + " (" + propertiesString + ") Values" + valueString + " "
        End Function

        Private Function GenerateValueString(ByRef attachmentIndex As Integer, ByVal param As Dictionary(Of String, Object), ByVal value As Object) As String
            Dim valueString As String = Nothing
            If value IsNot Nothing Then
                If TypeOf value Is DateTime Then
                    Dim dateValue As Date = CType(value, Date)
                    If (dateValue.Year > 1) Then
                        valueString = dateValue.ToMyDateTimeString()
                    Else
                        valueString = DBNull.Value.ToString()
                    End If
                ElseIf TypeOf value Is Byte() Then
                    attachmentIndex += 1
                    valueString = "@@attachment" + attachmentIndex + "@"
                    param.Add("@attachment" + attachmentIndex, value)
                Else valueString = value.ToString
                End If
            Else
                valueString = DBNull.Value.ToString()
            End If
            Return valueString.EscapeSqlString()
        End Function

        Private Function CheckForPrimaryKey(ByVal prop As PropertyInfo) As Boolean
            Dim attributes As Object() = prop.GetCustomAttributes(False)
            Dim isKey As Boolean = prop.IsDefined(GetType(PrimaryKeyAttribute), False)
            If isKey Then Return True Else Return False
        End Function

        Private Function GetPrimaryKeyIndex(ByVal prop As PropertyInfo) As Integer
            Dim att As PrimaryKeyAttribute
            If prop.IsDefined(GetType(PrimaryKeyAttribute), False) Then
                att = CType(Attribute.GetCustomAttribute(prop, GetType(PrimaryKeyAttribute)), PrimaryKeyAttribute)
                Return att.KeyIndex
            End If
        End Function

        Private Function CheckPrimaryKeyAutoIncrement(ByVal prop As PropertyInfo) As Integer
            Dim att As PrimaryKeyAttribute
            If prop.IsDefined(GetType(PrimaryKeyAttribute), False) Then
                att = CType(Attribute.GetCustomAttribute(prop, GetType(PrimaryKeyAttribute)), PrimaryKeyAttribute)
                Return att.AutoIncrement
            End If
        End Function

        Sub HandleMultipleTransactions(ParamArray transAray As SqlTransaction())
            Try
                For Each tran As SqlTransaction In transAray
                    tran.Commit()
                Next
            Catch ex As Exception
                For Each tran As SqlTransaction In transAray
                    tran.Rollback()
                Next
            End Try
        End Sub

        Function EstablishConnectionWithDataBase(Optional ConnectionString As String = "") As SqlConnection
            Try
                ConnectionString = If(StringIsNot_Null_Empty_WhiteSpace(ConnectionString), ConnectionString, DefaultConnectionString)
                Dim connection As New SqlConnection(ConnectionString)
                connection.Open()
                Return connection
            Catch ex As Exception
                'MessageBox.Show(ex.Message)
                Return Nothing
            End Try
        End Function

        'Function CheckColumnsExist(Of T)(colname As String, Optional connection As SqlConnection = Nothing) As Boolean
        '    Dim query As String = "declare @exists as bit = 0
        '                            IF EXISTS(Select 1 FROM sys.columns
        '                                       WHERE Name = N'" + colname + "'
        '                                       And Object_ID = Object_ID(N'" + GetType(T).Name + "'))
        '                                set @exists = 1
        '                            ELSE
        '                                set @exists = 0
        '                                       select @exists as 'Exists'"
        '    Return List(Of ItemExist)(query, connection:=connection).SingleOrDefault().Exist
        'End Function

        Private Function IsIgnored(ByVal prop As PropertyInfo) As Boolean
            Dim attributes As List(Of Object) = prop.GetCustomAttributes(False).ToList()
            Dim ignore As Boolean = False
            If prop.IsDefined(GetType(DBIgnore), False) Then
                ignore = True
            End If
            Return ignore
        End Function

        Private Function IsUpdateIgnored(ByVal prop As PropertyInfo) As Boolean
            Dim attributes As List(Of Object) = prop.GetCustomAttributes(False).ToList()
            Dim ignore As Boolean = False
            If prop.IsDefined(GetType(UpdateIgnored), False) Then
                ignore = True
            End If
            Return ignore
        End Function

        Private Function FilterProperties(ByRef properties As PropertyInfo()) As List(Of PrimaryKey)
            Dim propertiesDelete As New List(Of PropertyInfo)
            Dim propertiesTemp As List(Of PropertyInfo) = properties.ToList()
            Dim primaryKeys As New List(Of PrimaryKey)
            For Each prop As PropertyInfo In properties
                If CheckForPrimaryKey(prop) Then primaryKeys.Add(New PrimaryKey(prop.Name, Nothing))
                If IsIgnored(prop) OrElse CheckPrimaryKeyAutoIncrement(prop) Then propertiesDelete.Add(prop)
            Next
            For Each prop As PropertyInfo In propertiesDelete
                propertiesTemp.Remove(prop)
            Next
            properties = propertiesTemp.ToArray()
            Return primaryKeys
        End Function

        Function GetTableNames(Optional connection As SqlConnection = Nothing) As List(Of String)
            Dim query As String = "Select TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' and TABLE_NAME not like 'conflict_dbo_%' order by TABLE_NAME  "
            Dim tableList As List(Of String) = Get_Data(query, connection:=connection).To1DList("TABLE_NAME")
            Return tableList
        End Function

        Function GetCatalogueNames(Optional connection As SqlConnection = Nothing)
            Dim query As String = "Select name From sys.databases order by name"
            Dim catalogueList As List(Of String) = Get_Data(query, connection:=connection).To1DList("name")
            Return catalogueList
        End Function

        Function GetColumnNames(tableName As String, Optional connection As SqlConnection = Nothing) As List(Of String)
            Dim query As String = "SELECT Column_name From INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = N'" + tableName + "' order by Column_name "
            Dim columnList As List(Of String) = Get_Data(query, connection:=connection).To1DList("Column_name")
            Return columnList
        End Function

        Function GenerateClassFromTable(tableName As String, Optional connection As SqlConnection = Nothing) As String
            Dim params As New Dictionary(Of String, Object) From {{"TableName", tableName}}
            Dim classStructure As String = Get_Stored_Proceedure_Data("SQL_Table_To_VB_Class", params:=params, connection:=connection).ToList(Of SQLClass).FirstOrDefault().ClassStructure
            Return classStructure
        End Function

        Class SQLClass
            Property ClassStructure As String
        End Class


        'Function CreateSQLTableClass(tableName As String, Optional classtype As SQLTableClassType = SQLTableClassType.VB, Optional connection As SqlConnection = Nothing)
        '    Dim sql_to_class_query As String = If(classtype = SQLTableClassType.VB, sql_to_vb, sql_to_c_sharp)
        '    Dim query As String = "declare @TableName sysname = " + tableName.ToDBString()
        '    query += sql_to_class_query
        '    Dim result As String = Get_Data(query, connection).Rows(0)("Result").ToString()
        '    Return result
        'End Function

        'Sub ConvertDatabaseTablesToClasses(classtype As SQLTableClassType, connection As SqlConnection)
        '    Dim classes As New List(Of String)
        '    Dim tableList = GetTableNames(connection)
        '    For Each table As String In tableList
        '        Dim result As String = CreateSQLTableClass(table, classtype, connection)
        '        classes.Add(result)
        '    Next
        '    File.WriteAllLines("classes.txt", classes.ToArray())
        '    'Process.Start("classes.txt")
        'End Sub
        '#Region "sql to vb"
        '    Private sql_to_vb As String = "
        '        Declare @TableName sysname = 'tablename'
        'Declare @result varchar(max) = 'Public Class ' + @TableName 
        'Select Case@Result = @Result +' '+ primaryKey+'
        'Property '+ ColumnName+' as '+ RTRIM(Columntype) + NullableSign+' '
        'from
        '(
        '    Select Case
        '        replace(col.name, ' ', '_') ColumnName,
        '        column_id,
        '        Case typ.name
        '             when 'bigint' then 'long'
        '        when 'binary' then 'byte'
        '        when 'bit' then 'boolean'
        '        when 'char' then 'string'
        '        when 'date' then 'DateTime'
        '        when 'datetime' then 'DateTime'
        '        when 'datetime2' then 'DateTime'
        '        when 'datetimeoffset' then 'DateTimeOffset'
        '        when 'decimal' then 'decimal'
        '        when 'float' then 'double'
        '        when 'image' then 'byte'
        '        when 'int' then 'integer'
        '        when 'money' then 'decimal'
        '        when 'nchar' then 'string'
        '        when 'ntext' then 'string'
        '        when 'numeric' then 'decimal'
        '        when 'nvarchar' then 'string'
        '        when 'real' then 'double'
        '        when 'smalldatetime' then 'DateTime'
        '        when 'smallint' then 'short'
        '        when 'smallmoney' then 'decimal'
        '        when 'text' then 'string'
        '        when 'time' then 'TimeSpan'
        '        when 'timestamp' then 'DateTime'
        '        when 'tinyint' then 'byte'
        '        when 'uniqueidentifier' then 'Guid'
        '        when 'varbinary' then 'byte'
        '        when 'varchar' then 'string'
        '        Else 'UNKNOWN_' + typ.name
        '        End + Case When col.is_nullable=1 And typ.name Not In ('binary', 'varbinary', 'image', 'text', 'ntext', 'varchar', 'nvarchar', 'char', 'nchar') THEN ' ' ELSE '' END ColumnType,
        '        			Case 
        '		when col.is_nullable = 1 And typ.name in ('bigint', 'bit', 'date', 'datetime', 'datetime2', 'datetimeoffset', 'decimal', 'float', 'int', 'money', 'numeric', 'real', 'smalldatetime', 'smallint', 'smallmoney', 'time', 'tinyint', 'uniqueidentifier') 
        '		then '?' 
        '		Else '' 
        '        End NullableSign,
        '				    Case
        '                when col.name in (SELECT column_name
        '                                    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS As TC
        '                                    inner join
        '                                        INFORMATION_SCHEMA.KEY_COLUMN_USAGE AS KU
        '                                            On TC.CONSTRAINT_TYPE = 'PRIMARY KEY' AND
        '                                                TC.CONSTRAINT_NAME = KU.CONSTRAINT_NAME And 
        '                                                KU.table_name=@TableName)
        '                                            then '
        '            <PrimaryKey>'
        '                Else ' '	
        '                End primaryKey,
        '		colDesc.colDesc AS ColumnDesc
        '    from sys.columns col
        '        join sys.types typ On
        '            col.system_type_id = typ.system_type_id And col.user_type_id = typ.user_type_id
        '    OUTER APPLY(
        '    SELECT TOP 1 CAST(value As NVARCHAR(max)) As colDesc
        '    FROM
        '       sys.extended_properties
        '    WHERE
        '       major_id = col.object_id
        '       And
        '       minor_id = COLUMNPROPERTY(major_id, col.name, 'ColumnId')
        '    ) colDesc            
        '    where object_id = object_id(@TableName)
        ') t
        'order by column_id

        'Set @result = @result  + '
        'End Class '

        'print @result
        '#End Region

        '#Region "sql to vb"
        '    Private sql_to_c_sharp As String = "
        'declare @Result varchar(max) = 'Public Class' + @TableName + '
        '{'
        'select @Result = @Result +' '+ primeky+'
        'public '+ Columntype + NullableSign +'  ' + ColumnName + ' { get; set; } '
        'from
        '(
        'select 
        '    replace(col.name,' ','_') ColumnName,
        '    column_id ColumnID,
        '    case typ.name
        '        when 'bigint' then 'long'
        '        when 'binary' then 'byte[]'
        '        when 'bit' then 'boolean'
        '        when 'char' then 'string'
        '        when 'date' then 'DateTime'
        '        when 'datetime' then 'DateTime'
        '        when 'datetime2' then 'DateTime'
        '        when 'datetimeoffset' then 'DateTimeOffset'
        '        when 'decimal' then 'decimal'
        '        when 'float' then 'double'
        '        when 'image' then 'byte[]'
        '        when 'int' then 'integer'
        '        when 'money' then 'decimal'
        '        when 'nchar' then 'string'
        '        when 'ntext' then 'string'
        '        when 'numeric' then 'decimal'
        '        when 'nvarchar' then 'string'
        '        when 'real' then 'double'
        '        when 'smalldatetime' then 'DateTime'
        '        when 'smallint' then 'short'
        '        when 'smallmoney' then 'decimal'
        '        when 'text' then 'string'
        '        when 'time' then 'TimeSpan'
        '        when 'timestamp' then 'DateTime'
        '        when 'tinyint' then 'byte'
        '        when 'uniqueidentifier' then 'Guid'
        '        when 'varbinary' then 'byte[]'
        '        when 'varchar' then 'string'
        '        else 'UNKNOWN_' + typ.name
        '    end ColumnType,
        '    case typ.name
        '        when 'binary' then replace(col.name,' ','_')+'()'
        '        when 'image' then replace(col.name,' ','_')+'()'
        '        when 'varbinary' then replace(col.name,' ','_')+'()'
        '        else replace(col.name,' ','_')+'()'
        '    end ColumnName,
        '    case
        '        when col.name in (SELECT column_name
        '                            FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS AS TC
        '                            inner join 
        '                                INFORMATION_SCHEMA.KEY_COLUMN_USAGE AS KU
        '                                    ON TC.CONSTRAINT_TYPE = 'PRIMARY KEY' AND
        '                                        TC.CONSTRAINT_NAME = KU.CONSTRAINT_NAME and 
        '                                        KU.table_name=@TableNames)
        '                                    then '
        '    [PrimaryKey]'
        '        else ' '
        '        end primaryKey,
        '        case 
        '        when col.is_nullable = 1 and type.name in ('bigint','bit','data','datetime',datetime2','datetimeoffset','decimal','float','int','money','numeric','real','smalldatetime','smallint','smallmoney')
        '        then ''
        '        else ''
        '        end NullableSign
        '       from sys.columns col
        'join sys.types typ on col.system_type_id = type.system_type_id AND col.user_type_id = typ.user_type_id
        'where object_id = object_id(@TableName)
        ') t
        'order by ColumnID
        'set @Result = @Result + '
        '}'

        'select @Result as 'Result'
        '"
        '#End Region


    End Module
End Namespace
