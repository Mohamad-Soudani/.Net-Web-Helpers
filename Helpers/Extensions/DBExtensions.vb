Imports System.Data
Imports System.IO
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json.Serialization
Imports tdh.Classes.Common
Imports UnitTest.tdh.Classes.Common
Imports ErrorEventArgs = Newtonsoft.Json.Serialization.ErrorEventArgs

Namespace Helpers.Extensions
    Public Module DBExtensions
        <Extension>
        Public Function ToObject(Of T As {Class, New})(ByVal row As DataRow, Optional index As Integer = 0) As T
            Dim obj As T = New T()
            Return GenerateObject(obj, row, index)
        End Function

        <Extension>
        Public Function ToList(Of T As {Class, New})(ByVal table As DataTable) As List(Of T)
            Dim list As List(Of T) = New List(Of T)
            GenerateObject_Logs.Clear()
            Dim i As Integer = 0
            For Each row As DataRow In table.AsEnumerable()
                Dim obj = row.ToObject(Of T)(i)
                list.Add(obj)
                i += 1
            Next
            Return list
            Return Nothing
        End Function

        <Extension>
        Public Function ToObject(ByVal row As DataRow, T As Type) As Object
            Dim obj = Activator.CreateInstance(T)
            Return GenerateObject(obj, row, 0)
        End Function

        Private Function GenerateObject(ob As Object, row As DataRow, rowindex As Integer)
            Dim className As String = ob.GetType().Name
            Dim allProps As PropertyInfo() = ob.GetType().GetProperties()
            Dim jobj As String = "{ "
            Dim i As Integer = 1
            For Each col As DataColumn In row.Table.Columns
                Dim colValue As Object = row(col.ColumnName)
                Dim filteredColName = col.ColumnName
                Dim prop As PropertyInfo = allProps.Where(Function(x) x.Name = filteredColName).FirstOrDefault()
                If (prop IsNot Nothing AndAlso StringIsNot_Null_Empty_WhiteSpace(colValue.ToString())) Then
                    Try
                        Dim propType = If(Nullable.GetUnderlyingType(prop.PropertyType) Is Nothing, prop.PropertyType, Nullable.GetUnderlyingType(prop.PropertyType))
                        Dim parsedValue As Object
                        If propType Is GetType(Boolean) Then
                            Dim intResult As Integer
                            If Integer.TryParse(colValue, intResult) AndAlso {0, 1}.Any(Function(x) x = intResult) Then
                                parsedValue = Convert.ToBoolean(intResult)
                            Else
                                parsedValue = Boolean.Parse(colValue)
                            End If
                        Else
                            parsedValue = Convert.ChangeType(colValue, propType)
                        End If
                        If (parsedValue IsNot Nothing) Then
                            Dim valueString As String = ""
                            If TypeOf parsedValue Is DateTime Then
                                valueString = CType(parsedValue, Date).ToMyDateTimeString()
                            Else
                                valueString = parsedValue.ToString()
                            End If
                            valueString = valueString.Replace("""", "'")
                            valueString = valueString.Replace("\", "\\")
                            jobj += col.ColumnName + ":""" + valueString.Trim() + ""","
                        End If
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine(ex)
                        Dim log As New Log
                        With log
                            .ColumnName = filteredColName
                            .RowIndex = rowindex
                            .Message = ex.Message
                            .Value = colValue.ToString()
                            .OccuranceTime = DateTime.Now
                        End With
                        Utility.GenerateObject_Logs.Add(log)
                    End Try
                End If
                i += 1
            Next
            jobj += " }"
            Dim obj = Newtonsoft.Json.JsonConvert.DeserializeObject(jobj, ob.GetType)
            Return obj
        End Function



        Public Sub HandleDeserializationError(ByRef sender As Object, ByRef errorArgs As ErrorEventArgs)
            Dim currentError = errorArgs.ErrorContext.Error.Message
            errorArgs.ErrorContext.Handled = True
        End Sub


        <Extension>
        Public Function ToList(ByVal table As DataTable, T As Type) As ArrayList
            Dim list As New ArrayList()
            For Each row As DataRow In table.AsEnumerable
                Dim obj = row.ToObject(T)
                list.Add(obj)
            Next
            Return list
        End Function

        <Extension>
        Public Function ToList(ByVal table As DataTable, propertyName As String) As List(Of String)
            Dim list As New List(Of String)
            For Each row As DataRow In table.AsEnumerable
                Dim obj = row(propertyName)
                list.Add(obj)
            Next
            Return list
        End Function

        <Extension>
        Public Function SingleQuotes(value As String) As String
            Return "'" + value.Trim() + "'"
        End Function

        <Extension>
        Public Function LeftRightLike(ByVal value As String) As String
            Return "%" + value.Trim() + "%".SingleQuotes()
        End Function

        <Extension>
        Public Function ObjectToList(Of T)(obj As T) As List(Of T)
            Dim list As New List(Of T)
            list.Add(obj)
            Return list
        End Function

        <Extension>
        Public Function ObjectToDataTable(Of T)(obj As T) As DataTable
            Return obj.ObjectToList().ToDataTable()
        End Function

        <Extension>
        Public Function ToDBConvertToDate(value As String) As String
            Return "Convert(date," & value & ")"
        End Function

        <Extension>
        Public Function ToDBConvertToYear(value As String) As String
            Return "Year(date," & value & ")"
        End Function

        <Extension>
        Public Function ToDBConvertToMonth(value As String) As String
            Return "Month(date," & value & ")"
        End Function

        <Extension>
        Public Function ToDBConvertToDay(value As String) As String
            Return "Day(date," & value & ")"
        End Function

        <Extension>
        Public Function ToList(dtrow As DataRow, Optional removeEmpty As Boolean = True) As List(Of String)
            Dim list As New List(Of String)
            For Each item As Object In dtrow.ItemArray
                list.Add(item.ToString)
            Next
            If removeEmpty Then
                list = list.Where(Function(x) Not x.ToLower.Contains("null") AndAlso Utility.StringIsNot_Null_Empty_WhiteSpace(x)).ToList()
            End If
            Return list
        End Function
        <Extension>
        Public Sub LoadFullObject(Of T As {Class, New})(ByRef instance As T, Optional connection As SqlClient.SqlConnection = Nothing)
            Dim primaryKeys As String() = DBHelper.CollectPrimaryKeys(instance, GetType(T).GetProperties).Select(Function(x) x.Value).ToArray
            instance = DBHelper.GetItem(Of T)(primaryKeys, connection:=connection)
        End Sub

        <Extension>
        Public Sub LoadProxyObject(Of T As {New, BaseClass(Of T)})(ByRef instance As T)
            instance.Proxy = ConvertToType(Of T)(instance)
        End Sub

    End Module
End Namespace
