Imports System.ComponentModel
Imports System.Data
Imports System.Runtime.CompilerServices
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Namespace Helpers.Extensions

    Public Module Extensions

        <Extension()>
        Function ToBase64(ByVal bytes As Byte()) As String
            Return Convert.ToBase64String(bytes)
        End Function

        <Extension()>
        Function FileNameFromPath(path As String) As String
            Return path.Split("\").LastOrDefault()
        End Function

        <Extension()>
        Function Truncate(ByVal value As String, ByVal maxChars As Integer) As String
            Return If(value.Length <= maxChars, value, value.Substring(0, maxChars))
        End Function

        <Extension()>
        Function ToMyDateTimeString(dt As Date, Optional removeTimeColons As Boolean = False) As String
            Dim datetimeString As String = dt.ToString("yyyy-MM-dd HH:mm:ss")
            If (removeTimeColons) Then
                datetimeString = datetimeString.Replace(":", "")
            End If
            Return datetimeString
        End Function



        <Extension()>
        Function ToMyDateString(ByVal dt As DateTime) As String
            Return dt.ToString("yyyy-MM-dd")
        End Function

        <Extension()>
        Function ToDBDateTimeString(ByVal dt As DateTime) As String
            Return dt.ToMyDateTimeString().ToDBString()
        End Function

        <Extension()>
        Function ToDBDateString(ByVal dt As DateTime) As String
            Return dt.ToMyDateTimeString().ToDBString()
        End Function

        <Extension()>
        Function ToDataBaseCommaSeparatedString(Of T)(ByVal itemList As IEnumerable(Of T), ByVal [property] As String) As String
            Dim concatValues = ""
            If itemList.Count() > 0 Then
                Dim propertyInfo As Reflection.PropertyInfo = GetType(T).GetProperty([property])
                Dim selectedValuesList As List(Of String) = itemList.ToList().Select(Of String)(Function(item) propertyInfo.GetValue(item, Nothing)).ToList()
                concatValues = ConcatinateListToDBCommaSeparatedString(selectedValuesList)
            End If
            Return concatValues
        End Function

        <Extension()>
        Function ToCommaSeparatedString(Of T)(ByVal itemList As IEnumerable(Of T)) As String
            Dim concatValues = ""
            If itemList.Count() > 0 Then
                concatValues = "" + String.Join(",", itemList)
            End If
            Return concatValues
        End Function

        <Extension()>
        Function ToDataBaseCommaSeparatedString(Of T)(ByVal list As IEnumerable(Of T), Optional addParanthesis As Boolean = True, Optional addPrefix As Boolean = False) As String
            Dim concatValues = "('')"
            If list.Count() > 0 Then
                Dim selectedValuesList As IEnumerable(Of String) = Helpers.ConvertToType(list, GetType(IEnumerable(Of String)))
                Dim prefix As String = If(addPrefix, "N", "")
                concatValues = ConcatinateListToDBCommaSeparatedString(selectedValuesList, prefix)
                If addParanthesis Then
                    concatValues = "(" + concatValues + ")"
                End If
            End If
            Return concatValues
        End Function

        Private Function ConcatinateListToDBCommaSeparatedString(ByRef selectedValuesList As IEnumerable(Of String), Optional prefix As String = "") As String
            Return "'" + String.Join("'," + prefix + "'", selectedValuesList) + "'"
        End Function

        <Extension()>
        Function ToDBString(ByVal value As Object) As String
            Return "'" + value.ToString().Trim() + "'"
        End Function

        <Extension()>
        Function RemoveAttachmentPrints(ByVal value As String) As String
            Return value.Replace("'@", "").Replace("@'", "")
        End Function

        <Extension()>
        Function ToDataTable(Of T)(data As IEnumerable(Of T)) As DataTable
            Dim properties As PropertyDescriptorCollection = TypeDescriptor.GetProperties(GetType(T))
            Dim table As New DataTable()
            For Each prop As PropertyDescriptor In properties
                table.Columns.Add(prop.Name, If(Nullable.GetUnderlyingType(prop.PropertyType), prop.PropertyType))
            Next
            For Each item As T In data
                Dim row As DataRow = table.NewRow()
                For Each prop As PropertyDescriptor In properties
                    row(prop.Name) = If(prop.GetValue(item), DBNull.Value)
                Next
                table.Rows.Add(row)
            Next
            Return table
        End Function

        <Extension>
        Function To1DList(data As DataTable, column As String) As List(Of String)
            Dim temp As New List(Of String)
            If Helpers.StringIsNot_Null_Empty_WhiteSpace(column) And data IsNot Nothing Then
                For Each row As DataRow In data.Rows
                    Dim cell As String = row(column).ToString()
                    Dim splitValues As String() = cell.Split(New String() {vbCrLf, vbLf, " "c}, StringSplitOptions.None)
                    temp.AddRange(splitValues)
                Next
            End If
            temp = temp.Where(Function(x) Helpers.StringIsNot_Null_Empty_WhiteSpace(x)).ToList()
            Return temp
        End Function

        <Extension()>
        Function IsEmptyString(ByVal value As String) As Boolean
            Return Helpers.StringIs_Null_Empty_WhiteSpace(value)
        End Function

        <Extension()>
        Function IsNotEmptyString(ByVal value As String) As Boolean
            Return Helpers.StringIsNot_Null_Empty_WhiteSpace(value)
        End Function

        <Extension()>
        Function GetTime(value As Date) As String
            Return value.ToString("hh:mm tt")
        End Function

        <Extension()>
        Function ConvertTimeToDate(ByVal time) As Date
            Dim dateString = Date.Now.ToString("yyyy/MMM/dd")
            dateString &= " " + time
            Return dateString
        End Function

        <Extension()>
        Function FormatDate_yyyyMMMdd(time As Date) As String
            Dim dateString = time.ToString("yyyy/MMM/dd")
            Return dateString
        End Function

        <Extension()>
        Function EscapeSqlString(text As String) As String
            Return text.Replace("'", "''").Trim()
        End Function

        <Extension()>
        Function ToList(Rows As DataRowCollection) As List(Of DataRow)
            Dim list As List(Of DataRow) = New List(Of DataRow)
            For Each row As DataRow In Rows
                list.Add(row)
            Next
            Return list
        End Function

        <Extension()>
        Function RemoveMultiple(str As String, replaceArr As String()) As String
            Dim concatList As New List(Of Char)
            Dim arr As Char() = str.ToCharArray()
            For Each c As Char In arr
                If Not replaceArr.Any(Function(value) c = value) Then
                    concatList.Add(c)
                End If
            Next
            str = String.Join("", concatList)
            Return str
        End Function

        <Extension()>
        Function ConvertListToType(Of T As {Class, New}, R As {Class, New})(lst As List(Of T)) As List(Of R)
            Dim resultList As New List(Of R)
            For Each item As Object In lst
                Dim result As R = Helpers.ConvertToType(Of R)(item)
                resultList.Add(result)
            Next
            Return resultList
        End Function

        <Extension()>
        Function ReadLines(value As String) As List(Of String)
            Dim valueList As List(Of String) = value.Split(vbCrLf).ToList()
            Return valueList
        End Function

        <Extension()>
        Function Serialize_Indented(Of T As {New, Class})(value As T) As String
            Dim settings As New JsonSerializerSettings()
            With settings
                .Formatting = Formatting.Indented
            End With
            Dim json As String = Newtonsoft.Json.JsonConvert.SerializeObject(value, settings)
            Return json
        End Function

        <Extension()>
        Function Serialize(Of T As {New, Class})(value As T) As String
            Dim json As String = Newtonsoft.Json.JsonConvert.SerializeObject(value)
            Return json
        End Function

        <Extension>
        Function DeSerialize(Of T As {New, Class})(value As String) As T
            Try
                Dim obj As T = Newtonsoft.Json.JsonConvert.DeserializeObject(Of T)(value)
                Return obj
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        <Extension>
        Function JsonObjectDetails(Jobj As JObject) As List(Of String)
            Try
                Dim objTokens As JEnumerable(Of JToken) = Jobj.Children(Of JToken)
                Dim embeddedObjects As New List(Of String)
                For Each token As JToken In objTokens
                    Dim name As String = token.Path
                    Try
                        Dim arr As JArray = Jobj(name)
                        embeddedObjects.Add(name & ": " & arr.Count & " items (Array)")
                    Catch ex As Exception
                        embeddedObjects.Add(name & ": (Property/Object")
                    End Try
                Next
                Return embeddedObjects
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        <Extension>
        Sub CopyToClipboard(value As String)
            'Try
            '    System.Windows.Forms.Clipboard.Clear()
            '    System.Windows.Forms.Clipboard.SetText(value)
            '    MessageBox.Show("Copied to clipboard!")
            'Catch ex As Exception
            '    MessageBox.Show(ex.Message)
            'End Try
        End Sub

        '<Extension>
        'Function GetHeaders(data As DataTable, Optional keepHeaders As Boolean = True) As List(Of String)
        '    Dim headers As List(Of String) = data.Rows(0).ToList()
        '    If Not keepHeaders Then
        '        RemoveHeaders(data)
        '    End If
        '    Return headers
        'End Function

        <Extension>
        Public Sub RemoveHeaders(data As DataTable)
            data.Rows.RemoveAt(0)
        End Sub
        ''' <summary>
        ''' Calculates the year difference between two different date times and returns an absolute (positive) result
        ''' </summary>
        ''' <param name="date1"></param>
        ''' <param name="date2"></param>
        ''' <returns></returns>
        <Extension>
        Public Function SubtractYears(date1 As DateTime, date2 As DateTime) As Double
            Dim diff As TimeSpan = date1.Subtract(date2)
            Return Math.Abs(diff.Days / 365)
        End Function

    End Module
End Namespace
