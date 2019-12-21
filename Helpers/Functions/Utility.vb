Imports System.ComponentModel
Imports System.Data
Imports System.Diagnostics
Imports System.IO
Imports ClosedXML.Excel
Imports Newtonsoft.Json
Imports tdh.Classes.Common
Imports tdh.Helpers.Extensions
Imports UnitTest.tdh.Classes.Common
Imports UnitTest.tdh.Helpers.Extensions

Namespace Helpers

    Public Module Utility
        Dim counter As Integer = 0
        Public GenerateObject_Logs As New List(Of Log)
        Function StringIsNot_Null_Empty_WhiteSpace(ByVal value As String) As String
            If String.IsNullOrEmpty(value) Or String.IsNullOrWhiteSpace(value) Then
                Return False
            End If
            Return True
        End Function

        Function StringIs_Null_Empty_WhiteSpace(ByVal value As String) As String
            If Not String.IsNullOrEmpty(value) Or Not String.IsNullOrWhiteSpace(value) Then
                Return False
            End If
            Return True
        End Function

        Function RemoveQuote(ByVal text As String) As String
            Return text.Replace("'", "").Trim()
        End Function


        Public Enum ExportMode
            CSV
            Tab
        End Enum

        Public Enum ExportExtensions
            csv
            txt
        End Enum



        Public Sub ShowConfirmationBox(msg As String, title As String, ByRef func As Action)
            'title = If(title.IsEmptyString(), Application.ProductName, title)
            'Dim confirmationResult As DialogResult = MessageBox.Show(msg, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            'If confirmationResult.Equals(DialogResult.Yes) Then
            '    func?.Invoke()
            'End If
        End Sub

        Public Function ConvertTotype(Of TS As {Class, New}, TR As {Class, New})(ByVal obj As TS) As TR
            Dim jsonString As String = Newtonsoft.Json.JsonConvert.SerializeObject(obj)
            Return Newtonsoft.Json.JsonConvert.DeserializeObject(Of TR)(jsonString)
        End Function

        Public Function ConvertToType(Of TS As {Class, New})(ByVal obj As Object) As TS
            Dim T As Type = GetType(TS)
            Return ConvertToType(obj, T)
        End Function

        Public Function ConvertToType(ByVal obj As Object, t As Type)
            Dim settings = New JsonSerializerSettings()
            With settings
                .PreserveReferencesHandling = PreserveReferencesHandling.Objects
                .Formatting = Formatting.Indented
            End With
            Dim jsonString As String = Newtonsoft.Json.JsonConvert.SerializeObject(obj, settings)
            Return Newtonsoft.Json.JsonConvert.DeserializeObject(jsonString, t)
        End Function

        Public Sub Done()
            'MessageDialoges.InformationMsg("Done")
        End Sub

        'Public Function GetFrame(Of T As {Class, New})() As Form
        '    Dim OpenForms = Application.OpenForms
        '    Dim foundFrame As Boolean
        '    Dim frm As Form
        '    For Each form As Form In OpenForms
        '        If form.GetType() Is GetType(T) Then
        '            frm = form
        '            foundFrame = Trues
        '        End If
        '    Next
        'End Function

        Public Sub WriteTotextFile(name As String, lines As IEnumerable(Of String))
            Dim textFile As String = name & ".txt"
            File.WriteAllLines(textFile, lines)
            Process.Start(textFile)
        End Sub

        Public Sub WriteTextFile(name As String, text As String)
            Dim textFile As String = name + ".txt"
            File.WriteAllText(textFile, text)
            Process.Start(textFile)
        End Sub

        Public Function CreateGuid() As Guid
            Dim seed As Guid = Guid.NewGuid()
            Return seed
        End Function

        Function SeedString(ByRef val As String) As String
            Dim nameParts As String() = val.Split(".")
            Dim seed As String = Date.Now.ToMyDateTimeString(True)
            val = nameParts.FirstOrDefault() + seed.ToString() + "." + nameParts.LastOrDefault()
            Return val
        End Function

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

        Function FileNameFromPath(path As String) As String
            Return path.Split("\").LastOrDefault()
        End Function



    End Module
End Namespace
