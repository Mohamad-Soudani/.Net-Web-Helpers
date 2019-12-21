
Imports System.Net.Http
Imports System.Text
Imports Newtonsoft.Json.Linq
Imports tdh.Helpers.Extensions
Imports UnitTest.tdh.Helpers.Extensions

Module HttpHelper
    Function PostData(url As String, obj As String) As String
        Try
            Dim client As New HttpClient
            Dim json As String = obj
            Dim content As New StringContent(json, Encoding.UTF8, "application/json")
            Dim response As String = client.PostAsync(url, content).Result.Content.ReadAsStringAsync().Result
            Dim Jobj As JObject = response.DeSerialize(Of JObject)
            Return Jobj.Serialize_Indented()
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Function GetData(Of T As {New, Class})(url As String) As T
        Try
            Dim json As String = GetData(url)
            Dim obj As T = json.DeSerialize(Of T)
            Return obj
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Function GetData(url As String) As String
        Try
            Dim client As New HttpClient()
            Dim response As String = client.GetAsync(url).Result.Content.ReadAsStringAsync().Result
            Dim obj As JArray = response.DeSerialize(Of JArray)
            Return obj.Serialize_Indented()
        Catch ex As Exception
            Throw ex
        End Try
    End Function

End Module
