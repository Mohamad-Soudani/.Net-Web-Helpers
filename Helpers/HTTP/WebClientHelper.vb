Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Module WebClientHelper

    Public Function GetData(Of T As {New, Class})(url As String, headers As Dictionary(Of String, String)) As T
        Dim client As New System.Net.WebClient
        Using client
            For Each item In headers
                client.Headers.Add(item.Key, item.Value)
            Next
            client.Encoding = Encoding.UTF8
            Dim result As String = client.DownloadString(url)
            Dim data As T = JsonConvert.DeserializeObject(Of T)(result)
            Return data
        End Using
    End Function

    Public Sub DownloadFile(url As String, filename As String, headers As Dictionary(Of String, String))
        Dim client As New System.Net.WebClient
        Using client
            For Each item In headers
                client.Headers.Add(item.Key, item.Value)
            Next
            client.Headers.Add("user-agent", " Mozilla/5.0 (Windows NT 6.1; WOW64; rv:25.0) Gecko/20100101 Firefox/25.0")
            client.UseDefaultCredentials = True
            client.Proxy.Credentials = System.Net.CredentialCache.DefaultCredentials
            client.DownloadFile(url, filename)
        End Using
    End Sub


End Module
