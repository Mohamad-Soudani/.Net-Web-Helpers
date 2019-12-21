Imports System.Text.RegularExpressions
Imports tdh.Helpers.Extensions
Imports UnitTest.tdh.Helpers.Extensions

Namespace Helpers
    Public Module RegexHelper
        Function GenerateStringPattern(rawText As String) As String
            Dim charArray As Char() = rawText.ToCharArray()
            Dim pattern As String = ""
            For i As Integer = 0 To charArray.Length - 1
                Dim c As Char = charArray(i)
                If c = "" Then
                    pattern = pattern + "\s?"
                End If
                If (i > 0 AndAlso i Mod 6 = 0) Then
                    pattern += c + "|"
                Else
                    pattern += c
                End If
            Next
            pattern = If(pattern.Length > 5, String.Join("|", pattern.Split("|").ToList().Where(Function(s) s.Count >= 5)), pattern)
            pattern.RemoveMultiple({" ", "(", ")"})
            pattern = ".*" & pattern & ".*"
            Return pattern
        End Function

        Function MatchPartsOfString(searchText As String, searchInText As String) As Boolean
            Dim pattern As String = GenerateStringPattern(searchText)
            Dim match As Boolean = Regex.IsMatch(searchInText, pattern, RegexOptions.IgnoreCase)
            Return match
        End Function

        Function IsEmail(ByVal email As String) As Boolean
            Dim emailRegex As New Regex("^[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,4})$")
            Return emailRegex.IsMatch(email)
        End Function

    End Module
End Namespace
