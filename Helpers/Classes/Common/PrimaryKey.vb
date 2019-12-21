Namespace Classes.Common
    Public Class PrimaryKey
        Property Name As String
        Property Value As String
        Public Sub New(ByVal name As String, ByVal value As String)
            Me.Name = name
            Me.Value = value
        End Sub
    End Class
End Namespace
