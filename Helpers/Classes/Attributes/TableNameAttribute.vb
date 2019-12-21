Namespace Classes.Attributes

    <System.AttributeUsage(System.AttributeTargets.Property)>
    Public Class TableNameAttribute : Inherits Attribute
        Property TableName As String
        Public Sub New(tableName As String)
            Me.TableName = tableName
        End Sub
    End Class

End Namespace