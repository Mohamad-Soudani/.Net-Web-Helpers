Imports dotNetWebHelpers.Classes.Attributes
Imports dotNetWebHelpers.Helpers

Namespace Classes.Attributes

    <System.AttributeUsage(System.AttributeTargets.Property)>
    Public Class ForeignObjectAttribute : Inherits DBIgnore
        Property TableName As String
        Property PrimaryKey As String
        Property ForeignKey As String
        Property Type As Type
        Property SingleObject As Boolean

        Public Sub New(type As Type, primaryKey As String, foreignKey As String)
            Me.Type = type
            Me.TableName = DBHelper.FixTableNameWithAttributes(type.Name, type)
            Me.PrimaryKey = primaryKey
            Me.ForeignKey = foreignKey
        End Sub


    End Class

End Namespace