Namespace Classes.Attributes

    <System.AttributeUsage(System.AttributeTargets.Property)>
    Public Class PrimaryKeyAttribute : Inherits Attribute
        Property KeyIndex As Integer = 0
        Property AutoIncrement As Boolean = False
        Public Sub New()

        End Sub

        Public Sub New(autoIncrement As Boolean)
            Me.AutoIncrement = autoIncrement
        End Sub

        Public Sub New(keyIndex As Integer)
            Me.KeyIndex = keyIndex
        End Sub

        Public Sub New(keyIndex As Integer, autoIncrement As Boolean)
            Me.KeyIndex = keyIndex
            Me.AutoIncrement = autoIncrement
        End Sub

    End Class

End Namespace