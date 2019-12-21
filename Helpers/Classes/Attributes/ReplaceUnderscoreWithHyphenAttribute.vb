Namespace Classes.Attributes

    <System.AttributeUsage(System.AttributeTargets.Class)>
    Public Class ReplaceUnderscoreWithHyphenAttribute : Inherits Attribute
        Property HyphenLocations As Integer()

        Public Sub New(ParamArray hyphenLocations As Integer())
            Me.HyphenLocations = hyphenLocations
        End Sub
    End Class

End Namespace