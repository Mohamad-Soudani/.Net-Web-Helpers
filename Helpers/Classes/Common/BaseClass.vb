Imports dotNetWebHelpers.Classes.Attributes

Namespace Classes.Common

    Public Class BaseClass(Of T)
        <ProxyObject>
        Property Proxy As T
    End Class

End Namespace