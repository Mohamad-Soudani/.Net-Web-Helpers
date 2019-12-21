Imports System.Runtime.CompilerServices
Imports System.Web.UI.WebControls
Imports Microsoft.VisualBasic
Namespace Helpers.Controls

    Public Module DropDownUtility

        <Extension>
        Public Sub Bind(ByRef drp As DropDownList, data As Object)
            drp.DataSource = data
            drp.DataBind()
        End Sub

        <Extension>
        Public Sub Bind(ByRef drp As DropDownList, data As Object, datatextfield As String, datavaluefield As String)
            drp.DataSource = data
            drp.DataTextField = datatextfield
            drp.DataValueField = datavaluefield
            drp.DataBind()
        End Sub

        <Extension>
        Public Sub BindEnum(ByRef drp As DropDownList, t As Type)
            Dim filterNames As String() = [Enum].GetNames(t).Select(Function(x) x.Replace("_", " ")).ToArray
            Dim filterValues As Integer() = [Enum].GetValues(t)
            For i As Integer = 0 To filterNames.Length - 1
                Dim item As New ListItem(filterNames(i), filterValues(i))
                drp.Items.Add(item)
            Next
        End Sub

    End Module

End Namespace