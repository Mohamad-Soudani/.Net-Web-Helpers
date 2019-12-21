Imports System.Runtime.CompilerServices
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports Microsoft.VisualBasic
Imports tdh.Interfaces
Imports UnitTest.tdh.Interfaces

Namespace Helpers.Controls

    Public Class GridViewExtensions
        Implements IBind
        Public Sub Bind(ctrl As Control, data As Object) Implements IBind.Bind
            Dim grd As GridView = CType(ctrl, GridView)
            grd.DataSource = data
            grd.DataBind()
        End Sub
    End Class

    Public Module GridViewUtility
        Dim ext As New GridViewExtensions
        <Extension>
        Public Sub Bind(grd As GridView, data As Object)
            ext.Bind(grd, data)
        End Sub

        Public Function ExtractValues(Of TControl)(grd As GridView, controlName As String) As Object
            Select Case GetType(TControl)
                Case GetType(TextBox)
                    Return (TryCast(grd.FooterRow.FindControl("F_txtActivitySection"), TextBox)).Text
                Case GetType(CheckBox)
                    Return (TryCast(grd.FooterRow.FindControl("F_txtActivitySection"), CheckBox)).Checked
                Case Else
                    Return Nothing
            End Select
        End Function

    End Module

End Namespace