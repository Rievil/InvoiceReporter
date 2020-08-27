Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1
    Dim WS As ThisAddIn
    Property Sheet As Object
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        MsgBox("Iniciace Ribbon1")
    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click

    End Sub

    Private Sub Workbook_Actiate()
        MsgBox("Aktivoval jsi workbook")
    End Sub
End Class
