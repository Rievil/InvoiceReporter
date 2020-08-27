
Public Class ThisAddIn

    Public SSMain As Object
    Public Property CurrentSheet As Object

    Dim Test As Ribbon1

    Private Sub ThisAddIn_Startup() Handles Me.Startup

        CurrentSheet = Me.Application.Worksheets("List1")
        MsgBox("Iniciace ThisAddIn")
    End Sub

    Public Function GetCurrSheet() As Object
        GetCurrSheet = CurrentSheet
    End Function

End Class
