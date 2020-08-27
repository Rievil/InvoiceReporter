Public Class ThisAddIn
    Public CS = Me.Application.Worksheets("List1")

    Private Sub ThisAddIn_Startup() Handles Me.Startup


        CS.Cells(1, 1) = "Test1"

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
