Class MainWindow
    Private Sub button_Click(sender As Object, e As RoutedEventArgs) Handles button.Click
        Dim oDrawing As New cl_VB_CATIALib.Drawing
        MsgBox(oDrawing.PartsList.Item(1).PartNo)
        'MsgBox(oDrawing.Select2DTable.GetCellString(1, 1))
    End Sub
End Class
