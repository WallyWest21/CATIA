Class MainWindow
    Dim oDrawing As New cl_VB_CATIALib.Drawing
    Private Sub button_Click(sender As Object, e As RoutedEventArgs) Handles button.Click
        Dim oDrawing As New cl_VB_CATIALib.Drawing
        'MsgBox(oDrawing.PartsList.Item(1).PartNo)
        MyListBox.ItemsSource = oDrawing.PartsList
        'MsgBox(oDrawing.Select2DTable.GetCellString(1, 1))
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        'MyListBox.ItemsSource = oDrawing.PartsList
    End Sub
End Class
