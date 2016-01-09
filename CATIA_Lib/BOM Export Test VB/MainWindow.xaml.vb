Class MainWindow
    Dim oProduct As New CATIA_Lib.Cl_CATIA._3D.Product
    Private Sub button_Click(sender As Object, e As RoutedEventArgs) Handles button.Click
        oProduct.test()

        PartNumber.Text = oProduct.SelectSingle3DProduct

    End Sub
End Class
