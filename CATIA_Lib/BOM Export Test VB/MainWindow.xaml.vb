﻿Class MainWindow
    Dim oProduct As New CATIA_Lib.Cl_CATIA._3D.Product
    Private Sub button_Click(sender As Object, e As RoutedEventArgs) Handles button.Click
        oProduct.test()

        PartNumber.Text = oProduct.SelectSingle3DProduct
        'MsgBox(oDrawing.PartsList.Item(1).PartNo)
    End Sub
End Class
