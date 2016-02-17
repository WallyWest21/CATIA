Class MainWindow
    Dim oProduct As New CATIA_Lib.Cl_CATIA._3D.oProduct
    Dim oDrawing As New CATIA_Lib.Cl_CATIA.Drawing
    Dim Panel As New CATIA_Lib.Cl_CATIA.UDF.Panel
    Dim oPart As New CATIA_Lib.Cl_CATIA._3D.oPart

    Private Sub button_Click(sender As Object, e As RoutedEventArgs) Handles button.Click
        'oProduct.test()

        'PartNumber.Text = oProduct.SelectSingle3DProduct
        'MsgBox(oDrawing.PartsList.Item(1).PartNo)
        'MsgBox(oDrawing.PartsList.Item(1).PartNo)
        'MsgBox(oProduct.PartsList.Item(0).PartNo)
    End Sub

    Private Sub button1_Click(sender As Object, e As RoutedEventArgs) Handles button1.Click
        MsgBox(oProduct.PartsList.Item(0).PartNo)
    End Sub

    Private Sub button2_Click(sender As Object, e As RoutedEventArgs) Handles button2.Click
        'Panel.pad()

        Dim Drawer As New CATIA_Lib.Cl_CATIA.UDF.Drawer(50, 40, 15, 20, 25, 14)
        Drawer.CreateFrontPanel()
        'oPart.CreatePlanefromOffset("Bottom")
        'oPart.Pad("Bottom")
    End Sub
End Class
