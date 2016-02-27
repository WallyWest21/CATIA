Class MainWindow
    Dim oProduct As New CATIA_Lib.Cl_CATIA._3D.oProduct
    Dim oDrawing As New CATIA_Lib.Cl_CATIA.Drawing
    Dim Panel As New CATIA_Lib.Cl_CATIA.UDF.Panel
    Dim oPart As New CATIA_Lib.Cl_CATIA._3D.oPart
    Dim CA As New CATIA_Lib.Cl_CATIA.UDF.ClashAnalysis

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
        Drawer.Create()
        'oPart.CreatePlanefromOffset("Bottom")
        'oPart.Pad("Bottom")
    End Sub

    Private Sub button3_Click(sender As Object, e As RoutedEventArgs) Handles button3.Click
        'oProduct.CreateANewProduct("Drawer")
        oProduct.SubInsertANewPart()

    End Sub

    Private Sub button4_Click(sender As Object, e As RoutedEventArgs) Handles button4.Click
        CA.ActiveProductClash()
    End Sub

    Private Sub button5_Click(sender As Object, e As RoutedEventArgs) Handles button5.Click
        Dim Notes As New List(Of String)(New String() {"REFER TO DOCUMENT  B6119GAFP AS THE DESIGN AUTHORITY FOR ALL MATERIAL AND FINISH SPECIFICATIONS. SYMBOL  INDICATES GRAIN DIRECTION.",
                              "amazon",
                              "yangtze",
                              "mississippi",
                              "yellow"})
        oDrawing.WriteNotesToDrawing(Notes)
    End Sub
End Class
