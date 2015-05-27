Public Class Cl_XL
    Shared XL As Object

    Shared Function GetXL() As Object

        XL = GetObject(, "CATIA.Application")
        If XL = GetObject(, "CATIA.Application") Is Nothing Or Err.Number <> 0 Then
            MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a CATIA session", vbCritical, "Open a CATIA Session ")
            'Environment.Exit(0)
            XL = CreateObject("Excel.Application")
            XL.Visible = True
        End If

        GetXL = XL
    End Function
End Class
