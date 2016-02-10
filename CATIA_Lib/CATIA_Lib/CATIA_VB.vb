Imports DrawingViews = DRAFTINGITF.DrawingViews
Imports DrawingView = DRAFTINGITF.DrawingView
Imports DrawingSheets = DRAFTINGITF.DrawingSheets
Imports DrawingSheet = DRAFTINGITF.DrawingSheet

'Imports DRAFTINGITF.IID_DraftingInterfaces
Imports DRAFTINGITF
Imports DRAFTINGITF.CatTextProperty
Imports DRAFTINGITF.CatTablePosition

Imports MECMOD

Imports ProductStructureTypeLib
Imports ProductStructureTypeLib.CatWorkModeType 'apply design mode

Imports INFITF
Imports INFITF.CATMultiSelectionMode
Imports System.Linq

Public Class Cl_CATIA
    Shared oCATIA As INFITF.Application
    Shared Function GetCATIA() As INFITF.Application
        oCATIA = GetObject(, "CATIA.Application")
        If oCATIA Is Nothing Or Err.Number <> 0 Then
            MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a CATIA session", vbCritical, "Open a CATIA Session ")
            Err.Clear()
            Exit Function
            'Environment.Exit(0)
            '       Set CATIA = CreateObject("CATIA.Application")
            '       CATIA.Visible = True
        End If

        GetCATIA = oCATIA
    End Function
    Public Function IsCATIAOpen() As Boolean
        'Try
        '    oCATIA = GetObject(, "CATIA.Application")
        'Catch ex As Exception
        '    MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a CATIA session", vbCritical, "Open a CATIA Session ")
        '    Return False
        '    Exit Function
        'End Try

        If oCATIA Is Nothing Or Err.Number <> 0 Then
            MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a CATIA session", vbCritical, "Open a CATIA Session ")
            Err.Clear()
            Return False
            Exit Function            '       Set CATIA = CreateObject("CATIA.Application")
            '       CATIA.Visible = True
        Else
            Return True
        End If
        'Return True

    End Function
    Public Class _3D
        Public Class oProduct
            Public Sub test()
                MsgBox("hi")
            End Sub
            Function GetProductDocument() As ProductDocument
                oCATIA = GetCATIA()
                Dim MyProductDocument As ProductDocument

                On Error Resume Next
                MyProductDocument = oCATIA.ActiveDocument
                If MyProductDocument Is Nothing Or Err.Number <> 0 Then
                    ' MsgBox "No CATIA Active Document found "
                    MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a Product" & vbCrLf & "in the Active session", vbCritical, "Open a Product")
                    Err.Clear()
                    Environment.Exit(0)
                End If
                GetProductDocument = MyProductDocument
            End Function
            Public Function IsAProductDocumentOpen() As Boolean
                oCATIA = GetCATIA()
                Dim MyProductDocument As ProductDocument

                On Error Resume Next
                MyProductDocument = oCATIA.ActiveDocument
                If MyProductDocument Is Nothing Or Err.Number <> 0 Then
                    ' MsgBox "No CATIA Active Document found "
                    MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a Product" & vbCrLf & "in the Active session", vbCritical, "Open a Product")
                    Err.Clear()
                    Return False
                    Exit Function
                End If
                Return True
            End Function
            Public Function SelectSingle3DProduct() As Product
                Dim ActiveProductDocument As ProductDocument, ActiveProduct As Product

                ActiveProductDocument = GetProductDocument()

                Dim What(0) 'As Object
                What(0) = "Product"

                Dim SelectedProduct As Selection
                SelectedProduct = ActiveProductDocument.Selection
                SelectedProduct.Clear()

                Dim e 'As String
                e = SelectedProduct.SelectElement3(What, "Select a Product or a Component", False, CATMultiSelTriggWhenUserValidatesSelection, False)

                ActiveProduct = SelectedProduct.Item(1).Value
                SelectedProduct.Clear()

                Return ActiveProduct
            End Function
            Public Function PartsList() As List(Of cl_PartsList)
                Dim cl_PL As New cl_PartsList, oPartsList As New List(Of cl_PartsList)
                Dim item As Integer, ActiveProduct As Product

                ActiveProduct = SelectSingle3DProduct()

                For item = 1 To ActiveProduct.Products.Count
                    cl_PL = New cl_PartsList

                    cl_PL.PartNo = ActiveProduct.Products.Item(item).PartNumber
                    cl_PL.Nomenclature = ActiveProduct.Products.Item(item).Nomenclature
                    'cl_PL.Material = ActiveProduct.Products.Item(item).Parameters.Item("Material").Value
                    'cl_PL.Manufacturer = ActiveProduct.Products.Item(item).Parameters.Item("Manufacturer").Value

                    oPartsList.Add(cl_PL)

                Next item
                cl_PL = Nothing

                Call WalkDownTree(ActiveProduct)

                Dim Children = From Child In AllPartsList
                               Group Child By Child.PartNo, Child.Nomenclature, Child.ParentPartNo, Child.ParentNomenclature Into Group
                               Order By PartNo
                               Select Quantity = Group.Count, PartNo = PartNo, Nomenclature = Nomenclature, ParentPartNo = ParentPartNo, ParentNomenclature = ParentNomenclature

                Dim ChildrenList As New List(Of cl_PartsList)

                For Each child In Children
                    cl_PL = New cl_PartsList
                    cl_PL.Quantity = child.Quantity
                    cl_PL.PartNo = child.PartNo
                    cl_PL.Nomenclature = child.Nomenclature

                    ChildrenList.Add(cl_PL)
                Next

                cl_PL = Nothing

                Return ChildrenList
            End Function
            Public AllPartsList As New List(Of cl_PartsList)
            Sub WalkDownTree(ActiveProduct As Product)
                Dim cl_PL As New cl_PartsList, oInstances As Products = ActiveProduct.Products

                '-----No instances found then this is CATPart
                If oInstances.Count = 0 Then
                    Exit Sub
                End If

                Try
                    Parallel.For(1, oInstances.Count + 1, Sub(k)
                                                              cl_PL = New cl_PartsList
                                                              Dim oInst As Product

                                                              oInst = oInstances.Item(k)
                                                              oInst.ApplyWorkMode(DESIGN_MODE)   'apply design mode

                                                              cl_PL.PartNo = oInst.PartNumber
                                                              cl_PL.Nomenclature = oInst.Nomenclature
                                                              cl_PL.ParentPartNo = oInst.Parent.Parent.PartNumber
                                                              cl_PL.ParentNomenclature = oInst.Parent.Parent.Nomenclature
                                                              AllPartsList.Add(cl_PL)
                                                              Call WalkDownTree(oInst)

                                                          End Sub)
                    cl_PL = Nothing

                Catch ex As Exception
                    MsgBox("You need a multicore computer")
                End Try

            End Sub
            Public Function SelectMultiple3DProducts() As List(Of Product)

                'Dim SelectedProducts As Products
                Dim ActiveProductDocument As ProductDocument, ActiveProducts As New List(Of Product)
                Dim counter As Integer

                ActiveProductDocument = GetProductDocument()

                Dim What(0) As Object
                What(0) = "Product"

                Dim SelectedProducts As Selection
                SelectedProducts = oCATIA.ActiveDocument.Selection
                SelectedProducts.Clear()

                Dim e As String
                e = SelectedProducts.SelectElement3(What, "Select a Product or a Component", False, 2, False)

                For counter = 1 To SelectedProducts.Count
                    ActiveProducts.Add(SelectedProducts.Item(counter).Value)
                Next

                SelectedProducts.Clear()

                Return ActiveProducts
            End Function

            Public Class cl_PartsList

                Private _PartNo As String
                Public Property PartNo() As String
                    Get
                        Return _PartNo
                    End Get
                    Set(ByVal value As String)
                        _PartNo = value
                    End Set
                End Property

                Private _Quantity As Integer
                Public Property Quantity() As Integer
                    Get
                        Return _Quantity
                    End Get
                    Set(ByVal value As Integer)
                        _Quantity = value
                    End Set
                End Property

                Private _Nomenclature As String
                Public Property Nomenclature() As String
                    Get
                        Return _Nomenclature
                    End Get
                    Set(ByVal value As String)
                        If value = vbNullString Then
                            _Nomenclature = "N/A"
                        Else
                            _Nomenclature = value
                        End If
                    End Set
                End Property

                Private _Description As String
                Public Property Description() As String
                    Get
                        Return _Description
                    End Get
                    Set(ByVal value As String)
                        If value = vbNullString Then
                            _Description = "N/A"
                        Else
                            _Description = value
                        End If
                    End Set
                End Property

                Private _Manufacturer As String
                Public Property Manufacturer() As String
                    Get
                        Return _Manufacturer
                    End Get
                    Set(ByVal value As String)
                        If value = vbNullString Then
                            _Manufacturer = "N/A"
                        Else
                            _Manufacturer = value
                        End If
                    End Set
                End Property

                Private _Material As String
                Public Property Material() As String
                    Get
                        Return _Material
                    End Get
                    Set(ByVal value As String)

                        If value = vbNullString Then
                            _Material = "N/A"
                        Else
                            _Material = value
                        End If
                    End Set
                End Property

                Private _ParentPartNo As String
                Public Property ParentPartNo() As String
                    Get
                        Return _ParentPartNo
                    End Get
                    Set(ByVal value As String)
                        _ParentPartNo = value
                    End Set
                End Property
                Private _ParentNomenclature As String
                Public Property ParentNomenclature() As String
                    Get
                        Return _ParentNomenclature
                    End Get
                    Set(ByVal value As String)
                        _ParentNomenclature = value
                    End Set
                End Property

                Public Class cl_Parent
                    Public PartNo As String
                    Public Nomenclature As String
                End Class
            End Class
        End Class
        Public Class oPart
            Function GetPartDocument() As PartDocument
                oCATIA = GetCATIA()
                Dim MyPartDocument As PartDocument

                On Error Resume Next
                MyPartDocument = oCATIA.ActiveDocument
                If MyPartDocument Is Nothing Or Err.Number <> 0 Then
                    MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a Product" & vbCrLf & "in the Active session", vbCritical, "Open a Product")
                    Environment.Exit(0)
                End If

                GetPartDocument = MyPartDocument
            End Function
            Public Class PartMetaData
                Dim PartNo, Nomenclature, Description, Parent, JobNo, RT As String
            End Class
            Function Select3DPart() As oPart

                Dim CATIA As Object, ActivePartDocument As ProductDocument, ActivePart As oPart

                CATIA = GetCATIA()
                ActivePartDocument = GetPartDocument()
                ActivePart = ActivePartDocument.Part

                Dim What(0) As Object
                What(0) = "Part"

                Dim UserSel 'As SelectedElement
                UserSel = CATIA.ActiveDocument.Selection
                UserSel.Clear()

                Dim e As String
                e = UserSel.SelectElement3(What, "Select a Part", False, 2, False)

                Dim SelectedElement As Long, ActiveProduct

                ActivePart = UserSel.Item(1).Value
                UserSel.Clear()

                Select3DPart = ActivePart
            End Function

            'Function SelectAxis() As Axis

            'End Function

            'Function SelectMatingFace() As Face

            'End Function

        End Class
    End Class
    Public Class Drawing

        Public Class cl_PartsList

            Private _PartNo As String
            Public Property PartNo() As String
                Get
                    Return _PartNo
                End Get
                Set(ByVal value As String)
                    _PartNo = value
                End Set
            End Property

            Private _Quantity As String
            Public Property Quantity() As String
                Get
                    Return _Quantity
                End Get
                Set(ByVal value As String)
                    _Quantity = value
                End Set
            End Property

            Private _Nomenclature As String
            Public Property Nomenclature() As String
                Get
                    Return _Nomenclature
                End Get
                Set(ByVal value As String)
                    _Nomenclature = value
                End Set
            End Property

            Private _ItemNo As String
            Public Property ItemNo() As String
                Get
                    Return _ItemNo
                End Get
                Set(ByVal value As String)
                    _ItemNo = value
                End Set
            End Property

            Private _Material As String
            Public Property Material() As String
                Get
                    Return _Material
                End Get
                Set(ByVal value As String)
                    _Material = value
                End Set
            End Property
            Private _ParentDashNo As String
            Public Property ParentDashNo() As String
                Get
                    Return _ParentDashNo
                End Get
                Set(ByVal value As String)
                    _ParentDashNo = value
                End Set
            End Property
            Private _ParentNomenclature As String
            Public Property ParentNomenclature() As String
                Get
                    Return _ParentNomenclature
                End Get
                Set(ByVal value As String)
                    _ParentNomenclature = value
                End Set
            End Property
            Private _DrawingNo As String
            Public Property DrawingNo() As String
                Get
                    Return _DrawingNo
                End Get
                Set(ByVal value As String)
                    _DrawingNo = value
                End Set
            End Property
            Private _ParentPartNo As String
            Public Property ParentPartNo() As String
                Get
                    Return DrawingNo + ParentDashNo
                End Get
                Set(ByVal value As String)
                    _ParentPartNo = value
                End Set
            End Property
            Private _DrawingName As String
            Public Property DrawingName() As String
                Get
                    Return _DrawingName
                End Get
                Set(ByVal value As String)
                    _DrawingName = value
                End Set
            End Property
            Public Class cl_Parent
                Public PartNo As String
                Public Nomenclature As String
            End Class

            Public ParentsDashNos As New List(Of String)

        End Class
        ''' <summary>
        '''
        ''' </summary>
        ''' <returns></returns>
        Public Function PartsList() As List(Of cl_PartsList)
            Dim cl_PL As New cl_PartsList, oPartsList As New List(Of cl_PartsList), row As Integer, column As Integer, item As Integer, newitem As Integer
            Dim tempParentDasNosList As New List(Of String), tempQtyList As New List(Of String)
            Dim cellValue1 As String, cellValue As String

            Dim Active2DTablesList As List(Of DrawingTable)
            Dim Active2DTable As DrawingTable 'One-based index where cell (1,1) is at the top left of the table

            Active2DTablesList = Select2DTable()




            'Active2DTable = Active2DTablesList(0)



            Dim ParentsDashNos As New List(Of String)

            For column = 1 To Active2DTablesList(0).NumberOfColumns

                cellValue = Active2DTablesList(0).GetCellString(Active2DTablesList(0).NumberOfRows - 1, column)

                If IsItAValidParentDashNo(cellValue.ToString) = True Then
                    ParentsDashNos.Add(Trim(cellValue.ToString))
                End If
            Next column


            For Each Active2DTable In Active2DTablesList
                For row = 1 To Active2DTable.NumberOfRows

                    If IsNumeric(Trim(Active2DTable.GetCellString(row, Active2DTable.NumberOfColumns))) Then

                        cl_PL = New cl_PartsList
                        tempParentDasNosList = New List(Of String)
                        tempQtyList = New List(Of String)
                        cl_PL.DrawingNo = "B471356"
                        cl_PL.PartNo = Active2DTable.GetCellString(row, Active2DTable.NumberOfColumns - 3).ToString
                        cl_PL.ItemNo = Active2DTable.GetCellString(row, Active2DTable.NumberOfColumns)
                        cl_PL.Material = Active2DTable.GetCellString(row, Active2DTable.NumberOfColumns - 1).ToString
                        cl_PL.Nomenclature = Active2DTable.GetCellString(row, Active2DTable.NumberOfColumns - 2)

                        For item = 0 To ParentsDashNos.Count - 1

                            cellValue1 = Active2DTable.GetCellString(row, item + 1)

                            If (cellValue1) <> vbNullString Then
                                tempQtyList.Add(cellValue1)
                                tempParentDasNosList.Add(ParentsDashNos(item))
                            End If

                        Next item

                        For newitem = 0 To tempParentDasNosList.Count - 1
                            cl_PL.Quantity = tempQtyList(newitem)
                            cl_PL.ParentDashNo = tempParentDasNosList(newitem)
                            cl_PL.PartNo = Active2DTable.GetCellString(row, Active2DTable.NumberOfColumns - 3).ToString
                            cl_PL.ItemNo = Active2DTable.GetCellString(row, Active2DTable.NumberOfColumns)
                            cl_PL.Material = Active2DTable.GetCellString(row, Active2DTable.NumberOfColumns - 1).ToString
                            cl_PL.Nomenclature = Active2DTable.GetCellString(row, Active2DTable.NumberOfColumns - 2)
                            cl_PL.DrawingNo = "B471356"

                            oPartsList.Add(cl_PL)

                            cl_PL = New cl_PartsList
                        Next newitem


                    End If
                Next row


            Next
            cl_PL = Nothing
            Return oPartsList
        End Function
        Function IsItAValidParentDashNo(ParentDashNo As Object) As Boolean
            ParentDashNo = Trim(ParentDashNo)

            If Len(ParentDashNo) = 4 Then
                If Left(ParentDashNo, 1) = "-" Then
                    If IsNumeric(Mid(ParentDashNo, 2, 3)) = True Then
                        If Mid(ParentDashNo, 2, 1) = "5" Or Mid(ParentDashNo, 2, 1) = "6" Or Mid(ParentDashNo, 2, 1) = "7" Then
                            Return True
                            Exit Function
                        End If
                    End If
                End If
            End If
            Return False
        End Function
        Function IsItAValidFirstDwgTable() As Boolean
            Return False
        End Function
        Public Function GetDrawingDocument() As DrawingDocument
            oCATIA = GetCATIA()
            Dim MyDrawingDocument As DrawingDocument

            On Error Resume Next
            MyDrawingDocument = oCATIA.ActiveDocument
            If MyDrawingDocument Is Nothing Or Err.Number <> 0 Then
                MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a Drawing" & vbCrLf & "In the Active session", vbCritical, "Open a Drawing")
                Err.Clear()
                Environment.Exit(0)
            End If
            GetDrawingDocument = MyDrawingDocument
        End Function
        Public Function IsADrawingDocumentOpen() As Boolean


            oCATIA = GetCATIA()
            Dim MyDrawingDocument As DrawingDocument

            On Error Resume Next
            MyDrawingDocument = oCATIA.ActiveDocument
            If MyDrawingDocument Is Nothing Or Err.Number <> 0 Then
                MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a Drawing" & vbCrLf & "In the Active session", vbCritical, "Open a Drawing")
                Err.Clear()
                Return False
                Exit Function
            End If
            Return True


        End Function
        Public Function Select2DTable() As List(Of DrawingTable)
            Dim ActiveDrawingDocument As DrawingDocument, ActiveTablesList As New List(Of DrawingTable), SelectedTable As Selection, e As String
            Dim What(0)

            oCATIA = GetCATIA()
            ActiveDrawingDocument = GetDrawingDocument()

            SelectedTable = ActiveDrawingDocument.Selection
            SelectedTable.Clear()

            What(0) = "DrawingTable"
            e = SelectedTable.SelectElement3(What, "Select a DrawingTable", True, CATMultiSelectionMode.CATMultiSelTriggWhenUserValidatesSelection, False)

            Dim table As Integer
            For table = 1 To SelectedTable.Count
                ActiveTablesList.Add(SelectedTable.Item(table).Value)
            Next table

            SelectedTable.Clear()

            Select2DTable = ActiveTablesList

        End Function
        Sub Clean2DTable()
            Dim Table2D As DrawingTable
            Table2D = Select2DTable()

            Dim SplittedPartNo() As String, i As Integer

            For i = 1 To Table2D.NumberOfRows
                If InStr(UCase(Table2D.GetCellString(i, 4)), " MULT") > 0 Then
                    SplittedPartNo = Split(UCase(Table2D.GetCellString(i, 4)), " MULT")
                    Table2D.SetCellString(i, 4, SplittedPartNo(0))
                End If
            Next i
        End Sub
        Sub DrwFlagNote(ByRef oDrawingText As DrawingText)

            Dim SplitoDrwTxt() As String
            Dim Position As Integer
            Dim Length As Integer

            Position = 1

            SplitoDrwTxt = Split(oDrawingText.Text, " ")

            For Each FlagNote In SplitoDrwTxt
                Length = Len(FlagNote)
                If IsNumeric(Trim(FlagNote)) = True Then
                    oDrawingText.SetParameterOnSubString(catBorder, Position, Length, 6)
                End If

                Position = Position + Length + 1
            Next

            'oDrawingText.SetParameterOnSubString catBorder, 1, 2, 6

        End Sub
        Function GetNotes(ByRef oDrawingText As DrawingText) As Collection

            'Dim FlagNotes() As String
            'Dim FlagnotesCollection As New Collection

            'Dim Notes() As String
            'Dim NotesCollection As New Collection

            'Dim k As Integer
            'Dim j As Integer

            'Dim Delimiter As String
            'Delimiter = Chr(10) '+ Chr(32)

            'Dim StartNote As New Collection
            'Dim StartFlagnote As New Collection
            'Dim GeneralNote As New Collection

            'Dim TempNote As String
            'Dim TempFlagnote As String

            'Notes = Split(oDrawingText.Text, Delimiter)

            'For j = LBound(Notes) To UBound(Notes)
            '    If Trim(Notes(j)) = vbNullString Then
            '        StartNote.Add(j)
            '    End If
            'Next j
            'StartNote.Add(j)

            'TempNote = vbNullString

            'For k = 2 To StartNote.count
            '    For j = StartNote(k - 1) To StartNote(k) - 1
            '        If Trim(Notes(j)) <> "" Then
            '            TempNote = TempNote + Space(1) + Trim(Notes(j))
            '        End If
            '    Next

            '    If Trim(TempNote) <> "" Then
            '        NotesCollection.Add(TempNote)
            '    End If
            '    TempNote = vbNullString
            'Next k

            ''Last value

            ''For Each Note In NotesCollection
            ''Cells(i, 26) = Note
            ''i = i + 1
            ''Next

            'GetNotes = NotesCollection
        End Function
        Function GetPartAndDescFromCallout(Callout As String, Table2D As DrawingTable) As String
            'Dim PartAndDescFromCallout As New Collection
            'Set GetPartAndDescFromCallout = PartAndDescFromCallout
            Dim SplittedPartNo
            For i = 1 To Table2D.NumberOfRows

                If Trim(UCase(Table2D.GetCellString(i, 4))) = Trim(UCase(Callout)) Or Trim(UCase(Table2D.GetCellString(i, 1))) = Trim(UCase(Callout)) Then
                    GetPartAndDescFromCallout = UCase(Table2D.GetCellString(i, 4))
                End If

                If InStr(Trim(UCase(Callout)), " MULT") > 0 Then
                    SplittedPartNo = Split(UCase(Callout), " MULT")

                    If Trim(UCase(Table2D.GetCellString(i, 4))) = Trim(UCase(SplittedPartNo(0))) Then
                        GetPartAndDescFromCallout = UCase(Table2D.GetCellString(i, 4))
                    End If
                End If
            Next i
        End Function
        Sub AllViewsAutomaticBalloonCallouts() '(Table As ListObject)

            Dim CATIA As Object
            CATIA = GetObject(, "CATIA.Application")

            Dim oDrawingDocuments As Documents
            oDrawingDocuments = CATIA.Documents

            Dim oDrawingDocument As Document
            oDrawingDocument = CATIA.ActiveDocument

            Dim oDrawingSheets As DrawingSheets
            oDrawingSheets = oDrawingDocument.Sheets

            Dim oDrawingSheet As DrawingSheet
            oDrawingSheet = oDrawingSheets.ActiveSheet

            Dim oDrawingViews As DrawingViews
            oDrawingViews = oDrawingSheet.Views

            Dim oDrawingView As DrawingView
            oDrawingView = oDrawingViews.ActiveView

            Dim SelectedView As Selection
            SelectedView = CATIA.ActiveDocument.Selection

            Dim oDrawingtables As DrawingTables
            oDrawingtables = oDrawingView.Tables

            Dim oDrawingtable As DrawingTable
            'Set oDrawingtable = oDrawingtables.Tables

            Dim Table2D As DrawingTable
            Table2D = Select2DTable()

            For Each oDrawingSheet In oDrawingSheets
                If InStr(oDrawingSheet.Name, "STARTUP") = 0 And InStr(oDrawingSheet.Name, "General_Tolerances") = 0 Then

                    For Each oDrawingView In oDrawingSheet.Views

                        Select Case oDrawingView.ViewType

                            Case 12 To 15
                            Case 0
                                'Case 9

                            Case Else

                                If InStr(oDrawingView.Name, "EC_") = 0 And oDrawingView.Name <> "Gen_Tol_A0&J2" And InStr(UCase(oDrawingView.Name), "NOTE") = 0 And oDrawingView.Texts.Count > 0 Then
                                    Dim i As Integer

                                    For Each Callout In oDrawingView.Texts
                                        Dim SplittedPartNo() As String

                                        If Callout.FrameType = 53 Then 'Balloon Frame
                                            For i = 1 To Table2D.NumberOfRows
                                                If Trim(UCase(Table2D.GetCellString(i, 4))) = Trim(UCase(Callout.Text)) Then
                                                    Callout.Text = UCase(Table2D.GetCellString(i, 1))
                                                End If

                                                If InStr(Trim(UCase(Callout.Text)), " MULT") > 0 Then
                                                    SplittedPartNo = Split(UCase(Callout.Text), " MULT")

                                                    If Trim(UCase(Table2D.GetCellString(i, 4))) = Trim(UCase(SplittedPartNo(0))) Then
                                                        Callout.Text = UCase(Table2D.GetCellString(i, 1))
                                                    End If

                                                End If
                                            Next i
                                        End If
                                    Next Callout
                                End If

                        End Select

                    Next oDrawingView
                End If
            Next oDrawingSheet

        End Sub

        'Sub AddRevisionBalloon(TableNotes As ListObject)

        '    Dim CATIA As Object
        '    CATIA = GetObject(, "CATIA.Application")

        '    Dim oDrawingDocuments As Documents
        '    oDrawingDocuments = CATIA.Documents

        '    Dim oDrawingDocument As DrawingDocument
        '    oDrawingDocument = CATIA.ActiveDocument

        '    Dim oDrawingSheets As DrawingSheets
        '    oDrawingSheets = oDrawingDocument.Sheets

        '    Dim oDrawingSheet As DrawingSheet
        '    oDrawingSheet = oDrawingSheets.ActiveSheet

        '    Dim oDrawingViews As DrawingViews
        '    oDrawingViews = oDrawingSheet.Views

        '    oDrawingViews.Add("Notes")

        '    Dim oDrawingView As DrawingView
        '    oDrawingView = oDrawingViews.Item("Notes")

        '    Dim oDrawingText As DrawingText
        '    Dim Y As Integer, NoteCntr As Integer, Note As Range

        '    NoteCntr = 1

        '    oDrawingText = oDrawingView.Texts.Add(" GENERAL NOTES", 30, 530 - Y) 'First line of Text ' NOTES: UNLESS OTHERWISE SPECIFIED
        '    'oDrawingText.TextProperties.Justification = catLeft

        '    oDrawingText.WrappingWidth = 500
        '    oDrawingText.SetParameterOnSubString(catFontSize, 1, 13, 6350) 'oDrawingText.SetParameterOnSubString catBold, 1, 6, 1
        '    Y = Y + 15

        '    'For Each Note In TableNotes.ListColumns(3).DataBodyRange
        '    '    Dim SelectedNote As Range
        '    '    Set SelectedNote = Note.Offset(0, 4)
        '    '
        '    ''    If SelectedNote.Value2 = True Then
        '    ''
        '    '        Set oDrawingText = oDrawingView.Texts.Add(NoteCntr & ". " & Note.Value, 30, 530 - Y)
        '    '        oDrawingText.TextProperties.Justification = catLeft
        '    '        oDrawingText.WrappingWidth = 500
        '    '
        '    '        If Note.Offset(0, -1).Value <> "" Then
        '    '            oDrawingText.Name = "Note_ID_" & Note.Offset(0, 3).Value
        '    '        End If
        '    '
        '    '        'If Note.Offset(0, 2).Value = "Yes" Then Call DrwFlagNote(oDrawingText)
        '    '
        '    '        Y = Y + 15
        '    '        NoteCntr = NoteCntr + 1
        '    ''    End If
        '    'Next
        'End Sub

        Function DrawingZone() As String
            Dim CATIA As Object
            CATIA = GetObject(, "CATIA.Application")

            Dim ActiveDrawingDocument As DrawingDocument
            'Set ActiveDwgDocument = CATIA.ActiveDocument

            On Error Resume Next
            'Set oDrawingDocument = CATIA.ActiveDocument
            ActiveDrawingDocument = CATIA.ActiveDocument
            If Err.Number <> 0 Then
                MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a Drawing" & vbCrLf & "In the Active session", vbCritical, "Open a Drawing")
                Exit Function
            End If

            Dim What(2) 'As String
            What(0) = "DrawingTable"
            What(1) = "DrawingView"
            What(2) = "DrawingText"

            Dim UserSel As SelectedElement
            UserSel = CATIA.ActiveDocument.Selection
            UserSel.Clear()

            Dim e 'As String
            e = UserSel.SelectElement3(What, "Select a Drawing Element", True, 2, False)

            Dim DrawingObject As Object
            DrawingObject = UserSel.Item(1).Value

            'DrawingObject.X = 0 + 28
            'DrawingObject.Y = 0 + 156
            UserSel.Clear()

            Dim DrawingZoneLetter(6, 2) As String
            Dim DrawingZoneNumber(8, 2) As String
            'DrawingZoneLetter
            'DrawingZoneLetter
            'GetDrawingZone = DrawingObject.X & vbCrLf & DrawingObject.Y
            DrawingZone = DrawingObject.X & ", " & DrawingObject.Y
        End Function

        Sub BOMFromViews()
            Dim CATIA As Object
            CATIA = GetObject(, "CATIA.Application")

            Dim oDrawingDocuments As Documents
            oDrawingDocuments = CATIA.Documents

            Dim oDrawingDocument As DrawingDocument
            oDrawingDocument = CATIA.ActiveDocument

            Dim oDrawingSheets As DrawingSheets
            oDrawingSheets = oDrawingDocument.Sheets

            Dim oDrawingSheet As DrawingSheet
            oDrawingSheet = oDrawingSheets.ActiveSheet

            Dim oDrawingViews As DrawingViews
            oDrawingViews = oDrawingSheet.Views

            Dim oDrawingView As DrawingView
            oDrawingView = oDrawingViews.ActiveView

            Dim SelectedView As Selection
            SelectedView = CATIA.ActiveDocument.Selection

            Dim oDrawingtables As DrawingTables
            oDrawingtables = oDrawingView.Tables

            Dim oDrawingtable As DrawingTable
            'Set oDrawingtable = oDrawingtables.Tables

            Dim CalloutsDict
            CalloutsDict = CreateObject("Scripting.Dictionary")

            Dim Table2D As DrawingTable
            Table2D = Select2DTable()

            For Each oDrawingSheet In oDrawingSheets

                For Each oDrawingView In oDrawingSheet.Views

                    Select Case oDrawingView.ViewType

                        Case 12 To 15
                        Case 0
                            'Case 9

                        Case Else

                            If InStr(oDrawingView.Name, "EC_") = 0 And oDrawingView.Name <> "Gen_Tol_A0&J2" And InStr(UCase(oDrawingView.Name), "NOTE") = 0 And oDrawingView.Texts.Count > 0 Then
                                Dim i As Integer

                                '                        Dim SplittedPartNo() As String
                                '
                                Dim Callout
                                Dim Qty As Integer
                                For Each Callout In oDrawingView.Texts
                                    If Callout.FrameType = 53 And IsNumeric(Trim(Callout.Text)) = True Then  'Balloon Frame

                                        If Not CalloutsDict.exists(Callout.Text) Then
                                            Qty = 1
                                            CalloutsDict.Add(Key:=Callout.Text, Item:=Qty)

                                        ElseIf CalloutsDict.exists(Callout.Text) Then
                                            CalloutsDict.Item(Callout.Text) = CalloutsDict.Item(Callout.Text) + 1
                                        End If

                                        '                            For i = 1 To Table2D.NumberOfRows
                                        '                                If Trim(UCase(Table2D.GetCellString(i, 4))) = Trim(UCase(Callout.Text)) Then
                                        '                                    Callout.Text = UCase(Table2D.GetCellString(i, 1))
                                        '                                End If
                                        '
                                        '                                If InStr(Trim(UCase(Callout.Text)), " MULT") > 0 Then
                                        '                                    SplittedPartNo = Split(UCase(Callout.Text), " MULT")
                                        '
                                        '                                     If Trim(UCase(Table2D.GetCellString(i, 4))) = Trim(UCase(SplittedPartNo(0))) Then
                                        '                                        Callout.Text = UCase(Table2D.GetCellString(i, 1))
                                        '                                     End If
                                        '
                                        '                                End If
                                        '                            Next i
                                    End If
                                Next Callout

                            End If

                    End Select

                Next oDrawingView
            Next oDrawingSheet

            'i = 0

            'Dim XLTable3D As Range
            'XLTable3D = ActiveSheet.ListObjects("Table3D").DataBodyRange

            'On Error Resume Next
            'XLTable3D.Delete()
            'On Error GoTo 0

            'For Each Key In CalloutsDict
            '    Dim ItemNo As String, QtyReqd As String, PartNo As String, Nomenclature As String
            '    QtyReqd = CalloutsDict(Key)
            '    ItemNo = Key
            '    PartNo = D2.GetPartAndDescFromCallout(ItemNo, Table2D)
            '    ActiveSheet.Cells(i + 14, 1).Value = QtyReqd
            '    'ActiveSheet.Cells(i + 14, 2).Value = ItemNo
            '    ActiveSheet.Cells(i + 14, 2).Value = PartNo
            '    i = i + 1
            'Next
            'Call Export2DTable(Table2D, ActiveSheet.ListObjects("Table2D"))
        End Sub

        '        Sub Export2DTable(Table2D As DrawingTable, XLTable As ListObject)
        '            ActiveSheet.Cells(10, 8).Value = Null
        '            ActiveSheet.Cells(11, 7).Value = Null

        '            Dim CATIA As Object, oDrawingDocument As DrawingDocument
        '            CATIA = GetCATIA()
        '            oDrawingDocument = GetCATIADrawingDocument

        '            Dim oRow As Range
        '            Dim i As Integer
        '            i = 1

        '            On Error Resume Next
        '            XLTable.DataBodyRange.Delete()
        '            XLTable.DataBodyRange.ClearContents()
        '            XLTable.DataBodyRange.ClearFormats()
        '            XLTable.DataBodyRange.WrapText = True
        '            On Error GoTo 0

        '            Application.ScreenUpdating = False

        '            Dim PartNumber As String, Qty As String, Nomenclature As String, ItemNo As String, Material As String

        '            For i = 1 To Table2D.NumberOfRows Step 1

        '                XLTable.ListRows.Add()
        '                Dim TotalRows As Integer

        '                TotalRows = XLTable.ListRows.Count

        '                PartNumber = Table2D.GetCellString(i, 4)
        '                Qty = Table2D.GetCellString(i, 3)
        '                Nomenclature = Table2D.GetCellString(i, 2)
        '                ItemNo = Table2D.GetCellString(i, 1)

        '                On Error Resume Next
        '                Material = Table2D.GetCellString(i, 5)
        '                On Error GoTo 0

        '                'XLTable.DataBodyRange.Rows.NumberFormat = "@"
        '                XLTable.DataBodyRange.Rows(TotalRows).Columns(1).Value = Qty

        '                XLTable.DataBodyRange.Rows(TotalRows).Columns(2).NumberFormat = "@"
        '                XLTable.DataBodyRange.Rows(TotalRows).Columns(2).Value = PartNumber

        '                XLTable.DataBodyRange.Rows(TotalRows).Columns(3).NumberFormat = "@"
        '                XLTable.DataBodyRange.Rows(TotalRows).Columns(3).Value2 = Nomenclature

        '                On Error Resume Next
        '                XLTable.DataBodyRange.Rows(TotalRows).Columns(4).NumberFormat = "@"
        '                XLTable.DataBodyRange.Rows(TotalRows).Columns(4).Value = Material
        '                On Error GoTo 0

        '                XLTable.DataBodyRange.Rows(TotalRows).Columns(5).Value = ItemNo

        '                '    i = i + 1

        '            Next

        '            On Error Resume Next
        '            XLTable.DataBodyRange.ClearFormats()
        '            XLTable.DataBodyRange.WrapText = True
        '            On Error GoTo 0
        '            'ActiveSheet.Cells(10, 8).Value = oDrawingDocument.Parameters.Item("DRAWING_NUMBER").Value
        '            'ActiveSheet.Cells(11, 7).Value = oDrawingDocument.Parameters.Item("DRAWING_TITLE").Value
        '            'ActiveSheet.Cells(12, 8).Value = Date & " " & Time()
        'ActiveSheet.Shapes.Range(Array("TimeStamp2D")).TextFrame2.TextRange.Characters.Text = Date & " " & Time()

        '            ActiveSheet.Shapes.Range(Array("2DPartNo")).TextFrame2.TextRange.Characters.Text = oDrawingDocument.Parameters.Item("DRAWING_NUMBER").Value
        '            ActiveSheet.Shapes.Range(Array("2DDescription")).TextFrame2.TextRange.Characters.Text = oDrawingDocument.Parameters.Item("DRAWING_TITLE").Value
        '        End Sub


        Public Sub ExportToDrawing(otherlist As List(Of Object))
            Dim ActiveDrawingDocument As DrawingDocument, NumberOfRows As Integer, NumberOfColumns As Integer

            oCATIA = GetCATIA()
            ActiveDrawingDocument = GetDrawingDocument()

            Dim oDrwTables As DrawingTables = ActiveDrawingDocument.Sheets.ActiveSheet.Views.ActiveView.Tables          'Dim oDrwSheets As DrawingSheets = ActiveDrawingDocument.Sheets, oDrwSheet As DrawingSheet = oDrwSheets.ActiveSheet, oDrwView As DrawingView = oDrwSheet.Views.ActiveView
            Dim oDrwTable As DrawingTable

            NumberOfRows = otherlist.Count
            NumberOfColumns = 6
            oDrwTable = oDrwTables.Add(896.650497436523, 126.999582529068, NumberOfRows, NumberOfColumns, 13.094, 20)  ' double  iPositionX,double  iPositionY, long  iNumberOfRow, long  iNumberOfColumn, double  iRowHeight, double  iColumnWidth)

            oDrwTable.Name = "Parts List"
            oDrwTable.MergeCells(NumberOfRows, NumberOfColumns - 4, 1, 5)
            'Set the column sizes
            oDrwTable.SetColumnSize(1, 14.144)
            oDrwTable.SetColumnSize(2, 14.144)
            oDrwTable.SetColumnSize(3, 50.576)
            oDrwTable.SetColumnSize(4, 99.153)
            oDrwTable.SetColumnSize(5, 41.468)
            oDrwTable.SetColumnSize(6, 17.18)

            oDrwTable.AnchorPoint = CatTableBottomLeft

            oDrwTable.SetCellString(NumberOfRows, NumberOfColumns - 4, "PARTS LIST")
            oDrwTable.SetCellAlignment(NumberOfRows, NumberOfColumns - 4, CatTableMiddleCenter)

            oDrwTable.SetCellString(NumberOfRows - 1, NumberOfColumns, "ITEM" & vbLf & "NO.")
            oDrwTable.SetCellAlignment(NumberOfRows - 1, NumberOfColumns, CatTableMiddleCenter)

            oDrwTable.SetCellString(NumberOfRows - 1, NumberOfColumns - 1, "MATERIAL" & vbLf & "SPECIFICATION")
            oDrwTable.SetCellAlignment(NumberOfRows - 1, NumberOfColumns - 1, CatTableMiddleCenter)

            oDrwTable.SetCellString(NumberOfRows - 1, NumberOfColumns - 2, "NOMENCLATURE" & vbLf & "OR DESCRIPTION")
            oDrwTable.SetCellAlignment(NumberOfRows - 1, NumberOfColumns - 2, CatTableMiddleCenter)

            oDrwTable.SetCellString(NumberOfRows - 1, NumberOfColumns - 3, "PART OR" & vbLf & "IDENTIFYIONG NO.")
            oDrwTable.SetCellAlignment(NumberOfRows - 1, NumberOfColumns - 3, CatTableMiddleCenter)

            oDrwTable.SetCellString(NumberOfRows - 1, NumberOfColumns - 4, "CAGE" & vbLf & "CODE")
            oDrwTable.SetCellAlignment(NumberOfRows - 1, NumberOfColumns - 4, CatTableMiddleCenter)
        End Sub


    End Class
    Public Class UDF
        Public Class Panel

        End Class
        Public Class Drawer

        End Class

        Public Class BondedStructure
        End Class

        Public Class DecoPanel

        End Class
        Public Class Monument

        End Class

        Public Class Hardware
            Public Class Fastener

            End Class

            Public Class Insert

            End Class

            Public Class Washer

            End Class
        End Class

    End Class
End Class