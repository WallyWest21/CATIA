Imports INFITF
Imports MECMOD
Imports ProductStructureTypeLib
Imports INFITF.CATMultiSelectionMode
Imports DRAFTINGITF
Imports DRAFTINGITF.CatTextProperty

Public Class cl_VB_CATIALib
    Shared oCATIA As INFITF.Application 'http://www.coe.org/p/fo/et/thread=22850
    Shared Function GetCATIA() As INFITF.Application
        oCATIA = GetObject(, "CATIA.Application")
        If oCATIA Is Nothing Or Err.Number <> 0 Then
            MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a CATIA session", vbCritical, "Open a CATIA Session ")
            Environment.Exit(0)
            '       Set CATIA = CreateObject("CATIA.Application")
            '       CATIA.Visible = True
        End If

        GetCATIA = oCATIA
    End Function
    Public Class _3D
        Public Class Product
            Public Sub test()
                MsgBox("hi")
            End Sub
            Public Part As String
            Function GetProductDocument() As ProductDocument
                oCATIA = GetCATIA()
                Dim MyProductDocument As ProductDocument

                On Error Resume Next
                MyProductDocument = oCATIA.ActiveDocument
                If MyProductDocument Is Nothing Or Err.Number <> 0 Then
                    ' MsgBox "No CATIA Active Document found "
                    MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a Product" & vbCrLf & "in the Active session", vbCritical, "Open a Product")
                    Environment.Exit(0)
                End If
                GetProductDocument = MyProductDocument
            End Function
            Public Function SelectSingle3DProduct() As String

                Dim ActiveProductDocument As ProductDocument, ActiveProduct As Products

                ActiveProductDocument = GetProductDocument()

                Dim What(0) 'As Object
                What(0) = "Product"

                Dim SelectedProduct As SelectedElement
                'SelectedProduct = ActiveProductDocument.Selection
                'SelectedProduct.Clear()

                Dim e 'As String
                e = SelectedProduct.SelectElement3(What, "Select a Product or a Component", False, CATMultiSelTriggWhenUserValidatesSelection, False)


                ActiveProduct = SelectedProduct.Item(1).Value.partnumber
                SelectedProduct.Clear()

                Return ActiveProduct
            End Function
            '            Public Function PartsList(ActiveProducts As List(Of Products)) As List(Of cl_PartsList)
            '                Dim testobject As String
            '                Dim cl_PL As New cl_PartsList
            '                Dim ActiveProduct As Products
            '                Dim counter As Integer
            '                For Each ActiveProduct In ActiveProducts

            '                    If ActiveProduct.Count = 0 Then
            '                        Exit Function
            '                    End If

            '                    For counter = 1 To ActiveProduct.Count
            '                        cl_PL.PartNo = ActiveProduct.Partnumber
            '                        cl_PL.Nomenclature = ActiveProduct.Nomenclature

            '                    Next

            '                    Dim oInstances As Products

            '                    oInstances = oInProduct.Products

            '                    '-----No instances found then this is CATPart


            '                    If oInstances.Count = 0 Then

            '                        'MsgBox "This is a CATPart with part number " & oInProduct.PartNumber


            '                        Exit Function

            '                    End If


            '                    '-----Found an instance therefore it is a CATProduct

            '                    'MsgBox "This is a CATProduct with part number " & oInProduct.ReferenceProduct.PartNumber


            '                    Dim k As Integer

            '                    For k = 1 To oInstances.Count

            '                        Dim oInst As Product

            '   Set oInst = oInstances.Item(k)


            ''apply design mode

            'oInstances.Item(k).ApplyWorkMode DESIGN_MODE

            '   Call WalkDownTree(oInst)

            '                    Next


            '                    End Sub
            '                Next


            '                'Try
            '                '    Parallel.For(1, ActiveProduct.Count + 1, Sub(k)

            '                '                                                 Dim oInst As INFITF.AnyObject
            '                '                                                 oInst = ActiveProduct.Item(k)
            '                '                                                 'ActiveProduct.Item(k).ApplyWorkMode(DESIGN_MODE)   'apply design mode
            '                '                                                 'testobject = oInst.partnumber
            '                '                                                 'If Validation.IsComponent(oInst) = False And oInstances.Item(k).Parent.Parent.PartNumber = oInProduct.partnumber Then
            '                '                                                 'If Validation.IsComponent(oInst) = False And RealParent(oInst) = oInProduct.partnumber Then

            '                '                                                 cl_PL.PartNo = oInst.Partnumber
            '                '                                                 cl_PL.Nomenclature = oInst.Nomenclature

            '                '                                                 MsgBox(cl_PL.Nomenclature)

            '                '                                                 PartsList.Add(cl_PL)


            '                '                                                 'Realchildren3D.Add(oInst.partnumber)
            '                '                                                 '    comp.Add(oInst.partnumber)

            '                '                                                 'End If

            '                '                                                 'If Validation.IsComponent(oInst) = True And RealParent(oInst) = oInProduct.partnumber Then
            '                '                                                 '    Call WalkDownTree(oInst)
            '                '                                                 '    test = RealParent(oInst)

            '                '                                                 'End If

            '                '                                             End Sub)

            '                'Catch ex As Exception
            '                '    MsgBox("You need a multicore computer")
            '                'End Try

            '                For Each Child In ActiveProduct
            '                    cl_PL.PartNo = Child.Partnumber
            '                    cl_PL.Nomenclature = Child.Nomenclature
            '                    'cl_PL.cl_Parent.

            '                    MsgBox(cl_PL.Nomenclature)

            '                    PartsList.Add(cl_PL)
            '                Next


            '                Return PartsList(ActiveProduct)
            '            End Function
            'Sub WalkDownTree(ByVal oInProduct As Object)
            '    'As Product)

            '    Dim Validation As New Validation
            '    Dim test As String
            '    Dim testobject As String
            '    Dim oInstances As Products
            '    oInstances = oInProduct.Products

            '    '-----No instances found then this is CATPart

            '    If oInstances.Count = 0 Then

            '        Exit Sub
            '    End If


            '    Try
            '        Parallel.For(1, oInstances.Count + 1, Sub(k)

            '                                                  Dim oInst As INFITF.AnyObject
            '                                                  oInst = oInstances.Item(k)
            '                                                  oInstances.Item(k).ApplyWorkMode(DESIGN_MODE)   'apply design mode
            '                                                  testobject = oInst.partnumber
            '                                                  'If Validation.IsComponent(oInst) = False And oInstances.Item(k).Parent.Parent.PartNumber = oInProduct.partnumber Then
            '                                                  If Validation.IsComponent(oInst) = False And RealParent(oInst) = oInProduct.partnumber Then

            '                                                      Children3D.Add(oInst)
            '                                                      Realchildren3D.Add(oInst.partnumber)
            '                                                      comp.Add(oInst.partnumber)

            '                                                  End If

            '                                                  If Validation.IsComponent(oInst) = True And RealParent(oInst) = oInProduct.partnumber Then
            '                                                      Call WalkDownTree(oInst)
            '                                                      test = RealParent(oInst)

            '                                                  End If

            '                                              End Sub)

            '    Catch ex As Exception
            '        MsgBox("You need a multicore computer")
            '    End Try
            '    'Realchildren3D.Add("klhkjhklhjkhkjlhkljl")


            '    '    lst1.Add("New Item")

            '    'ListBox1.ItemsSource = ChildrenList.
            '    comp.Add("comparator")


            'End Sub

            'Sub Select3DProduct()

            '    Dim CATIA As Object, ActiveProductDocument As ProductDocument, ActProd As Products

            '    CATIA = GetCATIA()
            '    ActiveProductDocument = GetProductDocument()


            '    Dim What(0) As Object
            '    What(0) = "Product"

            '    Dim UserSel As Object
            '    UserSel = CATIA.ActiveDocument.Selection
            '    UserSel.Clear()

            '    Dim e As String
            '    e = UserSel.SelectElement3(What, "Select a Product or a Component", False, 2, False)

            '    Dim SelectedElement As Long, ActiveProduct

            '    ActiveProduct = UserSel.Item(1).Value
            '    UserSel.Clear()

            '    ActiveProduct.ApplyWorkMode(DESIGN_MODE)

            '    Dim DictStrChildren As Dictionary
            '    DictStrChildren = CreateObject("Scripting.Dictionary")

            '    Dim DictChildren As Dictionary
            '    DictChildren = CreateObject("Scripting.Dictionary")

            '    Dim ChildrenName As String

            '    For i = 1 To ActiveProduct.Products.count
            '        On Error Resume Next

            '        ChildrenName = ActiveProduct.Products.Item(i).PartNumber


            '        If ActiveProduct.Products.Item(i).ReferenceProduct.Name = Null Then
            '            MsgBox("Make sure that all your part are loaded", vbCritical)
            '            Exit Sub
            '        End If

            '        If Not DictStrChildren.exists(ChildrenName) Then 'And IsComponent(SelectedItem.Products.Item(i)) = False Then

            '            Qty = 1
            '            DictChildren.Add(Key:=ChildrenName, Item:=ActiveProduct.Products.Item(i))
            '            DictStrChildren.Add(Key:=ChildrenName, Item:=Qty)

            '        ElseIf DictStrChildren.exists(ChildrenName) Then
            '            DictStrChildren.Item(ChildrenName) = DictStrChildren.Item(ChildrenName) + 1
            '        End If

            '    Next i


            '    Dim BOM1 As New BOM
            '    On Error Resume Next
            '    Call BOM1.ClearFormat3D()
            '    Range("Table3D").ClearContents()
            '    Range("Table3D").ClearFormats()
            '    Range("Table3D").WrapText = True
            '    On Error GoTo 0

            '    i = 0

            '    For Each Key In DictStrChildren
            '        ActiveSheet.Cells(i + 14, 1).Value = DictStrChildren(Key)
            '        ActiveSheet.Cells(i + 14, 2).Value = Key
            '        ActiveSheet.Cells(i + 14, 3).Value = DictChildren.Item(Key).DescriptionRef
            '        i = i + 1
            '    Next

            '    ActiveSheet.Shapes.Range(Array("3DPartNo")).TextFrame2.TextRange.Characters.Text = ActiveProduct.Name
            '    ActiveSheet.Shapes.Range(Array("3DDescription")).TextFrame2.TextRange.Characters.Text = ActiveProduct.DescriptionRef

            '    ActiveSheet.Cells(12, 3).Value = Date & " " & Time()
            'End Sub
            Public Function SelectMultiple3DProducts() As List(Of Products)

                'Dim SelectedProducts As Products
                Dim ActiveProductDocument As ProductDocument, ActiveProducts As Products
                Dim counter As Integer

                ActiveProductDocument = GetProductDocument()

                Dim What(0) As Object
                What(0) = "Product"

                Dim SelectedProducts As SelectedElement
                SelectedProducts = oCATIA.ActiveDocument.Selection
                SelectedProducts.Clear()

                Dim e As String
                e = SelectedProducts.SelectElement3(What, "Select a Product or a Component", False, 2, False)


                For counter = 1 To SelectedProducts.count
                    ActiveProducts.add(SelectedProducts.Item(counter).Value)
                Next

                SelectedProducts.Clear()

                Return ActiveProducts
            End Function
            Public Class cl_PartsList
                Public PartNo As String
                Public Quantity As String
                Public Nomenclature As String
                Public ItemNo As String
                Public Material As String
                Public Class cl_Parent
                    Public PartNo As String
                    Public Nomenclature As String
                End Class
            End Class
        End Class
        Public Class Part
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
            Function Select3DPart() As Part

                Dim CATIA As Object, ActivePartDocument As ProductDocument, ActivePart As Part

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
            Public PartNo As String
            Public Quantity As String
            Public Nomenclature As String
            Public ItemNo As String
            Public Material As String
            Public Class cl_Parent
                Public PartNo As String
                Public Nomenclature As String
            End Class
        End Class
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <returns></returns>
        Public Function PartsList() As List(Of cl_PartsList)
            Dim cl_PL As New cl_PartsList, item As Integer
            Dim oPartsList As New List(Of cl_PartsList)
            Dim Active2DTable As DrawingTable

            Active2DTable = Select2DTable()

            For item = 1 To Active2DTable.NumberOfRows

                'cl_PL.ItemNo = Active2DTable.GetCellString(1, 1)
                'cl_PL.Nomenclature = Active2DTable.GetCellString(item, 2)
                'cl_PL.Quantity = Active2DTable.GetCellString(item, 3)
                'cl_PL.PartNo = Active2DTable.GetCellString(item, 4)
                'cl_PL.Material = Active2DTable.GetCellString(item, 5)

                cl_PL.PartNo = Active2DTable.GetCellString(1, 1)

                oPartsList.Add(cl_PL)
                cl_PL = Nothing

            Next item


            Return oPartsList
        End Function
        Public Function GetDrawingDocument() As DrawingDocument
            oCATIA = GetCATIA()
            Dim MyDrawingDocument As DrawingDocument

            On Error Resume Next
            MyDrawingDocument = oCATIA.ActiveDocument
            If MyDrawingDocument Is Nothing Or Err.Number <> 0 Then
                MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a Drawing" & vbCrLf & "in the Active session", vbCritical, "Open a Drawing")
                Environment.Exit(0)
            End If
            GetDrawingDocument = MyDrawingDocument
        End Function
        Public Function Select2DTable() As DrawingTable
            Dim ActiveDrawingDocument As DrawingDocument, ActiveTable As DrawingTable
            Dim SelectedTable As Selection
            Dim e As String
            Dim What(0)

            oCATIA = GetCATIA()
            ActiveDrawingDocument = GetDrawingDocument()

            What(0) = "DrawingTable"
            'MsgBox("hi")

            SelectedTable = ActiveDrawingDocument.Selection
            SelectedTable.Clear()

            e = SelectedTable.SelectElement3(What, "Select a DrawingTable", True, CATMultiSelectionMode.CATMultiSelTriggWhenUserValidatesSelection, False)
            'e = SelectedTable.SelectElement2(What, "Select a DrawingTable", False)

            ActiveTable = SelectedTable.Item(1).Value
            SelectedTable.Clear()

            Select2DTable = ActiveTable
            MsgBox("hi")
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
                MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a Drawing" & vbCrLf & "in the Active session", vbCritical, "Open a Drawing")
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


    End Class
End Class
