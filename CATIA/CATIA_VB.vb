'Imports DrawingDocument = DRAFTINGITF.IID_DraftingIDLItf
Imports DrawingSheets = DRAFTINGITF.DrawingSheets
Imports DrawingSheet = DRAFTINGITF.DrawingSheet

Imports DrawingViews = DRAFTINGITF.DrawingViews
Imports DrawingView = DRAFTINGITF.DrawingView

Imports DRAFTINGITF.IID_DraftingInterfaces
Imports DRAFTINGITF
Imports DRAFTINGITF.CatTextProperty
Imports DRAFTINGITF.CatTablePosition

Imports MECMOD

Imports ProductStructureTypeLib

Imports INFITF

Imports Microsoft.Office.Interop.Excel
Public Class Cl_CATIA
    Shared CATIA As Object
    Shared Function GetCATIA() As Object

        CATIA = GetObject(, "CATIA.Application")
        If CATIA Is Nothing Or Err.Number <> 0 Then
            MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a CATIA session", vbCritical, "Open a CATIA Session ")
            Environment.Exit(0)
            '       Set CATIA = CreateObject("CATIA.Application")
            '       CATIA.Visible = True
        End If

        GetCATIA = CATIA
    End Function
    Public Class _3D
        Public Class Product
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
            Function GetProductDocument() As ProductDocument
                CATIA = GetCATIA()
                Dim MyProductDocument As ProductDocument

                On Error Resume Next
                MyProductDocument = CATIA.ActiveDocument
                If MyProductDocument Is Nothing Or Err.Number <> 0 Then
                    ' MsgBox "No CATIA Active Document found "
                    MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a Product" & vbCrLf & "in the Active session", vbCritical, "Open a Product")
                    Environment.Exit(0)
                End If
                GetProductDocument = MyProductDocument
            End Function
            Function SelectSingle3DProduct() As Products
                'Dim CATIA As Object,
                Dim ActiveProductDocument As ProductDocument, ActiveProduct As Products

                'CATIA = GetCATIA()
                ActiveProductDocument = GetProductDocument()

                Dim What(0) As Object
                What(0) = "Product"

                Dim SelectedProduct As SelectedElement
                SelectedProduct = CATIA.ActiveDocument.Selection
                SelectedProduct.Clear()

                Dim e As String
                e = SelectedProduct.SelectElement3(What, "Select a Product or a Component", False, 2, False)

                ActiveProduct = SelectedProduct.Item(1).Value
                SelectedProduct.Clear()

                Return ActiveProduct
            End Function

            Function PartsList(ActiveProduct As Products) As List(Of cl_PartsList)
                Dim testobject As String
                Dim cl_PL As New cl_PartsList
                If ActiveProduct.Count = 0 Then
                    Exit Function
                End If


                'Try
                '    Parallel.For(1, ActiveProduct.Count + 1, Sub(k)

                '                                                 Dim oInst As INFITF.AnyObject
                '                                                 oInst = ActiveProduct.Item(k)
                '                                                 'ActiveProduct.Item(k).ApplyWorkMode(DESIGN_MODE)   'apply design mode
                '                                                 'testobject = oInst.partnumber
                '                                                 'If Validation.IsComponent(oInst) = False And oInstances.Item(k).Parent.Parent.PartNumber = oInProduct.partnumber Then
                '                                                 'If Validation.IsComponent(oInst) = False And RealParent(oInst) = oInProduct.partnumber Then

                '                                                 cl_PL.PartNo = oInst.Partnumber
                '                                                 cl_PL.Nomenclature = oInst.Nomenclature

                '                                                 MsgBox(cl_PL.Nomenclature)

                '                                                 PartsList.Add(cl_PL)


                '                                                 'Realchildren3D.Add(oInst.partnumber)
                '                                                 '    comp.Add(oInst.partnumber)

                '                                                 'End If

                '                                                 'If Validation.IsComponent(oInst) = True And RealParent(oInst) = oInProduct.partnumber Then
                '                                                 '    Call WalkDownTree(oInst)
                '                                                 '    test = RealParent(oInst)

                '                                                 'End If

                '                                             End Sub)

                'Catch ex As Exception
                '    MsgBox("You need a multicore computer")
                'End Try

                For Each Child In ActiveProduct
                    cl_PL.PartNo = Child.Partnumber
                    cl_PL.Nomenclature = Child.Nomenclature
                    'cl_PL.cl_Parent.

                    MsgBox(cl_PL.Nomenclature)

                    PartsList.Add(cl_PL)
                Next


                Return PartsList(ActiveProduct)
            End Function
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
            Function SelectMultiple3DProducts() As List(Of Product)
                Dim SelectedProducts As Products

                Return SelectedProducts
            End Function
        End Class
        Public Class Part
            Function GetPartDocument() As PartDocument
                CATIA = GetCATIA()
                Dim MyPartDocument As PartDocument

                On Error Resume Next
                MyPartDocument = CATIA.ActiveDocument
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

            Function SelectAxis() As Axis

            End Function

            Function SelectMatingFace() As Face

            End Function

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

        Function PartsList() As List(Of cl_PartsList)
            Return PartsList
        End Function


        Function GetDrawingDocument() As DrawingDocument
            CATIA = GetCATIA()
            Dim MyDrawingDocument As DrawingDocument

            On Error Resume Next
            MyDrawingDocument = CATIA.ActiveDocument
            If MyDrawingDocument Is Nothing Or Err.Number <> 0 Then
                MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a Drawing" & vbCrLf & "in the Active session", vbCritical, "Open a Drawing")
                Environment.Exit(0)
            End If
            GetDrawingDocument = MyDrawingDocument
        End Function

        Function Select2DTable() As DrawingTable
            Dim CATIA As Object, ActiveDrawingDocument As DrawingDocument, ActiveTable As DrawingTable
            Dim e As String
            Dim What(0) As Object
            Dim SelectedTable As SelectedElement
            CATIA = GetCATIA()
            ActiveDrawingDocument = GetDrawingDocument()

            What(0) = "DrawingTable"

            SelectedTable = CATIA.ActiveDocument.Selection
            SelectedTable.Clear()

            e = SelectedTable.SelectElement3(What, "Select a DrawingTable", True, 2, False)

            ActiveTable = SelectedTable.Item(1).Value

            SelectedTable.Clear()
            Select2DTable = ActiveTable

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

        Sub Export2DTable(Table2D As DrawingTable, XLTable As ListObject)
            '            ActiveSheet.Cells(10, 8).Value = Null
            '            ActiveSheet.Cells(11, 7).Value = Null

            '            Dim CATIA As Object, oDrawingDocument As DrawingDocument
            '            CATIA = GetCATIA
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

            '                TotalRows = XLTable.ListRows.count

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
            'ActiveSheet.Cells(12, 8).Value = Date & " " & Time()

            '            ActiveSheet.Shapes.Range(Array("2DPartNo")).TextFrame2.TextRange.Characters.Text = oDrawingDocument.Parameters.Item("DRAWING_NUMBER").Value
            '            ActiveSheet.Shapes.Range(Array("2DDescription")).TextFrame2.TextRange.Characters.Text = oDrawingDocument.Parameters.Item("DRAWING_TITLE").Value
            '        End Sub

            '        Sub XLTo2D(Table As ListObject, QtyCol As Integer, PartNoCol As Integer, NomenclatureCol As Integer, Optional MatSpecCol As Integer, Optional ItemNoCol As Integer)

            '            Dim CATIA As Object
            '            CATIA = GetObject(, "CATIA.Application")

            '            Dim oDrawingDocument As DrawingDocument
            '            'Set oDrawingDocument = CATIA.ActiveDocument

            '            On Error Resume Next
            '            oDrawingDocument = CATIA.ActiveDocument

            '            If Err.Number <> 0 Then
            '                MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a Drawing" & vbCrLf & "in the Active session", vbCritical, "Open a Drawing")
            '                Exit Sub
            '            End If


            '            oDrwSheets = oDrawingDocument.Sheets
            '            oDrwSheet = oDrwSheets.ActiveSheet
            '            oDrwView = oDrwSheet.Views.ActiveView
            '            'oDrView.Select

            '            'Retrieve the view's tables collection
            '            Dim oDrwTables As DrawingTables
            '            oDrwTables = oDrwView.Tables

            '            ' Table.Range.AutoFilter Field:=2, Criteria1:=RGB(255, 255, 255), Operator:=xlFilterNoFill

            '            Dim TableCount As Integer

            '            Dim oRow As Range                           'See if there's parts that needs to be omitted
            '            For Each oRow In Table.DataBodyRange.Rows
            '                If oRow.Columns(2).Interior.Color <> 15773696 Then
            '                    TableCount = TableCount + 1
            '                End If
            '            Next oRow

            '            ' Create a new drawing table
            '            Dim oDrwTable As DrawingTable
            '            oDrwTable = oDrwTables.Add(896.650497436523, 126.999582529068, TableCount, 6, 5, 20)  ' double  iPositionX,double  iPositionY, long  iNumberOfRow, long  iNumberOfColumn, double  iRowHeight, double  iColumnWidth)

            '            ' Set the drawing table's name
            '            oDrwTable.Name = "Parts List"

            '            'Set the column sizes
            '            oDrwTable.SetColumnSize(1, 15.24)
            '            oDrwTable.SetColumnSize(2, 127)
            '            oDrwTable.SetColumnSize(3, 15.24)
            '            oDrwTable.SetColumnSize(4, 50.8)
            '            oDrwTable.SetColumnSize(5, 55.88)
            '            oDrwTable.SetColumnSize(6, 15.24)
            '            oDrwTable.AnchorPoint = CatTableBottomLeft

            '            Dim i As Integer
            '            i = 1


            '            Dim TotalRows As Integer, RowNbr As Integer
            '            'Totalrows = Table.Rows.count


            '            For Each oRow In Table.DataBodyRange.Rows

            '                If oRow.Columns(2).Interior.Color <> 15773696 Then

            '                    Dim ItemNo As String
            '                    ItemNo = i

            '        oDrwTable.SetCellString (i), 4, oRow.Columns(2).Value   'Parts
            '        oDrwTable.SetCellString (i), 3, oRow.Columns(1).Value   'Qty
            '        oDrwTable.SetCellString (i), 2, oRow.Columns(3).Value   'Nomenclature

            '                    If ItemNoCol = 0 Then
            '                        oDrwTable.SetCellString(i, 1, oDrwTable.NumberOfRows + 1 - ItemNo) 'ItemNo
            '                    Else
            '                        oDrwTable.SetCellString(i, 1, oRow.Columns(ItemNoCol).Value)
            '                    End If

            '                    If MatSpecCol <> 0 Then                                               'MatSpec
            '                        oDrwTable.SetCellString(i, 5, oRow.Columns(MatSpecCol).Value)
            '                    End If

            '        oDrwTable.SetCellAlignment (i), 2, CatTableMiddleLeft

            '                    i = i + 1
            '                End If
            '            Next oRow

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

        Sub AddRevisionBalloon(TableNotes As ListObject)

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

            oDrawingViews.Add("Notes")

            Dim oDrawingView As DrawingView
            oDrawingView = oDrawingViews.Item("Notes")

            Dim oDrawingText As DrawingText
            Dim Y As Integer, NoteCntr As Integer, Note As Range

            NoteCntr = 1

            oDrawingText = oDrawingView.Texts.Add(" GENERAL NOTES", 30, 530 - Y) 'First line of Text ' NOTES: UNLESS OTHERWISE SPECIFIED
            'oDrawingText.TextProperties.Justification = catLeft

            oDrawingText.WrappingWidth = 500
            oDrawingText.SetParameterOnSubString(catFontSize, 1, 13, 6350) 'oDrawingText.SetParameterOnSubString catBold, 1, 6, 1
            Y = Y + 15

            'For Each Note In TableNotes.ListColumns(3).DataBodyRange
            '    Dim SelectedNote As Range
            '    Set SelectedNote = Note.Offset(0, 4)
            '
            ''    If SelectedNote.Value2 = True Then
            ''
            '        Set oDrawingText = oDrawingView.Texts.Add(NoteCntr & ". " & Note.Value, 30, 530 - Y)
            '        oDrawingText.TextProperties.Justification = catLeft
            '        oDrawingText.WrappingWidth = 500
            '
            '        If Note.Offset(0, -1).Value <> "" Then
            '            oDrawingText.Name = "Note_ID_" & Note.Offset(0, 3).Value
            '        End If
            '
            '        'If Note.Offset(0, 2).Value = "Yes" Then Call DrwFlagNote(oDrawingText)
            '
            '        Y = Y + 15
            '        NoteCntr = NoteCntr + 1
            ''    End If
            'Next
        End Sub

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

        Sub XLTo2D(Table As ListObject, QtyCol As Integer, PartNoCol As Integer, NomenclatureCol As Integer, MatSpecCol As Integer, ItemNoCol As Integer)

            Dim CATIA As Object

            CATIA = GetObject(, "CATIA.Application")

            Dim oDrawingDocument 'As DrawingDocument
            'Set oDrawingDocument = CATIA.ActiveDocument

            On Error Resume Next

            oDrawingDocument = CATIA.ActiveDocument

            If Err.Number <> 0 Then
                MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a Drawing" & vbCrLf & "in the Active session", vbCritical, "Open a Drawing")
                Exit Sub
            End If

            Dim oDrwSheets As DrawingSheets
            Dim oDrwSheet As DrawingSheet

            Dim oDrwView As DrawingView
            oDrwSheets = oDrawingDocument.IID_DrawingSheets

            oDrwView = oDrwSheet.Views.ActiveView
            'oDrView.Select

            'Retrieve the view's tables collection
            Dim oDrwTables As DrawingTables
            oDrwTables = oDrwView.Tables

            ' Table.Range.AutoFilter Field:=2, Criteria1:=RGB(255, 255, 255), Operator:=xlFilterNoFill

            Dim TableCount As Integer

            Dim oRow As Range                           'See if there's parts that needs to be omitted
            For Each oRow In Table.DataBodyRange.Rows
                If oRow.Columns(2).Interior.Color <> 15773696 Then
                    TableCount = TableCount + 1
                End If
            Next oRow

            ' Create a new drawing table
            Dim oDrwTable As DrawingTable
            oDrwTable = oDrwTables.Add(896.650497436523, 126.999582529068, TableCount, 6, 5, 20)  ' double  iPositionX,double  iPositionY, long  iNumberOfRow, long  iNumberOfColumn, double  iRowHeight, double  iColumnWidth)

            ' Set the drawing table's name
            oDrwTable.Name = "Parts List"

            'Set the column sizes
            oDrwTable.SetColumnSize(1, 15.24)
            oDrwTable.SetColumnSize(2, 127)
            oDrwTable.SetColumnSize(3, 15.24)
            oDrwTable.SetColumnSize(4, 50.8)
            oDrwTable.SetColumnSize(5, 55.88)
            oDrwTable.SetColumnSize(6, 15.24)
            oDrwTable.AnchorPoint = CatTableBottomLeft

            Dim i As Integer
            i = 1


            Dim TotalRows As Integer, RowNbr As Integer
            'Totalrows = Table.Rows.count


            For Each oRow In Table.DataBodyRange.Rows

                If oRow.Columns(2).Interior.Color <> 15773696 Then

                    Dim ItemNo As String
                    ItemNo = i

                    ' ********FIX THIS YOU LAZY CODER 'oDrwTable.SetCellString (i), 4, oRow.Columns(2).Value   'Parts
                    'oDrwTable.SetCellString (i), 3, oRow.Columns(1).Value   'Qty
                    'oDrwTable.SetCellString (i), 2, oRow.Columns(3).Value   'Nomenclature

                    If ItemNoCol = 0 Then
                        oDrwTable.SetCellString(i, 1, oDrwTable.NumberOfRows + 1 - ItemNo) 'ItemNo
                    Else
                        oDrwTable.SetCellString(i, 1, oRow.Columns(ItemNoCol).Value)
                    End If

                    If MatSpecCol <> 0 Then                                               'MatSpec
                        oDrwTable.SetCellString(i, 5, oRow.Columns(MatSpecCol).Value)
                    End If

                    oDrwTable.SetCellAlignment(i, 2, CatTableMiddleLeft)

                    i = i + 1
                End If
            Next oRow

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
