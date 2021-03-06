﻿Imports DrawingDocument = DRAFTINGITF.IID_DraftingIDLItf
Imports MECMOD
Imports ProductStructureTypeLib
Imports LAYOUT2DITF
Imports INFITF
Imports CATANNOTITF
Imports CATANNOTITF.CatTextProperty


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
            Function GetProductDocument() As ProductDocument
                CATIA = GetCATIA()
                Dim MyProductDocument As ProductDocument

                On Error Resume Next
                MyProductDocument = CATIA.ActiveDocument
                If MyProductDocument Is Nothing Or Err.Number <> 0 Then
                    ' MsgBox "No Catia Active Document found "
                    MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a Product" & vbCrLf & "in the Active session", vbCritical, "Open a Product")
                    Environment.Exit(0)
                End If
                GetProductDocument = MyProductDocument
            End Function

            Sub Select3DProduct()

                Dim CATIA As Object, ActiveProductDocument As ProductDocument, ActProd As Products

                CATIA = GetCATIA()
                ActiveProductDocument = GetProductDocument()


                Dim What(0) As Object
                What(0) = "Product"

                Dim UserSel As Object
                UserSel = CATIA.ActiveDocument.Selection
                UserSel.Clear()

                Dim e As String
                e = UserSel.SelectElement3(What, "Select a Product or a Component", False, 2, False)

                Dim SelectedElement As Long, ActiveProduct

                ActiveProduct = UserSel.Item(1).Value
                UserSel.Clear()

                ActiveProduct.ApplyWorkMode(DESIGN_MODE)

                Dim DictStrChildren As Dictionary
                DictStrChildren = CreateObject("Scripting.Dictionary")

                Dim DictChildren As Dictionary
                DictChildren = CreateObject("Scripting.Dictionary")

                Dim ChildrenName As String

                For i = 1 To ActiveProduct.Products.count
                    On Error Resume Next

                    ChildrenName = ActiveProduct.Products.Item(i).PartNumber


                    If ActiveProduct.Products.Item(i).ReferenceProduct.Name = Null Then
                        MsgBox("Make sure that all your part are loaded", vbCritical)
                        Exit Sub
                    End If

                    If Not DictStrChildren.exists(ChildrenName) Then 'And IsComponent(SelectedItem.Products.Item(i)) = False Then

                        Qty = 1
                        DictChildren.Add(Key:=ChildrenName, Item:=ActiveProduct.Products.Item(i))
                        DictStrChildren.Add(Key:=ChildrenName, Item:=Qty)

                    ElseIf DictStrChildren.exists(ChildrenName) Then
                        DictStrChildren.Item(ChildrenName) = DictStrChildren.Item(ChildrenName) + 1
                    End If

                Next i


                Dim BOM1 As New BOM
                On Error Resume Next
                Call BOM1.ClearFormat3D()
                Range("Table3D").ClearContents()
                Range("Table3D").ClearFormats()
                Range("Table3D").WrapText = True
                On Error GoTo 0

                i = 0

                For Each Key In DictStrChildren
                    ActiveSheet.Cells(i + 14, 1).Value = DictStrChildren(Key)
                    ActiveSheet.Cells(i + 14, 2).Value = Key
                    ActiveSheet.Cells(i + 14, 3).Value = DictChildren.Item(Key).DescriptionRef
                    i = i + 1
                Next

                ActiveSheet.Shapes.Range(Array("3DPartNo")).TextFrame2.TextRange.Characters.Text = ActiveProduct.Name
                ActiveSheet.Shapes.Range(Array("3DDescription")).TextFrame2.TextRange.Characters.Text = ActiveProduct.DescriptionRef

ActiveSheet.Cells(12, 3).Value = Date & " " & Time()
            End Sub
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

            Function Select3DPart() As Part

                Dim CATIA As Object, ActiveProductDocument As ProductDocument, ActPart As Part

                CATIA = GetCATIA()
                ActivePartDocument = GetCATIAPartDocument()
                ActivePart = ActivePartDocument.Part

                Dim What(0) As Object
                What(0) = "Part"

                Dim UserSel 'As SelectedElement
                UserSel = CATIA.ActiveDocument.Selection
                UserSel.Clear()

                Dim e As String
                e = UserSel.SelectElement3(What, "Select a Part", False, 2, False)

                Dim SelectedElement As Long, ActiveProduct

                ActPart = UserSel.Item(1).Value
                UserSel.Clear()

                Select3DPart = ActivePart
            End Function
        End Class
    End Class
    Public Class Drawing

        Function GetDrawingDocument() As DrawingDocument
            CATIA = GetCATIA()
            Dim MyDrawingDocument As DrawingDocument

            On Error Resume Next
            MyDrawingDocument = CATIA.ActiveDocument
            If MyDrawingDocument Is Nothing Or Err.Number <> 0 Then
                MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a Drawing" & vbCrLf & "in the Active session", vbCritical, "Open a Drawing")
                End
            End If
            GetDrawingDocument = MyDrawingDocument
        End Function

        Function Select2DTable() As DrawingTable
            Dim CATIA As Object, ActiveDrawingDocument As DrawingDocument, UserSel As SelectedElement, ActTable As DrawingTable
            Dim e As String
            Dim What(0) As Object

            CATIA = GetCATIA()
            ActiveDrawingDocument = GetDrawingDocument()

            What(0) = "DrawingTable"

            UserSel2D = CATIA.ActiveDocument.Selection
            UserSel2D.Clear()

            e = UserSel2D.SelectElement3(What, "Select a DrawingTable", True, 2, False)

            ActTable = UserSel2D.Item(1).Value

            UserSel2D.Clear()
            Select2DTable = ActTable

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

            Dim FlagNotes() As String
            Dim FlagnotesCollection As New Collection

            Dim Notes() As String
            Dim NotesCollection As New Collection

            Dim k As Integer
            Dim j As Integer

            Dim Delimiter As String
            Delimiter = Chr(10) '+ Chr(32)

            Dim StartNote As New Collection
            Dim StartFlagnote As New Collection
            Dim GeneralNote As New Collection

            Dim TempNote As String
            Dim TempFlagnote As String

            Notes = Split(oDrawingText.Text, Delimiter)

            For j = LBound(Notes) To UBound(Notes)
                If Trim(Notes(j)) = vbNullString Then
                    StartNote.Add(j)
                End If
            Next j
            StartNote.Add(j)

            TempNote = vbNullString

            For k = 2 To StartNote.count
                For j = StartNote(k - 1) To StartNote(k) - 1
                    If Trim(Notes(j)) <> "" Then
                        TempNote = TempNote + Space(1) + Trim(Notes(j))
                    End If
                Next

                If Trim(TempNote) <> "" Then
                    NotesCollection.Add(TempNote)
                End If
                TempNote = vbNullString
            Next k

            'Last value






            'For Each Note In NotesCollection
            'Cells(i, 26) = Note
            'i = i + 1
            'Next

            GetNotes = NotesCollection
        End Function
    End Class

End Class
