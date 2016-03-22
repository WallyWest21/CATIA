Imports DrawingViews = DRAFTINGITF.DrawingViews
Imports DrawingView = DRAFTINGITF.DrawingView
Imports DrawingSheets = DRAFTINGITF.DrawingSheets
Imports DrawingSheet = DRAFTINGITF.DrawingSheet

'Imports DRAFTINGITF.IID_DraftingInterfaces
Imports DRAFTINGITF
Imports DRAFTINGITF.CatTextProperty
Imports DRAFTINGITF.CatTablePosition

Imports MECMOD
Imports MECMOD.CatConstraintType
Imports MECMOD.CatConstraintMode
Imports HybridShapeTypeLib


Imports ProductStructureTypeLib
Imports ProductStructureTypeLib.CatWorkModeType 'apply design mode

Imports INFITF
Imports INFITF.CATMultiSelectionMode
Imports System.Linq

Imports SPATypeLib 'Clash Analysis
Imports System.Xml
Imports System.IO


Imports System.Text.RegularExpressions

'Imports PARTITF

Public Class Cl_CATIA
    Public Shared oCATIA As INFITF.Application
    Public Shared Function GetCATIA() As INFITF.Application
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

            Public Sub CreateANewProduct(Optional NameOfProduct As String = "Product")
                Dim CATIADocuments As Documents
                CATIADocuments = oCATIA.Documents

                Dim NewProduct As ProductDocument
                NewProduct = CATIADocuments.Add("Product")


                Dim product1 As Product
                product1 = NewProduct.Product

                Dim products1 As Products
                products1 = product1.Products
            End Sub
            'Public Function InsertANewPart(Optional oPartNumber As String = "") As ProductDocument
            '    Dim ActiveProductDocument As ProductDocument

            '    ActiveProductDocument = GetProductDocument()

            '    Dim product1 As Product = ActiveProductDocument.Product

            '    Dim products1 As Product = product1.Products

            '    Dim product2 As Product = products1.AddNewComponent("Part", oPartNumber)

            '    Dim partDocument1 As PartDocument
            '    'partDocument1 = ActiveProductDocument.Item("Part3.CATPart")

            '    'Dim part1 As Part
            '    'part1 = partDocument1.Part
            '    Return ActiveProductDocument


            'End Function

            Public Sub InsertANewPart(PartNumberList As List(Of String))
                oCATIA = GetCATIA()
                Dim oDocuments As Documents
                oDocuments = oCATIA.Documents

                Dim oProductDocument As ProductDocument
                oProductDocument = oDocuments.Add("Product")


                'Dim ActiveProductDocument As ProductDocument

                'ActiveProductDocument = GetProductDocument()

                Dim product1 As Product = oProductDocument.Product

                Dim products1 As Products = product1.Products

                Dim product2 As Product

                For Each oPartNumber In PartNumberList 'http://www.maruf.ca/files/caadoc/CAAScdPstUseCases/CAAPstAddExternalComponent.htm

                    Dim oPartDocument As Document
                    oPartDocument = oCATIA.Documents.Item(oPartNumber + ".CATPart")

                    'MsgBox(oPartDocument.Name)

                    product2 = products1.AddExternalComponent(oPartDocument)

                    product2.Update()

                    oPartDocument.Close()
                Next

                product1.PartNumber = "B475101"

                'Dim partDocument1 As PartDocument
                'partDocument1 = ActiveProductDocument.Item("Part3.CATPart")

                'Dim part1 As Part
                'part1 = partDocument1.Part
                'Return ActiveProductDocument


            End Sub
            Public Sub SubInsertANewPart(Optional oPartNumber As String = "")
                Dim ActiveProductDocument As ProductDocument
                ActiveProductDocument = GetProductDocument()


                ActiveProductDocument.Activate()
                Dim product1 As Product = ActiveProductDocument.Product

                Dim products1 As Products = product1.Products

                'Threading.Thread.Sleep(10000)


                Dim product2 As Product
                Try
                    product2 = products1.AddNewComponent("Part", oPartNumber)
                    product2.Update()
                    product2.Update()
                Catch ex As Exception

                    MsgBox("Let's wait a while together")
                    Threading.Thread.Sleep(30000)

                End Try



                'Dim product2 As Product = AddNewComponent(products1, oPartNumber)



                'Dim partDocument1 As PartDocument
                'partDocument1 = ActiveProductDocument.Item(oPartNumber + ".CATPart")

                'Dim part1 As Part
                'part1 = partDocument1.Part

                Threading.Thread.Sleep(2000)

                'product2.Update()
                'product2.Update()

            End Sub
            Function AddNewComponent(Products1 As Products, opartnumber As String) As Product
                AddNewComponent = Products1.AddNewComponent("Part", opartnumber)
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
            Public Function GetPartDocument() As PartDocument
                oCATIA = GetCATIA()
                Dim MyPartDocument As PartDocument

                On Error Resume Next
                MyPartDocument = oCATIA.ActiveDocument
                If MyPartDocument Is Nothing Or Err.Number <> 0 Then
                    MsgBox("To avoid a beep" & vbCrLf & "Or a rude message" & vbCrLf & "Just open a Part" & vbCrLf & "in the Active session", vbCritical, "Open a Product")
                    Environment.Exit(0)
                End If

                GetPartDocument = MyPartDocument
            End Function

            Public Function CreateANewPartDocumnet(Optional PartNumber As String = "Part") As Document
                oCATIA = GetCATIA()
                Dim oDocuments As Documents
                oDocuments = oCATIA.Documents

                Dim oPartDocument As PartDocument
                oPartDocument = oDocuments.Add("Part")

                Dim oPart As Part
                oPart = oPartDocument.Part

                Dim oProduct As Product
                oProduct = oPartDocument.GetItem(Replace(oPartDocument.Name, ".CATPART", ""))
                oProduct.PartNumber = PartNumber
                oProduct.Update()



                Return oPartDocument
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
            Public Sub CreatePlanefromOffset(NameOfPlane As String, ActivePartDocument As PartDocument)
                'Dim ActivePartDocument As PartDocument
                Dim Part1 As Part

                'ActivePartDocument = GetPartDocument()
                Part1 = ActivePartDocument.Part

                Dim hybridShapeFactory1 As HybridShapeFactory = Part1.HybridShapeFactory

                Dim originElements1 As OriginElements = Part1.OriginElements

                Dim hybridShapePlaneExplicit1 As HybridShapePlaneExplicit = originElements1.PlaneXY

                Dim reference1 As Reference = Part1.CreateReferenceFromObject(hybridShapePlaneExplicit1)

                Dim hybridShapePlaneOffset1 As HybridShapePlaneOffset = hybridShapeFactory1.AddNewPlaneOffset(reference1, 20.0, False)
                hybridShapePlaneOffset1.Name = NameOfPlane

                Dim hybridBodies1 As HybridBodies = Part1.HybridBodies

                Dim hybridBody1 As HybridBody = hybridBodies1.Item("Geometrical Set.1")

                hybridBody1.AppendHybridShape(hybridShapePlaneOffset1)

            End Sub
            Public Function fctCreatePlanefromOffset(NameofPlane As String, Offset As Double, ActivePartDocument As PartDocument) As Reference
                'Dim ActivePartDocument As PartDocument


                'ActivePartDocument = GetPartDocument()
                Dim Part1 As Part
                Part1 = ActivePartDocument.Part

                Dim hybridShapeFactory1 As HybridShapeFactory
                hybridShapeFactory1 = Part1.HybridShapeFactory

                Dim originElements1 As OriginElements
                originElements1 = Part1.OriginElements

                Dim hybridShapePlaneExplicit1 As HybridShapePlaneExplicit


                Dim reference1 As Reference
                Dim PlaneOrientaion As Boolean
                Select Case UCase(NameofPlane)
                    Case "FRONT"
                        'reference1 = originElements1.PlaneXY
                        'reference1 = CreatePlanefromOffset()
                        hybridShapePlaneExplicit1 = originElements1.PlaneYZ
                        reference1 = originElements1.PlaneYZ
                        PlaneOrientaion = False
                    Case "REAR"
                        hybridShapePlaneExplicit1 = originElements1.PlaneYZ
                        reference1 = originElements1.PlaneYZ
                        PlaneOrientaion = True
                    Case "FWD"
                        hybridShapePlaneExplicit1 = originElements1.PlaneZX
                        reference1 = originElements1.PlaneZX
                        PlaneOrientaion = False
                    Case "AFT"
                        hybridShapePlaneExplicit1 = originElements1.PlaneZX
                        reference1 = originElements1.PlaneZX
                        PlaneOrientaion = True
                    Case "BOTTOM"
                        hybridShapePlaneExplicit1 = originElements1.PlaneXY
                        reference1 = originElements1.PlaneXY
                        PlaneOrientaion = True
                    Case "TOP"
                        hybridShapePlaneExplicit1 = originElements1.PlaneXY
                        reference1 = originElements1.PlaneXY
                        PlaneOrientaion = False
                    Case Else
                        MsgBox("Choose a proper Sketch Support Name")
                        Exit Function
                End Select



                reference1 = Part1.CreateReferenceFromObject(hybridShapePlaneExplicit1)

                Dim hybridShapePlaneOffset1 As HybridShapePlaneOffset
                hybridShapePlaneOffset1 = hybridShapeFactory1.AddNewPlaneOffset(reference1, Offset, PlaneOrientaion)
                hybridShapePlaneOffset1.Name = NameofPlane

                Dim hybridBodies1 As HybridBodies
                hybridBodies1 = Part1.HybridBodies

                Dim hybridBody1 As HybridBody
                hybridBody1 = hybridBodies1.Item("Geometrical Set.1")

                hybridBody1.AppendHybridShape(hybridShapePlaneOffset1)

                'Part1.InWorkObject = hybridShapePlaneOffset1

                Part1.Update()

                Return hybridBody1.HybridShapes.Item(NameofPlane)
            End Function
            Public Sub Split(SplittingElement As String, oSplitside As Boolean, Optional PartBodyName As String = "PartBody")

                Dim partDocument1 As PartDocument
                partDocument1 = GetPartDocument()

                Dim part1 As Part
                part1 = partDocument1.Part

                Dim bodies1 As Bodies
                bodies1 = part1.Bodies

                Dim body1 As Body
                body1 = bodies1.Item(PartBodyName)

                part1.InWorkObject = body1

                Dim shapeFactory1 As PARTITF.ShapeFactory
                shapeFactory1 = part1.ShapeFactory

                Dim reference1 As Reference
                reference1 = part1.CreateReferenceFromName("")

                Dim split1 As PARTITF.Split
                split1 = shapeFactory1.AddNewSplit(reference1, oSplitside)

                Dim hybridBodies1 As HybridBodies
                hybridBodies1 = part1.HybridBodies

                Dim hybridBody1 As HybridBody
                hybridBody1 = hybridBodies1.Item("Geometrical Set.1")

                Dim hybridShapes1 As HybridShapes
                hybridShapes1 = hybridBody1.HybridShapes

                Dim hybridShapePlaneOffset1 As HybridShapePlaneOffset
                hybridShapePlaneOffset1 = hybridShapes1.Item(SplittingElement)

                Dim reference2 As Reference
                reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)

                split1.Surface = reference2

                'split1.SplitSide = PARTITF.CatSplitSide.catPositiveSide

                part1.Update()


            End Sub
            Public Function CreateACenteredRectangle(Width As Double, Height As Double, Optional SketchSupport As String = "Bottom", Optional CenterX As Double = 0, Optional CenterY As Double = 0) As Sketch
                Dim partDocument1 As PartDocument
                partDocument1 = GetPartDocument()

                Dim part1 As Part = partDocument1.Part
                Dim hybridBodies1 As HybridBodies = part1.HybridBodies
                Dim hybridBody1 As HybridBody = hybridBodies1.Item("Geometrical Set.1")
                Dim sketches1 As Sketches = hybridBody1.HybridSketches
                Dim originElements1 As OriginElements = part1.OriginElements
                Dim reference1 As Reference


                Dim sketch1 As Sketch
                reference1 = hybridBody1.HybridShapes.Item(SketchSupport)
                sketch1 = sketches1.Add(reference1)

                Dim arrayOfVariantOfDouble1(8)
                arrayOfVariantOfDouble1(0) = 0#
                arrayOfVariantOfDouble1(1) = 0#
                arrayOfVariantOfDouble1(2) = 0#
                arrayOfVariantOfDouble1(3) = 1.0#
                arrayOfVariantOfDouble1(4) = 0#
                arrayOfVariantOfDouble1(5) = 0#
                arrayOfVariantOfDouble1(6) = 0#
                arrayOfVariantOfDouble1(7) = 1.0#
                arrayOfVariantOfDouble1(8) = 0#
                Dim sketch1Variant
                sketch1Variant = sketch1
                sketch1Variant.SetAbsoluteAxisData(arrayOfVariantOfDouble1)

                part1.InWorkObject = sketch1

                Dim factory2D1 As Factory2D = sketch1.OpenEdition()

                Dim geometricElements1 As GeometricElements = sketch1.GeometricElements

                Dim axis2D1 As Axis2D = geometricElements1.Item("AbsoluteAxis")
                Dim line2D1 As Line2D = axis2D1.GetItem("HDirection")
                Dim line2D2 As Line2D = axis2D1.GetItem("VDirection")
                line2D1.ReportName = 1
                line2D2.ReportName = 2

                Dim point2D1 As Point2D = factory2D1.CreatePoint(CenterX, CenterY)
                Dim point2D2 As Point2D = factory2D1.CreatePoint(CenterX + Width / 2, CenterY + Height / 2)
                Dim point2D3 As Point2D = factory2D1.CreatePoint(CenterX + Width / 2, CenterY - Height / 2)
                Dim line2D3 As Line2D = factory2D1.CreateLine(CenterX + Width / 2, CenterY + Height / 2, CenterX + Width / 2, CenterY - Height / 2)
                point2D1.ReportName = 3
                point2D2.ReportName = 4
                point2D3.ReportName = 5
                line2D3.ReportName = 6
                line2D3.StartPoint = point2D2
                line2D3.EndPoint = point2D3


                Dim point2D4 As Point2D = factory2D1.CreatePoint(CenterX - Width / 2, CenterY - Height / 2)
                Dim line2D4 As Line2D = factory2D1.CreateLine(CenterX + Width / 2, CenterY - Height / 2, CenterX - Width / 2, CenterY - Height / 2)
                point2D4.ReportName = 7
                line2D4.ReportName = 8
                line2D4.StartPoint = point2D3
                line2D4.EndPoint = point2D4


                Dim point2D5 As Point2D = factory2D1.CreatePoint(CenterX - Width / 2, CenterY + Height / 2)
                Dim line2D5 As Line2D = factory2D1.CreateLine(CenterX - Width / 2, CenterY - Height / 2, CenterX - Width / 2, CenterY + Height / 2)
                point2D5.ReportName = 9
                line2D5.ReportName = 10
                line2D5.StartPoint = point2D4
                line2D5.EndPoint = point2D5


                Dim line2D6 As Line2D = factory2D1.CreateLine(CenterX - Width / 2, CenterY + Height / 2, CenterX + Width / 2, CenterY + Height / 2)
                line2D6.ReportName = 11
                line2D6.StartPoint = point2D5
                line2D6.EndPoint = point2D2

                Dim constraints1 As Constraints = sketch1.Constraints

                Dim reference2 As Reference = part1.CreateReferenceFromObject(line2D3)
                Dim reference3 As Reference = part1.CreateReferenceFromObject(line2D2)
                Dim constraint1 As Constraint = constraints1.AddBiEltCst(catCstTypeVerticality, reference2, reference3)
                constraint1.Mode = catCstModeDrivingDimension

                Dim reference4 As Reference = part1.CreateReferenceFromObject(line2D4)
                Dim reference5 As Reference = part1.CreateReferenceFromObject(line2D1)
                Dim constraint2 As Constraint = constraints1.AddBiEltCst(catCstTypeHorizontality, reference4, reference5)
                constraint2.Mode = catCstModeDrivingDimension

                Dim reference6 As Reference = part1.CreateReferenceFromObject(line2D5)
                Dim reference7 As Reference = part1.CreateReferenceFromObject(line2D2)
                Dim constraint3 As Constraint = constraints1.AddBiEltCst(catCstTypeVerticality, reference6, reference7)
                constraint3.Mode = catCstModeDrivingDimension

                Dim reference8 As Reference = part1.CreateReferenceFromObject(line2D6)
                Dim reference9 As Reference = part1.CreateReferenceFromObject(line2D1)
                Dim constraint4 As Constraint = constraints1.AddBiEltCst(catCstTypeHorizontality, reference8, reference9)
                constraint4.Mode = catCstModeDrivingDimension

                Dim reference10 As Reference = part1.CreateReferenceFromObject(line2D3)
                Dim reference11 As Reference = part1.CreateReferenceFromObject(line2D5)
                Dim reference12 As Reference = part1.CreateReferenceFromObject(point2D1)

                Dim constraint5 As Constraint = constraints1.AddTriEltCst(catCstTypeEquidistance, reference10, reference11, reference12)
                Dim reference13 As Reference = part1.CreateReferenceFromObject(line2D4)
                Dim reference14 As Reference = part1.CreateReferenceFromObject(line2D6)
                constraint5.Mode = catCstModeDrivingDimension

                Dim reference15 As Reference = part1.CreateReferenceFromObject(point2D1)
                Dim constraint6 As Constraint = constraints1.AddTriEltCst(catCstTypeEquidistance, reference13, reference14, reference15)
                constraint6.Mode = catCstModeDrivingDimension

                sketch1.CloseEdition()

                part1.InWorkObject = hybridBody1

                Return sketch1
            End Function
            Public Function CreateASketchCircle(iX As Double, iY As Double, Diameter As Double, Optional SketchSupport As String = "XY") As Sketch
                Dim partDocument1 As PartDocument
                partDocument1 = GetPartDocument()

                Dim part1 As Part
                part1 = partDocument1.Part

                Dim hybridBodies1 As HybridBodies
                hybridBodies1 = part1.HybridBodies

                Dim hybridBody1 As HybridBody
                hybridBody1 = hybridBodies1.Item("Geometrical Set.1")

                Dim sketches1 As Sketches
                sketches1 = hybridBody1.HybridSketches

                Dim originElements1 As OriginElements
                originElements1 = part1.OriginElements

                Dim reference1 As Reference

                Select Case SketchSupport
                    Case "XY"
                        reference1 = originElements1.PlaneXY
                    Case "YZ"
                        reference1 = originElements1.PlaneYZ
                    Case "ZX"
                        reference1 = originElements1.PlaneZX
                End Select

                Dim sketch1 As Sketch
                sketch1 = sketches1.Add(reference1)

                Dim arrayOfVariantOfDouble1(8)
                arrayOfVariantOfDouble1(0) = 0#
                arrayOfVariantOfDouble1(1) = 0#
                arrayOfVariantOfDouble1(2) = 0#
                arrayOfVariantOfDouble1(3) = 1.0#
                arrayOfVariantOfDouble1(4) = 0#
                arrayOfVariantOfDouble1(5) = 0#
                arrayOfVariantOfDouble1(6) = 0#
                arrayOfVariantOfDouble1(7) = 1.0#
                arrayOfVariantOfDouble1(8) = 0#

                Dim sketch1Variant
                sketch1Variant = sketch1
                sketch1Variant.SetAbsoluteAxisData(arrayOfVariantOfDouble1)

                part1.InWorkObject = sketch1

                Dim factory2D1 As Factory2D
                factory2D1 = sketch1.OpenEdition()

                Dim geometricElements1 As GeometricElements
                geometricElements1 = sketch1.GeometricElements

                Dim axis2D1 As Axis2D
                axis2D1 = geometricElements1.Item("AbsoluteAxis")

                Dim line2D1 As Line2D
                line2D1 = axis2D1.GetItem("HDirection")

                line2D1.ReportName = 1

                Dim line2D2 As Line2D
                line2D2 = axis2D1.GetItem("VDirection")

                line2D2.ReportName = 2

                Dim point2D1 As Point2D
                point2D1 = factory2D1.CreatePoint(iX, iY)

                point2D1.ReportName = 3

                Dim circle2D1 As Circle2D
                circle2D1 = factory2D1.CreateClosedCircle(iX, iY, Diameter)

                circle2D1.CenterPoint = point2D1

                circle2D1.ReportName = 4

                sketch1.CloseEdition()

                part1.InWorkObject = hybridBody1

                'part1.Update()
                Return sketch1
            End Function

            Public Sub Pad(SketchSupport As String, Optional PartBodyName As String = "PartBody", Optional Thickness As Double = 12.7)
                Dim partDocument1 As PartDocument = GetPartDocument()
                Dim part1 As Part = partDocument1.Part
                Dim bodies1 As Bodies = part1.Bodies
                Dim body1 As Body = bodies1.Item("PartBody")
                body1.Name = PartBodyName
                part1.InWorkObject = body1


                Dim shapeFactory1 As PARTITF.ShapeFactory = part1.ShapeFactory
                Dim reference1 As Reference = part1.CreateReferenceFromName("")
                Dim pad1 As PARTITF.Pad = shapeFactory1.AddNewPadFromRef(reference1, Thickness)
                Dim hybridBodies1 As HybridBodies = part1.HybridBodies
                Dim hybridBody1 As HybridBody = hybridBodies1.Item("Geometrical Set.1")

                Dim reference2 As Reference = part1.CreateReferenceFromObject(CreateACenteredRectangle(3000, 3000, SketchSupport))
                pad1.SetProfileElement(reference2)

                body1.Name = PartBodyName
                part1.Update()
            End Sub

        End Class
    End Class
    Public Class Drawing

        Public Function ListOfOpenDrawingDocuments() As List(Of String)
            ListOfOpenDrawingDocuments = New List(Of String)

            Dim oDocuments As Documents
            oDocuments = GetCATIA.Documents

            Dim oDocument As Document

            Dim strType As String

            Dim AvailableDocsPartNo() As String
            For Each oDocument In oDocuments
                strType = TypeName(oDocument)

                Select Case strType
                    Case "DrawingDocument"

                        AvailableDocsPartNo = oDocument.Name.Split(".")
                        ListOfOpenDrawingDocuments.Add(AvailableDocsPartNo(0))

                End Select
            Next
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
                        tempQtyList = New List(Of String)                                                           'Find a way to improve on getting quantity from table, CUZ THIS IS GARBAGE!!!!!
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

        Public Sub WriteNotesToDrawing()

        End Sub
        Public Sub WriteNotesToDrawing(Notes As List(Of String))

            Dim oDrawingDocument As DrawingDocument = GetDrawingDocument()
            Dim oDrawingSheets As DrawingSheets = oDrawingDocument.Sheets
            Dim oDrawingSheet As DrawingSheet = oDrawingSheets.ActiveSheet
            Dim oDrawingViews As DrawingViews = oDrawingSheet.Views

            oDrawingViews.Add("Notes")

            Dim oDrawingView As DrawingView = oDrawingViews.Item("Notes")
            Dim oDrawingText As DrawingText

            Dim Y As Integer

            oDrawingText = oDrawingView.Texts.Add("NOTES: UNLESS OTHERWISE SPECIFIED", 30, 530 - Y)
            oDrawingText.TextProperties.Justification = CatJustification.catLeft

            oDrawingText.WrappingWidth = 500
            oDrawingText.SetParameterOnSubString(catBold, 1, 6, 1)
            Y = Y + 15

            Dim count As Integer = 1
            For Each Note In Notes

                oDrawingText = oDrawingView.Texts.Add(count & ". " & Note, 30, 530 - Y)
                oDrawingText.TextProperties.Justification = CatJustification.catLeft

                oDrawingText.WrappingWidth = 500
                oDrawingText.Name = "Note_ID_" & count

                If (count Mod 2) = 0 Then Call DrwFlagNote(oDrawingText)

                Y = Y + 15
                count += 1
            Next

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

        Shared Function GetProductDocument() As ProductDocument
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

        Public Class ClashAnalysis
            Public Sub ActiveProductClash()

                GetCATIA.DisplayFileAlerts = False

                Dim Message, Style, Title, Response, MyString

                Message = ("If you have a big Assembly, be aware that this macro will start to put everything in Design Mode. If you don't want, put a omment in front of lines 3 to 12" &
                        (Chr(13)) & "" & (Chr(13)) & "   Do you want to continue ?")

                Style = vbYesNo + vbDefaultButton1   'Define buttons.

                Title = "Purpose "

                Response = MsgBox(Message, Style, Title)

                If Response = vbYes Then   ' User chose Yes.

                    MyString = "Yes"

                    Dim productDocument1 As Document
                    productDocument1 = GetProductDocument()

                    Dim product1 As Product
                    product1 = productDocument1.Product
                    product1.ApplyWorkMode(DESIGN_MODE)

                    Dim cClashes As Clashes
                    cClashes = GetCATIA.ActiveDocument.Product.GetTechnologicalObject("Clashes")

                    Dim oClash As Clash
                    oClash = cClashes.AddFromSel
                    oClash.ComputationType = CatClashComputationType.catClashComputationTypeInsideOne
                    oClash.InterferenceType = CatClashInterferenceType.catClashInterferenceTypeContact

                    oClash.Compute()


                    'Probably Garbage!!!!!!!!!
                    '************************************************
                    Dim cConflicts As Conflicts
                    cConflicts = oClash.Conflicts

                    Dim I As Integer
                    For I = 1 To cConflicts.Count

                        Dim oConflict As Conflict
                        oConflict = cConflicts.Item(I)

                        'MsgBox oConflict.Value

                        If (oConflict.Type = CatConflictType.catConflictTypeClash) Then

                            If (oConflict.Value <> 0) Then

                                oConflict.Status = CatConflictStatus.catConflictStatusRelevant
                                Dim dFilterValue
                                oConflict.Comment = "Automatic filter : penetration less than " & dFilterValue
                            End If
                        End If

                    Next

                    '**********************************************
                    oClash.Export(CatClashExportType.CatClashExportTypeXMLResultOnly, "c:\Temp\sample.xml")

                End If
            End Sub
            Public Function ListOfClashes() As List(Of cl_CA_Metadata)

                ListOfClashes = New List(Of cl_CA_Metadata)
                Dim CA As New cl_CA_Metadata

                Dim di As New DirectoryInfo("C:\Temp\DataBase")
                ' Get a reference to each file in that directory.
                Dim fiArr As FileInfo() = di.GetFiles()
                ' Display the names of the files.
                Dim fri As FileInfo
                For Each fri In fiArr

                    CA = New cl_CA_Metadata

                    Try

                        Dim m_xmld As XmlDocument
                        Dim m_nodelist As XmlNodeList
                        Dim m_node As XmlNode
                        'Create the XML Document
                        m_xmld = New XmlDocument()
                        'Load the Xml file
                        m_xmld.Load(fri.FullName)
                        'Get the list of name nodes 

                        m_node = m_xmld.SelectSingleNode("/InterferenceElement")

                        CA.ProductDocument = m_node.Attributes.GetNamedItem("CATDocument").Value
                        Dim SplitProductDocument As String()
                        SplitProductDocument = Split(CA.ProductDocument, "\")
                        CA.ProductDocument = SplitProductDocument(SplitProductDocument.Count - 1)

                        m_node = m_xmld.SelectSingleNode("/InterferenceElement/Interference")
                        CA.Status = m_node.Attributes.GetNamedItem("Status").Value
                        CA.Type = m_node.Attributes.GetNamedItem("ResultType").Value


                        m_nodelist = m_xmld.SelectNodes("/InterferenceElement/Interference/Product")
                        'Loop through the nodes
                        Dim count As Integer = 1
                        For Each m_node In m_nodelist
                            Select Case count
                                Case 1
                                    CA.Product1 = m_node.Attributes.GetNamedItem("Alias").Value
                                Case 2
                                    CA.Product2 = m_node.Attributes.GetNamedItem("Alias").Value
                            End Select

                            count += 1
                        Next
                        m_node = m_xmld.SelectSingleNode("/InterferenceElement/Interference/GeometricAspect")
                        CA.Value = m_node.Attributes.GetNamedItem("ExtractOrDistValue").Value

                        m_node = m_xmld.SelectSingleNode("/InterferenceElement/Interference/Picture")
                        CA.PictureHRef = m_node.Attributes.GetNamedItem("HRef").Value

                    Catch errorVariable As Exception
                        'Error trapping
                        MsgBox("Oops! Can not load XML file")
                        Exit Function
                    End Try
                    ListOfClashes.Add(CA)
                Next fri

                CA = Nothing
                Return ListOfClashes
            End Function
            Public Class cl_CA_Metadata

                Private _Priority As Integer
                Public Property Priority() As Integer
                    Get
                        'Dim rnd As New Random
                        Select Case Math.Abs((CType(Value, Double)))
                            Case 0
                                _Priority = 3

                            Case Is <= 0.125
                                _Priority = 2

                            Case Is > 0.125
                                _Priority = 1
                        End Select

                        Return _Priority
                    End Get
                    Set(ByVal value As Integer)

                        _Priority = value
                    End Set
                End Property
                Private _PriorityColor As String
                Public Property PriorityColor() As String
                    Get
                        Select Case Priority
                            Case 1
                                _PriorityColor = "Crimson"
                            Case 2
                                _PriorityColor = "DarkBlue"
                            Case 3
                                _PriorityColor = "DarkGreen"
                        End Select

                        Return _PriorityColor
                    End Get
                    Set(ByVal value As String)

                        _PriorityColor = value
                    End Set
                End Property
                Private _Product1 As String
                Public Property Product1() As String
                    Get
                        Return _Product1
                    End Get
                    Set(ByVal value As String)
                        _Product1 = value
                    End Set
                End Property

                Private _Product2 As String
                Public Property Product2() As String
                    Get
                        Return _Product2
                    End Get
                    Set(ByVal value As String)
                        _Product2 = value
                    End Set
                End Property

                Private _ProductDocument As String
                Public Property ProductDocument() As String
                    Get
                        Return _ProductDocument
                    End Get
                    Set(ByVal value As String)
                        _ProductDocument = value
                    End Set
                End Property

                Private _Type As String
                Public Property Type() As String
                    Get
                        Return _Type
                    End Get
                    Set(ByVal value As String)
                        _Type = value
                    End Set
                End Property

                Private _Value As String
                Public Property Value() As String
                    Get
                        Return _Value
                    End Get
                    Set(ByVal value As String)

                        'Dim regex As Regex = New Regex("/^\d*\.?\d*$/")
                        'Dim match As Match = regex.Match(value)
                        'If match.Success Then
                        '    Console.WriteLine(match.Value)
                        'End If

                        value = value.Replace("mm", vbNullString)
                        _Value = value.Replace("in", vbNullString)
                    End Set
                End Property

                Private _Status As String
                Public Property Status() As String
                    Get
                        Return _Status
                    End Get
                    Set(ByVal value As String)

                        _Status = value
                    End Set
                End Property

                Private _Comment As String
                Public Property Comment() As String
                    Get
                        Return _Comment
                    End Get
                    Set(ByVal value As String)
                        _Comment = value
                    End Set
                End Property

                Private _Level As String
                Public Property Level() As Integer
                    Get
                        Return _Level
                    End Get
                    Set(ByVal value As Integer)
                        _Level = value
                    End Set
                End Property

                Private _PictureHRef As String
                Public Property PictureHRef() As String
                    Get
                        Return _PictureHRef
                    End Get
                    Set(ByVal value As String)
                        _PictureHRef = value
                    End Set
                End Property
            End Class
        End Class
        Public Class Panel

        End Class
        Public Class Drawer
            Dim ooCATIA As New Cl_CATIA
            Dim oPart As New _3D.oPart
            Dim oProduct As New _3D.oProduct
            Dim oFS As Double, oWL As Double, oRBL As Double
            Dim oHeight As Double, oDepth As Double, oWidth As Double
            Dim FrontPlane As Reference, BottomPlane As Reference, FWDPlane As Reference, AFTPlane As Reference, TopPlane As Reference, RearPlane As Reference
            Dim ActiveProductDocument As ProductDocument, CATIADocuments As Documents
            Dim oPartNumber As New List(Of String)
            Sub New(FS As Double, WL As Double, RBL As Double, Height As Double, Depth As Double, Width As Double)
                oFS = FS
                oWL = WL
                oRBL = RBL
                oHeight = Height + oWL
                oDepth = Depth + oRBL
                oWidth = Width + oFS

                Dim ooPartnumber As New List(Of String)(New String() {"B472402-501", "B472402-503", "B472402-505", "B472402-507", "B472402-509"})
                oPartNumber = ooPartnumber
            End Sub
            Public Sub Create()
                CreateFrontPanel()
                CreateRearPanel()
                CreateFWDPanel()
                CreateAFTPanel()
                CreateBottomPanel()

                oProduct.InsertANewPart(oPartNumber)
            End Sub
            Sub CreateReferencePlanes(count As Integer)
                'CATIADocuments = GetCATIA.Documents
                'oProduct.SubInsertANewPart(oPartNumber(count))

                'Threading.Thread.Sleep(2000)


                Dim ActivePartDocument As PartDocument

                Dim stroPartnumber As String = oPartNumber(count) + ".CATPart"
                'ActivePartDocument = CATIADocuments.Item(stroPartnumber)
                ActivePartDocument = oPart.CreateANewPartDocumnet(oPartNumber(count))

                FrontPlane = oPart.fctCreatePlanefromOffset("FRONT", oRBL, ActivePartDocument)
                RearPlane = oPart.fctCreatePlanefromOffset("REAR", oDepth, ActivePartDocument)
                FWDPlane = oPart.fctCreatePlanefromOffset("FWD", oFS, ActivePartDocument)
                AFTPlane = oPart.fctCreatePlanefromOffset("AFT", oWidth, ActivePartDocument)
                BottomPlane = oPart.fctCreatePlanefromOffset("BOTTOM", oWL, ActivePartDocument)
                TopPlane = oPart.fctCreatePlanefromOffset("TOP", oHeight, ActivePartDocument)
            End Sub
            Sub TrimPanel(TrimmingPlane As String, TrimSide As Boolean, Optional PartBodyName As String = "PartBody")
                oPart.Split(TrimmingPlane, TrimSide, PartBodyName)
            End Sub
            Public Sub CreateFrontPanel()
                Dim PartBodyName As String
                PartBodyName = "FRONT PANEL"

                CreateReferencePlanes(0)

                oPart.Pad("FRONT", PartBodyName)

                TrimPanel("TOP", 1, PartBodyName)
                TrimPanel("BOTTOM", 0, PartBodyName)
                TrimPanel("FWD", 1, PartBodyName)
                TrimPanel("AFT", 0, PartBodyName)
            End Sub
            Public Sub CreateRearPanel()
                Dim PartBodyName As String
                PartBodyName = "REAR PANEL"

                CreateReferencePlanes(1)

                oPart.Pad("REAR", PartBodyName)

                TrimPanel("TOP", 1, PartBodyName)
                TrimPanel("BOTTOM", 0, PartBodyName)
                TrimPanel("FWD", 1, PartBodyName)
                TrimPanel("AFT", 0, PartBodyName)
            End Sub
            Public Sub CreateBottomPanel()
                Dim PartBodyName As String
                PartBodyName = "BOTTOM PANEL"

                CreateReferencePlanes(2)

                oPart.Pad("BOTTOM", PartBodyName)

                TrimPanel("REAR", 0, PartBodyName)
                TrimPanel("FRONT", 1, PartBodyName)
                TrimPanel("FWD", 1, PartBodyName)
                TrimPanel("AFT", 0, PartBodyName)
            End Sub
            Public Sub CreateFWDPanel()
                Dim PartBodyName As String
                PartBodyName = "FWD PANEL"

                CreateReferencePlanes(3)

                oPart.Pad("FWD", PartBodyName)

                TrimPanel("TOP", 1, PartBodyName)
                TrimPanel("BOTTOM", 0, PartBodyName)
                TrimPanel("REAR", 0, PartBodyName)
                TrimPanel("FRONT", 1, PartBodyName)
            End Sub
            Public Sub CreateAFTPanel()
                Dim PartBodyName As String
                PartBodyName = "AFT PANEL"

                CreateReferencePlanes(4)

                oPart.Pad("AFT", PartBodyName)

                TrimPanel("TOP", 1, PartBodyName)
                TrimPanel("BOTTOM", 0, PartBodyName)
                TrimPanel("REAR", 0, PartBodyName)
                TrimPanel("FRONT", 1, PartBodyName)
            End Sub
        End Class

        Public Class ECNMaker
            Public Class Drawing_Metadata
                Public Function ListofSheets() As List(Of DrawingSheet)
                    Return ListofSheets
                End Function

                Public Function ListofViews() As List(Of DrawingView)
                    Return ListofViews()
                End Function
                Public Function ListofTables() As List(Of DrawingTable)
                    Return ListofTables()
                End Function

                Public Function ListofCallouts() As List(Of DrawingText)
                    Return ListofCallouts()
                End Function

                Public Function ListofDimensions() As List(Of DrawingDimension)
                    Return ListofDimensions()
                End Function
                Public Function ListofNotes() As List(Of String)
                    Return ListofNotes()
                End Function

            End Class
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