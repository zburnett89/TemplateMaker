Imports System.Math

Class MainWindow

    Public corelApp As Corel.Interop.CorelDRAW.Application
    Public corelDoc As Corel.Interop.CorelDRAW.Document

    Dim ctrlRectangle, cornerRect, regDot, vertFlutes,
            horzFlutes, stkDot6x24, stkDot10x30, grommet As Corel.Interop.CorelDRAW.Shape
    Dim radTxt, holeSzTxt, holeDistTxt, holePlcTxt As String

    Private Sub ckbxCLRgroms_Checked(sender As Object, e As RoutedEventArgs) Handles ckbxCLRgroms.Checked, ckbxCLRgroms.Unchecked
        If ckbxCLRgroms.IsChecked Then
            ckbxLRspacing.IsEnabled = False
            ckbxLRspacing.IsChecked = False
        End If
        If ckbxCLRgroms.IsChecked = False Then
            ckbxLRspacing.IsEnabled = True
        End If
    End Sub

    Private Sub ckbxCTBGroms_Checked(sender As Object, e As RoutedEventArgs) Handles ckbxCTBGroms.Checked, ckbxCTBGroms.Unchecked
        If ckbxCTBGroms.IsChecked Then
            ckbxTBspacing.IsEnabled = False
            ckbxTBspacing.IsChecked = False
        End If
        If ckbxCTBGroms.IsChecked = False Then
            ckbxTBspacing.IsEnabled = True
        End If
    End Sub

    Private Sub ckbxLRspacing_Checked(sender As Object, e As RoutedEventArgs) Handles ckbxLRspacing.Checked, ckbxLRspacing.Unchecked
        If ckbxLRspacing.IsChecked Then
            ckbxCLRgroms.IsEnabled = False
            ckbxCLRgroms.IsChecked = False
        End If
        If ckbxLRspacing.IsChecked = False Then
            ckbxCLRgroms.IsEnabled = True
        End If
    End Sub

    Private Sub ckbxTBspacing_Checked(sender As Object, e As RoutedEventArgs) Handles ckbxTBspacing.Checked, ckbxTBspacing.Unchecked

        If ckbxTBspacing.IsChecked Then
            ckbxCTBGroms.IsEnabled = False
            ckbxCTBGroms.IsChecked = False
        End If
        If ckbxTBspacing.IsChecked = False Then
            ckbxCTBGroms.IsEnabled = True
        End If
    End Sub

    Dim txtDimWidth, txtDimHeight As Corel.Interop.VGCore.Shape

    Private Sub lstFlutes_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lstFlutes.SelectionChanged
        If lstFlutes.SelectedIndex = 1 Then
            lblStkDots.IsEnabled = False
            lstStkDots.IsEnabled = False 'Removes stake dot option for horizontal flutes
        Else
            lblStkDots.IsEnabled = True
            lstStkDots.IsEnabled = True
        End If

    End Sub

    Public Sub lstHoleQty_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lstHoleQty.SelectionChanged
        If lstHoleQty.SelectedIndex <> 0 And lstHoleQty.SelectedIndex <> -1 Then
            lblHoleSz.IsEnabled = True
            lstHoleSz.IsEnabled = True
            lblHoleDistance.IsEnabled = True
            lblHolePosition.IsEnabled = True
            lblHolesTBDist.IsEnabled = True
            txtTBDist.IsEnabled = True
            txtLRDist.IsEnabled = True
            lblHolesLRDist.IsEnabled = True
            ckbxUL.IsEnabled = True
            ckbxUC.IsEnabled = True
            ckbxUR.IsEnabled = True
            ckbxCL.IsEnabled = True
            ckbxCR.IsEnabled = True
            ckbxLL.IsEnabled = True
            ckbxLC.IsEnabled = True
            ckbxLR.IsEnabled = True
        Else
            lblHoleSz.IsEnabled = False
            lstHoleSz.IsEnabled = False
            lstHoleSz.SelectedIndex = -1
            lblHoleDistance.IsEnabled = False
            lblHolePosition.IsEnabled = False
            lblHolesTBDist.IsEnabled = False
            lblHolesLRDist.IsEnabled = False
            ckbxUL.IsEnabled = False
            ckbxUC.IsEnabled = False
            ckbxUR.IsEnabled = False
            ckbxCL.IsEnabled = False
            ckbxCR.IsEnabled = False
            ckbxLL.IsEnabled = False
            ckbxLC.IsEnabled = False
            ckbxLR.IsEnabled = False
        End If
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs) Handles btnClose.Click
        Close()
    End Sub

    Public Sub New()

        InitializeComponent()

        If corelApp Is Nothing Then
            corelApp = CType(CreateObject("CorelDRAW.Application"), Corel.Interop.CorelDRAW.Application)
        Else
            corelApp = CType(GetObject(, "CorelDRAW.Application"), Corel.Interop.CorelDRAW.Application)
        End If



    End Sub

    Public Sub lstMaterial_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lstMaterial.SelectionChanged

        If lstMaterial.SelectedIndex = 0 Then 'Coroplast
            lblFlutes.IsEnabled = True
            lstFlutes.IsEnabled = True
            lblStkDots.IsEnabled = True
            lstStkDots.IsEnabled = True
            lblCorners.IsEnabled = False
            lstCorners.IsEnabled = False
            lblHoles.IsEnabled = False
            lstHoleQty.IsEnabled = False
            lstCorners.SelectedIndex = -1
            lstHoleQty.SelectedIndex = -1
            lstHoleSz.SelectedIndex = -1
            txtLRDist.IsEnabled = False
            txtLRDist.Clear()
            txtTBDist.IsEnabled = False
            txtTBDist.Clear()
            ckbxUL.IsChecked = False
            ckbxUC.IsChecked = False
            ckbxUR.IsChecked = False
            ckbxCL.IsChecked = False
            ckbxCR.IsChecked = False
            ckbxLL.IsChecked = False
            ckbxLC.IsChecked = False
            ckbxLR.IsChecked = False
        ElseIf lstMaterial.SelectedIndex = -1 Or lstMaterial.SelectedIndex = 0 Then
            ckbxCR.IsEnabled = False
        ElseIf lstMaterial.SelectedIndex = 1 Or lstMaterial.SelectedIndex = 2 Then
            lblCorners.IsEnabled = True
            lstCorners.IsEnabled = True
            lblHoles.IsEnabled = True
            lstHoleQty.IsEnabled = True
            lblFlutes.IsEnabled = False
            lstFlutes.IsEnabled = False
            lstFlutes.SelectedIndex = -1
            lblStkDots.IsEnabled = False
            lstStkDots.IsEnabled = False
            lstStkDots.SelectedIndex = -1
        End If

    End Sub

    Public Sub BtnGo_Click(sender As Object, e As RoutedEventArgs) Handles btnGo.Click

        corelApp.Visible = True

        Dim pgHeight As Double = Val(txtHeight.Text)
        Dim pgWidth As Double = Val(txtWidth.Text)
        Dim valTBspacing As Single = Val(txtTBspacing.Text)
        Dim valLRspacing As Single = Val(txtLRspacing.Text)

        If ckbxTBspacing.IsChecked And Not Integer.TryParse(pgWidth / valTBspacing, 0) Then
            MsgBox("Width must be evenly divisible by grommet spacing distance.")
            txtTBspacing.Clear()
            Exit Sub
        End If

        If ckbxLRspacing.IsChecked And Not Integer.TryParse(pgHeight / valLRspacing, 0) Then
            MsgBox("Height must be evenly divisible by grommet spacing distance.")
            txtLRspacing.Clear()
            Exit Sub
        End If
        corelDoc = corelApp.CreateDocumentFromTemplate("C:\Program Files\Corel\CorelDRAW Graphics Suite 2017\Draw\SignMasterTemplate\TestTemplate.cdt")
        'corelDoc.Activate()


        'Sets the page size
        corelDoc.ActivePage.SizeWidth = pgWidth
        corelDoc.ActivePage.SizeHeight = pgHeight

        'Sets the control rectangle size & position
        ctrlRectangle = corelDoc.ActivePage.Shapes("ctrlRectangle")
        ctrlRectangle.SizeHeight = pgHeight
        ctrlRectangle.SizeWidth = pgWidth
        ctrlRectangle.SetPosition(0, pgHeight)

        txtDimWidth = corelDoc.ActivePage.Shapes("txtDimWidth")
        txtDimHeight = corelDoc.ActivePage.Shapes("txtDimHeight")

        'txtDimWidth.FontProperties.Size = 400

        'Sets flutes & position
        vertFlutes = corelDoc.ActivePage.Shapes("vertFlutes")

        horzFlutes = corelDoc.ActivePage.Shapes("horzFlutes")

        If lstFlutes.SelectedIndex = 0 Then 'Vertical flutes
            horzFlutes.Delete()
            vertFlutes.SetPosition(pgWidth + 2, pgHeight / 2 + vertFlutes.SizeHeight / 2)
        ElseIf lstFlutes.SelectedIndex = 1 Then 'Horizontal flutes
            vertFlutes.Delete()
            horzFlutes.SetPosition(pgWidth + 2, pgHeight / 2 + horzFlutes.SizeHeight / 2)
        ElseIf lstFlutes.SelectedIndex = 2 Then 'Both flutes
            vertFlutes.SetPosition(pgWidth + 2, pgHeight / 2 + vertFlutes.SizeHeight + 1)
            horzFlutes.SetPosition(pgWidth + 2, vertFlutes.PositionY - 4)
        Else 'No flutes
            horzFlutes.Delete()
            vertFlutes.Delete()
        End If

        'Sets stake dots & position
        stkDot6x24 = corelDoc.ActivePage.Shapes("stkDot6x24")
        stkDot10x30 = corelDoc.ActivePage.Shapes("stkDot10x30")

        If lstStkDots.SelectedIndex = -1 Or lstStkDots.SelectedIndex = 2 Then
            stkDot6x24.Delete()
            stkDot10x30.Delete()
        ElseIf lstFlutes.SelectedIndex = 1 Then '6x24
            stkDot10x30.Delete()
            stkDot6x24.SetPosition(pgWidth / 2 - stkDot6x24.SizeWidth / 2, 0.3)
        ElseIf lstStkDots.SelectedIndex = 0 Then '10x30
            stkDot6x24.Delete()
            stkDot10x30.SetPosition(pgWidth / 2 - stkDot10x30.SizeWidth / 2, 0.3)
        End If

        Select Case lstCorners.SelectedIndex
            Case 1
                cornerRect = corelDoc.ActiveLayer.CreateRectangle(0, 0, pgWidth, pgHeight)
                ctrlRectangle.Outline.Width = 0
                cornerRect.Rectangle.RadiusUpperLeft = 0.25
                cornerRect.Rectangle.RadiusLowerLeft = 0.25
                cornerRect.Rectangle.RadiusLowerRight = 0.25
                cornerRect.Rectangle.RadiusUpperRight = 0.25
            Case 2
                cornerRect = corelDoc.ActiveLayer.CreateRectangle(0, 0, pgWidth, pgHeight)
                ctrlRectangle.Outline.Width = 0
                cornerRect.Rectangle.RadiusUpperLeft = 0.5
                cornerRect.Rectangle.RadiusLowerLeft = 0.5
                cornerRect.Rectangle.RadiusLowerRight = 0.5
                cornerRect.Rectangle.RadiusUpperRight = 0.5
            Case 3
                cornerRect = corelDoc.ActiveLayer.CreateRectangle(0, 0, pgWidth, pgHeight)
                ctrlRectangle.Outline.Width = 0
                cornerRect.Rectangle.RadiusUpperLeft = 0.75
                cornerRect.Rectangle.RadiusLowerLeft = 0.75
                cornerRect.Rectangle.RadiusLowerRight = 0.75
                cornerRect.Rectangle.RadiusUpperRight = 0.75
            Case 4
                cornerRect = corelDoc.ActiveLayer.CreateRectangle(0, 0, pgWidth, pgHeight)
                ctrlRectangle.Outline.Width = 0
                cornerRect.Rectangle.RadiusUpperLeft = 1
                cornerRect.Rectangle.RadiusLowerLeft = 1
                cornerRect.Rectangle.RadiusLowerRight = 1
                cornerRect.Rectangle.RadiusUpperRight = 1
            Case 5
                cornerRect = corelDoc.ActiveLayer.CreateRectangle(0, 0, pgWidth, pgHeight)
                ctrlRectangle.Outline.Width = 0
                cornerRect.Rectangle.RadiusUpperLeft = 1.5
                cornerRect.Rectangle.RadiusLowerLeft = 1.5
                cornerRect.Rectangle.RadiusLowerRight = 1.5
                cornerRect.Rectangle.RadiusUpperRight = 1.5

        End Select

        Dim holeSz As Double = 0

        Select Case lstHoleSz.SelectedIndex
            Case 0
                holeSz = 0.1875
            Case 1
                holeSz = 0.25
            Case 2
                holeSz = 0.3125
            Case 3
                holeSz = 0.375
        End Select

        Dim tbDist As Double = Val(txtTBDist.Text)
        Dim lrDist As Double = Val(txtLRDist.Text)
        If ckbxUL.IsChecked Then
            corelApp.ActiveLayer.CreateEllipse2(lrDist, pgHeight - tbDist, holeSz / 2)
        End If
        If ckbxUC.IsChecked Then
            corelApp.ActiveLayer.CreateEllipse2(pgWidth / 2, pgHeight - tbDist, holeSz / 2)
        End If
        If ckbxUR.IsChecked Then
            corelApp.ActiveLayer.CreateEllipse2(pgWidth - lrDist, pgHeight - tbDist, holeSz / 2)
        End If
        If ckbxCL.IsChecked Then
            corelApp.ActiveLayer.CreateEllipse2(lrDist, pgHeight / 2, holeSz / 2)
        End If
        If ckbxCR.IsChecked Then
            corelApp.ActiveLayer.CreateEllipse2(pgWidth - lrDist, pgHeight / 2, holeSz / 2)
        End If
        If ckbxLL.IsChecked Then
            corelApp.ActiveLayer.CreateEllipse2(lrDist, tbDist, holeSz / 2)
        End If
        If ckbxLC.IsChecked Then
            corelApp.ActiveLayer.CreateEllipse2(pgWidth / 2, tbDist, holeSz / 2)
        End If
        If ckbxLR.IsChecked Then
            corelApp.ActiveLayer.CreateEllipse2(pgWidth - lrDist, tbDist, holeSz / 2)
        End If

        'Writing corner and hole description
        If lstCorners.SelectedIndex <> -1 Or lstHoleSz.SelectedIndex <> -1 Then


            If lstCorners.SelectedIndex <> -1 Then

                radTxt = lstCorners.SelectionBoxItem.ToString + " radius corners"

            End If
            If lstHoleSz.SelectedIndex <> -1 Then

                holeSzTxt = lstHoleSz.SelectionBoxItem.ToString + " holes"

                If tbDist = lrDist And ckbxUL.IsChecked And ckbxUR.IsChecked Then
                    holePlcTxt = tbDist.ToString
                    If tbDist = 0.25 Then
                        holePlcTxt = "1/4"
                    ElseIf tbDist = 0.375 Then
                        holePlcTxt = "3/8"
                    ElseIf tbDist = 0.5 Then
                        holePlcTxt = "1/2"
                    ElseIf tbDist = 0.625 Then
                        holePlcTxt = "5/8"
                    ElseIf tbDist = 0.75 Then
                        holePlcTxt = "3/4"
                    End If
                    holeSzTxt = holeSzTxt + ", " + holePlcTxt + "'' from edge"

                End If

            End If

            Dim holesAndCornersTxt As Corel.Interop.CorelDRAW.Shape = corelDoc.ActiveLayer.CreateArtisticText(0, 0, radTxt + vbCrLf + holeSzTxt, , ,
                                                    "Arial", , , , , Corel.Interop.VGCore.cdrAlignment.cdrCenterAlignment)

            holesAndCornersTxt.SetSize(pgWidth)
            holesAndCornersTxt.SetPosition(0, 0 - (pgHeight * 0.125))
            holesAndCornersTxt.AlignToShape(Corel.Interop.VGCore.cdrAlignType.cdrAlignHCenter, ctrlRectangle)


        End If

        'Grommets
        grommet = corelDoc.ActivePage.Shapes("grommet")
        Dim grommetUL, grommetUC, grommetTBSpacing, grommetCL, grommetLRSpacing As Corel.Interop.VGCore.Shape
        If ckbxCornerGroms.IsChecked Then

            grommetUL = grommet.Duplicate()
            grommetUL.SetPosition(0.625, pgHeight - 0.625)
            grommetUL.Duplicate(pgWidth - 2)
            grommetUL.Duplicate(, 0 - pgHeight + 2)
            grommetUL.Duplicate(pgWidth - 2, 0 - pgHeight + 2)

        End If

        If ckbxCTBGroms.IsChecked Then

            grommetUC = grommet.Duplicate()
            grommetUC.SetPosition(pgWidth / 2 - 0.375, pgHeight - 0.625)
            grommetUC.Duplicate(, 0 - pgHeight + 2)

        End If

        If ckbxCLRgroms.IsChecked Then

            grommetCL = grommet.Duplicate
            grommetCL.SetPosition(0.625, pgHeight / 2 + 0.375)
            grommetCL.Duplicate(pgWidth - 2)

        End If

        If ckbxTBspacing.IsChecked Then

            Dim totalTBdist As Single = valTBspacing - 0.375
            Dim i As Integer
            grommetTBSpacing = grommet.Duplicate()
            grommetTBSpacing.SetPosition(totalTBdist, pgHeight - 0.625)
            grommetTBSpacing.Duplicate(, 0 - pgHeight + 2)

            For i = 1 To Floor(pgWidth / valTBspacing - 2)
                grommetTBSpacing.Duplicate(valTBspacing * i)
                grommetTBSpacing.Duplicate(valTBspacing * i, 0 - pgHeight + 2)
            Next i
        End If

        If ckbxLRspacing.IsChecked Then

            Dim totalLRdist As Single = valLRspacing + 0.375
            Dim i As Integer
            grommetLRSpacing = grommet.Duplicate()
            grommetLRSpacing.SetPosition(0.625, totalLRdist)
            grommetLRSpacing.Duplicate(pgWidth - 2)

            For i = 1 To Floor(pgHeight / valLRspacing - 2)
                grommetLRSpacing.Duplicate(0, valLRspacing * i)
                grommetLRSpacing.Duplicate(pgWidth - 2, valLRspacing * i)
            Next i
        End If

        grommet.Delete()


        'Change view to fit everything on Layer 1

        Dim layer1 As Corel.Interop.CorelDRAW.ShapeRange
        layer1 = corelDoc.ActiveLayer.Shapes.All
        corelApp.ActiveWindow.ActiveView.ToFitShapeRange(layer1)
        'Sets registration dot position

        regDot = corelDoc.ActivePage.Shapes("RegDot")
        regDot.SetPosition(0 - 6, pgHeight + 6)

    End Sub
End Class
