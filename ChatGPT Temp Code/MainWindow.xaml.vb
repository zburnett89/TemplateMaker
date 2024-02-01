Imports System.Math
Imports System.Xml.Linq
Imports Corel.Interop.VGCore

Class MainWindow

    Public corelApp As Corel.Interop.VGCore.Application
    Public corelDoc As Corel.Interop.VGCore.Document

    Dim ctrlRectangle, cornerRect, regDot, vertFlutes,
            horzFlutes, stkDot6x24, stkDot10x30, grommet, tagBorder, tagHoles As Corel.Interop.VGCore.Shape


    Private Sub ckbxCLRgroms_Checked(sender As Object, e As RoutedEventArgs) Handles ckbxCLRgroms.Checked, ckbxCLRgroms.Unchecked
        If ckbxCLRgroms.IsChecked Then
            ckbxLRspacing.IsEnabled = False
            ckbxLRspacing.IsChecked = False
        End If
        If ckbxCLRgroms.IsChecked = False Then
            ckbxLRspacing.IsEnabled = True
        End If
    End Sub

    Private Sub btnXML_Click(sender As Object, e As RoutedEventArgs) Handles btnXML.Click
        Dim openFileDialog As New Microsoft.Win32.OpenFileDialog()
        Dim zchFile As New XDocument

        ' Set properties of the dialog
        openFileDialog.Filter = "ZCH Files (*.zch)|*.zch|All Files (*.*)|*.*"
        openFileDialog.Title = "Select ZCH File"

        ' Show the dialog and get the selected file path
        If openFileDialog.ShowDialog() = True Then
            Dim selectedFilePath As String = openFileDialog.FileName
            zchFile = XDocument.Load(selectedFilePath)
        Else
            Return
        End If

        Dim orderNode As XElement = zchFile.Element("order")
        Dim id As String = orderNode.Attribute("id").Value
        Dim material As Integer = Integer.Parse(orderNode.Element("material").Value)
        Dim height As Single = Single.Parse(orderNode.Element("height").Value)
        Dim width As Single = Single.Parse(orderNode.Element("width").Value)
        Dim flutes As Integer = Integer.Parse(orderNode.Element("flutes").Value)
        Dim stakes As Integer = Integer.Parse(orderNode.Element("stakes").Value)
        Dim corners As Integer = Integer.Parse(orderNode.Element("corners").Value)
        Dim holes As Integer = Integer.Parse(orderNode.Element("holes").Value)
        Dim holesize As Integer = Integer.Parse(orderNode.Element("holesize").Value)
        Dim holedistvert As Single = Single.Parse(orderNode.Element("holedistvert").Value)
        Dim holedisthorz As Single = Single.Parse(orderNode.Element("holedisthorz").Value)
        Dim holeul As Boolean = Boolean.Parse(orderNode.Element("holeul").Value)
        Dim holeuc As Boolean = Boolean.Parse(orderNode.Element("holeuc").Value)
        Dim holeur As Boolean = Boolean.Parse(orderNode.Element("holeur").Value)
        Dim holecl As Boolean = Boolean.Parse(orderNode.Element("holecl").Value)
        Dim holecr As Boolean = Boolean.Parse(orderNode.Element("holecr").Value)
        Dim holell As Boolean = Boolean.Parse(orderNode.Element("holell").Value)
        Dim holelc As Boolean = Boolean.Parse(orderNode.Element("holelc").Value)
        Dim holelr As Boolean = Boolean.Parse(orderNode.Element("holelr").Value)
        Dim gromcorn As Boolean = Boolean.Parse(orderNode.Element("gromcorn").Value)
        Dim gromtbc As Boolean = Boolean.Parse(orderNode.Element("gromtbc").Value)
        Dim gromlrc As Boolean = Boolean.Parse(orderNode.Element("gromlrc").Value)
        Dim gromtb As Boolean = Boolean.Parse(orderNode.Element("gromtb").Value)
        Dim gromtbdist As Single = Single.Parse(orderNode.Element("gromtbdist").Value)
        Dim gromlr As Boolean = Boolean.Parse(orderNode.Element("gromlr").Value)
        Dim gromlrdist As Single = Single.Parse(orderNode.Element("gromlrdist").Value)

        lstMaterial.SelectedIndex = material
        txtWidth.Text = width
        txtHeight.Text = height
        lstFlutes.SelectedIndex = flutes
        lstStkDots.SelectedIndex = stakes
        lstCorners.SelectedIndex = corners
        lstHoleSz.SelectedIndex = holesize
        txtLRDist.Text = holedisthorz
        txtTBDist.Text = holedistvert
        ckbxUL.IsChecked = holeul
        ckbxUC.IsChecked = holeuc
        ckbxUR.IsChecked = holeur
        ckbxCL.IsChecked = holecl
        ckbxCR.IsChecked = holecr
        ckbxLL.IsChecked = holell
        ckbxLC.IsChecked = holelc
        ckbxLR.IsChecked = holelr
        ckbxCornerGroms.IsChecked = gromcorn
        ckbxCTBGroms.IsChecked = gromtbc
        ckbxCLRgroms.IsChecked = gromlrc
        ckbxTBspacing.IsChecked = gromtb
        ckbxLRspacing.IsChecked = gromlr
        txtTBspacing.Text = gromtbdist
        txtLRspacing.Text = gromlrdist

    End Sub
    Private Sub txtTBDist_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtTBDist.GotFocus
        txtTBDist.SelectAll()
    End Sub
    Private Sub txtLRDist_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtLRDist.GotFocus
        txtLRDist.SelectAll()
    End Sub

    Private Sub txtHeight_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtHeight.GotFocus
        txtHeight.SelectAll()
    End Sub

    Private Sub txtWidth_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtWidth.GotFocus
        txtWidth.SelectAll()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As RoutedEventArgs) Handles btnClear.Click
        lstMaterial.SelectedIndex = -1
        lblFlutes.IsEnabled = False
        lstFlutes.IsEnabled = False
        lblStkDots.IsEnabled = False
        lstStkDots.IsEnabled = False
        lblCorners.IsEnabled = False
        lstCorners.IsEnabled = False
        lblHoles.IsEnabled = False
        lstHoleSz.IsEnabled = False
        lstCorners.SelectedIndex = -1
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
        txtWidth.IsEnabled = False
        txtWidth.Clear()
        txtHeight.IsEnabled = False
        txtHeight.Clear()
        ckbxCLRgroms.IsEnabled = False
        ckbxCLRgroms.IsChecked = False
        ckbxCTBGroms.IsEnabled = False
        ckbxCTBGroms.IsChecked = False
        ckbxCornerGroms.IsEnabled = False
        ckbxCornerGroms.IsChecked = False
        ckbxTBspacing.IsEnabled = False
        ckbxTBspacing.IsChecked = False
        ckbxLRspacing.IsEnabled = False
        ckbxLRspacing.IsChecked = False
        txtTBspacing.IsEnabled = False
        txtTBspacing.Clear()
        txtLRspacing.IsEnabled = False
        txtLRspacing.Clear()
        lstFlutes.SelectedIndex = -1
        lstStkDots.SelectedIndex = -1
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

    Private Sub lstFlutes_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lstFlutes.SelectionChanged
        If lstFlutes.SelectedIndex = 1 Then
            lblStkDots.IsEnabled = False
            lstStkDots.IsEnabled = False 'Removes stake dot option for horizontal flutes
            lstStkDots.SelectedIndex = -1
        Else
            lblStkDots.IsEnabled = True
            lstStkDots.IsEnabled = True
        End If

    End Sub

    Public Sub lstHoleSz_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lstHoleSz.SelectionChanged
        If lstHoleSz.SelectedIndex <> -1 Then
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
            corelApp = CType(CreateObject("CorelDRAW.Application"), Corel.Interop.VGCore.Application)
        Else
            corelApp = CType(GetObject(, "CorelDRAW.Application"), Corel.Interop.VGCore.Application)
        End If



    End Sub

    Public Sub lstMaterial_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lstMaterial.SelectionChanged
        If lstMaterial.SelectedIndex = -1 Then
            lblFlutes.IsEnabled = False
            lstFlutes.IsEnabled = False
            lblStkDots.IsEnabled = False
            lstStkDots.IsEnabled = False
            lblCorners.IsEnabled = False
            lstCorners.IsEnabled = False
            lblHoles.IsEnabled = False
            lstCorners.SelectedIndex = -1
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
            txtWidth.IsEnabled = False
            txtWidth.Clear()
            txtHeight.IsEnabled = False
            txtHeight.Clear()
            ckbxCLRgroms.IsEnabled = False
            ckbxCLRgroms.IsChecked = False
            ckbxCTBGroms.IsEnabled = False
            ckbxCTBGroms.IsChecked = False
            ckbxCornerGroms.IsEnabled = False
            ckbxCornerGroms.IsChecked = False
            ckbxTBspacing.IsEnabled = False
            ckbxTBspacing.IsChecked = False
            ckbxLRspacing.IsEnabled = False
            ckbxLRspacing.IsChecked = False
            txtTBspacing.IsEnabled = False
            txtTBspacing.Clear()
            txtLRspacing.IsEnabled = False
            txtLRspacing.Clear()
            lstFlutes.SelectedIndex = -1
            lstStkDots.SelectedIndex = -1
        End If
        If lstMaterial.SelectedIndex = 0 Then 'Coroplast
            txtHeight.IsEnabled = True
            txtWidth.IsEnabled = True
            lblFlutes.IsEnabled = True
            lstFlutes.IsEnabled = True
            lblStkDots.IsEnabled = True
            lstStkDots.IsEnabled = True
            lblCorners.IsEnabled = False
            lstCorners.IsEnabled = False
            lblHoles.IsEnabled = False
            lstCorners.SelectedIndex = -1
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
            ckbxCLRgroms.IsEnabled = True
            ckbxCTBGroms.IsEnabled = True
            ckbxCornerGroms.IsEnabled = True
            ckbxTBspacing.IsEnabled = True
            ckbxLRspacing.IsEnabled = True
            txtTBspacing.IsEnabled = True
            txtLRspacing.IsEnabled = True
        ElseIf lstMaterial.SelectedIndex = 1 Then 'Aluminum
            txtHeight.IsEnabled = True
            txtWidth.IsEnabled = True
            lblCorners.IsEnabled = True
            lstCorners.IsEnabled = True
            lblHoles.IsEnabled = True
            lstHoleSz.IsEnabled = True
            lblFlutes.IsEnabled = False
            lstFlutes.IsEnabled = False
            lstFlutes.SelectedIndex = -1
            lblStkDots.IsEnabled = False
            lstStkDots.IsEnabled = False
            lstStkDots.SelectedIndex = -1
            ckbxCLRgroms.IsEnabled = False
            ckbxCLRgroms.IsChecked = False
            ckbxCTBGroms.IsEnabled = False
            ckbxCTBGroms.IsChecked = False
            ckbxCornerGroms.IsEnabled = False
            ckbxCornerGroms.IsChecked = False
            ckbxTBspacing.IsEnabled = False
            ckbxTBspacing.IsChecked = False
            ckbxLRspacing.IsEnabled = False
            ckbxLRspacing.IsChecked = False
            txtTBspacing.IsEnabled = False
            txtTBspacing.Clear()
            txtLRspacing.IsEnabled = False
            txtLRspacing.Clear()
        ElseIf lstMaterial.SelectedIndex = 2 Or lstMaterial.SelectedIndex = 3 Then 'ACM or Plastic
            txtHeight.IsEnabled = True
            txtWidth.IsEnabled = True
            lblCorners.IsEnabled = True
            lstCorners.IsEnabled = True
            lblHoles.IsEnabled = True
            lstHoleSz.IsEnabled = True
            lblFlutes.IsEnabled = False
            lstFlutes.IsEnabled = False
            lstFlutes.SelectedIndex = -1
            lblStkDots.IsEnabled = False
            lstStkDots.IsEnabled = False
            lstStkDots.SelectedIndex = -1
            ckbxCLRgroms.IsEnabled = True
            ckbxCTBGroms.IsEnabled = True
            ckbxCornerGroms.IsEnabled = True
            ckbxTBspacing.IsEnabled = True
            ckbxLRspacing.IsEnabled = True
            txtTBspacing.IsEnabled = True
            txtLRspacing.IsEnabled = True
        ElseIf lstMaterial.SelectedIndex = 4 Then 'Vinyl
            txtHeight.IsEnabled = True
            txtWidth.IsEnabled = True
            lstFlutes.SelectedIndex = -1
            lstStkDots.SelectedIndex = -1
            lblFlutes.IsEnabled = False
            lstFlutes.IsEnabled = False
            lblStkDots.IsEnabled = False
            lstStkDots.IsEnabled = False
            lblCorners.IsEnabled = False
            lstCorners.IsEnabled = False
            lblHoles.IsEnabled = False
            lstHoleSz.IsEnabled = False
            lstCorners.SelectedIndex = -1
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
            ckbxCLRgroms.IsEnabled = False
            ckbxCLRgroms.IsChecked = False
            ckbxCTBGroms.IsEnabled = False
            ckbxCTBGroms.IsChecked = False
            ckbxCornerGroms.IsEnabled = False
            ckbxCornerGroms.IsChecked = False
            ckbxTBspacing.IsEnabled = False
            ckbxTBspacing.IsChecked = False
            ckbxLRspacing.IsEnabled = False
            ckbxLRspacing.IsChecked = False
            txtTBspacing.IsEnabled = False
            txtTBspacing.Clear()
            txtLRspacing.IsEnabled = False
            txtLRspacing.Clear()

        ElseIf lstMaterial.SelectedIndex = 5 Then 'Banner
            txtHeight.IsEnabled = True
            txtWidth.IsEnabled = True
            lstFlutes.SelectedIndex = -1
            lstStkDots.SelectedIndex = -1
            lblFlutes.IsEnabled = False
            lstFlutes.IsEnabled = False
            lblStkDots.IsEnabled = False
            lstStkDots.IsEnabled = False
            lblCorners.IsEnabled = False
            lstCorners.IsEnabled = False
            lblHoles.IsEnabled = False
            lstHoleSz.IsEnabled = False
            lstCorners.SelectedIndex = -1
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
            txtHeight.IsEnabled = True
            txtWidth.IsEnabled = True
            ckbxCLRgroms.IsEnabled = True
            ckbxCTBGroms.IsEnabled = True
            ckbxCornerGroms.IsEnabled = True
            ckbxTBspacing.IsEnabled = True
            ckbxLRspacing.IsEnabled = True
            txtTBspacing.IsEnabled = True
            txtLRspacing.IsEnabled = True

        ElseIf lstMaterial.SelectedIndex = 6 Then 'Poster
            txtHeight.IsEnabled = True
            txtWidth.IsEnabled = True
            lstFlutes.SelectedIndex = -1
            lstStkDots.SelectedIndex = -1
            lblFlutes.IsEnabled = False
            lstFlutes.IsEnabled = False
            lblStkDots.IsEnabled = False
            lstStkDots.IsEnabled = False
            lblCorners.IsEnabled = False
            lstCorners.IsEnabled = False
            lblHoles.IsEnabled = False
            lstHoleSz.IsEnabled = False
            lstCorners.SelectedIndex = -1
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
            ckbxCLRgroms.IsEnabled = False
            ckbxCLRgroms.IsChecked = False
            ckbxCTBGroms.IsEnabled = False
            ckbxCTBGroms.IsChecked = False
            ckbxCornerGroms.IsEnabled = False
            ckbxCornerGroms.IsChecked = False
            ckbxTBspacing.IsEnabled = False
            ckbxTBspacing.IsChecked = False
            ckbxLRspacing.IsEnabled = False
            ckbxLRspacing.IsChecked = False
            txtTBspacing.IsEnabled = False
            txtTBspacing.Clear()
            txtLRspacing.IsEnabled = False
            txtLRspacing.Clear()
        ElseIf lstMaterial.SelectedIndex = 7 Then 'Tags
            lblFlutes.IsEnabled = False
            lstFlutes.IsEnabled = False
            lblStkDots.IsEnabled = False
            lstStkDots.IsEnabled = False
            lblCorners.IsEnabled = False
            lstCorners.IsEnabled = False
            lblHoles.IsEnabled = False
            lstHoleSz.IsEnabled = False
            lstCorners.SelectedIndex = -1
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
            txtWidth.IsEnabled = False
            txtWidth.Text = 12
            txtHeight.IsEnabled = False
            txtHeight.Text = 6
            ckbxCLRgroms.IsEnabled = False
            ckbxCLRgroms.IsChecked = False
            ckbxCTBGroms.IsEnabled = False
            ckbxCTBGroms.IsChecked = False
            ckbxCornerGroms.IsEnabled = False
            ckbxCornerGroms.IsChecked = False
            ckbxTBspacing.IsEnabled = False
            ckbxTBspacing.IsChecked = False
            ckbxLRspacing.IsEnabled = False
            ckbxLRspacing.IsChecked = False
            txtTBspacing.IsEnabled = False
            txtTBspacing.Clear()
            txtLRspacing.IsEnabled = False
            txtLRspacing.Clear()
            lstFlutes.SelectedIndex = -1
            lstStkDots.SelectedIndex = -1

        End If

    End Sub

    Public Sub BtnGo_Click(sender As Object, e As RoutedEventArgs) Handles btnGo.Click

        corelApp.Visible = True



        Dim pgHeight As Single = Val(txtHeight.Text)
        Dim pgWidth As Single = Val(txtWidth.Text)
        Dim valTBspacing As Single = Val(txtTBspacing.Text)
        Dim valLRspacing As Single = Val(txtLRspacing.Text)
        Dim pgDimsRange As Corel.Interop.VGCore.IVGShapeRange, pgDimsShape As Corel.Interop.VGCore.Shape


        If ckbxTBspacing.IsChecked And Not Integer.TryParse(pgWidth / valTBspacing, 0) Then
            MsgBox("Width must be evenly divisible by grommet spacing distance.", , Title:="Error!")
            txtTBspacing.Clear()
            Exit Sub
        End If

        If ckbxLRspacing.IsChecked And Not Integer.TryParse(pgHeight / valLRspacing, 0) Then
            MsgBox("Height must be evenly divisible by grommet spacing distance.", , Title:="Error!")
            txtLRspacing.Clear()
            Exit Sub
        End If

        If lstMaterial.SelectedIndex = 0 And lstFlutes.SelectedIndex = -1 Then
            MsgBox("Must choose flute direction on coroplast!", , Title:="Error!")
            Exit Sub
        End If


        If pgHeight = 0 Or pgWidth = 0 Then
            MsgBox("Invalid page size.", , Title:="Error!")
            Exit Sub
        End If

        Dim appDirectory As String = AppDomain.CurrentDomain.BaseDirectory
        Dim templateFilePath As String = appDirectory & "TestTemplate.cdt"
        corelDoc = corelApp.CreateDocumentFromTemplate(templateFilePath)
        'corelDoc.Activate()
        Dim regDots18x24 As ShapeRange = corelDoc.ActiveLayer.FindShapes("regDots18x24")
        Dim cutLines18x24 As ShapeRange = corelDoc.ActiveLayer.FindShapes("cutLines18x24")
        Dim lyrOne18x24 As ShapeRange = corelDoc.ActiveLayer.FindShapes("lyrOne18x24")
        Dim regmark As Layer = corelDoc.ActivePage.AllLayers.Find("Regmark")
        Dim thrucut As Layer = corelDoc.ActivePage.AllLayers.Find("Through Cut")
        Dim layer1 As Layer = corelDoc.ActivePage.AllLayers.Find("Layer 1")


        'Sets the page size
        corelDoc.ActivePage.SizeWidth = pgWidth
        corelDoc.ActivePage.SizeHeight = pgHeight

        'Sets the control rectangle size & position
        ctrlRectangle = corelDoc.ActivePage.Shapes("ctrlRectangle")
        ctrlRectangle.SizeHeight = pgHeight
        ctrlRectangle.SizeWidth = pgWidth
        ctrlRectangle.SetPosition(0, pgHeight)

        pgDimsRange = corelDoc.ActiveLayer.Shapes.All
        For Each pgDimsShape In pgDimsRange
            If pgDimsShape.Type = cdrShapeType.cdrLinearDimensionShape Then
                pgDimsShape.Dimension.TextShape.Text.Story.Size = 1.8 * ((pgWidth + pgHeight) / 2) + 34
            End If
        Next

        If lstMaterial.SelectedIndex = 0 And pgHeight = 18 And pgWidth = 24 Then
            regDots18x24.MoveToLayer(regmark)
            regDots18x24.Ungroup()
            cutLines18x24.MoveToLayer(thrucut)
            cutLines18x24.Ungroup()
            lyrOne18x24.Ungroup()
        ElseIf lstMaterial.SelectedIndex = 0 And pgHeight = 24 And pgWidth = 18 Then
            regDots18x24.MoveToLayer(regmark)
                regDots18x24.Ungroup()
                cutLines18x24.MoveToLayer(thrucut)
                cutLines18x24.Ungroup()
                lyrOne18x24.Ungroup()
            Else
            regDots18x24.Delete()
            cutLines18x24.Delete()
            lyrOne18x24.Delete()
        End If


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
        ElseIf lstStkDots.SelectedIndex = 1 Then '6x24
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

        Dim radTxt, holeSzTxt, holeLRDistTxt, holeTBDistTxt, holePlcTxt As String
        Dim tbDist As Double = Val(txtTBDist.Text)
        Dim lrDist As Double = Val(txtLRDist.Text)
        If tbDist = Nothing Or tbDist = 0 Then
            If lrDist = Nothing Or lrDist = 0 Then
                MsgBox("Distance from edge cannot be 0.", , Title:="Error!")
                Exit Sub
            Else
                tbDist = lrDist
            End If
        End If
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


            If lstCorners.SelectedIndex <> -1 And lstCorners.SelectedIndex <> 0 Then

                radTxt = lstCorners.SelectionBoxItem.ToString + " radius corners"

            Else
                radTxt = ""

            End If
            If lstHoleSz.SelectedIndex <> -1 Then

                holeSzTxt = lstHoleSz.SelectionBoxItem.ToString + " holes"

                If tbDist = lrDist Then
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
                    holeSzTxt = holeSzTxt + ", " + holePlcTxt + "'' from edge" + vbCrLf
                Else
                    If tbDist = 0.25 Then
                        holeTBDistTxt = "1/4'' from top and bottom edge" + vbCrLf
                    ElseIf tbDist = 0.375 Then
                        holeTBDistTxt = "3/8'' from top and bottom edge" + vbCrLf
                    ElseIf tbDist = 0.5 Then
                        holeTBDistTxt = "1/2'' from top and bottom edge" + vbCrLf
                    ElseIf tbDist = 0.625 Then
                        holeTBDistTxt = "5/8'' from top and bottom edge" + vbCrLf
                    ElseIf tbDist = 0.75 Then
                        holeTBDistTxt = "3/4'' from top and bottom edge" + vbCrLf
                    ElseIf txtTBDist Is Nothing Or tbDist = 0 Then
                        holeTBDistTxt = ""
                    Else
                        holeTBDistTxt = txtTBDist.Text + "'' from top and bottom edge" + vbCrLf
                    End If
                    If lrDist = 0.25 Then
                        holeLRDistTxt = "1/4'' from left and right edge" + vbCrLf
                    ElseIf lrDist = 0.375 Then
                        holeLRDistTxt = "3/8'' from left and right edge" + vbCrLf
                    ElseIf lrDist = 0.5 Then
                        holeLRDistTxt = "1/2'' from left and right edge" + vbCrLf
                    ElseIf lrDist = 0.625 Then
                        holeLRDistTxt = "5/8'' from left and right edge" + vbCrLf
                    ElseIf lrDist = 0.75 Then
                        holeLRDistTxt = "3/4'' from left and right edge" + vbCrLf
                    ElseIf txtLRDist Is Nothing Or lrDist = 0 Then
                        holeLRDistTxt = ""
                    Else
                        holeLRDistTxt = txtLRDist.Text + "'' from left and right edge" + vbCrLf
                    End If
                    holeSzTxt = holeSzTxt + vbCrLf + holeTBDistTxt + holeLRDistTxt

                End If
                If ckbxUC.IsChecked = True And ckbxLC.IsChecked = True Then
                    holeSzTxt = holeSzTxt + "in center at top & bottom" + vbCrLf
                ElseIf ckbxUC.IsChecked = True And ckbxLC.IsChecked = False Then
                    holeSzTxt = holeSzTxt + "in center at top" + vbCrLf
                End If
                If ckbxCL.IsChecked = True And ckbxCR.IsChecked = True Then
                    holeSzTxt = holeSzTxt + "in center at left & right" + vbCrLf
                End If
                If ckbxUL.IsChecked = True And ckbxUR.IsChecked = True And ckbxLL.IsChecked = False And ckbxLR.IsChecked = False Then
                    holeSzTxt = holeSzTxt + "in top two corners" + vbCrLf
                End If
                If ckbxUL.IsChecked = True And ckbxUR.IsChecked = True And ckbxLL.IsChecked = True And ckbxLR.IsChecked = True Then
                    holeSzTxt = holeSzTxt + "one in each corner" + vbCrLf
                End If
            Else
                    holeSzTxt = ""
            End If

            Dim holesAndCornersTxt As Corel.Interop.VGCore.Shape = corelDoc.ActiveLayer.CreateArtisticText(0, 0, radTxt + vbCrLf + holeSzTxt, , ,
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

        'grommet & banner text
        Dim bannerSizeTxt, grommetTxt As String
        Dim bannerArtTxt, grommetArtTxt As Corel.Interop.VGCore.Shape

        grommetTxt = "3/8'' Brass Grommets "

        If lstMaterial.SelectedIndex = 5 Then
            bannerSizeTxt = "Banner Finish Size is " + txtHeight.Text + "''x" + txtWidth.Text + "''"
            bannerArtTxt = corelDoc.ActiveLayer.CreateArtisticText(0, 0, bannerSizeTxt, , ,
                                                    "Arial", 3 * ((pgWidth + pgHeight) / 2) + 34, , , , Corel.Interop.VGCore.cdrAlignment.cdrLeftAlignment)
            bannerArtTxt.SetPosition(0, -bannerArtTxt.SizeHeight - 1)
            If ckbxCornerGroms.IsChecked Then
                grommetTxt += "- One in each corner"
            End If
            If ckbxCTBGroms.IsChecked Then
                grommetTxt += ", in center of top & bottom"
            End If
            If ckbxCLRgroms.IsChecked Then
                grommetTxt += ", in center of left & right"
            End If
            If ckbxTBspacing.IsChecked Then
                grommetTxt += ", every " + txtTBspacing.Text + "'' along top & bottom"
            End If
            If ckbxLRspacing.IsChecked Then
                grommetTxt += ", every " + txtLRspacing.Text + "'' along left & right"
            End If
            grommetArtTxt = corelDoc.ActiveLayer.CreateArtisticText(0, 0, grommetTxt, , ,
                                                    "Arial", 1.5 * ((pgWidth + pgHeight) / 2) + 34, , , , Corel.Interop.VGCore.cdrAlignment.cdrLeftAlignment)
            grommetArtTxt.SetPosition(0, -3 * bannerArtTxt.SizeHeight)
        End If

        'Tags
        tagBorder = corelDoc.ActiveLayer.Shapes.FindShape("tagBorder")
        tagHoles = corelDoc.ActiveLayer.Shapes.FindShape("tagHoles")

        If lstMaterial.SelectedIndex = 7 Then
            tagBorder.SetPosition(0, 6)
            tagHoles.SetPosition(1.85, 5.5)
            ctrlRectangle.Outline.Width = 0
        Else
            tagBorder.Delete()
            tagHoles.Delete()
        End If

        'Sets registration dot position

        regDot = corelDoc.ActivePage.Shapes("RegDot")
        regDot.SetPosition(0 - 6, pgHeight + 6)


        'Change view to fit everything on Layer 1
        corelApp.ActiveWindow.ActiveView.ToFitPage()
        corelDoc.ClearSelection()

    End Sub
End Class
