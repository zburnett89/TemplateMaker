Imports System.Math
Imports Corel.Interop.VGCore

Class MainWindow

    Public corelApp As Corel.Interop.VGCore.Application
    Public corelDoc As Corel.Interop.VGCore.Document

    Dim ctrlRectangle, cornerRect, vertFlutes,
            horzFlutes, stkDot6x24, stkDot10x30, grommet, tagBorder, tagHoles As Corel.Interop.VGCore.Shape
    Dim pSizeA, pSizeB As Double

    'Private Sub btnXML_Click(sender As Object, e As RoutedEventArgs) Handles btnXML.Click
    '    Dim openFileDialog As New Microsoft.Win32.OpenFileDialog()
    '    Dim zchFile As New XDocument

    '    ' Set properties of the dialog
    '    openFileDialog.Filter = "ZCH Files (*.zch)|*.zch|All Files (*.*)|*.*"
    '    openFileDialog.Title = "Select ZCH File"

    '    ' Show the dialog and get the selected file path
    '    If openFileDialog.ShowDialog() = True Then
    '        Dim selectedFilePath As String = openFileDialog.FileName
    '        zchFile = XDocument.Load(selectedFilePath)
    '    Else
    '        Return
    '    End If

    '    Dim orderNode As XElement = zchFile.Element("order")
    '    Dim id As String = orderNode.Attribute("id").Value
    '    Dim material As Integer = Integer.Parse(orderNode.Element("material").Value)
    '    Dim height As Single = Single.Parse(orderNode.Element("height").Value)
    '    Dim width As Single = Single.Parse(orderNode.Element("width").Value)
    '    Dim flutes As Integer = Integer.Parse(orderNode.Element("flutes").Value)
    '    Dim stakes As Integer = Integer.Parse(orderNode.Element("stakes").Value)
    '    Dim corners As Integer = Integer.Parse(orderNode.Element("corners").Value)
    '    Dim holes As Integer = Integer.Parse(orderNode.Element("holes").Value)
    '    Dim holesize As Integer = Integer.Parse(orderNode.Element("holesize").Value)
    '    Dim holedistvert As Single = Single.Parse(orderNode.Element("holedistvert").Value)
    '    Dim holedisthorz As Single = Single.Parse(orderNode.Element("holedisthorz").Value)
    '    Dim holeul As Boolean = Boolean.Parse(orderNode.Element("holeul").Value)
    '    Dim holeuc As Boolean = Boolean.Parse(orderNode.Element("holeuc").Value)
    '    Dim holeur As Boolean = Boolean.Parse(orderNode.Element("holeur").Value)
    '    Dim holecl As Boolean = Boolean.Parse(orderNode.Element("holecl").Value)
    '    Dim holecr As Boolean = Boolean.Parse(orderNode.Element("holecr").Value)
    '    Dim holell As Boolean = Boolean.Parse(orderNode.Element("holell").Value)
    '    Dim holelc As Boolean = Boolean.Parse(orderNode.Element("holelc").Value)
    '    Dim holelr As Boolean = Boolean.Parse(orderNode.Element("holelr").Value)
    '    Dim gromcorn As Boolean = Boolean.Parse(orderNode.Element("gromcorn").Value)
    '    Dim gromtbc As Boolean = Boolean.Parse(orderNode.Element("gromtbc").Value)
    '    Dim gromlrc As Boolean = Boolean.Parse(orderNode.Element("gromlrc").Value)
    '    Dim gromtb As Boolean = Boolean.Parse(orderNode.Element("gromtb").Value)
    '    Dim gromtbdist As Single = Single.Parse(orderNode.Element("gromtbdist").Value)
    '    Dim gromlr As Boolean = Boolean.Parse(orderNode.Element("gromlr").Value)
    '    Dim gromlrdist As Single = Single.Parse(orderNode.Element("gromlrdist").Value)

    '    lstMaterial.SelectedIndex = material
    '    txtWidth.Text = width
    '    txtHeight.Text = height
    '    lstFlutes.SelectedIndex = flutes
    '    lstStkDots.SelectedIndex = stakes
    '    lstCorners.SelectedIndex = corners
    '    lstHoleSz.SelectedIndex = holesize
    '    txtLRDist.Text = holedisthorz
    '    txtTBDist.Text = holedistvert
    '    ckbxUL.IsChecked = holeul
    '    ckbxUC.IsChecked = holeuc
    '    ckbxUR.IsChecked = holeur
    '    ckbxCL.IsChecked = holecl
    '    ckbxCR.IsChecked = holecr
    '    ckbxLL.IsChecked = holell
    '    ckbxLC.IsChecked = holelc
    '    ckbxLR.IsChecked = holelr
    '    ckbxCornerGroms.IsChecked = gromcorn
    '    ckbxCTBGroms.IsChecked = gromtbc
    '    ckbxCLRgroms.IsChecked = gromlrc
    '    ckbxTBspacing.IsChecked = gromtb
    '    ckbxLRspacing.IsChecked = gromlr
    '    txtTBspacing.Text = gromtbdist
    '    txtLRspacing.Text = gromlrdist

    'End Sub 'uploading an xml file
    Private Sub txtTDist_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtTDist.GotFocus
        txtTDist.SelectAll()
    End Sub
    Private Sub txtBDist_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtBDist.GotFocus
        txtBDist.SelectAll()
    End Sub
    Private Sub txtLDist_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtLDist.GotFocus
        txtLDist.SelectAll()
    End Sub
    Private Sub txtRDist_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtRDist.GotFocus
        txtRDist.SelectAll()
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
        txtLDist.IsEnabled = False
        txtLDist.Clear()
        txtRDist.IsEnabled = False
        txtRDist.Clear()
        txtTDist.IsEnabled = False
        txtTDist.Clear()
        txtBDist.IsEnabled = False
        txtBDist.Clear()
        btnTBLink.IsEnabled = False
        btnLRLink.IsEnabled = False
        buttonTextBlock.Text = ChrW(&HE72E).ToString()
        buttonTextBlock1.Text = ChrW(&HE72E).ToString()
        ckbxUL.IsChecked = False
        ckbxUC.IsChecked = False
        ckbxUR.IsChecked = False
        ckbxCL.IsChecked = False
        ckbxCR.IsChecked = False
        ckbxLL.IsChecked = False
        ckbxLC.IsChecked = False
        ckbxLR.IsChecked = False
        txtWidth.IsEnabled = False
        txtHeight.IsEnabled = False
        pSizeA = Nothing
        pSizeB = Nothing
        txtWidth.Clear()
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
        ckbxTBQty.IsEnabled = False
        ckbxTBQty.IsChecked = False
        ckbxLRQty.IsEnabled = False
        ckbxLRQty.IsChecked = False
        txtTBQty.IsEnabled = False
        txtTBQty.Clear()
        txtLRQty.IsEnabled = False
        txtLRQty.Clear()
        txtTBspacing.IsEnabled = False
        txtTBspacing.Clear()
        txtLRspacing.IsEnabled = False
        txtLRspacing.Clear()
        lstFlutes.SelectedIndex = -1
        lstStkDots.SelectedIndex = -1
        lstPresets.SelectedIndex = -1
    End Sub 'clear button

    Private Sub ckbxTBQty_Checked(sender As Object, e As RoutedEventArgs) Handles ckbxTBQty.Click
        If ckbxTBQty.IsChecked = True Then
            ckbxTBspacing.IsEnabled = False
            ckbxTBspacing.IsChecked = False
            ckbxCTBGroms.IsEnabled = False
            ckbxCTBGroms.IsChecked = False
            txtTBspacing.IsEnabled = False
            txtTBspacing.Clear()
            ckbxCornerGroms.IsChecked = True
            txtTBQty.IsEnabled = True
        End If
        If ckbxTBQty.IsChecked = False Then
            ckbxTBspacing.IsEnabled = True
            ckbxCTBGroms.IsEnabled = True
            txtTBspacing.IsEnabled = True
            txtTBQty.Clear()
            txtTBQty.IsEnabled = False
        End If
    End Sub 'lr quantity grommet checkbox click

    Private Sub ckbxLRQty_Checked(sender As Object, e As RoutedEventArgs) Handles ckbxLRQty.Click
        If ckbxLRQty.IsChecked = True Then
            ckbxCornerGroms.IsChecked = True
            ckbxCLRgroms.IsEnabled = False
            ckbxCLRgroms.IsChecked = False
            ckbxLRspacing.IsEnabled = False
            ckbxLRspacing.IsChecked = False
            txtLRspacing.IsEnabled = False
            txtLRspacing.Clear()
            txtLRQty.IsEnabled = True
        End If
        If ckbxLRQty.IsChecked = False Then
            ckbxCLRgroms.IsEnabled = True
            ckbxLRspacing.IsEnabled = True
            txtLRspacing.IsEnabled = True
            txtLRQty.Clear()
            txtLRQty.IsEnabled = False
        End If
    End Sub 'lr quantity grommet checkbox click

    Private Sub ckbxCTBGroms_Checked(sender As Object, e As RoutedEventArgs) Handles ckbxCTBGroms.Click
        If ckbxCTBGroms.IsChecked = True Then
            ckbxTBspacing.IsEnabled = False
            ckbxTBspacing.IsChecked = False
            ckbxTBQty.IsEnabled = False
            ckbxTBQty.IsChecked = False
            txtTBQty.IsEnabled = False
            txtTBspacing.IsEnabled = False
            txtTBspacing.Clear()
            txtTBQty.Clear()
        End If
        If ckbxCTBGroms.IsChecked = False Then
            ckbxTBspacing.IsEnabled = True
            ckbxTBQty.IsEnabled = True
        End If
    End Sub 'ctb grommet checkbox click
    Private Sub ckbxCLRGroms_Checked(sender As Object, e As RoutedEventArgs) Handles ckbxCLRgroms.Click
        If ckbxCLRgroms.IsChecked = True Then
            ckbxLRspacing.IsEnabled = False
            ckbxLRspacing.IsChecked = False
            ckbxLRQty.IsEnabled = False
            ckbxLRQty.IsChecked = False
            txtLRQty.IsEnabled = False
            txtLRspacing.IsEnabled = False
            txtLRspacing.Clear()
            txtLRQty.Clear()
        End If
        If ckbxCLRgroms.IsChecked = False Then
            ckbxLRspacing.IsEnabled = True
            ckbxLRQty.IsEnabled = True
        End If
    End Sub 'clr grommet checkbox click

    Private Sub txtTBspacing_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtTBspacing.GotFocus
        'ckbxTBspacing.IsChecked = True
        'ckbxCTBGroms.IsChecked = False
        'ckbxCTBGroms.IsEnabled = False
        'ckbxTBQty.IsChecked = False
        'ckbxTBQty.IsEnabled = False
        'txtTBQty.IsEnabled = False
        'txtTBQty.Clear()
        txtTBspacing.SelectAll()
    End Sub 'tb grommet spacing txt gets focus
    Private Sub txtTBQty_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtTBQty.GotFocus
        'ckbxCornerGroms.IsChecked = True
        'ckbxCTBGroms.IsChecked = False
        'ckbxCTBGroms.IsEnabled = False
        'ckbxTBspacing.IsEnabled = False
        'ckbxTBQty.IsChecked = True
        'ckbxTBQty.IsEnabled = True
        txtTBQty.SelectAll()
    End Sub 'tb grommet quantity txt gets focus

    Private Sub txtLRspacing_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtLRspacing.GotFocus
        'ckbxLRspacing.IsChecked = True
        'ckbxCLRgroms.IsChecked = False
        'ckbxCLRgroms.IsEnabled = False
        'ckbxLRQty.IsChecked = False
        'ckbxLRQty.IsEnabled = False
        'txtLRQty.IsEnabled = False
        'txtLRQty.Clear()
        txtLRspacing.SelectAll()
    End Sub 'lr grommet spacing txt gets focus
    Private Sub txtLRQty_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtLRQty.GotFocus
        'ckbxCornerGroms.IsChecked = True
        'ckbxCLRgroms.IsChecked = False
        'ckbxCLRgroms.IsEnabled = False
        'ckbxLRspacing.IsEnabled = False
        'ckbxLRQty.IsChecked = True
        'ckbxLRQty.IsEnabled = True
        txtLRQty.SelectAll()
    End Sub 'lr grommet quantity txt gets focus

    Private Sub btnSwitchSize_Click(sender As Object, e As RoutedEventArgs) Handles btnSwitchSize.Click
        pSizeA = Val(txtHeight.Text)
        pSizeB = Val(txtWidth.Text)
        txtHeight.Text = pSizeB.ToString
        txtWidth.Text = pSizeA.ToString
    End Sub 'switch sizes button

    Private Sub btnTBLink_Click(sender As Object, e As RoutedEventArgs) Handles btnTBLink.Click
        If buttonTextBlock.Text = ChrW(&HE72E).ToString() Then
            buttonTextBlock.Text = ChrW(&HE785).ToString()
            lblHolesBDist.IsEnabled = True
            txtBDist.IsEnabled = True
        Else
            buttonTextBlock.Text = ChrW(&HE72E).ToString()
            txtBDist.Text = txtTDist.Text
            lblHolesBDist.IsEnabled = False
            txtBDist.IsEnabled = False
        End If
    End Sub

    Private Sub txtTDist_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtTDist.TextChanged
        If buttonTextBlock.Text = ChrW(&HE72E) Then
            txtBDist.Text = txtTDist.Text
        End If
    End Sub

    Private Sub btnLRLink_Click(sender As Object, e As RoutedEventArgs) Handles btnLRLink.Click
        If buttonTextBlock1.Text = ChrW(&HE72E).ToString() Then
            buttonTextBlock1.Text = ChrW(&HE785).ToString()
            lblHolesRDist.IsEnabled = True
            txtRDist.IsEnabled = True
        Else
            buttonTextBlock1.Text = ChrW(&HE72E).ToString()
            txtRDist.Text = txtLDist.Text
            lblHolesRDist.IsEnabled = False
            txtRDist.IsEnabled = False
        End If
    End Sub

    Private Sub txtLDist_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtLDist.TextChanged
        If buttonTextBlock1.Text = ChrW(&HE72E) Then
            txtRDist.Text = txtLDist.Text
        End If
    End Sub

    Private Sub ComboBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Select Case lstPresets.SelectedIndex
            Case 1
                pSizeA = 6
                pSizeB = 9
            Case 2
                pSizeA = 6
                pSizeB = 12
            Case 3
                pSizeA = 6
                pSizeB = 24
            Case 4
                pSizeA = 12
                pSizeB = 18
            Case 5
                pSizeA = 12
                pSizeB = 24
            Case 6
                pSizeA = 18
                pSizeB = 24
            Case 7
                pSizeA = 24
                pSizeB = 24
            Case 8
                pSizeA = 24
                pSizeB = 36
            Case 9
                pSizeA = 24
                pSizeB = 48
            Case 10
                pSizeA = 32
                pSizeB = 48
            Case 11
                pSizeA = 48
                pSizeB = 48
            Case 12
                pSizeA = 48
                pSizeB = 96
        End Select
        txtHeight.Text = pSizeA.ToString
        txtWidth.Text = pSizeB.ToString
    End Sub 'preset size dropdown

    Private Sub ckbxLRspacing_Checked(sender As Object, e As RoutedEventArgs) Handles ckbxLRspacing.Click

        If ckbxLRspacing.IsChecked Then
            ckbxCLRgroms.IsEnabled = False
            ckbxCLRgroms.IsChecked = False
            ckbxLRQty.IsChecked = False
            ckbxLRQty.IsEnabled = False
            txtLRQty.Clear()
            txtLRQty.IsEnabled = False
            txtLRspacing.IsEnabled = True
        End If
        If ckbxLRspacing.IsChecked = False Then
            ckbxCLRgroms.IsEnabled = True
            ckbxLRQty.IsEnabled = True
            txtLRQty.IsEnabled = True
            txtLRspacing.Clear()
            txtLRspacing.IsEnabled = False

        End If
    End Sub 'left right grommet spacing

    Private Sub ckbxTBspacing_Checked(sender As Object, e As RoutedEventArgs) Handles ckbxTBspacing.Click

        If ckbxTBspacing.IsChecked Then
            ckbxCTBGroms.IsEnabled = False
            ckbxCTBGroms.IsChecked = False
            ckbxTBQty.IsChecked = False
            ckbxTBQty.IsEnabled = False
            txtTBQty.Clear()
            txtTBQty.IsEnabled = False
            txtTBspacing.IsEnabled = True
        End If
        If ckbxTBspacing.IsChecked = False Then
            ckbxCTBGroms.IsEnabled = True
            ckbxTBQty.IsEnabled = True
            txtTBQty.IsEnabled = True
            txtTBspacing.Clear()
            txtTBspacing.IsEnabled = False
        End If
    End Sub 'top bottom grommet spacing

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
        If lstHoleSz.SelectedIndex <> -1 And lstHoleSz.SelectedIndex <> 0 Then
            lblHoleDistance.IsEnabled = True
            lblHolePosition.IsEnabled = True
            lblHolesTDist.IsEnabled = True
            lblHolesBDist.IsEnabled = True
            txtTDist.IsEnabled = True
            txtBDist.IsEnabled = False
            txtLDist.IsEnabled = True
            txtRDist.IsEnabled = False
            lblHolesLDist.IsEnabled = True
            lblHolesRDist.IsEnabled = True
            btnTBLink.IsEnabled = True
            btnLRLink.IsEnabled = True
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
            lblHolesTDist.IsEnabled = False
            lblHolesBDist.IsEnabled = False
            txtTDist.IsEnabled = False
            txtBDist.IsEnabled = False
            txtLDist.IsEnabled = False
            txtRDist.IsEnabled = False
            lblHolesLDist.IsEnabled = False
            lblHolesRDist.IsEnabled = False
            btnTBLink.IsEnabled = False
            btnLRLink.IsEnabled = False
            buttonTextBlock.Text = ChrW(&HE72E).ToString()
            buttonTextBlock1.Text = ChrW(&HE72E).ToString()
            ckbxUL.IsEnabled = False
            ckbxUC.IsEnabled = False
            ckbxUR.IsEnabled = False
            ckbxCL.IsEnabled = False
            ckbxCR.IsEnabled = False
            ckbxLL.IsEnabled = False
            ckbxLC.IsEnabled = False
            ckbxLR.IsEnabled = False
        End If
    End Sub 'allow hole selection

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs) Handles btnClose.Click
        Close()
    End Sub 'close button

    Public Sub New()

        InitializeComponent()

        If corelApp Is Nothing Then
            corelApp = CType(CreateObject("CorelDRAW.Application"), Corel.Interop.VGCore.Application)
        Else
            corelApp = CType(GetObject(, "CorelDRAW.Application"), Corel.Interop.VGCore.Application)
        End If



    End Sub 'initialize corel

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
            txtRDist.IsEnabled = False
            txtRDist.Clear()
            txtBDist.IsEnabled = False
            txtBDist.Clear()
            txtLDist.IsEnabled = False
            txtLDist.Clear()
            txtTDist.IsEnabled = False
            txtTDist.Clear()
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
            txtTBQty.IsEnabled = False
            txtTBQty.Clear()
            txtLRQty.IsEnabled = False
            txtLRQty.Clear()
            lstFlutes.SelectedIndex = -1
            lstStkDots.SelectedIndex = -1
            lstPresets.IsEnabled = False
            btnSwitchSize.IsEnabled = False

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
            txtLDist.IsEnabled = False
            txtLDist.Clear()
            txtRDist.IsEnabled = False
            txtRDist.Clear()
            txtTDist.IsEnabled = False
            txtTDist.Clear()
            txtBDist.IsEnabled = False
            txtBDist.Clear()
            btnTBLink.IsEnabled = False
            btnLRLink.IsEnabled = False
            buttonTextBlock.Text = ChrW(&HE72E).ToString()
            buttonTextBlock1.Text = ChrW(&HE72E).ToString()
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
            ckbxTBQty.IsEnabled = True
            ckbxLRQty.IsEnabled = True
            lstPresets.IsEnabled = True
            btnSwitchSize.IsEnabled = True
            'txtTBQty.IsEnabled = True
            'txtLRQty.IsEnabled = True
            'txtTBspacing.IsEnabled = True
            'txtLRspacing.IsEnabled = True
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
            ckbxTBQty.IsChecked = False
            ckbxTBQty.IsEnabled = False
            ckbxLRQty.IsChecked = False
            ckbxLRQty.IsEnabled = False
            txtTBQty.IsEnabled = False
            txtLRQty.IsEnabled = False
            txtTBQty.Clear()
            txtLRQty.Clear()
            txtTBspacing.IsEnabled = False
            txtTBspacing.Clear()
            txtLRspacing.IsEnabled = False
            txtLRspacing.Clear()
            lstPresets.IsEnabled = True
            btnSwitchSize.IsEnabled = True
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
            ckbxTBQty.IsEnabled = True
            ckbxLRQty.IsEnabled = True
            lstPresets.IsEnabled = True
            btnSwitchSize.IsEnabled = True
            'txtTBQty.IsEnabled = True
            'txtLRQty.IsEnabled = True
            'txtTBspacing.IsEnabled = True
            'txtLRspacing.IsEnabled = True
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
            txtLDist.IsEnabled = False
            txtLDist.Clear()
            txtRDist.IsEnabled = False
            txtRDist.Clear()
            txtTDist.IsEnabled = False
            txtTDist.Clear()
            txtBDist.IsEnabled = False
            txtBDist.Clear()
            btnTBLink.IsEnabled = False
            btnLRLink.IsEnabled = False
            buttonTextBlock.Text = ChrW(&HE72E).ToString()
            buttonTextBlock1.Text = ChrW(&HE72E).ToString()
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
            lstPresets.IsEnabled = True
            btnSwitchSize.IsEnabled = True
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
            txtLDist.IsEnabled = False
            txtLDist.Clear()
            txtRDist.IsEnabled = False
            txtRDist.Clear()
            txtTDist.IsEnabled = False
            txtTDist.Clear()
            txtBDist.IsEnabled = False
            txtBDist.Clear()
            btnTBLink.IsEnabled = False
            btnLRLink.IsEnabled = False
            buttonTextBlock.Text = ChrW(&HE72E).ToString()
            buttonTextBlock1.Text = ChrW(&HE72E).ToString()
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
            ckbxTBQty.IsEnabled = True
            ckbxLRQty.IsEnabled = True
            lstPresets.IsEnabled = True
            btnSwitchSize.IsEnabled = True
            'txtTBQty.IsEnabled = True
            'txtLRQty.IsEnabled = True
            'txtTBspacing.IsEnabled = True
            'txtLRspacing.IsEnabled = True

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
            txtLDist.IsEnabled = False
            txtLDist.Clear()
            txtRDist.IsEnabled = False
            txtRDist.Clear()
            txtTDist.IsEnabled = False
            txtTDist.Clear()
            txtBDist.IsEnabled = False
            txtBDist.Clear()
            btnTBLink.IsEnabled = False
            btnLRLink.IsEnabled = False
            buttonTextBlock.Text = ChrW(&HE72E).ToString()
            buttonTextBlock1.Text = ChrW(&HE72E).ToString()
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
            ckbxTBQty.IsChecked = False
            ckbxTBQty.IsEnabled = False
            ckbxLRQty.IsChecked = False
            ckbxLRQty.IsEnabled = False
            txtTBQty.IsEnabled = False
            txtLRQty.IsEnabled = False
            txtTBQty.Clear()
            txtLRQty.Clear()
            txtTBspacing.IsEnabled = False
            txtTBspacing.Clear()
            txtLRspacing.IsEnabled = False
            txtLRspacing.Clear()
            lstPresets.IsEnabled = True
            btnSwitchSize.IsEnabled = True

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
            txtLDist.IsEnabled = False
            txtLDist.Clear()
            txtRDist.IsEnabled = False
            txtRDist.Clear()
            txtTDist.IsEnabled = False
            txtTDist.Clear()
            txtBDist.IsEnabled = False
            txtBDist.Clear()
            btnTBLink.IsEnabled = False
            btnLRLink.IsEnabled = False
            buttonTextBlock.Text = ChrW(&HE72E).ToString()
            buttonTextBlock1.Text = ChrW(&HE72E).ToString()
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
            ckbxTBQty.IsChecked = False
            ckbxTBQty.IsEnabled = False
            ckbxLRQty.IsChecked = False
            ckbxLRQty.IsEnabled = False
            txtTBQty.IsEnabled = False
            txtLRQty.IsEnabled = False
            txtTBQty.Clear()
            txtLRQty.Clear()
            txtTBspacing.Clear()
            txtLRspacing.IsEnabled = False
            txtLRspacing.Clear()
            lstFlutes.SelectedIndex = -1
            lstStkDots.SelectedIndex = -1
            lstPresets.IsEnabled = False
            btnSwitchSize.IsEnabled = False

        End If

    End Sub

    Public Sub BtnGo_Click(sender As Object, e As RoutedEventArgs) Handles btnGo.Click

        corelApp.Visible = True



        Dim pgHeight As Single = Val(txtHeight.Text)
        Dim pgWidth As Single = Val(txtWidth.Text)
        Dim valTBspacing As Single = Val(txtTBspacing.Text)
        Dim valLRspacing As Single = Val(txtLRspacing.Text)
        Dim pgDimsRange As Corel.Interop.VGCore.IVGShapeRange, pgDimsShape As Corel.Interop.VGCore.Shape
        Dim tDist As Double = Val(txtTDist.Text)
        Dim bDist As Double = Val(txtBDist.Text)
        Dim lDist As Double = Val(txtLDist.Text)
        Dim rDist As Double = Val(txtRDist.Text)
        Dim tbQty As Single = Val(txtTBQty.Text)
        Dim lrQty As Single = Val(txtLRQty.Text)

        'Error messages

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

        If lstHoleSz.SelectedIndex <> -1 Then
            If ckbxUL.IsChecked = True Or ckbxUR.IsChecked = True Or
                ckbxLL.IsChecked = True Or ckbxLR.IsChecked = True Then
                If tDist = 0 Or lDist = 0 Or bDist = 0 Or rDist = 0 Then
                    MsgBox("Distance from edge cannot be 0.", , Title:="Error!")
                    Exit Sub
                End If
            End If
        End If

        If lstHoleSz.SelectedIndex <> -1 Then
            If tDist > pgHeight / 2 Or bDist > pgHeight / 2 Or lDist > pgWidth / 2 Or rDist > pgWidth / 2 Then
                MsgBox("Invalid hole distance.", , Title:="Error!")
                Exit Sub
            End If
        End If

        If lstHoleSz.SelectedIndex <> -1 Then
            If ckbxUL.IsChecked = False And ckbxUC.IsChecked = False And ckbxUR.IsChecked = False And
                ckbxCL.IsChecked = False And ckbxCR.IsChecked = False And
                ckbxLL.IsChecked = False And ckbxLC.IsChecked = False And ckbxLR.IsChecked = False Then
                MsgBox("Must choose hole position.", , Title:="Error!")
                Exit Sub
            End If
        End If

        If pgHeight = 0 Or pgWidth = 0 Then
            MsgBox("Invalid page size.", , Title:="Error!")
            Exit Sub
        End If

        Select Case lstMaterial.SelectedIndex
            Case 0 ' coroplast
                If Val(txtHeight.Text) > 48 And Val(txtWidth.Text) > 48 Then
                    MsgBox("Invalid page size. Coro sheet size is 48''x96''.", , Title:="Error!")
                    Exit Sub
                End If
                If Val(txtHeight.Text) > 96 Or Val(txtWidth.Text) > 96 Then
                    MsgBox("Invalid page size. Coro sheet size is 48''x96''.", , Title:="Error!")
                    Exit Sub
                End If
                If Val(txtHeight.Text) > 48 Then
                    If lstFlutes.SelectedIndex = 1 Then
                        MsgBox("Flutes must be vertical at this size.", , Title:="Error!")
                        Exit Sub
                    End If
                End If
                If Val(txtWidth.Text) > 48 Then
                    If lstFlutes.SelectedIndex = 0 Then
                        MsgBox("Flutes must be horizontal at this size.", , Title:="Error!")
                        Exit Sub
                    End If
                End If

            Case 1 ' aluminum
                If Val(txtHeight.Text) > 48 And Val(txtWidth.Text) > 48 Then
                    MsgBox("Invalid page size. Aluminum sheet size is 48''x96''.", , Title:="Error!")
                    Exit Sub
                End If
                If Val(txtHeight.Text) > 96 Or Val(txtWidth.Text) > 96 Then
                    MsgBox("Invalid page size. Aluminum sheet size is 48''x96''.", , Title:="Error!")
                    Exit Sub
                End If
                If Val(txtHeight.Text) > 48 Then
                    If Val(txtWidth.Text) <> 48 Then
                        MsgBox("Cannot shear this size on aluminum. Shear can only cut up to 48''.")
                        Exit Sub
                    End If
                End If
                If Val(txtWidth.Text) > 48 Then
                    If Val(txtHeight.Text) <> 48 Then
                        MsgBox("Cannot shear this size on aluminum. Shear can only cut up to 48''.")
                        Exit Sub
                    End If
                End If

            Case 2 ' acm
                If Val(txtHeight.Text) > 48 And Val(txtWidth.Text) > 48 Then
                    MsgBox("Invalid page size. ACM sheet size is 48''x96''.", , Title:="Error!")
                    Exit Sub
                End If
                If Val(txtHeight.Text) > 96 Or Val(txtWidth.Text) > 96 Then
                    MsgBox("Invalid page size. ACM sheet size is 48''x96''.", , Title:="Error!")
                    Exit Sub
                End If
            Case 3 ' plastic
                If Val(txtHeight.Text) > 48 And Val(txtWidth.Text) > 48 Then
                    MsgBox("Invalid page size. Most plastic has a sheet size of 48''x96''.", , Title:="Error!")
                    Exit Sub
                End If
                If Val(txtHeight.Text) > 96 Or Val(txtWidth.Text) > 96 Then
                    MsgBox("Invalid page size. Most plastic has a sheet size of 48''x96''.", , Title:="Error!")
                    Exit Sub
                End If
                If Val(txtHeight.Text) >= 48 Or Val(txtWidth.Text) >= 48 Then
                    If Val(txtWidth.Text) > 36 Or Val(txtHeight.Text) > 36 Then
                        MsgBox("Careful! If this is polyethylene, the largest we can do is 36''x48''.", , Title:="Warning!")
                    End If
                End If

            Case 4 ' vinyl
                If Val(txtHeight.Text) > 50 And Val(txtWidth.Text) > 50 Then
                    MsgBox("Invalid page size. One side must be no more than 50''.", , Title:="Error!")
                    Exit Sub
                End If

            Case 5 ' banner
                If Val(txtHeight.Text) > 120 And Val(txtWidth.Text) > 120 Then
                    MsgBox("Invalid page size. One side must be no more than 120''.", , Title:="Error!")
                    Exit Sub
                End If

            Case 6 ' poster
            Case 7 ' tags

        End Select

        'end of error messages

        Dim appDirectory As String = AppDomain.CurrentDomain.BaseDirectory
        Dim templateFilePath As String = appDirectory & "TestTemplate.cdt"
        corelDoc = corelApp.CreateDocumentFromTemplate(templateFilePath)
        'corelDoc.Activate()
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

        'Sets flutes & position
        vertFlutes = corelDoc.ActivePage.Shapes("vertFlutes")

        horzFlutes = corelDoc.ActivePage.Shapes("horzFlutes")

        If lstFlutes.SelectedIndex = 0 Then 'Vertical flutes
            horzFlutes.Delete()
            vertFlutes.SetSize(, pgHeight / 5)
            vertFlutes.SetPosition(pgWidth + 2, pgHeight / 2 + vertFlutes.SizeHeight / 2)
        ElseIf lstFlutes.SelectedIndex = 1 Then 'Horizontal flutes
            vertFlutes.Delete()
            horzFlutes.SetSize(, pgHeight / 5)
            horzFlutes.SetPosition(pgWidth + 2, pgHeight / 2 + horzFlutes.SizeHeight / 2)
        ElseIf lstFlutes.SelectedIndex = 2 Then 'Both flutes
            vertFlutes.SetSize(, pgHeight / 5)
            horzFlutes.SetSize(, pgHeight / 5)
            vertFlutes.SetPosition(pgWidth + 2, pgHeight / 2 + vertFlutes.SizeHeight + 1)
            horzFlutes.SetPosition(pgWidth + 2, vertFlutes.PositionY - horzFlutes.SizeHeight - 2)
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

        Dim crnrSz As Double

        Select Case lstCorners.SelectedIndex
            Case 1
                crnrSz = 0.25
            Case 2
                crnrSz = 0.5
            Case 3
                crnrSz = 0.75
            Case 4
                crnrSz = 1
            Case 5
                crnrSz = 1.5

        End Select

        If lstCorners.SelectedIndex > 0 Then
            cornerRect = corelDoc.ActiveLayer.CreateRectangle(0, 0, pgWidth, pgHeight)
            ctrlRectangle.Outline.Width = 0
            cornerRect.Rectangle.RadiusUpperLeft = crnrSz
            cornerRect.Rectangle.RadiusLowerLeft = crnrSz
            cornerRect.Rectangle.RadiusLowerRight = crnrSz
            cornerRect.Rectangle.RadiusUpperRight = crnrSz
        End If

        Dim holeSz As Double = 0

        Select Case lstHoleSz.SelectedIndex
            Case 1
                holeSz = 0.1875
            Case 2
                holeSz = 0.25
            Case 3
                holeSz = 0.3125
            Case 4
                holeSz = 0.375
        End Select

        Dim radTxt, holeSzTxt, holeLDistTxt, holeRDistTxt, holeLRDistTxt, holeTDistTxt, holeBDistTxt, holeTBDistTxt, holePlcTxt,
        holeCTBTxt, holeCTTxt, holeCLRTxt, holeTCText, holeCornerText, holeEQDist As String

        If ckbxUL.IsChecked Then
            corelApp.ActiveLayer.CreateEllipse2(lDist, pgHeight - tDist, holeSz / 2)
        End If
        If ckbxUC.IsChecked Then
            corelApp.ActiveLayer.CreateEllipse2(pgWidth / 2, pgHeight - tDist, holeSz / 2)
        End If
        If ckbxUR.IsChecked Then
            corelApp.ActiveLayer.CreateEllipse2(pgWidth - rDist, pgHeight - tDist, holeSz / 2)
        End If
        If ckbxCL.IsChecked Then
            corelApp.ActiveLayer.CreateEllipse2(lDist, pgHeight / 2, holeSz / 2)
        End If
        If ckbxCR.IsChecked Then
            corelApp.ActiveLayer.CreateEllipse2(pgWidth - rDist, pgHeight / 2, holeSz / 2)
        End If
        If ckbxLL.IsChecked Then
            corelApp.ActiveLayer.CreateEllipse2(lDist, bDist, holeSz / 2)
        End If
        If ckbxLC.IsChecked Then
            corelApp.ActiveLayer.CreateEllipse2(pgWidth / 2, bDist, holeSz / 2)
        End If
        If ckbxLR.IsChecked Then
            corelApp.ActiveLayer.CreateEllipse2(pgWidth - rDist, bDist, holeSz / 2)
        End If

        'Writing corner and hole description
        If lstCorners.SelectedIndex <> -1 Or lstHoleSz.SelectedIndex <> -1 Then


            If lstCorners.SelectedIndex <> -1 And lstCorners.SelectedIndex <> 0 Then

                radTxt = lstCorners.SelectionBoxItem.ToString + " radius corners"

            Else
                radTxt = ""

            End If
            If lstHoleSz.SelectedIndex > 0 Then

                holeSzTxt = lstHoleSz.SelectionBoxItem.ToString + " holes"
                If tDist = lDist And bDist = rDist And lDist = bDist Then
                    If tDist = 0.25 Then
                        holeEQDist = "1/4"
                    ElseIf tDist = 0.375 Then
                        holeEQDist = "3/8"
                    ElseIf tDist = 0.5 Then
                        holeEQDist = "1/2"
                    ElseIf tDist = 0.625 Then
                        holeEQDist = "5/8"
                    ElseIf tDist = 0.75 Then
                        holeEQDist = "3/4"
                    Else
                        holeEQDist = tDist.ToString
                    End If
                    holeSzTxt = holeSzTxt + ", " + holeEQDist + "'' from edge"
                Else
                    If tDist = bDist And tDist <> lDist Then
                        If tDist = 0.25 Then
                            holeTBDistTxt = "1/4'' from top & bottom edge" + vbCrLf
                        ElseIf tDist = 0.375 Then
                            holeTBDistTxt = "3/8'' from top & bottom edge" + vbCrLf
                        ElseIf tDist = 0.5 Then
                            holeTBDistTxt = "1/2'' from top & bottom edge" + vbCrLf
                        ElseIf tDist = 0.625 Then
                            holeTBDistTxt = "5/8'' from top & bottom edge" + vbCrLf
                        ElseIf tDist = 0.75 Then
                            holeTBDistTxt = "3/4'' from top & bottom edge" + vbCrLf
                        ElseIf txtTDist Is Nothing Or tDist = 0 Then
                            holeTBDistTxt = ""
                        Else
                            holeTBDistTxt = txtTDist.Text + "'' from top & bottom edge" + vbCrLf
                        End If
                    End If
                    If lDist = rDist And lDist <> tDist Then
                        If lDist = 0.25 Then
                            holeLRDistTxt = "1/4'' from left & right edge" + vbCrLf
                        ElseIf lDist = 0.375 Then
                            holeLRDistTxt = "3/8'' from left & right edge" + vbCrLf
                        ElseIf lDist = 0.5 Then
                            holeLRDistTxt = "1/2'' from left & right edge" + vbCrLf
                        ElseIf lDist = 0.625 Then
                            holeLRDistTxt = "5/8'' from left & right edge" + vbCrLf
                        ElseIf lDist = 0.75 Then
                            holeLRDistTxt = "3/4'' from left & right edge" + vbCrLf
                        ElseIf txtLDist Is Nothing Or lDist = 0 Then
                            holeLRDistTxt = ""
                        Else
                            holeLRDistTxt = txtLDist.Text + "'' from left & right edge" + vbCrLf
                        End If
                    End If
                    If tDist <> bDist Then
                        If tDist = 0.25 Then
                            holeTDistTxt = "1/4'' from top edge"
                        ElseIf tDist = 0.375 Then
                            holeTDistTxt = "3/8'' from top edge"
                        ElseIf tDist = 0.5 Then
                            holeTDistTxt = "1/2'' from top edge"
                        ElseIf tDist = 0.625 Then
                            holeTDistTxt = "5/8'' from top edge"
                        ElseIf tDist = 0.75 Then
                            holeTDistTxt = "3/4'' from top edge"
                        Else
                            holeTDistTxt = txtTDist.Text + "'' from top edge"
                        End If
                        If bDist = 0.25 Then
                            holeBDistTxt = "1/4'' from bottom edge"
                        ElseIf bDist = 0.375 Then
                            holeBDistTxt = "3/8'' from bottom edge"
                        ElseIf bDist = 0.5 Then
                            holeBDistTxt = "1/2'' from bottom edge"
                        ElseIf bDist = 0.625 Then
                            holeBDistTxt = "5/8'' from bottom edge"
                        ElseIf bDist = 0.75 Then
                            holeBDistTxt = "3/4'' from bottom edge"
                        Else
                            holeBDistTxt = txtBDist.Text + "'' from bottom edge"
                        End If
                        holeTBDistTxt = holeTDistTxt + ", " + holeBDistTxt + vbCrLf
                    End If
                    If lDist <> rDist Then
                        If lDist = 0.25 Then
                            holeLDistTxt = "1/4'' from left edge"
                        ElseIf lDist = 0.375 Then
                            holeLDistTxt = "3/8'' from left edge"
                        ElseIf lDist = 0.5 Then
                            holeLDistTxt = "1/2'' from left edge"
                        ElseIf lDist = 0.625 Then
                            holeLDistTxt = "5/8'' from left edge"
                        ElseIf lDist = 0.75 Then
                            holeLDistTxt = "3/4'' from left edge"
                        Else
                            holeLDistTxt = txtLDist.Text + "'' from left edge"
                        End If
                        If rDist = 0.25 Then
                            holeRDistTxt = "1/4'' from right edge"
                        ElseIf rDist = 0.375 Then
                            holeRDistTxt = "3/8'' from right edge"
                        ElseIf rDist = 0.5 Then
                            holeRDistTxt = "1/2'' from right edge"
                        ElseIf rDist = 0.625 Then
                            holeRDistTxt = "5/8'' from right edge"
                        ElseIf rDist = 0.75 Then
                            holeRDistTxt = "3/4'' from right edge"
                        Else
                            holeRDistTxt = txtRDist.Text + "'' from right edge"
                        End If
                        holeLRDistTxt = holeLDistTxt + ", " + holeRDistTxt + vbCrLf

                    End If

                End If

                If ckbxUC.IsChecked = True And ckbxLC.IsChecked = True Then
                    holeCTBTxt = "in center at top & bottom" + vbCrLf
                    If ckbxCL.IsChecked = False And ckbxCR.IsChecked = False Then
                        holeLRDistTxt = ""
                    End If
                ElseIf ckbxUC.IsChecked = True And ckbxLC.IsChecked = False Then
                    holeCTTxt = "in center at top" + vbCrLf
                    If ckbxCL.IsChecked = False And ckbxCR.IsChecked = False Then
                        holeLRDistTxt = ""
                    End If
                    If tDist = 0.25 Then
                        holeTDistTxt = "1/4'' from top edge"
                    ElseIf tDist = 0.375 Then
                        holeTDistTxt = "3/8'' from top edge"
                    ElseIf tDist = 0.5 Then
                        holeTDistTxt = "1/2'' from top edge"
                    ElseIf tDist = 0.625 Then
                        holeTDistTxt = "5/8'' from top edge"
                    ElseIf tDist = 0.75 Then
                        holeTDistTxt = "3/4'' from top edge"
                    Else
                        holeTDistTxt = txtTDist.Text + "'' from top edge"
                    End If
                    holeTBDistTxt = holeTDistTxt + vbCrLf
                    holeSzTxt = lstHoleSz.SelectionBoxItem.ToString + " holes"
                    If tDist = lDist And lDist = rDist Then
                        If tDist = 0.25 Then
                            holeEQDist = "1/4"
                        ElseIf tDist = 0.375 Then
                            holeEQDist = "3/8"
                        ElseIf tDist = 0.5 Then
                            holeEQDist = "1/2"
                        ElseIf tDist = 0.625 Then
                            holeEQDist = "5/8"
                        ElseIf tDist = 0.75 Then
                            holeEQDist = "3/4"
                        End If
                        holeSzTxt = holeSzTxt + ", " + holeEQDist + "'' from edge"
                        holeTBDistTxt = ""
                        holeLRDistTxt = ""
                    End If
                End If
                If ckbxCL.IsChecked = True And ckbxCR.IsChecked = True Then
                    holeCLRTxt = "in center at left & right" + vbCrLf
                    If ckbxUC.IsChecked = False And ckbxLC.IsChecked = False Then
                        holeTBDistTxt = ""
                    End If
                End If
                If ckbxUL.IsChecked = True And ckbxUR.IsChecked = True And ckbxLL.IsChecked = False And ckbxLR.IsChecked = False Then
                    holeTCText = "in top two corners" + vbCrLf
                End If
                If ckbxUL.IsChecked = True And ckbxUR.IsChecked = True And ckbxLL.IsChecked = True And ckbxLR.IsChecked = True Then
                    holeCornerText = "one in each corner" + vbCrLf
                End If
                holePlcTxt = holePlcTxt + holeCTBTxt + holeCTTxt + holeCLRTxt + holeTCText + holeCornerText
                ' MsgBox(holePlcTxt.ToString)
                holeSzTxt = holeSzTxt + vbCrLf + holeTBDistTxt + holeLRDistTxt + holePlcTxt

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
        Dim grommetUL, grommetUC, grommetTBSpacing, grommetCL, grommetLRSpacing, grommetTBQty, grommetLRQty As Corel.Interop.VGCore.Shape
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

        If ckbxTBQty.IsChecked Then
            Dim tbQtyDist As Single = ((pgWidth - 2)) / (tbQty - 1)
            Dim i As Integer
            grommetTBQty = grommet.Duplicate()
            grommetTBQty.SetPosition(tbQtyDist + 0.625, pgHeight - 0.625)
            grommetTBQty.Duplicate(, 0 - pgHeight + 2)

            For i = 1 To tbQty - 3
                grommetTBQty.Duplicate((tbQtyDist) * i)
                grommetTBQty.Duplicate((tbQtyDist) * i, 0 - pgHeight + 2)
            Next i
        End If

        If ckbxLRQty.IsChecked Then
            Dim LRQtyDist As Single = ((pgHeight - 2)) / (lrQty - 1)
            Dim i As Integer
            grommetLRQty = grommet.Duplicate()
            grommetLRQty.SetPosition(0.625, LRQtyDist + 0.625)
            grommetLRQty.Duplicate(pgWidth - 2)

            For i = 1 To lrQty - 3
                grommetLRQty.Duplicate(0, (LRQtyDist) * i)
                grommetLRQty.Duplicate(pgWidth - 2, LRQtyDist * i)
            Next i
        End If

        grommet.Delete()

        'grommet & banner text
        Dim bannerSizeTxt, grommetTxt As String
        Dim bannerArtTxt, grommetArtTxt As Corel.Interop.VGCore.Shape

        If ckbxCLRgroms.IsChecked = True Or ckbxCornerGroms.IsChecked = True Or ckbxCTBGroms.IsChecked = True Or ckbxLRspacing.IsChecked = True Or ckbxTBspacing.IsChecked = True Then
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
                If ckbxTBQty.IsChecked Then
                    grommetTxt += ", " + txtTBQty.Text + " grommets along top & bottom"
                End If
                If ckbxLRQty.IsChecked Then
                    grommetTxt += ", " + txtLRQty.Text + " grommets along left & right"
                End If
                grommetArtTxt = corelDoc.ActiveLayer.CreateArtisticText(0, 0, grommetTxt, , ,
                                                    "Arial", 1.5 * ((pgWidth + pgHeight) / 2) + 34, , , , Corel.Interop.VGCore.cdrAlignment.cdrLeftAlignment)
                grommetArtTxt.SetPosition(0, -3 * bannerArtTxt.SizeHeight)
            Else
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
                If ckbxTBQty.IsChecked Then
                    grommetTxt += ", " + txtTBQty.Text + " grommets along top & bottom"
                End If
                If ckbxLRQty.IsChecked Then
                    grommetTxt += ", " + txtLRQty.Text + " grommets along left & right"
                End If
                grommetArtTxt = corelDoc.ActiveLayer.CreateArtisticText(0, 0, grommetTxt, , ,
                                                    "Arial", 1.5 * ((pgWidth + pgHeight) / 2) + 34, , , , Corel.Interop.VGCore.cdrAlignment.cdrLeftAlignment)
                grommetArtTxt.SetPosition(0, -grommetArtTxt.SizeHeight - 1)
            End If
        End If

        'Tags
        tagBorder = corelDoc.ActiveLayer.Shapes.FindShape("tagBorder")
        tagHoles = corelDoc.ActiveLayer.Shapes.FindShape("tagHoles")

        If lstMaterial.SelectedIndex = 7 Then
            tagBorder.SetPosition(0, 6)
            tagHoles.SetPosition(1.85, 5.5)
            ctrlRectangle.Outline.Width = 0
            radTxt = "1/2'' radius corners"
            Dim tagText As Corel.Interop.VGCore.Shape = corelDoc.ActiveLayer.CreateArtisticText(0, 0, radTxt, , ,
                                                    "Arial", , , , , Corel.Interop.VGCore.cdrAlignment.cdrCenterAlignment)

            tagText.SetSize(pgWidth)
            tagText.SetPosition(0, 0 - (pgHeight * 0.125))
            tagText.AlignToShape(Corel.Interop.VGCore.cdrAlignType.cdrAlignHCenter, ctrlRectangle)
        Else
            tagBorder.Delete()
            tagHoles.Delete()
        End If

        'Change view to fit everything on Layer 1
        corelApp.ActiveWindow.ActiveView.ToFitPage()
        corelDoc.ClearSelection()

    End Sub
End Class
