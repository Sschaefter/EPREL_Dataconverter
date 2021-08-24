Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Globalization


Imports <xmlns:ns3="http://eprel.ener.ec.europa.eu/services/productModelService/modelRegistrationService/v2">
Imports <xmlns:ns2="http://eprel.ener.ec.europa.eu/productModel/productCore/v2">
Imports <xmlns:ns5="http://eprel.ener.ec.europa.eu/commonTypes/EnergyLabelTypes/v2">

Public Class Form1

    Public items() As String
    Public _EPREL_MODEL_REGISTRATION_NUMBER() As String
    Public _MODEL_IDENTIFIER() As String
    Public _CONSIDER_GENERATED_LABEL_AS_PROVIDED() As String
    Public _ON_MARKET_START_DATE() As String
    Public _ON_MARKET_END_DATE() As String
    Public _VISIBLE_TO_UK_MSA() As String
    Public _LIGHTING_TECHNOLOGY() As String
    Public _DIRECTIONAL() As String
    Public _CAP_TYPE() As String
    Public _MAINS() As String
    Public _CONNECTED_LIGHT_SOURCE() As String
    Public _COLOUR_TUNEABLE_LIGHT_SOURCE() As String
    Public _ENVELOPE() As String
    Public _HIGH_LUMINANCE_LIGHT_SOURCE() As String
    Public _ANTI_GLARE_SHIELD() As String
    Public _DIMMABLE() As String
    Public _ENERGY_CONS_ON_MODE() As String
    Public _ENERGY_CLASS() As String
    Public _LUMINOUS_FLUX() As String
    Public _BEAM_ANGLE_CORRESPONDENCE() As String
    Public _CORRELATED_COLOUR_TEMP_TYPE() As String
    Public _CORRELATED_COLOUR_TEMP_SINGLE() As String
    Public _CORRELATED_COLOUR_TEMP_MIN() As String
    Public _CORRELATED_COLOUR_TEMP_MAX() As String
    Public _CORRELATED_COLOUR_TEMP_1() As String
    Public _CORRELATED_COLOUR_TEMP_2() As String
    Public _CORRELATED_COLOUR_TEMP_3() As String
    Public _CORRELATED_COLOUR_TEMP_4() As String
    Public _POWER_ON_MODE() As String
    Public _POWER_STANDBY() As String
    Public _POWER_STANDBY_NETWORKED() As String
    Public _COLOUR_RENDERING_INDEX() As String
    Public _MIN_COLOUR_RENDERING_INDEX() As String
    Public _MAX_COLOUR_RENDERING_INDEX() As String
    Public _DIMENSION_HEIGHT() As String
    Public _DIMENSION_WIDTH() As String
    Public _DIMENSION_DEPTH() As String
    Public _SPECTRAL_POWER_DISTRIBUTION_IMAGE() As String
    Public _CLAIM_EQUIVALENT_POWER() As String
    Public _EQUIVALENT_POWER() As String
    Public _CHROMATICITY_COORD_X() As String
    Public _CHROMATICITY_COORD_Y() As String
    Public _DLS_PEAK_LUMINOUS_INTENSITY() As String
    Public _DLS_BEAM_ANGLE() As String
    Public _DLS_MIN_BEAM_ANGLE() As String
    Public _DLS_MAX_BEAM_ANGLE() As String
    Public _LED_R9_COLOUR_RENDERING_INDEX() As String
    Public _LED_SURVIVAL_FACTOR() As String
    Public _LED_LUMEN_MAINTENANCE_FACTOR() As String
    Public _LED_MLS_DISPLACEMENT_FACTOR() As String
    Public _LED_MLS_COLOUR_CONSISTENCY() As String
    Public _LED_MLS_CLAIM_LED_REPLACE_FLUORESCENT() As String
    Public _LED_MLS_FL_REPLACEMENT_CLAIM() As String
    Public _LED_MLS_FLICKER_METRIC() As String
    Public _LED_MLS_STROBOSCOPIC_EFFECT_METRIC() As String
    Public row As Integer
    Public col As String
    Public sheet As String

    Structure _TECHNICAL_DOCUMENTATION
        Public _TD_MODEL_IDENTIFIER As String
        Public _TD_DESCRIPTION As String
        Public _TD_LANGUAGE As String
        Public _TD_ADDITIONAL_PART As Boolean
        Public _TD_CALCULATIONS As Boolean
        Public _TD_GENERAL_DESCRIPTION As Boolean
        Public _TD_MESURED_TECHNICAL_PARAMETERS As Boolean
        Public _TD_REFERENCES_TO_HARMONIZED_STANDARDS As Boolean
        Public _TD_SPECIFIC_PRECAUTIONS As Boolean
        Public _TD_TESTING_CONDITIONS As Boolean
        Public _TD_FILE_NAME As String
    End Structure
    Public _TD() As _TECHNICAL_DOCUMENTATION
    Public dummy, dummy2 As Integer
    'Public doc As XmlDocument = New XmlDocument()
    Public doc As XDocument = New XDocument()
    Public state As Boolean = False
    Public errorstate As Boolean = False
    '---Form & GUI
    Private Sub CB_OperationType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_OperationType.SelectedIndexChanged
        If CB_OperationType.SelectedItem = "UPDATE_PRODUCT_MODEL" Then
            CB_ReasonChange.Enabled = True
            CB_ReasonChange.SelectedIndex = 0
        Else
            CB_ReasonChange.Enabled = False
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '---steback of states
        errorstate = False
        state = False
        'If CheckB_Log.Checked = True Then
        '    Form2.Visible = True
        'End If


        If CB_OperationType.SelectedItem = "REGISTER_PRODUCT_MODEL" Then
            Select Case MsgBox("Please make shure, that all attachments are named liked in the source table and are located in one folder!", vbOKCancel)
                Case MsgBoxResult.Cancel
                    Exit Sub
                Case MsgBoxResult.Ok
                    Exit Select
            End Select
            Form2.LB_Log.Items.Add("REGISTER_PRODUCT_MODEL")
            SELECT_INPUT()
            If state = True Then
                Exit Sub
            End If
            REGISTRATION()
            If state = True Then
                Exit Sub
            End If
            OUTPUT()
        ElseIf CB_OperationType.SelectedItem = "PREREGISTER_PRODUCT_MODEL" Then
            Form2.LB_Log.Items.Add("PREREGISTER_PRODUCT_MODEL")
            SELECT_INPUT()
            If state = True Then
                Exit Sub
            End If
            PREREGISTRATION()
            If state = True Then
                Exit Sub
            End If
            OUTPUT()
        ElseIf CB_OperationType.SelectedItem = "UPDATE_PRODUCT_MODEL" Then
            Select Case MsgBox("Please make shure, that all attachments are named liked in the source table and are located in one folder!", vbOKCancel)
                Case MsgBoxResult.Cancel
                    Exit Sub
                Case MsgBoxResult.Ok
                    Exit Select
            End Select
            Form2.LB_Log.Items.Add("UPDATE_PRODUCT_MODEL")
            SELECT_INPUT()
            If state = True Then
                Exit Sub
            End If
            UPDATE_PRODUCT()
            If state = True Then
                Exit Sub
            End If
            OUTPUT()
        ElseIf CB_OperationType.SelectedItem = "DECLARE_END_DATE_OF_PLACEMENT_ON_MARKET" Then
            Form2.LB_Log.Items.Add("DECLARE_END_DATE_OF_PLACEMENT_ON_MARKET")
            SELECT_INPUT()
            If state = True Then
                Exit Sub
            End If
            DECLARE_END_DATE_OF_PLACEMENT_ON_MARKET()
            If state = True Then
                Exit Sub
            End If
            OUTPUT()
        End If



            Select Case MsgBox("Validate Zip File?", vbYesNo)
            Case MsgBoxResult.Yes
                Validate_ZIP()
            Case MsgBoxResult.No
                Exit Select
        End Select

        If CheckB_Log.Checked = True Then
            Save_Log_XML()
        End If

        'Close()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Validate_ZIP()
    End Sub
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("mailto:m.planeck@nimbus-group.com")
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form_Contact.Show()
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "EPREL Dataconverter " & My.Application.Info.Version.ToString
        Form2.Show()
        Form2.Hide()
        CB_RegistrantNature.SelectedIndex = 0

    End Sub
    Private Sub CheckB_Log_CheckedChanged(sender As Object, e As EventArgs) Handles CheckB_Log.CheckedChanged
        Select Case CheckB_Log.Checked
            Case True
                Form2.Show()
            Case False
                Form2.Hide()
        End Select
    End Sub

    '---Input
    Public Sub SELECT_INPUT()
        '----------------------------Validierung, ob Felder befüllt
        If Txt_Request.TextLength = 0 Or Txt_TrademarkRef.TextLength = 0 Then
            MsgBox("Please fill Values!")
            state = True
            Exit Sub
        End If

        If Txt_ContactRef.Text = "" And CB_OperationType.SelectedItem = "UPDATE_PRODUCT_MODEL" Then
            MsgBox("Please fill Values!")
            state = True
            Exit Sub
        End If
        '----------------------------Datei Auswählen und öffnen----------------------

        'xlApp.Visible = False
        Dim quelle As New OpenFileDialog
        quelle.Title = "Please select the source file!"
        quelle.Filter = "Excel files (*.xlsx)|*.xlsx"
        quelle.ShowDialog()
        If quelle.FileName = "" Then
            MsgBox("Error!")
            state = True
            Exit Sub
        End If

        Dim flname As String = quelle.FileName

        Select Case CB_OperationType.SelectedItem
            Case "UPDATE_PRODUCT_MODEL"
                PARSE_UPDATE(flname)
            Case "PREREGISTER_PRODUCT_MODEL"
                PARSE_PREREGISTER(flname)
            Case "REGISTER_PRODUCT_MODEL"
                PARSE_REGISTER(flname)
            Case "DECLARE_END_DATE_OF_PLACEMENT_ON_MARKET"
                PARSE_DEOP(flname)
                Exit Sub
        End Select



        'PARSE -DATA

        'Dim book = xlApp.Workbooks.Open(quelle.FileName)

        ''Dim book = xlApp.Workbooks.Open("C:\Users\User79\Desktop\EPREL_Datenkonvertierung_Python_20210201\quelle.xlsx")



        'Dim xltab1 = book.Worksheets("Tabelle1")
        ''Dim items() As String
        'Dim xlUP As Object = Excel.XlDirection.xlUp
        'Dim lastentry As Object
        ''Dim dummy As Integer

        'dummy = book.Sheets(1).Range("A" & xltab1.Rows.Count).End(xlUP).Row
        'lastentry = xltab1.Range("A1:A" & dummy).Value
        'ReDim items(dummy - 1)
        'For i = 1 To dummy - 1
        '    items(i - 1) = xltab1.Range("A" & i + 1).Value
        'Next
        'xlApp.Workbooks.Close()
        'xlApp.Quit()
        'PARSE_END


    End Sub

    '---Parsing
    Sub PARSE_PREREGISTER(ByVal quelle As String)

        Dim xlApp As New Excel.Application
        Dim book = xlApp.Workbooks.Open(quelle)
        Dim xltab1 = book.Worksheets("PREREGISTRATION")
        Dim xlUP As Object = Excel.XlDirection.xlUp
        Dim lastentry As Object
        Try
            dummy = xltab1.Range("A" & xltab1.Rows.Count).End(xlUP).Row

            lastentry = xltab1.Range("A1:A" & dummy).Value
            ReDim items(dummy - 1)
            For i = 1 To dummy - 1
                items(i - 1) = xltab1.Range("A" & i + 1).Value
            Next
        Catch ex As Exception
            ErrorDlg("parse", ex)

        End Try

        xlApp.Workbooks.Close()
            xlApp.Quit()
        'PARSE_END

    End Sub
    Sub PARSE_REGISTER(ByVal quelle As String)
        Dim xlApp As New Excel.Application
        Dim book = xlApp.Workbooks.Open(quelle)
        Dim xltab1 = book.Worksheets("REGISTER_PRODUCT_MODEL")
        Dim xltab2 = book.Worksheets("attachments")
        Dim xlUP As Object = Excel.XlDirection.xlUp
        Dim lastentry As Object

        Try
            dummy = xltab1.Range("A" & xltab1.Rows.Count).End(xlUP).Row
            lastentry = xltab1.Range("A1:A" & dummy).Value
            ReDim _MODEL_IDENTIFIER(dummy - 1)
            ReDim _CONSIDER_GENERATED_LABEL_AS_PROVIDED(dummy - 1)
            ReDim _ON_MARKET_START_DATE(dummy - 1)
            ReDim _ON_MARKET_END_DATE(dummy - 1)
            ReDim _VISIBLE_TO_UK_MSA(dummy - 1)
            ReDim _LIGHTING_TECHNOLOGY(dummy - 1)
            ReDim _DIRECTIONAL(dummy - 1)
            ReDim _CAP_TYPE(dummy - 1)
            ReDim _MAINS(dummy - 1)
            ReDim _CONNECTED_LIGHT_SOURCE(dummy - 1)
            ReDim _COLOUR_TUNEABLE_LIGHT_SOURCE(dummy - 1)
            ReDim _ENVELOPE(dummy - 1)
            ReDim _HIGH_LUMINANCE_LIGHT_SOURCE(dummy - 1)
            ReDim _ANTI_GLARE_SHIELD(dummy - 1)
            ReDim _DIMMABLE(dummy - 1)
            ReDim _ENERGY_CONS_ON_MODE(dummy - 1)
            ReDim _ENERGY_CLASS(dummy - 1)
            ReDim _LUMINOUS_FLUX(dummy - 1)
            ReDim _BEAM_ANGLE_CORRESPONDENCE(dummy - 1)
            ReDim _CORRELATED_COLOUR_TEMP_TYPE(dummy - 1)
            ReDim _CORRELATED_COLOUR_TEMP_SINGLE(dummy - 1)
            ReDim _CORRELATED_COLOUR_TEMP_MIN(dummy - 1)
            ReDim _CORRELATED_COLOUR_TEMP_MAX(dummy - 1)
            ReDim _CORRELATED_COLOUR_TEMP_1(dummy - 1)
            ReDim _CORRELATED_COLOUR_TEMP_2(dummy - 1)
            ReDim _CORRELATED_COLOUR_TEMP_3(dummy - 1)
            ReDim _CORRELATED_COLOUR_TEMP_4(dummy - 1)
            ReDim _POWER_ON_MODE(dummy - 1)
            ReDim _POWER_STANDBY(dummy - 1)
            ReDim _POWER_STANDBY_NETWORKED(dummy - 1)
            ReDim _COLOUR_RENDERING_INDEX(dummy - 1)
            ReDim _MIN_COLOUR_RENDERING_INDEX(dummy - 1)
            ReDim _MAX_COLOUR_RENDERING_INDEX(dummy - 1)
            ReDim _DIMENSION_HEIGHT(dummy - 1)
            ReDim _DIMENSION_WIDTH(dummy - 1)
            ReDim _DIMENSION_DEPTH(dummy - 1)
            ReDim _SPECTRAL_POWER_DISTRIBUTION_IMAGE(dummy - 1)
            ReDim _CLAIM_EQUIVALENT_POWER(dummy - 1)
            ReDim _EQUIVALENT_POWER(dummy - 1)
            ReDim _CHROMATICITY_COORD_X(dummy - 1)
            ReDim _CHROMATICITY_COORD_Y(dummy - 1)
            ReDim _DLS_PEAK_LUMINOUS_INTENSITY(dummy - 1)
            ReDim _DLS_BEAM_ANGLE(dummy - 1)
            ReDim _DLS_MIN_BEAM_ANGLE(dummy - 1)
            ReDim _DLS_MAX_BEAM_ANGLE(dummy - 1)
            ReDim _LED_R9_COLOUR_RENDERING_INDEX(dummy - 1)
            ReDim _LED_SURVIVAL_FACTOR(dummy - 1)
            ReDim _LED_LUMEN_MAINTENANCE_FACTOR(dummy - 1)
            ReDim _LED_MLS_DISPLACEMENT_FACTOR(dummy - 1)
            ReDim _LED_MLS_COLOUR_CONSISTENCY(dummy - 1)
            ReDim _LED_MLS_CLAIM_LED_REPLACE_FLUORESCENT(dummy - 1)
            ReDim _LED_MLS_FL_REPLACEMENT_CLAIM(dummy - 1)
            ReDim _LED_MLS_FLICKER_METRIC(dummy - 1)
            ReDim _LED_MLS_STROBOSCOPIC_EFFECT_METRIC(dummy - 1)
            Dim provider As CultureInfo = New CultureInfo("en-EN")
            Dim dmy1 As Date
            Dim dmy2 As Double

            For i = 1 To dummy - 1
                sheet = "REGISTER_PRODUCT_MODEL"
                row = i + 1
                col = "A"
                _MODEL_IDENTIFIER(i - 1) = xltab1.Range(col & i + 1).Value
                col = "B"
                _CONSIDER_GENERATED_LABEL_AS_PROVIDED(i - 1) = xltab1.Range(col & i + 1).Value
                col = "C"
                dmy1 = xltab1.Range(col & i + 1).Value
                '-Format date to yyyy-mm-dd+hh:mm
                _ON_MARKET_START_DATE(i - 1) = dmy1.ToString("yyyy") & "-" & dmy1.ToString("MM") & "-" & dmy1.ToString("dd") & dmy1.ToString("zzz")
                col = "D"
                _VISIBLE_TO_UK_MSA(i - 1) = xltab1.Range("D" & i + 1).Value
                col = "E"
                _LIGHTING_TECHNOLOGY(i - 1) = xltab1.Range("E" & i + 1).Value
                col = "F"
                _DIRECTIONAL(i - 1) = xltab1.Range("F" & i + 1).Value
                col = "G"
                _CAP_TYPE(i - 1) = xltab1.Range("G" & i + 1).Value
                col = "H"
                _MAINS(i - 1) = xltab1.Range("H" & i + 1).Value
                col = "I"
                _CONNECTED_LIGHT_SOURCE(i - 1) = xltab1.Range("I" & i + 1).Value
                col = "J"
                _COLOUR_TUNEABLE_LIGHT_SOURCE(i - 1) = xltab1.Range("J" & i + 1).Value
                col = "K"
                _ENVELOPE(i - 1) = xltab1.Range("K" & i + 1).Value
                col = "L"
                _HIGH_LUMINANCE_LIGHT_SOURCE(i - 1) = xltab1.Range("L" & i + 1).Value
                col = "M"
                _ANTI_GLARE_SHIELD(i - 1) = xltab1.Range("M" & i + 1).Value
                col = "N"
                _DIMMABLE(i - 1) = xltab1.Range("N" & i + 1).Value
                col = "O"
                _ENERGY_CONS_ON_MODE(i - 1) = Math.Round(Convert.ToDecimal(xltab1.Range("O" & i + 1).Value))
                col = "P"
                '_ENERGY_CONS_ON_MODE(i - 1) = String.Format("{0000}", xltab1.Range("P" & i + 1).Value)
                _ENERGY_CLASS(i - 1) = xltab1.Range("P" & i + 1).Value
                col = "R"
                _LUMINOUS_FLUX(i - 1) = Math.Round(Convert.ToDecimal(xltab1.Range("R" & i + 1).Value))
                col = "S"
                '_LUMINOUS_FLUX(i - 1) = String.Format("{00000}", xltab1.Range("S" & i + 1).Value)
                _BEAM_ANGLE_CORRESPONDENCE(i - 1) = xltab1.Range("S" & i + 1).Value
                col = "T"
                _CORRELATED_COLOUR_TEMP_TYPE(i - 1) = xltab1.Range("T" & i + 1).Value

                'CCT - Single
                col = "U"
                dmy2 = Math.Ceiling(xltab1.Range("U" & i + 1).Value / 100)
                _CORRELATED_COLOUR_TEMP_SINGLE(i - 1) = String.Format("{0000}", dmy2 * 100)
                'CCT - Range
                'MIN
                col = "V"
                dmy2 = Math.Ceiling(xltab1.Range("V" & i + 1).Value / 100)
                _CORRELATED_COLOUR_TEMP_MIN(i - 1) = String.Format("{00000}", dmy2 * 100)
                'MAX
                col = "W"
                dmy2 = Math.Ceiling(xltab1.Range("W" & i + 1).Value / 100)
                _CORRELATED_COLOUR_TEMP_MAX(i - 1) = String.Format("{00000}", dmy2 * 100)
                'CCT - Steps
                col = "X"
                dmy2 = Math.Ceiling(xltab1.Range("X" & i + 1).Value / 100)
                _CORRELATED_COLOUR_TEMP_1(i - 1) = String.Format("{00000}", dmy2 * 100)
                col = "Y"
                dmy2 = Math.Ceiling(xltab1.Range("Y" & i + 1).Value / 100)
                _CORRELATED_COLOUR_TEMP_2(i - 1) = String.Format("{00000}", dmy2 * 100)
                col = "Z"
                dmy2 = Math.Ceiling(xltab1.Range("Z" & i + 1).Value / 100)
                _CORRELATED_COLOUR_TEMP_3(i - 1) = String.Format("{00000}", dmy2 * 100)
                col = "AA"
                dmy2 = Math.Ceiling(xltab1.Range("AA" & i + 1).Value / 100)
                _CORRELATED_COLOUR_TEMP_4(i - 1) = String.Format("{00000}", dmy2 * 100)
                'Power_ON_MODE
                col = "AB"
                dmy2 = xltab1.Range("AB" & i + 1).Value
                _POWER_ON_MODE(i - 1) = String.Format(provider, "{0:###0.0}", dmy2)
                col = "AC"
                dmy2 = xltab1.Range("AC" & i + 1).Value
                _POWER_STANDBY(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
                col = "AD"
                dmy2 = xltab1.Range("AD" & i + 1).Value
                _POWER_STANDBY_NETWORKED(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
                col = "AE"
                dmy2 = xltab1.Range("AE" & i + 1).Value
                _COLOUR_RENDERING_INDEX(i - 1) = String.Format(provider, "{0:###}", dmy2)
                col = "AF"
                dmy2 = xltab1.Range("AF" & i + 1).Value
                _MIN_COLOUR_RENDERING_INDEX(i - 1) = String.Format(provider, "{0:###}", dmy2)
                col = "AG"
                dmy2 = xltab1.Range("AG" & i + 1).Value
                _MAX_COLOUR_RENDERING_INDEX(i - 1) = String.Format(provider, "{0:###}", dmy2)
                col = "AH"
                dmy2 = xltab1.Range("AH" & i + 1).Value
                _DIMENSION_HEIGHT(i - 1) = String.Format(provider, "{0:#####}", dmy2)
                col = "AI"
                dmy2 = xltab1.Range("AI" & i + 1).Value
                _DIMENSION_WIDTH(i - 1) = String.Format(provider, "{0:#####}", dmy2)
                col = "AJ"
                dmy2 = xltab1.Range("AJ" & i + 1).Value
                _DIMENSION_DEPTH(i - 1) = String.Format(provider, "{0:#####}", dmy2)
                col = "AK"
                _SPECTRAL_POWER_DISTRIBUTION_IMAGE(i - 1) = xltab1.Range("AK" & i + 1).Value
                _CLAIM_EQUIVALENT_POWER(i - 1) = xltab1.Range("AL" & i + 1).Value
                col = "AM"
                dmy2 = xltab1.Range("AM" & i + 1).Value
                _EQUIVALENT_POWER(i - 1) = String.Format(provider, "{0:#####}", dmy2)
                col = "AN"
                dmy2 = xltab1.Range("AN" & i + 1).Value
                _CHROMATICITY_COORD_X(i - 1) = String.Format(provider, "{0:0.000}", dmy2)
                col = "AO"
                dmy2 = xltab1.Range("AO" & i + 1).Value
                _CHROMATICITY_COORD_Y(i - 1) = String.Format(provider, "{0:0.000}", dmy2)
                col = "AP"
                dmy2 = xltab1.Range("AP" & i + 1).Value
                _DLS_PEAK_LUMINOUS_INTENSITY(i - 1) = String.Format(provider, "{0:######}", dmy2)
                col = "AQ"
                dmy2 = xltab1.Range("AQ" & i + 1).Value
                _DLS_BEAM_ANGLE(i - 1) = String.Format(provider, "{0:###}", dmy2)
                col = "AR"
                dmy2 = xltab1.Range("AR" & i + 1).Value
                _DLS_MIN_BEAM_ANGLE(i - 1) = String.Format(provider, "{0:###}", dmy2)
                col = "AS"
                dmy2 = xltab1.Range("AS" & i + 1).Value
                _DLS_MAX_BEAM_ANGLE(i - 1) = String.Format(provider, "{0:###}", dmy2)
                ' R9
                col = "AT"
                _LED_R9_COLOUR_RENDERING_INDEX(i - 1) = Math.Round(Convert.ToDecimal(xltab1.Range("AT" & i + 1).Value))
                'dmy2 = xltab1.Range("AU" & i + 1).Value
                '_LED_R9_COLOUR_RENDERING_INDEX(i - 1) = String.Format(provider, "{0:###}", dmy2)
                col = "AU"
                dmy2 = xltab1.Range("AU" & i + 1).Value
                _LED_SURVIVAL_FACTOR(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
                col = "AV"
                dmy2 = xltab1.Range("AV" & i + 1).Value
                _LED_LUMEN_MAINTENANCE_FACTOR(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
                col = "AW"
                dmy2 = xltab1.Range("AW" & i + 1).Value
                _LED_MLS_DISPLACEMENT_FACTOR(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
                col = "AX"
                dmy2 = xltab1.Range("AX" & i + 1).Value
                _LED_MLS_COLOUR_CONSISTENCY(i - 1) = String.Format(provider, "{0:#}", dmy2)
                col = "AY"
                _LED_MLS_CLAIM_LED_REPLACE_FLUORESCENT(i - 1) = xltab1.Range("AY" & i + 1).Value
                col = "AZ"
                dmy2 = xltab1.Range("AZ" & i + 1).Value
                _LED_MLS_FL_REPLACEMENT_CLAIM(i - 1) = String.Format(provider, "{0:##}", dmy2)
                col = "BA"
                dmy2 = xltab1.Range("BA" & i + 1).Value
                _LED_MLS_FLICKER_METRIC(i - 1) = String.Format(provider, "{0:0.0}", dmy2)
                col = "BB"
                dmy2 = xltab1.Range("BB" & i + 1).Value
                _LED_MLS_STROBOSCOPIC_EFFECT_METRIC(i - 1) = String.Format(provider, "{0:0.0}", dmy2)

            Next

            dummy2 = book.Worksheets("attachments").Range("A" & xltab2.Rows.Count).End(xlUP).Row
            lastentry = xltab2.Range("A1:A" & dummy2).Value

            dummy2 = dummy2 - 2
            ReDim _TD(dummy2)
            For i = 0 To dummy2
                sheet = "attachments"
                row = i + 1
                col = "A"
                _TD(i)._TD_MODEL_IDENTIFIER = xltab2.Range("A" & i + 2).Value
                col = "B"
                _TD(i)._TD_DESCRIPTION = xltab2.Range("B" & i + 2).Value
                col = "C"
                _TD(i)._TD_LANGUAGE = xltab2.Range("C" & i + 2).Value
                col = "D"
                _TD(i)._TD_ADDITIONAL_PART = xltab2.Range("D" & i + 2).Value
                col = "E"
                _TD(i)._TD_CALCULATIONS = xltab2.Range("E" & i + 2).Value
                col = "F"
                _TD(i)._TD_GENERAL_DESCRIPTION = xltab2.Range("F" & i + 2).Value
                col = "G"
                _TD(i)._TD_MESURED_TECHNICAL_PARAMETERS = xltab2.Range("G" & i + 2).Value
                col = "H"
                _TD(i)._TD_REFERENCES_TO_HARMONIZED_STANDARDS = xltab2.Range("H" & i + 2).Value
                col = "I"
                _TD(i)._TD_TESTING_CONDITIONS = xltab2.Range("I" & i + 2).Value
                col = "J"
                _TD(i)._TD_SPECIFIC_PRECAUTIONS = xltab2.Range("J" & i + 2).Value
                col = "K"
                _TD(i)._TD_FILE_NAME = xltab2.Range("K" & i + 2).Value
            Next

        Catch ex As Exception
            ErrorDlg("parse", ex, row, col, sheet)
        End Try


        book.Close(False)
        xlApp.Quit()

        'System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
        'xlApp = Nothing
    End Sub
    Sub PARSE_UPDATE(ByVal quelle As String)
        Dim xlApp As New Excel.Application
        Dim book = xlApp.Workbooks.Open(quelle)
        Dim xltab1 = book.Worksheets("UPDATE_PRODUCT_MODEL")
        Dim xltab2 = book.Worksheets("attachments")
        Dim xlUP As Object = Excel.XlDirection.xlUp
        Dim lastentry As Object

        Try
            dummy = xltab1.Range("A" & xltab1.Rows.Count).End(xlUP).Row

            'dummy = book.Sheets(1).Range("A" & xltab1.Rows.Count).End(xlUP).Row
            lastentry = xltab1.Range("A1:A" & dummy).Value
            ReDim _EPREL_MODEL_REGISTRATION_NUMBER(dummy - 1)
            ReDim _MODEL_IDENTIFIER(dummy - 1)
            ReDim _CONSIDER_GENERATED_LABEL_AS_PROVIDED(dummy - 1)
            ReDim _ON_MARKET_START_DATE(dummy - 1)
            ReDim _ON_MARKET_END_DATE(dummy - 1)
            ReDim _VISIBLE_TO_UK_MSA(dummy - 1)
            ReDim _LIGHTING_TECHNOLOGY(dummy - 1)
            ReDim _DIRECTIONAL(dummy - 1)
            ReDim _CAP_TYPE(dummy - 1)
            ReDim _MAINS(dummy - 1)
            ReDim _CONNECTED_LIGHT_SOURCE(dummy - 1)
            ReDim _COLOUR_TUNEABLE_LIGHT_SOURCE(dummy - 1)
            ReDim _ENVELOPE(dummy - 1)
            ReDim _HIGH_LUMINANCE_LIGHT_SOURCE(dummy - 1)
            ReDim _ANTI_GLARE_SHIELD(dummy - 1)
            ReDim _DIMMABLE(dummy - 1)
            ReDim _ENERGY_CONS_ON_MODE(dummy - 1)
            ReDim _ENERGY_CLASS(dummy - 1)
            ReDim _LUMINOUS_FLUX(dummy - 1)
            ReDim _BEAM_ANGLE_CORRESPONDENCE(dummy - 1)
            ReDim _CORRELATED_COLOUR_TEMP_TYPE(dummy - 1)
            ReDim _CORRELATED_COLOUR_TEMP_SINGLE(dummy - 1)
            ReDim _CORRELATED_COLOUR_TEMP_MIN(dummy - 1)
            ReDim _CORRELATED_COLOUR_TEMP_MAX(dummy - 1)
            ReDim _CORRELATED_COLOUR_TEMP_1(dummy - 1)
            ReDim _CORRELATED_COLOUR_TEMP_2(dummy - 1)
            ReDim _CORRELATED_COLOUR_TEMP_3(dummy - 1)
            ReDim _CORRELATED_COLOUR_TEMP_4(dummy - 1)
            ReDim _POWER_ON_MODE(dummy - 1)
            ReDim _POWER_STANDBY(dummy - 1)
            ReDim _POWER_STANDBY_NETWORKED(dummy - 1)
            ReDim _COLOUR_RENDERING_INDEX(dummy - 1)
            ReDim _MIN_COLOUR_RENDERING_INDEX(dummy - 1)
            ReDim _MAX_COLOUR_RENDERING_INDEX(dummy - 1)
            ReDim _DIMENSION_HEIGHT(dummy - 1)
            ReDim _DIMENSION_WIDTH(dummy - 1)
            ReDim _DIMENSION_DEPTH(dummy - 1)
            ReDim _SPECTRAL_POWER_DISTRIBUTION_IMAGE(dummy - 1)
            ReDim _CLAIM_EQUIVALENT_POWER(dummy - 1)
            ReDim _EQUIVALENT_POWER(dummy - 1)
            ReDim _CHROMATICITY_COORD_X(dummy - 1)
            ReDim _CHROMATICITY_COORD_Y(dummy - 1)
            ReDim _DLS_PEAK_LUMINOUS_INTENSITY(dummy - 1)
            ReDim _DLS_BEAM_ANGLE(dummy - 1)
            ReDim _DLS_MIN_BEAM_ANGLE(dummy - 1)
            ReDim _DLS_MAX_BEAM_ANGLE(dummy - 1)
            ReDim _LED_R9_COLOUR_RENDERING_INDEX(dummy - 1)
            ReDim _LED_SURVIVAL_FACTOR(dummy - 1)
            ReDim _LED_LUMEN_MAINTENANCE_FACTOR(dummy - 1)
            ReDim _LED_MLS_DISPLACEMENT_FACTOR(dummy - 1)
            ReDim _LED_MLS_COLOUR_CONSISTENCY(dummy - 1)
            ReDim _LED_MLS_CLAIM_LED_REPLACE_FLUORESCENT(dummy - 1)
            ReDim _LED_MLS_FL_REPLACEMENT_CLAIM(dummy - 1)
            ReDim _LED_MLS_FLICKER_METRIC(dummy - 1)
            ReDim _LED_MLS_STROBOSCOPIC_EFFECT_METRIC(dummy - 1)
            Dim provider As CultureInfo = New CultureInfo("en-EN")
            Dim dmy1 As Date
            Dim dmy2 As Double

            For i = 1 To dummy - 1
                _EPREL_MODEL_REGISTRATION_NUMBER(i - 1) = xltab1.Range("A" & i + 1).Value
                _MODEL_IDENTIFIER(i - 1) = xltab1.Range("B" & i + 1).Value
                _CONSIDER_GENERATED_LABEL_AS_PROVIDED(i - 1) = xltab1.Range("C" & i + 1).Value
                dmy1 = xltab1.Range("D" & i + 1).Value
                '-Format date to yyyy-mm-dd+hh:mm
                _ON_MARKET_START_DATE(i - 1) = dmy1.ToString("yyyy") & "-" & dmy1.ToString("MM") & "-" & dmy1.ToString("dd") & dmy1.ToString("zzz")
                _VISIBLE_TO_UK_MSA(i - 1) = xltab1.Range("E" & i + 1).Value
                _LIGHTING_TECHNOLOGY(i - 1) = xltab1.Range("F" & i + 1).Value
                _DIRECTIONAL(i - 1) = xltab1.Range("G" & i + 1).Value
                _CAP_TYPE(i - 1) = xltab1.Range("H" & i + 1).Value
                _MAINS(i - 1) = xltab1.Range("I" & i + 1).Value
                _CONNECTED_LIGHT_SOURCE(i - 1) = xltab1.Range("J" & i + 1).Value
                _COLOUR_TUNEABLE_LIGHT_SOURCE(i - 1) = xltab1.Range("K" & i + 1).Value
                _ENVELOPE(i - 1) = xltab1.Range("L" & i + 1).Value
                _HIGH_LUMINANCE_LIGHT_SOURCE(i - 1) = xltab1.Range("M" & i + 1).Value
                _ANTI_GLARE_SHIELD(i - 1) = xltab1.Range("N" & i + 1).Value
                _DIMMABLE(i - 1) = xltab1.Range("O" & i + 1).Value
                _ENERGY_CONS_ON_MODE(i - 1) = Math.Round(Convert.ToDecimal(xltab1.Range("P" & i + 1).Value))
                '_ENERGY_CONS_ON_MODE(i - 1) = String.Format("{0000}", xltab1.Range("P" & i + 1).Value)
                _ENERGY_CLASS(i - 1) = xltab1.Range("Q" & i + 1).Value
                _LUMINOUS_FLUX(i - 1) = Math.Round(Convert.ToDecimal(xltab1.Range("S" & i + 1).Value))
                '_LUMINOUS_FLUX(i - 1) = String.Format("{00000}", xltab1.Range("S" & i + 1).Value)
                _BEAM_ANGLE_CORRESPONDENCE(i - 1) = xltab1.Range("T" & i + 1).Value
                _CORRELATED_COLOUR_TEMP_TYPE(i - 1) = xltab1.Range("U" & i + 1).Value
                'CCT - Single
                dmy2 = Math.Ceiling(xltab1.Range("V" & i + 1).Value / 100)
                _CORRELATED_COLOUR_TEMP_SINGLE(i - 1) = String.Format("{0000}", dmy2 * 100)
                'CCT - Range
                'MIN
                dmy2 = Math.Ceiling(xltab1.Range("W" & i + 1).Value / 100)
                _CORRELATED_COLOUR_TEMP_MIN(i - 1) = String.Format("{00000}", dmy2 * 100)
                'MAX
                dmy2 = Math.Ceiling(xltab1.Range("X" & i + 1).Value / 100)
                _CORRELATED_COLOUR_TEMP_MAX(i - 1) = String.Format("{00000}", dmy2 * 100)
                'CCT - Steps
                dmy2 = Math.Ceiling(xltab1.Range("Y" & i + 1).Value / 100)
                _CORRELATED_COLOUR_TEMP_1(i - 1) = String.Format("{00000}", dmy2 * 100)
                dmy2 = Math.Ceiling(xltab1.Range("Z" & i + 1).Value / 100)
                _CORRELATED_COLOUR_TEMP_2(i - 1) = String.Format("{00000}", dmy2 * 100)
                dmy2 = Math.Ceiling(xltab1.Range("AA" & i + 1).Value / 100)
                _CORRELATED_COLOUR_TEMP_3(i - 1) = String.Format("{00000}", dmy2 * 100)
                dmy2 = Math.Ceiling(xltab1.Range("AB" & i + 1).Value / 100)
                _CORRELATED_COLOUR_TEMP_4(i - 1) = String.Format("{00000}", dmy2 * 100)
                dmy2 = xltab1.Range("AC" & i + 1).Value
                _POWER_ON_MODE(i - 1) = String.Format(provider, "{0:###0.0}", dmy2)
                dmy2 = xltab1.Range("AD" & i + 1).Value
                _POWER_STANDBY(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
                dmy2 = xltab1.Range("AE" & i + 1).Value
                _POWER_STANDBY_NETWORKED(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
                dmy2 = xltab1.Range("AF" & i + 1).Value
                _COLOUR_RENDERING_INDEX(i - 1) = String.Format(provider, "{0:###}", dmy2)
                dmy2 = xltab1.Range("AG" & i + 1).Value
                _MIN_COLOUR_RENDERING_INDEX(i - 1) = String.Format(provider, "{0:###}", dmy2)
                dmy2 = xltab1.Range("AH" & i + 1).Value
                _MAX_COLOUR_RENDERING_INDEX(i - 1) = String.Format(provider, "{0:###}", dmy2)
                dmy2 = xltab1.Range("AI" & i + 1).Value
                _DIMENSION_HEIGHT(i - 1) = String.Format(provider, "{0:#####}", dmy2)
                dmy2 = xltab1.Range("AJ" & i + 1).Value
                _DIMENSION_WIDTH(i - 1) = String.Format(provider, "{0:#####}", dmy2)
                dmy2 = xltab1.Range("AK" & i + 1).Value
                _DIMENSION_DEPTH(i - 1) = String.Format(provider, "{0:#####}", dmy2)
                _SPECTRAL_POWER_DISTRIBUTION_IMAGE(i - 1) = xltab1.Range("AL" & i + 1).Value
                _CLAIM_EQUIVALENT_POWER(i - 1) = xltab1.Range("AM" & i + 1).Value
                dmy2 = xltab1.Range("AN" & i + 1).Value
                _EQUIVALENT_POWER(i - 1) = String.Format(provider, "{0:#####}", dmy2)
                dmy2 = xltab1.Range("AO" & i + 1).Value
                _CHROMATICITY_COORD_X(i - 1) = String.Format(provider, "{0:0.000}", dmy2)
                dmy2 = xltab1.Range("AP" & i + 1).Value
                _CHROMATICITY_COORD_Y(i - 1) = String.Format(provider, "{0:0.000}", dmy2)
                dmy2 = xltab1.Range("AQ" & i + 1).Value
                _DLS_PEAK_LUMINOUS_INTENSITY(i - 1) = String.Format(provider, "{0:######}", dmy2)
                dmy2 = xltab1.Range("AR" & i + 1).Value
                _DLS_BEAM_ANGLE(i - 1) = String.Format(provider, "{0:###}", dmy2)
                dmy2 = xltab1.Range("AS" & i + 1).Value
                _DLS_MIN_BEAM_ANGLE(i - 1) = String.Format(provider, "{0:###}", dmy2)
                dmy2 = xltab1.Range("AT" & i + 1).Value
                _DLS_MAX_BEAM_ANGLE(i - 1) = String.Format(provider, "{0:###}", dmy2)

                _LED_R9_COLOUR_RENDERING_INDEX(i - 1) = Math.Round(Convert.ToDecimal(xltab1.Range("AU" & i + 1).Value))
                'dmy2 = xltab1.Range("AU" & i + 1).Value
                '_LED_R9_COLOUR_RENDERING_INDEX(i - 1) = String.Format(provider, "{0:###}", dmy2)

                dmy2 = xltab1.Range("AV" & i + 1).Value
                _LED_SURVIVAL_FACTOR(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
                dmy2 = xltab1.Range("AW" & i + 1).Value
                _LED_LUMEN_MAINTENANCE_FACTOR(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
                dmy2 = xltab1.Range("AX" & i + 1).Value
                _LED_MLS_DISPLACEMENT_FACTOR(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
                dmy2 = xltab1.Range("AY" & i + 1).Value
                _LED_MLS_COLOUR_CONSISTENCY(i - 1) = String.Format(provider, "{0:#}", dmy2)
                _LED_MLS_CLAIM_LED_REPLACE_FLUORESCENT(i - 1) = xltab1.Range("AZ" & i + 1).Value
                dmy2 = xltab1.Range("BA" & i + 1).Value
                _LED_MLS_FL_REPLACEMENT_CLAIM(i - 1) = String.Format(provider, "{0:##}", dmy2)
                dmy2 = xltab1.Range("BB" & i + 1).Value
                _LED_MLS_FLICKER_METRIC(i - 1) = String.Format(provider, "{0:0.0}", dmy2)
                dmy2 = xltab1.Range("BC" & i + 1).Value
                _LED_MLS_STROBOSCOPIC_EFFECT_METRIC(i - 1) = String.Format(provider, "{0:0.0}", dmy2)

            Next

            'dummy2 = book.Worksheets("attachments").Range("A" & xltab2.Rows.Count).End(xlUP).Row
            dummy2 = xltab2.Range("A" & xltab2.Rows.Count).End(xlUP).Row
            lastentry = xltab2.Range("A1:A" & dummy2).Value

            dummy2 = dummy2 - 2
            ReDim _TD(dummy2)
            For i = 0 To dummy2
                _TD(i)._TD_MODEL_IDENTIFIER = xltab2.Range("A" & i + 2).Value
                _TD(i)._TD_DESCRIPTION = xltab2.Range("B" & i + 2).Value
                _TD(i)._TD_LANGUAGE = xltab2.Range("C" & i + 2).Value
                _TD(i)._TD_ADDITIONAL_PART = xltab2.Range("D" & i + 2).Value
                _TD(i)._TD_CALCULATIONS = xltab2.Range("E" & i + 2).Value
                _TD(i)._TD_GENERAL_DESCRIPTION = xltab2.Range("F" & i + 2).Value
                _TD(i)._TD_MESURED_TECHNICAL_PARAMETERS = xltab2.Range("G" & i + 2).Value
                _TD(i)._TD_REFERENCES_TO_HARMONIZED_STANDARDS = xltab2.Range("H" & i + 2).Value
                _TD(i)._TD_TESTING_CONDITIONS = xltab2.Range("I" & i + 2).Value
                _TD(i)._TD_SPECIFIC_PRECAUTIONS = xltab2.Range("J" & i + 2).Value
                _TD(i)._TD_FILE_NAME = xltab2.Range("K" & i + 2).Value
            Next

        Catch ex As Exception
            ErrorDlg("parse", ex)
        End Try
        xlApp.ActiveWorkbook.Close(False)
        xlApp.Quit()


    End Sub
    Sub PARSE_DEOP(ByRef quelle As String)
        Dim xlApp As New Excel.Application
        Dim book = xlApp.Workbooks.Open(quelle)
        Dim xltab1 = book.Worksheets("END_OF_PLACEMENT")
        Dim xlUP As Object = Excel.XlDirection.xlUp
        Dim lastentry As Object

        Try
            dummy = xltab1.Range("A" & xltab1.Rows.Count).End(xlUP).Row


            lastentry = xltab1.Range("A1:A" & dummy).Value
            ReDim _EPREL_MODEL_REGISTRATION_NUMBER(dummy - 1)
            ReDim _MODEL_IDENTIFIER(dummy - 1)
            ReDim _ON_MARKET_START_DATE(dummy - 1)
            ReDim _ON_MARKET_END_DATE(dummy - 1)
            Dim provider As CultureInfo = New CultureInfo("en-EN")
            Dim dmy1 As Date


            For i = 1 To dummy - 1
                _EPREL_MODEL_REGISTRATION_NUMBER(i - 1) = xltab1.Range("A" & i + 1).Value
                _MODEL_IDENTIFIER(i - 1) = xltab1.Range("B" & i + 1).Value
                dmy1 = xltab1.Range("C" & i + 1).Value
                '-Format date to yyyy-mm-dd+hh:mm
                _ON_MARKET_START_DATE(i - 1) = dmy1.ToString("yyyy") & "-" & dmy1.ToString("MM") & "-" & dmy1.ToString("dd") & dmy1.ToString("zzz")
                dmy1 = xltab1.Range("D" & i + 1).Value
                '-Format date to yyyy-mm-dd+hh:mm
                _ON_MARKET_END_DATE(i - 1) = dmy1.ToString("yyyy") & "-" & dmy1.ToString("MM") & "-" & dmy1.ToString("dd") & dmy1.ToString("zzz")

            Next

        Catch ex As Exception
            ErrorDlg("parse", ex)
        End Try
        xlApp.ActiveWorkbook.Close(False)
        xlApp.Quit()
    End Sub

    '---Generating XML
    Public Sub REGISTRATION()

        Try
            '---Declaration -M
            Dim decl As XDeclaration = New XDeclaration(encoding:="UTF-8", standalone:="yes", version:="1.0")
            doc.Declaration = decl

            '---Registration -M
            Dim REGISTRATION As XElement = <ns3:ProductModelRegistrationRequest xmlns:ns2="http://eprel.ener.ec.europa.eu/productModel/productCore/v2" REQUEST_ID="nothing"/>

            '---Request -M
            Dim REQUEST_ID As XAttribute = REGISTRATION.Attribute("REQUEST_ID")
            REQUEST_ID.Value = Txt_Request.Text

            For i = 0 To dummy - 2
                '---product Operation
                Dim productOperation As XElement = <productOperation OPERATION_TYPE="nothing" OPERATION_ID="nothing"/>
                REGISTRATION.Add(productOperation)
                Dim OPERATION_TYPE As XAttribute = productOperation.Attribute("OPERATION_TYPE")
                OPERATION_TYPE.Value = CB_OperationType.SelectedItem
                Dim OPERATION_ID As XAttribute = productOperation.Attribute("OPERATION_ID")
                OPERATION_ID.Value = i

                '---Model Version -M
                Dim MODEL_VERSION As XElement = <MODEL_VERSION/>
                productOperation.Add(MODEL_VERSION)

                '---Model Identifier -M
                Dim MODEL_IDENTIFIER As XElement = <MODEL_IDENTIFIER/>
                MODEL_IDENTIFIER.Value = _MODEL_IDENTIFIER(i)
                MODEL_VERSION.Add(MODEL_IDENTIFIER)

                '---Supplier -M
                If CB_Trademark.Checked = True Then
                    Dim SUPPLIER_NAME_OR_TRADEMARK As XElement = <SUPPLIER_NAME_OR_TRADEMARK/>
                    SUPPLIER_NAME_OR_TRADEMARK.Value = Txt_TrademarkRef.Text
                    MODEL_VERSION.Add(SUPPLIER_NAME_OR_TRADEMARK)
                Else
                    Dim TRADEMARK_REFERENCE As XElement = <TRADEMARK_REFERENCE/>
                    TRADEMARK_REFERENCE.Value = Txt_TrademarkRef.Text
                    MODEL_VERSION.Add(TRADEMARK_REFERENCE)
                End If

                '---Delegated Act -M
                Dim DELEGATED_ACT As XElement = <DELEGATED_ACT/>
                DELEGATED_ACT.Value = "EU_2019_2015"
                MODEL_VERSION.Add(DELEGATED_ACT)

                '---Product Group
                Dim PRODUCT_GROUP As XElement = <PRODUCT_GROUP/>
                PRODUCT_GROUP.Value = "LAMP"
                MODEL_VERSION.Add(PRODUCT_GROUP)

                '---Energy Label 
                If _CONSIDER_GENERATED_LABEL_AS_PROVIDED(i) <> "" Then
                    Dim ENERGY_LABEL As XElement = <ENERGY_LABEL xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:ns5="http://eprel.ener.ec.europa.eu/commonTypes/EnergyLabelTypes/v2" xsi:type="ns5:GeneratedEnergyLabel"/>
                    Dim CONSIDER_GENERATED_LABEL_AS_PROVIDED As XElement = <CONSIDER_GENERATED_LABEL_AS_PROVIDED/>

                    CONSIDER_GENERATED_LABEL_AS_PROVIDED.Value = _CONSIDER_GENERATED_LABEL_AS_PROVIDED(i)
                    ENERGY_LABEL.Add(CONSIDER_GENERATED_LABEL_AS_PROVIDED)
                    MODEL_VERSION.Add(ENERGY_LABEL)
                Else
                    Form2.LB_Log.Items.Add("Energy Label for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If


                '---Market Start Date YYYY-MM-DD
                Dim ON_MARKET_START_DATE As XElement = <ON_MARKET_START_DATE/>
                ON_MARKET_START_DATE.Value = _ON_MARKET_START_DATE(i)
                MODEL_VERSION.Add(ON_MARKET_START_DATE)

                '---Registrant Nature
                Dim REGISTRANT_NATURE As XElement = <REGISTRANT_NATURE/>
                REGISTRANT_NATURE.Value = CB_RegistrantNature.SelectedItem
                MODEL_VERSION.Add(REGISTRANT_NATURE)

                '---UK MSA
                Dim VISIBLE_TO_UK_MSA As XElement = <VISIBLE_TO_UK_MSA/>
                VISIBLE_TO_UK_MSA.Value = _VISIBLE_TO_UK_MSA(i)
                MODEL_VERSION.Add(VISIBLE_TO_UK_MSA)

                '---technical Documentation
                Try
                    Dim flag As Boolean = False
                    Dim test As String = MODEL_IDENTIFIER.Value
                    Dim TECHNICAL_DOCUMENTATION As XElement = <TECHNICAL_DOCUMENTATION xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="ns2:TechnicalDocumentationDetail"/>
                    For j = 0 To dummy2
                        If _TD(j)._TD_MODEL_IDENTIFIER = test Then
                            Dim DOCUMENT As XElement = <DOCUMENT/>

                            Dim DESCRIPTION As XElement = <ns2:DESCRIPTION/>
                            DESCRIPTION.Value = _TD(j)._TD_DESCRIPTION
                            DOCUMENT.Add(DESCRIPTION)

                            Dim lng As String = _TD(j)._TD_LANGUAGE

                            For Each elmnt In lng.Split(";")
                                Dim LANGUAGE As XElement = <LANGUAGE/>
                                LANGUAGE.Value = elmnt
                                DOCUMENT.Add(New XElement(LANGUAGE))
                            Next

                            Dim TECHNICAL_PART As XElement = <TECHNICAL_PART/>

                            If _TD(j)._TD_ADDITIONAL_PART = True Then
                                TECHNICAL_PART.Value = "ADDITIONAL_PART"
                                DOCUMENT.Add(New XElement(TECHNICAL_PART))
                            End If

                            If _TD(j)._TD_CALCULATIONS = True Then
                                TECHNICAL_PART.Value = "CALCULATIONS"
                                DOCUMENT.Add(New XElement(TECHNICAL_PART))
                            End If

                            If _TD(j)._TD_GENERAL_DESCRIPTION = True Then
                                TECHNICAL_PART.Value = "GENERAL_DESCRIPTION"
                                DOCUMENT.Add(New XElement(TECHNICAL_PART))
                            End If

                            If _TD(j)._TD_MESURED_TECHNICAL_PARAMETERS = True Then
                                TECHNICAL_PART.Value = "MESURED_TECHNICAL_PARAMETERS"
                                DOCUMENT.Add(New XElement(TECHNICAL_PART))
                            End If

                            If _TD(j)._TD_REFERENCES_TO_HARMONIZED_STANDARDS = True Then
                                TECHNICAL_PART.Value = "REFERENCES_TO_HARMONISED_STANDARDS"
                                DOCUMENT.Add(New XElement(TECHNICAL_PART))
                            End If

                            If _TD(j)._TD_TESTING_CONDITIONS = True Then
                                TECHNICAL_PART.Value = "TESTING_CONDITIONS"
                                DOCUMENT.Add(New XElement(TECHNICAL_PART))
                            End If

                            If _TD(j)._TD_SPECIFIC_PRECAUTIONS = True Then
                                TECHNICAL_PART.Value = "SPECIFIC_PRECAUTIONS"
                                DOCUMENT.Add(New XElement(TECHNICAL_PART))
                            End If

                            Dim FILE_PATH As XElement = <FILE_PATH/>
                            FILE_PATH.Value = "/attachments/" & _TD(j)._TD_FILE_NAME
                            DOCUMENT.Add(FILE_PATH)

                            TECHNICAL_DOCUMENTATION.Add(DOCUMENT)
                            flag = True
                        End If
                    Next
                    If flag = True Then
                        MODEL_VERSION.Add(TECHNICAL_DOCUMENTATION)
                    End If
                Catch ex As Exception
                    errorstate = True
                    Form2.LB_Log.Items.Add("technical Information for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    Continue For
                End Try



                '---Kontakt
                Select Case Form_Contact.CB_ContactDetails.Checked

                    Case False
                        '---Contact Details
                        Dim CONTACT_DETAILS As XElement = <CONTACT_DETAILS xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="ns2:ContactByReference"/>
                        '--- Contact Reference
                        Dim CONTACT_REFERENCE As XElement = <CONTACT_REFERENCE/>
                        CONTACT_REFERENCE.Value = Txt_ContactRef.Text
                        CONTACT_DETAILS.Add(CONTACT_REFERENCE)
                        MODEL_VERSION.Add(CONTACT_DETAILS)

                    Case True
                        '---Contact details
                        Dim CONTACT_DETAILS As XElement = <CONTACT_DETAILS xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="ns2:ModelSpecificContactDetails"/>
                        Dim CONTACT_NAME As XElement = <CONTACT_NAME/>
                        CONTACT_NAME.Value = Form_Contact.TB_ContactName.Text
                        CONTACT_DETAILS.Add(CONTACT_NAME)

                        With Form_Contact
                            If .TB_StreetName.Text <> "" Or .TB_Number.Text <> "" Or .TB_City.Text <> "" Or .TB_Municipality.Text <> "" Or .TB_Province.Text <> "" Or .TB_Postcode.Text <> "" Or .CBox_Country.SelectedItem <> "" Then
                                Dim ADDRESS As XElement = <ADDRESS xmlns:ns5="http://eprel.ener.ec.europa.eu/commonTypes/baseTypes/v1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="ns5:DetailedAddress"/>

                                If .TB_StreetName.Text <> "" Then
                                    Dim STREET_NAME As XElement = <STREET_NAME/>
                                    STREET_NAME.Value = .TB_StreetName.Text
                                    ADDRESS.Add(STREET_NAME)
                                End If

                                If .TB_Number.Text <> "" Then
                                    Dim STREET_NUMBER As XElement = <STREET_NUMBER/>
                                    STREET_NUMBER.Value = .TB_Number.Text
                                    ADDRESS.Add(STREET_NUMBER)
                                End If

                                If .TB_City.Text <> "" Then
                                    Dim CITY As XElement = <CITY/>
                                    CITY.Value = .TB_City.Text
                                    ADDRESS.Add(CITY)
                                End If

                                If .TB_Municipality.Text <> "" Then
                                    Dim MUNICIPALITY As XElement = <MUNICIPALITY/>
                                    MUNICIPALITY.Value = .TB_Municipality.Text
                                    ADDRESS.Add(MUNICIPALITY)
                                End If

                                If .TB_Province.Text <> "" Then
                                    Dim PROVINCE As XElement = <PROVINCE/>
                                    PROVINCE.Value = .TB_Province.Text
                                    ADDRESS.Add(PROVINCE)
                                End If

                                If .TB_Postcode.Text <> "" Then
                                    Dim POSTCODE As XElement = <POSTCODE/>
                                    POSTCODE.Value = .TB_Postcode.Text
                                    ADDRESS.Add(POSTCODE)
                                End If

                                If .CBox_Country.SelectedItem <> "" Then
                                    Dim COUNTRY As XElement = <COUNTRY/>
                                    COUNTRY.Value = .CBox_Country.SelectedItem
                                    ADDRESS.Add(COUNTRY)
                                End If
                                CONTACT_DETAILS.Add(ADDRESS)
                            End If
                        End With

                        Dim FIRST_NAME As XElement = <FIRST_NAME/>
                        FIRST_NAME.Value = Form_Contact.TB_FirstName.Text
                        CONTACT_DETAILS.Add(FIRST_NAME)

                        Dim LAST_NAME As XElement = <LAST_NAME/>
                        LAST_NAME.Value = Form_Contact.TB_LastName.Text
                        CONTACT_DETAILS.Add(LAST_NAME)

                        Dim PHONE_NUMBER As XElement = <PHONE_NUMBER/>
                        PHONE_NUMBER.Value = Form_Contact.TB_PhoneNumber.Text
                        CONTACT_DETAILS.Add(PHONE_NUMBER)

                        If Form_Contact.TB_Email.Text <> "" Then
                            Dim EMAIL_ADDRESS As XElement = <EMAIL_ADDRESS/>
                            EMAIL_ADDRESS.Value = Form_Contact.TB_Email.Text
                            CONTACT_DETAILS.Add(EMAIL_ADDRESS)
                        End If

                        If Form_Contact.TB_URL.Text <> "" Then
                            Dim URL As XElement = <URL/>
                            URL.Value = Form_Contact.TB_URL.Text
                            CONTACT_DETAILS.Add(URL)
                        End If


                        MODEL_VERSION.Add(CONTACT_DETAILS)
                End Select




                '---Product Group Detail
                Dim PRODUCT_GROUP_DETAIL As XElement = <PRODUCT_GROUP_DETAIL xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:ns5="http://eprel.ener.ec.europa.eu/productModel/productGroups/lightsource/v1" xsi:type="ns5:LightSource"/>

                '---Lighting technology
                Dim LIGHTING_TECHNOLOGY As XElement = <LIGHTING_TECHNOLOGY/>
                LIGHTING_TECHNOLOGY.Value = _LIGHTING_TECHNOLOGY(i)
                PRODUCT_GROUP_DETAIL.Add(LIGHTING_TECHNOLOGY)

                '---Captype
                If _CAP_TYPE(i) <> "" Then
                    Dim CAP_TYPE As XElement = <CAP_TYPE/>
                    CAP_TYPE.Value = _CAP_TYPE(i)
                    PRODUCT_GROUP_DETAIL.Add(CAP_TYPE)
                Else
                    Form2.LB_Log.Items.Add("Cap Type for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exeption Occured")
                    errorstate = True
                    Continue For
                End If

                '---Directional
                If _DIRECTIONAL(i) <> "" Then
                    Dim DIRECTIONAL As XElement = <DIRECTIONAL/>
                    DIRECTIONAL.Value = _DIRECTIONAL(i)
                    PRODUCT_GROUP_DETAIL.Add(DIRECTIONAL)
                Else
                    Form2.LB_Log.Items.Add("Direction for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                '---Mains
                If _MAINS(i) <> "" Then
                    Dim MAINS As XElement = <MAINS/>
                    MAINS.Value = _MAINS(i)
                    PRODUCT_GROUP_DETAIL.Add(MAINS)
                Else
                    Form2.LB_Log.Items.Add("MAINS for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                '---Connected lightsource
                If _CONNECTED_LIGHT_SOURCE(i) <> "" Then
                    Dim CONNECTED_LIGHT_SOURCE As XElement = <CONNECTED_LIGHT_SOURCE/>
                    CONNECTED_LIGHT_SOURCE.Value = _CONNECTED_LIGHT_SOURCE(i)
                    PRODUCT_GROUP_DETAIL.Add(CONNECTED_LIGHT_SOURCE)
                Else
                    Form2.LB_Log.Items.Add(" for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _COLOUR_TUNEABLE_LIGHT_SOURCE(i) <> "" Then
                    Dim COLOUR_TUNEABLE_LIGHT_SOURCE As XElement = <COLOUR_TUNEABLE_LIGHT_SOURCE/>
                    COLOUR_TUNEABLE_LIGHT_SOURCE.Value = _COLOUR_TUNEABLE_LIGHT_SOURCE(i)
                    PRODUCT_GROUP_DETAIL.Add(COLOUR_TUNEABLE_LIGHT_SOURCE)
                Else
                    Form2.LB_Log.Items.Add("CTLS for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                '---Envelope
                If _LIGHTING_TECHNOLOGY(i) = "MIXED" Or _LIGHTING_TECHNOLOGY(i) = "OTHER_HID" Then
                    Select Case _ENVELOPE(i)
                        Case "NO"
                            Dim ENVELOPE As XElement = <ENVELOPE/>
                            ENVELOPE.Value = "NO"
                            PRODUCT_GROUP_DETAIL.Add(ENVELOPE)
                        Case "SECOND"
                            Dim ENVELOPE As XElement = <ENVELOPE/>
                            ENVELOPE.Value = "SECOND"
                            PRODUCT_GROUP_DETAIL.Add(ENVELOPE)
                        Case "NON_CLEAR"
                            Dim ENVELOPE As XElement = <ENVELOPE/>
                            ENVELOPE.Value = "NON_CLEAR"
                            PRODUCT_GROUP_DETAIL.Add(ENVELOPE)
                        Case Else
                            errorstate = True
                            Form2.LB_Log.Items.Add("Envelope is missing for Modelidentifier" & _MODEL_IDENTIFIER(i) & " is missing!")
                            Continue For
                    End Select
                End If

                '---High luminance Light source
                If _HIGH_LUMINANCE_LIGHT_SOURCE(i) <> "" Then
                    Dim HIGH_LUMINANCE_LIGHT_SOURCE As XElement = <HIGH_LUMINANCE_LIGHT_SOURCE/>
                    HIGH_LUMINANCE_LIGHT_SOURCE.Value = _HIGH_LUMINANCE_LIGHT_SOURCE(i)
                    PRODUCT_GROUP_DETAIL.Add(HIGH_LUMINANCE_LIGHT_SOURCE)
                Else
                    Form2.LB_Log.Items.Add("High luminance for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _ANTI_GLARE_SHIELD(i) <> "" Then
                    Dim ANTI_GLARE_SHIELD As XElement = <ANTI_GLARE_SHIELD/>
                    ANTI_GLARE_SHIELD.Value = _ANTI_GLARE_SHIELD(i)
                    PRODUCT_GROUP_DETAIL.Add(ANTI_GLARE_SHIELD)
                Else
                    Form2.LB_Log.Items.Add("Anti Glare shield for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _DIMMABLE(i) <> "" Then
                    Dim DIMMABLE As XElement = <DIMMABLE/>
                    DIMMABLE.Value = _DIMMABLE(i)
                    PRODUCT_GROUP_DETAIL.Add(DIMMABLE)
                Else
                    Form2.LB_Log.Items.Add("Dimmable for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _ENERGY_CONS_ON_MODE(i) <> "" Then
                    Dim ENERGY_CONS_ON_MODE As XElement = <ENERGY_CONS_ON_MODE/>
                    ENERGY_CONS_ON_MODE.Value = _ENERGY_CONS_ON_MODE(i)
                    PRODUCT_GROUP_DETAIL.Add(ENERGY_CONS_ON_MODE)
                Else
                    Form2.LB_Log.Items.Add("Energy consumption for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _ENERGY_CLASS(i) <> "" Then
                    Dim ENERGY_CLASS As XElement = <ENERGY_CLASS/>
                    ENERGY_CLASS.Value = _ENERGY_CLASS(i)
                    PRODUCT_GROUP_DETAIL.Add(ENERGY_CLASS)
                Else
                    Form2.LB_Log.Items.Add("Energyclass for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _LUMINOUS_FLUX(i) <> "" Then
                    Dim LUMINOUS_FLUX As XElement = <LUMINOUS_FLUX/>
                    LUMINOUS_FLUX.Value = _LUMINOUS_FLUX(i)
                    PRODUCT_GROUP_DETAIL.Add(LUMINOUS_FLUX)
                Else
                    Form2.LB_Log.Items.Add("Luminus flux for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _BEAM_ANGLE_CORRESPONDENCE(i) <> "" Then
                    Dim BEAM_ANGLE_CORRESPONDENCE As XElement = <BEAM_ANGLE_CORRESPONDENCE/>
                    BEAM_ANGLE_CORRESPONDENCE.Value = _BEAM_ANGLE_CORRESPONDENCE(i)
                    PRODUCT_GROUP_DETAIL.Add(BEAM_ANGLE_CORRESPONDENCE)
                Else
                    Form2.LB_Log.Items.Add("Beam angle correspondence for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If


                If _CORRELATED_COLOUR_TEMP_TYPE(i) <> "" Then

                    Dim CORRELATED_COLOUR_TEMP_TYPE As XElement = <CORRELATED_COLOUR_TEMP_TYPE/>
                    CORRELATED_COLOUR_TEMP_TYPE.Value = _CORRELATED_COLOUR_TEMP_TYPE(i)
                    PRODUCT_GROUP_DETAIL.Add(CORRELATED_COLOUR_TEMP_TYPE)

                    Dim CORRELATED_COLOUR_TEMP As XElement = <CORRELATED_COLOUR_TEMP/>
                    Select Case CORRELATED_COLOUR_TEMP_TYPE.Value
                        Case "SINGLE_VALUE"
                            CORRELATED_COLOUR_TEMP.Value = _CORRELATED_COLOUR_TEMP_SINGLE(i)
                            PRODUCT_GROUP_DETAIL.Add(CORRELATED_COLOUR_TEMP)
                        Case "STEPS"
                            'Dim CORRELATED_COLOUR_TEMP As XElement = <CORRELATED_COLOUR_TEMP/>
                            CORRELATED_COLOUR_TEMP.Value = _CORRELATED_COLOUR_TEMP_1(i)
                            PRODUCT_GROUP_DETAIL.Add(New XElement(CORRELATED_COLOUR_TEMP))
                            CORRELATED_COLOUR_TEMP.Value = _CORRELATED_COLOUR_TEMP_2(i)
                            PRODUCT_GROUP_DETAIL.Add(New XElement(CORRELATED_COLOUR_TEMP))
                            CORRELATED_COLOUR_TEMP.Value = _CORRELATED_COLOUR_TEMP_3(i)
                            PRODUCT_GROUP_DETAIL.Add(New XElement(CORRELATED_COLOUR_TEMP))
                            CORRELATED_COLOUR_TEMP.Value = _CORRELATED_COLOUR_TEMP_4(i)
                            PRODUCT_GROUP_DETAIL.Add(New XElement(CORRELATED_COLOUR_TEMP))
                        Case "RANGE"
                            CORRELATED_COLOUR_TEMP.Value = _CORRELATED_COLOUR_TEMP_MIN(i)
                            PRODUCT_GROUP_DETAIL.Add(New XElement(CORRELATED_COLOUR_TEMP))
                            CORRELATED_COLOUR_TEMP.Value = _CORRELATED_COLOUR_TEMP_MAX(i)
                            PRODUCT_GROUP_DETAIL.Add(New XElement(CORRELATED_COLOUR_TEMP))
                    End Select

                Else
                    Form2.LB_Log.Items.Add("Correlated colour temperature type for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If


                If _POWER_ON_MODE(i) <> "" Then
                    Dim POWER_ON_MODE As XElement = <POWER_ON_MODE/>
                    POWER_ON_MODE.Value = _POWER_ON_MODE(i)

                    PRODUCT_GROUP_DETAIL.Add(POWER_ON_MODE)
                Else
                    Form2.LB_Log.Items.Add("Power for on mode for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _POWER_STANDBY(i) <> "" Then
                    Dim POWER_STANDBY As XElement = <POWER_STANDBY/>
                    POWER_STANDBY.Value = _POWER_STANDBY(i)
                    PRODUCT_GROUP_DETAIL.Add(POWER_STANDBY)
                Else
                    Form2.LB_Log.Items.Add("Standby power for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If


                If _CONNECTED_LIGHT_SOURCE(i) <> "" Then
                    If _CONNECTED_LIGHT_SOURCE(i) = "true" Then
                        Dim POWER_STANDBY_NETWORKED As XElement = <POWER_STANDBY_NETWORKED/>
                        POWER_STANDBY_NETWORKED.Value = _POWER_STANDBY_NETWORKED(i)
                        PRODUCT_GROUP_DETAIL.Add(POWER_STANDBY_NETWORKED)
                    End If
                Else
                    Form2.LB_Log.Items.Add("Standby networked power for on mode for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _COLOUR_RENDERING_INDEX(i) <> "" Then
                    Dim COLOUR_RENDERING_INDEX As XElement = <COLOUR_RENDERING_INDEX/>
                    COLOUR_RENDERING_INDEX.Value = _COLOUR_RENDERING_INDEX(i)
                    PRODUCT_GROUP_DETAIL.Add(COLOUR_RENDERING_INDEX)
                End If


                If _MIN_COLOUR_RENDERING_INDEX(i) <> "" Then
                    Dim MIN_COLOUR_RENDERING_INDEX As XElement = <MIN_COLOUR_RENDERING_INDEX/>
                    MIN_COLOUR_RENDERING_INDEX.Value = _MIN_COLOUR_RENDERING_INDEX(i)
                    PRODUCT_GROUP_DETAIL.Add(MIN_COLOUR_RENDERING_INDEX)
                End If

                If _MAX_COLOUR_RENDERING_INDEX(i) <> "" Then
                    Dim MAX_COLOUR_RENDERING_INDEX As XElement = <MAX_COLOUR_RENDERING_INDEX/>
                    MAX_COLOUR_RENDERING_INDEX.Value = _MAX_COLOUR_RENDERING_INDEX(i)
                    PRODUCT_GROUP_DETAIL.Add(MAX_COLOUR_RENDERING_INDEX)
                End If

                If _DIMENSION_HEIGHT(i) <> "" Then
                    Dim DIMENSION_HEIGHT As XElement = <DIMENSION_HEIGHT/>
                    DIMENSION_HEIGHT.Value = _DIMENSION_HEIGHT(i)
                    PRODUCT_GROUP_DETAIL.Add(DIMENSION_HEIGHT)
                Else
                    Form2.LB_Log.Items.Add("Dimension height for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If


                If _DIMENSION_WIDTH(i) <> "" Then
                    Dim DIMENSION_WIDTH As XElement = <DIMENSION_WIDTH/>
                    DIMENSION_WIDTH.Value = _DIMENSION_WIDTH(i)
                    PRODUCT_GROUP_DETAIL.Add(DIMENSION_WIDTH)
                Else
                    Form2.LB_Log.Items.Add("Dimension width for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _DIMENSION_DEPTH(i) <> "" Then
                    Dim DIMENSION_DEPTH As XElement = <DIMENSION_DEPTH/>
                    DIMENSION_DEPTH.Value = _DIMENSION_DEPTH(i)
                    PRODUCT_GROUP_DETAIL.Add(DIMENSION_DEPTH)
                Else
                    Form2.LB_Log.Items.Add("Dimension depth for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _SPECTRAL_POWER_DISTRIBUTION_IMAGE(i) <> "" Then
                    Dim SPECTRAL_POWER_DISTRIBUTION_IMAGE As XElement = <SPECTRAL_POWER_DISTRIBUTION_IMAGE/>
                    SPECTRAL_POWER_DISTRIBUTION_IMAGE.Value = "/attachments/" & _SPECTRAL_POWER_DISTRIBUTION_IMAGE(i)
                    PRODUCT_GROUP_DETAIL.Add(SPECTRAL_POWER_DISTRIBUTION_IMAGE)
                Else
                    Form2.LB_Log.Items.Add("Spectral power distribution image for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                '---Claim equivalent power
                If _CLAIM_EQUIVALENT_POWER(i) <> "" Then
                    Dim CLAIM_EQUIVALENT_POWER As XElement = <CLAIM_EQUIVALENT_POWER/>
                    CLAIM_EQUIVALENT_POWER.Value = _CLAIM_EQUIVALENT_POWER(i)
                    PRODUCT_GROUP_DETAIL.Add(CLAIM_EQUIVALENT_POWER)
                Else
                    form2.LB_Log.Items.Add("Claim equivalent power for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    errorstate = True
                    Continue For
                End If
                '---Equivalent Power
                Select Case _CLAIM_EQUIVALENT_POWER(i)
                    Case "true"
                        Dim EQUIVALENT_POWER As XElement = <EQUIVALENT_POWER/>
                        EQUIVALENT_POWER.Value = _EQUIVALENT_POWER(i)
                        PRODUCT_GROUP_DETAIL.Add(EQUIVALENT_POWER)
                    Case "false"
                        Exit Select
                    Case Else
                        Form2.LB_Log.Items.Add("Equivalent power for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                        errorstate = True
                        Continue For
                End Select



                If _CHROMATICITY_COORD_X(i) <> "" Then
                    Dim CHROMATICITY_COORD_X As XElement = <CHROMATICITY_COORD_X/>
                    CHROMATICITY_COORD_X.Value = _CHROMATICITY_COORD_X(i)
                    PRODUCT_GROUP_DETAIL.Add(CHROMATICITY_COORD_X)
                Else
                    Form2.LB_Log.Items.Add("Chromaticity coordinate for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _CHROMATICITY_COORD_Y(i) <> "" Then
                    Dim CHROMATICITY_COORD_Y As XElement = <CHROMATICITY_COORD_Y/>
                    CHROMATICITY_COORD_Y.Value = _CHROMATICITY_COORD_Y(i)
                    PRODUCT_GROUP_DETAIL.Add(CHROMATICITY_COORD_Y)
                Else
                    Form2.LB_Log.Items.Add("Chromaticity coordinate for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                '---IF DLS
                If _DIRECTIONAL(i) = "DLS" Then
                    If _DLS_PEAK_LUMINOUS_INTENSITY(i) <> "" Then
                        Dim PEAK_LUMINOUS_INTENSITY As XElement = <PEAK_LUMINOUS_INTENSITY/>
                        PEAK_LUMINOUS_INTENSITY.Value = _DLS_PEAK_LUMINOUS_INTENSITY(i)
                        PRODUCT_GROUP_DETAIL.Add(PEAK_LUMINOUS_INTENSITY)
                    Else
                        Form2.LB_Log.Items.Add("Peak luminous intensity for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                        'Throw New ArgumentException("Exception Occured")
                        errorstate = True
                        Continue For
                    End If

                    If _DLS_BEAM_ANGLE(i) <> "" Then
                        Dim BEAM_ANGLE As XElement = <BEAM_ANGLE/>
                        BEAM_ANGLE.Value = _DLS_BEAM_ANGLE(i)
                        PRODUCT_GROUP_DETAIL.Add(BEAM_ANGLE)
                    Else
                        Form2.LB_Log.Items.Add("Beam angle for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                        'Throw New ArgumentException("Exception Occured")
                        errorstate = True
                        Continue For
                    End If

                    If _DLS_BEAM_ANGLE(i) = "" Then
                        If _DLS_MIN_BEAM_ANGLE(i) <> "" Then
                            Dim MIN_BEAM_ANGLE As XElement = <MIN_BEAM_ANGLE/>
                            MIN_BEAM_ANGLE.Value = _DLS_MIN_BEAM_ANGLE(i)
                            PRODUCT_GROUP_DETAIL.Add(MIN_BEAM_ANGLE)
                        Else
                            Form2.LB_Log.Items.Add("Min beam angle for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                            'Throw New ArgumentException("Exception Occured")
                            errorstate = True
                            Continue For
                        End If

                        If _DLS_MAX_BEAM_ANGLE(i) <> "" Then
                            Dim MAX_BEAM_ANGLE As XElement = <MAX_BEAM_ANGLE/>
                            MAX_BEAM_ANGLE.Value = _DLS_MAX_BEAM_ANGLE(i)
                            PRODUCT_GROUP_DETAIL.Add(MAX_BEAM_ANGLE)
                        Else
                            Form2.LB_Log.Items.Add("Max beam angle for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                            'Throw New ArgumentException("Exception Occured")
                            errorstate = True
                            Continue For
                        End If
                    End If

                End If

                If _LIGHTING_TECHNOLOGY(i) = "LED" Or _LIGHTING_TECHNOLOGY(i) = "OLED" Then
                    If _LED_R9_COLOUR_RENDERING_INDEX(i) <> "" Then
                        Dim R9_COLOUR_RENDERING_INDEX As XElement = <R9_COLOUR_RENDERING_INDEX/>
                        R9_COLOUR_RENDERING_INDEX.Value = _LED_R9_COLOUR_RENDERING_INDEX(i)
                        PRODUCT_GROUP_DETAIL.Add(R9_COLOUR_RENDERING_INDEX)
                    Else
                        Form2.LB_Log.Items.Add("R9 Value for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                        'Throw New ArgumentException("Exception Occured")
                        errorstate = True
                        Continue For
                    End If

                    If _LED_SURVIVAL_FACTOR(i) <> "" Then
                        Dim SURVIVAL_FACTOR As XElement = <SURVIVAL_FACTOR/>
                        SURVIVAL_FACTOR.Value = _LED_SURVIVAL_FACTOR(i)
                        PRODUCT_GROUP_DETAIL.Add(SURVIVAL_FACTOR)
                    Else
                        Form2.LB_Log.Items.Add("Survival Factor for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                        'Throw New ArgumentException("Exception Occured")
                        errorstate = True
                        Continue For
                    End If

                    If _LED_LUMEN_MAINTENANCE_FACTOR(i) <> "" Then
                        Dim LUMEN_MAINTENANCE_FACTOR As XElement = <LUMEN_MAINTENANCE_FACTOR/>
                        LUMEN_MAINTENANCE_FACTOR.Value = _LED_LUMEN_MAINTENANCE_FACTOR(i)
                        PRODUCT_GROUP_DETAIL.Add(LUMEN_MAINTENANCE_FACTOR)
                    Else
                        Form2.LB_Log.Items.Add("Lumen maintnance factor for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                        'Throw New ArgumentException("Exception Occured")
                        errorstate = True
                        Continue For
                    End If

                    If _MAINS(i) = "MLS" Then
                        If _LED_MLS_DISPLACEMENT_FACTOR(i) <> "" Then
                            Dim DISPLACEMENT_FACTOR As XElement = <DISPLACEMENT_FACTOR/>
                            DISPLACEMENT_FACTOR.Value = _LED_MLS_DISPLACEMENT_FACTOR(i)
                            PRODUCT_GROUP_DETAIL.Add(DISPLACEMENT_FACTOR)
                        Else
                            Form2.LB_Log.Items.Add("Displacementfactor for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                            'Throw New ArgumentException("Exception Occured")
                            errorstate = True
                            Continue For
                        End If

                        If _LED_MLS_COLOUR_CONSISTENCY(i) <> "" Then
                            Dim COLOUR_CONSISTENCY As XElement = <COLOUR_CONSISTENCY/>
                            COLOUR_CONSISTENCY.Value = _LED_MLS_COLOUR_CONSISTENCY(i)
                            PRODUCT_GROUP_DETAIL.Add(COLOUR_CONSISTENCY)
                        Else
                            Form2.LB_Log.Items.Add("Colour consistency for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                            'Throw New ArgumentException("Exception Occured")
                            errorstate = True
                            Continue For
                        End If

                        Select Case _LED_MLS_CLAIM_LED_REPLACE_FLUORESCENT(i)
                            Case "true"
                                Dim CLAIM_LED_REPLACE_FLOURESCENT As XElement = <CLAIM_LED_REPLACE_FLOURESCENT/>
                                CLAIM_LED_REPLACE_FLOURESCENT.Value = _LED_MLS_CLAIM_LED_REPLACE_FLUORESCENT(i)
                                PRODUCT_GROUP_DETAIL.Add(CLAIM_LED_REPLACE_FLOURESCENT)

                                If _LED_MLS_FL_REPLACEMENT_CLAIM(i) <> "" Then
                                    Dim REPLACEMENT_CLAIM As XElement = <REPLACEMENT_CLAIM/>
                                    REPLACEMENT_CLAIM.Value = _LED_MLS_FL_REPLACEMENT_CLAIM(i)
                                    PRODUCT_GROUP_DETAIL.Add(REPLACEMENT_CLAIM)
                                Else
                                    Form2.LB_Log.Items.Add("Replacement claim for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                                    'Throw New ArgumentException("Exception Occured")
                                    errorstate = True
                                    Continue For
                                End If

                            Case "false"
                                Dim CLAIM_LED_REPLACE_FLUORESCENT As XElement = <CLAIM_LED_REPLACE_FLUORESCENT/>
                                CLAIM_LED_REPLACE_FLUORESCENT.Value = _LED_MLS_CLAIM_LED_REPLACE_FLUORESCENT(i)
                                PRODUCT_GROUP_DETAIL.Add(CLAIM_LED_REPLACE_FLUORESCENT)
                            Case Else
                                Form2.LB_Log.Items.Add("Replacement claim for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                                'Throw New ArgumentException("Exception Occured")
                                errorstate = True
                                Continue For
                        End Select

                        If _LED_MLS_FLICKER_METRIC(i) <> "" Then
                            Dim FLICKER_METRIC As XElement = <FLICKER_METRIC/>
                            FLICKER_METRIC.Value = _LED_MLS_FLICKER_METRIC(i)
                            PRODUCT_GROUP_DETAIL.Add(FLICKER_METRIC)
                        Else
                            Form2.LB_Log.Items.Add("Flicker Metric for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                            'Throw New ArgumentException("Exception Occured")
                            errorstate = True
                            Continue For
                        End If


                        If _LED_MLS_STROBOSCOPIC_EFFECT_METRIC(i) <> "" Then
                            Dim STROBOSCOPIC_EFFECT_METRIC As XElement = <STROBOSCOPIC_EFFECT_METRIC/>
                            STROBOSCOPIC_EFFECT_METRIC.Value = _LED_MLS_STROBOSCOPIC_EFFECT_METRIC(i)
                            PRODUCT_GROUP_DETAIL.Add(STROBOSCOPIC_EFFECT_METRIC)
                        Else
                            Form2.LB_Log.Items.Add("Stroboscopic effect for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                            'Throw New ArgumentException("Exception Occured")
                            errorstate = True
                            Continue For
                        End If


                    End If


                End If

                MODEL_VERSION.Add(PRODUCT_GROUP_DETAIL)

            Next

            doc.Add(REGISTRATION)

#If DEBUG Then
            Console.WriteLine("Display the modified XML...")
            Console.WriteLine(doc)
            doc.Save(Console.Out)
#End If

        Catch ex As Exception
            ErrorDlg("xml", ex)
        End Try

        If errorstate = True Then
            ErrorDlg("xml")
        End If

    End Sub
    Public Sub UPDATE_PRODUCT()
        Try

            '---Decleration
            Dim decl As XDeclaration = New XDeclaration(encoding:="UTF-8", standalone:="yes", version:="1.0")
            doc.Declaration = decl

            '---Registration
            Dim REGISTRATION As XElement = <ns3:ProductModelRegistrationRequest xmlns:ns2="http://eprel.ener.ec.europa.eu/productModel/productCore/v2" REQUEST_ID="nothing"/>

            '---Request-ID
            Dim REQUEST_ID As XAttribute = REGISTRATION.Attribute("REQUEST_ID")
            REQUEST_ID.Value = Txt_Request.Text

            For i = 0 To dummy - 2
                '---product Operation
                Dim productOperation As XElement = <productOperation OPERATION_TYPE="nothing" OPERATION_ID="nothing" REASON_FOR_CHANGE="nothing"/>
                REGISTRATION.Add(productOperation)
                Dim OPERATION_TYPE As XAttribute = productOperation.Attribute("OPERATION_TYPE")
                OPERATION_TYPE.Value = CB_OperationType.SelectedItem
                Dim REASON_FOR_CHANGE As XAttribute = productOperation.Attribute("REASON_FOR_CHANGE")
                REASON_FOR_CHANGE.Value = CB_ReasonChange.SelectedItem
                Dim OPERATION_ID As XAttribute = productOperation.Attribute("OPERATION_ID")
                OPERATION_ID.Value = i

                '-Model Version
                Dim MODEL_VERSION As XElement = <MODEL_VERSION/>
                productOperation.Add(MODEL_VERSION)

                '---EPREL REGISTRATION Number
                If _EPREL_MODEL_REGISTRATION_NUMBER(i) <> "" Then
                    Dim EPREL_REGISTRATION_NUMBER As XElement = <EPREL_MODEL_REGISTRATION_NUMBER/>
                    EPREL_REGISTRATION_NUMBER.Value = _EPREL_MODEL_REGISTRATION_NUMBER(i)
                    MODEL_VERSION.Add(EPREL_REGISTRATION_NUMBER)
                Else
                    Form2.LB_Log.Items.Add("EPREL Model Registration Number for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exeption Occured")
                    errorstate = True
                    Continue For
                End If

                '---Model Identifier -M
                Dim MODEL_IDENTIFIER As XElement = <MODEL_IDENTIFIER/>
                MODEL_IDENTIFIER.Value = _MODEL_IDENTIFIER(i)
                MODEL_VERSION.Add(MODEL_IDENTIFIER)

                '---Supplier -M
                If CB_Trademark.Checked = True Then
                    Dim SUPPLIER_NAME_OR_TRADEMARK As XElement = <SUPPLIER_NAME_OR_TRADEMARK/>
                    SUPPLIER_NAME_OR_TRADEMARK.Value = Txt_TrademarkRef.Text
                    MODEL_VERSION.Add(SUPPLIER_NAME_OR_TRADEMARK)
                Else
                    Dim TRADEMARK_REFERENCE As XElement = <TRADEMARK_REFERENCE/>
                    TRADEMARK_REFERENCE.Value = Txt_TrademarkRef.Text
                    MODEL_VERSION.Add(TRADEMARK_REFERENCE)
                End If

                '---Delegated Act -M
                Dim DELEGATED_ACT As XElement = <DELEGATED_ACT/>
                DELEGATED_ACT.Value = "EU_2019_2015"
                MODEL_VERSION.Add(DELEGATED_ACT)

                '---Product Group
                Dim PRODUCT_GROUP As XElement = <PRODUCT_GROUP/>
                PRODUCT_GROUP.Value = "LAMP"
                MODEL_VERSION.Add(PRODUCT_GROUP)

                '---Energy Label 
                If _CONSIDER_GENERATED_LABEL_AS_PROVIDED(i) <> "" Then
                    Dim ENERGY_LABEL As XElement = <ENERGY_LABEL xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:ns5="http://eprel.ener.ec.europa.eu/commonTypes/EnergyLabelTypes/v2" xsi:type="ns5:GeneratedEnergyLabel"/>
                    Dim CONSIDER_GENERATED_LABEL_AS_PROVIDED As XElement = <CONSIDER_GENERATED_LABEL_AS_PROVIDED/>

                    CONSIDER_GENERATED_LABEL_AS_PROVIDED.Value = _CONSIDER_GENERATED_LABEL_AS_PROVIDED(i)
                    ENERGY_LABEL.Add(CONSIDER_GENERATED_LABEL_AS_PROVIDED)
                    MODEL_VERSION.Add(ENERGY_LABEL)
                Else
                    Form2.LB_Log.Items.Add("Energy Label for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If


                '---Market Start Date YYYY-MM-DD
                Dim ON_MARKET_START_DATE As XElement = <ON_MARKET_START_DATE/>
                ON_MARKET_START_DATE.Value = _ON_MARKET_START_DATE(i)
                MODEL_VERSION.Add(ON_MARKET_START_DATE)

                '---Registrant Nature
                Dim REGISTRANT_NATURE As XElement = <REGISTRANT_NATURE/>
                REGISTRANT_NATURE.Value = CB_RegistrantNature.SelectedItem
                MODEL_VERSION.Add(REGISTRANT_NATURE)

                '---UK MSA
                Dim VISIBLE_TO_UK_MSA As XElement = <VISIBLE_TO_UK_MSA/>
                VISIBLE_TO_UK_MSA.Value = _VISIBLE_TO_UK_MSA(i)
                MODEL_VERSION.Add(VISIBLE_TO_UK_MSA)

                '---technical Documentation
                Try
                    Dim flag As Boolean = False
                    Dim test As String = MODEL_IDENTIFIER.Value
                    Dim TECHNICAL_DOCUMENTATION As XElement = <TECHNICAL_DOCUMENTATION xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="ns2:TechnicalDocumentationDetail"/>
                    For j = 0 To dummy2
                        If _TD(j)._TD_MODEL_IDENTIFIER = test Then
                            Dim DOCUMENT As XElement = <DOCUMENT/>

                            Dim DESCRIPTION As XElement = <ns2:DESCRIPTION/>
                            DESCRIPTION.Value = _TD(j)._TD_DESCRIPTION
                            DOCUMENT.Add(DESCRIPTION)

                            Dim lng As String = _TD(j)._TD_LANGUAGE

                            For Each elmnt In lng.Split(";")
                                Dim LANGUAGE As XElement = <LANGUAGE/>
                                LANGUAGE.Value = elmnt
                                DOCUMENT.Add(New XElement(LANGUAGE))
                            Next

                            Dim TECHNICAL_PART As XElement = <TECHNICAL_PART/>

                            If _TD(j)._TD_ADDITIONAL_PART = True Then
                                TECHNICAL_PART.Value = "ADDITIONAL_PART"
                                DOCUMENT.Add(New XElement(TECHNICAL_PART))
                            End If

                            If _TD(j)._TD_CALCULATIONS = True Then
                                TECHNICAL_PART.Value = "CALCULATIONS"
                                DOCUMENT.Add(New XElement(TECHNICAL_PART))
                            End If

                            If _TD(j)._TD_GENERAL_DESCRIPTION = True Then
                                TECHNICAL_PART.Value = "GENERAL_DESCRIPTION"
                                DOCUMENT.Add(New XElement(TECHNICAL_PART))
                            End If

                            If _TD(j)._TD_MESURED_TECHNICAL_PARAMETERS = True Then
                                TECHNICAL_PART.Value = "MESURED_TECHNICAL_PARAMETERS"
                                DOCUMENT.Add(New XElement(TECHNICAL_PART))
                            End If

                            If _TD(j)._TD_REFERENCES_TO_HARMONIZED_STANDARDS = True Then
                                TECHNICAL_PART.Value = "REFERENCES_TO_HARMONISED_STANDARDS"
                                DOCUMENT.Add(New XElement(TECHNICAL_PART))
                            End If

                            If _TD(j)._TD_TESTING_CONDITIONS = True Then
                                TECHNICAL_PART.Value = "TESTING_CONDITIONS"
                                DOCUMENT.Add(New XElement(TECHNICAL_PART))
                            End If

                            If _TD(j)._TD_SPECIFIC_PRECAUTIONS = True Then
                                TECHNICAL_PART.Value = "SPECIFIC_PRECAUTIONS"
                                DOCUMENT.Add(New XElement(TECHNICAL_PART))
                            End If

                            Dim FILE_PATH As XElement = <FILE_PATH/>
                            FILE_PATH.Value = "/attachments/" & _TD(j)._TD_FILE_NAME
                            DOCUMENT.Add(FILE_PATH)

                            TECHNICAL_DOCUMENTATION.Add(DOCUMENT)
                            flag = True
                        End If
                    Next
                    If flag = True Then
                        MODEL_VERSION.Add(TECHNICAL_DOCUMENTATION)
                    End If
                Catch ex As Exception
                    errorstate = True
                    Form2.LB_Log.Items.Add("technical Information for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    Continue For
                End Try



                '---Kontakt
                Select Case Form_Contact.CB_ContactDetails.Checked

                    Case False
                        '---Contact Details
                        Dim CONTACT_DETAILS As XElement = <CONTACT_DETAILS xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="ns2:ContactByReference"/>
                        '--- Contact Reference
                        Dim CONTACT_REFERENCE As XElement = <CONTACT_REFERENCE/>
                        CONTACT_REFERENCE.Value = Txt_ContactRef.Text
                        CONTACT_DETAILS.Add(CONTACT_REFERENCE)
                        MODEL_VERSION.Add(CONTACT_DETAILS)

                    Case True
                        '---Contact details
                        Dim CONTACT_DETAILS As XElement = <CONTACT_DETAILS xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="ns2:ModelSpecificContactDetails"/>
                        Dim CONTACT_NAME As XElement = <CONTACT_NAME/>
                        CONTACT_NAME.Value = Form_Contact.TB_ContactName.Text
                        CONTACT_DETAILS.Add(CONTACT_NAME)

                        With Form_Contact
                            If .TB_StreetName.Text <> "" Or .TB_Number.Text <> "" Or .TB_City.Text <> "" Or .TB_Municipality.Text <> "" Or .TB_Province.Text <> "" Or .TB_Postcode.Text <> "" Or .CBox_Country.SelectedItem <> "" Then
                                Dim ADDRESS As XElement = <ADDRESS xmlns:ns5="http://eprel.ener.ec.europa.eu/commonTypes/baseTypes/v1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="ns5:DetailedAddress"/>

                                If .TB_StreetName.Text <> "" Then
                                    Dim STREET_NAME As XElement = <STREET_NAME/>
                                    STREET_NAME.Value = .TB_StreetName.Text
                                    ADDRESS.Add(STREET_NAME)
                                End If

                                If .TB_Number.Text <> "" Then
                                    Dim STREET_NUMBER As XElement = <STREET_NUMBER/>
                                    STREET_NUMBER.Value = .TB_Number.Text
                                    ADDRESS.Add(STREET_NUMBER)
                                End If

                                If .TB_City.Text <> "" Then
                                    Dim CITY As XElement = <CITY/>
                                    CITY.Value = .TB_City.Text
                                    ADDRESS.Add(CITY)
                                End If

                                If .TB_Municipality.Text <> "" Then
                                    Dim MUNICIPALITY As XElement = <MUNICIPALITY/>
                                    MUNICIPALITY.Value = .TB_Municipality.Text
                                    ADDRESS.Add(MUNICIPALITY)
                                End If

                                If .TB_Province.Text <> "" Then
                                    Dim PROVINCE As XElement = <PROVINCE/>
                                    PROVINCE.Value = .TB_Province.Text
                                    ADDRESS.Add(PROVINCE)
                                End If

                                If .TB_Postcode.Text <> "" Then
                                    Dim POSTCODE As XElement = <POSTCODE/>
                                    POSTCODE.Value = .TB_Postcode.Text
                                    ADDRESS.Add(POSTCODE)
                                End If

                                If .CBox_Country.SelectedItem <> "" Then
                                    Dim COUNTRY As XElement = <COUNTRY/>
                                    COUNTRY.Value = .CBox_Country.SelectedItem
                                    ADDRESS.Add(COUNTRY)
                                End If
                                CONTACT_DETAILS.Add(ADDRESS)
                            End If
                        End With

                        Dim FIRST_NAME As XElement = <FIRST_NAME/>
                        FIRST_NAME.Value = Form_Contact.TB_FirstName.Text
                        CONTACT_DETAILS.Add(FIRST_NAME)

                        Dim LAST_NAME As XElement = <LAST_NAME/>
                        LAST_NAME.Value = Form_Contact.TB_LastName.Text
                        CONTACT_DETAILS.Add(LAST_NAME)

                        Dim PHONE_NUMBER As XElement = <PHONE_NUMBER/>
                        PHONE_NUMBER.Value = Form_Contact.TB_PhoneNumber.Text
                        CONTACT_DETAILS.Add(PHONE_NUMBER)

                        If Form_Contact.TB_Email.Text <> "" Then
                            Dim EMAIL_ADDRESS As XElement = <EMAIL_ADDRESS/>
                            EMAIL_ADDRESS.Value = Form_Contact.TB_Email.Text
                            CONTACT_DETAILS.Add(EMAIL_ADDRESS)
                        End If

                        If Form_Contact.TB_URL.Text <> "" Then
                            Dim URL As XElement = <URL/>
                            URL.Value = Form_Contact.TB_URL.Text
                            CONTACT_DETAILS.Add(URL)
                        End If


                        MODEL_VERSION.Add(CONTACT_DETAILS)
                End Select




                '---Product Group Detail
                Dim PRODUCT_GROUP_DETAIL As XElement = <PRODUCT_GROUP_DETAIL xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:ns5="http://eprel.ener.ec.europa.eu/productModel/productGroups/lightsource/v1" xsi:type="ns5:LightSource"/>

                '---Lighting technology
                Dim LIGHTING_TECHNOLOGY As XElement = <LIGHTING_TECHNOLOGY/>
                LIGHTING_TECHNOLOGY.Value = _LIGHTING_TECHNOLOGY(i)
                PRODUCT_GROUP_DETAIL.Add(LIGHTING_TECHNOLOGY)

                '---Captype
                If _CAP_TYPE(i) <> "" Then
                    Dim CAP_TYPE As XElement = <CAP_TYPE/>
                    CAP_TYPE.Value = _CAP_TYPE(i)
                    PRODUCT_GROUP_DETAIL.Add(CAP_TYPE)
                Else
                    Form2.LB_Log.Items.Add("Cap Type for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exeption Occured")
                    errorstate = True
                    Continue For
                End If

                '---Directional
                If _DIRECTIONAL(i) <> "" Then
                    Dim DIRECTIONAL As XElement = <DIRECTIONAL/>
                    DIRECTIONAL.Value = _DIRECTIONAL(i)
                    PRODUCT_GROUP_DETAIL.Add(DIRECTIONAL)
                Else
                    Form2.LB_Log.Items.Add("Direction for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                '---Mains
                If _MAINS(i) <> "" Then
                    Dim MAINS As XElement = <MAINS/>
                    MAINS.Value = _MAINS(i)
                    PRODUCT_GROUP_DETAIL.Add(MAINS)
                Else
                    Form2.LB_Log.Items.Add("MAINS for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                '---Connected lightsource
                If _CONNECTED_LIGHT_SOURCE(i) <> "" Then
                    Dim CONNECTED_LIGHT_SOURCE As XElement = <CONNECTED_LIGHT_SOURCE/>
                    CONNECTED_LIGHT_SOURCE.Value = _CONNECTED_LIGHT_SOURCE(i)
                    PRODUCT_GROUP_DETAIL.Add(CONNECTED_LIGHT_SOURCE)
                Else
                    Form2.LB_Log.Items.Add(" for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _COLOUR_TUNEABLE_LIGHT_SOURCE(i) <> "" Then
                    Dim COLOUR_TUNEABLE_LIGHT_SOURCE As XElement = <COLOUR_TUNEABLE_LIGHT_SOURCE/>
                    COLOUR_TUNEABLE_LIGHT_SOURCE.Value = _COLOUR_TUNEABLE_LIGHT_SOURCE(i)
                    PRODUCT_GROUP_DETAIL.Add(COLOUR_TUNEABLE_LIGHT_SOURCE)
                Else
                    Form2.LB_Log.Items.Add("CTLS for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                '---Envelope
                If _LIGHTING_TECHNOLOGY(i) = "MIXED" Or _LIGHTING_TECHNOLOGY(i) = "OTHER_HID" Then
                    Select Case _ENVELOPE(i)
                        Case "NO"
                            Dim ENVELOPE As XElement = <ENVELOPE/>
                            ENVELOPE.Value = "NO"
                            PRODUCT_GROUP_DETAIL.Add(ENVELOPE)
                        Case "SECOND"
                            Dim ENVELOPE As XElement = <ENVELOPE/>
                            ENVELOPE.Value = "SECOND"
                            PRODUCT_GROUP_DETAIL.Add(ENVELOPE)
                        Case "NON_CLEAR"
                            Dim ENVELOPE As XElement = <ENVELOPE/>
                            ENVELOPE.Value = "NON_CLEAR"
                            PRODUCT_GROUP_DETAIL.Add(ENVELOPE)
                        Case Else
                            errorstate = True
                            Form2.LB_Log.Items.Add("Envelope is missing for Modelidentifier" & _MODEL_IDENTIFIER(i) & " is missing!")
                            Continue For
                    End Select
                End If

                '---High luminance Light source
                If _HIGH_LUMINANCE_LIGHT_SOURCE(i) <> "" Then
                    Dim HIGH_LUMINANCE_LIGHT_SOURCE As XElement = <HIGH_LUMINANCE_LIGHT_SOURCE/>
                    HIGH_LUMINANCE_LIGHT_SOURCE.Value = _HIGH_LUMINANCE_LIGHT_SOURCE(i)
                    PRODUCT_GROUP_DETAIL.Add(HIGH_LUMINANCE_LIGHT_SOURCE)
                Else
                    Form2.LB_Log.Items.Add("High luminance for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _ANTI_GLARE_SHIELD(i) <> "" Then
                    Dim ANTI_GLARE_SHIELD As XElement = <ANTI_GLARE_SHIELD/>
                    ANTI_GLARE_SHIELD.Value = _ANTI_GLARE_SHIELD(i)
                    PRODUCT_GROUP_DETAIL.Add(ANTI_GLARE_SHIELD)
                Else
                    Form2.LB_Log.Items.Add("Anti Glare shield for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _DIMMABLE(i) <> "" Then
                    Dim DIMMABLE As XElement = <DIMMABLE/>
                    DIMMABLE.Value = _DIMMABLE(i)
                    PRODUCT_GROUP_DETAIL.Add(DIMMABLE)
                Else
                    Form2.LB_Log.Items.Add("Dimmable for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _ENERGY_CONS_ON_MODE(i) <> "" Then
                    Dim ENERGY_CONS_ON_MODE As XElement = <ENERGY_CONS_ON_MODE/>
                    ENERGY_CONS_ON_MODE.Value = _ENERGY_CONS_ON_MODE(i)
                    PRODUCT_GROUP_DETAIL.Add(ENERGY_CONS_ON_MODE)
                Else
                    Form2.LB_Log.Items.Add("Energy consumption for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _ENERGY_CLASS(i) <> "" Then
                    Dim ENERGY_CLASS As XElement = <ENERGY_CLASS/>
                    ENERGY_CLASS.Value = _ENERGY_CLASS(i)
                    PRODUCT_GROUP_DETAIL.Add(ENERGY_CLASS)
                Else
                    Form2.LB_Log.Items.Add("Energyclass for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _LUMINOUS_FLUX(i) <> "" Then
                    Dim LUMINOUS_FLUX As XElement = <LUMINOUS_FLUX/>
                    LUMINOUS_FLUX.Value = _LUMINOUS_FLUX(i)
                    PRODUCT_GROUP_DETAIL.Add(LUMINOUS_FLUX)
                Else
                    Form2.LB_Log.Items.Add("Luminus flux for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _BEAM_ANGLE_CORRESPONDENCE(i) <> "" Then
                    Dim BEAM_ANGLE_CORRESPONDENCE As XElement = <BEAM_ANGLE_CORRESPONDENCE/>
                    BEAM_ANGLE_CORRESPONDENCE.Value = _BEAM_ANGLE_CORRESPONDENCE(i)
                    PRODUCT_GROUP_DETAIL.Add(BEAM_ANGLE_CORRESPONDENCE)
                Else
                    Form2.LB_Log.Items.Add("Beam angle correspondence for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If


                If _CORRELATED_COLOUR_TEMP_TYPE(i) <> "" Then

                    Dim CORRELATED_COLOUR_TEMP_TYPE As XElement = <CORRELATED_COLOUR_TEMP_TYPE/>
                    CORRELATED_COLOUR_TEMP_TYPE.Value = _CORRELATED_COLOUR_TEMP_TYPE(i)
                    PRODUCT_GROUP_DETAIL.Add(CORRELATED_COLOUR_TEMP_TYPE)

                    Dim CORRELATED_COLOUR_TEMP As XElement = <CORRELATED_COLOUR_TEMP/>
                    Select Case CORRELATED_COLOUR_TEMP_TYPE.Value
                        Case "SINGLE_VALUE"
                            CORRELATED_COLOUR_TEMP.Value = _CORRELATED_COLOUR_TEMP_SINGLE(i)
                            PRODUCT_GROUP_DETAIL.Add(CORRELATED_COLOUR_TEMP)
                        Case "STEPS"
                            'Dim CORRELATED_COLOUR_TEMP As XElement = <CORRELATED_COLOUR_TEMP/>
                            CORRELATED_COLOUR_TEMP.Value = _CORRELATED_COLOUR_TEMP_1(i)
                            PRODUCT_GROUP_DETAIL.Add(New XElement(CORRELATED_COLOUR_TEMP))
                            CORRELATED_COLOUR_TEMP.Value = _CORRELATED_COLOUR_TEMP_2(i)
                            PRODUCT_GROUP_DETAIL.Add(New XElement(CORRELATED_COLOUR_TEMP))
                            CORRELATED_COLOUR_TEMP.Value = _CORRELATED_COLOUR_TEMP_3(i)
                            PRODUCT_GROUP_DETAIL.Add(New XElement(CORRELATED_COLOUR_TEMP))
                            CORRELATED_COLOUR_TEMP.Value = _CORRELATED_COLOUR_TEMP_4(i)
                            PRODUCT_GROUP_DETAIL.Add(New XElement(CORRELATED_COLOUR_TEMP))
                        Case "RANGE"
                            CORRELATED_COLOUR_TEMP.Value = _CORRELATED_COLOUR_TEMP_MIN(i)
                            PRODUCT_GROUP_DETAIL.Add(New XElement(CORRELATED_COLOUR_TEMP))
                            CORRELATED_COLOUR_TEMP.Value = _CORRELATED_COLOUR_TEMP_MAX(i)
                            PRODUCT_GROUP_DETAIL.Add(New XElement(CORRELATED_COLOUR_TEMP))
                    End Select

                Else
                    Form2.LB_Log.Items.Add("Correlated colour temperature type for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If


                If _POWER_ON_MODE(i) <> "" Then
                    Dim POWER_ON_MODE As XElement = <POWER_ON_MODE/>
                    POWER_ON_MODE.Value = _POWER_ON_MODE(i)

                    PRODUCT_GROUP_DETAIL.Add(POWER_ON_MODE)
                Else
                    Form2.LB_Log.Items.Add("Power for on mode for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _POWER_STANDBY(i) <> "" Then
                    Dim POWER_STANDBY As XElement = <POWER_STANDBY/>
                    POWER_STANDBY.Value = _POWER_STANDBY(i)
                    PRODUCT_GROUP_DETAIL.Add(POWER_STANDBY)
                Else
                    Form2.LB_Log.Items.Add("Standby power for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If


                If _CONNECTED_LIGHT_SOURCE(i) <> "" Then
                    If _CONNECTED_LIGHT_SOURCE(i) = "true" Then
                        Dim POWER_STANDBY_NETWORKED As XElement = <POWER_STANDBY_NETWORKED/>
                        POWER_STANDBY_NETWORKED.Value = _POWER_STANDBY_NETWORKED(i)
                        PRODUCT_GROUP_DETAIL.Add(POWER_STANDBY_NETWORKED)
                    End If
                Else
                    Form2.LB_Log.Items.Add("Standby networked power for on mode for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _COLOUR_RENDERING_INDEX(i) <> "" Then
                    Dim COLOUR_RENDERING_INDEX As XElement = <COLOUR_RENDERING_INDEX/>
                    COLOUR_RENDERING_INDEX.Value = _COLOUR_RENDERING_INDEX(i)
                    PRODUCT_GROUP_DETAIL.Add(COLOUR_RENDERING_INDEX)
                End If


                If _MIN_COLOUR_RENDERING_INDEX(i) <> "" Then
                    Dim MIN_COLOUR_RENDERING_INDEX As XElement = <MIN_COLOUR_RENDERING_INDEX/>
                    MIN_COLOUR_RENDERING_INDEX.Value = _MIN_COLOUR_RENDERING_INDEX(i)
                    PRODUCT_GROUP_DETAIL.Add(MIN_COLOUR_RENDERING_INDEX)
                End If

                If _MAX_COLOUR_RENDERING_INDEX(i) <> "" Then
                    Dim MAX_COLOUR_RENDERING_INDEX As XElement = <MAX_COLOUR_RENDERING_INDEX/>
                    MAX_COLOUR_RENDERING_INDEX.Value = _MAX_COLOUR_RENDERING_INDEX(i)
                    PRODUCT_GROUP_DETAIL.Add(MAX_COLOUR_RENDERING_INDEX)
                End If

                If _DIMENSION_HEIGHT(i) <> "" Then
                    Dim DIMENSION_HEIGHT As XElement = <DIMENSION_HEIGHT/>
                    DIMENSION_HEIGHT.Value = _DIMENSION_HEIGHT(i)
                    PRODUCT_GROUP_DETAIL.Add(DIMENSION_HEIGHT)
                Else
                    Form2.LB_Log.Items.Add("Dimension height for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If


                If _DIMENSION_WIDTH(i) <> "" Then
                    Dim DIMENSION_WIDTH As XElement = <DIMENSION_WIDTH/>
                    DIMENSION_WIDTH.Value = _DIMENSION_WIDTH(i)
                    PRODUCT_GROUP_DETAIL.Add(DIMENSION_WIDTH)
                Else
                    Form2.LB_Log.Items.Add("Dimension width for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _DIMENSION_DEPTH(i) <> "" Then
                    Dim DIMENSION_DEPTH As XElement = <DIMENSION_DEPTH/>
                    DIMENSION_DEPTH.Value = _DIMENSION_DEPTH(i)
                    PRODUCT_GROUP_DETAIL.Add(DIMENSION_DEPTH)
                Else
                    Form2.LB_Log.Items.Add("Dimension depth for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _SPECTRAL_POWER_DISTRIBUTION_IMAGE(i) <> "" Then
                    Dim SPECTRAL_POWER_DISTRIBUTION_IMAGE As XElement = <SPECTRAL_POWER_DISTRIBUTION_IMAGE/>
                    SPECTRAL_POWER_DISTRIBUTION_IMAGE.Value = "/attachments/" & _SPECTRAL_POWER_DISTRIBUTION_IMAGE(i)
                    PRODUCT_GROUP_DETAIL.Add(SPECTRAL_POWER_DISTRIBUTION_IMAGE)
                Else
                    Form2.LB_Log.Items.Add("Spectral power distribution image for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                '---Claim equivalent power
                If _CLAIM_EQUIVALENT_POWER(i) <> "" Then
                    Dim CLAIM_EQUIVALENT_POWER As XElement = <CLAIM_EQUIVALENT_POWER/>
                    CLAIM_EQUIVALENT_POWER.Value = _CLAIM_EQUIVALENT_POWER(i)
                    PRODUCT_GROUP_DETAIL.Add(CLAIM_EQUIVALENT_POWER)
                Else
                    Form2.LB_Log.Items.Add("Claim equivalent power for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    errorstate = True
                    Continue For
                End If
                '---Equivalent Power
                Select Case _CLAIM_EQUIVALENT_POWER(i)
                    Case "true"
                        Dim EQUIVALENT_POWER As XElement = <EQUIVALENT_POWER/>
                        EQUIVALENT_POWER.Value = _EQUIVALENT_POWER(i)
                        PRODUCT_GROUP_DETAIL.Add(EQUIVALENT_POWER)
                    Case "false"
                        Exit Select
                    Case Else
                        Form2.LB_Log.Items.Add("Equivalent power for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                        errorstate = True
                        Continue For
                End Select



                If _CHROMATICITY_COORD_X(i) <> "" Then
                    Dim CHROMATICITY_COORD_X As XElement = <CHROMATICITY_COORD_X/>
                    CHROMATICITY_COORD_X.Value = _CHROMATICITY_COORD_X(i)
                    PRODUCT_GROUP_DETAIL.Add(CHROMATICITY_COORD_X)
                Else
                    Form2.LB_Log.Items.Add("Chromaticity coordinate for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                If _CHROMATICITY_COORD_Y(i) <> "" Then
                    Dim CHROMATICITY_COORD_Y As XElement = <CHROMATICITY_COORD_Y/>
                    CHROMATICITY_COORD_Y.Value = _CHROMATICITY_COORD_Y(i)
                    PRODUCT_GROUP_DETAIL.Add(CHROMATICITY_COORD_Y)
                Else
                    Form2.LB_Log.Items.Add("Chromaticity coordinate for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exception Occured")
                    errorstate = True
                    Continue For
                End If

                '---IF DLS
                If _DIRECTIONAL(i) = "DLS" Then
                    If _DLS_PEAK_LUMINOUS_INTENSITY(i) <> "" Then
                        Dim PEAK_LUMINOUS_INTENSITY As XElement = <PEAK_LUMINOUS_INTENSITY/>
                        PEAK_LUMINOUS_INTENSITY.Value = _DLS_PEAK_LUMINOUS_INTENSITY(i)
                        PRODUCT_GROUP_DETAIL.Add(PEAK_LUMINOUS_INTENSITY)
                    Else
                        Form2.LB_Log.Items.Add("Peak luminous intensity for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                        'Throw New ArgumentException("Exception Occured")
                        errorstate = True
                        Continue For
                    End If

                    If _DLS_BEAM_ANGLE(i) <> "" Then
                        Dim BEAM_ANGLE As XElement = <BEAM_ANGLE/>
                        BEAM_ANGLE.Value = _DLS_BEAM_ANGLE(i)
                        PRODUCT_GROUP_DETAIL.Add(BEAM_ANGLE)
                    Else
                        Form2.LB_Log.Items.Add("Beam angle for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                        'Throw New ArgumentException("Exception Occured")
                        errorstate = True
                        Continue For
                    End If

                    If _DLS_BEAM_ANGLE(i) = "" Then
                        If _DLS_MIN_BEAM_ANGLE(i) <> "" Then
                            Dim MIN_BEAM_ANGLE As XElement = <MIN_BEAM_ANGLE/>
                            MIN_BEAM_ANGLE.Value = _DLS_MIN_BEAM_ANGLE(i)
                            PRODUCT_GROUP_DETAIL.Add(MIN_BEAM_ANGLE)
                        Else
                            Form2.LB_Log.Items.Add("Min beam angle for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                            'Throw New ArgumentException("Exception Occured")
                            errorstate = True
                            Continue For
                        End If

                        If _DLS_MAX_BEAM_ANGLE(i) <> "" Then
                            Dim MAX_BEAM_ANGLE As XElement = <MAX_BEAM_ANGLE/>
                            MAX_BEAM_ANGLE.Value = _DLS_MAX_BEAM_ANGLE(i)
                            PRODUCT_GROUP_DETAIL.Add(MAX_BEAM_ANGLE)
                        Else
                            Form2.LB_Log.Items.Add("Max beam angle for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                            'Throw New ArgumentException("Exception Occured")
                            errorstate = True
                            Continue For
                        End If
                    End If

                End If

                If _LIGHTING_TECHNOLOGY(i) = "LED" Or _LIGHTING_TECHNOLOGY(i) = "OLED" Then
                    If _LED_R9_COLOUR_RENDERING_INDEX(i) <> "" Then
                        Dim R9_COLOUR_RENDERING_INDEX As XElement = <R9_COLOUR_RENDERING_INDEX/>
                        R9_COLOUR_RENDERING_INDEX.Value = _LED_R9_COLOUR_RENDERING_INDEX(i)
                        PRODUCT_GROUP_DETAIL.Add(R9_COLOUR_RENDERING_INDEX)
                    Else
                        Form2.LB_Log.Items.Add("R9 Value for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                        'Throw New ArgumentException("Exception Occured")
                        errorstate = True
                        Continue For
                    End If

                    If _LED_SURVIVAL_FACTOR(i) <> "" Then
                        Dim SURVIVAL_FACTOR As XElement = <SURVIVAL_FACTOR/>
                        SURVIVAL_FACTOR.Value = _LED_SURVIVAL_FACTOR(i)
                        PRODUCT_GROUP_DETAIL.Add(SURVIVAL_FACTOR)
                    Else
                        Form2.LB_Log.Items.Add("Survival Factor for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                        'Throw New ArgumentException("Exception Occured")
                        errorstate = True
                        Continue For
                    End If

                    If _LED_LUMEN_MAINTENANCE_FACTOR(i) <> "" Then
                        Dim LUMEN_MAINTENANCE_FACTOR As XElement = <LUMEN_MAINTENANCE_FACTOR/>
                        LUMEN_MAINTENANCE_FACTOR.Value = _LED_LUMEN_MAINTENANCE_FACTOR(i)
                        PRODUCT_GROUP_DETAIL.Add(LUMEN_MAINTENANCE_FACTOR)
                    Else
                        Form2.LB_Log.Items.Add("Lumen maintnance factor for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                        'Throw New ArgumentException("Exception Occured")
                        errorstate = True
                        Continue For
                    End If

                    If _MAINS(i) = "MLS" Then
                        If _LED_MLS_DISPLACEMENT_FACTOR(i) <> "" Then
                            Dim DISPLACEMENT_FACTOR As XElement = <DISPLACEMENT_FACTOR/>
                            DISPLACEMENT_FACTOR.Value = _LED_MLS_DISPLACEMENT_FACTOR(i)
                            PRODUCT_GROUP_DETAIL.Add(DISPLACEMENT_FACTOR)
                        Else
                            Form2.LB_Log.Items.Add("Displacementfactor for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                            'Throw New ArgumentException("Exception Occured")
                            errorstate = True
                            Continue For
                        End If

                        If _LED_MLS_COLOUR_CONSISTENCY(i) <> "" Then
                            Dim COLOUR_CONSISTENCY As XElement = <COLOUR_CONSISTENCY/>
                            COLOUR_CONSISTENCY.Value = _LED_MLS_COLOUR_CONSISTENCY(i)
                            PRODUCT_GROUP_DETAIL.Add(COLOUR_CONSISTENCY)
                        Else
                            Form2.LB_Log.Items.Add("Colour consistency for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                            'Throw New ArgumentException("Exception Occured")
                            errorstate = True
                            Continue For
                        End If

                        Select Case _LED_MLS_CLAIM_LED_REPLACE_FLUORESCENT(i)
                            Case "true"
                                Dim CLAIM_LED_REPLACE_FLOURESCENT As XElement = <CLAIM_LED_REPLACE_FLOURESCENT/>
                                CLAIM_LED_REPLACE_FLOURESCENT.Value = _LED_MLS_CLAIM_LED_REPLACE_FLUORESCENT(i)
                                PRODUCT_GROUP_DETAIL.Add(CLAIM_LED_REPLACE_FLOURESCENT)

                                If _LED_MLS_FL_REPLACEMENT_CLAIM(i) <> "" Then
                                    Dim REPLACEMENT_CLAIM As XElement = <REPLACEMENT_CLAIM/>
                                    REPLACEMENT_CLAIM.Value = _LED_MLS_FL_REPLACEMENT_CLAIM(i)
                                    PRODUCT_GROUP_DETAIL.Add(REPLACEMENT_CLAIM)
                                Else
                                    Form2.LB_Log.Items.Add("Replacement claim for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                                    'Throw New ArgumentException("Exception Occured")
                                    errorstate = True
                                    Continue For
                                End If

                            Case "false"
                                Dim CLAIM_LED_REPLACE_FLUORESCENT As XElement = <CLAIM_LED_REPLACE_FLUORESCENT/>
                                CLAIM_LED_REPLACE_FLUORESCENT.Value = _LED_MLS_CLAIM_LED_REPLACE_FLUORESCENT(i)
                                PRODUCT_GROUP_DETAIL.Add(CLAIM_LED_REPLACE_FLUORESCENT)
                            Case Else
                                Form2.LB_Log.Items.Add("Replacement claim for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                                'Throw New ArgumentException("Exception Occured")
                                errorstate = True
                                Continue For
                        End Select

                        If _LED_MLS_FLICKER_METRIC(i) <> "" Then
                            Dim FLICKER_METRIC As XElement = <FLICKER_METRIC/>
                            FLICKER_METRIC.Value = _LED_MLS_FLICKER_METRIC(i)
                            PRODUCT_GROUP_DETAIL.Add(FLICKER_METRIC)
                        Else
                            Form2.LB_Log.Items.Add("Flicker Metric for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                            'Throw New ArgumentException("Exception Occured")
                            errorstate = True
                            Continue For
                        End If


                        If _LED_MLS_STROBOSCOPIC_EFFECT_METRIC(i) <> "" Then
                            Dim STROBOSCOPIC_EFFECT_METRIC As XElement = <STROBOSCOPIC_EFFECT_METRIC/>
                            STROBOSCOPIC_EFFECT_METRIC.Value = _LED_MLS_STROBOSCOPIC_EFFECT_METRIC(i)
                            PRODUCT_GROUP_DETAIL.Add(STROBOSCOPIC_EFFECT_METRIC)
                        Else
                            Form2.LB_Log.Items.Add("Stroboscopic effect for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                            'Throw New ArgumentException("Exception Occured")
                            errorstate = True
                            Continue For
                        End If


                    End If


                End If

                MODEL_VERSION.Add(PRODUCT_GROUP_DETAIL)

            Next

            doc.Add(REGISTRATION)

#If DEBUG Then
            Console.WriteLine("Display the modified XML...")
            Console.WriteLine(doc)
            doc.Save(Console.Out)
#End If

        Catch ex As Exception
            ErrorDlg("xml", ex)
        End Try

        If errorstate = True Then
            ErrorDlg("xml")
        End If


    End Sub
    Public Sub PREREGISTRATION()
        Try
            '---Declaration
            Dim decl As XDeclaration = New XDeclaration(encoding:="UTF-8", standalone:="yes", version:="1.0")
            doc.Declaration = decl

            '---Registration
            Dim REGISTRATION As XElement = <ns3:ProductModelRegistrationRequest xmlns:ns2="http://eprel.ener.ec.europa.eu/productModel/productCore/v2" REQUEST_ID="nothing"/>

            '---Request ID
            Dim REQUEST_ID As XAttribute = REGISTRATION.Attribute("REQUEST_ID")
            REQUEST_ID.Value = Txt_Request.Text

            For i = 0 To dummy - 2
                '---Product Operation
                Dim productOperation As XElement = <productOperation OPERATION_TYPE="nothing" OPERATION_ID="nothing"/>
                REGISTRATION.Add(productOperation)
                Dim OPERATION_TYPE As XAttribute = productOperation.Attribute("OPERATION_TYPE")
                OPERATION_TYPE.Value = CB_OperationType.SelectedItem
                Dim OPERATION_ID As XAttribute = productOperation.Attribute("OPERATION_ID")
                OPERATION_ID.Value = i

                '---Model Version
                Dim MODEL_VERSION As XElement = <MODEL_VERSION/>
                productOperation.Add(MODEL_VERSION)

                '---Model Identifier
                Dim MODEL_IDENTIFIER As XElement = <MODEL_IDENTIFIER/>
                MODEL_IDENTIFIER.Value = items(i)
                MODEL_VERSION.Add(MODEL_IDENTIFIER)

                '---Supplier -M
                If CB_Trademark.Checked = True Then
                    Dim SUPPLIER_NAME_OR_TRADEMARK As XElement = <SUPPLIER_NAME_OR_TRADEMARK/>
                    SUPPLIER_NAME_OR_TRADEMARK.Value = Txt_TrademarkRef.Text
                    MODEL_VERSION.Add(SUPPLIER_NAME_OR_TRADEMARK)
                Else
                    Dim TRADEMARK_REFERENCE As XElement = <TRADEMARK_REFERENCE/>
                    TRADEMARK_REFERENCE.Value = Txt_TrademarkRef.Text
                    MODEL_VERSION.Add(TRADEMARK_REFERENCE)
                End If

                '---Delegated Act
                Dim DELEGATED_ACT As XElement = <DELEGATED_ACT/>
                DELEGATED_ACT.Value = "EU_2019_2015"
                MODEL_VERSION.Add(DELEGATED_ACT)

                '---Product Group
                Dim PRODUCT_GROUP As XElement = <PRODUCT_GROUP/>
                PRODUCT_GROUP.Value = "LAMP"
                MODEL_VERSION.Add(PRODUCT_GROUP)

                Form2.LB_Log.Items.Add(MODEL_VERSION.Value + " - Success!")


            Next

            doc.Add(REGISTRATION)

#If DEBUG Then
            Console.WriteLine("Display the modified XML...")
            Console.WriteLine(doc)
            doc.Save(Console.Out)
#End If

        Catch ex As Exception
            ErrorDlg("xml", ex)
        End Try


    End Sub
    Public Sub DECLARE_END_DATE_OF_PLACEMENT_ON_MARKET()
        Try

            '---Decleration
            Dim decl As XDeclaration = New XDeclaration(encoding:="UTF-8", standalone:="yes", version:="1.0")
            doc.Declaration = decl

            '---Registration
            Dim REGISTRATION As XElement = <ns3:ProductModelRegistrationRequest xmlns:ns2="http://eprel.ener.ec.europa.eu/productModel/productCore/v2" REQUEST_ID="nothing"/>

            '---Request-ID
            Dim REQUEST_ID As XAttribute = REGISTRATION.Attribute("REQUEST_ID")
            REQUEST_ID.Value = Txt_Request.Text

            For i = 0 To dummy - 2
                '---product Operation
                Dim productOperation As XElement = <productOperation OPERATION_TYPE="nothing" OPERATION_ID="nothing"/>
                REGISTRATION.Add(productOperation)
                Dim OPERATION_TYPE As XAttribute = productOperation.Attribute("OPERATION_TYPE")
                OPERATION_TYPE.Value = CB_OperationType.SelectedItem
                Dim OPERATION_ID As XAttribute = productOperation.Attribute("OPERATION_ID")
                OPERATION_ID.Value = i

                '-Model Version
                Dim MODEL_VERSION As XElement = <MODEL_VERSION/>
                productOperation.Add(MODEL_VERSION)

                '---EPREL REGISTRATION Number
                If _EPREL_MODEL_REGISTRATION_NUMBER(i) <> "" Then
                    Dim EPREL_REGISTRATION_NUMBER As XElement = <EPREL_MODEL_REGISTRATION_NUMBER/>
                    EPREL_REGISTRATION_NUMBER.Value = _EPREL_MODEL_REGISTRATION_NUMBER(i)
                    MODEL_VERSION.Add(EPREL_REGISTRATION_NUMBER)
                Else
                    Form2.LB_Log.Items.Add("EPREL Model Registration Number for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    'Throw New ArgumentException("Exeption Occured")
                    errorstate = True
                    Continue For
                End If

                '---Model Identifier -M
                Dim MODEL_IDENTIFIER As XElement = <MODEL_IDENTIFIER/>
                MODEL_IDENTIFIER.Value = _MODEL_IDENTIFIER(i)
                MODEL_VERSION.Add(MODEL_IDENTIFIER)

                '---Supplier -M
                If CB_Trademark.Checked = True Then
                    Dim SUPPLIER_NAME_OR_TRADEMARK As XElement = <SUPPLIER_NAME_OR_TRADEMARK/>
                    SUPPLIER_NAME_OR_TRADEMARK.Value = Txt_TrademarkRef.Text
                    MODEL_VERSION.Add(SUPPLIER_NAME_OR_TRADEMARK)
                Else
                    Dim TRADEMARK_REFERENCE As XElement = <TRADEMARK_REFERENCE/>
                    TRADEMARK_REFERENCE.Value = Txt_TrademarkRef.Text
                    MODEL_VERSION.Add(TRADEMARK_REFERENCE)
                End If

                '---Delegated Act -M
                Dim DELEGATED_ACT As XElement = <DELEGATED_ACT/>
                DELEGATED_ACT.Value = "EU_2019_2015"
                MODEL_VERSION.Add(DELEGATED_ACT)

                '---Product Group
                Dim PRODUCT_GROUP As XElement = <PRODUCT_GROUP/>
                PRODUCT_GROUP.Value = "LAMP"
                MODEL_VERSION.Add(PRODUCT_GROUP)

                '---Market Start Date YYYY-MM-DD
                Dim ON_MARKET_START_DATE As XElement = <ON_MARKET_START_DATE/>
                ON_MARKET_START_DATE.Value = _ON_MARKET_START_DATE(i)
                MODEL_VERSION.Add(ON_MARKET_START_DATE)

                '---Market END Date YYYY-MM-DD
                Dim ON_MARKET_END_DATE As XElement = <ON_MARKET_END_DATE/>
                ON_MARKET_END_DATE.Value = _ON_MARKET_END_DATE(i)
                MODEL_VERSION.Add(ON_MARKET_END_DATE)

                ''---Registrant Nature
                'Dim REGISTRANT_NATURE As XElement = <REGISTRANT_NATURE/>
                'REGISTRANT_NATURE.Value = CB_RegistrantNature.SelectedItem
                'MODEL_VERSION.Add(REGISTRANT_NATURE)

            Next

            doc.Add(REGISTRATION)

#If DEBUG Then
            Console.WriteLine("Display the modified XML...")
            Console.WriteLine(doc)
            doc.Save(Console.Out)
#End If

        Catch ex As Exception
            ErrorDlg("xml", ex)
        End Try

        If errorstate = True Then
            ErrorDlg("xml")
        End If
    End Sub

    '---Output
    Public Sub OUTPUT()

#If DEBUG Then
        '---------DEBUG!---------------------
        Dim dir As String = Directory.GetCurrentDirectory
        Directory.CreateDirectory("Data")
        Directory.CreateDirectory("Data\productModelRegistrationTable")
        doc.Save(".\Data\productModelRegistrationTable\registration-data.xml")
        Dim start As String = ".\Data\productModelRegistrationTable"
#Else
        '------RELEASE!--------
        Dim dir As String = Directory.GetCurrentDirectory
        Directory.GetAccessControl(dir + "\Data\")
        Directory.CreateDirectory(dir + "\Data\productModelRegistrationTable")
        doc.Save(dir + "\Data\productModelRegistrationTable\registration-data.xml")

        Dim start As String = ".\Data\productModelRegistrationTable\"
#End If




        Dim ziel As New SaveFileDialog
        ziel.Filter = "zip files (*.zip)|*.zip"
        ziel.FileName = "productModelRegistrationTable.zip"
        ziel.ShowDialog()

        If File.Exists(ziel.FileName) = True Then
            File.Delete(ziel.FileName)
        End If



        If ziel.FileName = "" Then
            MsgBox("Error!")
            Exit Sub
        End If

        If CB_OperationType.SelectedItem = "REGISTER_PRODUCT_MODEL" Or CB_OperationType.SelectedItem = "UPDATE_PRODUCT_MODEL" Then
Select_File:
            Dim SPECTRAL As New FolderBrowserDialog
            SPECTRAL.Description = "Please select folder with attachment data!"

            SPECTRAL.ShowDialog()
            If SPECTRAL.SelectedPath = "" Then
                Select Case MsgBox("Are you shure you do not want to upload any files?", MsgBoxStyle.YesNo)
                    Case MsgBoxResult.Yes
                        GoTo Done
                    Case MsgBoxResult.No
                        GoTo Select_File
                    Case Else
                        GoTo Select_File
                End Select
            End If

            Directory.CreateDirectory(dir + "\Data\productModelRegistrationTable\attachments\")

            Dim fle As String
            Dim target As String = ""

            For Each fle In Directory.GetFiles(SPECTRAL.SelectedPath)
                target = dir & "\Data\productModelRegistrationTable\attachments\" & Path.GetFileName(fle)
                File.Copy(fle, target)
            Next

        End If

Done:
        ZipFile.CreateFromDirectory(start, ziel.FileName)

        '---------------DEBUG!--------------------
        Directory.Delete(dir + "\Data\productModelRegistrationTable", True)

#If DEBUG Then
        Directory.Delete(".\Data", True)
#End If


        MsgBox("Done!")

    End Sub

    '---Validation
    Private Sub Validate_ZIP()
        'Dim zip_File As File
        Dim slct As New OpenFileDialog
        slct.Filter = "EPREL Zip-File (*.zip)|*.zip"
        slct.Title = "Select productModelRegistrationTable.zip for Validation"
        slct.ShowDialog()

        Dim valFile As String = ""

        If slct.FileName <> "" Then
            valFile = slct.FileName
        End If

#If DEBUG Then
        Dim val As String = "cmd.exe /k java -jar ./EprelExchangeModel-2.7.10-SNAPSHOT.jar """
        Dim strCMD As String
        valFile = valFile + """"
        strCMD = val & valFile
        Dim p As New Process
        Dim pi As New ProcessStartInfo
        pi.FileName = "cmd.exe"
        pi.CreateNoWindow = "false"
        pi.Arguments = strCMD
        p.StartInfo = pi
        p.Start()
        p.WaitForExit()
#Else
        Dim val As String = "cmd.exe /k java -jar ./EprelExchangeModel-2.7.10-SNAPSHOT.jar """
        Dim strCMD As String
        valFile = valFile + """"
        strCMD = val & valFile
        Dim p As New Process
        Dim pi As New ProcessStartInfo
        pi.FileName = "cmd.exe"
        pi.CreateNoWindow = "false"
        pi.Arguments = strCMD
        p.StartInfo = pi
        p.Start()
        p.WaitForExit()
#End If
    End Sub

    '---Errors
    Private Sub ErrorDlg(ByVal type As String, Optional ByVal reason As Exception = Nothing, Optional ByVal row As Integer = 0, Optional ByVal col As String = Nothing, Optional sheet As String = Nothing)
        If type = "xml" Then
            MsgBox("Error while processing xml! Please check your files and try again! For detailed information check log.")
        ElseIf type = "parse" Then
            MsgBox("Error while parsing Excel file! Please check your files and try again!" & vbNewLine & "Error: " & reason.Message & vbNewLine & "Please check row " & row & " at column " & col & " on sheet " & sheet & " !")
        Else
            MsgBox("Hard error")
            Close()
        End If

        state = True
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles BT_Tools.Click
        Form3.Show()
    End Sub


    '---Logging
    Private Sub Save_Log_XML()

        Dim result As DialogResult
        result = MessageBox.Show("Do you want to save Log-File?", "Save Log", MessageBoxButtons.YesNo)

        If result = DialogResult.Yes Then
            Dim filesave As New SaveFileDialog
            filesave.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
            filesave.ShowDialog()
            If filesave.FileName = "" Then
                Exit Sub
            Else
                IO.File.WriteAllLines(filesave.FileName, Form2.LB_Log.Items.Cast(Of String).ToArray)
            End If


        ElseIf result = DialogResult.No Then
            Exit Sub
        End If

    End Sub




End Class
