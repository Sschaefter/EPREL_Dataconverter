Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Globalization
Imports System.Xml
Imports System.Xml.Schema
Imports System.Xml.XPath

Imports <xmlns:ns3="http://eprel.ener.ec.europa.eu/services/productModelService/modelRegistrationService/v2">
Imports <xmlns:ns2="http://eprel.ener.ec.europa.eu/productModel/productCore/v2">
Imports <xmlns:ns5="http://eprel.ener.ec.europa.eu/commonTypes/EnergyLabelTypes/v2">


'Imports System.IO
'Imports System.IO.Compression
'Documentation

Public Class Form1
    Public xlApp As New Excel.Application
    Public wb As Excel.Workbook
    Public ws As Excel.Worksheet
    Public wbook As Excel.Workbooks
    Public items() As String
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

    Public dummy As Integer
    'Public doc As XmlDocument = New XmlDocument()
    Public doc As XDocument = New XDocument()
    Public state As Boolean = False

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If CheckB_Log.Checked = True Then

            Form2.Visible = True

        End If

        Cursor.Current = Cursors.WaitCursor

        If CB_OperationType.SelectedItem = "REGISTER_PRODUCT_MODEL" Then
            Form2.LB_Log.Items.Add("PREREGISTER_PRODUCT_MODEL")
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
            Form2.LB_Log.Items.Add("PREREGISTER_PRODUCT_MODEL")
            'SELECT_INPUT()
            If state = True Then
                Exit Sub
            End If
            UPDATE_PRODUCT()
            If state = True Then
                Exit Sub
            End If
            OUTPUT()
        ElseIf CB_OperationType.SelectedItem = "DECLARE_END_DATE_OF_PLACEMENT_ON_MARKET" Then
            MsgBox("Not available in this Version!")
            Exit Sub
        End If

        Cursor.Current = Cursors.Default

        If CheckB_Log.Checked = True Then
            Save_Log_XML()
        End If

        Close()
    End Sub
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

    Public Sub REGISTRATION()

        Dim decl As XDeclaration = New XDeclaration(encoding:="UTF-8", standalone:="yes", version:="1.0")
        doc.Declaration = decl

        Dim REGISTRATION As XElement = <ns3:ProductModelRegistrationRequest xmlns:ns2="http://eprel.ener.ec.europa.eu/productModel/productCore/v2" REQUEST_ID="nothing"/>

        Dim REQUEST_ID As XAttribute = REGISTRATION.Attribute("REQUEST_ID")
        REQUEST_ID.Value = Txt_Request.Text

        For i = 0 To dummy - 2
            '-product Operation
            Dim productOperation As XElement = <productOperation OPERATION_TYPE="nothing" OPERATION_ID="nothing"/>
            REGISTRATION.Add(productOperation)
            Dim OPERATION_TYPE As XAttribute = productOperation.Attribute("OPERATION_TYPE")
            OPERATION_TYPE.Value = CB_OperationType.SelectedItem
            Dim OPERATION_ID As XAttribute = productOperation.Attribute("OPERATION_ID")
            OPERATION_ID.Value = i

            '-Model Verion
            Dim MODEL_VERSION As XElement = <MODEL_VERSION/>
            productOperation.Add(MODEL_VERSION)

            '-Model Identifier
            Dim MODEL_IDENTIFIER As XElement = <MODEL_IDENTIFIER/>
            MODEL_IDENTIFIER.Value = _MODEL_IDENTIFIER(i)
            MODEL_VERSION.Add(MODEL_IDENTIFIER)

            '-Supplier
            Dim SUPPLIER_NAME_OR_TRADEMARK As XElement = <SUPPLIER_NAME_OR_TRADEMARK/>
            SUPPLIER_NAME_OR_TRADEMARK.Value = Txt_TrademarkRef.Text
            MODEL_VERSION.Add(SUPPLIER_NAME_OR_TRADEMARK)

            '-Delegated Act
            Dim DELEGATED_ACT As XElement = <DELEGATED_ACT/>
            DELEGATED_ACT.Value = "EU_2019_2015"
            MODEL_VERSION.Add(DELEGATED_ACT)

            '-Product Group
            Dim PRODUCT_GROUP As XElement = <PRODUCT_GROUP/>
            PRODUCT_GROUP.Value = "LAMP"
            MODEL_VERSION.Add(PRODUCT_GROUP)

            '-Energy Label
            Dim ENERGY_LABEL As XElement = <ENERGY_LABEL xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:ns5="http://eprel.ener.ec.europa.eu/commonTypes/EnergyLabelTypes/v2" xsi:type="ns5:GeneratedEnergyLabel"/>
            Dim CONSIDER_GENERATED_LABEL_AS_PROVIDED As XElement = <CONSIDER_GENERATED_LABEL_AS_PROVIDED/>
            CONSIDER_GENERATED_LABEL_AS_PROVIDED.Value = _CONSIDER_GENERATED_LABEL_AS_PROVIDED(i)
            ENERGY_LABEL.Add(CONSIDER_GENERATED_LABEL_AS_PROVIDED)
            MODEL_VERSION.Add(ENERGY_LABEL)


            '---Market Start Date YYYY-MM-DD
            Dim ON_MARKET_START_DATE As XElement = <ON_MARKET_START_DATE/>
            ON_MARKET_START_DATE.Value = _ON_MARKET_START_DATE(i)
            MODEL_VERSION.Add(ON_MARKET_START_DATE)

            Dim TECHNICAL_DOCUMENTATION As XElement = <TECHNICAL_DOCUMENTATION xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="ns2:TechnicalDocumentationDetail"/>
            Dim DOCUMENT As XElement =
            <DOCUMENT>
                <ns2:DESCRIPTION>test conditions</ns2:DESCRIPTION>
                <LANGUAGE>EN</LANGUAGE>
                <TECHNICAL_PART>TESTING_CONDITIONS</TECHNICAL_PART>
                <TECHNICAL_PART>CALCULATIONS</TECHNICAL_PART>
                <TECHNICAL_PART>GENERAL_DESCRIPTION</TECHNICAL_PART>
                <TECHNICAL_PART>MESURED_TECHNICAL_PARAMETERS</TECHNICAL_PART>
                <TECHNICAL_PART>REFERENCES_TO_HARMONISED_STANDARDS</TECHNICAL_PART>
                <TECHNICAL_PART>SPECIFIC_PRECAUTIONS</TECHNICAL_PART>
                <FILE_PATH>/attachments/testConditions.docx</FILE_PATH>
            </DOCUMENT>

            TECHNICAL_DOCUMENTATION.Add(DOCUMENT)
            MODEL_VERSION.Add(TECHNICAL_DOCUMENTATION)

            '-Kontakt (optional)

            '-Product Group Detail
            Dim PRODUCT_GROUP_DETAIL As XElement = <PRODUCT_GROUP_DETAIL xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:ns5="http://eprel.ener.ec.europa.eu/productModel/productGroups/lightsource/v1" xsi:type="ns5:LightSource"/>
            Dim LIGHTING_TECHNOLOGY As XElement = <LIGHTING_TECHNOLOGY/>
            LIGHTING_TECHNOLOGY.Value = _LIGHTING_TECHNOLOGY(i)
            PRODUCT_GROUP_DETAIL.Add(LIGHTING_TECHNOLOGY)

            Dim CAP_TYPE As XElement = <CAP_TYPE/>
            CAP_TYPE.Value = _CAP_TYPE(i)
            PRODUCT_GROUP_DETAIL.Add(CAP_TYPE)

            Dim DIRECTIONAL As XElement = <DIRECTIONAL/>
            DIRECTIONAL.Value = _DIRECTIONAL(i)
            PRODUCT_GROUP_DETAIL.Add(DIRECTIONAL)

            Dim MAINS As XElement = <MAINS/>
            MAINS.Value = _MAINS(i)
            PRODUCT_GROUP_DETAIL.Add(MAINS)

            Dim CONNECTED_LIGHT_SOURCE As XElement = <CONNECTED_LIGHT_SOURCE/>
            CONNECTED_LIGHT_SOURCE.Value = _CONNECTED_LIGHT_SOURCE(i)
            PRODUCT_GROUP_DETAIL.Add(CONNECTED_LIGHT_SOURCE)

            Dim COLOUR_TUNEABLE_LIGHT_SOURCE As XElement = <COLOUR_TUNEABLE_LIGHT_SOURCE/>
            COLOUR_TUNEABLE_LIGHT_SOURCE.Value = _COLOUR_TUNEABLE_LIGHT_SOURCE(i)
            PRODUCT_GROUP_DETAIL.Add(COLOUR_TUNEABLE_LIGHT_SOURCE)

            Dim HIGH_LUMINANCE_LIGHT_SOURCE As XElement = <HIGH_LUMINANCE_LIGHT_SOURCE/>
            HIGH_LUMINANCE_LIGHT_SOURCE.Value = _HIGH_LUMINANCE_LIGHT_SOURCE(i)
            PRODUCT_GROUP_DETAIL.Add(HIGH_LUMINANCE_LIGHT_SOURCE)

            Dim ANTI_GLARE_SHIELD As XElement = <ANTI_GLARE_SHIELD/>
            ANTI_GLARE_SHIELD.Value = _ANTI_GLARE_SHIELD(i)
            PRODUCT_GROUP_DETAIL.Add(ANTI_GLARE_SHIELD)

            Dim DIMMABLE As XElement = <DIMMABLE/>
            DIMMABLE.Value = _DIMMABLE(i)
            PRODUCT_GROUP_DETAIL.Add(DIMMABLE)

            Dim ENERGY_CONS_ON_MODE As XElement = <ENERGY_CONS_ON_MODE/>
            ENERGY_CONS_ON_MODE.Value = _ENERGY_CONS_ON_MODE(i)
            PRODUCT_GROUP_DETAIL.Add(ENERGY_CONS_ON_MODE)

            Dim ENERGY_CLASS As XElement = <ENERGY_CLASS/>
            ENERGY_CLASS.Value = _ENERGY_CLASS(i)
            PRODUCT_GROUP_DETAIL.Add(ENERGY_CLASS)

            Dim LUMINOUS_FLUX As XElement = <LUMINOUS_FLUX/>
            LUMINOUS_FLUX.Value = _LUMINOUS_FLUX(i)
            PRODUCT_GROUP_DETAIL.Add(LUMINOUS_FLUX)

            Dim BEAM_ANGLE_CORRESPONDENCE As XElement = <BEAM_ANGLE_CORRESPONDENCE/>
            BEAM_ANGLE_CORRESPONDENCE.Value = _BEAM_ANGLE_CORRESPONDENCE(i)
            PRODUCT_GROUP_DETAIL.Add(BEAM_ANGLE_CORRESPONDENCE)

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
                    PRODUCT_GROUP_DETAIL.Add(CORRELATED_COLOUR_TEMP)
                    CORRELATED_COLOUR_TEMP.Value = _CORRELATED_COLOUR_TEMP_2(i)
                    PRODUCT_GROUP_DETAIL.Add(CORRELATED_COLOUR_TEMP)
                    CORRELATED_COLOUR_TEMP.Value = _CORRELATED_COLOUR_TEMP_3(i)
                    PRODUCT_GROUP_DETAIL.Add(CORRELATED_COLOUR_TEMP)
                    CORRELATED_COLOUR_TEMP.Value = _CORRELATED_COLOUR_TEMP_4(i)
                    PRODUCT_GROUP_DETAIL.Add(CORRELATED_COLOUR_TEMP)
                Case "RANGE"
                    CORRELATED_COLOUR_TEMP.Value = _CORRELATED_COLOUR_TEMP_MIN(i)
                    PRODUCT_GROUP_DETAIL.Add(CORRELATED_COLOUR_TEMP)
                    CORRELATED_COLOUR_TEMP.Value = _CORRELATED_COLOUR_TEMP_MAX(i)
                    PRODUCT_GROUP_DETAIL.Add(CORRELATED_COLOUR_TEMP)
            End Select


            Dim POWER_ON_MODE As XElement = <POWER_ON_MODE/>
            POWER_ON_MODE.Value = _POWER_ON_MODE(i)
            PRODUCT_GROUP_DETAIL.Add(POWER_ON_MODE)

            Dim POWER_STANDBY As XElement = <POWER_STANDBY/>
            POWER_STANDBY.Value = _POWER_STANDBY(i)
            PRODUCT_GROUP_DETAIL.Add(POWER_STANDBY)

            If CONNECTED_LIGHT_SOURCE.Value = "true" Then
                Dim POWER_STANDBY_NETWORKED As XElement = <POWER_STANDBY_NETWORKED/>
                POWER_STANDBY_NETWORKED.Value = _POWER_STANDBY_NETWORKED(i)
                PRODUCT_GROUP_DETAIL.Add(POWER_STANDBY_NETWORKED)
            End If

            Dim COLOUR_RENDERING_INDEX As XElement = <COLOUR_RENDERING_INDEX/>
            COLOUR_RENDERING_INDEX.Value = _COLOUR_RENDERING_INDEX(i)
            PRODUCT_GROUP_DETAIL.Add(COLOUR_RENDERING_INDEX)

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

            Dim DIMENSION_HEIGHT As XElement = <DIMENSION_HEIGHT/>
            DIMENSION_HEIGHT.Value = _DIMENSION_HEIGHT(i)
            PRODUCT_GROUP_DETAIL.Add(DIMENSION_HEIGHT)

            Dim DIMENSION_WIDTH As XElement = <DIMENSION_WIDTH/>
            DIMENSION_WIDTH.Value = _DIMENSION_WIDTH(i)
            PRODUCT_GROUP_DETAIL.Add(DIMENSION_WIDTH)

            Dim DIMENSION_DEPTH As XElement = <DIMENSION_DEPTH/>
            DIMENSION_DEPTH.Value = _DIMENSION_DEPTH(i)
            PRODUCT_GROUP_DETAIL.Add(DIMENSION_DEPTH)

            Dim SPECTRAL_POWER_DISTRIBUTION_IMAGE As XElement = <SPECTRAL_POWER_DISTRIBUTION_IMAGE/>
            SPECTRAL_POWER_DISTRIBUTION_IMAGE.Value = "./SPECTRAL/" & _SPECTRAL_POWER_DISTRIBUTION_IMAGE(i)
            PRODUCT_GROUP_DETAIL.Add(SPECTRAL_POWER_DISTRIBUTION_IMAGE)

            Dim CLAIM_EQUIVALENT_POWER As XElement = <CLAIM_EQUIVALENT_POWER/>
            CLAIM_EQUIVALENT_POWER.Value = "false"
            PRODUCT_GROUP_DETAIL.Add(CLAIM_EQUIVALENT_POWER)

            'Dim EQUIVALENT_POWER As XElement = <EQUIVALENT_POWER/>
            'EQUIVALENT_POWER.Value = "1"
            'PRODUCT_GROUP_DETAIL.Add(EQUIVALENT_POWER)

            Dim CHROMATICITY_COORD_X As XElement = <CHROMATICITY_COORD_X/>
            CHROMATICITY_COORD_X.Value = _CHROMATICITY_COORD_X(i)
            PRODUCT_GROUP_DETAIL.Add(CHROMATICITY_COORD_X)

            Dim CHROMATICITY_COORD_Y As XElement = <CHROMATICITY_COORD_Y/>
            CHROMATICITY_COORD_Y.Value = _CHROMATICITY_COORD_Y(i)
            PRODUCT_GROUP_DETAIL.Add(CHROMATICITY_COORD_Y)

            If LIGHTING_TECHNOLOGY.Value = "LED" Then
                Dim R9_COLOUR_RENDERING_INDEX As XElement = <R9_COLOUR_RENDERING_INDEX/>
                R9_COLOUR_RENDERING_INDEX.Value = _LED_R9_COLOUR_RENDERING_INDEX(i)
                PRODUCT_GROUP_DETAIL.Add(R9_COLOUR_RENDERING_INDEX)

                Dim SURVIVAL_FACTOR As XElement = <SURVIVAL_FACTOR/>
                SURVIVAL_FACTOR.Value = _LED_SURVIVAL_FACTOR(i)
                PRODUCT_GROUP_DETAIL.Add(SURVIVAL_FACTOR)

                Dim LUMEN_MAINTENANCE_FACTOR As XElement = <LUMEN_MAINTENANCE_FACTOR/>
                LUMEN_MAINTENANCE_FACTOR.Value = _LED_LUMEN_MAINTENANCE_FACTOR(i)
                PRODUCT_GROUP_DETAIL.Add(LUMEN_MAINTENANCE_FACTOR)

            End If

            If LIGHTING_TECHNOLOGY.Value = "LED" And MAINS.Value = "MLS" Then
                Dim FLICKER_METRIC As XElement = <FLICKER_METRIC/>
                FLICKER_METRIC.Value = _LED_MLS_FLICKER_METRIC(i)
                PRODUCT_GROUP_DETAIL.Add(FLICKER_METRIC)

                Dim STROBOSCOPIC_EFFECT_METRIC As XElement = <STROBOSCOPIC_EFFECT_METRIC/>
                STROBOSCOPIC_EFFECT_METRIC.Value = _LED_MLS_STROBOSCOPIC_EFFECT_METRIC(i)
                PRODUCT_GROUP_DETAIL.Add(STROBOSCOPIC_EFFECT_METRIC)
            End If

            MODEL_VERSION.Add(PRODUCT_GROUP_DETAIL)

        Next

        doc.Add(REGISTRATION)
        Console.WriteLine("Display the modified XML...")
        Console.WriteLine(doc)
        doc.Save(Console.Out)

        'OUTPUT()

    End Sub
    Public Sub UPDATE_PRODUCT()



        Dim decl As XDeclaration = New XDeclaration(encoding:="UTF-8", standalone:="yes", version:="1.0")
        doc.Declaration = decl

        Dim REGISTRATION As XElement = <ns3:ProductModelRegistrationRequest xmlns:ns2="http://eprel.ener.ec.europa.eu/productModel/productCore/v2" REQUEST_ID="nothing"/>

        Dim REQUEST_ID As XAttribute = REGISTRATION.Attribute("REQUEST_ID")
        REQUEST_ID.Value = Txt_Request.Text

        'For i = 0 To dummy - 2
        '-product Operation
        Dim productOperation As XElement = <productOperation OPERATION_TYPE="nothing" OPERATION_ID="nothing" REASON_FOR_CHANGE="nothing"/>
            REGISTRATION.Add(productOperation)
            Dim OPERATION_TYPE As XAttribute = productOperation.Attribute("OPERATION_TYPE")
            OPERATION_TYPE.Value = CB_OperationType.SelectedItem
            Dim OPERATION_ID As XAttribute = productOperation.Attribute("OPERATION_ID")
            OPERATION_ID.Value = "1"
            Dim REASON_FOR_CHANGE As XAttribute = productOperation.Attribute("REASON_FOR_CHANGE")
            REASON_FOR_CHANGE.Value = "CORRECT_TYPO"


            '-Model Version
            Dim MODEL_VERSION As XElement = <MODEL_VERSION/>
            productOperation.Add(MODEL_VERSION)

            '-EPREL_MODEL_REGISTRATION_NUMBER
            Dim EPREL_MODEL_REGISTRATION_NUMBER As XElement = <EPREL_MODEL_REGISTRATION_NUMBER/>
            EPREL_MODEL_REGISTRATION_NUMBER.Value = "12345"
            MODEL_VERSION.Add(EPREL_MODEL_REGISTRATION_NUMBER)

            '-Model Identifier
            Dim MODEL_IDENTIFIER As XElement = <MODEL_IDENTIFIER/>
            MODEL_IDENTIFIER.Value = "Test"
            MODEL_VERSION.Add(MODEL_IDENTIFIER)

            '-Supplier
            Dim SUPPLIER_NAME_OR_TRADEMARK As XElement = <SUPPLIER_NAME_OR_TRADEMARK/>
            SUPPLIER_NAME_OR_TRADEMARK.Value = Txt_TrademarkRef.Text
            MODEL_VERSION.Add(SUPPLIER_NAME_OR_TRADEMARK)

            '-Delegated Act
            Dim DELEGATED_ACT As XElement = <DELEGATED_ACT/>
            DELEGATED_ACT.Value = "EU_2019_2015"
            MODEL_VERSION.Add(DELEGATED_ACT)

            '-Product Group
            Dim PRODUCT_GROUP As XElement = <PRODUCT_GROUP/>
            PRODUCT_GROUP.Value = "LAMP"
            MODEL_VERSION.Add(PRODUCT_GROUP)

            '-Energy Label
            Dim ENERGY_LABEL As XElement = <ENERGY_LABEL xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:ns5="http://eprel.ener.ec.europa.eu/commonTypes/EnergyLabelTypes/v2" xsi:type="ns5:GeneratedEnergyLabel"/>
            Dim CONSIDER_GENERATED_LABEL_AS_PROVIDED As XElement = <CONSIDER_GENERATED_LABEL_AS_PROVIDED/>
            CONSIDER_GENERATED_LABEL_AS_PROVIDED.Value = "true"
            ENERGY_LABEL.Add(CONSIDER_GENERATED_LABEL_AS_PROVIDED)
            MODEL_VERSION.Add(ENERGY_LABEL)


            '---Market Start Date YYYY-MM-DD
            Dim ON_MARKET_START_DATE As XElement = <ON_MARKET_START_DATE/>
            ON_MARKET_START_DATE.Value = "2021-05-01"
            MODEL_VERSION.Add(ON_MARKET_START_DATE)

            Dim TECHNICAL_DOCUMENTATION As XElement = <TECHNICAL_DOCUMENTATION xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="ns2:TechnicalDocumentationDetail"/>
            Dim DOCUMENT As XElement =
            <DOCUMENT>
                <ns2:DESCRIPTION>test conditions</ns2:DESCRIPTION>
                <LANGUAGE>EN</LANGUAGE>
                <TECHNICAL_PART>TESTING_CONDITIONS</TECHNICAL_PART>
                <TECHNICAL_PART>CALCULATIONS</TECHNICAL_PART>
                <TECHNICAL_PART>GENERAL_DESCRIPTION</TECHNICAL_PART>
                <TECHNICAL_PART>MESURED_TECHNICAL_PARAMETERS</TECHNICAL_PART>
                <TECHNICAL_PART>REFERENCES_TO_HARMONISED_STANDARDS</TECHNICAL_PART>
                <TECHNICAL_PART>SPECIFIC_PRECAUTIONS</TECHNICAL_PART>
                <FILE_PATH>/attachments/testConditions.docx</FILE_PATH>
            </DOCUMENT>

            TECHNICAL_DOCUMENTATION.Add(DOCUMENT)
            MODEL_VERSION.Add(TECHNICAL_DOCUMENTATION)

            '-Kontakt (optional)

            '-Product Group Detail
            Dim PRODUCT_GROUP_DETAIL As XElement = <PRODUCT_GROUP_DETAIL xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:ns5="http://eprel.ener.ec.europa.eu/productModel/productGroups/lightsource/v1" xsi:type="ns5:LightSource"/>
            Dim LIGHTING_TECHNOLOGY As XElement = <LIGHTING_TECHNOLOGY/>
            LIGHTING_TECHNOLOGY.Value = "LED"
            PRODUCT_GROUP_DETAIL.Add(LIGHTING_TECHNOLOGY)

            Dim CAP_TYPE As XElement = <CAP_TYPE/>
            CAP_TYPE.Value = "mycaptype"
            PRODUCT_GROUP_DETAIL.Add(CAP_TYPE)

            Dim DIRECTIONAL As XElement = <DIRECTIONAL/>
            DIRECTIONAL.Value = "DLS"
            PRODUCT_GROUP_DETAIL.Add(DIRECTIONAL)

            Dim MAINS As XElement = <MAINS/>
            MAINS.Value = "MLS"
            PRODUCT_GROUP_DETAIL.Add(MAINS)

            Dim CONNECTED_LIGHT_SOURCE As XElement = <CONNECTED_LIGHT_SOURCE/>
            CONNECTED_LIGHT_SOURCE.Value = "true"
            PRODUCT_GROUP_DETAIL.Add(CONNECTED_LIGHT_SOURCE)

            Dim COLOUR_TUNEABLE_LIGHT_SOURCE As XElement = <COLOUR_TUNEABLE_LIGHT_SOURCE/>
            COLOUR_TUNEABLE_LIGHT_SOURCE.Value = "true"
            PRODUCT_GROUP_DETAIL.Add(COLOUR_TUNEABLE_LIGHT_SOURCE)

            Dim HIGH_LUMINANCE_LIGHT_SOURCE As XElement = <HIGH_LUMINANCE_LIGHT_SOURCE/>
            HIGH_LUMINANCE_LIGHT_SOURCE.Value = "true"
            PRODUCT_GROUP_DETAIL.Add(HIGH_LUMINANCE_LIGHT_SOURCE)

            Dim ANTI_GLARE_SHIELD As XElement = <ANTI_GLARE_SHIELD/>
            ANTI_GLARE_SHIELD.Value = "true"
            PRODUCT_GROUP_DETAIL.Add(ANTI_GLARE_SHIELD)

            Dim DIMMABLE As XElement = <DIMMABLE/>
            DIMMABLE.Value = "YES"
            PRODUCT_GROUP_DETAIL.Add(DIMMABLE)

            Dim ENERGY_CONS_ON_MODE As XElement = <ENERGY_CONS_ON_MODE/>
            ENERGY_CONS_ON_MODE.Value = "1"
            PRODUCT_GROUP_DETAIL.Add(ENERGY_CONS_ON_MODE)

            Dim ENERGY_CLASS As XElement = <ENERGY_CLASS/>
            ENERGY_CLASS.Value = "A"
            PRODUCT_GROUP_DETAIL.Add(ENERGY_CLASS)

            Dim LUMINOUS_FLUX As XElement = <LUMINOUS_FLUX/>
            LUMINOUS_FLUX.Value = "100"
            PRODUCT_GROUP_DETAIL.Add(LUMINOUS_FLUX)

            Dim BEAM_ANGLE_CORRESPONDENCE As XElement = <BEAM_ANGLE_CORRESPONDENCE/>
            BEAM_ANGLE_CORRESPONDENCE.Value = "SPHERE_360"
            PRODUCT_GROUP_DETAIL.Add(BEAM_ANGLE_CORRESPONDENCE)

            Dim CORRELATED_COLOUR_TEMP_TYPE As XElement = <CORRELATED_COLOUR_TEMP_TYPE/>
            CORRELATED_COLOUR_TEMP_TYPE.Value = "SINGLE_VALUE"
            PRODUCT_GROUP_DETAIL.Add(CORRELATED_COLOUR_TEMP_TYPE)

            Dim CORRELATED_COLOUR_TEMP As XElement = <CORRELATED_COLOUR_TEMP/>
            CORRELATED_COLOUR_TEMP.Value = "1000"
            PRODUCT_GROUP_DETAIL.Add(CORRELATED_COLOUR_TEMP)

            Dim POWER_ON_MODE As XElement = <POWER_ON_MODE/>
            POWER_ON_MODE.Value = "1.5"
            PRODUCT_GROUP_DETAIL.Add(POWER_ON_MODE)

            Dim POWER_STANDBY As XElement = <POWER_STANDBY/>
            POWER_STANDBY.Value = "0.4"
            PRODUCT_GROUP_DETAIL.Add(POWER_STANDBY)

            Dim COLOUR_RENDERING_INDEX As XElement = <COLOUR_RENDERING_INDEX/>
            COLOUR_RENDERING_INDEX.Value = "90"
            PRODUCT_GROUP_DETAIL.Add(COLOUR_RENDERING_INDEX)

            Dim MIN_COLOUR_RENDERING_INDEX As XElement = <MIN_COLOUR_RENDERING_INDEX/>
            MIN_COLOUR_RENDERING_INDEX.Value = "10"
            PRODUCT_GROUP_DETAIL.Add(MIN_COLOUR_RENDERING_INDEX)

            Dim MAX_COLOUR_RENDERING_INDEX As XElement = <MAX_COLOUR_RENDERING_INDEX/>
            MAX_COLOUR_RENDERING_INDEX.Value = "100"
            PRODUCT_GROUP_DETAIL.Add(MAX_COLOUR_RENDERING_INDEX)

            Dim DIMENSION_HEIGHT As XElement = <DIMENSION_HEIGHT/>
            DIMENSION_HEIGHT.Value = "10"
            PRODUCT_GROUP_DETAIL.Add(DIMENSION_HEIGHT)

            Dim DIMENSION_WIDTH As XElement = <DIMENSION_WIDTH/>
            DIMENSION_WIDTH.Value = "10"
            PRODUCT_GROUP_DETAIL.Add(DIMENSION_WIDTH)

            Dim DIMENSION_DEPTH As XElement = <DIMENSION_DEPTH/>
            DIMENSION_DEPTH.Value = "10"
            PRODUCT_GROUP_DETAIL.Add(DIMENSION_DEPTH)

            Dim SPECTRAL_POWER_DISTRIBUTION_IMAGE As XElement = <SPECTRAL_POWER_DISTRIBUTION_IMAGE/>
            SPECTRAL_POWER_DISTRIBUTION_IMAGE.Value = "./image.png"
            PRODUCT_GROUP_DETAIL.Add(SPECTRAL_POWER_DISTRIBUTION_IMAGE)

        '-if cell empty default false
        Dim CLAIM_EQUIVALENT_POWER As XElement = <CLAIM_EQUIVALENT_POWER/>
            CLAIM_EQUIVALENT_POWER.Value = "true"
            PRODUCT_GROUP_DETAIL.Add(CLAIM_EQUIVALENT_POWER)

            Dim EQUIVALENT_POWER As XElement = <EQUIVALENT_POWER/>
            EQUIVALENT_POWER.Value = "1"
            PRODUCT_GROUP_DETAIL.Add(EQUIVALENT_POWER)
        '---

        Dim CHROMATICITY_COORD_X As XElement = <CHROMATICITY_COORD_X/>
            CHROMATICITY_COORD_X.Value = "0.111"
            PRODUCT_GROUP_DETAIL.Add(CHROMATICITY_COORD_X)

            Dim CHROMATICITY_COORD_Y As XElement = <CHROMATICITY_COORD_Y/>
            CHROMATICITY_COORD_Y.Value = "0.111"
            PRODUCT_GROUP_DETAIL.Add(CHROMATICITY_COORD_Y)

            Dim R9_COLOUR_RENDERING_INDEX As XElement = <R9_COLOUR_RENDERING_INDEX/>
            R9_COLOUR_RENDERING_INDEX.Value = "90"
            PRODUCT_GROUP_DETAIL.Add(R9_COLOUR_RENDERING_INDEX)

            Dim SURVIVAL_FACTOR As XElement = <SURVIVAL_FACTOR/>
            SURVIVAL_FACTOR.Value = "1.04"
            PRODUCT_GROUP_DETAIL.Add(SURVIVAL_FACTOR)

            Dim LUMEN_MAINTENANCE_FACTOR As XElement = <LUMEN_MAINTENANCE_FACTOR/>
            LUMEN_MAINTENANCE_FACTOR.Value = "1.09"
            PRODUCT_GROUP_DETAIL.Add(LUMEN_MAINTENANCE_FACTOR)

            Dim FLICKER_METRIC As XElement = <FLICKER_METRIC/>
            FLICKER_METRIC.Value = "102.8"
            PRODUCT_GROUP_DETAIL.Add(FLICKER_METRIC)

            Dim STROBOSCOPIC_EFFECT_METRIC As XElement = <STROBOSCOPIC_EFFECT_METRIC/>
            STROBOSCOPIC_EFFECT_METRIC.Value = "0.9"
            PRODUCT_GROUP_DETAIL.Add(STROBOSCOPIC_EFFECT_METRIC)

            MODEL_VERSION.Add(PRODUCT_GROUP_DETAIL)

        '    Next


        doc.Add(REGISTRATION)
        Console.WriteLine("Display the modified XML...")

        Console.WriteLine(doc)
        doc.Save(Console.Out)

        'OUTPUT()

    End Sub


    Private Sub PREREGISTRATION()

        'SELECT_INPUT()

        Dim decl As XDeclaration = New XDeclaration(encoding:="UTF-8", standalone:="yes", version:="1.0")
        doc.Declaration = decl



        Dim REGISTRATION As XElement = <ns3:ProductModelRegistrationRequest xmlns:ns2="http://eprel.ener.ec.europa.eu/productModel/productCore/v2" REQUEST_ID="nothing"/>

        Dim REQUEST_ID As XAttribute = REGISTRATION.Attribute("REQUEST_ID")
        REQUEST_ID.Value = Txt_Request.Text

        For i = 0 To dummy - 2
            '-Product Operation
            Dim productOperation As XElement = <productOperation OPERATION_TYPE="nothing" OPERATION_ID="nothing"/>
            REGISTRATION.Add(productOperation)
            Dim OPERATION_TYPE As XAttribute = productOperation.Attribute("OPERATION_TYPE")
            OPERATION_TYPE.Value = CB_OperationType.SelectedItem
            Dim OPERATION_ID As XAttribute = productOperation.Attribute("OPERATION_ID")
            OPERATION_ID.Value = i

            '-Model Verion
            Dim MODEL_VERSION As XElement = <MODEL_VERSION/>
            productOperation.Add(MODEL_VERSION)

            '-Model Identifier
            Dim MODEL_IDENTIFIER As XElement = <MODEL_IDENTIFIER/>
            MODEL_IDENTIFIER.Value = items(i)
            MODEL_VERSION.Add(MODEL_IDENTIFIER)

            '-Supplier
            Dim SUPPLIER_NAME_OR_TRADEMARK As XElement = <SUPPLIER_NAME_OR_TRADEMARK/>
            SUPPLIER_NAME_OR_TRADEMARK.Value = Txt_TrademarkRef.Text
            MODEL_VERSION.Add(SUPPLIER_NAME_OR_TRADEMARK)

            '-Delegated Act
            Dim DELEGATED_ACT As XElement = <DELEGATED_ACT/>
            DELEGATED_ACT.Value = "EU_2019_2015"
            MODEL_VERSION.Add(DELEGATED_ACT)

            '-Product Group
            Dim PRODUCT_GROUP As XElement = <PRODUCT_GROUP/>
            PRODUCT_GROUP.Value = "LAMP"
            MODEL_VERSION.Add(PRODUCT_GROUP)

            Form2.LB_Log.Items.Add(MODEL_VERSION.Value + " - Success!")


        Next

        doc.Add(REGISTRATION)

        Console.WriteLine("Display the modified XML...")
        Console.WriteLine(doc)
        doc.Save(Console.Out)

        'OUTPUT()


    End Sub
    Public Sub OUTPUT()
        ''------RELEASE!--------
        'Dim dir As String = Directory.GetCurrentDirectory
        'Directory.GetAccessControl(dir + "\Data\")
        'doc.Save(dir + "\Data\registration-data.xml")
        '---------DEBUG!---------------------
        Directory.CreateDirectory("Data")
        doc.Save(".\Data\registration-data.xml")


        Dim start As String = ".\Data\"

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

        'If CB_OperationType.SelectedItem = "REGISTER_PRODUCT_MODEL" Then
        '    Dim SPECTRAL As New FolderBrowserDialog
        '    'Dim dmmy As String = Path.GetDirectoryName(ziel.FileName) & "\"
        '    'SPECTRAL.RootFolder = dmmy
        '    SPECTRAL.Description = "Please select folder with spectral data!"
        '    SPECTRAL.ShowDialog()

        '    Directory.CreateDirectory(start & "\SPECTRAL\")
        '    Dim fle As String
        '    Dim target As String = ""


        '    For Each fle In Directory.GetFiles(SPECTRAL.SelectedPath)
        '        target = start & "SPECTRAL\" & Path.GetFileName(fle)
        '        File.Copy(fle, target)
        '    Next

        'End If

        ZipFile.CreateFromDirectory(start, ziel.FileName)
            '---------------DEBUG!--------------------
            'Dim zipPath As String = ".\productModelRegistrationTable.zip"
            Directory.Delete(".\Data", True)

        MsgBox("Done!")

    End Sub


    Public Sub SELECT_INPUT()
        '----------------------------Validierung, ob Felder befüllt
        If Txt_Request.TextLength = 0 Or Txt_TrademarkRef.TextLength = 0 Then
            MsgBox("Please fill Values!")
            state = True
            Exit Sub
        End If
        '----------------------------Datei Auswählen und öffnen----------------------
        xlApp.Visible = False
        Dim quelle As New OpenFileDialog
        quelle.Title = "Please select the source file!"
        quelle.Filter = "Excel files (*.xlsx)|*.xlsx"
        quelle.ShowDialog()
        If quelle.FileName = "" Then
            MsgBox("Error!")
            state = True
            Exit Sub
        End If

        Select Case CB_OperationType.SelectedItem
            Case "UPDATE_PRODUCT_MODEL"
                PARSE_UPDATE(quelle)
            Case "PREREGISTER_PRODUCT_MODEL"
                PARSE_PREREGISTER(quelle)
            Case "REGISTER_PRODUCT_MODEL"
                PARSE_REGISTER(quelle)
            Case "DECLARE_END_DATE_OF_PLACEMENT_ON_MARKET"
                MsgBox("Not yet implemented")
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
    Sub PARSE_PREREGISTER(quelle)
        'PARSE -DATA

        Dim book = xlApp.Workbooks.Open(quelle.FileName)

        'Dim book = xlApp.Workbooks.Open("C:\Users\User79\Desktop\EPREL_Datenkonvertierung_Python_20210201\quelle.xlsx")



        Dim xltab1 = book.Worksheets("Tabelle1")
        'Dim items() As String
        Dim xlUP As Object = Excel.XlDirection.xlUp
        Dim lastentry As Object
        'Dim dummy As Integer

        dummy = book.Sheets(1).Range("A" & xltab1.Rows.Count).End(xlUP).Row
        lastentry = xltab1.Range("A1:A" & dummy).Value
        ReDim items(dummy - 1)
        For i = 1 To dummy - 1
            items(i - 1) = xltab1.Range("A" & i + 1).Value
        Next
        xlApp.Workbooks.Close()
        xlApp.Quit()
        'PARSE_END

    End Sub

    Sub PARSE_REGISTER(quelle)
        Dim book = xlApp.Workbooks.Open(quelle.FileName)
        Dim xltab1 = book.Worksheets("REGISTER_PRODUCT_MODEL")
        Dim xlUP As Object = Excel.XlDirection.xlUp
        Dim lastentry As Object

        dummy = book.Sheets(1).Range("A" & xltab1.Rows.Count).End(xlUP).Row
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
            _MODEL_IDENTIFIER(i - 1) = xltab1.Range("A" & i + 1).Value
            _CONSIDER_GENERATED_LABEL_AS_PROVIDED(i - 1) = xltab1.Range("B" & i + 1).Value
            '_ON_MARKET_START_DATE(i - 1) = xltab1.Range("C" & i + 1).Value
            dmy1 = xltab1.Range("C" & i + 1).Value
            _ON_MARKET_START_DATE(i - 1) = dmy1.ToString("yyyy") & "-" & dmy1.ToString("MM") & "-" & dmy1.ToString("dd") & dmy1.ToString("zzz")
            _VISIBLE_TO_UK_MSA(i - 1) = xltab1.Range("D" & i + 1).Value
            _LIGHTING_TECHNOLOGY(i - 1) = xltab1.Range("E" & i + 1).Value
            _DIRECTIONAL(i - 1) = xltab1.Range("F" & i + 1).Value
            _CAP_TYPE(i - 1) = xltab1.Range("G" & i + 1).Value
            _MAINS(i - 1) = xltab1.Range("H" & i + 1).Value
            _CONNECTED_LIGHT_SOURCE(i - 1) = xltab1.Range("I" & i + 1).Value
            _COLOUR_TUNEABLE_LIGHT_SOURCE(i - 1) = xltab1.Range("J" & i + 1).Value
            _HIGH_LUMINANCE_LIGHT_SOURCE(i - 1) = xltab1.Range("K" & i + 1).Value
            _ANTI_GLARE_SHIELD(i - 1) = xltab1.Range("L" & i + 1).Value
            _DIMMABLE(i - 1) = xltab1.Range("M" & i + 1).Value
            _ENERGY_CONS_ON_MODE(i - 1) = String.Format("{0000}", xltab1.Range("N" & i + 1).Value)
            _ENERGY_CLASS(i - 1) = xltab1.Range("O" & i + 1).Value
            _LUMINOUS_FLUX(i - 1) = String.Format("{00000}", xltab1.Range("Q" & i + 1).Value)
            _BEAM_ANGLE_CORRESPONDENCE(i - 1) = xltab1.Range("R" & i + 1).Value
            _CORRELATED_COLOUR_TEMP_TYPE(i - 1) = xltab1.Range("S" & i + 1).Value
            _CORRELATED_COLOUR_TEMP_SINGLE(i - 1) = String.Format("{00000}", xltab1.Range("T" & i + 1).Value)
            _CORRELATED_COLOUR_TEMP_MIN(i - 1) = String.Format("{00000}", xltab1.Range("U" & i + 1).Value)
            _CORRELATED_COLOUR_TEMP_MAX(i - 1) = String.Format("{00000}", xltab1.Range("V" & i + 1).Value)
            _CORRELATED_COLOUR_TEMP_1(i - 1) = String.Format("{00000}", xltab1.Range("W" & i + 1).Value)
            _CORRELATED_COLOUR_TEMP_2(i - 1) = String.Format("{00000}", xltab1.Range("X" & i + 1).Value)
            _CORRELATED_COLOUR_TEMP_3(i - 1) = String.Format("{00000}", xltab1.Range("Y" & i + 1).Value)
            _CORRELATED_COLOUR_TEMP_4(i - 1) = String.Format("{00000}", xltab1.Range("Z" & i + 1).Value)
            dmy2 = xltab1.Range("AA" & i + 1).Value
            _POWER_ON_MODE(i - 1) = String.Format(CultureInfo.CreateSpecificCulture("en-EN"), "{0:###0.0}", dmy2)
            dmy2 = xltab1.Range("AB" & i + 1).Value
            _POWER_STANDBY(i - 1) = String.Format(CultureInfo.CreateSpecificCulture("en-EN"), "{0:0.00}", dmy2)
            dmy2 = xltab1.Range("AC" & i + 1).Value
            _POWER_STANDBY_NETWORKED(i - 1) = String.Format(CultureInfo.CreateSpecificCulture("en-EN"), "{0:0.00}", dmy2)
            _POWER_STANDBY_NETWORKED(i - 1) = String.Format("{0,00}", _POWER_STANDBY_NETWORKED(i - 1))
            _COLOUR_RENDERING_INDEX(i - 1) = xltab1.Range("AD" & i + 1).Value
            _MIN_COLOUR_RENDERING_INDEX(i - 1) = xltab1.Range("AE" & i + 1).Value
            _MAX_COLOUR_RENDERING_INDEX(i - 1) = xltab1.Range("AF" & i + 1).Value
            dmy2 = xltab1.Range("AG" & i + 1).Value
            _DIMENSION_HEIGHT(i - 1) = String.Format(provider, "{0:#####}", dmy2)
            dmy2 = xltab1.Range("AH" & i + 1).Value
            _DIMENSION_WIDTH(i - 1) = String.Format(provider, "{0:#####}", dmy2)
            dmy2 = xltab1.Range("AI" & i + 1).Value
            _DIMENSION_DEPTH(i - 1) = String.Format(provider, "{0:#####}", dmy2)
            _SPECTRAL_POWER_DISTRIBUTION_IMAGE(i - 1) = xltab1.Range("AJ" & i + 1).Value
            dmy2 = xltab1.Range("AK" & i + 1).Value
            _CHROMATICITY_COORD_X(i - 1) = String.Format(provider, "{0:0.000}", dmy2)
            dmy2 = xltab1.Range("AL" & i + 1).Value
            _CHROMATICITY_COORD_Y(i - 1) = String.Format(provider, "{0:0.000}", dmy2)
            _DLS_PEAK_LUMINOUS_INTENSITY(i - 1) = xltab1.Range("AM" & i + 1).Value
            _DLS_BEAM_ANGLE(i - 1) = xltab1.Range("AN" & i + 1).Value
            _DLS_MIN_BEAM_ANGLE(i - 1) = xltab1.Range("AO" & i + 1).Value
            _DLS_MAX_BEAM_ANGLE(i - 1) = xltab1.Range("AP" & i + 1).Value
            dmy2 = xltab1.Range("AQ" & i + 1).Value
            _LED_R9_COLOUR_RENDERING_INDEX(i - 1) = String.Format(provider, "{0:###}", dmy2)
            dmy2 = xltab1.Range("AR" & i + 1).Value
            _LED_SURVIVAL_FACTOR(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
            dmy2 = xltab1.Range("AS" & i + 1).Value
            _LED_LUMEN_MAINTENANCE_FACTOR(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
            dmy2 = xltab1.Range("AT" & i + 1).Value
            _LED_MLS_DISPLACEMENT_FACTOR(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
            dmy2 = xltab1.Range("AU" & i + 1).Value
            _LED_MLS_COLOUR_CONSISTENCY(i - 1) = String.Format(provider, "{0:#}", dmy2)
            _LED_MLS_CLAIM_LED_REPLACE_FLUORESCENT(i - 1) = xltab1.Range("AV" & i + 1).Value
            dmy2 = xltab1.Range("AW" & i + 1).Value
            _LED_MLS_FL_REPLACEMENT_CLAIM(i - 1) = String.Format(provider, "{0:##}", dmy2)
            dmy2 = xltab1.Range("AX" & i + 1).Value
            _LED_MLS_FLICKER_METRIC(i - 1) = String.Format(provider, "{0:0.0}", dmy2)
            dmy2 = xltab1.Range("AY" & i + 1).Value
            _LED_MLS_STROBOSCOPIC_EFFECT_METRIC(i - 1) = String.Format(provider, "{0:0.0}", dmy2)

        Next
        xlApp.Workbooks.Close()
        xlApp.Quit()



    End Sub

    Sub PARSE_UPDATE(filename)

    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("mailto:m.planeck@nimbus-group.com")
    End Sub

    Private Sub CB_OperationType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_OperationType.SelectedIndexChanged
        If CB_OperationType.SelectedItem = "UPDATE_PRODUCT_MODEL" Then
            CB_ReasonChange.Enabled = True
            CB_ReasonChange.SelectedIndex = 0
        Else
            CB_ReasonChange.Enabled = False
        End If
    End Sub
End Class
