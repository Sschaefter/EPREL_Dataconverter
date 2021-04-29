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
    Public _EPREL_MODEL_REGISTRATION_NUMBER() As Integer
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
        Public _TD_FILE_NAME As String
    End Structure
    Public _TD() As _TECHNICAL_DOCUMENTATION

    Public dummy, dummy2 As Integer
    'Public doc As XmlDocument = New XmlDocument()
    Public doc As XDocument = New XDocument()
    Public state As Boolean = False

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If CheckB_Log.Checked = True Then

            Form2.Visible = True

        End If

        Cursor.Current = Cursors.WaitCursor

        If CB_OperationType.SelectedItem = "REGISTER_PRODUCT_MODEL" Then
            Select Case MsgBox("Please make shure, that all attachments are named liked in the source table and are located in one folder!", vbOKCancel)
                Case MsgBoxResult.Cancel
                    Exit Sub
                Case MsgBoxResult.Ok
                    Exit Select
            End Select
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
            Select Case MsgBox("Please make shure, that all attachments are named liked in the source table and are located in one folder!", vbOKCancel)
                Case MsgBoxResult.Cancel
                    Exit Sub
                Case MsgBoxResult.Ok
                    Exit Select
            End Select
            Form2.LB_Log.Items.Add("PREREGISTER_PRODUCT_MODEL")
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

        Try

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
                If _CONSIDER_GENERATED_LABEL_AS_PROVIDED(i) <> "" Then
                    Dim ENERGY_LABEL As XElement = <ENERGY_LABEL xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:ns5="http://eprel.ener.ec.europa.eu/commonTypes/EnergyLabelTypes/v2" xsi:type="ns5:GeneratedEnergyLabel"/>
                    Dim CONSIDER_GENERATED_LABEL_AS_PROVIDED As XElement = <CONSIDER_GENERATED_LABEL_AS_PROVIDED/>

                    CONSIDER_GENERATED_LABEL_AS_PROVIDED.Value = _CONSIDER_GENERATED_LABEL_AS_PROVIDED(i)
                    ENERGY_LABEL.Add(CONSIDER_GENERATED_LABEL_AS_PROVIDED)
                    MODEL_VERSION.Add(ENERGY_LABEL)
                Else
                    Form2.LB_Log.Items.Add("Energy Label for Modelidentifier " & _MODEL_IDENTIFIER(i) & " is missing!")
                    Throw New ArgumentException("Exception Occured")
                End If


                '---Market Start Date YYYY-MM-DD
                Dim ON_MARKET_START_DATE As XElement = <ON_MARKET_START_DATE/>
                ON_MARKET_START_DATE.Value = _ON_MARKET_START_DATE(i)
                MODEL_VERSION.Add(ON_MARKET_START_DATE)

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
                Finally

                End Try



                '-Kontakt

                Select Case Form_Contact.CB_ContactDetails.Checked
                    Case False
                        Dim CONTACT_DETAILS As XElement = <CONTACT_DETAILS xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="ns2:ContactByReference"/>
                        Dim CONTACT_REFERENCE As XElement = <CONTACT_REFERENCE/>
                        CONTACT_REFERENCE.Value = Txt_ContactRef.Text
                        CONTACT_DETAILS.Add(CONTACT_REFERENCE)
                        MODEL_VERSION.Add(CONTACT_DETAILS)
                    Case True
                        Dim CONTACT_DETAILS As XElement = <CONTACT_DETAILS xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="ns2:ModelSpecificContactDetails"/>
                        Dim CONTACT_NAME As XElement = <CONTACT_NAME/>
                        CONTACT_NAME.Value = Form_Contact.TB_ContactName.Text
                        CONTACT_DETAILS.Add(CONTACT_NAME)

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

                        With Form_Contact
                            If .TB_StreetName.Text <> "" Or .TB_Number.Text <> "" Or .TB_City.Text <> "" Or .TB_Municipality.Text <> "" Or .TB_Province.Text <> "" Or .TB_Postcode.Text <> "" Or .CBox_Country.SelectedItem <> "" Then
                                Dim ADDRESS As XElement = <ADDRESS xsi.type="ns5:DetailedAddress"/>

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
                        MODEL_VERSION.Add(CONTACT_DETAILS)
                End Select

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
                SPECTRAL_POWER_DISTRIBUTION_IMAGE.Value = "/attachments/" & _SPECTRAL_POWER_DISTRIBUTION_IMAGE(i)
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

        Catch ex As Exception
            MsgBox("Error while processing! Please check your files and try again!")
            state = True
        End Try


    End Sub
    Public Sub UPDATE_PRODUCT()
        Try

            Dim decl As XDeclaration = New XDeclaration(encoding:="UTF-8", standalone:="yes", version:="1.0")
            doc.Declaration = decl

            Dim REGISTRATION As XElement = <ns3:ProductModelRegistrationRequest xmlns:ns2="http://eprel.ener.ec.europa.eu/productModel/productCore/v2" REQUEST_ID="nothing"/>

            Dim REQUEST_ID As XAttribute = REGISTRATION.Attribute("REQUEST_ID")
            REQUEST_ID.Value = Txt_Request.Text

            For i = 0 To dummy - 2
                '-product Operation
                Dim productOperation As XElement = <productOperation OPERATION_TYPE="nothing" OPERATION_ID="nothing" REASON_FOR_CHANGE="nothing"/>
                REGISTRATION.Add(productOperation)
                Dim OPERATION_TYPE As XAttribute = productOperation.Attribute("OPERATION_TYPE")
                OPERATION_TYPE.Value = CB_OperationType.SelectedItem
                Dim REASON_FOR_CHANGE As XAttribute = productOperation.Attribute("REASON_FOR_CHANGE")
                REASON_FOR_CHANGE.Value = CB_ReasonChange.SelectedItem
                Dim OPERATION_ID As XAttribute = productOperation.Attribute("OPERATION_ID")
                OPERATION_ID.Value = i

                '-Model Verion
                Dim MODEL_VERSION As XElement = <MODEL_VERSION/>
                productOperation.Add(MODEL_VERSION)

                '-EPREL REGISTRATION Number
                Dim EPREL_REGISTRATION_NUMBER As XElement = <EPREL_MODEL_REGISTRATION_NUMBER/>
                EPREL_REGISTRATION_NUMBER.Value = _EPREL_MODEL_REGISTRATION_NUMBER(i)
                MODEL_VERSION.Add(EPREL_REGISTRATION_NUMBER)

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
                Finally

                End Try



                '-Kontakt

                Select Case Form_Contact.CB_ContactDetails.Checked
                    Case False
                        Dim CONTACT_DETAILS As XElement = <CONTACT_DETAILS xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="ns2:ContactByReference"/>
                        Dim CONTACT_REFERENCE As XElement = <CONTACT_REFERENCE/>
                        CONTACT_REFERENCE.Value = Txt_ContactRef.Text
                        CONTACT_DETAILS.Add(CONTACT_REFERENCE)
                        MODEL_VERSION.Add(CONTACT_DETAILS)
                    Case True
                        Dim CONTACT_DETAILS As XElement = <CONTACT_DETAILS xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="ns2:ModelSpecificContactDetails"/>
                        Dim CONTACT_NAME As XElement = <CONTACT_NAME/>
                        CONTACT_NAME.Value = Form_Contact.TB_ContactName.Text
                        CONTACT_DETAILS.Add(CONTACT_NAME)

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

                        With Form_Contact
                            If .TB_StreetName.Text <> "" Or .TB_Number.Text <> "" Or .TB_City.Text <> "" Or .TB_Municipality.Text <> "" Or .TB_Province.Text <> "" Or .TB_Postcode.Text <> "" Or .CBox_Country.SelectedItem <> "" Then
                                Dim ADDRESS As XElement = <ADDRESS xsi.type="ns5:DetailedAddress"/>

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
                        MODEL_VERSION.Add(CONTACT_DETAILS)
                End Select

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
                SPECTRAL_POWER_DISTRIBUTION_IMAGE.Value = "/attachments/" & _SPECTRAL_POWER_DISTRIBUTION_IMAGE(i)
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

        Catch ex As Exception
            MsgBox("Error while processing! Please check your files and try again!")
            state = True
        End Try

    End Sub
    Private Sub PREREGISTRATION()

        'SELECT_INPUT()
        Try
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
            If CheckB_Log.Checked = True Then
                Dim txt As TextWriter = Nothing
                doc.Save(txt)


            End If
            doc.Save(Console.Out)

        Catch ex As Exception
            MsgBox("Error while processing! Please check your files and try again!")
            state = True
        End Try


    End Sub
    Public Sub OUTPUT()

#If DEBUG Then
        '---------DEBUG!---------------------
        Directory.CreateDirectory("Data")
        doc.Save(".\Data\registration-data.xml")
#Else
        ''------RELEASE!--------
        Dim dir As String = Directory.GetCurrentDirectory
        Directory.GetAccessControl(dir + "\Data\")
        doc.Save(dir + "\Data\registration-data.xml")
#End If


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
        Try
            If CB_OperationType.SelectedItem = "REGISTER_PRODUCT_MODEL" Or CB_OperationType.SelectedItem = "UPDATE_PRODUCT_MODEL" Then
Select_File:
                Dim SPECTRAL As New FolderBrowserDialog
                'Dim dmmy As String = Path.GetDirectoryName(ziel.FileName) & "\"
                'SPECTRAL.RootFolder = dmmy
                SPECTRAL.Description = "Please select folder with attachment data!"

                SPECTRAL.ShowDialog()

                Directory.CreateDirectory(start & "\attachments\")
                Dim fle As String
                Dim target As String = ""


                For Each fle In Directory.GetFiles(SPECTRAL.SelectedPath)
                    target = start & "attachments\" & Path.GetFileName(fle)
                    File.Copy(fle, target)
                Next

            End If
        Catch

            Select Case MsgBox("Are you shure you do not want to upload any files?", MsgBoxStyle.YesNo)
                Case MsgBoxResult.Yes
                    Exit Try
                Case MsgBoxResult.No
                    GoTo Select_File
                Case Else
                    GoTo Select_File
            End Select
        End Try

        ZipFile.CreateFromDirectory(start, ziel.FileName)

#If DEBUG Then
        '---------------DEBUG!--------------------
        Directory.Delete(".\Data", True)
#End If
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



        Dim xltab1 = book.Worksheets("PREREGISTRATION")
        'Dim items() As String
        Dim xlUP As Object = Excel.XlDirection.xlUp
        Dim lastentry As Object
        'Dim dummy As Integer

        dummy = xltab1.Range("A" & xltab1.Rows.Count).End(xlUP).Row

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
        Dim xltab2 = book.Worksheets("attachments")
        Dim xlUP As Object = Excel.XlDirection.xlUp
        Dim lastentry As Object

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
            dmy1 = xltab1.Range("C" & i + 1).Value
            '-Format date to yyyy-mm-dd+hh:mm
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
            dmy2 = Math.Ceiling(xltab1.Range("T" & i + 1).Value / 100)
            _CORRELATED_COLOUR_TEMP_SINGLE(i - 1) = String.Format("{0000}", dmy2 * 100)
            dmy2 = Math.Ceiling(xltab1.Range("U" & i + 1).Value / 100)
            _CORRELATED_COLOUR_TEMP_MIN(i - 1) = String.Format("{00000}", dmy2 * 100)
            dmy2 = Math.Ceiling(xltab1.Range("V" & i + 1).Value / 100)
            _CORRELATED_COLOUR_TEMP_MAX(i - 1) = String.Format("{00000}", dmy2 * 100)
            dmy2 = Math.Ceiling(xltab1.Range("W" & i + 1).Value / 100)
            _CORRELATED_COLOUR_TEMP_1(i - 1) = String.Format("{00000}", dmy2 * 100)
            dmy2 = Math.Ceiling(xltab1.Range("X" & i + 1).Value / 100)
            _CORRELATED_COLOUR_TEMP_2(i - 1) = String.Format("{00000}", dmy2 * 100)
            dmy2 = Math.Ceiling(xltab1.Range("Y" & i + 1).Value / 100)
            _CORRELATED_COLOUR_TEMP_3(i - 1) = String.Format("{00000}", dmy2 * 100)
            dmy2 = Math.Ceiling(xltab1.Range("Z" & i + 1).Value / 100)
            _CORRELATED_COLOUR_TEMP_4(i - 1) = String.Format("{00000}", dmy2 * 100)
            dmy2 = xltab1.Range("AA" & i + 1).Value
            _POWER_ON_MODE(i - 1) = String.Format(provider, "{0:###0.0}", dmy2)
            dmy2 = xltab1.Range("AB" & i + 1).Value
            _POWER_STANDBY(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
            dmy2 = xltab1.Range("AC" & i + 1).Value
            _POWER_STANDBY_NETWORKED(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
            dmy2 = xltab1.Range("AD" & i + 1).Value
            _COLOUR_RENDERING_INDEX(i - 1) = String.Format(provider, "{0:###}", dmy2)
            dmy2 = xltab1.Range("AE" & i + 1).Value
            _MIN_COLOUR_RENDERING_INDEX(i - 1) = String.Format(provider, "{0:###}", dmy2)
            dmy2 = xltab1.Range("AF" & i + 1).Value
            _MAX_COLOUR_RENDERING_INDEX(i - 1) = String.Format(provider, "{0:###}", dmy2)
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
            dmy2 = xltab1.Range("AM" & i + 1).Value
            _DLS_PEAK_LUMINOUS_INTENSITY(i - 1) = String.Format(provider, "{0:######}", dmy2)
            dmy2 = xltab1.Range("AN" & i + 1).Value
            _DLS_BEAM_ANGLE(i - 1) = String.Format(provider, "{0:###}", dmy2)
            dmy2 = xltab1.Range("AO" & i + 1).Value
            _DLS_MIN_BEAM_ANGLE(i - 1) = String.Format(provider, "{0:###}", dmy2)
            dmy2 = xltab1.Range("AP" & i + 1).Value
            _DLS_MAX_BEAM_ANGLE(i - 1) = String.Format(provider, "{0:###}", dmy2)
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

        dummy2 = book.Worksheets("attachments").Range("A" & xltab2.Rows.Count).End(xlUP).Row
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
            _TD(i)._TD_SPECIFIC_PRECAUTIONS = xltab2.Range("I" & i + 2).Value
            _TD(i)._TD_FILE_NAME = xltab2.Range("J" & i + 2).Value
        Next


        xlApp.ActiveWorkbook.Close(False)
        xlApp.Quit()


    End Sub

    Sub PARSE_UPDATE(quelle)
        Dim book = xlApp.Workbooks.Open(quelle.FileName)
        Dim xltab1 = book.Worksheets("UPDATE_PRODUCT_MODEL")
        Dim xltab2 = book.Worksheets("attachments")
        Dim xlUP As Object = Excel.XlDirection.xlUp
        Dim lastentry As Object

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
            _HIGH_LUMINANCE_LIGHT_SOURCE(i - 1) = xltab1.Range("L" & i + 1).Value
            _ANTI_GLARE_SHIELD(i - 1) = xltab1.Range("M" & i + 1).Value
            _DIMMABLE(i - 1) = xltab1.Range("N" & i + 1).Value
            _ENERGY_CONS_ON_MODE(i - 1) = String.Format("{0000}", xltab1.Range("O" & i + 1).Value)
            _ENERGY_CLASS(i - 1) = xltab1.Range("P" & i + 1).Value
            _LUMINOUS_FLUX(i - 1) = String.Format("{00000}", xltab1.Range("R" & i + 1).Value)
            _BEAM_ANGLE_CORRESPONDENCE(i - 1) = xltab1.Range("S" & i + 1).Value
            _CORRELATED_COLOUR_TEMP_TYPE(i - 1) = xltab1.Range("T" & i + 1).Value
            dmy2 = Math.Ceiling(xltab1.Range("U" & i + 1).Value / 100)
            _CORRELATED_COLOUR_TEMP_SINGLE(i - 1) = String.Format("{0000}", dmy2 * 100)
            dmy2 = Math.Ceiling(xltab1.Range("V" & i + 1).Value / 100)
            _CORRELATED_COLOUR_TEMP_MIN(i - 1) = String.Format("{00000}", dmy2 * 100)
            dmy2 = Math.Ceiling(xltab1.Range("W" & i + 1).Value / 100)
            _CORRELATED_COLOUR_TEMP_MAX(i - 1) = String.Format("{00000}", dmy2 * 100)
            dmy2 = Math.Ceiling(xltab1.Range("X" & i + 1).Value / 100)
            _CORRELATED_COLOUR_TEMP_1(i - 1) = String.Format("{00000}", dmy2 * 100)
            dmy2 = Math.Ceiling(xltab1.Range("Y" & i + 1).Value / 100)
            _CORRELATED_COLOUR_TEMP_2(i - 1) = String.Format("{00000}", dmy2 * 100)
            dmy2 = Math.Ceiling(xltab1.Range("Z" & i + 1).Value / 100)
            _CORRELATED_COLOUR_TEMP_3(i - 1) = String.Format("{00000}", dmy2 * 100)
            dmy2 = Math.Ceiling(xltab1.Range("AA" & i + 1).Value / 100)
            _CORRELATED_COLOUR_TEMP_4(i - 1) = String.Format("{00000}", dmy2 * 100)
            dmy2 = xltab1.Range("AB" & i + 1).Value
            _POWER_ON_MODE(i - 1) = String.Format(provider, "{0:###0.0}", dmy2)
            dmy2 = xltab1.Range("AC" & i + 1).Value
            _POWER_STANDBY(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
            dmy2 = xltab1.Range("AD" & i + 1).Value
            _POWER_STANDBY_NETWORKED(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
            dmy2 = xltab1.Range("AE" & i + 1).Value
            _COLOUR_RENDERING_INDEX(i - 1) = String.Format(provider, "{0:###}", dmy2)
            dmy2 = xltab1.Range("AF" & i + 1).Value
            _MIN_COLOUR_RENDERING_INDEX(i - 1) = String.Format(provider, "{0:###}", dmy2)
            dmy2 = xltab1.Range("AG" & i + 1).Value
            _MAX_COLOUR_RENDERING_INDEX(i - 1) = String.Format(provider, "{0:###}", dmy2)
            dmy2 = xltab1.Range("AH" & i + 1).Value
            _DIMENSION_HEIGHT(i - 1) = String.Format(provider, "{0:#####}", dmy2)
            dmy2 = xltab1.Range("AI" & i + 1).Value
            _DIMENSION_WIDTH(i - 1) = String.Format(provider, "{0:#####}", dmy2)
            dmy2 = xltab1.Range("AJ" & i + 1).Value
            _DIMENSION_DEPTH(i - 1) = String.Format(provider, "{0:#####}", dmy2)
            _SPECTRAL_POWER_DISTRIBUTION_IMAGE(i - 1) = xltab1.Range("AK" & i + 1).Value
            dmy2 = xltab1.Range("AL" & i + 1).Value
            _CHROMATICITY_COORD_X(i - 1) = String.Format(provider, "{0:0.000}", dmy2)
            dmy2 = xltab1.Range("AM" & i + 1).Value
            _CHROMATICITY_COORD_Y(i - 1) = String.Format(provider, "{0:0.000}", dmy2)
            dmy2 = xltab1.Range("AN" & i + 1).Value
            _DLS_PEAK_LUMINOUS_INTENSITY(i - 1) = String.Format(provider, "{0:######}", dmy2)
            dmy2 = xltab1.Range("AO" & i + 1).Value
            _DLS_BEAM_ANGLE(i - 1) = String.Format(provider, "{0:###}", dmy2)
            dmy2 = xltab1.Range("AP" & i + 1).Value
            _DLS_MIN_BEAM_ANGLE(i - 1) = String.Format(provider, "{0:###}", dmy2)
            dmy2 = xltab1.Range("AQ" & i + 1).Value
            _DLS_MAX_BEAM_ANGLE(i - 1) = String.Format(provider, "{0:###}", dmy2)
            dmy2 = xltab1.Range("AR" & i + 1).Value
            _LED_R9_COLOUR_RENDERING_INDEX(i - 1) = String.Format(provider, "{0:###}", dmy2)
            dmy2 = xltab1.Range("AS" & i + 1).Value
            _LED_SURVIVAL_FACTOR(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
            dmy2 = xltab1.Range("AT" & i + 1).Value
            _LED_LUMEN_MAINTENANCE_FACTOR(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
            dmy2 = xltab1.Range("AU" & i + 1).Value
            _LED_MLS_DISPLACEMENT_FACTOR(i - 1) = String.Format(provider, "{0:0.00}", dmy2)
            dmy2 = xltab1.Range("AV" & i + 1).Value
            _LED_MLS_COLOUR_CONSISTENCY(i - 1) = String.Format(provider, "{0:#}", dmy2)
            _LED_MLS_CLAIM_LED_REPLACE_FLUORESCENT(i - 1) = xltab1.Range("AW" & i + 1).Value
            dmy2 = xltab1.Range("AX" & i + 1).Value
            _LED_MLS_FL_REPLACEMENT_CLAIM(i - 1) = String.Format(provider, "{0:##}", dmy2)
            dmy2 = xltab1.Range("AY" & i + 1).Value
            _LED_MLS_FLICKER_METRIC(i - 1) = String.Format(provider, "{0:0.0}", dmy2)
            dmy2 = xltab1.Range("AZ" & i + 1).Value
            _LED_MLS_STROBOSCOPIC_EFFECT_METRIC(i - 1) = String.Format(provider, "{0:0.0}", dmy2)

        Next

        dummy2 = book.Worksheets("attachments").Range("A" & xltab2.Rows.Count).End(xlUP).Row
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
            _TD(i)._TD_SPECIFIC_PRECAUTIONS = xltab2.Range("I" & i + 2).Value
            _TD(i)._TD_FILE_NAME = xltab2.Range("J" & i + 2).Value
        Next


        xlApp.ActiveWorkbook.Close(False)
        xlApp.Quit()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("mailto:m.planeck@nimbus-group.com")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form_Contact.Show()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "EPREL Dataconverter " & My.Application.Info.Version.ToString
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
