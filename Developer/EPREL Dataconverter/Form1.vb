Imports Excel = Microsoft.Office.Interop.Excel
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
            'SELECT_INPUT()
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
            'PREREGISTRATION()
            If state = True Then
                Exit Sub
            End If
            OUTPUT()
        ElseIf CB_OperationType.SelectedItem = "UPDATE_PRODUCT_MODEL" Then
            Form2.LB_Log.Items.Add("PREREGISTER_PRODUCT_MODEL")
            SELECT_INPUT()
            If state = True Then
                Exit Sub
            End If
            'UPDATE_PRODUCT()
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


        'SELECT_INPUT()

        'Dim decl As XmlDeclaration
        'decl = doc.CreateXmlDeclaration("1.0", Nothing, Nothing)
        'decl.Encoding = "UTF-8"
        'decl.Standalone = "yes"

        'Dim ns3 As XNamespace = ns3:"http://eprel.ener.ec.europa.eu/services/productModelService/modelRegistrationService/v2"

        Dim decl As XDeclaration = New XDeclaration(encoding:="UTF-8", standalone:="yes", version:="1.0")
        doc.Declaration = decl



        Dim REGISTRATION As XElement = <ns3:ProductModelRegistrationRequest xmlns:ns2="http://eprel.ener.ec.europa.eu/productModel/productCore/v2" REQUEST_ID="nothing"/>

        Dim REQUEST_ID As XAttribute = REGISTRATION.Attribute("REQUEST_ID")
        REQUEST_ID.Value = Txt_Request.Text
        '-product Operation
        Dim productOperation As XElement = <productOperation OPERATION_TYPE="nothing" OPERATION_ID="nothing"/>
        REGISTRATION.Add(productOperation)
        Dim OPERATION_TYPE As XAttribute = productOperation.Attribute("OPERATION_TYPE")
        OPERATION_TYPE.Value = CB_OperationType.SelectedItem
        Dim OPERATION_ID As XAttribute = productOperation.Attribute("OPERATION_ID")
        OPERATION_ID.Value = "1"

        '-Model Verion
        Dim MODEL_VERSION As XElement = <MODEL_VERSION/>
        productOperation.Add(MODEL_VERSION)

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
        COLOUR_TUNEABLE_LIGHT_SOURCE.Value="true"
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

        Dim CLAIM_EQUIVALENT_POWER As XElement = <CLAIM_EQUIVALENT_POWER/>
        CLAIM_EQUIVALENT_POWER.Value = "true"
        PRODUCT_GROUP_DETAIL.Add(CLAIM_EQUIVALENT_POWER)

        Dim EQUIVALENT_POWER As XElement = <EQUIVALENT_POWER/>
        EQUIVALENT_POWER.Value = "1"
        PRODUCT_GROUP_DETAIL.Add(EQUIVALENT_POWER)

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






        'Next
        doc.Add(REGISTRATION)
        Console.WriteLine("Display the modified XML...")
        Console.WriteLine(REGISTRATION)
        Console.WriteLine(doc)
        doc.Save(Console.Out)

        'OUTPUT()

    End Sub
    'Public Sub UPDATE_PRODUCT()


    '    'SELECT_INPUT()

    '    Dim decl As XmlDeclaration
    '    decl = doc.CreateXmlDeclaration("1.0", Nothing, Nothing)
    '    decl.Encoding = "UTF-8"
    '    decl.Standalone = "yes"


    '    Dim ProductModelRegistrationRequest As XmlNode = doc.CreateNode(XmlNodeType.Element, "ns3", "ProductModelRegistrationRequest", "http://eprel.ener.ec.europa.eu/services/productModelService/modelRegistrationService/v2")
    '    doc.AppendChild(ProductModelRegistrationRequest)

    '    doc.InsertBefore(decl, ProductModelRegistrationRequest)

    '    Dim REQUEST_ID As XmlAttribute = doc.CreateAttribute("REQUEST_ID")
    '    REQUEST_ID.Value = Txt_Request.Text
    '    doc.DocumentElement.SetAttributeNode(REQUEST_ID)

    '    For i = 0 To dummy - 2
    '        Dim productOperation As XmlNode = doc.CreateNode("element", "productOperation", "")
    '        ProductModelRegistrationRequest.AppendChild(productOperation)

    '        Dim OPERATION_TYPE As XmlNode = doc.CreateNode(XmlNodeType.Attribute, "OPERATION_TYPE", "")
    '        'OPERATION_TYPE.Value = Txt_OperationType.Text
    '        OPERATION_TYPE.Value = CB_OperationType.SelectedItem
    '        productOperation.Attributes.SetNamedItem(OPERATION_TYPE)

    '        Dim OPERATION_ID As XmlNode = doc.CreateNode(XmlNodeType.Attribute, "OPERATION_ID", "")
    '        OPERATION_ID.Value = i
    '        productOperation.Attributes.SetNamedItem(OPERATION_ID)

    '        Dim MODEL_VERSION As XmlNode = doc.CreateNode("element", "MODEL_VERSION", "")
    '        productOperation.AppendChild(MODEL_VERSION)

    '        Dim MODEL_IDENTIFIER As XmlElement = doc.CreateElement("MODEL_IDENTIFIER")
    '        MODEL_IDENTIFIER.InnerText = items(i)
    '        MODEL_VERSION.AppendChild(MODEL_IDENTIFIER)

    '        Dim TRADEMARK_REFERENCE As XmlElement = doc.CreateElement("TRADEMARK_REFERENCE")
    '        TRADEMARK_REFERENCE.InnerText = Txt_TrademarkRef.Text
    '        MODEL_VERSION.AppendChild(TRADEMARK_REFERENCE)

    '        Dim DELEGATED_ACT As XmlElement = doc.CreateElement("DELEGATED_ACT")
    '        DELEGATED_ACT.InnerText = "EU_2019_2015"
    '        MODEL_VERSION.AppendChild(DELEGATED_ACT)

    '        '---Energy label provided by supplier? TRUE/FALSE---
    '        Dim ENERGY_LABEL As XmlNode = doc.CreateNode(XmlNodeType.Element, "ns5", "ENERGY_LABEL", "http://eprel.ener.ec.europa.eu/commonTypes/EnergyLabelTypes/v2")
    '        MODEL_VERSION.AppendChild(ENERGY_LABEL)

    '        Dim CONSIDER_GENERATED_LABEL As XmlElement = doc.CreateElement("CONSIDER_GENERATED_LABEL")
    '        CONSIDER_GENERATED_LABEL.InnerText = "TRUE"
    '        ENERGY_LABEL.AppendChild(CONSIDER_GENERATED_LABEL)

    '        '---Market Start Date YYYY-MM-DD
    '        Dim ON_MARKET_START_DATE As XmlElement = doc.CreateElement("ON_MARKET_START_DATE")
    '        ON_MARKET_START_DATE.InnerText = "2021-05-01"
    '        MODEL_VERSION.AppendChild(ON_MARKET_START_DATE)

    '        '---Technical Documentation
    '        Dim TECHNICAL_DOCUMENTATION As XmlNode = doc.CreateNode(XmlNodeType.Element, "ns2", "TECHNICAL_DOCUMENTATION", "http://eprel.ener.ec.europa.eu/productModel/productCore/v2")
    '        MODEL_VERSION.AppendChild(TECHNICAL_DOCUMENTATION)

    '        Dim DOCUMENT As XmlNode = doc.CreateNode(XmlNodeType.Element, "DOCUMENT", "")
    '        TECHNICAL_DOCUMENTATION.AppendChild(DOCUMENT)

    '        Dim LANGUAGE As XmlElement = doc.CreateElement("LANGUAGE")
    '        LANGUAGE.InnerText = "DE"
    '        DOCUMENT.AppendChild(LANGUAGE)

    '        Dim TECHNICAL_PART As XmlElement = doc.CreateElement("TECHNICAL_PART")
    '        TECHNICAL_PART.InnerText = "TESTING_CONDICTIONS"
    '        DOCUMENT.AppendChild(TECHNICAL_PART)

    '        Dim FILE_PATH As XmlElement = doc.CreateElement("FILE_PATH")
    '        Dim j As String = (i + 1).ToString
    '        FILE_PATH.InnerText = "/attachment/Test" + j + ".pdf"
    '        DOCUMENT.AppendChild(FILE_PATH)

    '        '---Product Group Detail---
    '        Dim PRODUCT_GROUP_DETAIL As XmlNode = doc.CreateNode(XmlNodeType.Element, "ns5", "PRODUCT_GROUP_DETAIL", "http://eprel.ener.ec.europa.eu/productModel/productGroups/Lamp2019/v1")
    '        MODEL_VERSION.AppendChild(PRODUCT_GROUP_DETAIL)
    '        Dim ENERGY_CLASS As XmlElement = doc.CreateElement("ENERGY_CLASS")
    '        ENERGY_CLASS.InnerText = "A"
    '        PRODUCT_GROUP_DETAIL.AppendChild(ENERGY_CLASS)
    '        Dim WEIGHTED_ENERGY_CONS As XmlElement = doc.CreateElement("WEIGHTED_ENERY_CONS")
    '        WEIGHTED_ENERGY_CONS.InnerText = "1000"
    '        PRODUCT_GROUP_DETAIL.AppendChild(WEIGHTED_ENERGY_CONS)



    '        '--------Kontakt
    '        Dim CONTACT_DETAILS As XmlNode = doc.CreateNode(XmlNodeType.Element, "ns2", "CONTACT_DETAILS", "http://eprel.ener.ec.europa.eu/productModel/productGroupInterfaces/v1")
    '        MODEL_VERSION.AppendChild(CONTACT_DETAILS)
    '        Dim CONTACT_REFERENCE As XmlElement = doc.CreateElement("CONTACT_REFERENCE")
    '        CONTACT_REFERENCE.InnerText = "Kontakt #1"
    '        CONTACT_DETAILS.AppendChild(CONTACT_REFERENCE)



    '        'If CB_OperationType.SelectedItem = "UPDATE_PRODUCT_MODEL" Then
    '        '    Dim ENERGY_LABEL As XmlNode = doc.CreateNode(XmlNodeType.Element, "ns6", "ENERGY_LABEL", "http://eprel.ener.ec.europa.eu/commonTypes/EnergyLabelTypes/v2")
    '        '    MODEL_VERSION.AppendChild(ENERGY_LABEL)

    '        '    Dim CONSIDER_ENERGY_LABEL_AS_PROVIDED As XmlElement = doc.CreateElement("CONSIDERED_ENERGY_LABEL_AS_PROVIDED")
    '        '    CONSIDER_ENERGY_LABEL_AS_PROVIDED.InnerText = "TRUE"
    '        '    ENERGY_LABEL.AppendChild(CONSIDER_ENERGY_LABEL_AS_PROVIDED)

    '        'End If

    '        productOperation.AppendChild(MODEL_VERSION)



    '    Next

    '    Console.WriteLine("Display the modified XML...")
    '    Console.WriteLine(doc)
    '    doc.Save(Console.Out)

    '    'OUTPUT()

    'End Sub


    'Private Sub PREREGISTRATION()

    '    'SELECT_INPUT()

    '    Dim decl As XmlDeclaration
    '    decl = doc.CreateXmlDeclaration("1.0", Nothing, Nothing)
    '    decl.Encoding = "UTF-8"
    '    decl.Standalone = "yes"


    '    Dim ProductModelRegistrationRequest As XmlNode = doc.CreateNode(XmlNodeType.Element, "ns3", "ProductModelRegistrationRequest", "http://eprel.ener.ec.europa.eu/services/productModelService/modelRegistrationService/v2")
    '    doc.AppendChild(ProductModelRegistrationRequest)

    '    doc.InsertBefore(decl, ProductModelRegistrationRequest)

    '    Dim REQUEST_ID As XmlAttribute = doc.CreateAttribute("REQUEST_ID")
    '    REQUEST_ID.Value = Txt_Request.Text
    '    doc.DocumentElement.SetAttributeNode(REQUEST_ID)

    '    For i = 0 To dummy - 2
    '        Dim productOperation As XmlNode = doc.CreateNode("element", "productOperation", "")
    '        ProductModelRegistrationRequest.AppendChild(productOperation)

    '        Dim OPERATION_TYPE As XmlNode = doc.CreateNode(XmlNodeType.Attribute, "OPERATION_TYPE", "")
    '        'OPERATION_TYPE.Value = Txt_OperationType.Text
    '        OPERATION_TYPE.Value = CB_OperationType.SelectedItem
    '        productOperation.Attributes.SetNamedItem(OPERATION_TYPE)
    '        Dim OPERATION_ID As XmlNode = doc.CreateNode(XmlNodeType.Attribute, "OPERATION_ID", "")
    '        OPERATION_ID.Value = i
    '        productOperation.Attributes.SetNamedItem(OPERATION_ID)

    '        Dim MODEL_VERSION As XmlNode = doc.CreateNode("element", "MODEL_VERSION", "")
    '        productOperation.AppendChild(MODEL_VERSION)
    '        Dim MODEL_IDENTIFIER As XmlElement = doc.CreateElement("MODEL_IDENTIFIER")
    '        MODEL_IDENTIFIER.InnerText = items(i)
    '        MODEL_VERSION.AppendChild(MODEL_IDENTIFIER)
    '        Dim TRADEMARK_REFERENCE As XmlElement = doc.CreateElement("TRADEMARK_REFERENCE")
    '        TRADEMARK_REFERENCE.InnerText = Txt_TrademarkRef.Text
    '        MODEL_VERSION.AppendChild(TRADEMARK_REFERENCE)
    '        Dim DELEGATED_ACT As XmlElement = doc.CreateElement("DELEGATED_ACT")
    '        DELEGATED_ACT.InnerText = "EU_2019_2015"
    '        MODEL_VERSION.AppendChild(DELEGATED_ACT)
    '        Dim PRODUCT_GROUP As XmlElement = doc.CreateElement("PRODUCT_GROUP")
    '        PRODUCT_GROUP.InnerText = "LAMP"
    '        MODEL_VERSION.AppendChild(PRODUCT_GROUP)

    '        productOperation.AppendChild(MODEL_VERSION)

    '        Form2.LB_Log.Items.Add(MODEL_VERSION.InnerText + " - Success!")


    '    Next

    '    Console.WriteLine("Display the modified XML...")
    '    Console.WriteLine(doc)
    '    doc.Save(Console.Out)

    '    'OUTPUT()


    'End Sub
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

    End Sub


    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("mailto:m.planeck@nimbus-group.com")
    End Sub
End Class
