VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetEnvelopeXML_HTTP 
      Caption         =   "GetEnvelopeXML - MSXML2.XMLHTTP"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton cmdGetEnvelopeXMLSTK 
      Caption         =   "GetEnvelopeXML - Soap ToolKit"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "após a execução a saida é printada da janela de debug. "
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   4035
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'POC realizada para a chamada de SOAP web services


'este metodo realiza a chamada via 'baixo nivel' utilizando o objeto MSXML2.XMLHTTP
Private Sub cmdGetEnvelopeXML_HTTP_Click()
Dim strUrl          As String
Dim strSoapAction   As String

Dim strXML          As String
Dim xmlDocIN        As Object
Dim xmlDocOUT       As Object
Dim strXMLRet       As String

Dim HTTPReq         As Object
Dim rsReturn        As Object

    strUrl = "http://localhost:53656/WebService1.asmx/GetEnvelopesXML"
    strSoapAction = "http://tempuri.org/GetEnvelopesXML"
    
    strXML = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
    "  <soap:Body>" & _
    "    <GetEnvelopesXMLResponse xmlns=""http://tempuri.org/"">" & _
    "      <GetEnvelopesXMLResult>" & _
    "        <Envelope>" & _
    "          <EnvelopeID>string</EnvelopeID>" & _
    "          <EnvelopeValor>string</EnvelopeValor>" & _
    "        </Envelope>" & _
    "        <Envelope>" & _
    "          <EnvelopeID>string</EnvelopeID>" & _
    "          <EnvelopeValor>string</EnvelopeValor>" & _
    "        </Envelope>" & _
    "      </GetEnvelopesXMLResult>" & _
    "    </GetEnvelopesXMLResponse>" & _
    "  </soap:Body>" & _
    "</soap:Envelope>"
    
    Set xmlDocIN = CreateObject("MSXML2.DOMDocument")
    Set HTTPReq = CreateObject("MSXML2.XMLHTTP")
    
    '#####
    'Set rsReturn = New ADODB.Recordset
    Set rsReturn = CreateObject("ADODB.Recordset")
    
    rsReturn.fields.Append "EnvelopeID", adVarWChar, 255
    rsReturn.fields.Append "EnvelopeValor", adVarWChar, 255
    
    rsReturn.CursorLocation = adUseClient
    rsReturn.CursorType = adOpenStatic
    rsReturn.LockType = adLockOptimistic
    rsReturn.open
    
    xmlDocIN.async = False
    xmlDocIN.loadXML strXML
    HTTPReq.open "POST", strUrl, False
    
    HTTPReq.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    HTTPReq.setRequestHeader "SOAPAction", strSoapAction
    
    HTTPReq.send xmlDocIN.xml
    
    If Not HTTPReq.readyState = 4 And (HTTPReq.Status >= 200 And HTTPReq.Status <= 300) Then
        MsgBox "Erro em CallListarDadosDeclaradoEnvelopeBBox - HTTP Ready State: " & _
            HTTPReq.readyState & " HTTP Status: " & HTTPReq.Status
    End If
    
    Set xmlDocOUT = CreateObject("MSXML2.DOMDocument") 'MSXML2.DOMDocument60
    Set xmlDocOUT = HTTPReq.responseXML
    strXMLRet = HTTPReq.responseText
    
    Dim i As Long
    For i = 0 To xmlDocOUT.selectNodes("//Envelope").length - 1
        Debug.Print xmlDocOUT.selectNodes("//Envelope").Item(i).selectSingleNode("EnvelopeID").Text
        Debug.Print xmlDocOUT.selectNodes("//Envelope").Item(i).selectSingleNode("EnvelopeValor").Text
        rsReturn.AddNew
        rsReturn("EnvelopeID") = xmlDocOUT.selectNodes("//Envelope").Item(i).selectSingleNode("EnvelopeID").Text
        rsReturn("EnvelopeValor") = xmlDocOUT.selectNodes("//Envelope").Item(i).selectSingleNode("EnvelopeValor").Text
    Next i

    'resulado final é um recordset ADO
    rsReturn.MoveFirst
    Do While Not rsReturn.EOF
        Debug.Print rsReturn("EnvelopeID") & "-" & rsReturn("EnvelopeValor")
        rsReturn.MoveNext
    Loop
    
    MsgBox "Execução finalizada com sucesso."
End Sub

'este metodo realiza a chamada via lib Microsoft SOAP Tool Kit utilizando o objeto MSSOAP
Private Sub cmdGetEnvelopeXMLSTK_Click()
    'Dim Serializer As SoapSerializer30
    Dim Serializer
    Set Serializer = CreateObject("MSSOAP.SoapSerializer30")
    
    'Dim Reader As SoapReader30
    Dim Reader
    Set Reader = CreateObject("MSSOAP.SoapReader30")
        
    Dim ResultElm As IXMLDOMElement
    Dim FaultElm As IXMLDOMElement
    'Dim Connector As SoapConnector30
    Dim Connector
    Set Connector = CreateObject("MSSOAP.HttpConnector30")
    'Set Connector = New HttpConnector30
    Connector.Property("EndPointURL") = "http://localhost:53656/WebService1.asmx"
    Connector.Connect
    
    ' binding/operation/soapoperation
    'Connector.Property("SoapAction") = SoapAction & Method
    Connector.Property("SoapAction") = "http://tempuri.org/GetEnvelopesXML"
    Connector.BeginMessage
    
    'Set Serializer = New SoapSerializer30
    Set Serializer = CreateObject("MSSOAP.SoapSerializer30")
    Serializer.Init Connector.InputStream
    
    Serializer.StartEnvelope
    Serializer.StartBody
    'Serializer.startElement Method, CALC_NS
    Serializer.startElement "GetEnvelopesXML", "http://tempuri.org/"
    
        'Caso o webservice tenha parametros, é necessário a definição de cada um deles como abaixo
        'Serializer.startElement "a"
        'Serializer.WriteString CStr(A)
        'Serializer.endElement
    Serializer.endElement
    Serializer.EndBody
    Serializer.EndEnvelope
    
    Connector.EndMessage
        
    'Set Reader = New SoapReader30
    Set Reader = CreateObject("MSSOAP.SoapReader30")
    
    Reader.Load Connector.OutputStream
    
    If Not Reader.Fault Is Nothing Then
        'ocorreu erro na chamada ao ws
        MsgBox Reader.FaultString.Text, vbExclamation
    End If
    
    'Debug.Print y.Item(1).selectSingleNode("//ResultsCount").nodeTypedValue
    Dim rsReturn
    Set rsReturn = CreateObject("ADODB.Recordset")
    rsReturn.fields.Append "EnvelopeID", adVarWChar, 255
    rsReturn.fields.Append "EnvelopeValor", adVarWChar, 255
'
    rsReturn.CursorLocation = adUseClient
    rsReturn.CursorType = adOpenStatic
    rsReturn.LockType = adLockOptimistic
    rsReturn.open
    
    Dim xmlDocOUT       As Object
    Set xmlDocOUT = CreateObject("MSXML2.DOMDocument") 'MSXML2.DOMDocument60
    'Set xmlDocOUT = HTTPReq.responseXML
    xmlDocOUT.loadXML (Reader.RpcResult.xml)
    
    Dim i As Long
    For i = 0 To xmlDocOUT.selectNodes("//Envelope").length - 1
        Debug.Print xmlDocOUT.selectNodes("//Envelope").Item(i).selectSingleNode("EnvelopeID").Text
        Debug.Print xmlDocOUT.selectNodes("//Envelope").Item(i).selectSingleNode("EnvelopeValor").Text
        rsReturn.AddNew
        rsReturn("EnvelopeID") = xmlDocOUT.selectNodes("//Envelope").Item(i).selectSingleNode("EnvelopeID").Text
        rsReturn("EnvelopeValor") = xmlDocOUT.selectNodes("//Envelope").Item(i).selectSingleNode("EnvelopeValor").Text
    Next i

    'resulado final é um recordset ADO
    rsReturn.MoveFirst
    Do While Not rsReturn.EOF
        Debug.Print rsReturn("EnvelopeID") & "-" & rsReturn("EnvelopeValor")
        rsReturn.MoveNext
    Loop
    MsgBox "Execução finalizada com sucesso."
End Sub
