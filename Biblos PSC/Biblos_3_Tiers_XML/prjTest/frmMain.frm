VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command23 
      Caption         =   "Command23"
      Height          =   495
      Left            =   2520
      TabIndex        =   22
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Command22"
      Height          =   495
      Left            =   1320
      TabIndex        =   21
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Command21"
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Command20"
      Height          =   375
      Left            =   3360
      TabIndex        =   19
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Command19"
      Height          =   375
      Left            =   2280
      TabIndex        =   18
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Command18"
      Height          =   375
      Left            =   1200
      TabIndex        =   17
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Command17"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Command16"
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Command15"
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Command14"
      Height          =   375
      Left            =   1200
      TabIndex        =   13
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Command13"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Command11"
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oRsCustomers
Dim oRsUbicaciones

Private Declare Function GetClassName Lib "User32" Alias _
         "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As _
         String, ByVal nMaxCount As Long) As Long


Private Sub Command1_Click()
Dim caca As String

caca = "!#$%&/(()"

MsgBox CheckString(caca)
End Sub



Public Function RecordsetFromXMLDocument(XMLDOMDocument)
Dim oRecordset
    
    Set oRecordset = New ADODB.Recordset
    oRecordset.Open XMLDOMDocument 'pass the DOM Document instance as the Source argument
    Set RecordsetFromXMLDocument = oRecordset  'return the recordset
    Set oRecordset = Nothing

End Function

    '''''''''''''''''''''''''''''''''''''''

Public Sub LoadUbicaciones(XMLUbicaciones)
    Dim oDOM
    Dim childNode
    Dim i
    Set oDOM = New MSXML2.DOMDocument
    
    If oDOM.loadXML(XMLUbicaciones) Then
        Set oRsUbicaciones = RecordsetFromXMLDocument(oDOM)
        While Not oRsUbicaciones.EOF
            For i = 0 To oRsUbicaciones.fields.Count - 1
                Debug.Print oRsUbicaciones(i) & "<BR> "
            Next
            oRsUbicaciones.MoveNext
        Wend
        Set oRsUbicaciones = Nothing
    Else
        'Wrong
    End If
    Set oDOM = Nothing
End Sub


Private Sub Command13_Click()
    Dim x As New cPrestamoTipo
Dim xml As String
Dim le, de, eu
Dim iID
Dim oDOM As New MSXML2.DOMDocument

x.Search 1, xml, , "descripcion", "ASC", le, de, eu

'Debug.Print le & de & eu
Debug.Print xml


End Sub

Private Sub Command14_Click()
Dim x As New cPrestamo
Dim xml As String
Dim le, de, eu
Dim iID
Dim oDOM As New MSXML2.DOMDocument
With x
    .Fecha_Desde = "20070321"
    .Fecha_Hasta = "20070326"
    .UsuarioID = 5
    .BibliotecariaID = 6
    .ItemID = 25
    .Tipo_PrestamoID = 2
End With
x.Add 1



End Sub

Private Sub Command15_Click()
Dim lErrNum, sErrDesc, sErrSource
Dim oitem As New citem
Dim oReserva As New cReserva
Dim s As String
Dim strMSg As String
Dim strXML As String
Dim strRes As String
Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument

oRs.fields.Append "titulo", adBSTR
oRs.fields.Append "autor", adBSTR
oRs.fields.Append "fecha_reserva", adBSTR

'oRs.Open

'oRs.AddNew

'oRs(0) = "El principito"
'oRs(1) = "Saint Exupery, Antoine"
'oRs(2) = "27/03/2007"

'oRs.Update

'oRs.save oDOM, adPersistXML

oitem.SearchForBorrow 5, strXML

If oitem.SearchForReserve(5, strXML, oDOM.xml, , , , lErrNum, sErrDesc, sErrSource) Then
                Set oDOM = Nothing
                Set oRs = Nothing
                If oitem.Read(1, strXML, lErrNum, sErrDesc, sErrSource) Then
                    oReserva.ItemID = oitem.id
                    If oReserva.Search(1, strXML, "Fecha_reserva = " & "20070327", , , lErrNum, sErrDesc, sErrSource) Then
                        If oReserva.Read(1, strXML, lErrNum, sErrDesc, sErrSource) Then
                            Set oDOM = New MSXML2.DOMDocument
                            If oDOM.loadXML(strXML) Then
                                Set oRs = RecordsetFromXMLDocument(oDOM)
                                If Not oRs.EOF Then strMSg = "El item ha sido reservado, pero existen (" & oRs.RecordCount & ") reservas anteriores a la suya."
                            End If
                            Set oDOM = Nothing
                            Set oRs = Nothing
                        Else
                            strMSg = "item reservado con éxito."
                        End If
                        If oReserva.Add(1, lErrNum, sErrDesc, sErrSource) Then
                                Set oReserva = Nothing
                            'response.redirect "items_search.asp?msg=" & strMSg
                        Else
                            'response.redirect "error.asp?title=" & strTitle & "&message=Error Nro: " & CStr(lErrNum) & " <BR> Descripción: " & sErrDesc & " <BR> Origen:  " & sErrSource
                        End If
                    Else
                        'response.redirect "error.asp?title=" & strTitle & "&message=Error Nro: " & CStr(lErrNum) & " <BR> Descripción: " & sErrDesc & " <BR> Origen:  " & sErrSource
                    End If
                Else
                    'response.redirect "error.asp?title=" & strTitle & "&message=Error Nro: " & cStr(lErrNum) & " <BR> Descripción: " & sErrDesc & " <BR> Origen:  " & sErrSource
                    strMSg = "No existen copias disponibles para el ."
                    'response.redirect "items_search.asp?msg=" & strMSg
                End If
            Else
                'response.redirect "error.asp?title=" & strTitle & "&message=Error Nro: " & CStr(lErrNum) & " <BR> Descripción: " & sErrDesc & " <BR> Origen:  " & sErrSource
            End If








End Sub

Private Sub Command16_Click()
Dim le, ed, es
Dim x As New cReserva
Dim s As String
Dim strRes As String
Dim oRs As Recordset
Dim oDOM As MSXML2.DOMDocument

'x.Search 1, s, "Fecha_reserva = " & Now, , , le, ed, es
'47
'15/03/2007
'6
'6
x.id = 6
x.Fecha_Reserva = "15/03/2007"
x.ItemID = 47
x.UsuarioID = 6

x.Update 1, le, ed, es

Debug.Print s
Debug.Print le
Debug.Print ed
Debug.Print es
End Sub

Private Sub Command17_Click()

Dim le, ed, es
Dim x As New cPrestamo
Dim s As String
Dim strRes As String
Dim oRs As Recordset
Dim oDOM As MSXML2.DOMDocument

x.Search 1, s, "id_usuario = " & 5, , , le, ed, es
Debug.Print s
Debug.Print le
Debug.Print ed
Debug.Print es

End Sub

Private Sub Command18_Click()
Dim x As New cReserva
Dim strXML As String
Dim le, de, eu
Dim iID

x.Search 1, strXML, "Fecha_reserva = 20070327 AND titulo = 'El principito' AND itemtipoID = 2", , , le, de, eu

x.Read 1, strXML
End Sub

Private Sub Command19_Click()
Dim x As New citem
Dim xml As String
Dim le, de, eu
Dim iID
Dim oDOM As New MSXML2.DOMDocument

x.id = 78
x.Delete 1, le, de, eu

End Sub

''''''''''''''''''''''''
Private Sub Command2_Click()
Dim x As New cUsuario
Dim xml As String
Dim le, de, eu
Dim iID
Dim oDOM As New MSXML2.DOMDocument

'x.updateUbicacion 3, "hola", "chau"

'x.GetUbicacion 3, xml
'x.id = 18
'x.Titulo = "chausss"
'x.Update
'x.Add
'x.search xml, , "id", "asc", le, de, eu
'Dim ors As New Recordset

'ors.fields.Append "id", adBSTR
'ors.fields.Append "descripcion", adBSTR
'ors.fields.Append "titulo", adBSTR
'ors.fields.Append "fecha_baja", adBSTR

'ors.Open

'ors.AddNew

'ors(0) = 4

'ors.Update

'ors.save oDOM, adPersistXML

x.Search iID, xml, "username = 'gaston'", , , le, de, eu
Debug.Print le, de, eu
x.Read xml, le, de, eu

Debug.Print oDOM.xml
Debug.Print le, de, eu
'Debug.Print x.Titulo
'If Not x.GetUbicaciones(xml, , 1, le, de, eu) Then MsgBox le & de & eu



End Sub


Function CheckString(sValue)
Dim i
Dim invalidList
    i = 0
 
    ' set up a list of unacceptable characters
    ' this includes spaces, dashes and underscores
    ' you can leave these out of the list
    ' you may need to add other characters, e.g. copied from MSWord
 
    invalidList = ",<.>?;:'@#~]}[{=+)(*&^%$£!`¬| -_%!"
 
    ' check for " which can't be inside the string
 
    If InStr(sValue, Chr(34)) > 0 Then
       sValue = Replace(sValue, Chr(34), Chr(34) & Chr(34))
    Else
        ' loop through, making sure no characters
        ' are in the 'reserved characters' list
 
        For i = 1 To Len(invalidList)
            If InStr(sValue, Mid(invalidList, i, 1)) > 0 Then
                 sValue = Replace(sValue, Mid(invalidList, i, 1), Mid(invalidList, i, 1) & Mid(invalidList, i, 1))
            End If
        Next
    End If

    CheckString = sValue

End Function

Private Sub Command20_Click()
Dim le, de, eu
Dim x As New csubcategoria
Dim sXML
Dim asd As New ADODB.Stream

'x.Titulo = "2322"
'x.UsuarioID = 1
'x.Archivo = "<?xml version=""1.0""?> <root xmlns:dt=""urn:schemas-microsoft-com:datatypes""><file1 dt:dt=""bin.base64"">Z2FzdG9uIGFsZ2F6ZQ==</file1></root>"
'x.Add 1, le, de, eu

x.Search 5, sXML
Debug.Print sXML







Debug.Print le, de, eu
End Sub

Private Sub Command21_Click()
Dim le, de, eu
Dim x As New crol
Dim sXML
x.id = 7
x.Delete 1


End Sub

Private Sub Command22_Click()
Dim le, de, eu
Dim x As New citem
Dim sXML
x.id = 113
x.Delete 1
End Sub

Private Sub Command23_Click()
Dim le, de, eu
Dim x As New cficha
Dim sXML

x.Archivo = "<?xml version=""1.0""?> <root xmlns:dt=""urn:schemas-microsoft-com:datatypes""><file1 dt:dt=""bin.base64"">Z2FzdG9uIGFsZ2F6ZQ==</file1></root>"
x.Archivo_Nombre = "caca.txt"
x.Archivo_Tamaño = "123"
x.Titulo = "123"
x.UsuarioID = 20


x.Add 20
End Sub

Private Sub Command3_Click()
Dim le, de, eu
Dim x As New cUsuario
Dim sXML


With x
   '.id = 3
'    .Username = "username"
'    .Password = "password"
'    .nombre = "nombre123123"
'    .Apellido = "apellido"
'    .Mail = "mail"
'    .DNI = "dni"
'    .Matricula = "matricula"
'    .Fecha_Nacimiento = "20061212"
'    .Domicilio_Calle = "calle"
'    .Domicilio_Nro = "nro"
'    .Domicilio_Piso = "piso"
'    .Domicilio_Unidad = "unidad"
'    .Domicilio_Cod_Postal = "cod_postal"
'    .Tel1 = "tel1"
'    .Tel2 = "tel2"
 '.Search 1, sXML, "id = 1"
' .Add 1
'.Read 1, sXML
.Search 1, sXML, , "apellido", "asc"


MsgBox .Apellido

End With

Debug.Print le, de, eu
End Sub

Private Sub Command4_Click()
Dim le, ed, es
Dim x As New crol
Dim strXML As String
Dim z As New cPermiso
Dim a As New Dictionary



With x
z.RolID = 1
z.Tabla = 1
z.Permiso = 15
a.Add 1, z
Set z = Nothing
Set z = New cPermiso

z.RolID = 2
z.Tabla = 2
z.Permiso = 14
a.Add 2, z

x.SetPermisos a

End With

End Sub

Private Sub Command5_Click()
Dim le, ed, es
Dim x As New cRestriccion
Dim strXML As String
Dim z As New cPermiso
Dim a As New Dictionary



With x
'26 81 4 1234
.Search strXML, "id = 4"
.Read 1, strXML
'.Update
MsgBox .Campo


End With
End Sub

Private Sub Command6_Click()
Dim le, ed, es
Dim x As New cSecurityAgent
Dim s As String
x.Login s, "biblio", "123", le, ed, es
End Sub

Private Sub Command7_Click()
Dim le, ed, es
Dim oUsuario As New cUsuario
Dim s As String
Dim oRs As Recordset
Dim oDOM As MSXML2.DOMDocument

'x.Login s, "gaston", "123", le, ed, es

Set oUsuario = New cUsuario

    Set oRs = New ADODB.Recordset
    Set oDOM = New MSXML2.DOMDocument

    oRs.fields.Append "userID", adBSTR
    oRs.fields.Append "pwdold", adBSTR
    oRs.fields.Append "pwdnew1", adBSTR
    oRs.fields.Append "pwdnew2", adBSTR

    oRs.Open
    
    'categorias
    oRs.AddNew
    oRs(0) = "1"
    oRs(1) = "123"
    oRs(2) = "321"
    oRs(3) = "321"
    oRs.Update

    oRs.save oDOM, adPersistXML

    If oUsuario.ChangePassword("1", oDOM.xml, le, ed, es) Then
        Set oUsuario = Nothing
        Set oRs = Nothing
        Set oDOM = Nothing
        'response.redirect "roles_list.asp?msg=Objeto modificado con éxito."
    Else
        Set oUsuario = Nothing
        Set oRs = Nothing
        Set oDOM = Nothing
        'response.redirect "error.asp?title=" & strTitle & "&message=Error Nro: " & CStr(lErrNum) & " <BR> Descripción: " & sErrDesc & " <BR> Origen:  " & sErrSource
    End If
End Sub

Private Sub Command8_Click()
Dim le, ed, es
Dim x As New cCategoria
Dim s As String
Dim oRs As Recordset
Dim oDOM As MSXML2.DOMDocument

x.Search 1, s, "itemtipoID"
Debug.Print s


End Sub

Private Sub Command9_Click()
Dim le, ed, es
Dim x As New citem
Dim s As String
Dim strRes As String
Dim oRs As Recordset
Dim oDOM As MSXML2.DOMDocument

'x.id = 17
'x.Titulo = "asdf"
'x.Autor = "2222"
'x.Anno = "1999"
'x.CategoriaID = 1
'x.SubcategoriaID = 10
'x.EditorialID = 1
'x.ISBN = "2"
'x.UbicacionID = 52

'x.Update 1, le, ed, es
x.Search 1, s, "(autor LIKE '%doc%' OR titulo LIKE '%fue%') AND fecha_baja IS NULL "
Debug.Print s

Debug.Print le
Debug.Print ed
Debug.Print es


End Sub


Private Sub Command10_Click()
Dim le, ed, es
Dim x As New citem
Dim s As String
Dim strRes As String
Dim oRs As Recordset
Dim oDOM As MSXML2.DOMDocument

'x.id = 27
x.Search 1, s, "id = 30", "titulo", "asc"

'x.deletecopy 1, le, ed, es


x.Read 1, s, le, ed, es

Debug.Print le
Debug.Print ed
Debug.Print es
Debug.Print s


End Sub

Private Sub Command11_Click()
Dim le, ed, es
Dim x As New csubcategoria
Dim s As String

x.id = 5

x.Delete 1, le, ed, es

Debug.Print le & ed & es

End Sub

Private Sub Command12_Click()
Dim le, ed, es
Dim x As New cCategoria
Dim s As String

x.id = 15

x.Delete 1, le, ed, es

Debug.Print le & ed & es
End Sub

