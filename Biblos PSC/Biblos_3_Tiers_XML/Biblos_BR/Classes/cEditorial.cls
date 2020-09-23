VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEditorial"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'========================================================================
' Copyright © CHMR 2007 All rights reserved.
'========================================================================
'
'   Module         - cEditorial
'   Version Number - 1.0
'   Last Updated   - January 18th 2007 - 5:34 PM
'   Author         - Algaze, Gastón
'
'
'
'
'
'========================================================================
' This file contains trade secrets of CHMR No part
' may be reproduced or transmitted in any form by any means or for any purpose
' without the express written permission of CHMR.
'========================================================================

'========================================================================
'Variables miembro
'========================================================================
Private mID As Integer
Private mNombre As String
Private mMail As String
Private mWeb As String
Private mDomicilio_Calle As String
Private mDomicilio_Nro As String
Private mDomicilio_Piso As String
Private mDomicilio_Unidad As String
Private mDomicilio_Cod_Postal As String
Private mTel1 As String
Private mTel2 As String
Private mFecha_Alta As String
Private mFecha_Ult_Act As String
Private mFecha_Baja As String

'========================================================================
'Propiedades
'========================================================================
Public Property Get ID() As Integer
    ID = mID
End Property

Public Property Let ID(ByVal vNewValue As Integer)
    mID = vNewValue
End Property

Public Property Get Nombre() As String
    Nombre = mNombre
End Property

Public Property Let Nombre(ByVal vNewValue As String)
    mNombre = vNewValue
End Property

Public Property Get Mail() As String
    Mail = mMail
End Property

Public Property Let Mail(ByVal vNewValue As String)
    mMail = vNewValue
End Property

Public Property Get Web() As String
    Web = mWeb
End Property

Public Property Let Web(ByVal vNewValue As String)
    mWeb = vNewValue
End Property

Public Property Get Domicilio_Calle() As String
    Domicilio_Calle = mDomicilio_Calle
End Property

Public Property Let Domicilio_Calle(ByVal vNewValue As String)
    mDomicilio_Calle = vNewValue
End Property

Public Property Get Domicilio_Piso() As String
    Domicilio_Piso = mDomicilio_Piso
End Property

Public Property Let Domicilio_Piso(ByVal vNewValue As String)
    mDomicilio_Piso = vNewValue
End Property

Public Property Get Domicilio_Unidad() As String
    Domicilio_Unidad = mDomicilio_Unidad
End Property

Public Property Let Domicilio_Unidad(ByVal vNewValue As String)
    mDomicilio_Unidad = vNewValue
End Property

Public Property Get Domicilio_Nro() As String
    Domicilio_Nro = mDomicilio_Nro
End Property

Public Property Let Domicilio_Nro(ByVal vNewValue As String)
    mDomicilio_Nro = vNewValue
End Property

Public Property Get Domicilio_Cod_Postal() As String
    Domicilio_Cod_Postal = mDomicilio_Cod_Postal
End Property

Public Property Let Domicilio_Cod_Postal(ByVal vNewValue As String)
    mDomicilio_Cod_Postal = vNewValue
End Property

Public Property Get Tel1() As String
    Tel1 = mTel1
End Property

Public Property Let Tel1(ByVal vNewValue As String)
    mTel1 = vNewValue
End Property

Public Property Get Tel2() As String
    Tel2 = mTel2
End Property

Public Property Let Tel2(ByVal vNewValue As String)
    mTel2 = vNewValue
End Property

Public Property Get Fecha_Alta() As String
    Fecha_Alta = mFecha_Alta
End Property

Public Property Let Fecha_Alta(ByVal vNewValue As String)
    mFecha_Alta = vNewValue
End Property

Public Property Get Fecha_Ult_Act() As String
    Fecha_Ult_Act = mFecha_Ult_Act
End Property

Public Property Let Fecha_Ult_Act(ByVal vNewValue As String)
    mFecha_Ult_Act = vNewValue
End Property

Public Property Get Fecha_Baja() As String
    Fecha_Baja = mFecha_Baja
End Property

Public Property Let Fecha_Baja(ByVal vNewValue As String)
    mFecha_Baja = vNewValue
End Property


'========================================================================
'Métodos y Funciones
'========================================================================
'========================================================================
'
'   Title          - Función ValidateMe()
'   Description    - Valida que las propiedades del objeto sean correctas
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Private Function ValidateMe(Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    ValidateMe = False
    If Len(CStr(ID)) > 0 And ID <> 0 Then
        If Len(Nombre) = 0 Or _
            Len(Nombre) > 255 Or _
            Len(Mail) > 100 Or _
            Len(Mail) = 0 Or _
            Len(Domicilio_Calle) > 255 Or _
            Len(Domicilio_Calle) = 0 Or _
            Len(Domicilio_Nro) > 10 Or _
            Len(Domicilio_Nro) = 0 Or _
            Len(Domicilio_Cod_Postal) > 12 Or _
            Len(Domicilio_Cod_Postal) = 0 Or _
            Len(Tel1) > 16 Or _
            Len(Tel1) = 0 Or _
            Not IsNumeric(ID) Then
                ValidateMe = False
                Err.Raise -1, "cEditorial.ValidateMe", "Longitud o tipo de parámetros incorrecto."
        Else
            ValidateMe = True
        End If
    Else
        If Len(Nombre) = 0 Or _
            Len(Nombre) > 255 Or _
            Len(Mail) > 30 Or _
            Len(Mail) = 0 Or _
            Len(Domicilio_Calle) > 255 Or _
            Len(Domicilio_Calle) = 0 Or _
            Len(Domicilio_Nro) > 10 Or _
            Len(Domicilio_Nro) = 0 Or _
            Len(Domicilio_Cod_Postal) > 12 Or _
            Len(Domicilio_Cod_Postal) = 0 Or _
            Len(Tel1) > 16 Or _
            Len(Tel1) = 0 Then
                ValidateMe = False
                Err.Raise -1, "cEditorial.ValidateMe", "Longitud o tipo de parámetros incorrecto."
        Else
            ValidateMe = True
        End If
    End If
End Function

'========================================================================
'
'   Title          - Función Add()
'   Description    - Inserta un nuevo elemento
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Add(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo AddErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oHelper As cEditorial
Dim oSecAgent As New cSecurityAgent

    Add = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cEditorial", opInsert, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If Not ValidateMe(lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    'Primero me fijo que no exista el nuevo objeto en el sistema
    If Search(iUserID, strXML, "nombre = '" & Nombre & "' OR mail = '" & Mail & "'", , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cEditorial
        If oHelper.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cEditorial.Add", "El nuevo objeto ya existe en el sistema.<BR>Por favor verifique los datos."
        Else
            If lErrNum = "-2" Then
                Err.Raise -1, "cEditorial.Read", "Hay mas de un registro para leer en el archivo XML"
            End If
                       
            oRs.fields.Append "nombre", adBSTR
            oRs.fields.Append "mail", adBSTR
            oRs.fields.Append "web", adBSTR
            oRs.fields.Append "domicilio_calle", adBSTR
            oRs.fields.Append "domicilio_nro", adBSTR
            oRs.fields.Append "domicilio_piso", adBSTR
            oRs.fields.Append "domicilio_unidad", adBSTR
            oRs.fields.Append "domicilio_cod_postal", adBSTR
            oRs.fields.Append "tel1", adBSTR
            oRs.fields.Append "tel2", adBSTR
            oRs.fields.Append "fecha_alta", adBSTR
            oRs.fields.Append "fecha_ult_act", adBSTR
            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = Nombre
            oRs(1) = Mail
            oRs(2) = Web
            oRs(3) = Domicilio_Calle
            oRs(4) = Domicilio_Nro
            oRs(5) = Domicilio_Piso
            oRs(6) = Domicilio_Unidad
            oRs(7) = Domicilio_Cod_Postal
            oRs(8) = Tel1
            oRs(9) = Tel2
            oRs(10) = Format(Now(), "yyyymmdd")
            oRs(11) = Format(Now(), "yyyymmdd")
            
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Add = oDataAccessTier.SaveEditorial(oDOM.xml, lErrNum, sErrDesc, sErrSource)
            
            Set oRs = Nothing
            Set oDOM = Nothing
        End If
    End If
    
    If Add Then
        Exit Function
    Else
    
AddErr:
        
        Set oRs = Nothing
        Set oDOM = Nothing
        
        ShowError "cEditorial.Add", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
        
    End If
    
End Function

'========================================================================
'
'   Title          - Función Read()
'   Description    - Carga al objeto, leyendo desde un archivo XML
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Read(iUserID As Variant, sXML As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo ReadErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oSecAgent As New cSecurityAgent

    Read = False
    
    strXML = CStr(sXML)
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cEditorial", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    Set oRs = RecordsetFromXMLString(strXML)
    
    If Not oRs.EOF Then
        If oRs.RecordCount > 1 Then 'esto despues lo voy a tener uqe cambiar por el tema de los prestamos y del wrapper
            Err.Raise -1, "cEditorial.Read", "Hay mas de un registro para leer en el archivo XML"
        Else
            ID = oRs(0)
            Nombre = oRs(1)
            Mail = oRs(2)
            Web = oRs(3)
            Domicilio_Calle = oRs(4)
            Domicilio_Nro = oRs(5)
            Domicilio_Piso = oRs(6)
            Domicilio_Unidad = oRs(7)
            Domicilio_Cod_Postal = oRs(8)
            Tel1 = oRs(9)
            Tel2 = oRs(10)
            Fecha_Alta = IIf(IsNull(oRs("fecha_alta")), "", oRs("fecha_alta"))
            Fecha_Ult_Act = IIf(IsNull(oRs("fecha_ult_act")), "", oRs("fecha_ult_act"))
            Fecha_Baja = IIf(IsNull(oRs("fecha_baja")), "", oRs("fecha_baja"))

            Read = True
        End If
    Else
        Err.Raise -1, "cEditorial.Read", "No se encontro ningún registro para leer en el archivo XML"
    Set oRs = Nothing
    End If
    
    If Read Then
        Exit Function
    Else
    
ReadErr:
        
        Set oRs = Nothing
    
        ShowError "cEditorial.Read", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Search()
'   Description    - Realiza busquedas, devolviendo un archivo XML con el
'                    resultado.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Search(iUserID As Variant, sXML As Variant, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo searchErr

Dim strRestricciones As String
Dim strXML As String
Dim oSecAgent As New cSecurityAgent

    Search = False
    
    sSearch = Replace(sSearch, "id ", "Editoriales.id_Editorial", , , vbTextCompare)
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cEditorial", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    Else ' me fijo si tiene algun restriccion
        If oSecAgent.GetRestrictions(CInt(iUserID), "cEditorial", strRestricciones) Then
            If (Len(sSearch) > 0) And (Len(strRestricciones) > 0) Then
                sSearch = strRestricciones & "AND " & sSearch
            Else
                If Len(sSearch) = 0 Then
                    sSearch = strRestricciones
                End If
            End If
        End If
    End If
    
    
        If Len(sOrderBy) > 0 Then
            If oSecAgent.ValidateOrderBy(sOrderBy, "cEditorial") Then
                If Len(sSort) > 0 Then
                    If Not oSecAgent.ValidateSort(sSort) Then
                        Err.Raise -1, "cEditorial.Search", """SORT"" inválido, solo se permite ASC o DESC."
                    End If
                End If
            Else
                Err.Raise -1, "cEditorial.Search", """ORDER BY"" inválido, los campos no corresponden."
            End If
        End If

    Search = oDataAccessTier.SearchEditorial(strXML, sSearch, sOrderBy, sSort, lErrNum, sErrDesc, sErrSource)
    
    If Search Then
        sXML = strXML
        Exit Function
    Else
    
searchErr:
        ShowError "cEditorial.Search", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
        
    End If
    
End Function

'========================================================================
'
'   Title          - Función Delete()
'   Description    - Elimina lógica o físicamente al objeto.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Delete(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo DeleteErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim oItem As New cItem
Dim strXML As String
Dim oSecAgent As New cSecurityAgent

    Delete = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cEditorial", opDelete, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If Not IsNumeric(ID) Then
        Err.Raise -1, "cEditorial.Delete", "Longitud o tipo de parámetro incorrecto."
    End If
    
    oRs.fields.Append "id", adBSTR
    
    oRs.open
    
    oRs.AddNew

    oRs(0) = ID

    oRs.Update
    
    oRs.save oDOM, adPersistXML
    
    If oItem.Search(iUserID, strXML, "editorialID = " & ID, , , lErrNum, sErrDesc, sErrSource) Then
        If oItem.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cEditorial.Delete", "La editorial no se puede eliminar existen Items que la utilizan."
        Else
            Delete = oDataAccessTier.Deleteeditorial(oDOM.xml, 1, lErrNum, sErrDesc, sErrSource)
        End If
    End If
    
    Set oRs = Nothing
    Set oDOM = Nothing
    
    If Delete Then
        Exit Function
    Else
    
DeleteErr:
        Set oRs = Nothing
        Set oDOM = Nothing
    
        ShowError "cEditorial.Delete", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Update()
'   Description    - Actualiza al objeto.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Update(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo UpdateErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oHelper As cEditorial
Dim oSecAgent As New cSecurityAgent

    Update = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cEditorial", opUpdate, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If Search(iUserID, strXML, "id <> " & CStr(ID) & " AND nombre = '" & Nombre & "'", , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cEditorial
        If oHelper.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cEditorial.Update", "El nuevo objeto ya existe en el sistema.<BR>Por favor verifique los datos."
        Else
            If lErrNum = "-2" Then
                Err.Raise -1, "cEditorial.Update", "Hay mas de un registro para leer en el archivo XML"
            End If
            
            If Not ValidateMe Then
                Err.Raise -1, "cEditorial.Update", "Longitud o tipo de parámetro incorrecto."
            End If
            
            oRs.fields.Append "nombre", adBSTR
            oRs.fields.Append "mail", adBSTR
            oRs.fields.Append "web", adBSTR
            oRs.fields.Append "domicilio_calle", adBSTR
            oRs.fields.Append "domicilio_nro", adBSTR
            oRs.fields.Append "domicilio_piso", adBSTR
            oRs.fields.Append "domicilio_unidad", adBSTR
            oRs.fields.Append "domicilio_cod_postal", adBSTR
            oRs.fields.Append "tel1", adBSTR
            oRs.fields.Append "tel2", adBSTR
            oRs.fields.Append "fecha_alta", adBSTR
            oRs.fields.Append "fecha_ult_act", adBSTR
            oRs.fields.Append "id", adBSTR
            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = Nombre
            oRs(1) = Mail
            oRs(2) = Web
            oRs(3) = Domicilio_Calle
            oRs(4) = Domicilio_Nro
            oRs(5) = Domicilio_Piso
            oRs(6) = Domicilio_Unidad
            oRs(7) = Domicilio_Cod_Postal
            oRs(8) = Tel1
            oRs(9) = Tel2
            oRs(10) = Format(Now(), "yyyymmdd")
            oRs(11) = Format(Now(), "yyyymmdd")
            oRs(12) = ID
            
            oRs.Update
            
            oRs.save oDOM, adPersistXML
    
            Update = oDataAccessTier.UpdateEditorial(oDOM.xml, lErrNum, sErrDesc, sErrSource)
            Set oHelper = Nothing
            Set oRs = Nothing
            Set oDOM = Nothing
        End If
    Else
        Err.Raise -1, "cEditorial.Update", "No se pudo comprobar la integridad referencial."
    End If
    
    If Update Then
        Exit Function
    Else
    
UpdateErr:
        Set oHelper = Nothing
        Set oRs = Nothing
        Set oDOM = Nothing

        ShowError "cEditorial.Update", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Recover()
'   Description    - Recupera lógicamente al objeto.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Recover(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo RecoverErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim oSecAgent As New cSecurityAgent

    Recover = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cEditorial", opUpdate, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If Not IsNumeric(ID) Then
        Err.Raise -1, "cEditorial.Update", "Longitud o tipo de parámetro incorrecto."
    End If
    
    oRs.fields.Append "id", adBSTR
    
    oRs.open
    
    oRs.AddNew

    oRs(0) = ID

    oRs.Update
    
    oRs.save oDOM, adPersistXML
    
    Recover = oDataAccessTier.RecoverEditorial(oDOM.xml, lErrNum, sErrDesc, sErrSource)
    
    Set oRs = Nothing
    Set oDOM = Nothing
    
    If Recover Then
        Exit Function
    Else
    
RecoverErr:
        Set oRs = Nothing
        Set oDOM = Nothing
    
        ShowError "cEditorial.Recover", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function


