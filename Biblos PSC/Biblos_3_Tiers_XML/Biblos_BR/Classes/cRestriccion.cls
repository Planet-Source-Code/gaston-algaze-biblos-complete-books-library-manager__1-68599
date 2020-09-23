VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRestriccion"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'========================================================================
' Copyright © CHMR 2007 All rights reserved.
'========================================================================
'
'   Module         - cRestriccion
'   Version Number - 1.0
'   Last Updated   - February 25th 2007 - 4:34 PM
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
Private mRolID As Integer
Private mTabla As String
Private mCampo As String
Private mOperacionID As Integer
Private mOperacion As String
Private mValor As String
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

Public Property Get RolID() As Integer
    RolID = mRolID
End Property

Public Property Let RolID(ByVal vNewValue As Integer)
    mRolID = vNewValue
End Property

Public Property Get OperacionID() As Integer
    OperacionID = mOperacionID
End Property

Public Property Let OperacionID(ByVal vNewValue As Integer)
    mOperacionID = vNewValue
End Property

Public Property Get Valor() As String
    Valor = mValor
End Property

Public Property Let Valor(ByVal vNewValue As String)
    mValor = vNewValue
End Property

Public Property Get Tabla() As String
    Tabla = mTabla
End Property

Public Property Let Tabla(ByVal vNewValue As String)
    mTabla = vNewValue
End Property

Public Property Get Campo() As String
    Campo = mCampo
End Property

Public Property Let Campo(ByVal vNewValue As String)
    mCampo = vNewValue
End Property

Public Property Get Operacion() As String
    Operacion = mOperacion
End Property

Public Property Let Operacion(ByVal vNewValue As String)
    mOperacion = vNewValue
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
'Métodos y Restricciones
'========================================================================
'========================================================================
'
'   Title          - Función ValidateMe()
'   Description    - Valida que las propiedades del objeto sean correctas
'   Version Number - 1.0
'   Last Updated   - February 26th 2007 - 6:49 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Private Function ValidateMe(Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    ValidateMe = False
    If Len(CStr(ID)) > 0 Then
        If Len(Valor) > 255 Or _
           Len(Valor) = 0 Or Not IsNumeric(ID) Or _
           Not IsNumeric(OperacionID) Then
            ValidateMe = False
            Err.Raise -1, "cRestriccion.ValidateMe", "Longitud o tipo de parámetros incorrecto."
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
'   Last Updated   - February 26th 2007 - 6:49 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Add(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo AddErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oHelper As cRestriccion
Dim oSecAgent As New cSecurityAgent

    Add = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cRestriccion", opInsert, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If Not ValidateMe Then
        Err.Raise -1, "cRestriccion.Add", "Longitud de parámetros incorrecta."
    End If
    
    'Primero me fijo que no exista el nuevo objeto en el sistema
    If Search(iUserID, strXML, "rolID = " & RolID & " AND valor = '" & Valor & "' AND campo = '" & Campo & "' AND tabla = '" & Tabla & "' AND operacionID = " & OperacionID & " AND fecha_baja IS NULL", , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cRestriccion
        If oHelper.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cRestriccion.Add", "El nuevo objeto ya existe en el sistema.<BR>Por favor verifique los datos."
        Else
            If lErrNum = "-2" Then
                Err.Raise -1, "cRestriccion.Read", "Hay mas de un registro para leer en el archivo XML"
            End If
            
            'Si no existe lo creo...
            oRs.fields.Append "valor", adBSTR
            oRs.fields.Append "fecha_alta", adBSTR
            oRs.fields.Append "fecha_ult_act", adBSTR
            oRs.fields.Append "campo", adBSTR
            oRs.fields.Append "tabla", adBSTR
            oRs.fields.Append "id_operacion", adBSTR
            oRs.fields.Append "id_rol", adBSTR
            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = Valor
            oRs(1) = Format(Now(), "yyyymmdd")
            oRs(2) = Format(Now(), "yyyymmdd")
            oRs(3) = Campo
            oRs(4) = Tabla
            oRs(5) = OperacionID
            oRs(6) = RolID
            
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Add = oDataAccessTier.SaveRestriccion(oDOM.xml, lErrNum, sErrDesc, sErrSource)
            
            Set oHelper = Nothing
            Set oRs = Nothing
            Set oDOM = Nothing
        End If
    End If
    
    If Add Then
        Exit Function
    Else
    
AddErr:
        
        Set oHelper = Nothing
        Set oRs = Nothing
        Set oDOM = Nothing
    
        ShowError "cRestriccion.Add", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Read()
'   Description    - Carga al objeto, leyendo desde un archivo XML
'   Version Number - 1.0
'   Last Updated   - February 26th 2007 - 6:49 PM
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
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cRestriccion", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    strXML = CStr(sXML)
    
    Set oRs = RecordsetFromXMLString(strXML)
    
    If Not oRs.EOF Then
        If oRs.RecordCount > 1 Then
            Err.Raise -2, "cRestriccion.Read", "Hay mas de un registro para leer en el archivo XML"
        Else
            ID = oRs("id")
            Valor = oRs("valor")
            Tabla = oRs("tabla")
            Campo = oRs("campo")
            OperacionID = oRs("operacionID")
            RolID = oRs("rolID")
            Fecha_Alta = IIf(IsNull(oRs("fecha_alta")), "", oRs("fecha_alta"))
            Fecha_Ult_Act = IIf(IsNull(oRs("fecha_ult_act")), "", oRs("fecha_ult_act"))
            Fecha_Baja = IIf(IsNull(oRs("fecha_baja")), "", oRs("fecha_baja"))
            Read = True
        End If
    Else
        Err.Raise -1, "cRestriccion.Read", "No se encontro ningún registro para leer en el archivo XML"
    Set oRs = Nothing
    End If
    
    If Read Then
        Exit Function
    Else
    
ReadErr:
        
        Set oRs = Nothing
    
        ShowError "cRestriccion.Read", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Search()
'   Description    - Realiza busquedas, devolviendo un archivo XML con el
'                    resultado.
'   Version Number - 1.0
'   Last Updated   - February 26th 2007 - 6:49 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Search(iUserID As Variant, sXML As Variant, Optional sSearch As String, Optional sOrderBy As String, Optional sSort As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo searchErr

Dim strXML As String
Dim strRestricciones As String
Dim oSecAgent As New cSecurityAgent

    Search = False
    
    sSearch = Replace(sSearch, "RolID ", "roles.id_rol ", , , vbTextCompare)
    sSearch = Replace(sSearch, "fecha_baja ", "restricciones.fecha_baja ", , , vbTextCompare)
    sSearch = Replace(sSearch, "operacionID ", "restricciones.id_operacion ", , , vbTextCompare)
    sSearch = Replace(sSearch, "id ", "restricciones.id_restriccion ", , , vbTextCompare)
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cRestriccion", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    Else ' me fijo si tiene algun restriccion
        If oSecAgent.GetRestrictions(CInt(iUserID), "cRestriccion", strRestricciones) Then
            If (Len(sSearch) > 0) And (Len(strRestricciones) > 0) Then
                sSearch = strRestricciones & "AND " & sSearch
            Else
                If Len(sSearch) = 0 Then
                    sSearch = strRestricciones
                End If
            End If
        End If
    End If
    
        'If Validate(sSearch) Then
            If Len(sOrderBy) > 0 Then
                If oSecAgent.ValidateOrderBy(sOrderBy, "cRestriccion") Then
                    If Len(sSort) > 0 Then
                        If Not oSecAgent.ValidateSort(sSort) Then
                            Err.Raise -1, "cRestriccion.Search", """SORT"" inválido, solo se permite ASC o DESC."
                        End If
                    End If
                Else
                    Err.Raise -1, "cRestriccion.Search", """ORDER BY"" inválido, los campos no corresponden."
                End If
            End If
        'Else
        '    Err.Raise -1, "cRestriccion.Search", """SEARCH"" inválido, los campos no corresponden."
        'End If

    Search = oDataAccessTier.SearchRestriccion(strXML, sSearch, sOrderBy, sSort, lErrNum, sErrDesc, sErrSource)
    
    If Search Then
        sXML = strXML
        Exit Function
    Else
    
searchErr:
    
        ShowError "cRestriccion.Search", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función SearchByRol()
'   Description    - Realiza busquedas por rol, devolviendo un archivo XML con el
'                    resultado.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function SearchByRol(iUserID As Variant, sXML As Variant, iRol As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo SearchByRolErr

Dim strXML As String
Dim oSecAgent As New cSecurityAgent

    SearchByRol = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cRestriccion", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If

    SearchByRol = oDataAccessTier.SearchRestriccionByRol(strXML, CInt(iRol), lErrNum, sErrDesc, sErrSource)
    
    If SearchByRol Then
        sXML = strXML
        Exit Function
    Else
    
SearchByRolErr:
        ShowError "cRestriccion.SearchByRol", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource

    End If
    
End Function

'========================================================================
'
'   Title          - Función Delete()
'   Description    - Elimina lógica o físicamente al objeto.
'   Version Number - 1.0
'   Last Updated   - February 26th 2007 - 6:49 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Delete(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo DeleteErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim oSecAgent As New cSecurityAgent

    Delete = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cRestriccion", opDelete) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
   
    If Not IsNumeric(ID) Then
        Err.Raise -1, "cRestriccion.Delete", "Longitud o tipo de parámetro incorrecto."
    End If
    
    oRs.fields.Append "id", adBSTR
    
    oRs.open
    
    oRs.AddNew

    oRs(0) = ID

    oRs.Update
    
    oRs.save oDOM, adPersistXML
    
    Delete = oDataAccessTier.DeleteRestriccion(oDOM.xml, 1, lErrNum, sErrDesc, sErrSource)
    
    Set oRs = Nothing
    Set oDOM = Nothing
    
    If Delete Then
        Exit Function
    Else
    
DeleteErr:
        Set oRs = Nothing
        Set oDOM = Nothing
    
        ShowError "cRestriccion.Delete", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Update()
'   Description    - Actualiza al objeto.
'   Version Number - 1.0
'   Last Updated   - February 26th 2007 - 6:49 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Update(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo UpdateErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oHelper As cRestriccion
Dim oSecAgent As New cSecurityAgent

    Update = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cRestriccion", opUpdate, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If Search(iUserID, strXML, "id <> " & CStr(ID) & " AND rolID = " & RolID & " AND valor = '" & Valor & "' AND campo = '" & Campo & "' AND tabla = '" & Tabla & "' AND operacionID = " & OperacionID & " AND fecha_baja IS NULL", , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cRestriccion
        If oHelper.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cRestriccion.Update", "El nuevo objeto ya existe en el sistema.<BR>Por favor verifique los datos."
        Else
            If lErrNum = "-2" Then
                Err.Raise -1, "cRestriccion.Update", "Hay mas de un registro para leer en el archivo XML"
            End If
            If Not ValidateMe Then
                Err.Raise -1, "cRestriccion.Update", "Longitud o tipo de parámetro incorrecto."
            End If
            
            oRs.fields.Append "valor", adBSTR
            oRs.fields.Append "fecha_ult_act", adBSTR
            oRs.fields.Append "fecha_baja", adBSTR, , adFldIsNullable
            oRs.fields.Append "id", adBSTR
            oRs.fields.Append "campo", adBSTR
            oRs.fields.Append "tabla", adBSTR
            oRs.fields.Append "id_operacion", adBSTR
            
            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = Valor
            oRs(1) = Format(Now(), "yyyymmdd")
            oRs(2) = Fecha_Baja
            oRs(3) = ID
            oRs(4) = Campo
            oRs(5) = Tabla
            oRs(6) = OperacionID
        
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Update = oDataAccessTier.UpdateRestriccion(oDOM.xml, lErrNum, sErrDesc, sErrSource)
                            
            Set oHelper = Nothing
            Set oRs = Nothing
            Set oDOM = Nothing
        End If
    Else
        Err.Raise -1, "cRestriccion.Update", "No se pudo comprobar la integridad referencial."
    End If
    
    If Update Then
        Exit Function
    Else
    
UpdateErr:
        Set oHelper = Nothing
        Set oRs = Nothing
        Set oDOM = Nothing

        ShowError "cRestriccion.Update", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'
'   Title          - Función Recover()
'   Description    - Recupera lógicamente al objeto.
'   Version Number - 1.0
'   Last Updated   - February 26th 2007 - 6:49 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Recover(iUserID As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo RecoverErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String
Dim oHelper As cRestriccion
Dim oSecAgent As New cSecurityAgent

    Recover = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cRestriccion", opUpdate, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
        
    Set oHelper = New cRestriccion
        
    If Not IsNumeric(ID) Then
        Err.Raise -1, "cRestriccion.Recover", "Longitud o tipo de parámetro incorrecto."
    End If
    
    If Search(iUserID, strXML, "id = " & CStr(ID), , , lErrNum, sErrDesc, sErrSource) Then
        oHelper.Read strXML, lErrNum, sErrDesc, sErrSource
    End If
    
    If Search(iUserID, strXML, "id <> " & CStr(ID) & " AND valor = '" & oHelper.Valor & "' AND campo = '" & Campo & "' AND tabla = '" & Tabla & "' AND id_operacion = " & OperacionID, , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cRestriccion
        If oHelper.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cRestriccion.Recover", "El nuevo objeto ya existe en el sistema.<BR>Por favor verifique los datos."
        Else
            If lErrNum = "-2" Then
                Err.Raise -1, "cRestriccion.Read", "Hay mas de un registro para leer en el archivo XML"
            End If
            
            oRs.fields.Append "id", adBSTR
            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = ID
        
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Recover = oDataAccessTier.RecoverRestriccion(oDOM.xml, lErrNum, sErrDesc, sErrSource)
            
            Set oHelper = Nothing
            Set oRs = Nothing
            Set oDOM = Nothing
        End If
    End If
    
    If Recover Then
        Exit Function
    Else
    
RecoverErr:
        
        Set oHelper = Nothing
        Set oRs = Nothing
        Set oDOM = Nothing
    
        ShowError "cRestriccion.Recover", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function
