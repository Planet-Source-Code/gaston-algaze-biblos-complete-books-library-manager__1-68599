VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUbicacion"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'========================================================================
' Copyright © CHMR 2007 All rights reserved.
'========================================================================
'
'   Module         - cUbicacion
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
Private mDescripcion As String
Private mTitulo As String
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

Public Property Get Descripcion() As String
    Descripcion = mDescripcion
End Property

Public Property Let Descripcion(ByVal vNewValue As String)
    mDescripcion = vNewValue
End Property

Public Property Get Titulo() As String
    Titulo = mTitulo
End Property

Public Property Let Titulo(ByVal vNewValue As String)
    mTitulo = vNewValue
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
    If Len(CStr(ID)) > 0 Then
        If Len(Descripcion) > 255 Or Len(Titulo) > 30 Or _
           Len(Titulo) = 0 Or Not IsNumeric(ID) Then
            ValidateMe = False
            Err.Raise -1, "cUbicacion.ValidateMe", "Longitud o tipo de parámetros incorrecto."
        Else
            ValidateMe = True
        End If
    Else
        If Len(Descripcion) > 255 Or Len(Titulo) > 30 Or _
           Len(Titulo) = 0 Then
            ValidateMe = False
            Err.Raise -1, "cUbicacion.ValidateMe", "Longitud o tipo de parámetros incorrecto."
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
Dim oHelper As cUbicacion
Dim oSecAgent As New cSecurityAgent

    Add = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cUbicacion", opInsert, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If Not ValidateMe Then
        Err.Raise -1, "cUbicacion.Add", "Longitud de parámetros incorrecta."
    End If
    
    'Primero me fijo que no exista el nuevo objeto en el sistema
    If Search(iUserID, strXML, "titulo = '" & Titulo & "' AND fecha_baja IS NULL", , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cUbicacion
        If oHelper.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cUbicacion.Add", "El nuevo objeto ya existe en el sistema.<BR>Por favor verifique los datos."
        Else
            If lErrNum = "-2" Then
                Err.Raise -1, "cUbicacion.Read", "Hay mas de un registro para leer en el archivo XML"
            End If
            
            oRs.fields.Append "descripcion", adBSTR
            oRs.fields.Append "titulo", adBSTR
            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = Descripcion
            oRs(1) = Titulo
        
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Add = oDataAccessTier.SaveUbicacion(oDOM.xml, lErrNum, sErrDesc, sErrSource)
            
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
    
        ShowError "cUbicacion.Add", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
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
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cUbicacion", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    strXML = CStr(sXML)
    
    Set oRs = RecordsetFromXMLString(strXML)
    
    If Not oRs.EOF Then
        If oRs.RecordCount > 1 Then
            Err.Raise -2, "cUbicacion.Read", "Hay mas de un registro para leer en el archivo XML"
        Else
            ID = oRs("id")
            Descripcion = oRs("descripcion")
            Titulo = oRs("titulo")
            Fecha_Alta = IIf(IsNull(oRs("fecha_alta")), "", oRs("fecha_alta"))
            Fecha_Ult_Act = IIf(IsNull(oRs("fecha_ult_act")), "", oRs("fecha_ult_act"))
            Fecha_Baja = IIf(IsNull(oRs("fecha_baja")), "", oRs("fecha_baja"))
            Read = True
        End If
    Else
        Err.Raise -1, "cUbicacion.Read", "No se encontro ningún registro para leer en el archivo XML"
    Set oRs = Nothing
    End If
    
    If Read Then
        Exit Function
    Else
    
ReadErr:
        
        Set oRs = Nothing
    
        ShowError "cUbicacion.Read", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
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

Dim strXML As String
Dim oSecAgent As New cSecurityAgent
Dim strRestricciones As String

    Search = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cUbicacion", opRead, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    Else ' me fijo si tiene algun restriccion
        If oSecAgent.GetRestrictions(CInt(iUserID), "cUbicacion", strRestricciones) Then
            If (Len(sSearch) > 0) And (Len(strRestricciones) > 0) Then
                sSearch = strRestricciones & "AND " & sSearch
            Else
                If Len(sSearch) = 0 Then
                    sSearch = strRestricciones
                End If
            End If
        End If
    End If
    
    sSearch = Replace(sSearch, "id ", "ubicaciones.id_ubicacion ", , , vbTextCompare)
    
    
        If Len(sOrderBy) > 0 Then
            If oSecAgent.ValidateOrderBy(sOrderBy, "cUbicacion") Then
                If Len(sSort) > 0 Then
                    If Not oSecAgent.ValidateSort(sSort) Then
                        Err.Raise -1, "cUbicacion.Search", """SORT"" inválido, solo se permite ASC o DESC."
                    End If
                End If
            Else
                Err.Raise -1, "cUbicacion.Search", """ORDER BY"" inválido, los campos no corresponden."
            End If
        End If

    Search = oDataAccessTier.SearchUbicacion(strXML, sSearch, sOrderBy, sSort, lErrNum, sErrDesc, sErrSource)
    
    If Search Then
        sXML = strXML
        Exit Function
    Else
    
searchErr:
    
        ShowError "cUbicacion.Search", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
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
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cUbicacion", opDelete, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
   
    If Not IsNumeric(ID) Then
        Err.Raise -1, "cUbicacion.Update", "Longitud o tipo de parámetro incorrecto."
    End If
    
    oRs.fields.Append "id", adBSTR
    
    oRs.open
    
    oRs.AddNew

    oRs(0) = ID

    oRs.Update
    
    oRs.save oDOM, adPersistXML
    
    If oItem.Search(iUserID, strXML, "ubicacionID = " & ID, , , lErrNum, sErrDesc, sErrSource) Then
        If oItem.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cUbicacion.Delete", "La ubicacion no se puede eliminar existen Items que la utilizan."
        Else
            Delete = oDataAccessTier.DeleteUbicacion(oDOM.xml, 1, lErrNum, sErrDesc, sErrSource)
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
    
        ShowError "cUbicacion.Delete", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
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
Dim oHelper As cUbicacion
Dim oSecAgent As New cSecurityAgent

    Update = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cUbicacion", opUpdate, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
    
    If Search(iUserID, strXML, "id <> " & CStr(ID) & " AND titulo = '" & Titulo & "' AND fecha_baja IS NULL", , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cUbicacion
        If oHelper.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cUbicacion.Update", "El nuevo objeto ya existe en el sistema.<BR>Por favor verifique los datos."
        Else
            If lErrNum = "-2" Then
                Err.Raise -1, "cUbicacion.Update", "Hay mas de un registro para leer en el archivo XML"
            End If
            If Not ValidateMe Then
                Err.Raise -1, "cUbicacion.Update", "Longitud o tipo de parámetro incorrecto."
            End If
            
            oRs.fields.Append "descripcion", adBSTR
            oRs.fields.Append "titulo", adBSTR
            oRs.fields.Append "fecha_baja", adBSTR, , adFldIsNullable
            oRs.fields.Append "id", adBSTR
            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = Descripcion
            oRs(1) = Titulo
            oRs(2) = Fecha_Baja
            oRs(3) = ID
        
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Update = oDataAccessTier.UpdateUbicacion(oDOM.xml, lErrNum, sErrDesc, sErrSource)
                            
            Set oHelper = Nothing
            Set oRs = Nothing
            Set oDOM = Nothing
        End If
    Else
        Err.Raise -1, "cUbicacion.Update", "No se pudo comprobar la integridad referencial."
    End If
    
    If Update Then
        Exit Function
    Else
    
UpdateErr:
        Set oHelper = Nothing
        Set oRs = Nothing
        Set oDOM = Nothing

        ShowError "cUbicacion.Update", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
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
Dim strXML As String
Dim oHelper As cUbicacion
Dim oSecAgent As New cSecurityAgent

    Recover = False
    
    If Not oSecAgent.IsAllowed(CInt(iUserID), "cUbicacion", opUpdate, lErrNum, sErrDesc, sErrSource) Then
        Err.Raise lErrNum, sErrSource, sErrDesc
    End If
        
    Set oHelper = New cUbicacion
        
    If Not IsNumeric(ID) Then
        Err.Raise -1, "cUbicacion.Recover", "Longitud o tipo de parámetro incorrecto."
    End If
    
    If Search(iUserID, strXML, "id = " & CStr(ID), , , lErrNum, sErrDesc, sErrSource) Then
        oHelper.Read strXML, lErrNum, sErrDesc, sErrSource
    End If
    
    If Search(iUserID, strXML, "id <> " & CStr(ID) & " AND titulo = '" & oHelper.Titulo & "'", , , lErrNum, sErrDesc, sErrSource) Then
        Set oHelper = New cUbicacion
        If oHelper.Read(iUserID, strXML, lErrNum, sErrDesc, sErrSource) Then
            Err.Raise -1, "cUbicacion.Recover", "El nuevo objeto ya existe en el sistema.<BR>Por favor verifique los datos."
        Else
            If lErrNum = "-2" Then
                Err.Raise -1, "cUbicacion.Read", "Hay mas de un registro para leer en el archivo XML"
            End If
            
            oRs.fields.Append "id", adBSTR
            
            oRs.open
            
            oRs.AddNew
        
            oRs(0) = ID
        
            oRs.Update
            
            oRs.save oDOM, adPersistXML
            
            Recover = oDataAccessTier.RecoverUbicacion(oDOM.xml, lErrNum, sErrDesc, sErrSource)
            
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
    
        ShowError "cUbicacion.Recover", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function
