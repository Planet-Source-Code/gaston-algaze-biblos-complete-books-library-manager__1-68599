VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSecurityAgent"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'========================================================================
' Copyright © CHMR 2007 All rights reserved.
'========================================================================
'
'   Module         - cSecurityAgent
'   Version Number - 1.0
'   Last Updated   - March 6th 2007 - 5:34 PM
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
'Propiedades
'========================================================================

'========================================================================
'Métodos y Funciones
'========================================================================
'========================================================================
'
'   Title          - Función Login()
'   Description    - Login.
'   Version Number - 1.0
'   Last Updated   - March 6th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function Login(sXML As Variant, sUsername As Variant, sPassword As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo LoginErr
Dim strXML As String

    Login = oDataAccessTier.LoginUsuario(strXML, CStr(sUsername), CStr(sPassword), lErrNum, sErrDesc, sErrSource)
    
    If Login Then
        sXML = strXML
        Exit Function
    Else
    
LoginErr:
    ShowError "cSecurityAgent.Login", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
    
End Function

'========================================================================
'   Función CompareColumns()
'   Description    - Valida que los nombres de los campos de una tabla correspodan con los valores de entrada
'   Version Number - 1.0
'   Last Updated   - 21.03.2007
'   Author         - Algaze, Gastón
'
'   21.03.2007 - Creacion de la funcion.
'========================================================================
Public Function CompareColumns(sTable As String, sValue As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo CompareColumnsErr

Dim strAux() As String
Dim strValues() As String
Dim strColumnsXML As String
Dim oRs As New Recordset

Dim i As Integer
Dim j As Integer

i = 0
strValues() = Split(sValue, ",")

CompareColumns = False

If oDataAccessTier.GetColumns(sTable, strColumnsXML, lErrNum, sErrDesc, sErrSource) Then
    Set oRs = RecordsetFromXMLString(strColumnsXML)
    
    If Not oRs.EOF Then
        ReDim strAux(oRs.RecordCount)
    Else
        Err.Raise -1, "cSecurityAgent.CompareColumns", "No se pudieron recuperar los campos"
    End If
    
    While Not oRs.EOF
        strAux(i) = oRs("field")
        i = i + 1
        oRs.MoveNext
    Wend
    
    For i = 0 To UBound(strAux)
        For j = 0 To UBound(strValues)
            If StrComp(strValues(j), strAux(i), vbTextCompare) = 0 Then
                CompareColumns = True
                Exit For
            End If
        Next
        If CompareColumns Then Exit For
    Next
Else
    Err.Raise -1, "cSecurityAgent.CompareColumns", "No se pudieron recuperar los campos"
CompareColumnsErr:
        
    ShowError "cSecurityAgent.CompareColumns", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
End If
End Function

'========================================================================
'   Función ValidateOrderBy()
'   Description    - Valida que los nombres de los campos para ordenar sean validos
'   Version Number - 1.0
'   Last Updated   - 21.03.2007
'   Author         - Algaze, Gastón
'
'   21.03.2007 - Creacion de la funcion.
'========================================================================
Public Function ValidateOrderBy(sValue As String, sClass As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
    
    ValidateOrderBy = False

    Select Case sClass
        Case "cLink"
            If CompareColumns("links", sValue, lErrNum, sErrDesc, sErrSource) Then
                ValidateOrderBy = True
            Else
                Err.Raise -1, "cSecurityAgent.ValidateOrderBy", "No se pudieron recuperar los campos"
            End If
        Case "cFicha"
            If CompareColumns("fichas", sValue, lErrNum, sErrDesc, sErrSource) Then
                ValidateOrderBy = True
            Else
                Err.Raise -1, "cSecurityAgent.ValidateOrderBy", "No se pudieron recuperar los campos"
            End If
        Case "cItemTipo"
            If CompareColumns("items_tipos", sValue, lErrNum, sErrDesc, sErrSource) Then
                ValidateOrderBy = True
            Else
                Err.Raise -1, "cSecurityAgent.ValidateOrderBy", "No se pudieron recuperar los campos"
            End If
        Case "cReserva"
            If CompareColumns("reservas", sValue, lErrNum, sErrDesc, sErrSource) Then
                ValidateOrderBy = True
            Else
                Err.Raise -1, "cSecurityAgent.ValidateOrderBy", "No se pudieron recuperar los campos"
            End If
        Case "cSubCategoria"
            If CompareColumns("subcategorias", sValue, lErrNum, sErrDesc, sErrSource) Then
                ValidateOrderBy = True
            Else
                Err.Raise -1, "cSecurityAgent.ValidateOrderBy", "No se pudieron recuperar los campos"
            End If
        Case "cRestriccion"
            If CompareColumns("restricciones", sValue, lErrNum, sErrDesc, sErrSource) Then
                ValidateOrderBy = True
            Else
                Err.Raise -1, "cSecurityAgent.ValidateOrderBy", "No se pudieron recuperar los campos"
            End If
        Case "cPrestamo"
            If CompareColumns("prestamos", sValue, lErrNum, sErrDesc, sErrSource) Then
                ValidateOrderBy = True
            Else
                Err.Raise -1, "cSecurityAgent.ValidateOrderBy", "No se pudieron recuperar los campos"
            End If
        Case "cPermiso"
            If CompareColumns("permisos", sValue, lErrNum, sErrDesc, sErrSource) Then
                ValidateOrderBy = True
            Else
                Err.Raise -1, "cSecurityAgent.ValidateOrderBy", "No se pudieron recuperar los campos"
            End If
        Case "cOperacion"
            If CompareColumns("operaciones", sValue, lErrNum, sErrDesc, sErrSource) Then
                ValidateOrderBy = True
            Else
                Err.Raise -1, "cSecurityAgent.ValidateOrderBy", "No se pudieron recuperar los campos"
            End If
        Case "cEditorial"
            If CompareColumns("editoriales", sValue, lErrNum, sErrDesc, sErrSource) Then
                ValidateOrderBy = True
            Else
                Err.Raise -1, "cSecurityAgent.ValidateOrderBy", "No se pudieron recuperar los campos"
            End If
        Case "cCategoria"
            If CompareColumns("categorias", sValue, lErrNum, sErrDesc, sErrSource) Then
                ValidateOrderBy = True
            Else
                Err.Raise -1, "cSecurityAgent.ValidateOrderBy", "No se pudieron recuperar los campos"
            End If
        Case "cCampo"
            If CompareColumns("campos", sValue, lErrNum, sErrDesc, sErrSource) Then
                ValidateOrderBy = True
            Else
                Err.Raise -1, "cSecurityAgent.ValidateOrderBy", "No se pudieron recuperar los campos"
            End If
        Case "cUbicacion"
            If CompareColumns("ubicaciones", sValue, lErrNum, sErrDesc, sErrSource) Then
                ValidateOrderBy = True
            Else
                Err.Raise -1, "cSecurityAgent.ValidateOrderBy", "No se pudieron recuperar los campos"
            End If
        Case "cPrestamoTipo"
            If CompareColumns("prestamos_tipos", sValue, lErrNum, sErrDesc, sErrSource) Then
                ValidateOrderBy = True
            Else
                Err.Raise -1, "cSecurityAgent.ValidateOrderBy", "No se pudieron recuperar los campos"
            End If
        Case "cUsuario"
            If CompareColumns("usuarios", sValue, lErrNum, sErrDesc, sErrSource) Then
                ValidateOrderBy = True
            Else
                Err.Raise -1, "cSecurityAgent.ValidateOrderBy", "No se pudieron recuperar los campos"
            End If
        Case "cRol"
            If CompareColumns("roles", sValue, lErrNum, sErrDesc, sErrSource) Then
                ValidateOrderBy = True
            Else
                Err.Raise -1, "cSecurityAgent.ValidateOrderBy", "No se pudieron recuperar los campos"
            End If
        Case "cItem"
            If CompareColumns("Items", sValue, lErrNum, sErrDesc, sErrSource) Then
                ValidateOrderBy = True
            Else
                Err.Raise -1, "cSecurityAgent.ValidateOrderBy", "No se pudieron recuperar los campos"
            End If
        Case Else
            Err.Raise -1, "cSecurityAgent.ValidateOrderBy", "Clase no reconocida"
    End Select

End Function

Public Function ValidateSort(sSort As String) As Boolean
    ValidateSort = IIf(UCase(sSort) = "ASC" Or UCase(sSort) = "DESC", True, False)
End Function

'========================================================================
'
'   Title          - Función IsAllowed()
'   Description    - Valida si el usuario que ejecuta la operacion tiene los permisos necesarios para realizarla.
'   Version Number - 1.0
'   Last Updated   - February 25th 2007 - 3:37 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function IsAllowed(iUserID As Integer, sClass As String, iOperation As Integer, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo IsAllowedErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strTabla As String
Dim strPrivilegesXML As String

    IsAllowed = False
    
    Select Case sClass
        Case "cLink"
            strTabla = "links"
        Case "cFicha"
            strTabla = "fichas"
        Case "cItemTipo"
            strTabla = "items_tipos"
        Case "cReserva"
            strTabla = "reservas"
        Case "cPrestamoTipo"
            strTabla = "prestamos_tipos"
        Case "cPrestamo"
            strTabla = "prestamos"
        Case "cRestriccion"
            strTabla = "restricciones"
        Case "cItem"
            strTabla = "items"
        Case "cEditorial"
            strTabla = "editoriales"
        Case "cCampo"
            strTabla = "campos"
        Case "cCategoria"
            strTabla = "categorias"
        Case "cSubCategoria"
            strTabla = "subcategorias"
        Case "cOperacion"
            strTabla = "operaciones"
        Case "cPermiso"
            strTabla = "permisos"
        Case "cRol"
            strTabla = "roles"
        Case "cUsuario"
            strTabla = "usuarios"
        Case "cUbicacion"
            strTabla = "ubicaciones"
        Case Else
            Err.Raise -1, "cSecurityAgent.IsAllowed", "Nombre de clase desconocido."
    End Select
    
    Select Case iOperation
        Case 1 ' Alta
        Case 2 ' Baja
        Case 3 ' Modificacion
        Case 4 ' Lectura
        Case Else
            Err.Raise -1, "cSecurityAgent.IsAllowed", "Operacion no reconocida."
    End Select
    
    oRs.fields.Append "userID", adBSTR
    oRs.fields.Append "Tabla", adBSTR
    
    oRs.open
    
    oRs.AddNew

    oRs(0) = iUserID
    oRs("Tabla") = strTabla

    oRs.Update
    
    oRs.save oDOM, adPersistXML
    
     If oDataAccessTier.GetGlobalPrivileges(oDOM.xml, strPrivilegesXML, lErrNum, sErrDesc, sErrSource) Then
        Set oRs = Nothing
        Set oRs = RecordsetFromXMLString(strPrivilegesXML)
        If Not oRs.EOF Then
            Select Case iOperation
                Case 1 ' Alta [8 - 15]
                    IsAllowed = Int(oRs("permiso")) >= 8
                Case 2 ' Baja [4 - 7], [12 - 15]
                    IsAllowed = (Int(oRs("permiso")) >= 4 And Int(oRs("permiso")) <= 7) _
                             Or (Int(oRs("permiso")) >= 12 And Int(oRs("permiso")) <= 15)
                Case 3 ' Modificacion [2-3], [6-7], [10-11], [14-15]
                    IsAllowed = (Int(oRs("permiso")) >= 2 And Int(oRs("permiso")) <= 3) _
                             Or (Int(oRs("permiso")) >= 6 And Int(oRs("permiso")) <= 6) _
                             Or (Int(oRs("permiso")) >= 10 And Int(oRs("permiso")) <= 11) _
                             Or (Int(oRs("permiso")) >= 14 And Int(oRs("permiso")) <= 15)
                Case 4 ' Lectura 'los impares
                    IsAllowed = IIf(CLng(oRs("permiso")) Mod 2 = 0, False, True)
            End Select
        End If
    End If
    
    If Not IsAllowed Then
        If oDataAccessTier.GetPrivileges(oDOM.xml, strPrivilegesXML, lErrNum, sErrDesc, sErrSource) Then
            Set oRs = Nothing
            Set oRs = RecordsetFromXMLString(strPrivilegesXML)
            If Not oRs.EOF Then
                Select Case iOperation
                    Case 1 ' Alta [8 - 15]
                        IsAllowed = Int(oRs("permiso")) >= 8
                    Case 2 ' Baja [4 - 7], [12 - 15]
                        IsAllowed = (Int(oRs("permiso")) >= 4 And Int(oRs("permiso")) <= 7) _
                                 Or (Int(oRs("permiso")) >= 12 And Int(oRs("permiso")) <= 15)
                    Case 3 ' Modificacion [2-3], [6-7], [10-11], [14-15]
                        IsAllowed = (Int(oRs("permiso")) >= 2 And Int(oRs("permiso")) <= 3) _
                                 Or (Int(oRs("permiso")) >= 6 And Int(oRs("permiso")) <= 6) _
                                 Or (Int(oRs("permiso")) >= 10 And Int(oRs("permiso")) <= 11) _
                                 Or (Int(oRs("permiso")) >= 14 And Int(oRs("permiso")) <= 15)
                    Case 4 ' Lectura 'los impares
                        IsAllowed = IIf(CLng(oRs("permiso")) Mod 2 = 0, False, True)
                End Select
            End If
        End If
    End If
    
    Set oRs = Nothing
    Set oDOM = Nothing
    
    If IsAllowed Then
        Exit Function
    Else
        Err.Raise -1, "cSecurityAgent.IsAllowed", "El usuario no tiene privilegios suficientes para realizar la acción."
        
IsAllowedErr:
        
        ShowError "cSecurityAgent.IsAllowed", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource

    End If
    
End Function

'========================================================================
'
'   Title          - Función GetRestrictions()
'   Description    - GetRestrictions.
'   Version Number - 1.0
'   Last Updated   - March 6th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function GetRestrictions(iUserID As Integer, sClass As String, sRestricciones As String, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo GetRestrictionsErr

Dim oRs As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strTabla As String
Dim strRestrictionsXML As String


    GetRestrictions = False
    
    Select Case sClass
        Case "All"
            strTabla = "All"
        Case "cLink"
            strTabla = "links"
        Case "cFicha"
            strTabla = "fichas"
        Case "cItemTipo"
            strTabla = "items_tipos"
        Case "cReserva"
            strTabla = "reservas"
        Case "cPrestamoTipo"
            strTabla = "prestamos_tipos"
        Case "cPrestamo"
            strTabla = "prestamos"
        Case "cRestriccion"
            strTabla = "restricciones"
        Case "cItem"
            strTabla = "Items"
        Case "cEditorial"
            strTabla = "editoriales"
        Case "cCampo"
            strTabla = "campos"
        Case "cCategoria"
            strTabla = "categorias"
        Case "cSubCategoria"
            strTabla = "subcategorias"
        Case "cOperacion"
            strTabla = "operaciones"
        Case "cPermiso"
            strTabla = "permisos"
        Case "cRol"
            strTabla = "roles"
        Case "cUsuario"
            strTabla = "usuarios"
        Case "cUbicacion"
            strTabla = "ubicaciones"
        Case Else
            Err.Raise -1, "Biblos_BR.GetRestrictions", "Nombre de clase desconocido."
    End Select
    
    oRs.fields.Append "userID", adBSTR
    oRs.fields.Append "Tabla", adBSTR
    
    oRs.open
    
    oRs.AddNew

    oRs(0) = iUserID
    oRs("Tabla") = strTabla

    oRs.Update
    
    oRs.save oDOM, adPersistXML
    
    If oDataAccessTier.GetRestrictions(oDOM.xml, strRestrictionsXML, lErrNum, sErrDesc, sErrSource) Then
        Set oRs = Nothing
        Set oRs = RecordsetFromXMLString(strRestrictionsXML)
        While Not oRs.EOF
            sRestricciones = oRs("tabla") & "." & oRs("campo") & " " & oRs("operacion")
            If StrComp(oRs("operacion"), "LIKE", vbTextCompare) = 0 Then
                If IsDate(oRs("valor")) Or IsNumeric(oRs("valor")) Then  'si es numerico o fecha
                    sRestricciones = sRestricciones & " %" & oRs("valor") & "% "
                Else
                    sRestricciones = sRestricciones & " '%" & oRs("valor") & "%' "
                End If
            Else
                If IsDate(oRs("valor")) Or IsNumeric(oRs("valor")) Then  'si es numerico o fecha
                    sRestricciones = sRestricciones & " " & oRs("valor") & " "
                Else
                    sRestricciones = sRestricciones & " '" & oRs("valor") & "' "
                End If
            End If
            oRs.MoveNext
            If Not oRs.EOF Then sRestricciones = sRestricciones & "AND "
        Wend
        GetRestrictions = True
    End If
    
    Set oRs = Nothing
    Set oDOM = Nothing
    
    If GetRestrictions Then
        Exit Function
    Else
    
GetRestrictionsErr:
        
    ShowError "Biblos_BR.GetRestrictions", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource
    End If
End Function

