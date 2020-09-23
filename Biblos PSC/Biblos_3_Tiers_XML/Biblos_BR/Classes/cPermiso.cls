VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPermiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'========================================================================
' Copyright © CHMR 2007 All rights reserved.
'========================================================================
'
'   Module         - cPermiso
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
Private mTablaID As Integer
Private mTabla As String
Private mPermiso As Integer
Private mRolID As Integer

'========================================================================
'Propiedades
'========================================================================
Public Property Get Tabla() As String
    Tabla = mTabla
End Property

Public Property Let Tabla(ByVal vNewValue As String)
    mTabla = vNewValue
End Property

Public Property Get Permiso() As Integer
    Permiso = mPermiso
End Property

Public Property Let Permiso(ByVal vNewValue As Integer)
    mPermiso = vNewValue
End Property

Public Property Get RolID() As Integer
    RolID = mRolID
End Property

Public Property Let RolID(ByVal vNewValue As Integer)
    mRolID = vNewValue
End Property

'========================================================================
'
'   Title          - Función SearchByRol()
'   Description    - Realiza busquedas por Rol, devolviendo un archivo XML con el
'                    resultado.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================
Public Function SearchByRol(sXML As Variant, iRol As Variant, Optional lErrNum As Variant, Optional sErrDesc As Variant, Optional sErrSource As Variant) As Boolean
On Error GoTo SearchByRolErr

Dim oRs As New Recordset
Dim oRsTables As New Recordset
Dim oRsResult As New Recordset
Dim oDOM As New MSXML2.DOMDocument
Dim strXML As String

    SearchByRol = False
    
    If oDataAccessTier.SearchPermisoByRol(strXML, CInt(iRol), lErrNum, sErrDesc, sErrSource) Then
        Set oRs = RecordsetFromXMLString(strXML)
        If Not oRs.EOF Then
            sXML = strXML
            SearchByRol = True
        Else
            If oDataAccessTier.GetTables(strXML, lErrNum, sErrDesc, sErrSource) Then
                Set oRsTables = RecordsetFromXMLString(strXML)
                
                oRsResult.fields.Append "tabla", adBSTR
                oRsResult.fields.Append "permiso", adInteger
                
                oRsResult.Open
                
                While Not oRsTables.EOF
                    oRsResult.AddNew
                
                    oRsResult(0) = oRsTables(0)
                    oRsResult(1) = 0
                    
                    oRsResult.Update
                    oRsTables.MoveNext
                Wend
                
                oRsResult.save oDOM, adPersistXML
                sXML = oDOM.xml
                SearchByRol = True
            End If
        End If
    End If
    
    
    If SearchByRol Then
        Exit Function
    End If
    
SearchByRolErr:
    
    ShowError "cPermiso.SearchByRol", Err.Number, lErrNum, Err.Description, sErrDesc, Err.Source, sErrSource

End Function
