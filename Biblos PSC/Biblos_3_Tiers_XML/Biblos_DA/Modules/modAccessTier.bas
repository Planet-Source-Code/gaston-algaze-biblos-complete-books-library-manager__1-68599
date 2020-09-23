Attribute VB_Name = "modAccessTier"
Option Explicit
Option Private Module

'========================================================================
' Copyright © CHMR 2007 All rights reserved.
'========================================================================
'
'   Module         - modAccessTier
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
'
'   Functions      - Funciones XML String a Recordset
'   Description    - Funciones para pasar de XML a Recordset y viceversa.
'   Version Number - 1.0
'   Last Updated   - January 10th 2007 - 1:25 PM
'   Author         - Algaze, Gastón
'
'========================================================================


Public Function RecordsetFromXMLString(sXML As String) As Recordset

    Dim oStream As ADODB.Stream
    Set oStream = New ADODB.Stream
    
    oStream.Open
    oStream.WriteText sXML   'Give the XML string to the ADO Stream

    oStream.Position = 0    'Set the stream position to the start

    Dim oRecordset As ADODB.Recordset
    Set oRecordset = New ADODB.Recordset
       
    oRecordset.Open oStream    'Open a recordset from the stream

    oStream.Close
    Set oStream = Nothing

    Set RecordsetFromXMLString = oRecordset  'Return the recordset

    Set oRecordset = Nothing

End Function

Public Function RecordsetFromXMLFile(sXMLPath As String) As Recordset

    Dim oRecordset As ADODB.Recordset
    Set oRecordset = New ADODB.Recordset
       
    oRecordset.Open sXMLPath, "Provider=mspersist"

    Set RecordsetFromXMLFile = oRecordset  'Return the recordset

    Set oRecordset = Nothing

End Function

'3 salidas que son las salidas de las funciones que la llaman.
'lErrNumber, sErrDesc, sErrSource
Public Function ShowError(sMessage As String, _
                          lErrNumberLocal As Variant, _
                          lErrNumberEvent As Variant, _
                          sErrDescLocal As Variant, _
                          sErrDescEvent As Variant, _
                          sErrSourceLocal As Variant, _
                          sErrSourceEvent As Variant) As Boolean

    ShowError = False
                          
    If CLng(lErrNumberLocal) <> 0 Then
        lErrNumberEvent = lErrNumberLocal
        sErrDescEvent = sErrDescLocal
        sErrSourceEvent = IIf(Len(sMessage) = 0, sErrSourceLocal, sMessage & "->" & sErrSourceLocal)
        ShowError = True
    Else
        If CLng(lErrNumberEvent) <> 0 Then
            sErrSourceEvent = sMessage & "->" & sErrSourceEvent
            ShowError = True
        End If
    End If
End Function


