Attribute VB_Name = "Module1"


Option Explicit

' PRODUCTION VERSION:
' - Creates one PO for each unique vendor in Excel column B.
' - Includes every Excel row belonging to that vendor in the same PO.
' - Reads Requisitioner from Excel column F.
' - Saves each successfully prepared PO.
' - Writes the resulting PO number to column K for that vendor's rows.

Public Sub SAP_Final_Solution()
    Dim wsInput As Worksheet
    Dim lastRow As Long, i As Long
    Dim sapRow As Long, sapItemNum As String
    Dim itemKey As String
    Dim SapGuiAuto As Object, SAPGUIApp As Object
    Dim Connection As Object, session As Object
    Dim itemSuccess As Boolean, panelExpanded As Boolean
    Dim dropDown As Object, popup As Object
    Dim warnCount As Long
    Dim subIDs As Variant, sID As Variant
    Dim vID As Long

    Dim vendorRows As Object
    Dim vendorKey As Variant
    Dim rowVendor As String
    Dim testVendor As String
    Dim headerRow As Long
    Dim vendorItemIndex As Long
    Dim requisitionerSuccess As Boolean
    Dim itemSelected As Boolean
    Dim poNumber As String
    Dim statusMsg As String
    Dim completedPOs As Long

    subIDs = Array("0010", "0011", "0012", "0013", "0014", _
                   "0015", "0016", "0017", "0018", "0019", _
                   "0020", "0021", "0022")

    On Error GoTo ErrorHandler

    ' ==========================================
    ' 1. DEFINE EXCEL SHEET AND DATA RANGE
    ' ==========================================
    Set wsInput = ThisWorkbook.ActiveSheet
    lastRow = wsInput.Cells(wsInput.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No data found to process.", vbExclamation
        Exit Sub
    End If

    ' ==========================================
    ' 2. BUILD VENDOR LIST AND VALIDATE ALL INPUT
    ' ==========================================
    Set vendorRows = CreateObject("Scripting.Dictionary")
    vendorRows.CompareMode = vbTextCompare

    For i = 2 To lastRow
        rowVendor = Trim$(CStr(wsInput.Cells(i, "B").Value2))

        If Len(rowVendor) = 0 Then
            MsgBox "Vendor is missing in Excel cell B" & i & ".", _
                   vbCritical, "Missing Vendor"
            Exit Sub
        End If

        If Not vendorRows.Exists(rowVendor) Then
            vendorRows.Add rowVendor, i
        End If

        If Len(Trim$(CStr(wsInput.Cells(i, "A").Value2))) = 0 Then
            MsgBox "Material is missing in cell A" & i & ".", _
                   vbCritical, "Missing Material"
            Exit Sub
        End If

        If Len(Trim$(CStr(wsInput.Cells(i, "C").Value2))) = 0 Then
            MsgBox "Plant is missing in cell C" & i & ".", _
                   vbCritical, "Missing Plant"
            Exit Sub
        End If

        If Not IsNumeric(wsInput.Cells(i, "D").Value2) Then
            MsgBox "A valid positive quantity is required in cell D" & i & ".", _
                   vbCritical, "Invalid Quantity"
            Exit Sub
        End If

        If CDbl(wsInput.Cells(i, "D").Value2) <= 0 Then
            MsgBox "A valid positive quantity is required in cell D" & i & ".", _
                   vbCritical, "Invalid Quantity"
            Exit Sub
        End If

        If Not IsDate(wsInput.Cells(i, "E").Value) Then
            MsgBox "A valid delivery date is required in cell E" & i & ".", _
                   vbCritical, "Invalid Delivery Date"
            Exit Sub
        End If

        If Len(Trim$(CStr(wsInput.Cells(i, "F").Value2))) = 0 Then
            MsgBox "Requisitioner is missing in cell F" & i & ".", _
                   vbCritical, "Missing Requisitioner"
            Exit Sub
        End If

        If Len(Trim$(CStr(wsInput.Cells(i, "G").Value2))) = 0 Or _
           Len(Trim$(CStr(wsInput.Cells(i, "H").Value2))) = 0 Then
            MsgBox "Purchasing organization or purchasing group is missing on row " & i & ".", _
                   vbCritical, "Missing Header Data"
            Exit Sub
        End If
    Next i

    ' ==========================================
    ' 3. CONNECT TO ACTIVE SAP SESSION
    ' ==========================================
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    On Error GoTo ErrorHandler

    If SapGuiAuto Is Nothing Then
        MsgBox "SAP is not open. Please log into SAP.", _
               vbCritical, "SAP Error"
        Exit Sub
    End If

    Set SAPGUIApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SAPGUIApp.Children(0)
    Set session = Connection.Children(0)

    If session Is Nothing Then
        MsgBox "Could not find an active SAP window.", _
               vbCritical, "SAP Error"
        Exit Sub
    End If

    completedPOs = 0

    For Each vendorKey In vendorRows.Keys
        testVendor = CStr(vendorKey)
        headerRow = CLng(vendorRows(testVendor))
        poNumber = vbNullString
        statusMsg = vbNullString

    ' ==========================================
    ' 4. START TRANSACTION ME21N
    ' ==========================================
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NME21N"
    session.findById("wnd[0]/tbar[0]/btn[0]").press

    ' ==========================================
    ' 5. HEADER DATA
    ' ==========================================
    For Each sID In subIDs
        Err.Clear
        On Error Resume Next
        session.findById( _
            "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
            "/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/" & _
            "ctxtMEPO_TOPLINE-SUPERFIELD").Text = testVendor

        If Err.Number = 0 Then
            On Error GoTo ErrorHandler
            Exit For
        End If
        On Error GoTo ErrorHandler
    Next sID

    If Err.Number <> 0 Then
        Err.Raise vbObjectError + 1001, "SAP_Final_Solution", _
                  "Could not locate the SAP vendor field."
    End If

    session.findById("wnd[0]").sendVKey 0

    panelExpanded = False

    For Each sID In subIDs
        Err.Clear
        On Error Resume Next
        session.findById( _
            "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
            "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
            "subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9").Select

        If Err.Number = 0 Then
            panelExpanded = True
            On Error GoTo ErrorHandler
            Exit For
        End If
        On Error GoTo ErrorHandler
    Next sID

    If Not panelExpanded Then
        For Each sID In subIDs
            Err.Clear
            On Error Resume Next
            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                "/subSUB1:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4000/" & _
                "btnDYN_4000-BUTTON").press

            If Err.Number = 0 Then
                On Error GoTo ErrorHandler
                Exit For
            End If
            On Error GoTo ErrorHandler
        Next sID
    End If

    panelExpanded = False

    For Each sID In subIDs
        Err.Clear
        On Error Resume Next

        session.findById( _
            "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
            "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
            "subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9").Select

        If Err.Number = 0 Then
            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
                "subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/" & _
                "ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/" & _
                "ctxtMEPO1222-EKORG").Text = _
                    CStr(wsInput.Cells(headerRow, "G").Value2)

            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
                "subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/" & _
                "ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/" & _
                "ctxtMEPO1222-EKGRP").Text = _
                    CStr(wsInput.Cells(headerRow, "H").Value2)

            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
                "subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3").Select

            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
                "subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3/" & _
                "ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230/" & _
                "subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/" & _
                "cntlTEXT_EDITOR_0101/shellcont/shell").Text = _
                    "Please confirm the order Acknowledgement & Order confirmation to orders@example.com" & _
                    vbCr & vbCr & _
                    "Order Acknowledgment : Confirms the receipt of Purchase order" & _
                    vbCr & vbCr & _
                    "Order Confirmation : Confirms the ETA /Delivery date of the Materials along with the Incoterms" & _
                    vbCr & vbCr

            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
                "subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT5").Select

            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
                "subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT5/" & _
                "ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1225/" & _
                "txtMEPOCOMM-IHREZ").Text = "E-ORDER"

            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
                "subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT5/" & _
                "ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1225/" & _
                "txtMEPOCOMM-UNSEZ").Text = "PDP"

            If Err.Number = 0 Then panelExpanded = True
        End If

        On Error GoTo ErrorHandler
        If panelExpanded Then Exit For
    Next sID

    If Not panelExpanded Then
        Err.Raise vbObjectError + 1002, "SAP_Final_Solution", _
                  "Could not complete the SAP header data."
    End If

    session.findById("wnd[0]").sendVKey 0

    ' ==========================================
    ' 6. ENTER GRID DATA FOR SELECTED VENDOR
    ' ==========================================
    sapRow = 0

    For i = 2 To lastRow
        rowVendor = Trim$(CStr(wsInput.Cells(i, "B").Value2))

        If StrComp(rowVendor, testVendor, vbTextCompare) = 0 Then
            itemSuccess = False

            For Each sID In subIDs
                Err.Clear
                On Error Resume Next

                session.findById( _
                    "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                    "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
                    "subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/" & _
                    "ctxtMEPO1211-EMATN[4," & sapRow & "]").Text = _
                        wsInput.Cells(i, "A").Value

                If Err.Number = 0 Then
                    session.findById( _
                        "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                        "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
                        "subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/" & _
                        "ctxtMEPO1211-NAME1[15," & sapRow & "]").Text = _
                            wsInput.Cells(i, "C").Value

                    session.findById( _
                        "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                        "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
                        "subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/" & _
                        "txtMEPO1211-MENGE[6," & sapRow & "]").Text = _
                            wsInput.Cells(i, "D").Value

                    session.findById( _
                        "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                        "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
                        "subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/" & _
                        "ctxtMEPO1211-EEIND[9," & sapRow & "]").Text = _
                            Format$(wsInput.Cells(i, "E").Value, "mm/dd/yyyy")

                    If Err.Number = 0 Then
                        On Error GoTo ErrorHandler

                        requisitionerSuccess = SetSapRequisitioner( _
                            session, CStr(sID), sapRow, _
                            Trim$(CStr(wsInput.Cells(i, "F").Value2)))

                        If requisitionerSuccess Then
                            itemSuccess = True
                            Exit For
                        End If
                    End If
                End If

                On Error GoTo ErrorHandler
            Next sID

            If Not itemSuccess Then
                MsgBox "Could not enter all required fields for Excel row " & i & _
                       "." & vbCrLf & _
                       "Vendor: " & testVendor & vbCrLf & _
                       "Requisitioner: " & _
                       CStr(wsInput.Cells(i, "F").Value2) & vbCrLf & vbCrLf & _
                       "The PO has not been saved.", _
                       vbCritical, "SAP Line Entry Error"
                Exit Sub
            End If

            sapRow = sapRow + 1
        End If
    Next i

    session.findById("wnd[0]").sendVKey 0

    ClearSapWarnings session, 5

    ' ==========================================
    ' 7. EXPAND ITEM DETAILS PANEL
    ' ==========================================
    panelExpanded = False

    For Each sID In subIDs
        For vID = 1 To 4
            Err.Clear
            On Error Resume Next
            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                "/subSUB" & vID & ":SAPLMEVIEWS:1100/" & _
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
                "subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT5").Select

            If Err.Number = 0 Then
                panelExpanded = True
                On Error GoTo ErrorHandler
                Exit For
            End If
            On Error GoTo ErrorHandler
        Next vID

        If panelExpanded Then Exit For
    Next sID

    If Not panelExpanded Then
        For Each sID In subIDs
            For vID = 1 To 4
                Err.Clear
                On Error Resume Next
                session.findById( _
                    "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                    "/subSUB" & vID & ":SAPLMEVIEWS:1100/" & _
                    "subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON").press

                If Err.Number = 0 Then
                    panelExpanded = True
                    On Error GoTo ErrorHandler
                    Exit For
                End If
                On Error GoTo ErrorHandler
            Next vID

            If panelExpanded Then Exit For
        Next sID
    End If

    If Not panelExpanded Then
        Err.Raise vbObjectError + 1003, "SAP_Final_Solution", _
                  "Could not expand the SAP item-details panel."
    End If

    ' ==========================================
    ' 8. ITEM DETAILS FOR SELECTED VENDOR
    ' ==========================================
    ' PASS 1: Remain on Delivery Schedule while processing every item.
    vendorItemIndex = 0

    For i = 2 To lastRow
        rowVendor = Trim$(CStr(wsInput.Cells(i, "B").Value2))

        If StrComp(rowVendor, testVendor, vbTextCompare) = 0 Then
            vendorItemIndex = vendorItemIndex + 1
            sapItemNum = CStr(vendorItemIndex * 10)
            itemKey = Right$("    " & CStr(vendorItemIndex), 4)

            If Not SelectSapItem(session, subIDs, itemKey, sapItemNum) Then
                Err.Raise vbObjectError + 1004, "SAP_Final_Solution", _
                          "Could not select SAP item " & sapItemNum & _
                          " during delivery-schedule processing."
            End If

            If Not SetItemDeliveryTime(session, subIDs, "08:00:00") Then
                Err.Raise vbObjectError + 1005, "SAP_Final_Solution", _
                          "Could not set delivery time for SAP item " & sapItemNum & "."
            End If

            ClearSapWarnings session, 3
        End If
    Next i

    ' PASS 2: Move to Invoice once, then process every item on that tab.
    vendorItemIndex = 0

    For i = 2 To lastRow
        rowVendor = Trim$(CStr(wsInput.Cells(i, "B").Value2))

        If StrComp(rowVendor, testVendor, vbTextCompare) = 0 Then
            vendorItemIndex = vendorItemIndex + 1
            sapItemNum = CStr(vendorItemIndex * 10)
            itemKey = Right$("    " & CStr(vendorItemIndex), 4)

            If Not SelectSapItem(session, subIDs, itemKey, sapItemNum) Then
                Err.Raise vbObjectError + 1008, "SAP_Final_Solution", _
                          "Could not select SAP item " & sapItemNum & _
                          " during invoice processing."
            End If

            If Not SetInvoiceCheckboxes(session, subIDs) Then
                Err.Raise vbObjectError + 1006, "SAP_Final_Solution", _
                          "Could not set invoice checkboxes for SAP item " & sapItemNum & "."
            End If
        End If
    Next i

    ' PASS 3: Move to Confirmations once, then process every item on that tab.
    vendorItemIndex = 0

    For i = 2 To lastRow
        rowVendor = Trim$(CStr(wsInput.Cells(i, "B").Value2))

        If StrComp(rowVendor, testVendor, vbTextCompare) = 0 Then
            vendorItemIndex = vendorItemIndex + 1
            sapItemNum = CStr(vendorItemIndex * 10)
            itemKey = Right$("    " & CStr(vendorItemIndex), 4)

            If Not SelectSapItem(session, subIDs, itemKey, sapItemNum) Then
                Err.Raise vbObjectError + 1009, "SAP_Final_Solution", _
                          "Could not select SAP item " & sapItemNum & _
                          " during confirmation processing."
            End If

            itemSuccess = SetConfirmationCheckbox(session, subIDs)

            If Not itemSuccess Then
                MsgBox "Could not update details for SAP Item " & sapItemNum & _
                       " for vendor " & testVendor & "." & vbCrLf & _
                       "The PO has not been saved.", _
                       vbCritical, "Item Details Error"
                Exit Sub
            End If
        End If
    Next i

    ' ==========================================
    ' 9. SAVE CURRENT VENDOR PO
    ' ==========================================
    If Not SavePOAndGetNumber(session, poNumber, statusMsg) Then
        MsgBox "SAP did not confirm PO creation for vendor " & testVendor & "." & _
               vbCrLf & vbCrLf & _
               "SAP status: " & statusMsg & vbCrLf & vbCrLf & _
               "Processing has stopped. Any earlier vendor POs remain saved.", _
               vbCritical, "PO Save Not Confirmed"
        GoTo SafeExit
    End If

    For i = 2 To lastRow
        rowVendor = Trim$(CStr(wsInput.Cells(i, "B").Value2))

        If StrComp(rowVendor, testVendor, vbTextCompare) = 0 Then
            wsInput.Cells(i, "K").NumberFormat = "@"
            wsInput.Cells(i, "K").Value = poNumber
        End If
    Next i

    completedPOs = completedPOs + 1
    Next vendorKey

    MsgBox completedPOs & " PO(s) created successfully." & vbCrLf & _
           "The PO numbers were written to Excel column K.", _
           vbInformation, "SAP Automation Complete"

SafeExit:
    Set popup = Nothing
    Set dropDown = Nothing
    Set session = Nothing
    Set Connection = Nothing
    Set SAPGUIApp = Nothing
    Set SapGuiAuto = Nothing
    Set vendorRows = Nothing
    Set wsInput = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Procedure: SAP_Final_Solution" & vbCrLf & _
           "Excel row: " & i & vbCrLf & _
           "Vendor: " & testVendor & vbCrLf & _
           "SAP item: " & sapItemNum & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Completed POs before this error: " & completedPOs & vbCrLf & _
           "The current vendor PO was not confirmed as saved.", _
           vbCritical, "SAP Automation Error"
    Resume SafeExit
End Sub

Private Function SavePOAndGetNumber( _
    ByVal session As Object, _
    ByRef poNumber As String, _
    ByRef statusMsg As String) As Boolean

    Dim popup As Object
    Dim messageType As String
    Dim words As Variant
    Dim word As Variant
    Dim cleanedWord As String

    SavePOAndGetNumber = False
    poNumber = vbNullString
    statusMsg = vbNullString

    If Not WaitForSapReady(session, 30) Then
        statusMsg = "SAP was still busy before the save command."
        Exit Function
    End If

    session.findById("wnd[0]").sendVKey 11

    If Not WaitForSapReady(session, 60) Then
        statusMsg = "SAP did not complete the save within 60 seconds."
        Exit Function
    End If

    Set popup = Nothing
    Err.Clear
    On Error Resume Next
    Set popup = session.findById("wnd[1]", False)
    On Error GoTo 0

    If Not popup Is Nothing Then
        On Error Resume Next
        statusMsg = popup.Text
        On Error GoTo 0

        If Len(statusMsg) = 0 Then
            statusMsg = "SAP displayed a secondary window during save."
        End If
        Exit Function
    End If

    Err.Clear
    On Error Resume Next
    statusMsg = session.findById("wnd[0]/sbar").Text
    messageType = session.findById("wnd[0]/sbar").messageType

    If Err.Number <> 0 Then
        On Error GoTo 0
        statusMsg = "Could not read the SAP status bar after save."
        Exit Function
    End If
    On Error GoTo 0

    If messageType <> "S" Then Exit Function

    words = Split(statusMsg, " ")

    For Each word In words
        cleanedWord = Replace(CStr(word), ".", vbNullString)
        cleanedWord = Replace(cleanedWord, ",", vbNullString)
        cleanedWord = Replace(cleanedWord, ":", vbNullString)
        cleanedWord = Replace(cleanedWord, ";", vbNullString)

        If IsNumeric(cleanedWord) Then
            If Len(cleanedWord) = 10 Then
                poNumber = cleanedWord
                Exit For
            End If
        End If
    Next word

    If Len(poNumber) = 0 Then Exit Function

    SavePOAndGetNumber = True
End Function

Private Function SetSapRequisitioner( _
    ByVal session As Object, _
    ByVal subId As String, _
    ByVal sapRow As Long, _
    ByVal requisitioner As String) As Boolean

    Dim basePath As String
    Dim controlPath As String
    Dim sapControl As Object

    SetSapRequisitioner = False
    If Len(requisitioner) = 0 Then Exit Function

    basePath = _
        "wnd[0]/usr/subSUB0:SAPLMEGUI:" & subId & _
        "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
        "subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/"

    ' SAP GUI recorder confirmed:
    ' txtMEPO1211-AFNAM[21,0]
    ' Keep the recorded column index and substitute the current SAP row.
    controlPath = basePath & _
                  "txtMEPO1211-AFNAM[21," & sapRow & "]"

    Set sapControl = Nothing
    Err.Clear
    On Error Resume Next
    Set sapControl = session.findById(controlPath, False)

    If Not sapControl Is Nothing Then
        sapControl.Text = requisitioner

        If Err.Number = 0 Then
            sapControl.SetFocus
            sapControl.caretPosition = Len(requisitioner)

            If Err.Number = 0 Then
                SetSapRequisitioner = True
            End If
        End If
    End If

    On Error GoTo 0
End Function

Private Function WaitForSapReady( _
    ByVal session As Object, _
    ByVal timeoutSeconds As Double) As Boolean

    Dim startedAt As Double
    Dim elapsed As Double

    WaitForSapReady = False
    startedAt = Timer

    Do
        Err.Clear
        On Error Resume Next

        If Not session.Busy Then
            If Err.Number = 0 Then
                On Error GoTo 0
                WaitForSapReady = True
                Exit Function
            End If
        End If

        On Error GoTo 0
        DoEvents

        elapsed = Timer - startedAt
        If elapsed < 0 Then elapsed = elapsed + 86400#
        If elapsed >= timeoutSeconds Then Exit Do
    Loop
End Function

Private Function SelectSapItem( _
    ByVal session As Object, _
    ByVal subIDs As Variant, _
    ByVal itemKey As String, _
    ByVal sapItemNum As String) As Boolean

    Dim sID As Variant
    Dim vID As Long
    Dim dropDown As Object

    SelectSapItem = False

    For Each sID In subIDs
        For vID = 1 To 4
            Set dropDown = Nothing
            Err.Clear
            On Error Resume Next

            Set dropDown = session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                "/subSUB" & vID & ":SAPLMEVIEWS:1100/" & _
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
                "subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST", False)

            If Not dropDown Is Nothing Then
                dropDown.SetFocus
                dropDown.Key = itemKey

                If Err.Number = 0 Then
                    session.findById("wnd[0]").sendVKey 0

                    If Err.Number = 0 Then
                        On Error GoTo 0

                        If WaitForSapReady(session, 20) Then
                            SelectSapItem = True
                            Exit Function
                        End If
                    End If
                End If
            End If

            On Error GoTo 0
        Next vID
    Next sID
End Function

Private Sub ClearSapWarnings(ByVal session As Object, ByVal maximumAttempts As Long)
    Dim attempt As Long
    Dim popup As Object
    Dim messageType As String

    For attempt = 1 To maximumAttempts
        Set popup = Nothing
        Err.Clear
        On Error Resume Next
        Set popup = session.findById("wnd[1]", False)
        On Error GoTo 0

        If Not popup Is Nothing Then
            session.findById("wnd[1]").sendVKey 0
        Else
            messageType = vbNullString

            On Error Resume Next
            messageType = session.findById("wnd[0]/sbar").messageType
            On Error GoTo 0

            If messageType = "W" Or messageType = "I" Then
                session.findById("wnd[0]").sendVKey 0
            Else
                Exit For
            End If
        End If
    Next attempt
End Sub

Private Function SetItemDeliveryTime( _
    ByVal session As Object, _
    ByVal subIDs As Variant, _
    ByVal deliveryTime As String) As Boolean

    Dim sID As Variant
    Dim vID As Long
    Dim tabPath As String
    Dim timeFieldPath As String
    Dim timeField As Object
    Dim startedAt As Double
    Dim elapsed As Double

    SetItemDeliveryTime = False

    For Each sID In subIDs
        For vID = 1 To 4
            tabPath = _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                "/subSUB" & vID & ":SAPLMEVIEWS:1100/" & _
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
                "subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT5"

            timeFieldPath = tabPath & _
                "/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1320/" & _
                "tblSAPLMEGUITC_1320/ctxtMEPO1320-UZEIT[4,0]"

            Err.Clear
            On Error Resume Next

            session.findById(tabPath).Select

            If Err.Number = 0 Then
                On Error GoTo 0

                If WaitForSapReady(session, 20) Then
                    startedAt = Timer

                    Do
                        Set timeField = Nothing
                        Err.Clear
                        On Error Resume Next
                        Set timeField = session.findById(timeFieldPath, False)
                        On Error GoTo 0

                        If Not timeField Is Nothing Then
                            Err.Clear
                            On Error Resume Next
                            timeField.Text = deliveryTime

                            If Err.Number = 0 Then
                                timeField.SetFocus
                                timeField.caretPosition = Len(deliveryTime)
                                session.findById("wnd[0]").sendVKey 0
                                On Error GoTo 0

                                If WaitForSapReady(session, 20) Then
                                    SetItemDeliveryTime = True
                                    Exit Function
                                End If
                            End If

                            On Error GoTo 0
                        End If

                        DoEvents
                        elapsed = Timer - startedAt
                        If elapsed < 0 Then elapsed = elapsed + 86400#
                        If elapsed >= 10 Then Exit Do
                    Loop
                End If
            End If

            On Error GoTo 0
        Next vID
    Next sID
End Function

Private Function SetInvoiceCheckboxes( _
    ByVal session As Object, _
    ByVal subIDs As Variant) As Boolean

    Dim sID As Variant
    Dim vID As Long

    SetInvoiceCheckboxes = False

    For Each sID In subIDs
        For vID = 1 To 4
            Err.Clear
            On Error Resume Next

            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                "/subSUB" & vID & ":SAPLMEVIEWS:1100/" & _
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
                "subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7").Select

            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                "/subSUB" & vID & ":SAPLMEVIEWS:1100/" & _
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
                "subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/" & _
                "ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/" & _
                "chkMEPO1317-REPOS").Selected = True

            If Err.Number = 0 Then session.findById("wnd[0]").sendVKey 2

            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                "/subSUB" & vID & ":SAPLMEVIEWS:1100/" & _
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
                "subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/" & _
                "ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/" & _
                "chkMEPO1317-WEBRE").Selected = True

            If Err.Number = 0 Then
                session.findById("wnd[0]").sendVKey 2
                SetInvoiceCheckboxes = True
                On Error GoTo 0
                Exit Function
            End If

            On Error GoTo 0
        Next vID
    Next sID
End Function

Private Function SetConfirmationCheckbox( _
    ByVal session As Object, _
    ByVal subIDs As Variant) As Boolean

    Dim sID As Variant
    Dim vID As Long

    SetConfirmationCheckbox = False

    For Each sID In subIDs
        For vID = 1 To 4
            Err.Clear
            On Error Resume Next

            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                "/subSUB" & vID & ":SAPLMEVIEWS:1100/" & _
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
                "subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT17").Select

            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:" & sID & _
                "/subSUB" & vID & ":SAPLMEVIEWS:1100/" & _
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
                "subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT17/" & _
                "ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/" & _
                "subSUB1:SAPLMEGUI:1334/chkMEPO1334-KZABS").Selected = True

            If Err.Number = 0 Then
                session.findById("wnd[0]").sendVKey 2
                SetConfirmationCheckbox = True
                On Error GoTo 0
                Exit Function
            End If

            On Error GoTo 0
        Next vID
    Next sID
End Function



