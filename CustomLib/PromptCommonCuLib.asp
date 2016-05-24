<%'** Copyright © 1996-2012 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Function CO_GetAnyPromptError(oPromptsTempXML, bAnyError)
'**************************************************************************
'Purpose:   put sErrCode in oSinglePromptTempXML/<temp>
'Inputs:    oSinglePromptTempXML, sErrCode
'Outputs:   Err.number
'**************************************************************************

    On Error Resume Next
    Dim oSinglePromptTempXML
    Dim sErrCode

    bAnyError = False
    For Each oSinglePromptTempXML In oPromptsTempXML.selectNodes("/mi/pif")
        Call CO_GetPromptError(oSinglePromptTempXML, sErrCode)
        If CLng(sErrCode) > 0 Then
            bAnyError = True
            Exit Function
        End If
    Next

    CO_GetAnyPromptError = Err.Number
    Err.Clear
End Function

Function CO_SetPromptError(oSinglePromptTempXML, sErrCode)
'**************************************************************************
'Purpose:   put sErrCode in oSinglePromptTempXML/<temp> to be displayed
'Inputs:    oSinglePromptTempXML, sErrCode
'Outputs:   Err.number
'**************************************************************************

    On Error Resume Next
    Dim oRootAnswer
    Dim oTemp
    Dim oError

    Set oRootAnswer = oSinglePromptTempXML.selectSingleNode("/")

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If (oTemp Is Nothing) Then
        Set oTemp = oRootAnswer.createElement("temp")
        Call oSinglePromptTempXML.appendChild(oTemp)
    End If

    Set oError = oTemp.selectSingleNode("./error")
    If (oError Is Nothing) Then
        Set oError = oRootAnswer.createElement("error")
        Call oTemp.appendChild(oError)
    End If

    Call oError.setAttribute("code", CStr(sErrCode))

    Set oRootAnswer = Nothing
    Set oTemp = Nothing
    Set oError = Nothing

    CO_SetPromptError = Err.Number
    Err.Clear
End Function

Function CO_GetPromptError(oSinglePromptTempXML, sErrCode)
'**************************************************************************
'Purpose:   get sErrCode in oSinglePromptTempXML/<temp>
'Inputs:    oSinglePromptTempXML
'Outputs:   sErrCode
'**************************************************************************

    On Error Resume Next
    Dim oTemp
    Dim oError

    sErrCode = "0"
    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oError = oTemp.selectSingleNode("./error")
        If Not (oError Is Nothing) Then
            sErrCode = oError.getAttribute("code")
        End If
    End If

    Set oTemp = Nothing
    Set oError = Nothing

    CO_GetPromptError = Err.Number
    Err.Clear
End Function

Function CO_SetBlockBegin(oSinglePromptTempXML, sBlockBegin)
'**************************************************************************
'Purpose:   put sBlockBegin in oSinglePromptTempXML/<temp>
'Inputs:    oSinglePromptTempXML, sBlockBegin
'Outputs:   Err.number
'**************************************************************************

    On Error Resume Next
    Dim oRootAnswer
    Dim oTemp
    Dim oBB

    Set oRootAnswer = oSinglePromptTempXML.selectSingleNode("/")

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If (oTemp Is Nothing) Then
        Set oTemp = oRootAnswer.createElement("temp")
        Call oSinglePromptTempXML.appendChild(oTemp)
    End If

    Set oBB = oTemp.selectSingleNode("./blockbegin")
    If (oBB Is Nothing) Then
        Set oBB = oRootAnswer.createElement("blockbegin")
        Call oTemp.appendChild(oBB)
    End If

    Call oBB.setAttribute("value", CStr(sBlockBegin))

    Set oRootAnswer = Nothing
    Set oTemp = Nothing
    Set oBB = Nothing

    CO_SetBlockBegin = Err.Number
    Err.Clear
End Function

Function CO_GetBlockBegin(oSinglePromptTempXML, lBlockBegin)
'**************************************************************************
'Purpose:   get BlockBegin in oSinglePromptTempXML/<temp>
'Inputs:    oSinglePromptTempXML
'Outputs:   lBlockBegin
'**************************************************************************
    On Error Resume Next
    Dim oTemp
    Dim oBB
    Dim sBlockBegin

    lBlockBegin = 1
    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oBB = oTemp.selectSingleNode("./blockbegin")
        If Not (oBB Is Nothing) Then
            sBlockBegin = oBB.getAttribute("value")
            If Len(sBlockBegin) > 0 Then
                lBlockBegin = CLng(sBlockBegin)
            End If
        End If
    End If

    Set oTemp = Nothing
    Set oBB = Nothing

    CO_GetBlockBegin = Err.Number
    Err.Clear
End Function


Function CO_SetSearchField(oSinglePromptTempXML, sSearch)
'**************************************************************************
'Purpose:   put sSearch in oSinglePromptTempXML/<temp>
'Inputs:    oSinglePromptTempXML, sSearch
'Outputs:   Err.number
'**************************************************************************

    On Error Resume Next
    Dim oRootAnswer
    Dim oTemp
    Dim oSearch

    Set oRootAnswer = oSinglePromptTempXML.selectSingleNode("/")

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If (oTemp Is Nothing) Then
        Set oTemp = oRootAnswer.createElement("temp")
        Call oSinglePromptTempXML.appendChild(oTemp)
    End If

    Set oSearch = oTemp.selectSingleNode("./search")
    If (oSearch Is Nothing) Then
        Set oSearch = oRootAnswer.createElement("search")
        Call oTemp.appendChild(oSearch)
    End If

    Call oSearch.setAttribute("text", CStr(sSearch))

    Set oRootAnswer = Nothing
    Set oTemp = Nothing
    Set oSearch = Nothing

    CO_SetSearchField = Err.Number
    Err.Clear
End Function

Function CO_GetSearchField(oSinglePromptTempXML, sSearch)
'**************************************************************************
'Purpose:   get sSearch in oSinglePromptTempXML/<temp>
'Inputs:    oSinglePromptTempXML
'Outputs:   sSearch
'**************************************************************************

    On Error Resume Next
    Dim oTemp
    Dim oSearch

    sSearch = ""
    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oSearch = oTemp.selectSingleNode("./search")
        If Not (oSearch Is Nothing) Then
            sSearch = oSearch.getAttribute("text")
        End If
    End If

    Set oTemp = Nothing
    Set oSearch = Nothing

    CO_GetSearchField = Err.Number
    Err.Clear
End Function

Function CO_SetCurrentPin(oSinglePromptTempXML)
'**************************************************************************
'Purpose:   put <curpin flag="1"> in oSinglePromptTempXML/<temp>
'Inputs:    oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'**************************************************************************
    On Error Resume Next
    Dim oRootAnswer
    Dim oTemp
    Dim oCurpin

    Set oRootAnswer = oSinglePromptTempXML.selectSingleNode("/")

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If (oTemp Is Nothing) Then
        Set oTemp = oRootAnswer.createElement("temp")
        Call oSinglePromptTempXML.appendChild(oTemp)
    End If

    Set oCurpin = oTemp.selectSingleNode("./curpin")
    If (oCurpin Is Nothing) Then
        Set oCurpin = oRootAnswer.createElement("curpin")
        Call oTemp.appendChild(oCurpin)
    End If

    Call oCurpin.setAttribute("flag", "1")

    Set oRootAnswer = Nothing
    Set oTemp = Nothing
    Set oCurpin = Nothing

    CO_SetCurrentPin = Err.Number
    Err.Clear
End Function

Function CO_IsCurrentPin(oSinglePromptTempXML, bCurrent)
'**************************************************************************
'Purpose:   check If <curpin flag="1"> in oSinglePromptTempXML/<temp>
'Inputs:    oSinglePromptTempXML,
'Outputs:   bCurrent
'**************************************************************************
    On Error Resume Next
    Dim oTemp
    Dim oCurpin
    Dim sFlag

    bCurrent = False

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oCurpin = oTemp.selectSingleNode("./curpin")
        If Not (oCurpin Is Nothing) Then
            sFlag = oCurpin.getAttribute("flag")
            If StrComp(CStr(sFlag), "1", vbBinaryCompare) = 0 Then
                bCurrent = True
            End If
        End If
    End If

    Set oTemp = Nothing
    Set oCurpin = Nothing

    CO_IsCurrentPin = Err.Number
    Err.Clear
End Function

Function CO_ClearCurrentPin(oSinglePromptTempXML)
'**************************************************************************
'Purpose:   remove current flag in oPromptsTempXML/<temp>
'Inputs:    oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'**************************************************************************

    On Error Resume Next
    Dim oTemp
    Dim oCurpin

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oCurpin = oTemp.selectSingleNode("./curpin")
        If Not (oCurpin Is Nothing) Then
            Call oCurpin.setAttribute("flag", "0")
        End If
    End If

    Set oTemp = Nothing
    Set oCurpin = Nothing

    CO_ClearCurrentPin = Err.Number
    Err.Clear
End Function


Function CO_RemoveCurrentPin(oPromptsTempXML)
'**************************************************************************
'Purpose:   remove current prompt index in oPromptsTempXML
'Inputs:    oPromptsTempXML
'Outputs:   oPromptsTempXML
'**************************************************************************
    On Error Resume Next
    Dim oSinglePromptTempXML
    Dim oTemp
    Dim oCurpin

    For Each oSinglePromptTempXML In oPromptsTempXML.selectNodes("/mi/pif")
        Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
        If Not (oTemp Is Nothing) Then
            Set oCurpin = oTemp.selectSingleNode("./curpin")
            If Not (oCurpin Is Nothing) Then
                Call oCurpin.setAttribute("flag", "0")
            End If
        End If
    Next

    Set oSinglePromptTempXML = Nothing
    Set oTemp = Nothing
    Set oCurpin = Nothing

    CO_RemoveCurrentPin = Err.Number
    Err.Clear
End Function

Function CO_ClearRemove(oSinglePromptTempXML)
'**************************************************************************
'Purpose:   Remove <temp/remove> in oSinglePromptTempXML
'Inputs:    oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'**************************************************************************
    On Error Resume Next
    Dim oRemove
    Dim oRemoveParent

    Set oRemove = oSinglePromptTempXML.selectSingleNode("./temp/remove")
    If Not (oRemove Is Nothing) Then
        Set oRemoveParent = oRemove.parentNode
        oRemoveParent.removeChild (oRemove)
    End If

    Set oRemove = Nothing
    Set oRemoveParent = Nothing

    CO_ClearRemove = Err.Number
    Err.Clear
End Function

Function CO_SetAttributeforHIPrompt(oSinglePromptTempXML, sATName, sATDid)
'**************************************************************************
'Purpose:   put sBlockBegin in oSinglePromptTempXML/<temp>
'Inputs:    oSinglePromptTempXML, sATName, sATDid
'Outputs:   Err.number
'**************************************************************************
    On Error Resume Next
    Dim oRootAnswer
    Dim oTemp
    Dim oAT

    Set oRootAnswer = oSinglePromptTempXML.selectSingleNode("/")

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If (oTemp Is Nothing) Then
        Set oTemp = oRootAnswer.createElement("temp")
        Call oSinglePromptTempXML.appendChild(oTemp)
    End If

    Set oAT = oTemp.selectSingleNode("./at")
    If (oAT Is Nothing) Then
        Set oAT = oRootAnswer.createElement("at")
        Call oTemp.appendChild(oAT)
    End If

    Call oAT.setAttribute("n", CStr(sATName))
    Call oAT.setAttribute("did", CStr(sATDid))

    Set oRootAnswer = Nothing
    Set oTemp = Nothing
    Set oAT = Nothing

    CO_SetAttributeforHIPrompt = Err.Number
    Err.Clear
End Function

Function CO_GetAttributeforHIPrompt(oSinglePromptTempXML, sATName, sATDid)
'**************************************************************************
'Purpose:   get attribute did For a Hierachical prompt from oPromptsTempXML
'Inputs:    oSinglePromptTempXML
'Outputs:   sATName, sATDid
'**************************************************************************
    On Error Resume Next
    Dim oTemp
    Dim oAT

    sATDid = ""
    sATName = ""
    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oAT = oTemp.selectSingleNode("./at")
        If Not (oAT Is Nothing) Then
            sATDid = oAT.getAttribute("did")
            sATName = oAT.getAttribute("n")
        End If
    End If

    CO_GetAttributeforHIPrompt = Err.Number
    Err.Clear
End Function

Function CO_ClearAttributeforHIPrompt(oSinglePromptTempXML)
'**************************************************************************
'Purpose:   remove attribute did For a Hierachical prompt from oPromptsTempXML/<temp>
'Inputs:    oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'**************************************************************************
    On Error Resume Next
    Dim oTemp
    Dim oAT

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oAT = oTemp.selectSingleNode("./at")
        If Not (oAT Is Nothing) Then
            Call oTemp.removeChild(oAT)
        End If
    End If

    Set oTemp = Nothing
    Set oAT = Nothing

    CO_ClearAttributeforHIPrompt = Err.Number
    Err.Clear
End Function

Function CO_SetAttBeforeDrillforHIPrompt(oSinglePromptTempXML, sATName, sATDid)
'**************************************************************************
'Purpose:   Set attribute before drill For a Hierachical prompt from oPromptsTempXML
'Inputs:    oSinglePromptTempXML, sATDid, sATName
'Outputs:   oSinglePromptTempXML
'**************************************************************************
    On Error Resume Next
    Dim oRootAnswer
    Dim oTemp
    Dim oATBeforeDrill

    Set oRootAnswer = oSinglePromptTempXML.selectSingleNode("/")

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If (oTemp Is Nothing) Then
        Set oTemp = oRootAnswer.createElement("temp")
        Call oSinglePromptTempXML.appendChild(oTemp)
    End If

    Set oATBeforeDrill = oTemp.selectSingleNode("./atbeforedrill")
    If (oATBeforeDrill Is Nothing) Then
        Set oATBeforeDrill = oRootAnswer.createElement("atbeforedrill")
        Call oTemp.appendChild(oATBeforeDrill)
    End If

    Call oATBeforeDrill.setAttribute("n", CStr(sATName))
    Call oATBeforeDrill.setAttribute("did", CStr(sATDid))

    Set oRootAnswer = Nothing
    Set oTemp = Nothing
    Set oATBeforeDrill = Nothing

    CO_GetAttBeforeDrillforHIPrompt = Err.Number
    Err.Clear
End Function


Function CO_GetAttBeforeDrillforHIPrompt(oSinglePromptTempXML, sATName, sATDid)
'**************************************************************************
'Purpose:   get attribute before drill For a Hierachical prompt from oPromptsTempXML
'Inputs:    oSinglePromptTempXML
'Outputs:   sATDid, sATName
'**************************************************************************
    On Error Resume Next
    Dim oTemp
    Dim oATBeforeDrill

    sATDid = ""
    sATName = ""
    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oATBeforeDrill = oTemp.selectSingleNode("./atbeforedrill")
        If Not (oATBeforeDrill Is Nothing) Then
            sATDid = oATBeforeDrill.getAttribute("did")
            sATName = oATBeforeDrill.getAttribute("n")
        End If
    End If

    CO_GetAttBeforeDrillforHIPrompt = Err.Number
    Err.Clear
End Function

'???????????????????????????
Function CO_SetFolderXMLForHIBrowse(oSinglePromptTempXML, oSearchResMI)
'**************************************************************************
'Purpose:   put folderXML (search result) in <temp> For ObjectPrompt with HIBrowsing
'Inputs:    oSinglePromptTempXML, oSearchResMI
'Outputs:   oSinglePromptTempXML
'**************************************************************************
    On Error Resume Next
    Dim oRootAnswer
    Dim oTemp
    Dim oSearchRes

    Set oRootAnswer = oSinglePromptTempXML.selectSingleNode("/")

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If (oTemp Is Nothing) Then
        Set oTemp = oRootAnswer.createElement("temp")
        Call oSinglePromptTempXML.appendChild(oTemp)
    End If

    Set oSearchRes = oTemp.selectSingleNode("./searchres")
    If Not (oSearchRes Is Nothing) Then
        Call oTemp.removeChild(oSearchRes)
    End If

    Set oSearchRes = oRootAnswer.createElement("searchres")
    Call oTemp.appendChild(oSearchRes)

    Call oSearchRes.appendChild(oSearchResMI.cloneNode(True))

    Set oRootAnswer = Nothing
    Set oTemp = Nothing
    Set oSearchRes = Nothing

    CO_SetFolderXMLForHIBrowse = Err.Number
    Err.Clear
End Function

Function CO_SetPathXMLForHIBrowse(oSinglePromptTempXML, oPathMI)
'**************************************************************************
'Purpose:   put pathXML (ancestor info) in <temp> For ObjectPrompt with HIBrowsing
'Inputs:    oSinglePromptTempXML, oPathMI
'Outputs:   oSinglePromptTempXML
'**************************************************************************
    On Error Resume Next
    Dim oRootAnswer
    Dim oTemp
    Dim oPath

    Set oRootAnswer = oSinglePromptTempXML.selectSingleNode("/")

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If (oTemp Is Nothing) Then
        Set oTemp = oRootAnswer.createElement("temp")
        Call oSinglePromptTempXML.appendChild(oTemp)
    End If

    Set oPath = oTemp.selectSingleNode("./path")
    If Not (oPath Is Nothing) Then
        Call oTemp.removeChild(oPath)
    End If

    Set oPath = oRootAnswer.createElement("path")
    Call oTemp.appendChild(oPath)

    Call oPath.appendChild(oPathMI.cloneNode(True))

    Set oRootAnswer = Nothing
    Set oTemp = Nothing
    Set oPath = Nothing

    CO_SetPathXMLForHIBrowse = Err.Number
    Err.Clear
End Function

Function CO_GetPathXMLforHIBrowse(oSinglePromptTempXML, oPathMI)
'**************************************************************************
'Purpose:   get pathXML (ancestor info) in <temp> For ObjectPrompt with HIBrowsing
'Inputs:    oSinglePromptTempXML
'Outputs:   oPathMI
'**************************************************************************
    On Error Resume Next
    Dim oTemp

    Set oPathMI = Nothing
    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oPathMI = oTemp.selectSingleNode("./path/mi")
    End If

    CO_GetPathXMLforHIBrowse = Err.Number
    Err.Clear
End Function

Function CO_SetHiLinkforObjectPrompt(oSinglePromptTempXML, sFDDid)
'**************************************************************************
'Purpose:   Set current folder did For object prompt in <temp>
'Inputs:    oSinglePromptTempXML, sFDDid
'Outputs:   oSinglePromptTempXML
'**************************************************************************
    On Error Resume Next
    Dim oRootAnswer
    Dim oTemp
    Dim oHilink

    Set oRootAnswer = oSinglePromptTempXML.selectSingleNode("/")

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If (oTemp Is Nothing) Then
        Set oTemp = oRootAnswer.createElement("temp")
        Call oSinglePromptTempXML.appendChild(oTemp)
    End If

    Set oHilink = oTemp.selectSingleNode("./hilink")
    If (oHilink Is Nothing) Then
        Set oHilink = oRootAnswer.createElement("hilink")
        Call oTemp.appendChild(oHilink)
    End If

    Call oHilink.setAttribute("did", CStr(sFDDid))

    Set oRootAnswer = Nothing
    Set oTemp = Nothing
    Set oHilink = Nothing

    CO_SetHiLinkforObjectPrompt = Err.Number
    Err.Clear
End Function

Function CO_GetHiLinkforObjectPrompt(oSinglePromptTempXML, sFDDid)
'**************************************************************************
'Purpose:   get attribute before drill For a Hierachical prompt from oPromptsTempXML
'Inputs:    oSinglePromptTempXML
'Outputs:   sFDDid
'**************************************************************************
    On Error Resume Next
    Dim oTemp
    Dim oHilink

    sFDDid = ""
    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oHilink = oTemp.selectSingleNode("./hilink")
        If Not (oHilink Is Nothing) Then
            sFDDid = oHilink.getAttribute("did")
        End If
    End If

    CO_GetHiLinkforObjectPrompt = Err.Number
    Err.Clear
End Function

Function CO_SetFilterXMLForDrillInHIPrompt(oSinglePromptTempXML, sFilterExp_Drill)
'**************************************************************************
'Purpose:   put filterXML For Drill
'Inputs:    oSinglePromptTempXML, oFilterXML_Drill_F
'Outputs:   oSinglePromptTempXML
'**************************************************************************
    On Error Resume Next
    Dim oRootAnswer
    Dim oTemp
    Dim oFilterDrill

    Set oRootAnswer = oSinglePromptTempXML.selectSingleNode("/")

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If (oTemp Is Nothing) Then
        Set oTemp = oRootAnswer.createElement("temp")
        Call oSinglePromptTempXML.appendChild(oTemp)
    End If

    Set oFilterDrill = oTemp.selectSingleNode("./filterdrill")
    If (oFilterDrill Is Nothing) Then
        Set oFilterDrill = oRootAnswer.createElement("filterdrill")
        Call oTemp.appendChild(oFilterDrill)
    End If

    oFilterDrill.Text = sFilterExp_Drill

    Set oRootAnswer = Nothing
    Set oTemp = Nothing
    Set oFilterDrill = Nothing

    CO_SetFilterXMLForDrillInHIPrompt = Err.Number
    Err.Clear
End Function


Function CO_GetFilterXMLForDrillInHIPrompt(aConnectionInfo, oSinglePromptTempXML, sFilterExp_Drill)
'**************************************************************************
'Purpose:   get filterXML For Drill
'Inputs:    oSinglePromptTempXML
'Outputs:   oFilterXML_Drill_F
'**************************************************************************
    On Error Resume Next
    Dim oTemp
    Dim oFilterDrill

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oFilterDrill = oTemp.selectSingleNode("./filterdrill")
        If Not (oFilterDrill Is Nothing) Then
            sFilterExp_Drill = oFilterDrill.Text
        End If
    End If

    CO_GetFilterXMLForDrillInHIPrompt = Err.Number
    Err.Clear
End Function

Function CO_ClearFilterXMLForDrillInHIPrompt(oSinglePromptTempXML)
'**************************************************************************
'Purpose:   remove filterXML For Drill in AnswerXML
'Inputs:    oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'**************************************************************************
    On Error Resume Next
    Dim oTemp
    Dim oFilterXML_Drill

    Set oFilterXML_Drill = Nothing
    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oFilterXML_Drill = oTemp.selectSingleNode("./filterdrill")
        If Not (oFilterXML_Drill Is Nothing) Then
            Call oTemp.removeChild(oFilterXML_Drill)
        End If
    End If

    CO_ClearFilterXMLForDrillInHIPrompt = Err.Number
    Err.Clear
End Function

Function CO_GetFilterOperator(oSinglePromptTempXML, sOP)
'**************************************************************************
'Purpose:   get filter operator in oSinglePromptTempXML/<temp>
'Inputs:    oSinglePromptTempXML
'Outputs:   sOP
'**************************************************************************

    On Error Resume Next
    Dim oTemp
    Dim oFilterOP

    sOP = ""
    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oFilterOP = oTemp.selectSingleNode("./filterop")
        If Not (oFilterOP Is Nothing) Then
            sOP = oFilterOP.getAttribute("op")
        End If
    End If

    Set oTemp = Nothing
    Set oFilterOP = Nothing

    CO_GetFilterOperator = Err.Number
    Err.Clear
End Function

Function CO_SetHierachyDIDForHIPrompt(oSinglePromptTempXML, sHierachyDID)
'**************************************************************************
'Purpose:   put Hierachy DID in <temp> For hierachical prompt on all dimensions
'Inputs:    oSinglePromptTempXML, sHierachyDID
'Outputs:   oSinglePromptTempXML
'**************************************************************************

    On Error Resume Next
    Dim oRootAnswer
    Dim oTemp
    Dim oHierachy

    Set oRootAnswer = oSinglePromptTempXML.selectSingleNode("/")

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If (oTemp Is Nothing) Then
        Set oTemp = oRootAnswer.createElement("temp")
        Call oSinglePromptTempXML.appendChild(oTemp)
    End If

    Set oHierachy = oTemp.selectSingleNode("./hierDID")
    If (oHierachy Is Nothing) Then
        Set oHierachy = oRootAnswer.createElement("hierDID")
        Call oTemp.appendChild(oHierachy)
    End If

    Call oHierachy.setAttribute("did", CStr(sHierachyDID))

    Set oRootAnswer = Nothing
    Set oTemp = Nothing
    Set oHierachy = Nothing

    CO_SetHierachyDIDForHIPrompt = Err.Number
    Err.Clear
End Function

Function CO_GetHierachyDIDForHIPrompt(oSinglePromptTempXML, sHierachyDID)
'**************************************************************************
'Purpose:   get Hierachy in <temp> For ObjectPrompt with HIBrowsing
'Inputs:    oSinglePromptTempXML
'Outputs:   sHierachyDID
'**************************************************************************
    On Error Resume Next
    Dim oTemp
    Dim oHierachy

    sHierachyDID = ""
    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oHierachy = oTemp.selectSingleNode("./hierDID")
        If Not (oHierachy Is Nothing) Then
            sHierachyDID = oHierachy.getAttribute("did")
        End If
    End If

    Set oTemp = Nothing
    Set oHierachy = Nothing
    CO_GetHierachyDIDForHIPrompt = Err.Number
    Err.Clear
End Function

Function CO_ClearHierachyDIDForHIPrompt(oSinglePromptTempXML)
'**************************************************************************
'Purpose:   remove current hiearchy did For prompt on all dimension
'Inputs:    oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'**************************************************************************
    On Error Resume Next
    Dim oTemp
    Dim oHI

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oHI = oTemp.selectSingleNode("./hierDID")
        If Not (oHI Is Nothing) Then
            Call oTemp.removeChild(oHI)
        End If
    End If

    Set oTemp = Nothing
    Set oHI = Nothing

    CO_ClearHierachyDIDForHIPrompt = Err.Number
    Err.Clear
End Function

Function CO_SetHiFlagForHIPrompt(oSinglePromptTempXML, sFlag)
'**************************************************************************
'Purpose:   put HiFlag in oSinglePromptTempXML/<temp> to be displayed
'Inputs:    oSinglePromptTempXML, sFlag
'Outputs:   oSinglePromptTempXML
'**************************************************************************
    On Error Resume Next
    Dim oRootAnswer
    Dim oTemp
    Dim oFlag

    Set oRootAnswer = oSinglePromptTempXML.selectSingleNode("/")

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If (oTemp Is Nothing) Then
        Set oTemp = oRootAnswer.createElement("temp")
        Call oSinglePromptTempXML.appendChild(oTemp)
    End If

    Set oFlag = oTemp.selectSingleNode("./hiflag")
    If (oFlag Is Nothing) Then
        Set oFlag = oRootAnswer.createElement("hiflag")
        Call oTemp.appendChild(oFlag)
    End If

    Call oFlag.setAttribute("text", CStr(sFlag))

    Set oRootAnswer = Nothing
    Set oTemp = Nothing
    Set oFlag = Nothing

    CO_SetHiFlagForHIPrompt = Err.Number
    Err.Clear
End Function

Function CO_GetHiFlagForHIPrompt(oSinglePromptTempXML, bAllDimension, sFlag)
'**************************************************************************
'Purpose:   get HiFlag in oSinglePromptTempXML/<temp>
'Inputs:    oSinglePromptTempXML, bAllDimension
'Outputs:   sFlag
'**************************************************************************
    On Error Resume Next
    Dim oTemp
    Dim oFlag

    sFlag = ""
    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oFlag = oTemp.selectSingleNode("./hiflag")
        If Not (oFlag Is Nothing) Then
            sFlag = oFlag.getAttribute("text")
        End If
    End If
    If sFlag = "" Then
        If bAllDimension Then
            sFlag = "PICK_ELEM"
        Else
            sFlag = "ELEM"
        End If
    End If

    Set oTemp = Nothing
    Set oFlag = Nothing

    CO_GetHiFlagForHIPrompt = Err.Number
    Err.Clear
End Function

'*****************
Function CO_SetSubFolderForPromptAllDimensions(oSinglePromptTempXML, sSubFolderDID)
'**************************************************************************
'Purpose:   put sub folder in oSinglePromptTempXML/<temp> to be displayed
'Inputs:    oSinglePromptTempXML, sSubFolderDID
'Outputs:   oSinglePromptTempXML
'**************************************************************************
    On Error Resume Next
    Dim oRootAnswer
    Dim oTemp
    Dim oSubFolder

    Set oRootAnswer = oSinglePromptTempXML.selectSingleNode("/")

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If (oTemp Is Nothing) Then
        Set oTemp = oRootAnswer.createElement("temp")
        Call oSinglePromptTempXML.appendChild(oTemp)
    End If

    Set oSubFolder = oTemp.selectSingleNode("./sf")
    If (oSubFolder Is Nothing) Then
        Set oSubFolder = oRootAnswer.createElement("sf")
        Call oTemp.appendChild(oSubFolder)
    End If

    Call oSubFolder.setAttribute("did", CStr(sSubFolderDID))

    Set oRootAnswer = Nothing
    Set oTemp = Nothing
    Set oSubFolder = Nothing

    CO_SetSubFolderForPromptAllDimensions = Err.Number
    Err.Clear
End Function

Function CO_GetSubFolderForPromptAllDimensions(oSinglePromptTempXML, sSubFolderDID)
'**************************************************************************
'Purpose:   get subfolder in oSinglePromptTempXML/<temp>
'Inputs:    oSinglePromptTempXML
'Outputs:   sSubFolderDID
'**************************************************************************
    On Error Resume Next
    Dim oTemp
    Dim oSubFolder

    sSubFolderDID = ""
    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oSubFolder = oTemp.selectSingleNode("./sf")
        If Not (oSubFolder Is Nothing) Then
            sSubFolderDID = oSubFolder.getAttribute("did")
        End If
    End If

    Set oTemp = Nothing
    Set oSubFolder = Nothing

    CO_GetSubFolderForPromptAllDimensions = Err.Number
    Err.Clear
End Function

Function CO_GetBlockCount(iFlag, lBlockCount)
'**************************************************************************
'Purpose:   get blockcount from options, Or use default
'Inputs:    iFlag
'Outputs:   lBlockCount
'**************************************************************************
    On Error Resume Next
    If iFlag = BLOCKCOUNT_ELEPROMPT Then
        lBlockCount = ReadUserOption(ELE_PROMPT_BLOCK_COUNT_OPTION)
    ElseIf iFlag = BLOCKCOUNT_OBJPROMPT Then
        lBlockCount = ReadUserOption(OBJ_PROMPT_BLOCK_COUNT_OPTION)
    End If
    If Len(lBlockCount) > 0 Then
        lBlockCount = CLng(lBlockCount)
    Else
        If iFlag = BLOCKCOUNT_ELEPROMPT Then
            lBlockCount = CONST_ELEPROMPT_BLOCKCOUNT
        ElseIf iFlag = BLOCKCOUNT_OBJPROMPT Then
            lBlockCount = CONST_OBJPROMPT_BLOCKCOUNT
        End If
    End If

    CO_GetBlockCount = Err.Number
    Err.Clear
End Function

Function CO_GetPromptStyleFile(oSinglePrompt, oSinglePromptQuestionXML, sXSL)
'**************************************************************************
'Purpose:   get style file name from oSinglePromptQuestionXML
'Inputs:    oSinglePromptQuestionXML
'Outputs:   sXSL
'**************************************************************************
    On Error Resume Next
    Dim oPRXSL

    Set oPRXSL = oSinglePromptQuestionXML.selectSingleNode("./prs/pr[@n='PSXSL']")
    If Not (oPRXSL Is Nothing) Then
        sXSL = oPRXSL.getAttribute("v")
    Else
        sXSL = ""
    End If

    If Len(sXSL) = 0 Or StrComp(sXSL, "DefaultPSXSL.XSL", vbTextCompare) = 0 Then
        Select Case oSinglePrompt.PromptType
            Case DssXmlPromptObjects
                sXSL = "PromptObject_Cart.xsl"
            Case DssXmlPromptElements
                sXSL = "PromptElement_Cart.xsl"
            Case DssXmlPromptExpression
                If oSinglePrompt.ExpressionType = DssXmlFilterAllAttributeQual or oSinglePrompt.ExpressionType = DssXmlExpressionMDXSAPVariable Then
                    sXSL = "PromptExpression_HierCart_OptSearch_Drill_Qual.xsl"
                Else
                    sXSL = "PromptExpression_Cart.xsl"
                End If
            Case DssXmlPromptDimty
                sXSL = "PromptLevel_Cart.xsl"
            Case DssXmlPromptLong, DssXmlPromptString, DssXmlPromptDouble, DssXmlPromptDate
                 sXSL = "PromptConstant_Textbox.xsl"

        End Select
	End If

    CO_GetPromptStyleFile = Err.Number
    Err.Clear
End Function

Function CO_GetPromptCartProperty(oSinglePrompt, oSinglePromptQuestionXML, bIsCart)
'**************************************************************************
'Purpose:   get prompt cart property from oSinglePromptQuestionXML
'Inputs:    oSinglePromptQuestionXML
'Outputs:   bIsCart
'**************************************************************************
    On Error Resume Next
    Dim oPRCART
    Dim sIsCart
    Dim oPRXSL
    Dim sXSL
    Dim oRes
    Dim sRes

    bIsCart = False
    Select Case oSinglePrompt.PromptType
	Case DssXmlPromptLong, DssXmlPromptString, DssXmlPromptDouble, DssXmlPromptDate
	Case Else
	    Set oPRXSL = oSinglePromptQuestionXML.selectSingleNode("./prs/pr[@n='PSXSL']")
	    If Not (oPRXSL Is Nothing) Then
		sXSL = oPRXSL.getAttribute("v")
		If Len(sXSL) = 0 Or _
		   StrComp(sXSL, "DefaultPSXSL.XSL", vbTextCompare) = 0 Or _
		   StrComp(sXSL, "PromptExpression_cart.xsl", vbTextCompare) = 0 Or _
		   StrComp(sXSL, "PromptElement_cart.xsl", vbTextCompare) = 0 Or _
		   StrComp(sXSL, "PromptExpression_HierCart_OPTsearch_drill_qual.xsl", vbTextCompare) = 0 Or _
		   StrComp(sXSL, "PromptExpression_HierCart_OPTsearch_drill_Noqual.xsl", vbTextCompare) = 0 Or _
		   StrComp(sXSL, "PromptExpression_cart.xsl", vbTextCompare) = 0 Or _
		   StrComp(sXSL, "PromptObject_cart.xsl", vbTextCompare) = 0 Or _
		   StrComp(sXSL, "PromptObject_cart.xsl", vbTextCompare) = 0 Or _
		   StrComp(sXSL, "PromptExpression_HierCart_REQsearch_drill_NOqual.xsl", vbTextCompare) = 0 Or _
		   StrComp(sXSL, "PromptExpression_HierCart_REQsearch_drill_qual.xsl", vbTextCompare) = 0 Or _
		   StrComp(sXSL, "PromptExpression_HierTree.xsl", vbTextCompare) = 0 Or _
		   StrComp(sXSL, "PromptLevel_cart.xsl", vbTextCompare) = 0 Or _
		   StrComp(sXSL, "PromptExpression_textfile.xsl", vbTextCompare) = 0 Or _
		   StrComp(sXSL, "PromptExpression_textbox.xsl", vbTextCompare) = 0 _
		Then
			bIsCart = True
		End If
	   End If
    End Select

    Set oPRCART = Nothing
    Set oPRXSL = Nothing
    CO_GetPromptCartProperty = Err.Number
    Err.Clear
End Function

'*************
Function CO_SetRemove(oRemove, oSinglePromptTempXML)
'**************************************************************************
'Purpose:   add remove information to oSinglePromptTempXML/<temp>
'Inputs:    oRemove, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'**************************************************************************
    On Error Resume Next
    Dim oRootAnswer
    Dim oTemp

    Set oRootAnswer = oSinglePromptTempXML.selectSingleNode("/")

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If (oTemp Is Nothing) Then
        Set oTemp = oRootAnswer.createElement("temp")
        Call oSinglePromptTempXML.appendChild(oTemp)
    End If

    If Not (oRemove Is Nothing) Then
        Call CO_ClearRemove(oSinglePromptTempXML)
        Call oTemp.appendChild(oRemove)
    End If

    Set oRootAnswer = Nothing
    Set oTemp = Nothing

    CO_SetRemove = Err.Number
    Err.Clear
End Function

Function CO_SetRemoveByName(sATDid, sFMDid, sMTDid, sOP, oSinglePromptTempXML)
'**************************************************************************
'Purpose:   add remove information to oSinglePromptTempXML/<temp>
'Inputs:    sATName, sFMName, sMTName, sOP, oSinglePromptTempXML
'Outputs:   oSinglePromptTempXML
'**************************************************************************
    On Error Resume Next
    Dim oRootAnswer
    Dim oTemp
    Dim oRemove

    Set oRootAnswer = oSinglePromptTempXML.selectSingleNode("/")
    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If (oTemp Is Nothing) Then
        Set oTemp = oRootAnswer.createElement("temp")
        Call oSinglePromptTempXML.appendChild(oTemp)
    End If

    Call CO_ClearRemove(oSinglePromptTempXML)
    Set oRemove = oRootAnswer.createElement("remove")
    Call oRemove.setAttribute("atdid", sATDid)
    Call oRemove.setAttribute("fmdid", sFMDid)
    Call oRemove.setAttribute("mtdid", sMTDid)
    Call oRemove.setAttribute("op", sOP)

    If Not (oRemove Is Nothing) Then
        Call oTemp.appendChild(oRemove)
    End If

    Set oRootAnswer = Nothing
    Set oTemp = Nothing
    Set oRemove = Nothing

    CO_SetRemoveByName = Err.Number
    Err.Clear
End Function

Function CO_SetbUnknownDef(oSinglePromptTempXML, bUnknownDef)
'**************************************************************************
'Purpose:   put bUnknownDef in oSinglePromptTempXML/<temp> to be used later
'Inputs:    oSinglePromptTempXML, bUnknownDef
'Outputs:   oSinglePromptTempXML
'**************************************************************************
    On Error Resume Next
    Dim oRootAnswer
    Dim oTemp
    Dim oUnknownDef

    Set oRootAnswer = oSinglePromptTempXML.selectSingleNode("/")

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If (oTemp Is Nothing) Then
        Set oTemp = oRootAnswer.createElement("temp")
        Call oSinglePromptTempXML.appendChild(oTemp)
    End If

    Set oUnknownDef = oTemp.selectSingleNode("./unknowndef")
    If (oUnknownDef Is Nothing) Then
        Set oUnknownDef = oRootAnswer.createElement("unknowndef")
        Call oTemp.appendChild(oUnknownDef)
    End If

    If bUnknownDef Then
        Call oUnknownDef.setAttribute("value", "1")
    Else
        Call oUnknownDef.setAttribute("value", "0")
    End If

    Set oRootAnswer = Nothing
    Set oTemp = Nothing
    Set oUnknownDef = Nothing

    CO_SetbUnknownDef = Err.Number
    Err.Clear
End Function

Function CO_GetbUnknownDef(oSinglePromptTempXML, bUnknownDef)
'**************************************************************************
'Purpose:   get bUnknownDef in oSinglePromptTempXML/<temp>
'Inputs:    oSinglePromptTempXML
'Outputs:   bUnknownDef
'**************************************************************************

    On Error Resume Next
    Dim oTemp
    Dim oUnknownDef

    bUnknownDef = False
    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oUnknownDef = oTemp.selectSingleNode("./unknowndef")
        If Not (oUnknownDef Is Nothing) Then
            If StrComp(oUnknownDef.getAttribute("value"), "1") = 0 Then
                bUnknownDef = True
            End If
        End If
    End If

    Set oTemp = Nothing
    Set oUnknownDef = Nothing

    CO_GetbUnknownDef = Err.Number
    Err.Clear
End Function

Function CO_SetDisplayUnknownDef(oSinglePromptTempXML, sDisplayUnknownDef)
'**************************************************************************
'Purpose:   put sDisplayUnknownDef in oSinglePromptTempXML/<temp> to be used later
'Inputs:    oSinglePromptTempXML, sDisplayUnknownDef
'Outputs:   oSinglePromptTempXML
'**************************************************************************
    On Error Resume Next
    Dim oRootAnswer
    Dim oTemp
    Dim oDisplayUnknownDef

    Set oRootAnswer = oSinglePromptTempXML.selectSingleNode("/")

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If (oTemp Is Nothing) Then
        Set oTemp = oRootAnswer.createElement("temp")
        Call oSinglePromptTempXML.appendChild(oTemp)
    End If

    Set oDisplayUnknownDef = oTemp.selectSingleNode("./displayunknowndef")
    If (oDisplayUnknownDef Is Nothing) Then
        Set oDisplayUnknownDef = oRootAnswer.createElement("displayunknowndef")
        Call oTemp.appendChild(oDisplayUnknownDef)
    End If
    Call oDisplayUnknownDef.setAttribute("value", sDisplayUnknownDef)

    Set oRootAnswer = Nothing
    Set oTemp = Nothing
    Set oDisplayUnknownDef = Nothing

    CO_SetDisplayUnknownDef = Err.Number
    Err.Clear
End Function

Function CO_GetDisplayUnknownDef(oSinglePromptTempXML, sDisplayUnknownDef)
'**************************************************************************
'Purpose:   get sDisplayUnknownDef in oSinglePromptTempXML/<temp>
'Inputs:    oSinglePromptTempXML
'Outputs:   sDisplayUnknownDef
'**************************************************************************

    On Error Resume Next
    Dim oTemp
    Dim oDisplayUnknownDef

    sDisplayUnknownDef = "0"
    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oDisplayUnknownDef = oTemp.selectSingleNode("./displayunknowndef")
        If Not (oDisplayUnknownDef Is Nothing) Then
            sDisplayUnknownDef = oDisplayUnknownDef.getAttribute("value")
        End If
    End If

    Set oTemp = Nothing
    Set oDisplayUnknownDef = Nothing

    CO_GetDisplayUnknownDef = Err.Number
    Err.Clear
End Function

Function CO_SetMatchCase(oSinglePromptTempXML, sMatchCase)
'**************************************************************************
'Purpose:   put sMatchCase in oSinglePromptTempXML/<temp>
'Inputs:    oSinglePromptTempXML, sMatchCase
'Outputs:   Err.number
'**************************************************************************
    On Error Resume Next
    Dim oRootAnswer
    Dim oTemp
    Dim oCase

    Set oRootAnswer = oSinglePromptTempXML.selectSingleNode("/")

    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If (oTemp Is Nothing) Then
        Set oTemp = oRootAnswer.createElement("temp")
        Call oSinglePromptTempXML.appendChild(oTemp)
    End If

    Set oCase = oTemp.selectSingleNode("./case")
    If (oCase Is Nothing) Then
        Set oCase = oRootAnswer.createElement("case")
        Call oTemp.appendChild(oCase)
    End If

    Call oCase.setAttribute("value", CStr(sMatchCase))

    Set oRootAnswer = Nothing
    Set oTemp = Nothing
    Set oCase = Nothing

    CO_SetMatchCase = Err.Number
    Err.Clear
End Function

Function CO_GetMatchCase(oSinglePromptTempXML, sMatchCase)
'**************************************************************************
'Purpose:   get MatchCase in oSinglePromptTempXML/<temp>
'Inputs:    oSinglePromptTempXML
'Outputs:   sMatchCase
'**************************************************************************
    On Error Resume Next
    Dim oTemp
    Dim oCase

    sMatchCase = ""
    Set oTemp = oSinglePromptTempXML.selectSingleNode("./temp")
    If Not (oTemp Is Nothing) Then
        Set oCase = oTemp.selectSingleNode("./case")
        If Not (oCase Is Nothing) Then
            sMatchCase = oCase.getAttribute("value")
        End If
    End If

    Set oTemp = Nothing
    Set oCase = Nothing

    CO_GetMatchCase = Err.Number
    Err.Clear
End Function

%>

