<%@LANGUAGE=VBSCRIPT%>
<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Option Explicit
On Error Resume Next

Dim oBinaryRequest
Dim iRequestSize
Dim oRequest
Dim sErrDescription

Set oRequest = Nothing
iRequestSize = Request.TotalBytes
oBinaryRequest = Request.BinaryRead(iRequestSize)

Call GetRequestFromBinaryData(oBinaryRequest, iRequestSize, oRequest, sErrDescription)

Response.CacheControl = "no-cache"      'enforce go back to reexecute the page
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="CustomLib\CommonLib.asp" -->
<%
Function GetRequestFromBinaryData(oBinaryRequest, iRequestSize, oRequest, sErrDescription)
    On Error Resume Next
    Dim oStringUtilities
    Dim oTempDict, oTempDictForCollection
    Dim sFileRepresentationContents
    Dim sParameter
    Dim sValue
    Dim sDelimeter
    Dim iGeneralPosition
    Dim iStart
    Dim iEnd
    Dim lErrNumber

    	If iRequestSize > 0 Then
		Set oStringUtilities = Server.CreateObject("M9StrUtl.StringUtilities")
		lErrNumber = Err.number
		If lErrNumber <> NO_ERR Then
		    Call LogErrorXML(aConnectionInfo, CStr(lErrNumber), CStr(Err.description), Err.source, "PromptCuLib.asp", "GetFileContentsForFilterSpecifications", "", "Error creating an instance of the class M9StrUtl.StringUtilities", LogLevelError)
		    sErrDescription = asDescriptors(1257) 'Descriptor: The following file has not been registered correctly in the Web Server: M9StrUtl.dll. Please ask the Administrator for more details.
		Else
			Set oTempDict = Server.CreateObject("Scripting.Dictionary")
			Set oTempDictForCollection = Server.CreateObject("Scripting.Dictionary")
			sFileRepresentationContents = oStringUtilities.StringConverte(oBinaryRequest, 64, VBNull) 'vbUnicode = 64
			sDelimeter = Left(sFileRepresentationContents, InStr(1, sFileRepresentationContents, vbCrLf) - Len(vbCrLf))

			iGeneralPosition = 1
			While (iGeneralPosition > -1) And (iGeneralPosition < Len(sFileRepresentationContents) - Len(vbCrLf & vbCrLf))
				sParameter = ""
				sValue = ""
				iGeneralPosition = InStr(iGeneralPosition, sFileRepresentationContents, "name=""") + Len("name=""")
				sParameter = Mid(sFileRepresentationContents, iGeneralPosition, InStr(iGeneralPosition, sFileRepresentationContents, """") - iGeneralPosition)
				If InStr(iGeneralPosition, sFileRepresentationContents, """; filename=""") = iGeneralPosition + Len(sParameter) Then
					iStart = iGeneralPosition + Len(sParameter) + Len("""; filename=""")
					iEnd = InStr(iStart, sFileRepresentationContents, """")
					sValue = Mid(sFileRepresentationContents, iStart, iEnd - iStart)
					sValue = Right(sValue, Len(sValue) - InStr(1, sValue, "."))
					Call oTempDict.Add(LCase(sParameter) & "_ext", sValue)
				End If
				iStart = InStr(iGeneralPosition, sFileRepresentationContents, vbCrLf & vbCrLf) + Len(vbCrLf & vbCrLf)
				iEnd = InStr(iStart, sFileRepresentationContents, sDelimeter)
				sValue = Mid(sFileRepresentationContents, iStart, iEnd - iStart - Len(vbCrLf))
				If oTempDict.Exists(LCase(sParameter)) Then
                    If IsObject(oTempDict(LCase(sParameter))) Then
                        Call oTempDict(LCase(sParameter)).Add(sValue, sValue)
                    Else
                        Call oTempDictForCollection.RemoveAll
                        Call oTempDictForCollection.Add(oTempDict(LCase(sParameter)), oTempDict(LCase(sParameter)))
                        Call oTempDictForCollection.Add(sValue, sValue)
                        Call oTempDict.Remove(LCase(sParameter))
                        Call oTempDict.Add(LCase(sParameter), oTempDictForCollection)
                    End If
                Else
                    Call oTempDict.Add(LCase(sParameter), sValue)
                End If
                iGeneralPosition = InStr(iGeneralPosition, sFileRepresentationContents, sDelimeter) + Len(sDelimeter)
            Wend
        End If
    End If
    Set oRequest = Nothing
    Set oRequest = oTempDict

    GetRequestFromBinaryData = lErrNumber
    Err.Clear
End Function

Dim iIndex
Dim oItem
Dim sTemporal
Dim aKeys
Dim aItems

aKeys = oRequest.keys
aItems = oRequest.items
%>
<HTML>
    <HEAD>
        <TITLE>MicroStrategy Web</TITLE>
        <!-- <SCRIPT language="JavaScript" SRC="PromptFunctions.js"></SCRIPT> -->
    </HEAD>
    <BODY BGCOLOR="#FFFFFF" TOPMARGIN="0" LEFTMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" onLoad="window.document.PromptForm.submit();">
        <FORM ACTION="prompt.asp" METHOD="POST" NAME="PromptForm">
            <%For iIndex = 0 To UBound(aKeys)
                If IsObject(aItems(iIndex)) Then
                    sTemporal = ""
                    For Each oItem In aItems(iIndex)
                        Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""" & aKeys(iIndex) & """ VALUE=""" & Replace(Replace(Replace(oItem, "&", "&#38;"), "<", "&#60;"), """", "&#34;") & """ />" & vbNewLine
                    Next
                Else
                    Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""" & aKeys(iIndex) & """ VALUE=""" & Replace(Replace(Replace(aItems(iIndex), "&", "&#38;"), "<", "&#60;"), """", "&#34;") & """ />" & vbNewLine
                End If
            Next%>
        </FORM>
    </BODY>
</HTML>
