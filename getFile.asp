<%@ Language=VBScript %>
<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "private"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CommonDeclarations.asp" -->
<!-- #include file="CustomLib/ReportsCuLib.asp" -->
<%

'=--File information
Dim aAttachmInfo(3)

    Response.Buffer = True
	Response.AddHeader "Content-disposition", "inline;filename=" & Request("filename") & ";"

    'Get page parameters
    aAttachmInfo(ATT_DOC_ID) = Trim(CStr(oRequest("file")))
    aAttachmInfo(ATT_INDEX)  = Trim(CStr(oRequest("index")))

    lErr = NO_ERR

    'If we got missing parameters, return error:
    If ((Len(aAttachmInfo(ATT_DOC_ID)) = 0) Or (Len(aAttachmInfo(ATT_INDEX))  = 0)) Then
        lErr = URL_MISSING_PARAMETER
    End If

    If lErr = NO_ERR Then
        If Application.Value("Stream_Attachments") <> "0" Then
            'Stream attachment directly
            lErr = StreamAttachmentContent(aAttachmInfo)

        Else
            'Get attachment info, and then send it:
            lErr = GetAttachmentInfo(aAttachmInfo)

            If lErr = NO_ERR Then
                'Set the correct MIME type based on the attachment type:
                If StrComp(CStr(aAttachmInfo(ATT_TYPE)), "text/html", vbTextCompare) = 0 or StrComp(CStr(aAttachmInfo(ATT_TYPE)), "text/plain", vbTextCompare) = 0 or StrComp(CStr(aAttachmInfo(ATT_TYPE)), "text/xml", vbTextCompare) = 0 Then
                    Response.Write aAttachmInfo(ATT_BODY)
                Else
                    Response.ContentType = aAttachmInfo(ATT_TYPE)
                    Response.BinaryWrite aAttachmInfo(ATT_BODY)
                End If

            End If
        End If
    End If

    'If we got here and no errors, the information has already been sent; so
    'end the response
    If lErr = NO_ERR Then
        Erase aAttachmInfo
        Response.End
    End If
%>
<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE></TITLE>
</HEAD>
<BODY TOPMARGIN="0" LEFTMARGIN="0" BGCOLOR="ffffff" ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT="0" MARGINWIDTH="0">
<!-- #include file="header_multi.asp" -->
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
	<TR>
		<TD WIDTH="1%" VALIGN="TOP">
			<!-- begin search box -->
				<!-- #include file="searchbox.asp" -->
			<!-- end search box -->
			<!-- begin left menu -->
			<BR /><BR />
			<!-- end left menu -->
			<IMG SRC="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD WIDTH="98%" VALIGN="TOP">
			<!-- begin center panel -->
			<BR />
			<%
			If lErr <> NO_ERR Then
			    Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(250) & " home.asp", "home.asp") 'Descriptor: Return to:
			End If
			%>
			<!-- end center panel -->
		</TD>
		<TD WIDTH="1%">
			<IMG SRC="images/1ptrans.gif" WIDTH="15" HEIGHT="1" BORDER="0" ALT="">
		</TD>
	</TR>
</TABLE>
<BR />
<!-- begin footer -->
	<!-- #include file="footer.asp" -->
<!-- end footer -->
</BODY>
</HTML>
<%
    Erase aAttachmInfo
%>