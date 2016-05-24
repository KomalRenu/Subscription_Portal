<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<!-- #include file="../CoreLib/ISCoLib.asp" -->
<%
Function RenderISList(aInformationSources)
'********************************************************
'*Purpose:  Renders the list of IS so the user can select one:
'*Inputs:
'*Outputs:  sISXML: The XML with the list
'********************************************************
Const PROCEDURE_NAME = "RenderISList"
Dim lErr

Dim oDOM
Dim oISs
Dim i

Dim sId
Dim sName
Dim sDesc
Dim sCredential


    On Error Resume Next
    lErr = 0

    If IsEmpty(aInformationSources) Then
        Response.Write "<TR><TD COLSPAN=4 ALIGN=CENTER>"
        Response.Write "<BR/><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """><B>"
        Response.Write asDescriptors(677) '"There are no Information sources defined in this project"
        Response.Write "</B></FONT>"
        Response.Write "<BR />&nbsp;</TD></TR>"
    Else

        For i = 0 To UBound(aInformationSources)

            sId = Server.HTMLEncode(aInformationSources(i, 0))
            sName = Server.HTMLEncode(aInformationSources(i, 2))
            sDesc = Server.HTMLEncode(aInformationSources(i, 3))
            sCredential = aInformationSources(i, 1)

			Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """><B>" & sName & "</B><BR>" & sDesc & "</FONT></TD>"

			Response.Write "<TD ALIGN=CENTER VALIGN=MIDDLE><INPUT TYPE=radio NAME=""is." & sId & """  VALUE=2"
			If sCredential = "2" Then Response.Write " CHECKED"
			Response.Write "></TD>"

			Response.Write "<TD></TD>"

			Response.Write "<TD ALIGN=CENTER VALIGN=MIDDLE><INPUT TYPE=radio NAME=""is." & sId & """  VALUE=1"
			If sCredential = "1" Then Response.Write " CHECKED"
			Response.Write "></TD>"

			Response.Write "<TD ALIGN=CENTER VALIGN=MIDDLE><INPUT TYPE=radio NAME=""is." & sId & """  VALUE=0"
			If sCredential = "0" Then Response.Write " CHECKED"
			Response.Write "></TD>"

			Response.Write "<TR>"
			Response.Write "<TR><TD COLSPAN=5 HEIGHT=1 BGCOLOR=#6699CC><IMG SRC=""../images/1ptrans.gif"" HEIGHT=1></TD></TR>"
        Next

    End If

End Function
%>
