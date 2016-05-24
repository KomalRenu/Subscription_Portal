<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%

Function readFaqFile()
Dim sPath
Dim oFS
Dim oFaqFile
Dim sPageName

    On Error Resume Next

    sPageName = Left(aPageInfo(S_NAME_PAGE), Len(aPageInfo(S_NAME_PAGE)) - 4)

    'Overview pages gets a different naming convention.
    If StrComp(sPageName, "adminoverview", vbTextCompare) = 0 Then
        sPageName = "overview_sec" & lAdminSection
    End If

    If StrComp(sPageName, "adminsummary", vbTextCompare) = 0 Then
		sPageName = "summary_sec" & lAdminSection
    End If

    'Get the Path
    sPath = Server.MapPath("../help/faq_" & sPageName & ".htm")

    Set oFs = Server.CreateObject("Scripting.FileSystemObject")
    If(oFs.FileExists(sPath)) Then
        Set oFaqFile = oFS.OpenTextFile(sPath)

        Do While oFaqFile.AtEndOfStream <> True And Err.number = 0
            Call Response.Write(oFaqFile.ReadLine)
        Loop
        oFaqFile.Close
    Else
        Response.Write("No topics defined for this step.")
    End If

    Set oFaqFile = Nothing
    Set oFS = Nothing

    Err.Clear

End Function

%>
<%If (aPageInfo(N_TOOLBARS_PAGE) And HELP_TOOLBAR) > 0 Then %>
<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" BORDER="0" WIDTH="160" HEIGHT="100%">
  <TR HEIGHT="1%" >
    <TD ROWSPAN="5" BGCOLOR="#999999" WIDTH="1" HEIGHT="100%"><IMG SRC="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT="" /></TD>
    <TD ROWSPAN="5" BGCOLOR="#cccccc" WIDTH="5"><IMG SRC="../images/1ptrans.gif" WIDTH="5" HEIGHT="1" BORDER="0" ALT="" /></TD>
    <TD BGCOLOR="#cccccc" HEIGHT="5"><IMG SRC="../images/1ptrans.gif" WIDTH="1" HEIGHT="5" BORDER="0" ALT="" /></TD>
    <TD ROWSPAN="5" BGCOLOR="#cccccc" WIDTH="5"><IMG SRC="../images/1ptrans.gif" WIDTH="5" HEIGHT="1" BORDER="0" ALT="" /></TD>
  </TR>
  <TR>
    <TD BGCOLOR="#cccccc" VALIGN="TOP"><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1"><A HREF="<%=aPageInfo(S_NAME_PAGE)%>?showHelp=0&<%=aPageInfo(N_OPTIONS_WITH_LINKS_PAGE)%>"><IMG SRC="../images/close.gif" BORDER="0" ALT="" ALIGN="RIGHT" /></A><%Response.Write(asDescriptors(242))'Descriptor:Online help%></FONT></TD>
  </TR>
  <TR>
    <TD BGCOLOR="#cccccc"><HR SIZE="1" /></TD>
  </TR>
  <TR HEIGHT="98%" >
    <TD BGCOLOR="#cccccc">
      <TABLE BORDER="0" CELLPADDING="1" WIDTH="100%" HEIGHT="100%">
        <TR>
          <TD BGCOLOR="#000000">
            <TABLE BORDER="0" BGCOLOR="#FFFFFF" CELLSPACING="0" CELLPADDING="8" HEIGHT="100%" WIDTH="100%">
              <TR>
                <TD VALIGN="TOP">
                  <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1">
                    <%Call readFaqFile()%>
                  </FONT>
                </TD>
              </TR>
            </TABLE>
          </TD>
        </TR>
      </TABLE>
    </TD>
  </TR>
  <TR HEIGHT="1%" >
    <TD BGCOLOR="#cccccc" HEIGHT="5"><IMG SRC="../images/1ptrans.gif" HEIGHT="5" BORDER="0" ALT="" /></TD>
  </TR>
</TABLE>
<%End If%>