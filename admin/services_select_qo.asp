<%
  Option Explicit
  Response.CacheControl = "no-cache"
  Response.AddHeader "Pragma", "no-cache"
  Response.Expires = -1
  On Error Resume Next
%>
<!-- #include file="../CommonDeclarations.asp" -->
<!-- #include file="../CustomLib/AdminCuLib.asp" -->
<!-- #include file="../CustomLib/ServicesConfigCuLib.asp" -->
<%
Dim lStatus

Dim sFolderId
Dim sFolderLink

Dim i

Dim sFolderContentXML
Dim oFolderDOM
Dim oFolders
Dim oParent
Dim oItems
Dim oItem

Dim sIds
Dim sId

    If lErr = NO_ERR Then
        lErr = ParseRequestForSvcConfig(oRequest, aSvcConfigInfo)
    End If


    'Set the PageInfo to be used by the navigator bar and the header.
    If aSvcConfigInfo(SVCCFG_STEP) = STATIC_SS Then
        aPageInfo(S_TITLE_PAGE) = STEP_SERVICES_STATIC_SELECT_QO & " " & asDescriptors(786) '"Select Question Object"
    Else
        aPageInfo(S_TITLE_PAGE) = STEP_SERVICES_DYNAMIC_SELECT_QO & " " & asDescriptors(786) '"Select Question Object"
    End If
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SERVICES
    aPageInfo(N_OPTIONS_WITH_LINKS_PAGE) = CreateRequestForSvcConfig(aSvcConfigInfo)

    lStatus = checkSiteConfiguration()

    'Select the current folder
     If lErr = NO_ERR Then
        If Len(oRequest("fid")) > 0 Then
            sFolderId = oRequest("fid")
        Else
            If Len(aSvcConfigInfo(SVCCFG_AQ_PARENT_ID)) > 0 Then
                sFolderId = aSvcConfigInfo(SVCCFG_AQ_PARENT_ID)
            Else
                sFolderId = aSvcConfigInfo(SVCCFG_QO_PARENT_ID)
            End If
        End If

		'If the folder ID is that of Hidden Objects, change it to Blank which is the root - Applications
        If sFolderId = HIDDEN_OBJECTS_FOLDER_ID Then
			sFolderId = ""
        End If

    End If

    'Get the folder content:
    If lErr = NO_ERR Then
        lErr = co_getFolderXML(sFolderId, ROOT_APP_FOLDER_TYPE, sFolderContentXML)
        If lErr <> NO_ERR Then
            Call LogErrorXML(aConnectionInfo, lErr, "", "select_devices.asp", "", "", "Error calling co_getFolderXML", LogLevelTrace)
        Else
            lErr = LoadXMLDOMFromString(aConnectionInfo, sFolderContentXML, oFolderDOM)
        End If
    End If

    If lErr = NO_ERR Then
        Set oFolders = oFolderDOM.selectNodes("/mi/fct/oi[@tp=""" & TYPE_FOLDER & """]")
        Set oItems   = oFolderDOM.selectNodes("/mi/fct/oi[@tp=""" & TYPE_QUESTION & """]")
        Set oParent  = oFolderDOM.selectSingleNode("//fd[../a/fd[@id='" & sFolderId & "']]")

        If oItems.length > 0 Then

            Redim sIds(oItems.length -1)
            For i = 0 To oItems.length - 1
                sIds(i) = oItems.item(i).getAttribute("id")
            Next

            sId = selectDefaultValue(sIds, Array(aSvcConfigInfo(SVCCFG_AQ_ID)))

        End If
    End If

    If lErr = NO_ERR Then
        sFolderLink = "services_select_qo.asp?" & CreateRequestForSvcConfig(aSvcConfigInfo)

        aSvcConfigInfo(SVCCFG_AQ_ID) = ""
        aSvcConfigInfo(SVCCFG_AQ_NAME) = ""
        aSvcConfigInfo(SVCCFG_AQ_PARENT_ID) = ""
    End If

%>
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>

<!-- #include file="../NSStyleSheet.asp" -->

</HEAD>
<BODY BGCOLOR="FFFFFF" TOPMARGIN=0 LEFTMARGIN=0 ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT=0 MARGINWIDTH=0>
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" HEIGHT="100%">
  <TR>
    <TD COLSPAN="6" HEIGHT="1%">
      <!-- begin header -->
        <!-- #include file="admin_header.asp" -->
      <!-- end header -->
    </TD>
  </TR>
  <TR>
    <TD WIDTH="1%" valign="TOP">
      <!-- begin toolbar -->
        <!-- #include file="_toolbar_services.asp" -->
      <!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
      <%If lErr <> 0 Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & "Over view" , "services_overview.asp") 'Descriptor: Return to: 'Descriptor:Overview %>
      <%Else%>
        <BR />
        <% Call RenderSvcConfigPath(aSvcConfigInfo) %>
        <BR />
        <BR />
        <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
          <%Response.Write(asDescriptors(747)) 'Page-by will not be shown to subscribers. However, you may select an alternative question below that will be shown to subscribers.%><BR />
        </FONT>
        <BR />

        <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
          <%Response.Write(asDescriptors(748)) 'Please select a Question Object below.%><BR />
        </FONT>
        <BR />

        <FORM ACTION="services_select_map.asp" METHOD-"POST">
        <%Call RenderSvcConfigInputs(aSvcConfigInfo)%>

        <!-- begin folder content -->
        <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
          <TR>
            <TD>
              <%Call RenderFolderPath(oFolderDOM, sFolderLink, "fid") %>
            </TD>
          </TR>
          <TR>
            <TD>
              <BR />
            </TD>
          </TR>
          <TR>
            <TD>
              <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="100%">
                <TR>
                  <TD COLSPAN=13 BGCOLOR="#99ccff"><IMG SRC="images/1ptrans.gif" HEIGHT="1" WIDTH="1" BORDER="0" ALT="" /></TD>
                </TR>
                <TR BGCOLOR="#6699cc">
                  <TD WIDTH=16><IMG SRC="../images/1ptrans.gif" WIDTH=16></TD>
                  <TD WIDTH=16><IMG SRC="../images/1ptrans.gif" WIDTH=16></TD>
                  <TD><IMG SRC="../images/1ptrans.gif" WIDTH=2></TD>
                  <TD><font face="<%=aFontInfo(S_FAMILY_FONT)%>" color="#ffffff" size="<%=aFontInfo(N_SMALL_FONT)%>"><b><%Call Response.Write(asDescriptors(306)) 'Descriptor: Name%></b></font></TD>
                  <TD>&nbsp;&nbsp;</TD>
                  <TD>&nbsp;&nbsp;</TD>
                  <TD><font face="<%=aFontInfo(S_FAMILY_FONT)%>" color="#ffffff" size="<%=aFontInfo(N_SMALL_FONT)%>"><b><%Call Response.Write(asDescriptors(34)) 'Descriptor: Modified%></b></font></TD>
                  <TD>&nbsp;&nbsp;</TD>
                  <TD>&nbsp;&nbsp;</TD>
                  <TD><font face="<%=aFontInfo(S_FAMILY_FONT)%>" color="#ffffff" size="<%=aFontInfo(N_SMALL_FONT)%>"><b><%Call Response.Write(asDescriptors(22)) 'Descriptor: Description%></b></font></TD>
                  <TD>&nbsp;&nbsp;</TD>
                </TR>

                <TR>
                  <TD COLSPAN=13 BGCOLOR="#003366"><IMG SRC="images/1ptrans.gif" HEIGHT="1" WIDTH="1" BORDER="0" ALT="" /></TD>
                </TR>

                <%If (oFolders.length + oItems.length) > 0 Then %>

                  <%For Each oItem in oFolders%>
                    <TR>
                      <TD COLSPAN=13 BGCOLOR="#ffffff"><IMG SRC="images/1ptrans.gif" HEIGHT="1" WIDTH="1" BORDER="0" ALT="" /></TD>
                    </TR>

                    <TR>
                      <TD WIDTH=16><IMG SRC="../images/1ptrans.gif" WIDTH=16></TD>
                      <TD WIDTH=16><A HREF="<%=sFolderLink & "&fid=" & oItem.getAttribute("id")%>"><IMG SRC="../images/folder2.gif" HEIGHT="16" WIDTH="16" BORDER="0" ALT="" /></A></TD>
                      <TD><IMG SRC="../images/1ptrans.gif" WIDTH=2></TD>
                      <TD NOWRAP><A HREF="<%=sFolderLink & "&fid=" & oItem.getAttribute("id")%>"><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" color="#000000" size="<%=aFontInfo(N_SMALL_FONT)%>"><b><%=oItem.getAttribute("n")%></b></font></A></TD>
                      <TD></TD>
                      <TD></TD>
                      <TD><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>"><%=DisplayDateAndTime(CDate(oItem.getAttribute("mdt")), "")%></FONT></TD>
                      <TD></TD>
                      <TD></TD>
                      <TD><FONT face="<%=aFontInfo(S_FAMILY_FONT)%>" size="<%=aFontInfo(N_SMALL_FONT)%>"><%=oItem.getAttribute("des")%></FONT></TD>
                      <TD></TD>
                    </TR>

                    <TR>
                      <TD COLSPAN=13 BGCOLOR="#ffffff"><IMG SRC="images/1ptrans.gif" HEIGHT="2" WIDTH="1" BORDER="0" ALT="" /></TD>
                    </TR>

                    <TR>
                      <TD COLSPAN=13 BGCOLOR="#99ccff"><IMG SRC="images/1ptrans.gif" HEIGHT="1" WIDTH="1" BORDER="0" ALT="" /></TD>
                    </TR>

                  <%Next%>

                  <%For Each oItem in oItems %>

                    <TR>
                      <TD COLSPAN=13 BGCOLOR="#ffffff"><IMG SRC="images/1ptrans.gif" HEIGHT="1" WIDTH="1" BORDER="0" ALT="" /></TD>
                    </TR>

                    <TR>
                      <TD WIDTH="16"><INPUT TYPE="radio" NAME="aid" VALUE="<%=oItem.getAttribute("id")%>" <%If sId = oItem.getAttribute("id") Then Response.Write("CHECKED")%> /></TD>
                      <TD WIDTH="16"><IMG SRC="../images/1ptrans.gif" WIDTH="16" HEIGHT="16"></TD>
                      <TD><IMG SRC="../images/1ptrans.gif" WIDTH=2></TD>
                      <TD NOWRAP><font face="<%=aFontInfo(S_FAMILY_FONT)%>" color="#000000" size="<%=aFontInfo(N_SMALL_FONT)%>"><b><%=oItem.getAttribute("n")%></b></font></TD>
                      <TD></TD>
                      <TD></TD>
                      <TD><font face="<%=aFontInfo(S_FAMILY_FONT)%>" size="<%=aFontInfo(N_SMALL_FONT)%>"><%=DisplayDateAndTime(CDate(oItem.getAttribute("mdt")), "")%></font></TD>
                      <TD></TD>
                      <TD></TD>
                      <TD><font face="<%=aFontInfo(S_FAMILY_FONT)%>" size="<%=aFontInfo(N_SMALL_FONT)%>"><%=oItem.getAttribute("des")%></font></TD>
                      <TD></TD>
                    </TR>

                    <TR>
                      <TD COLSPAN=13 BGCOLOR="#ffffff"><IMG SRC="images/1ptrans.gif" HEIGHT="2" WIDTH="1" BORDER="0" ALT="" /></TD>
                    </TR>

                    <TR>
                      <TD COLSPAN=13 BGCOLOR="#99ccff"><IMG SRC="images/1ptrans.gif" HEIGHT="1" WIDTH="1" BORDER="0" ALT="" /></TD>
                    </TR>
                  <%Next%>

                <%Else%>

                    <TR>
                      <TD COLSPAN=13 BGCOLOR="#ffffff"><IMG SRC="images/1ptrans.gif" HEIGHT="10" WIDTH="1" BORDER="0" ALT="" /></TD>
                    </TR>

                    <TR>
                      <TD ALIGN="CENTER" COLSPAN=13 BGCOLOR="#ffffff">
                        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>">
                          <B><%Response.Write(asDescriptors(837)) 'There are no valid objects in this folder%></B>
                        </FONT></TD>
                      </TD>
                    </TR>

                    <TR>
                      <TD COLSPAN=13 BGCOLOR="#ffffff"><IMG SRC="images/1ptrans.gif" HEIGHT="10" WIDTH="1" BORDER="0" ALT="" /></TD>
                    </TR>

                    <%If Not oParent Is Nothing Then %>
                    <TR>
                      <TD COLSPAN=13 BGCOLOR="#ffffff">
                        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>">
                           <B><A HREF="<%=sFolderLink & "&fid=" & oParent.getAttribute("id")%>"><%Response.Write(asDescriptors(147)) 'Back to parent folder%></A></B>
                        </FONT></TD>
                      </TD>
                    </TR>
                    <%End If%>


                <%End If%>
              </TABLE>

            </TD>
          </TR>
        </TABLE>
        <!-- end folder content -->
        <BR />


        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
          <TR>
            <TD COLSPAN="2">
              <BR />
            </TD>
          </TR>

          <TR>
            <TD ALIGN="left" NOWRAP WIDTH="1%">
              <INPUT NAME=back TYPE=SUBMIT CLASS="buttonClass" VALUE="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
            </TD>
            <TD ALIGN="left" NOWRAP WIDTH="98%">
              <%If oItems.length > 0 Then%>
              <INPUT NAME=next TYPE=SUBMIT CLASS="buttonClass" VALUE="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;
              <%End If%>
            </TD>
          </TR>
        </TABLE>
        </FORM>
      <%End If %>
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="1%">
        <!-- #include file="help_widget.asp" -->
    </TD>
  </TR>
</TABLE>
</BODY>
</HTML>
<%
	Set oFolderDOM = Nothing
	Set oFolders = Nothing
	Set oItems = Nothing
	Set oItem = Nothing
%>