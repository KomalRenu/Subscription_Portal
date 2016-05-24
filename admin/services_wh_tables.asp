<%
  Option Explicit
  Response.CacheControl = "no-cache"
  Response.AddHeader "Pragma", "no-cache"
  Response.Expires = -1
  On Error Resume Next
%>
<!-- #include file="../CustomLib/ErrorCuLib.asp" -->
<!-- #include file="../CommonDeclarations.asp" -->
<!-- #include file="../CustomLib/AdminCuLib.asp" -->
<!-- #include file="../CustomLib/CommonCuLib.asp" -->
<!-- #include file="../CoreLib/SiteConfigCoLib.asp" -->
<%
Dim lStatus

Dim aDBAliases
Dim aDBList
Dim i

Dim aTables

Dim sMapDBAlias
Dim aMapTables
Dim aMapColumns


    If lErr = NO_ERR Then
        lErr = ParseRequestForSvcConfig(oRequest, aSvcConfigInfo)
    End If

    'Get a list of DBAliases from the Engine:
    If lErr = NO_ERR Then
        lErr = getDBAliases(aDBAliases)

        If Not IsEmpty(aDBAliases) Then
            Redim aDBList(UBound(aDBAliases), 1)

            For i = 0 to UBound(aDBAliases)
                aDBList(i, 0) = aDBAliases(i,0)

                If Len(aDBAliases(i,1)) > 0 Then
					aDBList(i, 1) = aDBAliases(i,1)
				Else
	                aDBList(i, 1) = aDBAliases(i,0)
	            End If
            Next
        End If

    End If

    'Get the previous values:
    If lErr = NO_ERR Then
        lErr = co_getFolderXML(sFolderId, ROOT_APP_FOLDER_TYPE, sFolderContentXML)
    End If

    sFolderLink = "services_select_qo.asp?id=" & sId

    'Set the PageInfo to be used by the navigator bar and the header.
    aPageInfo(S_TITLE_PAGE) = STEP_SERVICES_QO_WH_TABLES & " " & "Warehouse Storage"
    aPageInfo(N_CURRENT_OPTION_PAGE) = 3
    aPageInfo(N_OPTIONS_WITH_LINKS_PAGE) = CreateRequestForSvcConfig(aSvcConfigInfo)

    lStatus = checkSiteConfiguration()


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
        <!-- #include file="_toolbar_site_preferences.asp" -->
      <!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
      <%If lErr <> 0 Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(579) , "devices_config.asp") 'Descriptor: Return to: 'Descriptor:Site Preferences %>
      <%Else%>
        <BR />
        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" >
          Service: <B><%=sName%></B><BR />
          Publication: <B><%=sPubName%></B><BR />
          <BR />
         </FONT>
        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
          You have choosen to display this Question Object:<BR />
           &nbsp;<B><%=sNewQOName%></B><BR />
          in place of this slicing Question:<BR/>
           &nbsp;<B><%=sQOName%></B><BR />
          <BR />
        </FONT>

        <TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH="100%">
          <TR>
            <TD>
            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
              <B>Warehouse Storage</B>
            </FONT>
            </TD>
          </TR>

          <TR>
            <TD BGCOLOR="#c2c2c2" HEIGHT="1"><IMG SRC="../images/1ptrans.gif" HEIGHT="1" WIDTH="1"></TD>
          </TR>

          <TR>
            <TD>
            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
              Please select the database tables below where you wisth to store the answers to: <%=sNewQOName%>
            </FONT>
            </TD>
          </TR>

          <TR>
            <TD HEIGHT="15"><IMG SRC="../images/1ptrans.gif" HEIGHT="15" WIDTH="1"></TD>
          </TR>

          <TR>
            <TD>
              <TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
                <TR>
                  <TD>
                    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                      Data Warehouse:&nbsp;
                    </FONT>
                  </TD>
                  <TD><%Call RenderDropDownList("dw", aDBList, "Hydra", "")%>
                  <TD>
                    &nbsp;<INPUT TYPE="SUBMIT" NAME="refresh" class="buttonclass" VALUE="<%Response.Write(asDescriptors(269))'Descriptor:Refresh%>" />&nbsp;
                  </TD>
                </TR>
              </TABLE>
            </TD>
          </TR>
        </TABLE>

        <FORM ACTION="services_wh_columns.asp">
        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
          <TR>
            <TD>
              <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                Available Tables:
              </FONT>
            </TD>
            <TD></TD>
            <TD>
              <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                Selected Tables:
              </FONT>
            </TD>
          </TR>

          <TR>
            <TD>
              <SELECT STYLE="width:200" NAME=avail SIZE="10" class="pullDownClass" WIDTH="40">
                <OPTION VALUE="t3">Table 3
                <OPTION VALUE="t4">Table 4
                <OPTION VALUE="t5">Table 5
                <OPTION VALUE="t6">Table 6
                <OPTION VALUE="t7">Table 7
              </SELECT>
            </TD>
            <TD WIDTH="29" VALIGN="MIDDLE" ALIGN="CENTER">
              <IMG SRC="../images/btn_add.gif" WIDTH="25" HEIGHT="25" /><BR />
              <BR />
              <IMG SRC="../images/btn_remove.gif" WIDTH="25" HEIGHT="25" />
            </TD>
            <TD>
              <SELECT STYLE="width:200" NAME=tbls SIZE="10" class="pullDownClass" WIDTH="40">
                <OPTION VALUE="t1">Table 1
                <OPTION VALUE="t2">Table 2
              </SELECT>
            </TD>
          </TR>

        </TABLE>

        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
          <TR>
            <TD COLSPAN="2">
              <BR />
            </TD>
          </TR>

          <TR>
            <TD ALIGN="left" NOWRAP WIDTH="1%">
              <INPUT name=back type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
            </TD>
            <TD ALIGN="left" NOWRAP WIDTH="98%">
              <INPUT name=next type=submit class="buttonClass" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;
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
	Erase aDBAliases
	Erase aDBList
	Erase aTables
	Erase aMapTables
	Erase aMapColumns
%>