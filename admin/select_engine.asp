<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Option Explicit
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
On Error Resume Next
%>
<!-- #include file="../CommonDeclarations.asp" -->
<!-- #include file="../CustomLib/AdminCuLib.asp" -->
<!-- #include file="../CustomLib/SiteConfigCuLib.asp" -->
<%
Dim lStatus
Dim sEngine
Dim aRUEngines()

    aPageInfo(S_NAME_PAGE) = "select_engine.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_SELECT_ENGINE & " " & asDescriptors(582) 'Descriptor:Subscription Engine Location"
    aPageInfo(N_CURRENT_OPTION_PAGE) = 1

    lStatus = checkSiteConfiguration()

    If lErr = NO_ERR Then
        lErr = getMRUEngines(aRUEngines)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, "", "", "select_engine.asp", "", "", "Error calling getMRUEngines", LogLevelTrace)
    End If

    If lErr = NO_ERR Then
        sEngine = selectDefaultEngine(aRUEngines)
    End If

%>
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--

  function validateForm() {
  var sMsg

    sMsg = "";
    if ((FormEngine.se.value == "") || isBlank(FormEngine.se.value)) {
      <%Call Response.Write("sMsg += ""<LI>" & asDescriptors(598) & """;") 'Descriptor:Please provide the name of the machine running the Subscription Engine %>
    }

    if (sMsg != "") {
      if(document.all){
         document.all("validation").innerHTML = sMsg;
         document.all("validation").style.display = "block";
      }
      return false;
    }
  }
//-->
</SCRIPT>

<!-- #include file="validationJS.asp" -->
<!-- #include file="../NSStyleSheet.asp" -->

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
        <!-- #include file="_toolbar_engine_config.asp" -->
      <!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
      <%If lErr <> 0 Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(283), "welcome.asp") 'Descriptor: Return to: 'Descriptor:Welcome%>
      <%Else%>
      <BR />
      <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" >
        <%Response.Write(asDescriptors(658)) 'A typical installation will install the subscription engine on the same machine (i.e., localhost).%>&nbsp;
        <%Response.Write(asDescriptors(816)) 'If it has been installed on a separate machine enter the name of that machine here.%>
        <BR />
      </FONT>

      <BR />
      <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>"  COLOR="#ff0000"><DIV STYLE="display:none;" class="validation" id="validation"></DIV></FONT>

      <FORM NAME="FormEngine" ACTION="modify_engine.asp" METHOD="POST">
      <TABLE BORDER="0" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
        <TR>
          <TD VALIGN="MIDDLE">
            <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>">
              <%Response.Write(asDescriptors(565)) 'Descriptor:Enter the location of the Subscription Engine:%>
            </FONT>
           <INPUT NAME=se class="textBoxClass" SIZE=40 VALUE="<%=Server.HTMLEncode(sEngine)%>"></INPUT>
          </TD>
        </TR>
        <TR>
          <TD COLSPAN=2>
            <BR />
            <!-- Start MRU List -->
            <%RenderRUEngines(aRUEngines)%>
            <!-- End MRU List -->
          </TD>
        </TR>
      </TABLE>

      <TABLE BORDER="0" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
        <TR>
          <TD>
            <BR />
          </TD>

        <TR>
          <TD ALIGN="left" NOWRAP>
            <INPUT name=back   type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
            <INPUT name=next  onClick="return validateForm();" type=submit class="buttonClass" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;
          </TD>
        </TR>
       </TABLE>

       </FORM>
       <%End If%>

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
	Erase aRUEngines
%>