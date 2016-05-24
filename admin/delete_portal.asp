<%
  Option Explicit
  Response.CacheControl = "no-cache"
  Response.AddHeader "Pragma", "no-cache"
  Response.Expires = -1
  On Error Resume Next
%>
<!-- #include file="../CommonDeclarations.asp" -->
<!-- #include file="../CustomLib/AdminCuLib.asp" -->
<!-- #include file="../CustomLib/PortalConfigCuLib.asp" -->
<%
Dim lStatus

Dim sName
Dim sConfirmed
Dim sVirtualDirectory

Dim aPortals
Dim nCount
Dim i

    'Check for actions cancelled:
    If Len(CStr(oRequest("cancel"))) > 0 Then
        Call Response.Redirect("select_portal.asp")
    End If
    
    sName = oRequest("name")
    sConfirmed = oRequest("confirm")
    sVirtualDirectory = GetVirtualDirectoryName()

    
    'If no given name so far for the site:
    If lErr = NO_ERR Then
    
        'If confirmed, delete the site:
        If sConfirmed = "yes" Then
        
            'If we're on the same VD, move to a different one:
            If StrComp(sName, sVirtualDirectory, vbTextCompare) = 0 Then
                lErr = cu_GetAllPortals(aPortals, nCount)
                
                If lErr = NO_ERR Then 
                    For i = 0 To nCount - 1
                        If StrComp(aPortals(i, 6), sVirtualDirectory) <> 0 Then
                            Call Response.Redirect("../../" & Server.URLEncode(aPortals(i, 6)) & "/admin/delete_portal.asp?confirm=yes&name=" & sName)
                        End If
                    Next
                End If
                    
            Else
                lErr = deletePortal(sName)
                        
                'If everything went fine, redirect to the select_Portal page again:
                If lErr = NO_ERR Then 
                    Call Response.Redirect("select_portal.asp")
                End If
            End If
                                
        End If

    End If
    
    aPageInfo(S_TITLE_PAGE) = STEP_SELECT_PORTAL & " " & asDescriptors(252)'Descriptor:Delete Confirmation
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_PORTAL_MANAGEMENT
    aPageInfo(N_OPTIONS_WITH_LINKS_PAGE) = "name=" & sName
    

%>
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>

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
        <!-- #include file="_toolbar_portal_management.asp" -->
      <!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
      <%If lErr <> 0 Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(628) , "select_portal.asp") 'Descriptor: Return to:'Descriptor:Portal Management %>
      <%Else%>
      <BR />
      <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>"  COLOR="#ff0000"><DIV STYLE="display:none;" class="validation" id="validation"></DIV></FONT>
      
      <TABLE BORDER="0" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
        <FORM ACTION="delete_portal.asp" METHOD="POST">
          <TR>
            <TD>
              <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>">
                <%Response.write(asDescriptors(701)) 'Descriptor:Are you sure you want to delete this portal%> <B><%=GerPortalName(sName)%></B>
              </FONT>
            </TD>
          </TR>
          <TR>
            <TD ALIGN=CENTER>  
              <BR />
              <INPUT name=name    type=HIDDEN value="<%=sName%>"></INPUT>
              <INPUT name=confirm type=HIDDEN value="yes"   ></INPUT>
              <INPUT name=ok      type=submit class="buttonClass" value="<%Response.Write(asDescriptors(543)) 'Descriptor:Ok%>"></INPUT> &nbsp;
              <INPUT name=cancel  type=submit class="buttonClass" value="<%Response.Write(asDescriptors(120)) 'Descriptor:Cancel%>"></INPUT>
            </TD>
          </TR>
        </FORM>
      </TABLE>
         
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