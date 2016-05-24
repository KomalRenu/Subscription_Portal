<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
  Option Explicit
  Response.CacheControl = "no-cache"
  Response.AddHeader "Pragma", "no-cache"
  Response.Expires = -1
  On Error Resume Next
%>
<!-- #include file="../CustomLib/AdminCuLib.asp" -->
<!-- #include file="../CustomLib/DeviceTypesCuLib.asp" -->
<!-- #include file="../CommonDeclarations.asp" -->
<%
Dim sDeviceTypesXML
Dim sFolderID
Dim sGetFolderContentsXML
Dim lStatus
Dim aDeviceTypeInfo

    lErr = ParseRequestForDTFolders(oRequest, aDeviceTypeInfo, sFolderID)

    If oRequest("back").count > 0 Then
        lErr = ReadDeviceTypesXML(sDeviceTypesXML)
        If lErr = NO_ERR Then
            lErr = DeleteTempDFS(sDeviceTypesXML, aDeviceTypeInfo(DEV_TYPE_ID))
            If lErr = NO_ERR Then
                lErr = WriteDeviceTypesXML(sDeviceTypesXML)
                If lErr = NO_ERR Then
                    Response.Redirect "editDeviceType.asp?action=edit&dtID=" & aDeviceTypeInfo(DEV_TYPE_ID)
                End If
            End If
        End If
    End If

    If lErr = NO_ERR Then
        lErr = ReadDeviceTypesXML(sDeviceTypesXML)
        If lErr = NO_ERR Then
            lErr = GetVariablesFromXML_DTFolders(aDeviceTypeInfo, sDeviceTypesXML)
        End If
    End If

    If lErr = NO_ERR Then
        If Request.Form.Count > 0 Then
            If Len(CStr(oRequest("DTFolderAdd.x"))) > 0 Then
                lErr = AddFoldersToDT(oRequest, aDeviceTypeInfo(DEV_TYPE_ID), sDeviceTypesXML)
                If lErr = NO_ERR Then
                    lErr = WriteDeviceTypesXML(sDeviceTypesXML)
                End If
            ElseIf Len(CStr(oRequest("DTFolderRemove.x"))) > 0 Then
                lErr = RemoveFoldersFromDT(oRequest, aDeviceTypeInfo(DEV_TYPE_ID), sDeviceTypesXML)
                If lErr = NO_ERR Then
                    lErr = WriteDeviceTypesXML(sDeviceTypesXML)
                End If
            Else 'This is the SAVE case
                lErr = SaveDTFolders(aDeviceTypeInfo(DEV_TYPE_ID), sDeviceTypesXML)
                If lErr = NO_ERR Then
                    lErr = WriteDeviceTypesXML(sDeviceTypesXML)
                    If lErr = NO_ERR Then
                        Response.Redirect "deviceTypes.asp"
                    End If
                End If
            End If
        End If
    End If

    If lErr = NO_ERR Then
        lErr = cu_GetFolderContents(sFolderID, sGetFolderContentsXML)
        If lErr = NO_ERR Then
            lErr = LoadTempDFS(sDeviceTypesXML, aDeviceTypeInfo(DEV_TYPE_ID))
            If lErr = NO_ERR Then
                lErr = WriteDeviceTypesXML(sDeviceTypesXML)
            End If
        End If
    End If

    'Get the Channels list request from the request object:
    aPageInfo(S_NAME_PAGE) = "deviceTypeFolders.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_DEVICE_TYPES_FOLDER & " " & asDescriptors(496) & ": " & Server.HTMLEncode(aDeviceTypeInfo(DEV_TYPE_NAME)) 'Descriptor: Device Folders
    aPageInfo(N_CURRENT_OPTION_PAGE) = 3
    aPageInfo(N_OPTIONS_WITH_LINKS_PAGE) = "folderID=" & sFolderID & "&" & CreateRequestForDeviceType(aDeviceTypeInfo)

    lStatus = checkSiteConfiguration()

%>
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>

<!-- #include file="../NSStyleSheet.asp" -->

</HEAD>
<SCRIPT LANGUAGE="JavaScript">
<!--
var bLinkEnabled = true;
function test(){
    if (bLinkEnabled){
        bLinkEnabled = false;
        return true;
    } else {
        return false;
    }
}
-->
</SCRIPT>
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
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(493), "deviceTypes.asp") 'Descriptor:Device Types%>
      <%Else%>
        <FORM METHOD="POST" ACTION="deviceTypeFolders.asp" NAME="DevTypesFolders">
        <INPUT TYPE="HIDDEN" NAME="dtID" VALUE="<%Response.Write aDeviceTypeInfo(DEV_TYPE_ID)%>" />

        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
          <TR>
            <TD>
              <BR />
            </TD>
          </TR>
          <TR>
            <TD>
              <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>">
				<%Response.Write asDescriptors(672) & " " 'Descriptor:This screen allows you to add or remove elements from the folder for this specific device type.%>
				<%Response.Write asDescriptors(673) 'Descriptor:Elements on the left are available, but not active on this device type.  Elements on the right are currently selected for this specific device type.%>
		      <P></FONT>
		    </TD>
	      </TR>
	    </TABLE>


        <TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0>
          <TR>
            <TD BGCOLOR="#cccccc">
              <TABLE BORDER=0 CELLPADDING=3 CELLSPACING=0>
                <TR>
                  <TD BGCOLOR="#ffffff">
                    <TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0>
                      <TR>
                        <TD>
                          <% Call RenderPath_DeviceFolders(aDeviceTypeInfo(DEV_TYPE_ID), sGetFolderContentsXML) %>
                        </TD>
                      </TR>

                      <TR>
                        <TD>
                          <TABLE BORDER=0 CELLPADDING=3 CELLSPACING=0>
                            <TR>
                              <TD VALIGN=TOP>
                                <font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" size="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(517) 'Descriptor: Available:%></font>
                                <TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0>
                                  <TR>
                                    <TD BGCOLOR="#cccccc">
                                      <TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
                                        <TR>
                                          <TD BGCOLOR="#ffffff">
                                            <TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0>
                                              <TR>
                                                <TD>
                                                  <% Call RenderAvailableDTFolders(sDeviceTypesXML, aDeviceTypeInfo(DEV_TYPE_ID), sGetFolderContentsXML) %>
                                                </TD>
                                              </TR>

                                              <TR>
                                                <TD><IMG SRC="../images/1ptrans.gif" HEIGHT="25" WIDTH="100" ALT="" BORDER="0" /></TD>
                                              </TR>
                                            </TABLE>
                                          </TD>
                                        </TR>
                                      </TABLE>
                                    </TD>
                                  </TR>
                                </TABLE>
                              </TD>

                              <TD>
                                <BR /><BR />
                                <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
                                  <TR>
                                    <TD><INPUT onClick="return test();" TYPE="IMAGE" NAME="DTFolderAdd" SRC="../images/btn_add.gif" BORDER="0" ALT="<%Response.Write asDescriptors(207) 'Descriptor: Add%>" /></TD>
                                  </TR>
                                  <TR>
                                    <TD><INPUT onClick="return test();" TYPE="IMAGE" NAME="DTFolderRemove" SRC="../images/btn_remove.gif" BORDER="0" ALT="<%Response.Write asDescriptors(106) 'Descriptor: Remove%>" /></TD>
                                  </TR>
                                </TABLE>

                                <BR /><BR /><BR />
                              </TD>

                              <TD VALIGN=TOP>
                                <font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" size="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(518) 'Descriptor: Selected:%></font>
                                <TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0>
                                  <TR>
                                    <TD BGCOLOR="#cccccc">
                                      <TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
                                        <TR>
                                          <TD BGCOLOR="#ffffff">
                                            <TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0>
                                              <TR>
                                                <TD>
                                                  <% Call RenderSelectedDTFolders(sDeviceTypesXML, aDeviceTypeInfo(DEV_TYPE_ID)) %>
                                                </TD>
                                              </TR>

                                              <TR>
                                                <TD><IMG SRC="../images/1ptrans.gif" HEIGHT="25" WIDTH="100" ALT="" BORDER="0" /></TD>
                                              </TR>
                                            </TABLE>
                                          </TD>
                                        </TR>
                                      </TABLE>
                                    </TD>
                                  </TR>
                                </TABLE>
                              </TD>
                            </TR>
                          </TABLE>
                        </TD>
                      </TR>
                    </TABLE>
                  </TD>
                </TR>
              </TABLE>
            </TD>
          </TR>
        </TABLE>

        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
          <TR>
            <TD>
              <BR />
            </TD>
          </TR>

          <TR>
            <TD ALIGN="left" NOWRAP>
              <INPUT name=back type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
              <INPUT name=next type=submit class="buttonClass" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;
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