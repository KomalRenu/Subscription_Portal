<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CustomLib/AddressesCuLib.asp" -->
<!-- #include file="CustomLib/DeviceTypesCuLib.asp" -->
<!-- #include file="CustomLib/FoldersCuLib.asp" -->
<!-- #include file="CommonDeclarations.asp" -->
<%
	'Check if user is logged in.  If not, send user to login page.
	If Len(LoggedInStatus()) = 0 Then
		Response.Redirect "login.asp"
	End If

	Dim sDeviceTypeID
	Dim sDeviceTypesXML
	Dim sGetUserAddressesXML
	Dim sDeviceDescXML
	Dim sGetFolderContentsXML
	Dim sAddressName
	Dim sPhysicalAddress
	Dim sAction
	Dim sEditAddID
	Dim sDeviceTypeName
	Dim sDeviceTypeImage
	Dim asDTFolders()

	Dim sCategoryName
	Dim sCategoryID

	sAddressesStyle = ""
	iAddressWizardStep = 1
	Redim asDTFolders(-1, 1)

    lErr = ParseRequestForAddressWiz(oRequest, sDeviceTypeID, sAddressName, sPhysicalAddress, sAction, sEditAddID, iAddressWizardStep, sCategoryID, sCategoryName)

    If oRequest("addWizCancel").Count > 0 Then
        Response.Redirect "addresses.asp"
    End If

	If lErr = NO_ERR Then
    	lErr = ReadDeviceTypesXML(sDeviceTypesXML)
    	If lErr = NO_ERR Then
    	    lErr = GetDeviceTypesProperties_AddressWizard(sDeviceTypeID, sDeviceTypesXML, sDeviceTypeName, sDeviceTypeImage, asDTFolders)
    	    If lErr = NO_ERR Then
    	        'Check if there is only one folder
    	        If UBound(asDTFolders) = 0 Then
    	            'If so, get the folder's contents
    	            lErr = cu_GetFolderContents(asDTFolders(0, 0), sGetFolderContentsXML)
    	            If lErr = NO_ERR Then
    	                'Check to see if there are sub-folders and use those
    	                lErr = CheckForSubFolders_AddressWizard(sGetFolderContentsXML, asDTFolders)
    	            End If
    	        End If
    	    End If
    	End If
	End If

	If lErr = NO_ERR Then
	    If iAddressWizardStep = 2 Then
	        lErr = GetDeviceDescForFolder_Wizard(sCategoryID, sDeviceDescXML)
	    End If
	End If

    If oRequest("addWizBack").Count > 0 Then
        If iAddressWizardStep = 1 Then
            Response.Redirect "addresses.asp?action=" & sAction & "&devicetypeID=" & sDeviceTypeID & "&editAddID=" & sEditAddID & "&wadn=" & Server.URLEncode(sAddressName) & "&wpa=" & Server.URLEncode(sPhysicalAddress)
        ElseIf iAddressWizardStep = 2 Then
            Response.Redirect "address_wiz.asp?action=" & sAction & "&devicetypeID=" & sDeviceTypeID & "&AddressName=" & Server.URLEncode(sAddressName) & "&PhysicalAddress=" & Server.URLEncode(sPhysicalAddress) & "&editAddID=" & sEditAddID
        End If
    End If
%>
<!-- #include file="CheckError.asp" -->

<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE><%Response.Write asDescriptors(361)'Descriptor: Addresses%> - MicroStrategy Narrowcast Server</TITLE>
</HEAD>
<BODY TOPMARGIN="0" LEFTMARGIN="0" BGCOLOR="ffffff" ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT="0" MARGINWIDTH="0">
<!-- #include file="header_multi.asp" -->
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
	<TR>
		<TD WIDTH="1%" VALIGN="TOP">
			<!-- begin left menu -->
		    <TABLE BORDER="0" CELLPADDING="3" CELLSPACING="0">
				<TR>
				    <TD>
				        <!-- #include file="_toolbar_Address.asp" -->
				    </TD>
				</TR>
                <TR>
                    <TD><IMG SRC="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT=""></TD>
                </TR>
		    </TABLE>
			<!-- end left menu -->
			<IMG SRC="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD WIDTH="1%">
			<IMG SRC="images/1ptrans.gif" WIDTH="15" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD WIDTH="97%" valign="TOP">
			<!-- begin center panel -->
			<% If lErr <> NO_ERR Then %>
				<% Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(381), "addresses.asp") 'Descriptor: Back to Addresses%>
			<% Else %>
            <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
                <TR>
                    <TD VALIGN="TOP">
                        <%If iAddressWizardStep = 1 Then%>
                            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write asDescriptors(604) 'Descriptor: Select a category.%></B></FONT>
                            <BR />
                            <%If Not (Len(sAddressName) = 0 And Len(sPhysicalAddress) = 0) Then %>
								<FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(606) 'Descriptor: Click on a device category for:%></FONT>
							<%End If
                        ElseIf iAddressWizardStep = 2 Then%>
                            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write asDescriptors(605) 'Descriptor: Select a style.%></B></FONT>
                            <BR />
                            <%If Not (Len(sAddressName) = 0 And Len(sPhysicalAddress) = 0) Then %>
							    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(607) 'Descriptor: Click on a device style for:%></FONT>
							<%End If
                        End If%>
                        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
                            <%
                            'If Len(sAddressName) = 0 And Len(sPhysicalAddress) = 0 Then
                            '    Response.Write "<I>" & asDescriptors(367) & "</I>" 'Descriptor: Address
                            'Else
                            If Not (Len(sAddressName) = 0 And Len(sPhysicalAddress) = 0) Then
                                If Len(sAddressName) > 0 Then
                                    Response.Write Server.HTMLEncode(sAddressName)
                                End If
                                If Len(sPhysicalAddress) > 0 Then
                                    Response.Write " (" & Server.HTMLEncode(sPhysicalAddress) & ")"
                                End If
                            End If
                            %>
                        </FONT>
                    </TD>
                    <TD ALIGN="RIGHT"><IMG SRC="images/desktop_Addresses.gif" WIDTH="60" HEIGHT="60" BORDER="0" ALT="" /></TD>
                </TR>
            </TABLE>

            <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="400">
            <FORM METHOD="POST" ACTION="address_wiz.asp">
            <INPUT TYPE="HIDDEN" NAME="action" VALUE="<%Response.Write sAction%>" />
            <INPUT TYPE="HIDDEN" NAME="devicetypeID" VALUE="<%Response.Write sDeviceTypeID%>" />
            <INPUT TYPE="HIDDEN" NAME="awstep" VALUE="<%Response.Write iAddressWizardStep%>" />
            <INPUT TYPE="HIDDEN" NAME="AddressName" VALUE="<%Response.Write Server.HTMLEncode(sAddressName)%>" />
            <INPUT TYPE="HIDDEN" NAME="PhysicalAddress" VALUE="<%Response.Write Server.HTMLEncode(sPhysicalAddress)%>" />
            <INPUT TYPE="HIDDEN" NAME="editAddID" VALUE="<%Response.Write sEditAddID%>" />
            <TR>
                <TD COLSPAN="2">
                    <%If iAddressWizardStep = 1 Then%>
                        <TABLE BORDER="0" CELLPADDING="2" CELLSPACING="0" WIDTH="100%">
                            <TR>
                                <TD ALIGN="CENTER" NOWRAP WIDTH="1%"><B><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><%Response.Write sDeviceTypeName%></FONT></B></TD>
                                <TD WIDTH="1%"><IMG SRC="images/1ptrans.gif" HEIGHT="1" WIDTH="10" ALT="" BORDER="0" /></TD>
                                <TD WIDTH="98%"></TD>
                            </TR>
                            <TR>
                                <TD VALIGN="TOP" ALIGN="CENTER"><IMG SRC="<%Response.Write sDeviceTypeImage%>" ALT="" BORDER="0" /></TD>
                                <TD></TD>
                                <TD VALIGN="TOP">
                                    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>">
                                    <%
                                        If UBound(asDTFolders) <> -1 Then
                                            Dim i
                                            Dim iUBoundDTFolders
                                            iUBoundDTFolders = UBound(asDTFolders)
                                            For i=0 to iUBoundDTFolders
                                                Response.Write "<A HREF=""address_wiz.asp?awstep=2&devicetypeID=" & sDeviceTypeID & "&dtfn=" & Server.URLEncode(asDTFolders(i, 1)) & "&dtfid=" & asDTFolders(i, 0) & "&action=" & sAction & "&editAddID=" & sEditAddID & "&AddressName=" & Server.URLEncode(sAddressName) & "&PhysicalAddress=" & Server.URLEncode(sPhysicalAddress) & """>" & asDTFolders(i, 1) & "</A><BR />"
                                            Next
                                        End If
                                    %>
                                    </FONT>
                                </TD>
                            </TR>
                        </TABLE>
                    <%ElseIf iAddressWizardStep = 2 Then%>
                        <% Call RenderDeviceDescriptions(sDeviceDescXML, sDeviceTypeID, sAction, sEditAddID, sAddressName, sPhysicalAddress)%>
                    <%End If%>
                </TD>
            </TR>
            <TR>
                <TD COLSPAN="2"><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="10" ALT="" BORDER="0" /></TD>
            </TR>
            <TR>
                <TD COLSPAN="2" bgcolor="#cccccc"><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="1" ALT="" BORDER="0" /></TD>
            </TR>
            <TR>
                <TD COLSPAN="2"><IMG SRC="images/1ptrans.gif" WIDTH="1" HEIGHT="3" ALT="" BORDER="0" /></TD>
            </TR>
            <TR>
                <TD><INPUT TYPE="SUBMIT" NAME="addWizBack" CLASS="buttonClass" VALUE="<%Response.Write asDescriptors(149) 'Descriptor: Back%>" /></TD>
                <TD ALIGN="RIGHT"><INPUT TYPE="SUBMIT" NAME="addWizCancel" CLASS="buttonClass" VALUE="<%Response.Write asDescriptors(120) 'Descriptor: Cancel%>" /></TD>
            </TR>
            </FORM>
            </TABLE>
			<% End If %>
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
    Erase asDTFolders
%>