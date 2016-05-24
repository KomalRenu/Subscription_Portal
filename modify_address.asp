<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CustomLib/ModifyAddressCuLib.asp" -->
<!-- #include file="CommonDeclarations.asp" -->
<%
	'Check if user is logged in.  If not, send user to login page.
	If LoggedInStatus() = "" Then
		Response.Redirect "login.asp"
	End If

	Dim sAction
	Dim sDeviceTypeID
	Dim sDelAddrID
	Dim sEditAddID
	Dim sAddressName
	Dim sPhysicalAddress
	Dim sDevice
	Dim sPIN
	Dim sCallBlock
	Dim sTransPropsID
	Dim sAddressNameVld
	Dim sPhysicalAddressVld
	Dim sPINVld

	sAddressesStyle = ""

	If oRequest("addrCancel") <> "" Then
		Response.Redirect "addresses.asp" 'Go back to address page.
	End If

	lErr = ParseRequestForModifyAddress(oRequest, sAction, sDeviceTypeID, sAddressName, sPhysicalAddress, sDevice, sPIN, sCallBlock, sDelAddrID, sEditAddID, sAddressNameVld, sPhysicalAddressVld, sPINVld, sTransPropsID)

	If lErr = NO_ERR Then
		If sAction = "add" Then
			lValidationError = validate_AddressFields(sAddressName, sPhysicalAddress, sAddressNameVld, sPhysicalAddressVld)
			If lValidationError = NO_ERR Then
			    lErr = cu_AddAddress(sAddressName, sPhysicalAddress, sDevice, sPIN, sCallBlock)
			End If
		ElseIf sAction = "edit" Then
			lValidationError = validate_AddressFields(sAddressName, sPhysicalAddress, sAddressNameVld, sPhysicalAddressVld)
			If lValidationError = NO_ERR Then
			    lErr = cu_EditAddress(sEditAddID, sAddressName, sPhysicalAddress, sDevice, sCallBlock, sPIN, sTransPropsID)
			End If
		ElseIf sAction = "delete" Then
			lErr = cu_DeleteAddress(sDelAddrID)
		End If
	End If

	If lErr = NO_ERR And lValidationError = NO_ERR Then
		Response.Redirect "addresses.asp" 'Go back to address page.
	End If

	sErrorHeader = asDescriptors(387) 'Descriptor: Error during address operation
	If lValidationError <> NO_ERR Then
	    sErrorMessage = ""
	    If (lValidationError And ERR_ADDRESS_BLANKS) Then
	        sErrorMessage = sErrorMessage & "<LI>" & asDescriptors(389) & "</LI>" 'Descriptor: One or more address fields were blank.  Please try again.
	    End If
        If (lValidationError And ERR_ADDR_NAME_INVALID) Then
            sErrorMessage = sErrorMessage & "<LI>" & asDescriptors(417) & " " 'Descriptor: Please enter an address name without the following characters:
            Dim i
            Dim sChars
            sChars = ""
	    	For i = 0 to Ubound(asReservedChars)
	    		sChars = sChars & asReservedChars(i) & " "
	    	Next
	    	sChars = Left(sChars, Len(sChars) - 1)
	    	sErrorMessage = sErrorMessage & sChars & "</LI>"
        End If
        If (lValidationError And ERR_EMAIL_ADDR_INVALID) Then
            sErrorMessage = sErrorMessage & "<LI>" & asDescriptors(419) & "</LI>" 'Descriptor: Please enter an address in the form of: user@server.com
        End If
        If (lValidationError And ERR_NUMBER_ADDR_INVALID) Then
            sErrorMessage = sErrorMessage & "<LI>" & asDescriptors(529) & "</LI>" 'Descriptor: Please enter a numeric value for the address in the following form: #########
        End If
	Else
	    Select Case lErr
	    	Case ERR_XML_LOAD_FAILED
	    		sErrorHeader = asDescriptors(427) 'Descriptor: Error retrieving data
	    		sErrorMessage = asDescriptors(428) 'Descriptor: There was an error while loading XML.  Please contact the system administrator.
	    	Case URL_MISSING_PARAMETER
	    		sErrorHeader = asDescriptors(325) 'Descriptor: Wrong parameters in the URL
	    		sErrorMessage = asDescriptors(326) & " action" 'Descriptor: The following parameters are required in the URL:
	    	Case ERR_ADDR_OPERATION
	    		sErrorMessage = asDescriptors(388) 'Descriptor: Please try again or contact your system administrator
	    	Case Else
	    End Select
	End If

%>

<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE></TITLE>
</HEAD>
<BODY BGCOLOR="ffffff" TOPMARGIN=0 LEFTMARGIN=0 ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT=0 MARGINWIDTH=0>
<!-- #include file="header_multi.asp" -->
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=100%>
	<TR>
		<TD WIDTH="1%" valign="TOP">
			<!-- begin search box -->
				<!-- #include file="searchbox.asp" -->
			<!-- end search box -->
			<!-- begin left menu -->
			<BR>

			<BR>
			<!-- end left menu -->
			<img src="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD WIDTH="98%" valign="TOP">
			<!-- begin center panel -->
			<BR>
			<% If lErr <> NO_ERR Then %>
				<% Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(381), "addresses.asp") 'Descriptor: Back to Addresses%>
		    <% ElseIf lValidationError <> NO_ERR Then %>
		        <% Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(381), "addresses.asp") 'Descriptor: Back to Addresses%>
			<% Else %>

			<% End If %>
			<!-- end center panel -->
		</TD>
		<TD WIDTH="1%">
			<img src="images/1ptrans.gif" WIDTH="15" HEIGHT="1" BORDER="0" ALT="">
		</TD>
	</TR>
</TABLE>
<BR>
<!-- begin footer -->
	<!-- #include file="footer.asp" -->
<!-- end footer -->
</BODY>
</HTML>