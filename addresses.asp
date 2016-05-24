<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CustomLib/AddressesCuLib.asp" -->
<!-- #include file="CustomLib/ModifyAddressCuLib.asp" -->
<!-- #include file="CustomLib/DeviceTypesCuLib.asp" -->
<!-- #include file="CommonDeclarations.asp" -->
<%
	'Check if user is logged in.  If not, send user to login page.
	If Len(LoggedInStatus()) = 0 Then
		Response.Redirect "login.asp"
	End If

	Dim sAction
	Dim sDeviceTypeID
	Dim sDeviceTypesXML
	Dim sEditAddID
	Dim sGetUserAddressesXML
	Dim sGetDevicesFromFoldersXML

	Dim sAddressName
	Dim sPhysicalAddress
	Dim sAddressNameVld
	Dim sPhysicalAddressVld
	Dim sDevice
	Dim sPIN
	Dim sCallBlock
	Dim sDelAddrID
	Dim sTransPropsID
	Dim sWizardDeviceID
	Dim sWizardAddressName
	Dim sWizardPhysicalAddress

	sAddressesStyle = ""

	lErr = ParseRequestForAddresses(oRequest, sDeviceTypeID, sEditAddID, sAction, sAddressName, sPhysicalAddress, sAddressNameVld, sPhysicalAddressVld, sDevice, sPIN, sCallBlock, sDelAddrID, sTransPropsID, sWizardDeviceID, sWizardAddressName, sWizardPhysicalAddress)

	If oRequest("addrCancel").Count > 0 Then
		Call CleanRequestForAddresses(sDeviceTypeID, sEditAddID, sAction, sAddressName, sPhysicalAddress, sAddressNameVld, sPhysicalAddressVld, sDevice, sPIN, sCallBlock, sDelAddrID, sTransPropsID, sWizardDeviceID, sWizardAddressName, sWizardPhysicalAddress)
	End If

    If oRequest("addrSave").Count > 0 Then
        If StrComp(sAction, "add", vbBinaryCompare) = 0 Then
			lValidationError = validate_AddressFields(sAddressName, sPhysicalAddress, sAddressNameVld, sPhysicalAddressVld)
			If lValidationError = NO_ERR Then
			    lErr = cu_AddAddress(sAddressName, sPhysicalAddress, sDevice, sPIN, sCallBlock)
			End If
        ElseIf StrComp(sAction, "edit", vbBinaryCompare) = 0 Then
			lValidationError = validate_AddressFields(sAddressName, sPhysicalAddress, sAddressNameVld, sPhysicalAddressVld)
			If lValidationError = NO_ERR Then
			    lErr = cu_EditAddress(sEditAddID, sAddressName, sPhysicalAddress, sDevice, sCallBlock, sPIN, sTransPropsID)
			End If
        End If

        If lErr = NO_ERR And lValidationError = NO_ERR Then
		    Call CleanRequestForAddresses(sDeviceTypeID, sEditAddID, sAction, sAddressName, sPhysicalAddress, sAddressNameVld, sPhysicalAddressVld, sDevice, sPIN, sCallBlock, sDelAddrID, sTransPropsID, sWizardDeviceID, sWizardAddressName, sWizardPhysicalAddress)
	    End If
    ElseIf oRequest("addrHelp.x").Count > 0 Then
        Response.Redirect "address_wiz.asp?devicetypeID=" & sDeviceTypeID & "&AddressName=" & Server.URLEncode(sAddressName) & "&PhysicalAddress=" & Server.URLEncode(sPhysicalAddress) & "&action=" & sAction & "&editAddID=" & sEditAddID
    End If

    If StrComp(sAction, "delete", vbBinaryCompare) = 0 Then
        lErr = cu_GetUserAddresses(sGetUserAddressesXML)

	    If lErr = NO_ERR Then
            lErr = cu_DeleteAddress(sDelAddrID, sGetUserAddressesXML)
        End If

        If lErr = NO_ERR Then
            Call CleanRequestForAddresses(sDeviceTypeID, sEditAddID, sAction, sAddressName, sPhysicalAddress, sAddressNameVld, sPhysicalAddressVld, sDevice, sPIN, sCallBlock, sDelAddrID, sTransPropsID, sWizardDeviceID, sWizardAddressName, sWizardPhysicalAddress)
        End If
    End If

	If lErr = NO_ERR Then
		lErr = ReadDeviceTypesXML(sDeviceTypesXML)
	End If

	If lErr = NO_ERR Then
	    lErr = cu_GetDevicesInFolders(sDeviceTypesXML, sGetDevicesFromFoldersXML)
	End If

	If lErr = NO_ERR Then
		lErr = cu_GetUserAddresses(sGetUserAddressesXML)
	End If


%>
<!-- #include file="CheckError.asp" -->

<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE><%Response.Write asDescriptors(361)'Descriptor: Addresses%> - MicroStrategy Narrowcast Server</TITLE>

<%If StrComp(GetJavaScriptSetting(), "1", vbBinaryCompare) = 0 Then%>
<SCRIPT LANGUAGE="JAVASCRIPT">
<!--
  var bValidate = false;
  var bCheckPin = false;

  function validateAddressForm(addressForm)	{

    if (bValidate == false){
      return;
    }

    var msg = "";
    if ((addressForm.AddressName.value == "") || isBlank(addressForm.AddressName.value)) {
      msg = msg + "- <%Response.Write asDescriptors(416)'Descriptor: Please enter an address name%>\n";
    }

    if ((addressForm.PhysicalAddress.value == "") || isBlank(addressForm.PhysicalAddress.value)) {
      msg += "- <%Response.Write asDescriptors(418)'Descriptor: Please enter an address%>\n";
    } else {
      if (addressForm.PhysicalAddressVld.value == "email"){
        if (checkInvalidCharacters(addressForm.PhysicalAddress.value) == false){
      	  msg += "- <%Response.Write asDescriptors(417) & " "'Descriptor: Please enter an address name without the following characters:%>" + invalidChars() + "\n";
       } else {
          if (checkEmailFormat(addressForm.PhysicalAddress.value) == false){
            msg += "- <%Response.Write asDescriptors(419)'Descriptor: Please enter an address in the form of: user@server.com%>\n";
          }
    	  }
      } else {
        if(addressForm.PhysicalAddressVld.value == "number") {
          if (checkNumericFormat(addressForm.PhysicalAddress.value) == false){
            msg += "- <%Response.Write asDescriptors(614) 'Descriptor: Please enter a value for the address in the following form: any numbers and the following characters - ( )%>\n";
    	    }
    	  }
    	}
    }

    if (addressForm.elements["PIN"] != null) {
      if (addressForm.PIN.value != addressForm.ConfirmPIN.value) {
        msg += "- The values of the PIN do not match. Please set again\n";
      }
    }

    if (msg != ""){
      if(document.all){
          document.all("validation").innerText = msg;
          document.all("validation").style.display = "block";
      } else {
        alert(msg);
      }

      return false;
    }
  }

//-->
</SCRIPT>

<!-- #include file="js/validationJS.asp" -->

<%End If%>
</HEAD>
<BODY TOPMARGIN="0" LEFTMARGIN="0" BGCOLOR="ffffff" ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT="0" MARGINWIDTH="0">
<!-- #include file="header_multi.asp" -->
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
	<TR>
		<TD WIDTH="1%" VALIGN="TOP">
			<!-- begin left menu -->
		    <TABLE BORDER="0" CELLPADDING="1" CELLSPACING="0">
				<TR>
				    <TD>
				        <!-- #include file="_toolbar_NewAddress.asp" -->
				    </TD>
				</TR>
                <TR>
                    <TD><IMG SRC="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT=""></TD>
                </TR>
		    </TABLE>
			<!-- end left menu -->
			<IMG SRC="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT="">
		</TD>
		<TD WIDTH="98%" VALIGN="TOP">
			<!-- begin center panel -->
			<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
			    <TR>
			        <TD VALIGN="CENTER">
			            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(26) & " " 'Descriptor: You are here:%> <B><%Response.Write asDescriptors(361) 'Descriptor: Addresses%></B></FONT>
			        </TD>
			        <TD ALIGN="RIGHT"><IMG SRC="images/desktop_Addresses.gif" WIDTH="60" HEIGHT="60" BORDER="0" ALT="" /></TD>
			    </TR>
			</TABLE>
			<%
		    If lErr <> NO_ERR Then
				Call DisplayLoginError(sErrorHeader, sErrorMessage)
		    ElseIf lValidationError <> NO_ERR Then
		        Call DisplayLoginError(sErrorHeader, sErrorMessage)
			End If
			%>

			<% Call RenderAddresses(sDeviceTypeID, sEditAddID, sWizardDeviceID, sWizardAddressName, sWizardPhysicalAddress, sDeviceTypesXML, sGetUserAddressesXML, sGetDevicesFromFoldersXML) %>
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