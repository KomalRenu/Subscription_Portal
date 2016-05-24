<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>

<%
Function displayEditDBAliasWidget(aDBAlias, asDescriptors)
'********************************************************
'*Purpose:  Display Edit DB Alias widget. This function encapsulates the HTML
'			form for widget
'*Inputs:   aDBAlias (string array)
'*Outputs:  Error Code
'********************************************************
On Error Resume Next
Dim lErr
Dim sDBName
Dim sJServerName
Dim sODBCName
Dim sUserName
Dim sPassword
Dim sDatabaseType
Dim sSystemName
Dim sDefaultPrefix
Dim sPortNumber
Dim sPoolSize

Dim sName

If lErr = NO_ERR Then
	lErr = GetDBAliasProperties(aDBAlias(DBALIAS_NAME),sDBName,sJServerName,sODBCName,sUserName,sPassword,sDatabaseType, sSystemName, sDefaultPrefix, sPortNumber, sPoolSize)
    If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, "", "", "editDBAlias_widget.asp", "", "", "Error calling getDBAliaseProperties", LogLevelTrace)
End If

If UBound(aDBAlias) > 0 Then
    If Len(aDBAlias(DBALIAS_DECODED_NAME)) > 0 Then
		sName = aDBAlias(DBALIAS_DECODED_NAME)
	Else
		sName = aDBAlias(DBALIAS_NAME)
    End If

%>
<TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
    <TR BGCOLOR="#6699CC">
      <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="18" /></TD>
      <TD NOWRAP="1"><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#FFFFFF"> <%Response.Write(asDescriptors(156)) 'Descriptor: Properties%> </FONT></B></TD>
      <TD>&nbsp;&nbsp;</TD>
      <TD><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#FFFFFF"> <%Call Response.write(asDescriptors(203)) 'Value %></FONT></B></TD>
    </TR>

    <TR>
      <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="2"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="2" BORDER="0" ALT=""></TD>
    </TR>
    <TR>
      <TD WIDTH="1%">&nbsp;&nbsp;</TD>
      <TD NOWRAP="1"><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#000000"> <%Response.Write(asDescriptors(557)) 'Descriptor: Database Alias:%></FONT></B></TD>
      <TD>&nbsp;&nbsp;</TD>
      <TD align="left"><input type="hidden"" size="50" name="tAliasName" CLASS="textBoxClass" value="<%=sName%>" /><%=sName%></TD>
      		<INPUT TYPE="HIDDEN" NAME="tDecodeAlias" VALUE="<%=aDBAlias(DBALIAS_DECODED_NAME)%>" />
    </TR>
    <TR>
      <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="2"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="2" BORDER="0" ALT=""></TD>
    </TR>

    <TR>
    	<TD COLSPAN="4" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#6699CC"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
    </TR>
    <TR>
      <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="2"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="2" BORDER="0" ALT=""></TD>
    </TR>
    <TR>
      <TD WIDTH="1%">&nbsp;&nbsp;</TD>
      <TD NOWRAP="1"><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#000000"> <%Response.Write(asDescriptors(624)) 'Descriptor: Database%>:</FONT></B></TD>
      <TD>&nbsp;&nbsp;</TD>
      <TD align="left"><select name="tDatabase" CLASS="pullDownClass" />
      <option value="sqlserver" <%If Strcomp(sDatabaseType, "sqlserver", vbTextCompare) = 0 Then Response.write("SELECTED")%>/>SQL Server
      <option value="db2" <%If Strcomp(sDatabaseType, "db2", vbTextCompare) = 0 Then Response.write("SELECTED")%>/>DB2
      <option value="oracle" <%If Strcomp(sDatabaseType, "oracle", vbTextCompare) = 0 Then Response.write("SELECTED")%>/>Oracle
      <option value="access" <%If Strcomp(sDatabaseType, "access", vbTextCompare) = 0 Then Response.write("SELECTED")%>/>Access
      <option value="teradata" <%If Strcomp(sDatabaseType, "teradata", vbTextCompare) = 0 Then Response.write("SELECTED")%>/>Teradata</select></TD>
    </TR>
    <TR>
      <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="2"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="2" BORDER="0" ALT=""></TD>
    </TR>

    <TR>
    	<TD COLSPAN="4" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#6699CC"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
    </TR>

    <TR>
      <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="2"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="2" BORDER="0" ALT=""></TD>
    </TR>
    <TR>
      <TD WIDTH="1%">&nbsp;&nbsp;</TD>
      <TD NOWRAP="1"><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#000000"><%Response.Write(asDescriptors(778)) 'Descriptor: Server name%>:</FONT></B></TD>
      <TD>&nbsp;&nbsp;</TD>
      <TD align="left"><input type="textfield"" size="50" CLASS="textBoxClass" NAME="tServerName" value="<%=sJServerName%>" /></TD>
    </TR>
    <TR>
      <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="2"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="2" BORDER="0" ALT=""></TD>
    </TR>

    <TR>
    	<TD COLSPAN="4" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#6699CC"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
    </TR>

    <TR>
      <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="2"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="2" BORDER="0" ALT=""></TD>
    </TR>
    <TR>
      <TD WIDTH="1%">&nbsp;&nbsp;</TD>
      <TD NOWRAP="1"><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#000000"><%Response.Write(asDescriptors(707)) 'Descriptor: Database name:%></FONT></B></TD>
      <TD>&nbsp;&nbsp;</TD>
      <TD align="left"><INPUT TYPE="TEXTFIELD"" SIZE="50" NAME="tDBName" CLASS="textBoxClass" VALUE="<%=sDBName%>" /></TD>
    </TR>
    <TR>
      <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="2"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="2" BORDER="0" ALT=""></TD>
    </TR>

    <TR>
    	<TD COLSPAN="4" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#6699CC"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
    </TR>

    <TR>
      <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="2"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="2" BORDER="0" ALT=""></TD>
    </TR>
    <TR>
      <TD WIDTH="1%">&nbsp;&nbsp;</TD>
      <TD NOWRAP="1"><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#000000"><%Response.Write(asDescriptors(10)) 'Descriptor: User name%>:</FONT></B></TD>
      <TD>&nbsp;&nbsp;</TD>
      <TD ALIGN="LEFT"><INPUT TYPE="TEXTFIELD"" SIZE="50" NAME="tUser" CLASS="textBoxClass" VALUE="<%=sUserName%>" /></TD>
    </TR>
    <TR>
      <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="2"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="2" BORDER="0" ALT=""></TD>
    </TR>

    <TR>
    	<TD COLSPAN="4" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#6699CC"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
    </TR>

    <TR>
      <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="2"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="2" BORDER="0" ALT=""></TD>
    </TR>
    <TR>
      <TD WIDTH="1%">&nbsp;&nbsp;</TD>
      <TD NOWRAP="1"><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#000000"><%Response.Write(asDescriptors(11)) 'Descriptor: Password%>:</FONT></B></TD>
      <TD>&nbsp;&nbsp;</TD>
      <TD ALIGN="LEFT"><INPUT TYPE="PASSWORD"" SIZE="50" NAME="tPassword" CLASS="textBoxClass" VALUE="<%=aDBAlias(4)%>" /></TD>
    </TR>
    <TR>
      <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="2"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="2" BORDER="0" ALT=""></TD>
    </TR>

    <TR>
    	<TD COLSPAN="4" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#6699CC"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
    </TR>

    <TR>
      <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="2"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="2" BORDER="0" ALT=""></TD>
    </TR>
    <TR>
      <TD WIDTH="1%">&nbsp;&nbsp;</TD>
      <TD NOWRAP="1"><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#000000"><%Response.Write(asDescriptors(394)) 'Descriptor: Confirm password:%></FONT></B></TD>
      <TD>&nbsp;&nbsp;</TD>
      <TD ALIGN="LEFT"><INPUT TYPE="PASSWORD"" SIZE="50" NAME="tConfirmPassword" CLASS="textBoxClass" VALUE="<%=aDBAlias(4)%>" /></TD>
    </TR>
    <TR>
      <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="2"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="2" BORDER="0" ALT=""></TD>
    </TR>

    <TR>
   	  <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#6699CC"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
    </TR>
    <TR>
      <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="2"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="2" BORDER="0" ALT=""></TD>
    </TR>
    <TR>
      <TD WIDTH="1%">&nbsp;&nbsp;</TD>
      <TD NOWRAP="1"><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#000000"><%Response.Write(asDescriptors(935)) 'Descriptor: Number of pooled connections%>:</FONT></B></TD>
      <TD>&nbsp;&nbsp;</TD>
      <TD align="left"><input type="textfield"" size="50" CLASS="textBoxClass" NAME="tPoolSize" value="<%=sPoolSize%>" /></TD>
    </TR>
    <TR>
      <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="2"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="2" BORDER="0" ALT=""></TD>
    </TR>

    <TR>
    	<TD COLSPAN="4" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#6699CC"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
    </TR>
<!--
    <TR>
      <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="2"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="2" BORDER="0" ALT=""></TD>
    </TR>
    <TR>
    	<TD WIDTH="1%">&nbsp;&nbsp;</TD>
    	<TD NOWRAP="1"><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#000000"><%Response.Write(asDescriptors(708)) 'Descriptor: Test DB Alias connection?%></FONT></B></TD>
    	<TD>&nbsp;&nbsp;</TD>
    	<TD align="left"><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#000000"><INPUT NAME="tValidate" TYPE=checkbox VALUE="true" CHECKED /> <%Response.Write(asDescriptors(119)) 'Descriptor: Yes%></FONT></B></TD>
    </TR>
    <TR>
      <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="2"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="2" BORDER="0" ALT=""></TD>
    </TR>
-->
    <TR>
    	<TD WIDTH="1%">&nbsp;&nbsp;</TD>
    	<TD NOWRAP="1"><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#000000"></FONT></B></TD>
    	<TD>&nbsp;&nbsp;</TD>
    	<TD align="left"><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="1" COLOR="#000000"><INPUT NAME="tClassName" TYPE=hidden VALUE="com.mstr.transactorx.dbdriver.SequeLinkWrapper" CHECKED /></FONT></B></TD>
    </TR>

	<input type="hidden" name="tSystemName" value="<%=sSystemName%>" />
	<input type="hidden" name="tPrefix" value="<%=sDefaultPrefix%>" />
	<input type="hidden" name="tPortNumber" value="<%=sPortNumber%>" />

    <input type="hidden" name="submit" value="submitted" />

  <TR>
    <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="5" ><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="5" BORDER="0" ALT=""></TD>
  </TR>
</TABLE>
<%
	lErr = NO_ERR
Else
	lErr = ERR_EXCEPTION_THROWN
End If

displayeditDBAliasWidget = lErr

End Function
%>