<%'** Copyright � 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<TABLE WIDTH="157" BORDER="0" CELLSPACING="0" CELLPADDING="0">
  <TR>
    <TD WIDTH="1%" BGCOLOR="#999999"><IMG SRC="../images/1ptrans.gif" WIDTH="14" HEIGHT="30" ALT="" BORDER="0" /></TD>
    <TD WIDTH="99%" BGCOLOR="#999999"><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>" COLOR="#FFFFFF"><B><%=asDescriptors(173)%></B><BR /></FONT></TD>
  </TR>
</TABLE>

<TABLE WIDTH="157" BORDER="0" CELLSPACING="0" CELLPADDING="0" HEIGHT="100%" BGCOLOR="#666666" >
<TR>
  <TD WIDTH="14" ><IMG SRC="../images/1ptrans.gif" WIDTH="14" HEIGHT="1" ALT="" BORDER="0" /></TD>
  <TD VALIGN="TOP">
    <!-- #include file="DateSelector.asp" -->
    <BR />
    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>" COLOR="#FFFFFF">
   		<%Response.Write asDescriptors(95) 'Descriptor: Display:%><BR />
   		&nbsp;&nbsp;&nbsp;<INPUT TYPE="CHECKBOX" NAME="cbError" <%If sErr = "on" Then Response.Write("CHECKED=""1"" ")%>/><%Response.Write asDescriptors(278) 'Descriptors: Errors%><BR />
   		&nbsp;&nbsp;&nbsp;<INPUT TYPE="CHECKBOX" NAME="cbWarning" <%If sWarn = "on" Then Response.Write("CHECKED=""1"" ") %>/><%Response.Write asDescriptors(277) 'Descriptors: Warnings%><BR />
   		&nbsp;&nbsp;&nbsp;<INPUT TYPE="CHECKBOX" NAME="cbMessages" <%If sMessage = "on" Then Response.Write("CHECKED=""1"" ")%>/><%Response.Write asDescriptors(292) 'Descriptors: Messages%>
   	</FONT>
    <BR />
    <BR />
   	<INPUT TYPE="SUBMIT" NAME="bRefresh" CLASS="buttonClass" VALUE="<%Response.Write asDescriptors(269) 'Descriptor: Refresh%>" />
   	<BR /><BR />
	<INPUT TYPE="SUBMIT" NAME="bContinue" CLASS="GOLDBUTTON" VALUE="<%Response.Write asDescriptors(114) 'Descriptor: Continue%>" />
	<BR /><BR />
   	<!--
   	<INPUT TYPE="SUBMIT" NAME="bContinue" CLASS="buttonClass" VALUE="<%Response.Write asDescriptors(114) 'Descriptor: Continue%>" />
   	<BR /><BR />
   	-->
  </TD>
  <TD WIDTH="5" ><IMG SRC="../images/1ptrans.gif" WIDTH="5" HEIGHT="1" ALT="" BORDER="0" /></TD>
</TR>
</TABLE>



