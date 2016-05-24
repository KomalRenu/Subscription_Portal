<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
Function displayDBAliasWidget(aDBAliases, asDescriptors, sPrefix, lRepositoryType)
'********************************************************
'*Purpose:  Display DB Alias widget. This function encapsulates the HTML
'			form for widget
'*Inputs:   aDBAlias (string array)
'*Outputs:  Error Code
'********************************************************
On Error Resume Next
Dim lErr
%>
<TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
  <TR BGCOLOR="#6699CC">
    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="18" /></TD>
    <TD NOWRAP="1"><B><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>" COLOR="#FFFFFF"> <%Response.Write(asDescriptors(558)) 'Descriptor: Database connections%> </FONT></B></TD>
    <TD>&nbsp;&nbsp;</TD>
    <TD></TD>
  </TR>

  <TR>
    <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#000000"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
  </TR>

  <% If IsEmpty(aDBAliases) Then %>
    <TR>
      <TD COLSPAN=8 ALIGN=CENTER>
        <BR/>
        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>"><B>
          <%Call Response.Write(asDescriptors(718)) 'There are no Database Alias defined%>
        </B></FONT>
        <BR />&nbsp;
      </TD>
    </TR>
  <% Else %>
    <% For i = 0 to UBound(aDBAliases) %>
      <TR>
        <TD WIDTH="1%"><INPUT NAME=dba TYPE=radio VALUE="<%=aDBAliases(i,0)%>"  <%If (aDBAliases(i,0) = sDBAlias) Then Response.Write "CHECKED" %> /></TD>

		<%If Len(aDBAliases(i,1)) > 0 Then %>
			<TD NOWRAP="1"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>" COLOR="#000000"><%=aDBAliases(i,1)%></FONT></TD>
			<TD><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>" COLOR="#000000"></FONT></TD>
        	<TD><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>" COLOR="#000000"></FONT></TD>
        <%Else %>
			<TD NOWRAP="1"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>" COLOR="#000000"><%=aDBAliases(i,0)%></FONT></TD>
			<TD><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>" COLOR="#000000"></FONT></TD>
        	<TD ALIGN=CENTER><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>" COLOR="#000000"><A HREF="editDBAlias.asp?st=<%=aSvcConfigInfo(SVCCFG_STEP)%>&rep=<%=lRepositoryType%>&tAliasName=<%=aDBAliases(i,0)%>&tDecodeAlias=<%=aDBAliases(i,1)%>"><%Response.Write(asDescriptors(353)) 'Descriptor:Edit%></A> / <A HREF="delete_dbalias.asp?st=<%=aSvcConfigInfo(SVCCFG_STEP)%>&rep=<%=lRepositoryType%>&tAliasName=<%=aDBAliases(i,0)%>&tDecodeAlias=<%=aDBAliases(i,1)%>"><%Response.Write(asDescriptors(249)) 'Descriptor:Delete%></A></TD>
      	<%End If%>
      </TR>

      <TR>
        <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#6699CC"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
      </TR>
    <% Next %>
  <% End If %>

  <TR>
    <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="5" ><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="5" BORDER="0" ALT=""></TD>
  </TR>

  <TR>
    <TD></TD>
    <TD COLSPAN=3>
      <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>" COLOR="#000000">
		<%If Not IsEmpty(lRepositoryType) Then
			If IsEmpty(aSvcConfigInfo(SVCCFG_STEP)) Then
				Response.Write "<A HREF=""addDBAlias.asp?rep=" & lRepositoryType & """>" & asDescriptors(621) & "</A>" 'Descriptor:Add a New Database Alias
		  	Else
		  		Response.Write "<A HREF=""addDBAlias.asp?st=" & aSvcConfigInfo(SVCCFG_STEP) & "&rep=" & lRepositoryType & """>" & asDescriptors(621) & "</A>" 'Descriptor:Add a New Database Alias
		  	End If
		  Else
			Response.Write "<A HREF=""addDBAlias.asp"">" & asDescriptors(621) & "</A>" 'Descriptor:Add a New Database Alias
		  End If
		%>

      </FONT>
    </TD>
  </TR>

  <TR>
    <TD COLSPAN="4" ALIGN="CENTER" HEIGHT="15" ><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="15" BORDER="0" ALT=""></TD>
  </TR>
<% If (lRepositoryType <> REPOSITORY_WAREHOUSE) Then %>
  <TR>
    <TD COLSPAN=4>
      <TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
        <TR>
          <TD WIDTH="18"><IMG SRC="../images/1ptrans.gif" WIDTH="18" /></TD>
          <TD>
            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>" COLOR="#000000">
              <%Response.write(asDescriptors(583)) 'Descriptor:Tables Prefix:%>
            </FONT>
          </TD>
          <TD><img src="../images/1ptrans.gif" WIDTH="5" HEIGHT="1" BORDER="0" ALT=""></TD>
          <TD>
            <INPUT NAME=pre class="textBoxClass" SIZE=40 VALUE="<%=trim(sPrefix)%>"></INPUT>
          </TD>
        </TR>
      </TABLE>
    </TD>
  </TR>
 <% End If %>
</TABLE>
<%End Function%>
