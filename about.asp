<%'** Copyright ?1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
    Option Explicit
    Response.CacheControl = "no-cache"
    Response.AddHeader "Pragma", "no-cache"
    Response.Expires = -1
    On Error Resume Next
%>
<!-- #include file="CommonDeclarations.asp" -->
<%
    Dim aVersionInfo(1)

    sChannel = ""

    lErr = cu_GetVersions(aVersionInfo)
%>

<HTML>
<HEAD>
    <%Response.Write(putMETATagWithCharSet())%>
    <TITLE><%Response.Write asDescriptors(363)'Descriptor: About%> - MicroStrategy Narrowcast Server 8</TITLE>
</HEAD>
<BODY TOPMARGIN="0" LEFTMARGIN="0" BGCOLOR="ffffff" ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT="0" MARGINWIDTH="0">
<!-- #include file="login_header_multi.asp" -->
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
    <TR>
        <TD WIDTH="1%" valign="TOP">
            <!-- begin left menu -->
            <BR />
            <!-- end left menu -->
            <IMG SRC="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT="">
        </TD>
        <TD WIDTH="98%" valign="TOP">
            <!-- begin center panel -->
            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><B><%If aFontInfo(B_DOUBLE_BYTE_FONT) Then%><%Response.Write asDescriptors(363)'Descriptor: About%><%Else%><%Response.Write UCase(asDescriptors(363))'Descriptor: About%><%End If%></B></FONT>
            <BR><BR>
            <TABLE BORDER="0" CELLPADDING="1" CELLSPACING="0" WIDTH="100%">
                <FORM METHOD="GET" ACTION="default.asp">
                <TR>
                    <TD COLSPAN="2"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B>
					MicroStrategy Narrowcast Server <%Response.Write(GetVersion())%></B></FONT></TD>
                </TR>
                <TR>
                    <TD><IMG SRC="images/1ptrans.gif" WIDTH="10" HEIGHT="1" BORDER="0" ALT="" /></TD>
                    <TD>
                        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>">
                            <%Response.Write asDescriptors(270) 'Descriptor:ASP version:%>&nbsp;<SPGUIVERSION>10.2.0008.0052</SPGUIVERSION> <BR />
                            <%Response.Write asDescriptors(469) 'Descriptor:SDK version:%>&nbsp;<%=aVersionInfo(1)%>
                        </FONT>
                    </TD>
                </TR>
                <TR>
                    <TD COLSPAN="2"><BR /><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><B><%Response.Write asDescriptors(430) 'Descriptor: System Administrator Contact Information%></B></FONT></TD>
                </TR>
                <TR>
                    <TD></TD>
                    <TD>
                        <TABLE BORDER="0" CELLSPACING="0" CELLPADDING="1">
                            <TR>
                                <TD><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><%Response.Write asDescriptors(431) 'Descriptor: E-mail:%></FONT></TD>
                                <TD><A HREF="mailto:<%Response.Write sSysAdminEmail%>"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><%Response.Write sSysAdminEmail%></FONT></A></TD>
                            </TR>
                            <TR>
                                <TD><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><%Response.Write asDescriptors(432) 'Descriptor: Phone:%></FONT></TD>
                                <TD><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_MEDIUM_FONT)%>"><%Response.Write sSysAdminPhone%></FONT></TD>
                            </TR>
                        </TABLE>
                    </TD>
                </TR>
                <TR>
                    <TD COLSPAN="2"><BR /><BR /><INPUT TYPE="SUBMIT" CLASS="buttonClass" VALUE="<%Response.Write asDescriptors(1)'Descriptor: Home%>" />
                </TR>
                <TR>
                    <TD><BR/></TD>
                </TR>
                <TR>
                    <TD COLSPAN = "2">
                        <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%Response.Write aFontInfo(N_SMALL_FONT)%>"><%Response.Write asDescriptors(924) 'Descriptor: This product is patented. One or more of the following patents may apply to the product sold herein: US Patent Nos. 6,154,766, 6,173,310, 6,260,050, 6,263,051, 6,269,393, 6,279,033, 6,501,832, 6,567,796, 6,587,547, 6,606,596, 6,658,093, 6,658,432, 6,662,195, 6,671,715, 6,691,100, 6,694,316, 6,697,808 and 6,704,723. Other patent applications are pending.%></FONT>
                    </TD>
                </TR>
                </FORM>
            </TABLE>
            <!-- end center panel -->
        </TD>
        <TD WIDTH="1%">
            <IMG SRC="images/1ptrans.gif" WIDTH="15" HEIGHT="1" BORDER="0" ALT="">
        </TD>
    </TR>
</TABLE>
<BR>
<!-- begin footer -->
    <!-- #include file="footer.asp" -->
<!-- end footer -->


</BODY>
</HTML>
<%
    Erase aVersionInfo
%>
