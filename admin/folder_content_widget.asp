<%
Dim oFolderDOM
Dim oFolders
Dim oItems
Dim oItem
Dim sItemImage
Dim aIds
Dim nIndex

Dim oParent
Dim sParentId

    Set oFolderDOM = Server.CreateObject("Microsoft.XMLDOM")
    Call oFolderDOM.loadXML(sFolderContentXML)

    Set oFolders = oFolderDOM.selectNodes("/mi/fct/oi[@tp=""" & TYPE_FOLDER & """]")
    Set oItems = oFolderDOM.selectNodes("/mi/fct/oi[@tp=""" & sType & """]")
    Set oParent = oFolderDOM.selectSingleNode("//fd[../a/fd[@id='" & sFolderId & "']]")
    
    If sType = TYPE_SERVICE Then
        sItemImage = "../images/report2.gif"
    Else
        sItemImage = "../images/1ptrans.gif"
    End If
    
    If oItems.length > 0 Then
    
        Redim aIds(oItems.length -1)
        For nIndex = 0 To oItems.length - 1
            aIds(nIndex) = oItems.item(nIndex).attributes.getNamedItem("id").text
        Next
        
        sId = selectDefaultValue(aIds, Array(sId))
        
    End If
        
%>
<TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
  <TR>
    <TD>
      <%Call RenderFolderPath(oFolderDOM, sFolderLink, "fid") %>
    </TD>
  </TR>
  <TR>
    <TD>
      <BR />
    </TD>
  </TR>
  <TR>
    <TD>
      <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="100%">
        <TR>
          <TD COLSPAN=13 BGCOLOR="#99ccff"><IMG SRC="images/1ptrans.gif" HEIGHT="1" WIDTH="1" BORDER="0" ALT="" /></TD>
        </TR>
        <TR BGCOLOR="#6699cc">
          <TD WIDTH=16><IMG SRC="../images/1ptrans.gif" WIDTH=16></TD>
          <TD WIDTH=16><IMG SRC="../images/1ptrans.gif" WIDTH=16></TD>
          <TD><IMG SRC="../images/1ptrans.gif" WIDTH=2></TD>
          <TD><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" color="#ffffff" size="<%=aFontInfo(N_SMALL_FONT)%>"><b><%Call Response.Write(asDescriptors(306)) 'Descriptor: Name%></b></font></TD>
          <TD>&nbsp;&nbsp;</TD>
          <TD>&nbsp;&nbsp;</TD>
          <TD><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" color="#ffffff" size="<%=aFontInfo(N_SMALL_FONT)%>"><b><%Call Response.Write(asDescriptors(34)) 'Descriptor: Modified%></b></font></TD>
          <TD>&nbsp;&nbsp;</TD>
          <TD>&nbsp;&nbsp;</TD>
          <TD><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" color="#ffffff" size="<%=aFontInfo(N_SMALL_FONT)%>"><b><%Call Response.Write(asDescriptors(22)) 'Descriptor: Description%></b></font></TD>
          <TD>&nbsp;&nbsp;</TD>
        </TR>
        
        <TR>
          <TD COLSPAN=13 BGCOLOR="#003366"><IMG SRC="images/1ptrans.gif" HEIGHT="1" WIDTH="1" BORDER="0" ALT="" /></TD>
        </TR>
        
        <%If (oFolders.length + oItems.length) > 0 Then %>

          <%For Each oItem in oFolders%>
            <TR>
              <TD COLSPAN=13 BGCOLOR="#ffffff"><IMG SRC="images/1ptrans.gif" HEIGHT="1" WIDTH="1" BORDER="0" ALT="" /></TD>
            </TR>
            
            <TR>
              <TD WIDTH=16><IMG SRC="../images/1ptrans.gif" WIDTH=16></TD>
              <TD WIDTH=16><A HREF="<%=sFolderLink & "&fid=" & oItem.attributes.getNamedItem("id").text%>"><IMG SRC="../images/folder2.gif" HEIGHT="16" WIDTH="16" BORDER="0" ALT="" /></A></TD>
              <TD><IMG SRC="../images/1ptrans.gif" WIDTH=2></TD>
              <TD NOWRAP><A HREF="<%=sFolderLink & "&fid=" & oItem.attributes.getNamedItem("id").text%>"><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" color="#000000" size="<%=aFontInfo(N_SMALL_FONT)%>"><b><%=oItem.attributes.getNamedItem("n").text%></b></font></A></TD>
              <TD></TD>
              <TD></TD>
              <TD><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" size="<%=aFontInfo(N_SMALL_FONT)%>"><%=DisplayDateAndTime(CDate(oItem.attributes.getNamedItem("mdt").text), "")%></font></TD>
              <TD></TD>
              <TD></TD>
              <TD><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" size="<%=aFontInfo(N_SMALL_FONT)%>"><%=oItem.attributes.getNamedItem("des").text%></font></TD>
              <TD></TD>
            </TR>
                
            <TR>
              <TD COLSPAN=13 BGCOLOR="#ffffff"><IMG SRC="images/1ptrans.gif" HEIGHT="2" WIDTH="1" BORDER="0" ALT="" /></TD>
            </TR>
            
            <TR>
              <TD COLSPAN=13 BGCOLOR="#99ccff"><IMG SRC="images/1ptrans.gif" HEIGHT="1" WIDTH="1" BORDER="0" ALT="" /></TD>
            </TR>
            
          <%Next%>
            
          <%For Each oItem in oItems %>
          
            <TR>
              <TD COLSPAN=13 BGCOLOR="#ffffff"><IMG SRC="images/1ptrans.gif" HEIGHT="1" WIDTH="1" BORDER="0" ALT="" /></TD>
            </TR>
            
            <TR>
              <TD WIDTH="16"><INPUT TYPE="radio" NAME="id" VALUE="<%=oItem.attributes.getNamedItem("id").text%>" <%If sId = oItem.attributes.getNamedItem("id").text Then Response.Write("CHECKED")%> /></TD>
              <TD WIDTH="16"><IMG SRC="<%=sItemImage%>" WIDTH="16" HEIGHT="16"></TD>
              <TD><IMG SRC="../images/1ptrans.gif" WIDTH=2></TD>
              <TD NOWRAP><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" color="#000000" size="<%=aFontInfo(N_SMALL_FONT)%>"><b><%=oItem.attributes.getNamedItem("n").text%></b></font></TD>
              <TD></TD>
              <TD></TD>
              <TD><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" size="<%=aFontInfo(N_SMALL_FONT)%>"><%=DisplayDateAndTime(CDate(oItem.attributes.getNamedItem("mdt").text), "")%></font></TD>
              <TD></TD>
              <TD></TD>
              <TD><font face="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" size="<%=aFontInfo(N_SMALL_FONT)%>"><%=oItem.attributes.getNamedItem("des").text%></font></TD>
              <TD></TD>
            </TR>
                
            <TR>
              <TD COLSPAN=13 BGCOLOR="#ffffff"><IMG SRC="images/1ptrans.gif" HEIGHT="2" WIDTH="1" BORDER="0" ALT="" /></TD>
            </TR>
            
            <TR>
              <TD COLSPAN=13 BGCOLOR="#99ccff"><IMG SRC="images/1ptrans.gif" HEIGHT="1" WIDTH="1" BORDER="0" ALT="" /></TD>
            </TR>
            <%Next%>
        <%Else%>
                
            <TR>
              <TD COLSPAN=13 BGCOLOR="#ffffff"><IMG SRC="images/1ptrans.gif" HEIGHT="10" WIDTH="1" BORDER="0" ALT="" /></TD>
            </TR>
            
            <TR>
              <TD ALIGN="CENTER" COLSPAN=13 BGCOLOR="#ffffff">
                <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>">
                  <B><%Response.Write(asDescriptors(837)) 'There are no valid objects in this folder%></B>
                </FONT></TD>
              </TD>
            </TR>
            
            <TR>
              <TD COLSPAN=13 BGCOLOR="#ffffff"><IMG SRC="images/1ptrans.gif" HEIGHT="10" WIDTH="1" BORDER="0" ALT="" /></TD>
            </TR>

            <%If Not oParent Is Nothing Then %>
            <TR>
              <TD COLSPAN=13 BGCOLOR="#ffffff">
                <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>">
                   <B><A HREF="<%=sFolderLink & "&fid=" & oParent.getAttribute("id")%>"><%Response.Write(asDescriptors(147)) 'Back to parent folder%></A></B>
                </FONT></TD>
              </TD>
            </TR>
            <%End If%>
            
        <%End If%>

      </TABLE>
      
    </TD>
  </TR>
</TABLE>
