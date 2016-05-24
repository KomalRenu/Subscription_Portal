<TABLE WIDTH="100%" BORDER=0 CELLPADDING=0 CELLSPACING=0>
  <TR>
    <TD>
      <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
        <%Response.Write(asDescriptors(730)) 'Static subscription sets%>
        <INPUT NAME="old" TYPE=HIDDEN VALUE="<%=sAnswer%>" />
      </FONT>
      <BR />
    </TD>
  </TR>
  <TR>
    <TD NOWRAP><IMG SRC="../images/bullet.gif" ALIGN="LEFT">
      <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
        <%Response.Write(asDescriptors(732)) 'Hide page-by questions from subscribers%>
      </FONT>
      <BR />
    </TD>
  </TR>
    <TD NOWRAP VALIGN="TOP">
      <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
        <TR>
          <TD><IMG SRC="../images/bullet.gif" ALIGN="LEFT"><FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" ><%Response.Write(asDescriptors(733)) 'Answer page-by questions with:%> &nbsp;</FONT><%If IsEmpty(aAnswers) Then
                Response.Write "<B><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ >" & asDescriptors(315) & " (" & Mid(Application.Value("Default_slicing_answer"), Len("subscription.") + 1) & ")</FONT></B>"  'Descriptor: Default
            Else
                Response.Write "</TD><TD>"
          
                Call RenderDropDownList("ans", aAnswers, sAnswer, "")
            End If%></TD>
        </TR>
      </TABLE>
    </TD>
  </TR>
  <TR>
    <TD><IMG SRC="../images/1ptrans.gif" WIDTH="1" HEIGHT="10"></TD>
  </TR>
  <TR>
    <TD>
      <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
        <%Response.Write(asDescriptors(731)) 'Dynamic Subscription Sets%>
      </FONT>
      <BR />
    </TD>
  </TR>
  <TR>
    <TD NOWRAP><IMG SRC="../images/bullet.gif" ALIGN="LEFT">
      <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
        <%Response.Write(asDescriptors(734)) 'Hide unconfigured subscription set from subscribers%>
      </FONT>
      <BR />
    </TD>
  </TR>
  <TR>
    <TD><IMG SRC="../images/1ptrans.gif" WIDTH="1" HEIGHT="10"></TD>
  </TR>
</TABLE>
