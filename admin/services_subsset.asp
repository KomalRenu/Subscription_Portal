<%
  Option Explicit
  Response.CacheControl = "no-cache"
  Response.AddHeader "Pragma", "no-cache"
  Response.Expires = -1
  On Error Resume Next
%>
<!-- #include file="../CommonDeclarations.asp" -->
<!-- #include file="../CustomLib/AdminCuLib.asp" -->
<!-- #include file="../CustomLib/ServicesConfigCuLib.asp" -->
<%
Dim lStatus
Dim aAnswers

Dim aNormalQuestions
Dim aSlicingQuestions
Dim aExtraQuestions

Dim sNormalQOIds
Dim sSlicingQOIds
Dim sExtraQOIds

Dim i
Dim lCount


    If lErr = NO_ERR Then
        lErr = ParseRequestForSvcConfig(oRequest, aSvcConfigInfo)

        'The following information is not valid anymore:
        aSvcConfigInfo(SVCCFG_QO_ID) = ""
        aSvcConfigInfo(SVCCFG_QO_NAME) = ""
        aSvcConfigInfo(SVCCFG_QO_PARENT_ID) = ""
        aSvcConfigInfo(SVCCFG_AQ_ID) = ""
        aSvcConfigInfo(SVCCFG_AQ_NAME) = ""
        aSvcConfigInfo(SVCCFG_AQ_PARENT_ID) = ""

    End If


    If lErr = NO_ERR Then
        lErr = GetSubscriptionSetConfig(aSvcConfigInfo, aNormalQuestions, aSlicingQuestions, aExtraQuestions)
    End If

    If lErr = NO_ERR Then

        If Not IsEmpty(aSlicingQuestions) Then
            lCount = UBound(aSlicingQuestions)
            For i = 0 To lCount
                sSlicingQOIds = sSlicingQOIds & aSlicingQuestions(i, QO_ID) & ";"
            Next
        End If

        If Not IsEmpty(aNormalQuestions) Then
            lCount = UBound(aNormalQuestions)
            For i = 0 To lCount
                sNormalQOIds = sNormalQOIds & aNormalQuestions(i, QO_ID) & ";"
            Next
        End If

        If Not IsEmpty(aExtraQuestions) Then
            lCount = UBound(aExtraQuestions)
            For i = 0 To lCount
                sExtraQOIds = sExtraQOIds & aExtraQuestions(i, QO_ALTERNATE_ID) & ";"
            Next
        End If


        Redim aAnswers(3, 1)

        aAnswers(0, 0) = "subscription." & ANSWER_USER_ID
        aAnswers(0, 1) = ANSWER_USER_ID

        aAnswers(1, 0) = "subscription." & ANSWER_SUBSCRIPTION_ID
        aAnswers(1, 1) = ANSWER_SUBSCRIPTION_ID

        aAnswers(2, 0) = "subscription." & ANSWER_ADDRESS_ID
        aAnswers(2, 1) = ANSWER_ADDRESS_ID

        aAnswers(3, 0) = ANSWER_OTHER_ID
        aAnswers(3, 1) = asDescriptors(811)'"Answer to this question:"

    End If

    'Set the PageInfo to be used by the navigator bar and the header.
    aPageInfo(S_TITLE_PAGE) = STEP_SERVICES_STATIC_SUBSET & " " & asDescriptors(785)'"Configure Subscription Set"
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SERVICES
    aPageInfo(N_OPTIONS_WITH_LINKS_PAGE) = CreateRequestForSvcConfig(aSvcConfigInfo)

    lStatus = checkSiteConfiguration()


%>
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>
<SCRIPT LANGUAGE=javascript>
<!--
function checkSliceBy(sQOid) {
var thisCombo;
var thisDiv;
var questions;
var bNeedAnswer;
var questionCombo;
var i;

    thisCombo = document.all("v" + sQOid);

    if (thisCombo != null) {
        thisDiv = document.all("pick" + sQOid);

        if (thisCombo.selectedIndex == 3) {
          thisDiv.style.display = "";
        } else {
          thisDiv.style.display = "none";
        }
    }

    //Find how many questions still need an alternate answer:
    questions = document.all("slicing").value.split(";");

    bNeedAnswer = false;
    for(i=0;i<questions.length - 1;i++) {
      questionCombo = document.all("v" + questions[i]);
      if (questionCombo.selectedIndex == 3) {
        if (document.all("a" + questions[i]).value == "") {
          bNeedAnswer = true || bNeedAnswer;
        }
      }
    }

    //If we still need an alternate question, hide the next button:
    if (bNeedAnswer == true) {
      document.all("next").style.display = "none";
    } else {
      document.all("next").style.display = "";
    }

}


//-->
</SCRIPT>

<!-- #include file="../NSStyleSheet.asp" -->

</HEAD>
<BODY BGCOLOR="#ffffff" TOPMARGIN=0 LEFTMARGIN=0 ALINK="#ff0000" LINK="#0000ff" VLINK="#0000ff" MARGINHEIGHT="0" MARGINWIDTH="0" <% If Not IsEmpty(aSlicingQuestions) Then Response.write("onLoad=""checkSliceBy('');""") %> >
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" HEIGHT="100%">
  <TR>
    <TD COLSPAN="6" HEIGHT="1%"><!-- begin header --><!-- #include file="admin_header.asp" --><!-- end header -->
    </TD>
  </TR>
  <TR>
    <TD WIDTH="1%" valign="top"><!-- begin toolbar --><!-- #include file="_toolbar_services.asp" --><!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG alt="" border=0 height=1 src="../images/1ptrans.gif" width=21 ></TD>

    <TD WIDTH="96%" valign="top">
      <%If lErr <> NO_ERR Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(623), "select_site.asp") 'Descriptor:Site Definition%>
      <%Else%>
        <BR />
        <% Call RenderSvcConfigPath(aSvcConfigInfo) %>
        <BR />
        <BR />
        <FORM ACTION="services_subsset_modify.asp">
        <% Call RenderSvcConfigInputs(aSvcConfigInfo) %>
          <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" >
            <B><%Response.Write(asDescriptors(769)) 'Page-by Questions:%></B><BR />
          </FONT>
          <BR />
          <!-- Start Slicing Questions -->
          <% If IsEmpty(aSlicingQuestions) Then %>
          <P ALIGN="CENTER">
            <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
              <B><%Response.Write(asDescriptors(824)) 'This Service does not contain pabe-by questions to configure. %></B>
            </FONT>
          </P>
          <% Else %>

            <TABLE WIDTH="100%" BORDER=0 CELLPADDING=0 CELLSPACING=0>
              <TR BGCOLOR="#c2c2c2">
                <TD>&nbsp;</TD>
                <TD><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" ><%Call Response.Write(asDescriptors(789)) 'Question%></FONT></B></TD>
                <TD>&nbsp;</TD>
                <TD><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" ><%Call Response.Write(asDescriptors(790)) 'Publication(s)%></FONT></B></TD>
                <TD>&nbsp;</TD>
                <TD><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" ><%Call Response.Write(asDescriptors(282)) 'Page-by%></FONT></B>&nbsp;</TD>
              </TR>

              <% lCount = UBound(aSlicingQuestions)%>
              <% For i = 0 To lCount %>
                <TR>
                  <TD COLSPAN="6" HEIGHT="2"><IMG SRC="../images/1ptrans.gif" HEIGHT="2" WIDTH="1" /></TD>
                </TR>
                <TR>
                  <TD ROWSPAN="2" VALIGN="TOP"></TD>
                  <TD ROWSPAN="2" VALIGN="TOP"><INPUT TYPE=HIDDEN NAME="n<%=aSlicingQuestions(i, QO_ID)%>" VALUE="<%=Server.HTMLEncode(aSlicingQuestions(i, QO_NAME))%>" /><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" ><%=Server.HTMLEncode(aSlicingQuestions(i, QO_NAME))%></FONT></TD>
                  <TD ROWSPAN="2" VALIGN="TOP"></TD>
                  <TD ROWSPAN="2" VALIGN="TOP"><INPUT TYPE=HIDDEN NAME="d<%=aSlicingQuestions(i, QO_ID)%>" VALUE="<%=Server.HTMLEncode(aSlicingQuestions(i, QO_DESCRIPTION))%>" /><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" ><%=Server.HTMLEncode(aSlicingQuestions(i, QO_DESCRIPTION))%></FONT></TD>
                  <TD ROWSPAN="2" VALIGN="TOP"></TD>
                  <TD><%Call RenderDropDownList("v" & aSlicingQuestions(i, QO_ID), aAnswers, aSlicingQuestions(i, QO_VALUE), "checkSliceBy('" & aSlicingQuestions(i, QO_ID) & "')")%></TD>
                </TR>
                <TR>
                  <TD><DIV <%If aSlicingQuestions(i, QO_VALUE) <> ANSWER_OTHER_ID Then Call Response.Write("STYLE=""display:none;""") %> ID="pick<%=aSlicingQuestions(i, QO_ID)%>">
                    <INPUT NAME="a<%=aSlicingQuestions(i, QO_ID)%>" TYPE="HIDDEN" VALUE="<%=aSlicingQuestions(i, QO_ALTERNATE_ID)%>" />
                    <INPUT NAME="m<%=aSlicingQuestions(i, QO_ID)%>" TYPE="HIDDEN" VALUE="<%=aSlicingQuestions(i, QO_MAP_ID)%>" />
                    <INPUT NAME="i<%=aSlicingQuestions(i, QO_ID)%>" TYPE="HIDDEN" VALUE="<%=aSlicingQuestions(i, QO_IS_ID)%>" />
                    <INPUT NAME="p<%=aSlicingQuestions(i, QO_ID)%>" TYPE="HIDDEN" VALUE="<%=aSlicingQuestions(i, QO_PROMPT_COUNT)%>" />
                    <TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
                      <TR>
                        <TD>
                          <TABLE BORDER="0" CELLSPACING="1" CELLPADDING="2" BGCOLOR="#c2c2c2">
                            <TR>
                              <TD BGCOLOR="#ffffff" NOWRAP><%
                              If Len(aSlicingQuestions(i, QO_ALTERNATE_ID)) = 0 Then
                                Response.Write("<I><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""red"" >(" & asDescriptors(725) & ")</FONT></I>") 'Pick an Alternate question
                              Else
                                Response.Write("<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ >" & aExtraQuestions(GetExtraQuestionIndex(aExtraQuestions, aSlicingQuestions(i, QO_ALTERNATE_ID)), QO_NAME) & "</FONT>")
                              End If%></TD>
                            </TR>
                          </TABLE>
                        </TD>
                        <TD NOWRAP>&nbsp;&nbsp;<INPUT NAME="b<%=aSlicingQuestions(i, QO_ID)%>" TYPE="SUBMIT" class="buttonClass" VALUE="<%Call Response.Write(asDescriptors(810)) 'Browse%>" />&nbsp;&nbsp;</TD>
                      </TR>
                    </TABLE>
                  </DIV></TD>
                </TR>

                <TR>
                  <TD COLSPAN="6" HEIGHT="1"><IMG SRC="../images/1ptrans.gif" HEIGHT="2" WIDTH="1" /></TD>
                </TR>

                <TR>
                  <TD COLSPAN="6" HEIGHT="1" BGCOLOR="#6699CC"><IMG SRC="../images/1ptrans.gif" HEIGHT="1" WIDTH="1" /></TD>
                </TR>

              <%Next %>
            </TABLE>

            <INPUT TYPE=HIDDEN NAME="slicing" VALUE="<%=sSlicingQOIds%>" />
          <% End If%>
          <!-- End Slicing Questions -->
          <BR />
          <BR />
          <!-- Start Normal Questions -->
          <% If Not IsEmpty(aNormalQuestions) Then %>
              <% lCount = UBound(aNormalQuestions)%>
              <% For i = 0 To lCount %>
                  <INPUT TYPE=HIDDEN NAME="n<%=aNormalQuestions(i, QO_ID)%>" VALUE="<%=Server.HTMLEncode(aNormalQuestions(i, QO_NAME))%>" />
                  <INPUT TYPE=HIDDEN NAME="d<%=aNormalQuestions(i, QO_ID)%>" VALUE="<%=Server.HTMLEncode(aNormalQuestions(i, QO_DESCRIPTION))%>" />
                  <INPUT TYPE=HIDDEN NAME="v<%=aNormalQuestions(i, QO_ID)%>" VALUE="<%If aNormalQuestions(i, QO_ID) <> "false" Then Response.Write "show"%>" />
              <%Next %>
            <INPUT TYPE=HIDDEN NAME="normal" VALUE="<%=sNormalQOIds%>" />
          <%End If%>
          <!-- End Normal Questions -->


          <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" >
            <B><%Response.Write(asDescriptors(770)) 'Additional Questions%></B><BR />
          </FONT>
          <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
            <%Response.Write(asDescriptors(771)) 'If you would like to prompt subscribers with any additional Questions, you may specify the Questions and the storage of their answer below.%><BR />
          </FONT>

          <!-- Start Extra Questions -->
          <% If IsEmpty(aExtraQuestions) Then %>
          <!--<P ALIGN="CENTER">
            <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
              <B>This Service does not contain any extra Questions.</B>
            </FONT>
           -->
          </P>
          <% Else %>
            <TABLE WIDTH="100%" BORDER=0 CELLPADDING=0 CELLSPACING=0>
              <TR BGCOLOR="#c2c2c2">
                <TD>&nbsp;</TD>
                <TD><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" ><%Call Response.Write(asDescriptors(789)) 'Question%></FONT></B></TD>
                <TD>&nbsp;</TD>
                <TD><B><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" ><%Response.Write(asDescriptors(724)) 'Storage%></FONT></B></TD>
                <TD>&nbsp;</TD>
                <TD>&nbsp;</TD>
              </TR>

              <% lCount = UBound(aExtraQuestions)%>
              <% For i = 0 To lCount %>
                    <INPUT NAME="n<%=aExtraQuestions(i, QO_ALTERNATE_ID)%>" TYPE="HIDDEN" VALUE="<%=Server.HTMLEncode(aExtraQuestions(i, QO_NAME))%>" />
                    <INPUT NAME="d<%=aExtraQuestions(i, QO_ALTERNATE_ID)%>" TYPE="HIDDEN" VALUE="<%=Server.HTMLEncode(aExtraQuestions(i, QO_DESCRIPTION))%>" />
                    <INPUT NAME="v<%=aExtraQuestions(i, QO_ALTERNATE_ID)%>" TYPE="HIDDEN" VALUE="<%=aExtraQuestions(i, QO_VALUE)%>" />
                    <INPUT NAME="i<%=aExtraQuestions(i, QO_ALTERNATE_ID)%>" TYPE="HIDDEN" VALUE="<%=aExtraQuestions(i, QO_IS_ID)%>" />
                    <INPUT NAME="p<%=aExtraQuestions(i, QO_ALTERNATE_ID)%>" TYPE="HIDDEN" VALUE="<%=aExtraQuestions(i, QO_PROMPT_COUNT)%>" />
                    <INPUT NAME="a<%=aExtraQuestions(i, QO_ALTERNATE_ID)%>" TYPE="HIDDEN" VALUE="<%=aExtraQuestions(i, QO_ALTERNATE_ID)%>" />
                    <INPUT NAME="m<%=aExtraQuestions(i, QO_ALTERNATE_ID)%>" TYPE="HIDDEN" VALUE="<%=aExtraQuestions(i, QO_MAP_ID)%>" />
                  <% If aExtraQuestions(i, QO_VALUE) = "false" Then %>

                <TR>
                  <TD COLSPAN="6" HEIGHT="2"><IMG SRC="../images/1ptrans.gif" HEIGHT="2" WIDTH="1" /></TD>
                </TR>
                <TR>
                  <TD VALIGN="TOP"></TD>
                  <TD VALIGN="TOP"><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" ><%=Server.HTMLEncode(aExtraQuestions(i, QO_NAME))%></FONT></TD>
                  <TD VALIGN="TOP"></TD>
                  <TD VALIGN="TOP"><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" ><%=Server.HTMLEncode(aExtraQuestions(i, QO_DESCRIPTION))%></FONT></TD>
                  <TD VALIGN="TOP"></TD>
                  <TD VALIGN="TOP">
                    <INPUT TYPE=SUBMIT NAME="e<%=aExtraQuestions(i, QO_ALTERNATE_ID)%>" CLASS="buttonClass" VALUE="<%Response.Write(asDescriptors(353)) 'Edit%>" />&nbsp;<INPUT TYPE=SUBMIT NAME="r<%=aExtraQuestions(i, QO_ALTERNATE_ID)%>" CLASS="buttonClass" VALUE="<%Response.write(asDescriptors(106)) 'Remove%>" />&nbsp</TD>
                </TR>

                <TR>
                  <TD COLSPAN="6" HEIGHT="2"><IMG SRC="../images/1ptrans.gif" HEIGHT="2" WIDTH="1" /></TD>
                </TR>

                <TR>
                  <TD COLSPAN="6" HEIGHT="1" BGCOLOR="#6699CC"><IMG SRC="../images/1ptrans.gif" HEIGHT="1" WIDTH="1" /></TD>
                </TR>

                <%End If
              Next %>
            </TABLE>

            <INPUT TYPE=HIDDEN NAME="extra" VALUE="<%=sExtraQOIds%>" />
          <% End If %>

          <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
            <TR>
              <TD VALIGN="CENTER">
                <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                  <%Call Response.Write(asDescriptors(809)) 'Add a question:%>&nbsp;&nbsp;
                </FONT>
              </TD>
              <TD>
                <INPUT TYPE=SUBMIT NAME="addqo" CLASS="buttonClass" VALUE="<%Call Response.Write(asDescriptors(810)) 'Browse%>" />
              </TD>
            </TR>
          </TABLE>
          <!-- End Extra Questions -->
          <BR />
          <BR />

          <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="100%">
            <TR>
              <TD COLSPAN="2">
                <BR />
              </TD>
            </TR>

            <TR>
              <%If aSvcConfigInfo(SVCCFG_STEP) = STATIC_SS Then%>
              <TD COLSPAN="2" ALIGN="CENTER" NOWRAP WIDTH="100%">
                <INPUT NAME=next TYPE=SUBMIT CLASS="buttonClass" VALUE="<%Response.Write(asDescriptors(543)) 'Descriptor:OK%>"></INPUT> &nbsp;
                <INPUT NAME=back TYPE=SUBMIT CLASS="buttonClass" VALUE="<%Response.Write(asDescriptors(120)) 'Descriptor:Cancel%>"></INPUT> &nbsp;
              </TD>
              <%Else%>
              <TD ALIGN="left" NOWRAP WIDTH="1%">
                  <INPUT NAME=back TYPE=SUBMIT CLASS="buttonClass" VALUE="<%Response.Write(asDescriptors(334)) 'Descriptor:back%>"></INPUT> &nbsp;
              </TD>
              <TD ALIGN="left" NOWRAP WIDTH="98%">
                <INPUT NAME=next TYPE=SUBMIT CLASS="buttonClass" VALUE="<%Response.Write(asDescriptors(335)) 'Descriptor:OK%>"></INPUT> &nbsp;
              </TD>
              <%End If%>
            </TR>
          </TABLE>

        </FORM>
      <%End If %>
    </TD>

    <TD WIDTH="1%"><IMG alt="" border=0 height=1 src="../images/1ptrans.gif" width=21 ></TD>

    <TD WIDTH="1%"><!-- #include file="help_widget.asp" -->
    </TD>
  </TR>
</TABLE>
</BODY>
</HTML>
<%
	Erase aAnswers
	Erase aNormalQuestions
	Erase aSlicingQuestions
	Erase aExtraQuestions

%>