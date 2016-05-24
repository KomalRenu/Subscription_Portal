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
Dim aMapInfo
Dim aTablesInfo

Dim aMapList
Dim aMapIds
Dim aPromptList
Dim lPromptCount
Dim sISID
Dim nStorageType
Dim sRequestForSvcConfig

Dim sId
Dim lCount
Dim i

    If lErr = NO_ERR Then
        lErr = ParseRequestForMap(oRequest, aSvcConfigInfo, aMapInfo, aTablesInfo)
    End If

    'Check for actions cancelled:
    If Len(oRequest("back")) > 0 Then
        sRequestForSvcConfig = CreateRequestForSvcConfig(aSvcConfigInfo)
        Response.Redirect("services_subsset.asp?" & sRequestForSvcConfig)
    End If

    'Get the current Map:
    If lErr = NO_ERR Then
        If Len(aMapInfo(MAP_ID)) = 0 Then
            lErr = GetQuestionMap(aSvcConfigInfo, aMapInfo(MAP_ID))
        End If
    End If


    If lErr = NO_ERR Then
        lErr = GetMapsForQuestion(aSvcConfigInfo, aMapList)
    End If

    'Get the storage this question can have:
    If lErr = NO_ERR Then

        lErr = GetPromptsForQuestion(aSvcConfigInfo, aPromptList)

        If lErr = NO_ERR Then
            lErr = GetQuestionStorageType(aSvcConfigInfo, aPromptList, nStorageType)
        End If

    End If

    'Get information about the QO:
    If lErr = NO_ERR Then
        If Not IsEmpty(aPromptList) Then
            aMapInfo(MAP_QO_PROMPT_COUNT) = UBound(aPromptList) + 1
            aMapInfo(MAP_QO_IS) = aPromptList(0, PROMPT_ISID)
        End If
    End If

    'Get the count of storage mappings saved for this object:
    If lErr = NO_ERR Then
        If Not IsEmpty(aMapList) Then lCount = UBound(aMapList) + 1
    End If

    If lErr = NO_ERR Then

        If (nStorageType And STORAGE_SBR_ONLY) > 0 Then
            ReDim aMapIds(lCount)
            aMapIds(lCount) = "sbr"
        Else
            If lCount > 0 Then
                ReDim aMapIds(lCount - 1)
            End If
        End If

        For i = 0 To lCount - 1
            aMapIds(i) = aMapList(i, MAP_ID)
        Next

        aMapInfo(MAP_ID) = selectDefaultValue(aMapIds, Array(aMapInfo(MAP_ID), "sbr"))

    End If

    'Set the PageInfo to be used by the navigator bar and the header.
    If aSvcConfigInfo(SVCCFG_STEP) = STATIC_SS Then
        aPageInfo(S_TITLE_PAGE) = STEP_SERVICES_STATIC_SELECT_MAP & " " & asDescriptors(782) '"Select Storage"
    Else
        aPageInfo(S_TITLE_PAGE) = STEP_SERVICES_DYNAMIC_SELECT_MAP & " " & asDescriptors(782) '"Select Storage"
    End If
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SERVICES
    aPageInfo(N_OPTIONS_WITH_LINKS_PAGE) = CreateRequestForSvcConfig(aSvcConfigInfo)

    lStatus = checkSiteConfiguration()


%>
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>

  <!-- #include file="../NSStyleSheet.asp" -->

</HEAD>
<BODY BGCOLOR="#ffffff" TOPMARGIN=0 LEFTMARGIN=0 ALINK="#ff0000" LINK="#0000ff" VLINK="#0000ff" MARGINHEIGHT="0" MARGINWIDTH="0">
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
      <%
        If lErr <> NO_ERR Then

            Select Case lErr
            Case ERR_QUESTION_IN_SERVICE_DEF, ERR_QUESTION_ALREADY_USED:

                Call Response.Write("<BR />")
                Call RenderSvcConfigPath(aSvcConfigInfo)
                Call Response.Write("<BR /> <BR /> " & vbCrLf)
                Call Response.Write("<FORM ACTION=""services_modify_map.asp"" METHOD=""POST"">" & vbCrLf)
                Call RenderSvcConfigInputs(aSvcConfigInfo)
                Call Response.Write("<FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ ><P ALIGN=""CENTER""><B>")
                If lErr = ERR_QUESTION_IN_SERVICE_DEF Then
                    Call Response.Write(asDescriptors(883)) 'The question you have selected is already being used as part of the service definition, you cannot use it as an additional or alternate question.
                Else
                    If (aSvcConfigInfo(SVCCFG_QO_ID) = NEW_OBJECT_ID) Then
                        Call Response.Write(asDescriptors(885)) 'The question you have selected is already being used as an alternative question, you cannot use it as an additional question.
                    Else
                        Call Response.Write(asDescriptors(884)) 'The question you have selected is already being used as an additional question, you cannot use it as an alternate question.
                    End If
                End If

                Call Response.Write("</B></P><P>" & asDescriptors(799)) 'Please go back and select a different question.
                Call Response.Write("</P></FONT> " & vbCrLf)

                Call Response.Write("<INPUT name=back type=submit class=""buttonClass"" value=""< " & asDescriptors(334) & """></INPUT><BR />") 'Descriptor:Back)
                Call Response.Write("</FORM>")

            Case Else
                Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & "Services Overview", "services_overview.asp") 'Descriptor:Services Overview

            End Select

        Else

      %>
        <BR />
        <% Call RenderSvcConfigPath(aSvcConfigInfo) %>
        <BR />
        <BR />
        <FORM ACTION="services_modify_map.asp" METHOD=POST>
        <% RenderSvcConfigInputs(aSvcConfigInfo) %>
          <INPUT TYPE="HIDDEN" NAME="isid" VALUE="<%=aMapInfo(MAP_QO_IS)%>" />
          <INPUT TYPE="HIDDEN" NAME="pcnt" VALUE="<%=aMapInfo(MAP_QO_PROMPT_COUNT)%>" />

        <%If nStorageType = STORAGE_NONE Then %>
        <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
          <P ALIGN="CENTER"><B><%Call Response.Write(asDescriptors(798)) 'The question you have selected, cannot be use in the Subscription Portal. %></B></P>
          <%Call Response.Write(asDescriptors(799)) 'Please go back and select a different question.%>
        </FONT>

        <%ElseIf nStorageType = STORAGE_ALL Then %>
        <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
          <%Call Response.Write(asDescriptors(800)) 'The answer to the question you have selected, can be stored either on the subscription book repository or on some other location using a storage mapping.%><BR/>
          <%Call Response.Write(asDescriptors(801)) 'If you select to store the answer in the subscription book repository, the answer itself will be stored there. %><BR/>
          <%Call Response.Write(asDescriptors(802)) 'If you select to use a storage mapping, a Preference ID of the question will be stored in the subscription book repository, and the answers are stored based on the storage mapping definition; with the Preference ID you can link the tables in the storage mapping with the subscription book repository.%><BR />
          <BR/>
          <%Call Response.Write(asDescriptors(749)) 'Please select the location for the answers to this question.%><BR />
          <BR />
        </FONT>

        <%Else%>
        <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
          <%Call Response.Write(asDescriptors(803)) 'The answer to the question you have selected, has to be stored using a storage mapping. This is because it prompts the user to select from a list of elements, and the user may pick more than one.%><BR/>
          <%Call Response.Write(asDescriptors(804)) 'When you use a storage mapping to store the answers, a Preference ID of the question will be stored in the subscription book repository, and the answers are stored based on the storage mapping definition; with the Preference ID you can link the tables in the storage mapping with the subscription book repository.%><BR />
          <BR/>
          <%Call Response.Write(asDescriptors(805)) 'Please select the storage mapping to use with this question:%><BR />
          <BR />
        </FONT>
        <%End If%>

        <TABLE WIDTH="100%" BORDER=0 CELLPADDING=0 CELLSPACING=0>
        <%If nStorageType = STORAGE_ALL Then %>
          <TR>
            <TD VALIGN=MIDDLE WIDTH="1%"><INPUT TYPE="RADIO" NAME="mid" VALUE="sbr" <%If aMapInfo(MAP_ID) = "sbr" Then Response.Write "CHECKED" %> /></TD>
            <TD VALIGN=MIDDLE WIDTH="99%" COLSPAN="6">
              <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                <%Call Response.Write(Replace(asDescriptors(750), "#", "<B>" & asDescriptors(581) & "</B>")) 'Use the <B>subsciption book repoisitory</B>%>
              </FONT>
            </TD>
          </TR>

          <TR>
            <TD COLSPAN="7">
              <IMG SRC="../images/1ptrans.gif" HEIGHT="5" WIDTH="1">
            </TD>
          </TR>

          <TR>
            <TD COLSPAN="7">
              <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                <%Call Response.Write(asDescriptors(751)) '<B>OR </B>use one of the following storage mappings:%>
              </FONT>
            </TD>
          </TR>

          <TR>
            <TD COLSPAN="7">
              <IMG SRC="../images/1ptrans.gif" HEIGHT="5" WIDTH="1" BORDER="0" ALT="">
            </TD>
          </TR>
       <%End If%>

          <%If lCount > 0 Then %>
          <TR>
            <TD BGCOLOR="#c2c2c2" VALIGN=TOP WIDTH=15><IMG SRC="../images/1ptrans.gif" HEIGHT="1" WIDTH="15" BORDER="0" ALT=""></TD>
            <TD BGCOLOR="#c2c2c2" VALIGN="TOP">
              <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                <B><%Call Response.Write(asDescriptors(306)) 'Name%></B>
              </FONT>
            </TD>
            <TD BGCOLOR="#c2c2c2">&nbsp;&nbsp;</TD>

            <TD BGCOLOR="#c2c2c2" VALIGN="TOP">
              <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                <B><%Call Response.Write(asDescriptors(806)) 'DB Alias%></B>
              </FONT>
            </TD>
            <TD BGCOLOR="#c2c2c2">&nbsp;&nbsp;</TD>

            <TD BGCOLOR="#c2c2c2" VALIGN="TOP">
              <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                <B><%Call Response.Write(asDescriptors(752)) 'Tables%></B>
              </FONT>
            </TD>
            <TD BGCOLOR="#c2c2c2">&nbsp;&nbsp;</TD>

          </TR>

          <TR>
            <TD COLSPAN="7" BGCOLOR="#000000"><IMG SRC="../images/1ptrans.gif" HEIGHT="1" WIDTH="1" BORDER="0" ALT=""></TD>
          </TR>

            <%For i = 0 To lCount - 1%>
            <TR>
              <TD  WIDTH=15><INPUT TYPE="RADIO" NAME="mid" VALUE="<%=aMapList(i, MAP_ID)%>" <%If aMapList(i, MAP_ID) = aMapInfo(MAP_ID) Then Response.Write "CHECKED" %>/></TD>
              <TD><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>"><%=Server.HTMLEncode(aMapList(i, MAP_NAME))%></FONT></TD>
              <TD>&nbsp;&nbsp;</TD>
              <TD><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>"><%=Server.HTMLEncode(aMapList(i, MAP_DBALIAS))%></FONT></TD>
              <TD>&nbsp;&nbsp;</TD>
              <TD><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>"><%=Server.HTMLEncode(aMapList(i, MAP_DESC))%></FONT></TD>
              <TD><FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT)%>"><A HREF="services_map_tables.asp?mid=<%=Server.URLEncode(aMapList(i, MAP_ID))%>&mn=<%=Server.URLEncode(aMapList(i, MAP_NAME))%>&mf=<%=Server.URLEncode(aMapList(i, MAP_FILTER))%>&<%=CreateRequestForSvcConfig(aSvcConfigInfo)%>"><%Call Response.Write(asDescriptors(353)) 'edit%></A></FONT></TD>
            </TR>
            <TR>
              <TD COLSPAN="7" ALIGN="CENTER" HEIGHT="1" BGCOLOR="#6699CC"><img src="../images/1ptrans.gif" WIDTH="1" HEIGHT="1" BORDER="0" ALT=""></TD>
            </TR>
            <%Next%>

          </TR>
          <TR>
            <TD COLSPAN="7">
              <IMG SRC="../images/1ptrans.gif" HEIGHT="5" WIDTH="1" BORDER="0" ALT="">
            </TD>
          </TR>
        <%End If%>

        <%If nStorageType <> STORAGE_NONE Then %>
          <TR>
            <TD></TD>
            <TD COLSPAN="6">
              <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                <A HREF="services_map_tables.asp?mid=new&<%=CreateRequestForSvcConfig(aSvcConfigInfo)%>"><%Call Response.Write(asDescriptors(753)) 'Add a new Storage Mapping%></A>
              </FONT>
            </TD>
          </TR>
        <%End If%>
        </TABLE>

        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
          <TR>
            <TD COLSPAN="2">
              <BR />
            </TD>
          </TR>

          <TR>
            <TD ALIGN="left" NOWRAP WIDTH="1%">
              <INPUT name=back type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
            </TD>
            <TD ALIGN="left" NOWRAP WIDTH="98%">
              <%If Not ((nStorageType = STORAGE_NONE) Or ((nStorageType And STORAGE_SBR_ONLY) = 0 And lCount = 0)) Then %>
              <INPUT name=next type=submit class="buttonClass" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;
              <%End If%>
            </TD>
          </TR>
        </TABLE>
        </FORM>

      <%End If %>
    </TD>

    <TD WIDTH="1%"><IMG ALT="" BORDER=0 HEIGHT=1 SRC="../images/1ptrans.gif" width=21 ></TD>

    <TD WIDTH="1%"><!-- #include file="help_widget.asp" -->
    </TD>
  </TR>
</TABLE>
</BODY>
</HTML>
<%
	Erase aMapList
	Erase aPromptList
%>
