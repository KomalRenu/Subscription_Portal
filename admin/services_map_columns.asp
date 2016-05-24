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

Dim oMapDOM
Dim oColumn
Dim aMapInfo

Dim aTables
Dim aColumns
Dim aGUIDs
Dim aTablesInfo
Dim sField
Dim sValue

Dim aPromptList

Dim lCount, lColumnCount
Dim i, j, k, lIndex
Dim sZoneName

Dim aNormalQuestions
Dim aSlicingQuestions
Dim aExtraQuestions
Dim sQuestionID

    If lErr = NO_ERR Then
        lErr = ParseRequestForMapInfo(oRequest, aSvcConfigInfo, aMapInfo)
    End If

    If Len(oRequest("back")) > 0 Then

        If Len(aSvcConfigInfo(SVCCFG_QO_ID)) > 0 Then
            Response.Redirect("services_select_map.asp?mid=" & aMapInfo(MAP_ID) & "&" & CreateRequestForSvcConfig(aSvcConfigInfo))
        Else
            Response.Redirect("services_subsset.asp?" & CreateRequestForSvcConfig(aSvcConfigInfo))
        End If
    End If

    'Get MapInfo:
    If lErr = NO_ERR Then
        lErr = GetMapDOM(aMapInfo(MAP_ID), oMapDOM)
    End If

    'Get the list of configured tables:
    If lErr = NO_ERR Then
        If Len(aMapInfo(MAP_TABLES)) > 0 Then
            aTables = Split(aMapInfo(MAP_TABLES), ";")
            Redim Preserve aTables(UBound(aTables) - 1)
            lErr = GetTablesInfo(aMapInfo(MAP_DBALIAS), aTables, aTablesInfo)
        End If
    End If

    'Get the Map Name
    If lErr = NO_ERR Then
        If Len(aMapInfo(MAP_NAME)) = 0 Then
            If Len(aMapInfo(MAP_TABLES)) > 0 Then

                lCount = UBound(aTables)
                For i = 0 To lCount
                    aMapInfo(MAP_NAME) = aMapInfo(MAP_NAME) & aTables(i) & ", "
                Next

                aMapInfo(MAP_NAME) = Left(aMapInfo(MAP_NAME), Len(aMapInfo(MAP_NAME)) - 2)
            Else
                aMapInfo(MAP_NAME) = "New Mapping"
            End If
        End If
    End If

    'if this is the mapping of the question objects, the the Details for questions,
    'if not, get the subscription set configuration:
    If lErr = NO_ERR Then
        If Len(aSvcConfigInfo(SVCCFG_AQ_ID)) > 0 Then
            lErr = GetPromptsForQuestion(aSvcConfigInfo, aPromptList)
        Else
            lErr = GetSubscriptionSetConfig(aSvcConfigInfo, aNormalQuestions, aSlicingQuestions, aExtraQuestions)
        End If
    End If

    'Set the PageInfo to be used by the navigator bar and the header.
    If Len(aSvcConfigInfo(SVCCFG_AQ_ID)) > 0 Then
        If aSvcConfigInfo(SVCCFG_STEP) = STATIC_SS Then
            aPageInfo(S_TITLE_PAGE) = STEP_SERVICES_STATIC_MAP_COLUMNS & " " & asDescriptors(784) '"Select Columns"
        Else
            aPageInfo(S_TITLE_PAGE) = STEP_SERVICES_DYNAMIC_MAP_COLUMNS & " " & asDescriptors(784)'"Select Columns"
        End If
    Else
        aPageInfo(S_TITLE_PAGE) = STEP_SERVICES_DYNAMIC_COLUMNS & " " & asDescriptors(784) '"Select Columns"
    End If
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SERVICES
    aPageInfo(N_OPTIONS_WITH_LINKS_PAGE) = "mid=" & Server.URLEncode(aMapInfo(MAP_ID)) & "&dba=" & Server.URLEncode(aMapInfo(MAP_DBALIAS)) & "&mf=" & Server.URLEncode(aMapInfo(MAP_FILTER)) & "&tbls=" & Server.URLEncode(aMapInfo(MAP_TABLES)) & "&mn=" & Server.URLEncode(aMapInfo(MAP_NAME)) & "&" & CreateRequestForSvcConfig(aSvcConfigInfo)
    aSvcConfigInfo(SVCCFG_MAP_NAME) = ""

    lStatus = checkSiteConfiguration()



%>
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>

<SCRIPT LANGUAGE="JavaScript" SRC="../js/DHTMLapi.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../js/DNDapi.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
//OnLoad method of the page:
function init(){
var zone, element, i;

    //Create the page by zone:
    //Create the atts zone:
    zone = new objDNDZone("general");
    zone.dragMode = DRAG;
    zone.dropMode = NO_DROP;
    zones["general"] = zone;

    <%If Len(aSvcConfigInfo(SVCCFG_AQ_ID)) > 0 Then %>
    zone.add("subscription.<%=ANSWER_SUBSCRIPTION_ID%>", "<%=ANSWER_SUBSCRIPTION_ID%>");
    zone.add("subscription.<%=ANSWER_USER_ID%>", "<%=ANSWER_USER_ID%>");
    <%Else%>
    zone.add("subscription.<%=ANSWER_SUBSCRIPTION_ID%>", "<%=ANSWER_SUBSCRIPTION_ID%>");
    zone.add("subscription.<%=ANSWER_USER_ID%>", "<%=ANSWER_USER_ID%>");
    zone.add("subscription.<%=ANSWER_ACCOUNT_ID%>", "<%=ANSWER_ACCOUNT_ID%>");
    zone.add("subscription.<%=ANSWER_ADD_TRANS_PROPS%>", "<%=ANSWER_ADD_TRANS_PROPS%>");
    zone.add("subscription.<%=ANSWER_ADDRESS_ID%>", "<%=ANSWER_ADDRESS_ID%>");
    zone.add("subscription.<%=ANSWER_CREATED_BY%>", "<%=ANSWER_CREATED_BY%>");
    zone.add("subscription.<%=ANSWER_CREATED_DATE%>", "<%=ANSWER_CREATED_DATE%>");
    zone.add("subscription.<%=ANSWER_EXPIRATION_DATE%>", "<%=ANSWER_EXPIRATION_DATE%>");
    zone.add("subscription.<%=ANSWER_LAST_MOD_BY%>", "<%=ANSWER_LAST_MOD_BY%>");
    zone.add("subscription.<%=ANSWER_LAST_MOD_DATE%>", "<%=ANSWER_LAST_MOD_DATE%>");
    zone.add("subscription.<%=ANSWER_STATUS%>", "<%=ANSWER_STATUS%>");
    zone.add("subscription.<%=ANSWER_SUBSCRIPTION_GUID%>", "<%=ANSWER_SUBSCRIPTION_GUID%>");
    zone.add("subscription.<%=ANSWER_SUBSCRIPTION_SET_ID%>", "<%=ANSWER_SUBSCRIPTION_SET_ID%>");
    zone.add("subscription.<%=ANSWER_TRANS_PROPS_ID%>", "<%=ANSWER_TRANS_PROPS_ID%>");
    <%End If%>

    zone.render();

    <%
    If Not IsEmpty(aPromptList) Then
        lCount = UBound(aPromptList)
        sZoneName = "question"

        Response.Write "zone = new objDNDZone(""" & sZoneName & """);" & vbCrLf
        Response.Write "zones[""" & sZoneName & """] = zone;" & vbCrLf
        Response.Write "zone.dragMode = DRAG;" & vbCrLf
        Response.Write "zone.dropMode = NO_DROP;" & vbCrLf

        Response.Write "zone.add(""qo." & aSvcConfigInfo(SVCCFG_AQ_ID) & "." & ANSWER_PREFERENCE_ID & """, """ & Replace(asDescriptors(832), "#", Server.HTMLEncode(aSvcConfigInfo(SVCCFG_AQ_NAME))) & """); " & vbCrLf 'Descriptor: Preference ID for:

        For i = 0 To lCount
            Response.Write "zone.add(""qo." & aSvcConfigInfo(SVCCFG_AQ_ID) & "." & ANSWER_PROMPT_ANSWER & "." & i + 1 & """, """ & Replace(asDescriptors(833), "#", Server.HTMLEncode(aPromptList(i, PROMPT_TITLE))) & """);" & vbCrLf 'Descriptor: Answer for:
        Next

        Response.Write "zone.render();" & vbCrLf & vbCrLf

    End If
    %>

    <%
    If Not IsEmpty(aExtraQuestions) Then
        lCount = UBound(aExtraQuestions)
        sZoneName = "questions"

        Response.Write "zone = new objDNDZone(""" & sZoneName & """);" & vbCrLf
        Response.Write "zones[""" & sZoneName & """] = zone;" & vbCrLf
        Response.Write "zone.dragMode = DRAG;" & vbCrLf
        Response.Write "zone.dropMode = NO_DROP;" & vbCrLf


        For i = 0 To lCount
            If aExtraQuestions(i, QO_MAP_ID) = "sbr" Then
                Response.Write "zone.add(""qo." & aExtraQuestions(i, QO_ALTERNATE_ID) & "." & ANSWER_PROMPT_ANSWER & ".1"", """ & Replace(asDescriptors(833), "#", Server.HTMLEncode(aExtraQuestions(i, QO_NAME))) & """);" & vbCrLf 'Descriptor: Answer for:
            Else
                Response.Write "zone.add(""qo." & aExtraQuestions(i, QO_ALTERNATE_ID) & "." & ANSWER_PREFERENCE_ID & """, """ & Replace(asDescriptors(832), "#", Server.HTMLEncode(aExtraQuestions(i, QO_NAME))) & """); " & vbCrLf 'Descriptor: Preference ID For:
            End If
        Next

        Response.Write "zone.render();" & vbCrLf & vbCrLf

    End If
    %>

    <%
    If Not IsEmpty(aTablesInfo) Then
        lCount = UBound(aTablesInfo)

        For i = 0 To lCount
            aColumns = Split(aTablesInfo(i, TABLE_COLUMNS), ";")
            aGUIDs = Split(aTablesInfo(i, TABLE_COLUMN_GUIDS), ";")
            lColumnCount = UBound(aColumns) - 1

            For j = 0 To lColumnCount
                sZoneName = aGUIDs(j)

                Response.Write "zone = new objDNDZone(""" & sZoneName & """);" & vbCrLf
                Response.Write "zones[""" & sZoneName & """] = zone;" & vbCrLf
                Response.Write "zone.dragMode = DRAG_REMOVE;" & vbCrLf
                Response.Write "zone.dropMode = DROP;" & vbCrLf
                Response.Write "zone.maxElements = 1;" & vbCrLf

                Set oColumn = oMapDOM.selectSingleNode("//table[@id='" & aTablesInfo(i, TABLE_ID) & "']/col[@id='" & aColumns(j) & "']")
                If Not oColumn Is Nothing Then
                    sField = oColumn.getAttribute("field")

                    If Len(sField) > 0 Then

                        If Left(sField, 13) = "subscription." Then
                            Response.Write "zone.add(""" & sField & """, """ & Mid(sField, 14) & """);"

                        ElseIf Left(sField, 3) = "qo." Then

                            If Len(aSvcConfigInfo(SVCCFG_AQ_ID)) > 0 Then
                                If Mid(sField, Len(sField) - 1, 1) = "." Then
                                    lIndex = Clng(Mid(sField, Len(sField)))
                                    Response.Write "zone.add(""qo." & aSvcConfigInfo(SVCCFG_AQ_ID) & "." & ANSWER_PROMPT_ANSWER & "." & lIndex & """, """ & Replace(asDescriptors(833), "#", Server.HTMLEncode(aPromptList(lIndex - 1, PROMPT_TITLE))) & """);" & vbCrLf 'Descriptor: Answer for:
                                Else
                                    Response.Write "zone.add(""qo." & aSvcConfigInfo(SVCCFG_AQ_ID) & "." & ANSWER_PREFERENCE_ID & """, """ & Replace(asDescriptors(832), "#", Server.HTMLEncode(aSvcConfigInfo(SVCCFG_AQ_NAME))) & """); " & vbCrLf 'Descriptor: Preference ID for:
                                End If
                            Else

                                If Not IsEmpty(aExtraQuestions) Then
                                    sQuestionID = Mid(sField, 4, 32)

                                    For k = 0 To UBound(aExtraQuestions)
                                        If aExtraQuestions(k, QO_ALTERNATE_ID) = sQuestionID Then
                                            If Mid(sField, Len(sField) - 1, 1) = "." Then
                                                lIndex = Clng(Mid(sField, Len(sField)))
                                                Response.Write "zone.add(""qo." & sQuestionID & "." & ANSWER_PROMPT_ANSWER & "." & lIndex & """, """ & Replace(asDescriptors(833), "#", Server.HTMLEncode(aExtraQuestions(k, QO_NAME))) & """);" & vbCrLf 'Descriptor: Answer for:
                                            Else
                                                Response.Write "zone.add(""qo." & sQuestionID & "." & ANSWER_PREFERENCE_ID & """, """ & Replace(asDescriptors(832), "#", Server.HTMLEncode(aExtraQuestions(k, QO_NAME))) & """); " & vbCrLf 'Descriptor: Preference ID For:
                                            End If
                                        End If
                                    Next

                                End If

                            End If

                        ElseIf sField = ANSWER_CONSTANT Then
                            sValue = Server.HTMLEncode(oColumn.getAttribute("value"))
                            Response.Write "zone.add(""" & sValue & """, """ & sValue & """);"

                        Else
                            sValue = Server.HTMLEncode(sField)
                            Response.Write "zone.add(""" & sValue & """, """ & sValue & """);"

                        End If

                    End If
                End If

                Response.Write "zone.render();" & vbCrLf
            Next
        Next
    End If
    %>

    //Create the page by zone:
    //Create the atts zone:
    zone = new objDNDZone("custom");
    zone.dragMode = DRAG;
    zone.dropMode = NO_DROP;
    zones["custom"] = zone;
    zone.add("<%Call Response.Write(asDescriptors(808)) 'custom%>", "<%Call Response.Write(asDescriptors(808)) 'custom%>");
    zone.render();

    //Call the initDND() function:
    initDND();

}

//Method for updating the custom div
function customChange(element) {
var oDiv;
var oElem;
var txt;

    txt = element.value;

    if (txt != "") {
        oElem = zones["custom"].elements[0];
        oElem.value = txt;
        oElem.caption = txt;

        oDiv = new objIDiv("custom" + DELIMITER + "0");
        oDiv.setInnerHTML(txt);
    }
}

function expandItems(sImageName) {

var sSrc;

    sSrc = document.images[sImageName].src;

    if (sSrc.search("arrow_down") != -1) {
        removeObj(sImageName.substr(6));
        document.images[sImageName].src = "../images/arrow_right.gif";
    } else {
        displayObj(sImageName.substr(6));
        document.images[sImageName].src = "../images/arrow_down.gif";
    }

}

function scrollWindow() {
var obj;

    obj = getObj("components");

    if (document.body.scrollLeft > 168) {
        obj.style.left = document.body.scrollLeft + 10;
    } else {
        obj.style.left = 178;
    }

    if (document.body.scrollTop > 250) {
        obj.style.top = document.body.scrollTop + 10;
    } else {
        obj.style.top = "";
    }

}

//-->
</SCRIPT>

<LINK rel="stylesheet" type="text/css" href="../js/DND.css" />
<!-- #include file="../NSStyleSheet.asp" -->

</HEAD>
<BODY onscroll="scrollWindow();" BGCOLOR="FFFFFF" TOPMARGIN=0 LEFTMARGIN=0 ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT=0 MARGINWIDTH=0 onLoad="init();">
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" HEIGHT="100%">
  <TR>
    <TD COLSPAN="6" HEIGHT="1%">
      <!-- begin header -->
        <!-- #include file="admin_header.asp" -->
      <!-- end header -->
    </TD>
  </TR>
  <TR>
    <TD WIDTH="1%" valign="TOP">
      <!-- begin toolbar -->
        <!-- #include file="_toolbar_services.asp" -->
      <!-- end toolbar -->
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="96%" valign="TOP">
      <%If lErr <> 0 Then %>
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(579) , "devices_config.asp") 'Descriptor: Return to: 'Descriptor:Site Preferences %>
      <%Else%>
        <BR />
        <% Call RenderSvcConfigPath(aSvcConfigInfo) %>
        <BR />
        <BR />
        <FORM ACTION="services_map_save.asp">
        <% RenderSvcConfigInputs(aSvcConfigInfo) %>
        <INPUT TYPE="HIDDEN" NAME="mid" VALUE="<%=aMapInfo(MAP_ID)%>" />
        <INPUT TYPE="HIDDEN" NAME="dba" VALUE="<%=aMapInfo(MAP_DBALIAS)%>" />
        <INPUT TYPE="HIDDEN" NAME="mf" VALUE="<%=aMapInfo(MAP_FILTER)%>" />
        <TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH="100%">
          <TR>
            <TD>
            <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
              <%Call Response.Write(asDescriptors(760)) 'Please select the table columns below where you wisth to store the information.%><BR>
              <%Call Response.Write(asDescriptors(761)) 'Drag and drop components to the left onto table columns on the right.%><BR />
              <%If Len(aSvcConfigInfo(SVCCFG_AQ_ID)) > 0 Then
                    Call Response.Write(asDescriptors(886)) 'In order to edit or delete subscriptions, please make sure that all tables include either the SUBSCRIPTION_ID or the Preference ID of the question.
                Else
                    Call Response.Write(asDescriptors(887)) 'In order to edit or delete subscriptions, please make sure that all tables include the SUBSCRIPTION_ID.
                End IF%>
            </FONT>
            </TD>
          </TR>

         <%If Len(aSvcConfigInfo(SVCCFG_AQ_ID)) > 0 Then %>
          <TR>
            <TD HEIGHT="15"><IMG SRC="../images/1ptrans.gif" HEIGHT="15" WIDTH="1"></TD>
          </TR>

          <TR>
            <TD>
              <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
                <TR>
                  <TD>
                    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                      <%Call Response.Write(asDescriptors(793)) 'Save this storage mapping as:%>&nbsp;
                    </FONT>
                  </TD>
                  <TD>
                    <INPUT NAME=mn class="textBoxClass" SIZE=40 VALUE="<%=Server.HTMLEncode(aMapInfo(MAP_NAME))%>"></INPUT>
                  </TD>
                </TR>
              </TABLE>
            </TD>
          </TR>
          <%Else%>
            <INPUT NAME=mn TYPE="HIDDEN" VALUE="<%=aMapInfo(MAP_NAME)%>"></INPUT>
          <%End If%>

        </TABLE><BR />

        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
          <TR>
            <TD VALIGN="TOP" WIDTH="200">
              <DIV ID="components" STYLE="position:absolute;left:178;width:200;background-color:#d2d2d2;border:'1pt solid #6c6c6c'">
              <TABLE BORDER=0 CELLPADDING=3 CELLSPACING=0>
                <TR>
                  <TD COLSPAN="3">
                    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                      <%If Len(aSvcConfigInfo(SVCCFG_AQ_ID)) > 0 Then%>
                        <B><%Call Response.Write(asDescriptors(762)) 'Answer Components%></B>
                      <%Else%>
                        <B><%Call Response.Write(asDescriptors(764)) 'Subscription Components%></B>
                      <%End If%>
                    </FONT>
                    <BR />
                  </TD>
                </TR>

                <TR>
                  <TD>
                    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                     <A HREF="javascript:expandItems('arrow_general');"><IMG SRC="../images/arrow_down.gif" NAME="arrow_general" WIDTH="13" HEIGHT="13" ALT="" BORDER="" ALIGN="LEFT"/></A><%Call Response.Write(asDescriptors(763)) 'General:%><BR />
                    </FONT>
                    <!--The DIV of the general zone-->
                    <DIV ID="general"></DIV>
                  </TD>
                </TR>

                <%If Len(aSvcConfigInfo(SVCCFG_AQ_ID)) > 0 Then%>

                <TR>
                  <TD>
                    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                     <A HREF="javascript:expandItems('arrow_question');"><IMG SRC="../images/arrow_down.gif" NAME="arrow_question" WIDTH="13" HEIGHT="13" ALT="" BORDER="" ALIGN="LEFT"/></A><%=Server.HTMLEncode(aSvcConfigInfo(SVCCFG_AQ_NAME))%>:<BR />
                    </FONT>
                    <!--The DIV of the question: zone-->
                    <DIV ID="question"></DIV>
                  </TD>
                </TR>

                <%Else%>
                <TR>
                  <TD>
                    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                     <A HREF="javascript:expandItems('arrow_questions');"><IMG SRC="../images/arrow_down.gif" NAME="arrow_questions" WIDTH="13" HEIGHT="13" ALT="" BORDER="" ALIGN="LEFT"/></A>Questions:<BR />
                    </FONT>
                    <!--The DIV of the questions: zone-->
                    <DIV ID="questions"><%If IsEmpty(aExtraQuestions) Then %>
                      <TABLE BORDER = 0 CELLSPACING=5 CELLPADDING=0>
                        <TR>
                          <TD>
                            <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                              <I><%Call Response.Write(asDescriptors(888)) 'This service has no Questions to Map%></I>
                            </FONT>
                          </TD>
                        </TR>
                      </TABLE>
                    <%End If%></DIV>
                    </FONT>
                  </TD>
                </TR>

                <%End If%>

                <TR>
                  <TD>
                    <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                     <A HREF="javascript:expandItems('arrow_mainCustom');"><IMG SRC="../images/arrow_down.gif" NAME="arrow_mainCustom" WIDTH="13" HEIGHT="13" ALT="" BORDER="" ALIGN="LEFT"/></A><%Call Response.Write(asDescriptors(807)) 'Custom value:%><BR />
                    </FONT>
                    <!--The DIV of the custom zone-->
                    <DIV ID="mainCustom">
                      <TABLE CELLSPACING=0 CELLPADDING=0>
                        <TR>
                          <TD><INPUT NAME="cst" VALUE="<%Call Response.Write(asDescriptors(808)) 'custom%>" CLASS="textBoxClass" onkeyup="customChange(this);" /></TD>
                        </TR>
                        <TR>
                          <TD><DIV ID="custom"></DIV></TD>
                        </TR>
                      </TABLE>
                    </DIV>
                  </TD>
                </TR>

              </TABLE>
              </DIV>
              <IMG SRC="../images/1ptrans.gif" height="1" WIDTH="200">
            </TD>

            <TD WIDTH="20"><IMG SRC="../images/1ptrans.gif" height="1" WIDTH="20"></TD>

            <TD VALIGN="TOP">
              <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
                <%
                If Not IsEmpty(aTablesInfo) Then
                    lCount = UBound(aTablesInfo)

                    For i = 0 To lCount
                        aColumns = Split(aTablesInfo(i, TABLE_COLUMNS), ";")
                        aGUIDs = Split(aTablesInfo(i, TABLE_COLUMN_GUIDS), ";")
                        lColumnCount = UBound(aColumns) - 1

                        Response.Write "<TR><TD><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ >" & vbCrLf
                        Response.Write "Table:<B>&nbsp;" & aTablesInfo(i, 0) & "</B><BR />" & vbCrLf
                        Response.Write "</FONT><TABLE BORDER=0 CELLPADDING=4 CELLSPACING=1 BGCOLOR=""#c2c2c2"">" & vbCrLf

                        Response.Write ("<TR>"  & vbCrLf)
                        For j = 0 To lColumnCount
                            Response.Write "<TD BGCOLOR=""#000066""><FONT FACE=""" & aFontInfo(S_FAMILY_FONT) & """ SIZE=""" & aFontInfo(N_SMALL_FONT) & """ COLOR=""#ffffff"">" & vbCrLf
                            Response.Write "<B>" & aColumns(j) & "</B><BR /></FONT></TD>" & vbCrLf
                        Next
                        Response.Write ("</TR>" & vbCrLf)

                        Response.Write ("<TR>" & vbCrLf)
                        For j = 0 To lColumnCount
                            sZoneName = aGUIDs(j)
                            Response.Write "<TD BGCOLOR=""#ffffff"">" & vbCrLf
                            Response.Write "<DIV ID=""" & sZoneName & """></DIV>" & vbCrLf
                            Response.Write "</TD>" & vbCrLf
                        Next
                        Response.Write ("</TR>" & vbCrLf)

                        Response.Write "</TABLE><BR /><BR />" & vbCrLf

                        Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""c" & aTablesInfo(i, 0) & """ VALUE="""
                        For j = 0 To lColumnCount
                            Response.Write aColumns(j) & ";"
                        Next
                        Response.Write """ />" & vbCrLf

                        Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""g" & aTablesInfo(i, 0) & """ VALUE="""
                        For j = 0 To lColumnCount
                            Response.Write aGUIDs(j) & ";"
                        Next
                        Response.Write """ />" & vbCrLf

                        Response.Write "</TD></TR>" & vbCrLf
                    Next

                    Response.Write "<TR><TD>&nbsp;" & vbCrLf
                    Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""tbls"" VALUE="""
                    For i = 0 To lCount
                        Response.Write aTablesInfo(i, 0) & ";"
                    Next
                    Response.Write """ />" & vbCrLf
                    Response.Write "</TD></TR>" & vbCrLf

                End If
                %>
              </TABLE>
            </TD>
          </TR>
        </TABLE>

        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
          <TR>
            <TD COLSPAN="2">
              <BR />
            </TD>
          </TR>


          <TR>
            <TD ALIGN="right" NOWRAP WIDTH="387">
              <INPUT name=back type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
            </TD>
            <TD ALIGN="left" NOWRAP >
              <INPUT name=next type=submit class="buttonClass" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT>
            </TD>
          </TR>
        </TABLE>
        </FORM>

        <!--We need to create a DIV with id=drag-->
        <DIV class=DNDDrag id=drag>abc</DIV>

      <%End If %>
    </TD>

    <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

    <TD WIDTH="1%">
        <!-- #include file="help_widget.asp" -->
    </TD>
  </TR>
</TABLE>
</BODY>
</HTML>
<%
	Set oMapDOM = Nothing
	Set oColumn = Nothing

	Erase aTables
	Erase aColumns
	Erase aGUIDs
	Erase aTablesInfo
%>