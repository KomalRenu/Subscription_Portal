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
<!-- #include file="../CustomLib/SiteConfigCuLib.asp" -->
<!-- #include file="dbaliases_widget.asp" -->
<%
Dim lStatus
Dim aMapInfo

Dim aDBAliases
Dim i, j, lTable
Dim lCount

Dim oMapDOM
Dim oMapTables
Dim sMapDBAlias

Dim sMapFilter

Dim sTables
Dim aTables
Dim aAllTables
Dim aSelectedTables
Dim sDBAlias

Dim bFound


    If lErr = NO_ERR Then
        lErr = ParseRequestForMapInfo(oRequest, aSvcConfigInfo, aMapInfo)
    End If

    If lErr = NO_ERR Then
        lErr = GetMapDOM(aMapInfo(MAP_ID), oMapDOM)

        If lErr = NO_ERR Then
            Set oMapTables = oMapDOM.selectNodes("//table")

            If oMapTables.length > 0 Then
                sMapDBAlias = oMapTables(0).getAttribute("connection")
            End If

            sMapFilter = oMapDOM.selectSingleNode("//mapping").getAttribute("f")

        End If
    End If

    If lErr = NO_ERR Then
        If Len(aMapInfo(MAP_DBALIAS)) = 0 Then
            aMapInfo(MAP_DBALIAS) = sMapDBAlias
        End If

        If Len(aMapInfo(MAP_FILTER)) = 0 Then
            aMapInfo(MAP_FILTER) = sMapFilter
        End If


        If Len(aMapInfo(MAP_DBALIAS)) > 0 Then
            lErr = GetTables(aMapInfo(MAP_DBALIAS), aMapInfo(MAP_FILTER), aTables)
        End If
    End If


    If lErr = NO_ERR Then

        If Len(aMapInfo(MAP_TABLES)) > 0 Then
            aMapInfo(MAP_TABLES) = Left(aMapInfo(MAP_TABLES), Len(aMapInfo(MAP_TABLES)) - 1)
            aSelectedTables = Split(aMapInfo(MAP_TABLES), ";")
        Else
            If aMapInfo(MAP_DBALIAS) = sMapDBAlias Then
                If oMapTables.length > 0 Then

                    lCount = oMapTables.length - 1
                    Redim aSelectedTables(lCount)

                    For i = 0 To lCount
                        aSelectedTables(i) = oMapTables(i).getAttribute("id")
                    Next
                End If
            End If
        End If

        If Not IsEmpty(aTables) Then
            lCount = UBound(aTables)
            If IsEmpty(aSelectedTables) Then
                Redim aAllTables(lCount)
            Else
                Redim aAllTables(lCount - UBound(aSelectedTables))
            End If

            lTable = 0
            For i = 0 To lCount
                If IsEmpty(aSelectedTables) Then
                    aAllTables(i) = aTables(i)
                Else
                    bFound = False
                    For j = 0 To UBound(aSelectedTables)
                        If aSelectedTables(j) = aTables(i) Then
                            bFound = True
                            Exit For
                        End If
                    Next

                    If bFound = False Then
                        aAllTables(lTable) = aTables(i)
                        lTable = lTable + 1
                    End If
                End If
            Next
        End If
    End If

    'Get a list of DBAliases from the Engine:
    If lErr = NO_ERR Then
        lErr = getDBAliases(aDBAliases)
        If lErr <> NO_ERR Then Call LogErrorXML(aConnectionInfo, lErr, "", MODULE_NAME, "", "", "Error calling getDBAliases", LogLevelTrace)

    End If



    If Len(aSvcConfigInfo(SVCCFG_AQ_ID)) > 0 Then
        If aSvcConfigInfo(SVCCFG_STEP) = STATIC_SS Then
            aPageInfo(S_TITLE_PAGE) = STEP_SERVICES_STATIC_MAP_TABLES & " " & asDescriptors(783) '"Select Tables"
        Else
            aPageInfo(S_TITLE_PAGE) = STEP_SERVICES_DYNAMIC_MAP_TABLES & " " & asDescriptors(783) '"Select Tables"e"
        End If
    Else
        aPageInfo(S_TITLE_PAGE) = STEP_SERVICES_DYNAMIC_TABLES & " " & asDescriptors(783) '"Select Tables"
    End If
    aPageInfo(N_CURRENT_OPTION_PAGE) = SECTION_SERVICES
    aPageInfo(N_OPTIONS_WITH_LINKS_PAGE) = "mid=" & Server.URLEncode(aMapInfo(MAP_ID)) & "&dba=" & Server.URLEncode(aMapInfo(MAP_DBALIAS)) & "&mf=" & Server.URLEncode(aMapInfo(MAP_FILTER)) & "&tbls=" & Server.URLEncode(aMapInfo(MAP_TABLES)) & "&mn=" & Server.URLEncode(aMapInfo(MAP_NAME)) & "&" & CreateRequestForSvcConfig(aSvcConfigInfo)

    aSvcConfigInfo(SVCCFG_MAP_FILTER) = ""

    lStatus = checkSiteConfiguration()

	sDBAlias = aMapInfo(MAP_DBALIAS)

%>
<HTML>
<HEAD>
  <%Response.Write(putMETATagWithCharSet())%>
  <TITLE><%Response.Write asDescriptors(248) 'Descriptor: Administrator Page%> - MicroStrategy Narrowcast Server</TITLE>

<!-- #include file="../NSStyleSheet.asp" -->
<SCRIPT>
<!--
function MoveItemsbyListObject(oFromList, oToList) {
	//**************************************************************************
	//Purpose:  move selected items from sFromList to sToList referenced by object
	//Input:    sFromList, sToList
	//Output:   sFromList, sToList
	//**************************************************************************
	var oForm = document.tables;
	var i;
	var lLength;
	var aSelArray = new Array();

	// add to right-side
	lLength = oFromList.options.length;
	for (i=0; i<lLength; i++)
	{
		if (oFromList.options[i].selected && oFromList.options[i].value!= "-none-")
		{
			if (oToList.options.length==1)	//replace -none- with an item
			{
				if (oToList.options[0].value=="-none-" )
				{
					oToList.options[0] = null;
				}
			}
			var oOption = new Option(oFromList.options[i].text, oFromList.options[i].value, false, false)
			oToList.options[oToList.length] = oOption;
			oToList.options[oToList.length-1].selected = false;
    		oOption = null;
    	}
    }

	//put left-side seleted into a temp array
	for (i=lLength-1;i>=0;i--)
	{
    	if (oFromList.options[i].selected)
		{
			oFromList.options[i].selected = false;
			aSelArray[aSelArray.length] = i;
		}
	}

	for (i=0; i<aSelArray.length; i++)
	{
		oFromList.options[aSelArray[i]] = null;
	}

	if (oFromList.options.length==0)	 //put -none- when no items
	{
		var oOption = new Option("<%=asDescriptors(103) %>", "-none-", false, false)
		oFromList.options[oFromList.length] = oOption;
		oFromList.options[oFromList.length-1].selected = false;
		oOption = null;
	}
	oFromList.options[0].selected = true;

	EnableNext();
}

function EnableNext() {
	//**************************************************************************
	//Purpose:  enables/disables the next button based on the number of tables selected.
	//Input:
	//Output:
	//**************************************************************************
	var oOptions;
	var oSeletedList;
	var oNext;

	if (eval('document.tables.selected') != null )
	{
		oSeletedList = document.tables.selected;
		oOptions = oSeletedList.options;

		oNext = document.tables.next;

		if (oOptions[0].value != "-none-")
		{
		    oNext.style.display = "";
		} else {
		    oNext.style.display = "none";
		}
	}
	return(true);
}

function BuildUserSelections() {
	//**************************************************************************
	//Purpose:  collect all the information user selected and put it to a hidden input
	//Input:
	//Output:
	//**************************************************************************
	var oForm = document.tables;
	var lLength;
	var i;
	var j;
	var oOptions;
	var oSeletedList

	if (eval('document.tables.selected') != null )
	{
		oSeletedList = eval('document.tables.selected');
		oOptions = oSeletedList.options;
		lLength = oOptions.length;
		if (oOptions[0].value != "-none-")
		{
			for(i=0; i<lLength; i++)
			{
				oForm.tbls.value = oForm.tbls.value + oOptions[i].value + ";" ;
			}
		}
	}
	document.filter.tbls.value = oForm.tbls.value;
	return(true);
}

// -->
</SCRIPT>
</HEAD>
<BODY BGCOLOR="FFFFFF" TOPMARGIN=0 LEFTMARGIN=0 ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT=0 MARGINWIDTH=0>
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
        <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & asDescriptors(783) , "services_map_tables.asp?" & CreateRequestForSvcConfig(aSvcConfigInfo)) 'Descriptor: Return to: 'Descriptor:Select Tables %>
      <%Else%>
        <BR />
        <% Call RenderSvcConfigPath(aSvcConfigInfo) %>
        <BR />
        <BR />
        <FONT FACE="<%=aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" >
          <%Call Response.Write(asDescriptors(755)) 'Please select the database tables where the information will be stored.%><BR />
        </FONT>
        <BR />

        <FORM NAME="filter" ACTION="services_map_tables.asp" METHOD="POST" onSubmit="return(BuildUserSelections());">
        <%Call RenderSvcConfigInputs(aSvcConfigInfo)%>
        <INPUT TYPE="HIDDEN" NAME="mid" VALUE="<%=aMapInfo(MAP_ID)%>" />

        <TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0">
          <TR>
            <TD COLSPAN="2">
              <!--Start DBAlias list: -->
              <%Call displayDBAliasWidget(aDBAliases, asDescriptors, "", REPOSITORY_WAREHOUSE)%>
              <!--End DBAlias list -->
            </TD>
           </TR>
          <TR>
            <TD HEIGHT="6"><IMG SRC="../images/1ptrans.gif" WIDTH="1" HEIGHT="10" BORDER="0" ALT=""></TD>
          </TR>
          <TR>
            <TD>
              <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_MEDIUM_FONT) %>" >
                <%Call Response.Write(asDescriptors(932)) 'Filter table names:%>&nbsp;
              </FONT>
            </TD>
            <TD><INPUT NAME="mf" VALUE="<%=Server.HTMLEncode(aMapInfo(MAP_FILTER))%>" CLASS="textBoxClass" /></TD>
          </TR>
           <TR>
            <TD>
			   <INPUT TYPE="SUBMIT" NAME="refresh" class="buttonclass" VALUE="<%Response.Write(asDescriptors(269))'Descriptor:Refresh%>" />&nbsp;
            </TD>
          </TR>
        </TABLE>
        <INPUT TYPE="HIDDEN" NAME="tbls" VALUE="" />
        </FORM>

        <FORM ACTION="services_map_columns.asp" NAME="tables" METHOD="POST" onSubmit="return(BuildUserSelections());">
        <INPUT TYPE="HIDDEN" NAME="mid" VALUE="<%=aMapInfo(MAP_ID)%>" />
        <INPUT TYPE="HIDDEN" NAME="dba" VALUE="<%=aMapInfo(MAP_DBALIAS)%>" />
        <INPUT TYPE="HIDDEN" NAME="mf" VALUE="<%=Server.HTMLEncode(aMapInfo(MAP_FILTER))%>" />
        <%Call RenderSvcConfigInputs(aSvcConfigInfo)%>

        <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
          <TR>
            <TD>
              <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                <%Call Response.Write(asDescriptors(758)) 'Available Tables:%>
              </FONT>
            </TD>
            <TD></TD>
            <TD>
              <FONT FACE="<%Response.Write aFontInfo(S_FAMILY_FONT)%>" SIZE="<%=aFontInfo(N_SMALL_FONT) %>" >
                <%Call Response.Write(asDescriptors(759)) 'Selected Tables:%>
              </FONT>
            </TD>
          </TR>

          <TR>
            <TD>
              <SELECT MULTIPLE=1 NAME="all" SIZE="10" class="pullDownClass" WIDTH="40">
                <%
                If Not IsEmpty(aAllTables) Then
                    lCount = UBound(aAllTables)
                    For i = 0 To lCount
                        If Len(aAllTables(i)) > 0 Then
                            Response.Write("<OPTION VALUE=""" & aAllTables(i) & """>" & aAllTables(i) & "</OPTION>")
                        End If
                    Next
                Else
                    Response.Write "<OPTION VALUE=""-none-"">" & asDescriptors(103) & "</OPTION>"'Descriptor = -- none --
                End If
                %>
              </SELECT>
            </TD>
            <TD WIDTH="29" VALIGN="MIDDLE" ALIGN="CENTER">
              <A HREF="javascript:MoveItemsbyListObject(document.tables.all, document.tables.selected)"><IMG SRC="../images/btn_add.gif" WIDTH="25" HEIGHT="25" BORDER="0" ALT=""/></A><BR />
              <BR />
              <A HREF="javascript:MoveItemsbyListObject(document.tables.selected, document.tables.all)"><IMG SRC="../images/btn_remove.gif" WIDTH="25" HEIGHT="25" BORDER="0" ALT=""/></A>
            </TD>
            <TD>
              <SELECT MULTIPLE=1 NAME="selected" SIZE="10" class="pullDownClass" WIDTH="40">
                <%
                If Not IsEmpty(aSelectedTables) Then
                    lCount = UBound(aSelectedTables)
                    For i = 0 To lCount
                        If Len(aSelectedTables(i)) > 0 Then
                            Response.Write("<OPTION VALUE=""" & aSelectedTables(i) & """>" & aSelectedTables(i) & "</OPTION>")
                        End If
                    Next
                Else
                    Response.Write "<OPTION VALUE=""-none-"">" & asDescriptors(103) & "</OPTION>"'Descriptor = -- none --
                End If
                %>
              </SELECT>
              <INPUT TYPE="HIDDEN" NAME="tbls" VALUE="" />
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
            <TD ALIGN="left" NOWRAP WIDTH="1%">
              <INPUT name=back type=submit class="buttonClass" value="<%Response.Write(asDescriptors(334)) 'Descriptor:Back%>"></INPUT> &nbsp;
            </TD>
            <TD ALIGN="left" NOWRAP WIDTH="98%">
              <INPUT name=next type=submit <%If IsEmpty(aSelectedTables) Then Response.Write(" style=""display:none"" ") %> class="buttonClass" value="<%Response.Write(asDescriptors(335)) 'Descriptor:Next%>"></INPUT> &nbsp;
            </TD>
          </TR>
        </TABLE>
        </FORM>
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
	Erase aDBAliases
	Erase aTables
	Erase aAllTables
	Erase aSelectedTables

	Set oMapDOM = Nothing
	Set oMapTables = Nothing


%>