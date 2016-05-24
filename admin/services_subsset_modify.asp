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

Dim aNormalQuestions
Dim aSlicingQuestions
Dim aExtraQuestions
Dim sAction
Dim sAid
Dim sMid

Dim lCount
Dim i
Dim sId


    If lErr = NO_ERR Then
        lErr = ParseRequestForSubsSet(oRequest, aSvcConfigInfo, aNormalQuestions, aSlicingQuestions, aExtraQuestions, sAction)
    End If

    'Check for actions cancelled:
    If sAction = "back" Then
        If lErr = NO_ERR Then
            Call DeleteCache(GetSvcConfigCacheName(aSvcConfigInfo), SVC_CONFIG_CACHE_FOLDER)
            Erase aNormalQuestions
			Erase aSlicingQuestions
			Erase aExtraQuestions
                
            If aSvcConfigInfo(SVCCFG_STEP) = DYNAMIC_SS Then                
                Response.Redirect("services_dynamic.asp?" & CreateRequestForSvcConfig(aSvcConfigInfo))
            Else
                Response.Redirect("services_static.asp?" & CreateRequestForSvcConfig(aSvcConfigInfo))
            End If
        End If
    End If
    
    'If it was none of the previous actions, we must save
    'the current settings in cache, so we can retrieve them later:
    If lErr = NO_ERR Then
        lErr = CreateSubscriptionsSetCache(aSvcConfigInfo, aNormalQuestions, aSlicingQuestions, aExtraQuestions)
    End If    
    
    'Check for default actions
    If lErr = NO_ERR Then
        If sAction = "next" Then
			Erase aNormalQuestions
			Erase aSlicingQuestions
			Erase aExtraQuestions
			
            If aSvcConfigInfo(SVCCFG_STEP) = DYNAMIC_SS Then
                Response.Redirect("services_map_tables.asp?" & CreateRequestForSvcConfig(aSvcConfigInfo))
            Else
                Response.Redirect("services_subsset_save.asp?" & CreateRequestForSvcConfig(aSvcConfigInfo))
            End If
            
        ElseIf sAction = "addqo" Then
			Erase aNormalQuestions
			Erase aSlicingQuestions
			Erase aExtraQuestions
			
            Response.Redirect("services_select_qo.asp?qid=new" & "&" & CreateRequestForSvcConfig(aSvcConfigInfo))
            
        ElseIf Left(sAction, 1) = "b" Then
            lCount = UBound(aSlicingQuestions)
            sId = Mid(sAction, 2)
            
            For i = 0 To lCount
                If aSlicingQuestions(i, QO_ID) = sId Then			
					
					sAid =  aSlicingQuestions(i, QO_ALTERNATE_ID)
                
					Erase aNormalQuestions
					Erase aSlicingQuestions
					Erase aExtraQuestions
					
                    Response.Redirect("services_select_qo.asp?qid=" & sId & "&aid=" & sAid & "&" & CreateRequestForSvcConfig(aSvcConfigInfo))
                End If
            Next

        ElseIf Left(sAction, 1) = "e" Then
            lCount = UBound(aExtraQuestions)
            sId = Mid(sAction, 2)
            
            For i = 0 To lCount
                If aExtraQuestions(i, QO_ALTERNATE_ID) = sId Then
					sMid = aExtraQuestions(i, QO_MAP_ID)
                
					Erase aNormalQuestions
					Erase aSlicingQuestions
					Erase aExtraQuestions
					
                    Response.Redirect("services_select_map.asp?qid=new" & "&aid=" & sId & "&mid=" & sMid & "&" & CreateRequestForSvcConfig(aSvcConfigInfo))
                End If
            Next

        ElseIf Left(sAction, 1) = "r" Then
            lCount = UBound(aExtraQuestions)
            sId = Mid(sAction, 2)
            
            For i = 0 To lCount
                If aExtraQuestions(i, QO_ALTERNATE_ID) = sId Then
					Erase aNormalQuestions
					Erase aSlicingQuestions
					Erase aExtraQuestions
					
                    Response.Redirect("services_remove_question.asp?qid=" & "&aid=" & sId & "&" & CreateRequestForSvcConfig(aSvcConfigInfo))
                End If
            Next
        End If
    End If    

    'Get the Channels list request from the request object:
    aPageInfo(S_NAME_PAGE) = "services_subsset.asp"
    aPageInfo(S_TITLE_PAGE) = STEP_PREFERENCES & " " & asDescriptors(286) 'Descriptor:Preferences
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
<BODY BGCOLOR="FFFFFF" TOPMARGIN=0 LEFTMARGIN=0 ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT=0 MARGINWIDTH=0>
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" HEIGHT="100%">
  <TR>
    <TD VALIGN="TOP">
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
            <%  Call DisplayAdminError(sErrorHeader, sErrorMessage, lErr, asDescriptors(250) & " " & "Overview" , "services_overview.asp") 'Descriptor: Return to: 'Descriptor:Overview %>
          </TD>
          
          <TD WIDTH="1%"><IMG SRC="../images/1ptrans.gif" WIDTH="21" HEIGHT="1" ALT="" BORDER="0" /></TD>

          <TD WIDTH="1%" VALIGN="TOP">
              <!-- #include file="help_widget.asp" -->
          </TD>
        </TR>
      </TABLE>
    </TD>
  </TR>
</TABLE>
</BODY>
</HTML>
<%
	Erase aNormalQuestions
	Erase aSlicingQuestions
	Erase aExtraQuestions
%>