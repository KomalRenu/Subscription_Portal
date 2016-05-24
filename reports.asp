<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	On Error Resume Next
%>
<!-- #include file="CommonDeclarations.asp" -->
<!-- #include file="CustomLib/ReportsCuLib.asp" -->
<%
    Dim aDocInfo(5)
    Dim asSubscriptionGUIDS()

    Dim sSubsXML
    Dim sGetAvailableSubscriptionsXML

    Dim sPortalAddId

    'Check if user is logged in.  If not, send user to login page.
    If Len(LoggedInStatus()) = 0 Then
        Response.Redirect "login.asp"
    End If

    'For selected item in the bands:
    sReportsStyle = ""

    'Start with no errors
    lErr = NO_ERR

    'Get Data from Request object:
    If lErr = NO_ERR Then
        lErr = ParseRequestForReports(oRequest, aDocInfo)
    End If

    'Get the Subscriptions for the Portal Address
    If lErr = NO_ERR Then
        lErr = GetUserSubscriptions_Reports(sSubsXML)
    End If

    If lErr = NO_ERR Then
        lErr = GetSubscriptionsArray_Reports(sSubsXML, asSubscriptionGUIDS)
        If lErr = NO_ERR Then
            If UBound(asSubscriptionGUIDS) <> -1 Then
                lErr = cu_GetAvailableSubscriptions(asSubscriptionGUIDS, sGetAvailableSubscriptionsXML)
            End If
        End If
    End If

    'Get the content of a document, if any is requested:
    If lErr = NO_ERR Then
        lErr = GetDocumentContent(aDocInfo, sSubsXML, sGetAvailableSubscriptionsXML)
    End If
%>
<!-- #include file="CheckError.asp" -->

<HTML>
<HEAD>
	<%Response.Write(putMETATagWithCharSet())%>
	<TITLE><%Response.Write asDescriptors(360)'Descriptor: Reports%> - MicroStrategy Narrowcast Server</TITLE>
</HEAD>
<BODY BGCOLOR="FFFFFF" TOPMARGIN="0" LEFTMARGIN="0" ALINK="ff0000" LINK="0000ff" VLINK="0000ff" MARGINHEIGHT="0" MARGINWIDTH="0">
<!-- #include file="header_multi.asp" -->
<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
  <TR>
    <TD NOWRAP WIDTH="1%" VALIGN="TOP">
      <!-- begin left menu -->
        <!-- #include file="toolbarReports.asp" -->
      <!-- end left menu -->
      <BR />
      <IMG SRC="images/1ptrans.gif" WIDTH="160" HEIGHT="1" BORDER="0" ALT="">
    </TD>
    <TD WIDTH="98%" VALIGN="TOP">
      <!-- begin center panel -->
        <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
          <TR>
            <TD VALIGN="CENTER">
              <!-- begin path -->
                <% RenderReportPath(aDocInfo) %>
              <!-- end path -->
            </TD>
            <TD ALIGN="RIGHT"><IMG SRC="images/desktop_SharedReports.gif" WIDTH="60" HEIGHT="60" BORDER="0" ALT="" /></TD>
          </TR>
        </TABLE>
        <TABLE WIDTH="100%" CELLSPACING="0" CELLPADDING="0" BORDER="0">
          <TR>
            <TD WIDTH="15"><IMG SRC="images/1pTrans.gif" WIDTH="5"/></TD>
            <TD>
              <!-- begin body -->
                <%
                  If lErr <> NO_ERR Then
                    If Len(aDocInfo(DOC_SUBS_ID)) = 0 Then
                      Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(380), "home.asp") 'Descriptor: Back to Home
                    Else
                      Call DisplayError(sErrorHeader, sErrorMessage, asDescriptors(382), "reports.asp") 'Descriptor: Back to Reports
                    End If
                  Else
                    Call RenderSubscriptionContent(aDocInfo)
                  End If
                %>
              <!-- end body -->
            </TD>
            <TD WIDTH="15"><IMG SRC="images/1pTrans.gif" /></TD>
          </TR>
        </TABLE>
      <!-- end center panel -->
    </TD>
    <TD WIDTH="1%">
      <IMG SRC="images/1ptrans.gif" WIDTH="15" HEIGHT="1" BORDER="0" ALT="">
    </TD>
  </TR>
</TABLE>
<BR />
<!-- begin footer -->
	<!-- #include file="footer.asp" -->
<!-- end footer -->
</BODY>
</HTML>
<%
    Erase aDocInfo
    Erase asSubscriptionGUIDS
%>