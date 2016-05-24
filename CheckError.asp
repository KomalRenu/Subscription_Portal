<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<%
    If lErr = ERR_SESSION_TIMEOUT Then
        Call Logout()
        Response.Redirect "login.asp?status=timeout"
    End If

    Select Case LCase(aPageInfo(S_NAME_PAGE))
        Case "home.asp", "options.asp", "printreport.asp", "reports.asp", "services.asp", "subscriptions.asp"
            'Do nothing
        Case "login.asp"
	        sErrorHeader = asDescriptors(384) 'Descriptor: Error during login
	        Select Case lErr
	        	Case ERR_XML_LOAD_FAILED
	        		sErrorMessage = asDescriptors(428) 'Descriptor: There was an error while loading XML.  Please contact the system administrator.
	        	Case ERR_LOGIN_ERROR
	        		sErrorMessage = asDescriptors(385) 'Descriptor: Either the User name or Password was incorrect
	        	Case IS_ERR_LOGIN_ERROR
	        		sErrorMessage = asDescriptors(385) 'Descriptor: Either the User name or Password was incorrect
	        	Case ERR_USER_INACTIVE
	        	    sErrorHeader = asDescriptors(461) 'Descriptor: Your account has been deactivated.
	        	    sErrorMessage = asDescriptors(486) 'Descriptor: This account has been deactivated. Please contact the administrator to reactivate this account.
	        	Case ERR_LOGIN_BLANKS
	        		sErrorMessage = asDescriptors(386) 'Descriptor: Either the User name or Password was blank
	        	Case ERR_PRIMARY_KEY_VIOLATION
	        	    sErrorMessage = asDescriptors(423) 'Descriptor: This User name is already taken. Please enter a different User name.
	        	Case API_ERR_NT_NOT_LINKED
	        	    sErrorMessage = asDescriptors(897) 'Descriptor: Cannot log in by using the NT  user account. Please ask your administrator to link this NT user account to a MicroStrategy Intelligence Server user or set the Web server properly.
				Case IS_SERVER_NOT_FOUND
				    sErrorMessage = "MicroStrategy Intelligence Server was not found."
	        	Case Else
	        End Select
        Case "default.asp"
	        sErrorHeader = asDescriptors(384) 'Descriptor: Error during login
	        Select Case lErr
	        	Case ERR_XML_LOAD_FAILED
	        		sErrorMessage = asDescriptors(428) 'Descriptor: There was an error while loading XML.  Please contact the system administrator.
	        	Case Else
	        End Select
        Case "subscribe.asp"
	        Select Case lErr
	        	Case URL_MISSING_PARAMETER
	        		sErrorHeader = asDescriptors(325) 'Descriptor: Wrong parameters in the URL
	        		sErrorMessage = asDescriptors(326) & " serviceID" 'Descriptor: The following parameters are required in the URL:
	        	Case ERR_EMAIL_ADDR_INVALID
	        	    sErrorHeader = asDescriptors(390) 'Descriptor: Error during subscription operation
                    Select Case CStr(Application("Device_Validation"))
                        Case S_DEVICE_VALIDATION_EMAIL
                            sErrorMessage = asDescriptors(419) 'Descriptor: Please enter an address in the form of: user@server.com
                        Case S_DEVICE_VALIDATION_NUMBER
                            sErrorMessage = asDescriptors(614) 'Descriptor: Please enter a value for the address in the following form: any numbers and the following characters - ( )
                        Case S_DEVICE_VALIDATION_NONE
                            sErrorMessage = asDescriptors(635) 'Descriptor: Please enter an address in the following form: any text or numeric characters
                        Case Else
                    End Select
	        	Case Else
	        	    sErrorHeader = asDescriptors(390) 'Descriptor: Error during subscription operation
	        	    sErrorMessage = asDescriptors(388) 'Descriptor: Please try again or contact your system administrator.
	        End Select
        Case "personalize.asp"
	        Select Case lErr
	        	Case ERR_EMAIL_ADDR_INVALID
	        		sErrorMessage = asDescriptors(419) 'Descriptor: Please enter an address in the form of: user@server.com
	        	Case URL_MISSING_PARAMETER
	        		sErrorHeader = asDescriptors(325) 'Descriptor: Wrong parameters in the URL
	        		sErrorMessage = asDescriptors(326) & " eSGUID" 'Descriptor: The following parameters are required in the URL:
	        	Case ERR_USERDEFAULT_NOTEXIST
	        		sErrorHeader = asDescriptors(689) 'Descriptor: Error creating subscription
	        		sErrorMessage = asDescriptors(690) & " " & asDescriptors(159)	'Descriptor: Please make sure you have default profile for the question object 'Descriptor: Contact the system administrator.
	        	Case Else
	        End Select
        Case "subsconfirm.asp", "preprompt.asp", "postprompt.asp"
	        Select Case lErr
	        	Case ERR_XML_LOAD_FAILED
	        		sErrorHeader = asDescriptors(427) 'Descriptor: Error retrieving data
	        		sErrorMessage = asDescriptors(428) 'Descriptor: There was an error while loading XML.  Please contact the system administrator.
	        	Case URL_MISSING_PARAMETER
	        		sErrorHeader = asDescriptors(325) 'Descriptor: Wrong parameters in the URL
	        		sErrorMessage = asDescriptors(326) & " subGUID" 'Descriptor: The following parameters are required in the URL:
	        	Case Else
	        End Select
        Case "prompt.asp"
	        Select Case lErr
	        	Case URL_MISSING_PARAMETER
	        		sErrorHeader = asDescriptors(325) 'Descriptor: Wrong parameters in the URL
	        		sErrorMessage = asDescriptors(326) & " subGUID" 'Descriptor: The following parameters are required in the URL:
	        	Case ERR_GET_HYDRA_PROMPT
	        		sErrorHeader = asDescriptors(691)  'Descriptor: Error retrieving question object information from Narrowcast Server.
	        		sErrorMessage = asDescriptors(159)	'Descriptor: Please contact your system Administrator.
	        	Case ERR_API_NO_PROJECT_ACCESS
	        		sErrorHeader = asDescriptors(907) 'Descriptor:You do not have access to one or more of the projects that contain the reports in this service.
	        		sErrorMessage = asDescriptors(159)	'Descriptor: Please contact your system Administrator.
	        	Case ERR_API_NO_WRITE_ACCESS
	        		sErrorHeader = asDescriptors(908) 'Descriptor:You do not have access to one or more of the reports in this service.
	        		sErrorMessage = asDescriptors(159)	'Descriptor: Please contact your system Administrator.
	        	Case Else
	        End Select
        Case "modify_subscription.asp"
	        Select Case lErr
	        	Case ERR_XML_LOAD_FAILED
	        		sErrorHeader = asDescriptors(427) 'Descriptor: Error retrieving data
	        		sErrorMessage = asDescriptors(428) 'Descriptor: There was an error while loading XML.  Please contact the system administrator.
	        	Case URL_MISSING_PARAMETER
	        		sErrorHeader = asDescriptors(325) 'Descriptor: Wrong parameters in the URL
	        		sErrorMessage = asDescriptors(326) & " subGUID" 'Descriptor: The following parameters are required in the URL:
	        	Case ERR_SUBS_OPERATION
	        		sErrorHeader = asDescriptors(390) 'Descriptor: Error during subscription operation
	        		sErrorMessage = asDescriptors(388) 'Descriptor: Please try again or contact your system administrator
	        	Case Else
	        End Select
        Case "addresses.asp"
	        sErrorHeader = asDescriptors(387) 'Descriptor: Error during address operation
	        If lValidationError <> NO_ERR Then
	            sErrorMessage = ""
	            If (lValidationError And ERR_ADDRESS_BLANKS) Then
	                sErrorMessage = sErrorMessage & "<LI>" & asDescriptors(389) & "</LI>" 'Descriptor: One or more address fields were blank.  Please try again.
	            End If
                If (lValidationError And ERR_ADDR_NAME_INVALID) Then
                    sErrorMessage = sErrorMessage & "<LI>" & asDescriptors(417) & " " 'Descriptor: Please enter an address name without the following characters:
                    Dim j
                    Dim iUBoundResChars_CheckError
                    Dim sChars
                    iUBoundResChars_CheckError = Ubound(asReservedChars)
                    sChars = ""
	            	For j = 0 to iUBoundResChars_CheckError
	            		sChars = sChars & asReservedChars(j) & " "
	            	Next
	            	sChars = Left(sChars, Len(sChars) - 1)
	            	sErrorMessage = sErrorMessage & sChars & "</LI>"
                End If
                If (lValidationError And ERR_EMAIL_ADDR_INVALID) Then
                    sErrorMessage = sErrorMessage & "<LI>" & asDescriptors(419) & "</LI>" 'Descriptor: Please enter an address in the form of: user@server.com
                End If
                If (lValidationError And ERR_NUMBER_ADDR_INVALID) Then
                    sErrorMessage = sErrorMessage & "<LI>" & asDescriptors(529) & "</LI>" 'Descriptor: Please enter a numeric value for the address in the following form: #########
                End If
	        Else
	            Select Case lErr
	            	Case ERR_XML_LOAD_FAILED
	            		sErrorHeader = asDescriptors(427) 'Descriptor: Error retrieving data
	            		sErrorMessage = asDescriptors(428) 'Descriptor: There was an error while loading XML.  Please contact the system administrator.
	            	Case URL_MISSING_PARAMETER
	            		sErrorHeader = asDescriptors(325) 'Descriptor: Wrong parameters in the URL
	            		sErrorMessage = asDescriptors(326) & " action" 'Descriptor: The following parameters are required in the URL:
	            	Case ERR_ADDR_OPERATION
	            		sErrorMessage = asDescriptors(388) 'Descriptor: Please try again or contact your system administrator
	            	Case Else
	            End Select
	        End If
        Case "address_wiz.asp"
            Select Case lErr
                Case URL_MISSING_PARAMETER
	                sErrorHeader = asDescriptors(325) 'Descriptor: Wrong parameters in the URL
	            	sErrorMessage = asDescriptors(326) & " deviceTypeID" 'Descriptor: The following parameters are required in the URL:
                Case Else
            End Select
        Case "logout.asp"
	        Select Case lErr
	        	Case URL_MISSING_PARAMETER
	        		sErrorHeader = asDescriptors(325) 'Descriptor: Wrong parameters in the URL
	        		sErrorMessage = asDescriptors(326) & " action" 'Descriptor: The following parameters are required in the URL:
	        	Case Else
	        End Select
        Case "change_password.asp"
	        sErrorHeader = asDescriptors(425) 'Descriptor: Error updating user profile
	        sErrorMessage = asDescriptors(443) & " " & asDescriptors(440) 'Descriptor: One or more errors were encountered when processing your change password request. Please follow the red instructions below and try again.
	        Select Case lErr
	        	Case ERR_XML_LOAD_FAILED
	        		sErrorMessage = asDescriptors(428) 'Descriptor: There was an error while loading XML.  Please contact the system administrator.
	        	Case ERR_LOGIN_ERROR
	        		sErrorMessage = asDescriptors(426) 'Descriptor: There was an error while updating your user profile.
	        	Case Else
	        	    If lValidationError = NO_ERR Then
	        	        sErrorMessage = asDescriptors(426) 'Descriptor: There was an error while updating your user profile.
	        	    End If
	        End Select
        Case "newuser.asp"
	        sErrorHeader = asDescriptors(391) 'Descriptor: Error while creating a new account
	        sErrorMessage = asDescriptors(439) & " " & asDescriptors(440) 'Descriptor: One or more errors were encountered when processing your request for a new account. Please follow the red instructions below and try again.
	        Select Case lErr
	        	Case ERR_XML_LOAD_FAILED
	        		sErrorMessage = asDescriptors(428) 'Descriptor: There was an error while loading XML.  Please contact the system administrator.
	        	Case ERR_LOGIN_ERROR
	        		sErrorMessage = asDescriptors(393) 'Descriptor: Error creating new user account.
	        	Case ERR_PRIMARY_KEY_VIOLATION
	        	    sErrorMessage = asDescriptors(423) 'Descriptor: This User name is already taken. Please enter a different User name.
	        	Case ERR_EMPTY_GUID
	        		sErrorMessage = asDescriptors(919) 'Descriptor: Error creating user: ID generation failed.  Please notify the system administrator.
	        	Case Else
	        	    If lValidationError = NO_ERR Then
	        	        sErrorMessage = asDescriptors(393) 'Descriptor: Error creating new user account.
	        	    End If
	        End Select
        Case "authentications.asp"
	        sErrorHeader = asDescriptors(425) 'Descriptor: Error updating user profile
	        sErrorMessage = asDescriptors(468) & " " & asDescriptors(440) 'Descriptor: One or more errors were encountered when processing your change Information Source credentials request. Please follow the red instructions below and try again.
	        Select Case lErr
	        	Case ERR_XML_LOAD_FAILED
	        		sErrorMessage = asDescriptors(428) 'Descriptor: There was an error while loading XML.  Please contact the system administrator.
	        	Case ERR_LOGIN_ERROR
	        		sErrorMessage = asDescriptors(426) 'Descriptor: There was an error while updating your user profile.
	        	Case ERR_LOGIN_BLANKS
	        	    sErrorMessage = asDescriptors(385) 'Descriptor: Either the User name or Password was blank.  Please enter them again.
	        	Case Else
	        	    If lValidationError = NO_ERR Then
	                    sErrorMessage = asDescriptors(426) 'Descriptor: There was an error while updating your user profile.
	                End If
	        End Select
        Case "password_hint.asp"
	        sErrorHeader = asDescriptors(413) 'Descriptor: Error retrieving password hint
	        Select Case lErr
	        	Case ERR_XML_LOAD_FAILED
	        		sErrorMessage = asDescriptors(428) 'Descriptor: There was an error while loading XML.  Please contact the system administrator.
	        	Case ERR_LOGIN_BLANKS
	        		sErrorMessage = asDescriptors(487) 'Descriptor: The User name was blank. Please enter it again.
	        	Case ERR_LOGIN_ERROR
	        		sErrorMessage = asDescriptors(414) 'Descriptor: There was an error while retrieving your password hint
	        	Case ERR_USER_NOT_EXIST
	        		sErrorMessage = asDescriptors(834)	'Descriptor: This user does not exist.
	        	Case NO_ERR
	        	Case Else
	        		sErrorMessage = asDescriptors(414) 'Descriptor: There was an error while retrieving your password hint
	        End Select
        Case "deactivate.asp"
	        sErrorHeader = asDescriptors(425) 'Descriptor: Error updating user profile
	        Select Case lErr
	        	Case ERR_XML_LOAD_FAILED
	        		sErrorMessage = asDescriptors(428) 'Descriptor: There was an error while loading XML.  Please contact the system administrator.
	        	Case ERR_LOGIN_ERROR
	        		sErrorMessage = asDescriptors(426) 'Descriptor: There was an error while updating your user profile.
	        	Case ERR_LOGIN_BLANKS
	        	    sErrorMessage = asDescriptors(488) 'Descriptor: The Password was blank. Please enter it again.
	        	Case "C0045215"
	        	    sErrorMessage = asDescriptors(889) 'Descriptor: Invalid password - User not deactivated.
	        	Case NO_ERR
	        	Case Else
	                sErrorMessage = asDescriptors(426) 'Descriptor: There was an error while updating your user profile.
	        End Select
        Case "deleteprofile.asp"
            Select Case lErr
                Case URL_MISSING_PARAMETER
	                sErrorHeader = asDescriptors(325) 'Descriptor: Wrong parameters in the URL
	            	sErrorMessage = asDescriptors(326) & " deviceTypeID" 'Descriptor: The following parameters are required in the URL:
	            Case HYDRA_APIERROR_DELETE_PROFILE
					sErrorHeader = asDescriptors(625) 'Descriptor: Error deleting user profile
					sErrorMessage = asDescriptors(626)  'Descriptor: The profile can't be deleted because it's shared by more than 1 subscription
                Case Else
            End Select

        Case Else

    End Select
%>