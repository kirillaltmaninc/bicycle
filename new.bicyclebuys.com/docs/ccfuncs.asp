<%

' turns a cybersource cc auth response into res
' also set's the error message if needed
sub GetCCStatus( nStatus, varReply, res, vCCErrorMessage, vAVSResponseCode, vAVSResponseMessage )
	Select case nStatus

		' successful transmission
		case 0:
			dim decision
			decision = UCase( varReply( "decision" ) )
			
			if decision = "ACCEPT" then
				res = 1
			   vCCErrorMessage = ""

            ' Cart.CC.AuthCode is OK to re-assign
            Cart.CC.AuthCode = varReply( "ccAuthReply_authorizationCode" )
            vAVSResponseCode = varReply( "ccAuthReply_avsCode" )

            select case vAVSResponseCode
               case "A"
                  vAVSResponseMessage = "Address (Street) Matches, ZIP does not."
               case "B"
                  vAVSResponseMessage = "Address information not provided for AVS check."
               case "E"
                  vAVSResponseMessage = "AVS Error."
               case "G"
                  vAVSResponseMessage = "Non-US Card Issuing Bank."
               case "N"
                  vAVSResponseMessage = "No Match on Address (Street) or ZIP."
               case "P"
                  vAVSResponseMessage = "AVS not applicable for this transaction."
               case "R"
                  vAVSResponseMessage = "Retry - System unavailable or timed out."
               case "S"
                  vAVSResponseMessage = "Service not supported by issuer."
               case "U"
                  vAVSResponseMessage = "Address information is unavailable."
               case "W"
                  vAVSResponseMessage = "9 digit ZIP Matches, Address (Street) does not."
               case "X"
                  vAVSResponseMessage = "Address (Street) and 9 digit ZIP match."
               case "Y"
                  vAVSResponseMessage = "Address (Street) and 5 digit ZIP match."
               case "Z"
                  vAVSResponseMessage = "5 digit ZIP matches, Address (Street) does not."
            end select

			elseif decision = "REVIEW" then
				res = 2
   			vCCErrorMessage = "There was a problem completing your order.  We apologize for the inconvenience.  Please contact customer support to review your order."

			else ' REJECT or ERROR

				dim reasonCode
				reasonCode = varReply( "reasonCode" )
				res = 2

				select case reasonCode
               Case "101"
                  vCCErrorMessage = "The request is missing one or more required fields."
               Case "102"
                  vCCErrorMessage = "One or more fields in the request contains invalid data."
               Case "104"
                  vCCErrorMessage = "Duplicate order submission! Please contact customer support to review your order."
               Case "150"
                  vCCErrorMessage = "Error: General system failure. Please contact customer support to review your order."
               Case "151"
                  vCCErrorMessage = "Error: General system failure. Please contact customer support to review your order."
               Case "152"
                  vCCErrorMessage = "Error: General system failure. Please contact customer support to review your order."
               Case "201"
                  vCCErrorMessage = "Please contact customer support to review your order."
               Case "202"
                  vCCErrorMessage = "A problem was found with card expiration..  Please try again with another card or contact customer support."
               Case "203"
                  vCCErrorMessage = "Card declined.  Please try again with another card or contact customer support."
               Case "204"
         			vCCErrorMessage = "There are insufficient funds in your account.  Please use a different credit card."
               Case "205"
                  vCCErrorMessage = "Card declined.  Please try again with another card or contact customer support."
               Case "208"
                  vCCErrorMessage = "Card declined.  Please try again with another card or contact customer support."
               Case "210"
                  vCCErrorMessage = "Card declined.  Please try again with another card or contact customer support."
               Case "221"
                  vCCErrorMessage = "Card declined.  Please try again with another card or contact customer support."
               Case "231"
                  vCCErrorMessage = "Invalid card number.  Please try again with another card or contact customer support."
               Case "233"
                  vCCErrorMessage = "Card declined.  Please try again with another card or contact customer support."
               Case "236"
				      vCCErrorMessage = "Your order could not be completed.  We apologize for the inconvenience.  Please try again at a later time."
               Case "240"
                  vCCErrorMessage = "Invalid card number.  Please try again with another card or contact customer support."
               Case "520"
                  vCCErrorMessage = "Invalid card number.  Please try again with another card or contact customer support."

					case else
						if decision = "REJECT" then
                     vCCErrorMessage = "Your order could not be completed.  Please review the information you entered and try again."
						else ' ERROR
				         vCCErrorMessage = "Your order could not be completed.  We apologize for the inconvenience.  Please try again at a later time."
						end if
				end select
			end if
		case 1
			vCCErrorMessage = "The following error occurred before the request could be sent:"
			vCCErrorMessage = vCCErrorMessage & strErrorInfo
			res = 0
		
		case 2
			vCCErrorMessage = "The following error occurred while sending the request:"
			vCCErrorMessage = vCCErrorMessage & strErrorInfo
			res = 0

		case 3
			vCCErrorMessage = "The following error occurred while waiting for or retrieving the reply:"
			vCCErrorMessage = vCCErrorMessage & strErrorInfo
			HandleCriticalError nStatus, strErrorInfo, oRequest, varReply
			res = 0

		case 4
			vCCErrorMessage = "The following error occurred after receiving and during processing of the reply:"
			vCCErrorMessage = vCCErrorMessage & strErrorInfo
			HandleCriticalError nStatus, strErrorInfo, oRequest, varReply
			res = 0

		case 5
			vCCErrorMessage = "The server returned a CriticalServerError fault:"
			vCCErrorMessage = vCCErrorMessage & GetFaultContent( varReply )
			HandleCriticalError nStatus, strErrorInfo, oRequest, varReply
			res = 0
		
		case 6
			vCCErrorMessage = "The server returned a ServerError fault:"
			vCCErrorMessage = vCCErrorMessage & GetFaultContent( varReply )
			res = 0

		case 7
			vCCErrorMessage = "The server returned a fault:"
			vCCErrorMessage = vCCErrorMessage & GetFaultContent( varReply )
			res = 0
 
		Case 8
			vCCErrorMessage = "An HTTP error occurred:"
			vCCErrorMessage = vCCErrorMessage & strErrorInfo
			vCCErrorMessage = vCCErrorMessage & "Response Body: " & vbCrLf & varReply
			res = 0
	End select
End sub

'------------------------------------------------------------------------------
' If an error occurs after the request has been sent to the server, but the
' client can't determine whether the transaction was successful, then the error
' is considered critical.  If a critical error happens, the transaction may be
' complete in the CyberSource system but not complete in your order system.
' Because the transaction may have been successfully processed by CyberSource,
' you should not resend the transaction, but instead send the error information
' and the order information (customer name, order number, etc.) to the
' appropriate personnel at your company.  They should use the information as
' search criteria within the CyberSource Transaction Search Screens to find the
' transaction and determine if it was successfully processed. If it was, you
' should update your order system with the transaction information. Note that
' this is only a recommendation; it may not apply to your business model.
'------------------------------------------------------------------------------
sub HandleCriticalError( nStatus, strErrorInfo, oRequest, varReply )

	' varReply may be one of the following:
	'
	' A Fault object.
	' A raw reply string.
	' Nothing/Null.

	dim strReply, strReplyType
	if nStatus = 5 then
		strReply = GetFaultContent( varReply )
		strReplyType = "FAULT DETAILS: "
	elseif IsNull( varReply ) then
		strReply = ""
		strReplyType = "No Reply available."
	else
		strReply = varReply
		strReplyType = "RAW REPLY: "
	end if

	dim strMessageToSend
	strMessageToSend _
		= "STATUS: " & CStr( nStatus ) & vbCrLf & _
		  "ERROR INFO: " & strErrorInfo & vbCrLf & _
		  "REQUEST: " & vbCrLf & oRequest.Content( vbCrLf ) & _
		  vbCrLf & strReplyType & vbCrLf & strReply
		  
	' send strMessageToSend to the appropriate personnel at your company
	' using any suitable method, e.g. e-mail, multicast log, etc.
	'response.write vbCrLf & "This is a critical error.  Send the following information to the appropriate personnel at your company:"
	'response.write vbCrLf & strMessageToSend
		
end sub

Function GetFaultContent( oFault )

	dim strRequestID
	if (oFault.RequestID = "") then
		strRequestID = "(unavailable)"
	end if
	
	GetFaultContent = "Fault code: " & oFault.FaultCode & vbCrLf & _
					  "Fault string: " & oFault.FaultString & vbCrLf & _
					  "RequestID: " & strRequestID & vbCrLf & _
					  "Fault document: " & vbCrLf & oFault.FaultDocument
					  
end function

%>
