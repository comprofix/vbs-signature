' *********************************************************************
' IBS InfoTech
' 
' Outlook Signature Editor
'
' Details: Add to the login script for each user to modify a nominated
'   text/html/rtf file with values obtained from the AD user prifile.
'
' 22/06/2011
'
' Ver. 2.0
'
' Version History
' 2.0 - Included support for multiple files and pictures
' 1.0 - Original script
' *********************************************************************

'On Error Resume Next

' Constants
	Const ForReading = 1 ,ForWriting = 2, ForAppending = 8
    Const HKEY_CURRENT_USER = &H80000001
	Const cAppDataPath = "\AppData\Roaming\Microsoft\Signatures\"
	
' Create objects
	Set Fso = wScript.CreateObject("Scripting.FileSystemObject")
	Set Fso2 = wScript.CreateObject("Scripting.FileSystemObject")
	Set WshShell = WScript.CreateObject("WScript.Shell")
	Set WshProcessEnvironment = WshShell.Environment("Process")
	Set objSysInfo = CreateObject("ADSystemInfo")

' Variables
	Dim OfficePhone, OfficeFax, SourcePath, oSourceHTML, oSourceText, oSrcHTM, oSourceRText, oSourceRHTML, sSrcTXT, sUserProfile, sUser, objuser, oFile
	Dim sTitle
	Dim sDescription
	Dim sUID
	Dim sDisplayName
	Dim sFirstName
	Dim sLastName
	Dim sInitials
	Dim sEMail
	Dim sPhoneNumber
	Dim sFaxNumber
	Dim sMobileNumber
	Dim sHomePhone
	Dim sDepartment
	Dim sOfficeLocation
	Dim sWebAddress
	Dim sLastLogOff
	Dim sAddStreet
	Dim sAddCity
	Dim sAddPOBox
	Dim sAddPostCode
	Dim sAddState
	Dim sAddCountry

	' ############ Change Settings Here ############
	LogonServer = "\\comprofix.local"
	SourcePath= "\\fs01.comprofix.local\Signatures$\"
	SourceImagePath = "\\fs01.comprofix.local\Signatures$\Comprofix_files"
	
	
	sSrcHTM="Comprofix.htm"
	sSrcTXT="Comprofix.txt"
	

	' ############ Change Settings Here ############

	' start running the script
	Call Start

'#####################################################################################################

	Private Sub Start

		' Set the user object
		sUserProfile = WshProcessEnvironment("USERPROFILE")
		sUser = objSysInfo.UserName
		Set objUser = GetObject("LDAP://" & sUser)
	
		' If the object returned null exit the script
		If isNull(objUser) Then WScript.Quit
		If IsNull(objUser.Displayname) Then WScript.Quit
		If Len(objUser.Displayname) = 0 Then WScript.Quit
		
		' Assign the user profile values to the variables
		sTitle = objuser.title
		sDescription = objuser.description
		sUID = objuser.cn
		sDisplayName = objuser.displayName
		sFirstName = objuser.givenName
		sLastName = objuser.sn
		sInitials = objuser.initials
		sEMail = objuser.mail
		sPhoneNumber = objuser.telephoneNumber
		sFaxNumber = objuser.facsimileTelephoneNumber
		sMobileNumber = objuser.mobile
		sHomePhone= objUser.homePhone
		sDepartment = objuser.department
		sOfficeLocation = objuser.physicalDeliveryOfficeName
		sWebAddress = objuser.wWWHomePage
		sAddStreet = objuser.streetAddress
		sAddCity = objuser.l
		sAddPOBox = objuser.postOfficeBox
		sAddPostcode = objuser.postalCode
		sAddState = objuser.st
		sAddCountry = objuser.c
		
		' open the 2 files and set the users' values
		Call OpenAndReplaceFiles
		
		' Get the paths for the files to be saved to
		Call BuildSignaturePath

		' save the signature files to the path
		Call SaveSignature (oSourceHTML, sUserProfile & cAppDataPath & sSrcHTM)
		Call SaveSignature (oSourceText, sUserProfile & cAppDataPath & sSrcTXT)
		
		' copy any image files to the path
		Call CopyAdditionalImages(SourceImagePath, sUserProfile & cAppDataPath)
		
	
		' Set the new signature files as the default signature for the users' default Outlook profile
		SetDefaultSignature Left(sSrcHTM,Instr(sSrcHTM,".")-1),"",True
		
		' Exit the script
		WScript.Quit

	End Sub

'#####################################################################################################

	Sub OpenAndReplaceFiles
	
		' open the specified HTML file for read-only
		Set oFile = Fso.OpenTextFile(SourcePath & sSrcHTM,ForReading,True)
		If ofile.AtEndOfLine = True Then WScript.Echo "oFile = EOF"
		oSourceHTML= oFile.ReadAll
		oFile.Close
		Set oFile = Nothing

				
		' open the specified text file for read-only
		Set oFile = Fso.OpenTextFile(SourcePath & sSrcTXT,ForReading,True)
		oSourceText= oFile.ReadAll
		oFile.Close
		Set oFile = Nothing
  
		 
		' Replace variables in the Default HTML file
		oSourceHTML=replace(oSourceHTML,"#USERID#",sUID)
		oSourceHTML=replace(oSourceHTML,"#DISPLAYNAME#",sDisplayName)
		oSourceHTML=replace(oSourceHTML,"#TITLE#",sTitle)
		oSourceHTML=replace(oSourceHTML,"#USERDESCRIPTION#",sDescription)
		oSourceHTML=replace(oSourceHTML,"#EMAIL#",lcase(sEMail))
		oSourceHTML=replace(oSourceHTML,"#FIRSTNAME#",sFirstName)
		oSourceHTML=replace(oSourceHTML,"#LASTNAME#",sLastName)
		oSourceHTML=replace(oSourceHTML,"#INITIALS#",sInitials)
		oSourceHTML=replace(oSourceHTML,"#OFFICEPHONE#",sPhoneNumber)
		oSourceHTML=replace(oSourceHTML,"#FAXNUMBER#",sFaxNumber)
		oSourceHTML=replace(oSourceHTML,"#MOBILENUMBER#",sMobileNumber)
		oSourceHTML=replace(oSourceHTML,"#HOMEPHONE#",sHomePhone)
		oSourceHTML=replace(oSourceHTML,"#DEPARTMENT#",sDepartment)
		oSourceHTML=replace(oSourceHTML,"#OFFICELOCATION#",sOfficeLocation)
		oSourceHTML=replace(oSourceHTML,"#WEBADDRESS#",sWebAddress)
		oSourceHTML=replace(oSourceHTML,"#LASTLOGOFF#",sLastLogOff)
		oSourceHTML=replace(oSourceHTML,"#ADDSTREET#",sAddStreet)
		oSourceHTML=replace(oSourceHTML,"#ADDCITY#",sAddCity)
		oSourceHTML=replace(oSourceHTML,"#ADDPOBOX#",sAddPOBox)
		oSourceHTML=replace(oSourceHTML,"#ADDPOSTCODE#",sAddPostcode)
		oSourceHTML=replace(oSourceHTML,"#ADDSTATE#",sAddState)
		oSourceHTML=replace(oSourceHTML,"#ADDCOUNTRY#",sAddCountry)

		' Replace variables in the default text file
		oSourceText=replace(oSourceText,"#USERID#",sUID)
		oSourceText=replace(oSourceText,"#DISPLAYNAME#",sDisplayName)
		oSourceText=replace(oSourceText,"#TITLE#",sTitle)
		oSourceText=replace(oSourceText,"#USERDESCRIPTION#",sDescription)
		oSourceText=replace(oSourceText,"#EMAIL#",lcase(sEMail))
		oSourceText=replace(oSourceText,"#FIRSTNAME#",sFirstName)
		oSourceText=replace(oSourceText,"#LASTNAME#",sLastName)
		oSourceText=replace(oSourceText,"#INITIALS#",sInitials)
		oSourceText=replace(oSourceText,"#OFFICEPHONE#",sPhoneNumber)
		oSourceText=replace(oSourceText,"#FAXNUMBER#",sFaxNumber)
		oSourceText=replace(oSourceText,"#MOBILENUMBER#",sMobileNumber)
		oSourceText=replace(oSourceText,"#HOMEPHONE#",sHomePhone)
		oSourceText=replace(oSourceText,"#DEPARTMENT#",sDepartment)
		oSourceText=replace(oSourceText,"#OFFICELOCATION#",sOfficeLocation)
		oSourceText=replace(oSourceText,"#WEBADDRESS#",sWebAddress)
		oSourceText=replace(oSourceText,"#LASTLOGOFF#",sLastLogOff)
		oSourceText=replace(oSourceText,"#ADDSTREET#",sAddStreet)
		oSourceText=replace(oSourceText,"#ADDCITY#",sAddCity)
		oSourceText=replace(oSourceText,"#ADDPOBOX#",sAddPOBox)
		oSourceText=replace(oSourceText,"#ADDPOSTCODE#",sAddPostcode)
		oSourceText=replace(oSourceText,"#ADDSTATE#",sAddState)
		oSourceText=replace(oSourceText,"#ADDCOUNTRY#",sAddCountry)

		
		' add entries below for defaults based on content
		'If len(trim(sMobileNumber)) = 0 then oSourceHTML=replace(oSourceHTML,"#MOBILE#","")
		'If len(trim(sMobileNumber))  > 0 then oSourceHTML=replace(oSourceHTML,"#MOBILE#","<br>M " & sMobileNumber)
	
		' default		
		If len(trim(sMobileNumber)) = 0 then oSourceHTML=replace(oSourceHTML,"#MOBILE#","")
		If len(trim(sMobileNumber))  > 0 then oSourceHTML=replace(oSourceHTML,"#MOBILE#","<FONT face=Arial><SPAN style=""COLOR: #6a8a4c""> | </SPAN></FONT><FONT face=Arial><B><SPAN style=""FONT-SIZE: 8pt"" lang=EN-US>Mobile</B></SPAN></FONT><FONT face=Arial><SPAN style=""FONT-SIZE: 8pt"" lang=EN-US> " & sMobileNumber & "</SPAN></FONT>")

		' default
		If len(trim(sMobileNumber)) = 0 then oSourceText=replace(oSourceText,"#MOBILE#","")
		If len(trim(sMobileNumber))  > 0 then oSourceText=replace(oSourceText,"#MOBILE#"," | Mobile " & sMobileNumber)

				
'		If len(trim(Phone)) = 0 then SourceHtml=replace(oSourceHTML,"#PHONE#",cBrisbanePhone)
'		If len(trim(Phone))  > 0 then oSourceHTML=replace(oSourceHTML,"#PHONE#",Phone)
		
'		If len(trim(Phone)) = 0 then oSourceText=replace(oSourceText,"#PHONE#",cBrisbanePhone)
'		If len(trim(Phone))  > 0 then oSourceText=replace(oSourceText,"#PHONE#",Phone)
		
'		If len(trim(sFaxNumber)) = 0 then oSourceHTML=replace(oSourceHTML,"#FAX#",OfficeFax)
'		If len(trim(sFaxNumber))  > 0 then oSourceHTML=replace(oSourceHTML,"#FAX#",sFaxNumber)
		
'		If len(trim(sFaxNumber)) = 0 then oSourceText=replace(oSourceText,"#FAX#",OfficeFax)
'		If len(trim(sFaxNumber))  > 0 then oSourceText=replace(oSourceText,"#FAX#",sFaxNumber)
		
'		oSourceHTML=replace(oSourceHTML,"#OFFICEPHONE#",OfficePhone)
'		oSourceText=replace(oSourceText,"#OFFICEPHONE#",OfficePhone)

	End Sub


'#####################################################################################################

	Sub BuildSignaturePath
		
		' Create Base Folders if they do not exist
		
		If not fso.folderexists(sUserProfile & "\AppData") then
			fso.createFolder(sUserProfile & "\AppData")
		End If
		
		If not fso.folderexists(sUserProfile & "\AppData\Microsoft") then
			fso.createFolder(sUserProfile & "\AppData\Microsoft")
		End If
		
	 	If not fso.folderexists(sUserProfile & "\AppData\Microsoft\Signatures") then
			fso.createFolder(sUserProfile & "\AppData\Microsoft\Signatures")
		End if
	
	End Sub

'#####################################################################################################

	Function SaveSignature(SignatureText,SignaturePath)
		
		' write the file to the given path
	
	  	Set ca = fso.CreateTextFile(SignaturePath, ForWriting, True)
	  	
	  	ca.write(SignatureText)
	 	ca.close
		set ca=nothing
	
	End Function

'#####################################################################################################

	Function CopyAdditionalImages(SourcePath, SignaturePath)
	
		' copy the images from the source path to the destination
		fso.copyFolder SourcePath, SignaturePath
		On Error resume Next

		fso.copyFolder SourcePath, SignaturePath

		'fso.copyFile SourcePath & "*.png", SignaturePath
		'fso.copyFile SourcePath & "*.gif", SignaturePath
	
	End Function

'#####################################################################################################

	Sub SetDefaultSignature(strSigName, strProfile, blnSetReply)
	
		' modify the registry to set Outlook new and reply signatures 
		
		On error resume next
	    
	    Dim arrProfileKeys, objreg, MyArray
	    
	    strComputer = "."
	  
		Set objreg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
		strKeyPath = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\"
		strKeyPath2 = "SOFTWARE\Microsoft\Office\16.0\Common\MailSettings"
		
		' get default profile name if none specified
		If strProfile = "" Then
		    objreg.GetStringValue HKEY_CURRENT_USER, strKeyPath, "DefaultProfile", strProfile
		End If
		
		' build array from signature name
		myArray = StringToByteArray(strSigName, True)
		strKeyPath = strKeyPath & strProfile & "\9375CFF0413111d3B88A00104B2A6676"
		objreg.EnumKey HKEY_CURRENT_USER, strKeyPath, arrProfileKeys
		
		For Each subkey In arrProfileKeys
		    strsubkeypath = strKeyPath & "\" & subkey
		    'On Error Resume Next
		    If blnsetreply = False Then
		    	' set default
		    	objreg.SetBinaryValue HKEY_CURRENT_USER, strsubkeypath, "New Signature", MyArray
		    Else
		    	' set reply
		    	objreg.SetBinaryValue HKEY_CURRENT_USER, strsubkeypath, "Reply-Forward Signature", MyArray
		    End If
		Next

		Set objreg = nothing
		set arrProfileKeys=Nothing
		set MyArray=Nothing
	
	End Sub

'#####################################################################################################

	Public Function StringToByteArray (Data, NeedNullTerminator)
	    
	    Dim strAll
	    
	    strAll = StringToHex4(Data)
	    
	    If NeedNullTerminator Then
	        strAll = strAll & "0000"
	    End If
	    
	    intLen = Len(strAll) \ 2
	    ReDim arr(intLen - 1)
	    
	    For i = 1 To Len(strAll) \ 2
	        arr(i - 1) = CByte("&H" & Mid(strAll, (2 * i) - 1, 2))
	    Next
	    
	    StringToByteArray = arr
	
	End Function

'#####################################################################################################

	Public Function StringToHex4(Data)
	    
	    ' Input: normal text
	    ' Output: four-character string for each character,
	    '         e.g. "3204" for lower-case Russian B,
	    '        "6500" for ASCII e
	    ' Output: correct characters
	    ' needs to reverse order of bytes from 0432
	    
	    Dim strAll
	    
	    For i = 1 To Len(Data)
	        ' get the four-character hex for each character
	        strChar = Mid(Data, i, 1)
	        strTemp = Right("00" & Hex(AscW(strChar)), 4)
	        strAll = strAll & Right(strTemp, 2) & Left(strTemp, 2)
	    Next
	    
	    StringToHex4 = strAll
	
	End Function