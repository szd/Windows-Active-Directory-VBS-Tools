Option Explicit
  '
 ' Remove users given in the CSV file from an Active Directory Security Group
 '
 ' Author : Steve ZD
 '
Dim objRootDSE, strDNSDomain, objTrans, strNetBIOSDomain
 Dim strFile, objFSO, objFile, strLine, arrValues
 Dim strName, strNewEmail, strUserDN, objUser,groupName,strGroupDN,objGroup
 
Const ForReading = 1
 ' Constants for the NameTranslate object.
 Const ADS_NAME_INITTYPE_GC = 3
 Const ADS_NAME_TYPE_NT4 = 3
 Const ADS_NAME_TYPE_1779 = 1
 Const ADS_PROPERTY_CLEAR = 1
 Const ADS_PROPERTY_APPEND = 3
 Const USERNAME_ID_IN_CVS = 2 'column ID in CSV file where user login is stored (0,1,2,...)
 
 
' Determine DNS name of domain from RootDSE.
 Set objRootDSE = GetObject("LDAP://RootDSE")
 strDNSDomain = objRootDSE.Get("defaultNamingContext")
 
' Use the NameTranslate object to find the NetBIOS domain name from the
 ' DNS domain name.
 Set objTrans = CreateObject("NameTranslate")
 objTrans.Init ADS_NAME_INITTYPE_GC, ""
 objTrans.Set ADS_NAME_TYPE_1779, strDNSDomain
 strNetBIOSDomain = objTrans.Get(ADS_NAME_TYPE_NT4)
 ' Remove trailing backslash.
 strNetBIOSDomain = Left(strNetBIOSDomain, Len(strNetBIOSDomain) - 1)
 
 If WScript.Arguments.Count<>2 Then
	Wscript.Echo "CSV or AD security Group missing." & VbCrLf & "Syntax : RemoveUsersFromGroup.vbs <CSV_PATH> AD_GROUP_NAME"
Else
' Specify input file.
 strFile = WScript.Arguments(0)
 groupName = WScript.Arguments(1)
 'ouName = WScript.Arguments(2)
 
 
On Error Resume Next
objTrans.Set ADS_NAME_TYPE_NT4, strNetBIOSDomain & "\" & groupName
strGroupDN = objTrans.Get(ADS_NAME_TYPE_1779)	

On Error GoTo 0
 
' Open text file for read access.
 Set objFSO = CreateObject("Scripting.FileSystemObject")
 Set objFile = objFSO.OpenTextFile(strFile, ForReading)
 
' Read file one line at a time.
 ' Assume no header line.
 Do Until objFile.AtEndOfStream
     strLine = Trim(objFile.ReadLine)
     ' Skip blank lines.
     If (strLine <> "") Then
         arrValues = CSVParse(strLine) 
             strName = arrValues(USERNAME_ID_IN_CVS)
                 ' Use Set method to specify NT format of user name.
                 ' Trap error if user not found.
                 On Error Resume Next
                 objTrans.Set ADS_NAME_TYPE_NT4, strNetBIOSDomain & "\" & strName
                 If (Err.Number <> 0) Then
                     On Error GoTo 0                     
                 Else
                     On Error GoTo 0
                     ' Use the Get method to retrieve DN of user object.
                     strUserDN = objTrans.Get(ADS_NAME_TYPE_1779)
					 
					 Set objGroup = GetObject ("LDAP://" & strGroupDN)
					 objGroup.Remove("LDAP://" & strUserDN)
 
					 objGroup.SetInfo
                     
                 End If
             
         
     End If
 Loop
 
' Clean up.
 objFile.Close
End If
 
Function CSVParse(ByVal strLine)
     ' Function to parse comma delimited line and return array
     ' of field values.
     ' Based on program by Michael Harris (a Microsoft MVP).
 
    Dim arrFields, blnIgnore, intFieldCount, intCursor
     Dim intStart, strChar, strValue
 
    Const QUOTE = """"
     Const QUOTE2 = """"""
 
    ' Check for empty string and return empty array.
     If (Len(Trim(strLine)) = 0) then
         CSVParse = Array()
         Exit Function
     End If
 
    ' Initialize.
     blnIgnore = False
     intFieldCount = 0
     intStart = 1
     arrFields = Array()
 
    ' Add "," to delimit the last field.
     strLine = strLine & ";"
 
    ' Walk the string.
     For intCursor = 1 To Len(strLine)
         ' Get a character.
         strChar = Mid(strLine, intCursor, 1)
         Select Case strChar
             Case QUOTE
                 ' Toggle the ignore flag.
                 blnIgnore = Not blnIgnore
             Case ";"
                 If Not blnIgnore Then
                     ' Add element to the array.
                     ReDim Preserve arrFields(intFieldCount)
                     ' Makes sure the "field" has a non-zero length.
                     If (intCursor - intStart > 0) Then
                         ' Extract the field value.
                         strValue = Mid(strLine, intStart, _
                             intCursor - intStart)
                         ' If it's a quoted string, use Mid to
                         ' remove outer quotes and replace inner
                         ' doubled quotes with single.
                         If (Left(strValue, 1) = QUOTE) Then
                             arrFields(intFieldCount) = _
                                 Replace(Mid(strValue, 2, _
                                 Len(strValue) - 2), QUOTE2, QUOTE)
                         Else
                             arrFields(intFieldCount) = strValue
                         End If
                     Else
                         ' An empty field is an empty array element.
                         arrFields(intFieldCount) = Empty
                     End If
                     ' increment for next field.
                     intFieldCount = intFieldCount + 1
                     intStart = intCursor + 1
                 End If
         End Select
     Next
     ' Return the array.
     CSVParse = arrFields
 End Function