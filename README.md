<div align="center">

## basNamesAndDates


</div>

### Description

This is a .bas module that contains a few string manipulation functions I find useful in the real word. MakeProper replaces the limited proper case functions of VB with code that will format your string in title case, but not force mid-word capital letters to lower case. So "John Smith III" or "MacDonald" comes out correctly if typed are typed as "john smith III" or "macDonald". Initial letters of words are capitalized, but other letters are left as typed.

DateWord takes a date and converts it to the phrasing used on legal documents, so 1/1/2000 would return "1st day of January, 2000." MailingLabelText accepts a number of inputs and returns a UDT that offers many variations on the name and address for use in creating mailing labels and other reports containing name and address data. The proper business ettiquette is observed in that the presence of an honorific like "Esquire or MD" eliminates the "Mr." or "Dr."

Look for an update of this soon with more functions for string manipulation.

The other functions are used by these three.

LogError is pretty useful, too, come to think of it!

I don't care about winning any prizes, I just wanted to contribute to a site that has helped me out so much. Your feedback is welcome all the same. PLEASE STILL RATE THIS SO I WILL KNOW WHAT YOU THINK!
 
### More Info
 
MakeProper accepts a string.

DateWord accepts a date.

MailingLabelText accepts a number of string arguments, many optional.

MakeProper Calls MakeWordsLowerCase and passes it several words that are commonly left lower case. You may wish to edit my selections.

All of these functions return strings.

Not tested under VB5.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lisa Z\. Morgan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lisa-z-morgan.md)
**Level**          |Advanced
**User Rating**    |4.4 (40 globes from 9 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lisa-z-morgan-basnamesanddates__1-6659/archive/master.zip)





### Source Code

```
Option Explicit
Option Compare Text
'Developed by Lisa Z. Morgan
'Lairhaven Enterprises
'lairhavn@pinn.net
'© 2000 All rights reserved.
'Use under the standard terms of Planet-Source-Code.com
'Is explicitly permitted.
Public Type NameAndAddress
 FullName As String
 MailingName As String
 StreetAddress As String
 CompanyAddress As String
 FullText As String
End Type
Public Function MailingLabelText(LastName As String, FirstName As String, _
       Optional MI As String = "", _
       Optional Title As String = "", _
       Optional Honorific As String = "", _
       Optional CompanyName As String = "", _
       Optional AddrLine1 As String = "", _
       Optional AddrLine2 As String = "", _
       Optional City As String = "", _
       Optional State As String = "", _
       Optional ZipCode As String = "" _
       ) As NameAndAddress
'Generates a full address or as much as is available
 On Error GoTo HandleErr
 Dim strName As String
 Dim strAddress As String
'Build the name
 If Len(MI) = 0 Then
 strName = FirstName & " " & LastName
 Else
 strName = FirstName & " " & MI & " " & LastName
 End If
'Assign the name to the FullName element
 MailingLabelText.FullName = strName
'Add title or honorific if present
 If Len(Honorific) = 0 Then
 If Len(Title) > 0 Then
  strName = Title & " " & strName
 End If
 Else
 strName = strName & ", " & Honorific
 End If
'assign the full name to the MailingName element
 MailingLabelText.MailingName = strName
'Build the Address
 If Len(AddrLine1) > 0 Then
 strAddress = AddrLine1
 End If
 If Len(AddrLine2) > 0 Then
 strAddress = strAddress & vbCrLf & AddrLine2
 End If
 If Len(City) > 0 Then
 strAddress = strAddress & vbCrLf & City
 If Len(State) > 0 Then
  strAddress = strAddress & ", " & State
 End If
  If Len(ZipCode) > 0 Then
  If Right(ZipCode, 1) = "-" Then
   ZipCode = Left(ZipCode, Len(ZipCode) - 1)
  End If
  strAddress = strAddress & " " & ZipCode
  End If
 End If
 'Assign the string to the streetaddress element
 MailingLabelText.StreetAddress = strAddress
 With MailingLabelText
 'Assign the other combinations as appropriate
 If Len(CompanyName) > 0 Then
  .CompanyAddress = CompanyName & vbCrLf & strAddress
 End If
 If (Len(strName) > 0 And Len(CompanyName) > 0) Then
  .FullText = strName & vbCrLf & CompanyName & vbCrLf & strAddress
 ElseIf (Len(strName) > 0 And Len(CompanyName) = 0) Then
  .FullText = strName & vbCrLf & strAddress
 ElseIf (Len(strName) = 0 And Len(CompanyName) > 0) Then
  .FullText = CompanyName & vbCrLf & strAddress
 Else
  .FullText = strAddress
 End If
 End With
ExitHere:
 Exit Function
HandleErr:
 Select Case Err.Number
 Case Else
  LogError "MailingLabelText", Err.Number, Err.Description, Err.Source
  Resume ExitHere
 End Select
End Function
Public Function MakeProper(StringIn As Variant) As String
'Upper-Cases the first letter of each word in in a string
 On Error GoTo HandleErr
 Dim strBuild As String
 Dim intLength As Integer
 Dim intCounter As Integer
 Dim strChar As String
 Dim strPrevChar As String
intLength = Len(StringIn)
'Bail out if there is nothing there
If intLength > 0 Then
 strBuild = UCase(Left(StringIn, 1))
 For intCounter = 1 To intLength
 strPrevChar = Mid$(StringIn, intCounter, 1)
 strChar = Mid$(StringIn, intCounter + 1, 1)
 Select Case strPrevChar
  Case Is = " ", ".", "/"
  strChar = UCase(strChar)
  Case Else
 End Select
 strBuild = strBuild & strChar
 Next intCounter
 MakeProper = strBuild
 strBuild = MakeWordsLowerCase(strBuild, " and ", " or ", " the ", " a ", " to ")
 MakeProper = strBuild
End If
ExitHere:
 Exit Function
HandleErr:
 Select Case Err.Number
 Case Else
  LogError "MakeProper", Err.Number, Err.Description, Err.Source
  Resume ExitHere
 End Select
End Function
Function MakeWordsLowerCase(StringIn As String, _
       ParamArray WordsToCheck()) As String
'Looks for the words in the WordsToCheck Array within
'the StringIn string and makes them lower case
 On Error GoTo HandleErr
 Dim strWordToFind As String
 Dim intWordStarts As Integer
 Dim intWordEnds As Integer
 Dim intStartLooking As Integer
 Dim strResult As String
 Dim intLength As Integer
 Dim intCounter As Integer
 'Initialize the variables
 strResult = StringIn
 intLength = Len(strResult)
 intStartLooking = 1
 For intCounter = LBound(WordsToCheck) To UBound(WordsToCheck)
 strWordToFind = WordsToCheck(intCounter)
 Do
  intWordStarts = InStr(intStartLooking, strResult, strWordToFind)
  If intWordStarts = 0 Then Exit Do
  intWordEnds = intWordStarts + Len(strWordToFind)
  strResult = Left(strResult, intWordStarts - 1) & _
  LCase(strWordToFind) & _
  Mid$(strResult, intWordEnds, (intLength - intWordEnds) + 1)
  intStartLooking = intWordEnds
 Loop While intWordStarts > 0
 intStartLooking = 1
 Next intCounter
 MakeWordsLowerCase = strResult
ExitHere:
 Exit Function
HandleErr:
 Select Case Err.Number
 Case Else
  LogError "MakeWordsLowerCase", Err.Number, Err.Description, Err.Source
  Resume ExitHere
 End Select
End Function
Function OrdinalNumber(NumberIn As Long) As String
'Formats a number as an ordinal number
 On Error GoTo HandleErr
 Dim intLastDigit As Integer
 Dim intLastTwoDigits As Integer
 intLastDigit = NumberIn Mod 10
 intLastTwoDigits = NumberIn Mod 100
 Select Case intLastTwoDigits
 Case 11 To 19
  OrdinalNumber = CStr(NumberIn) & "th"
 Case Else
  Select Case intLastDigit
  Case Is = 1
   OrdinalNumber = CStr(NumberIn) & "st"
  Case Is = 2
   OrdinalNumber = CStr(NumberIn) & "nd"
  Case Is = 3
   OrdinalNumber = CStr(NumberIn) & "rd"
  Case Else
   OrdinalNumber = CStr(NumberIn) & "th"
  End Select
 End Select
ExitHere:
 Exit Function
HandleErr:
 Select Case Err.Number
 Case Else
  LogError "OrdinalNumber", Err.Number, Err.Description, Err.Source
  Resume ExitHere
 End Select
End Function
Function MonthName(DateIn As Date) As String
'Returns the full name of the month of the date passed in
On Error GoTo HandleErr
Dim dv As New DevTools
 Select Case Month(DateIn)
 Case Is = 1
  MonthName = "January"
 Case Is = 2
  MonthName = "February"
 Case Is = 3
  MonthName = "March"
 Case Is = 4
  MonthName = "April"
 Case Is = 5
  MonthName = "May"
 Case Is = 6
  MonthName = "June"
 Case Is = 7
  MonthName = "July"
 Case Is = 8
  MonthName = "August"
 Case Is = 9
  MonthName = "September"
 Case Is = 10
  MonthName = "October"
 Case Is = 11
  MonthName = "November"
 Case Is = 12
  MonthName = "December"
 End Select
ExitHere:
 Exit Function
HandleErr:
 Select Case Err.Number
 Case Else
  LogError "MonthName", Err.Number, Err.Description, Err.Source
  Resume ExitHere
 End Select
End Function
Function DateWord(DateIn As Date) As String
'Accepts: DateIn--the date to be converted
'Returns: DateWord--the date in "5th day of August, 1997" format
'Comments: Calls OrdinalNum for the day value and MonthName for the Month
'*****************************************************************************
 On Error GoTo HandleErr
 Dim strDay As String
 Dim strMonth As String
 Dim strYear As String
 Dim lngIntDayNum As Long
 strMonth = MonthName(DateIn)
 strYear = CStr(Year(DateIn))
 lngIntDayNum = CInt(Day(DateIn))
 strDay = OrdinalNum(lngIntDayNum)
DateWord = strDay & _
 " day of " & strMonth & _
 ", " & strYear
ExitHere:
 Exit Function
HandleErr:
 Select Case Err.Number
 Case Else
  LogError "DateWord", Err.Number, Err.Description, Err.Source
  Resume ExitHere
 End Select
End Function
Public Sub LogError(ProcedureName As String, ErrorNumber As Long, _
   ErrorDescription As String, ErrorSource As String)
 On Error GoTo HandleErr
 Dim lngFileNo As Long
 Dim strTextFile As String
 Dim strPath As String
 Dim strLogText As String
 'Build a text entry for the error log file
 strLogText = vbCrLf & Space(14) & " * BEGIN ERROR RECORD * " & vbCrLf
 strLogText = strLogText & "Error " & ErrorNumber
 strLogText = strLogText & " in Procedure " & ProcedureName & " at " & Now() & vbCrLf
 strLogText = strLogText & ErrorDescription & vbCrLf
 strLogText = strLogText & Space(14) & "* END ERROR RECORD * " & vbCrLf & vbCrLf
 'place the file in the application directory and name it Error Log.txt
 strPath = App.Path
 strTextFile = strPath & "\Error Log.txt"
 'Open the file
 lngFileNo = FreeFile
 Open strTextFile For Append As #lngFileNo
 'Write the error entry
 Write #lngFileNo, strLogText
 'Close the file
 Close #lngFileNo
ExitHere:
 Exit Sub
HandleErr:
 Debug.Print "Error in LogError"
 Resume ExitHere
End Sub
```

