<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #INCLUDE file="include.asp" -->
<%
	' CHANGES
	' v 1.02 (2/10/2004)
	' * Support an optional dd/mm/yyyy format with new inFormat variable
	' Possible format values are...
	' 0 or blank - mm/dd/yyyy (default)
	' 1 - mm/dd/yy
	' 2 - dd/mm/yyyy
	' 3 - dd/mm/yy
	' 4 - yyyy/mm/dd
	' 5 - yyyy/dd/mm
	' 6 - yy/dd/mm
	' 7 - yy/mm/dd
	' * Fixed minor display bug on Gecko browsers
	' * Clicking a date now displays it with a new class clickedDate

	' v 1.01 
	' * IFrame method is now supported for embedding calendar in page.
	
	' IMPORTANT COPYRIGHT NOTICE
	' The information in this file and all supporting files that this file references (heretofore
	' known as the "software package") are copyright of Philip Dearmore.  You may not reproduce or 
	' distribute the software package in any way, except as authorized by your license as issued
	' by Philip Dearmore or ASPTools.Biz.
	' This software package may have been included in an ASP web application.  If so, you may
	' only use this software package with the application it was included with.  You are not
	' permitted to copy the software package, in whole or in part, to be used in any other way.
	' If you have purchased a license to redistribute this software package as part of one of
	' your applications, this copyright notice MUST remain intact.
	' To see about obtaining a license to use this software package, visit 
	' http://www.asptools.biz/datepicker.asp
	Err.Clear
	' Input variable definitions
	'  ONE OF THE FOLLOWING MUST BE POPULATED WHEN OPENING THE CALENDAR ASP PAGE. 
	'   (The first one populated will be the pass-back method used--the others will be disregarded)
	' 
	'  inField : Field to send selected date back to
	'  inVariable : Javascript variable to send selected date back to
	'  inCallback : Callback function that takes selected date as argument
	'
	'  OPTIONAL VARIABLES
	' 
	'  inDefault : Default selected date
	'  inBlnClose : Close calendar window after date selected?
	'  inBlnWeekends : Allow weekends to be selected?
	'  inBlnSelected : If this is False (default is True) the default date is not highlighted
	Dim inField, inVariable, inCallback, inDefault, inDayStyle, inBlnClose, inBlnWeekends, inBlnSelected
	' v1.01
	Dim inBlnIframe
	' v1.02
	Dim inFormat, inDelimiter
	
	' Internal Variables
	Dim strErrMsg : strErrMsg = Null
	Dim datCurrent : datCurrent = Date()
	Dim datFirstOfMonth, datEndOfMonth
	Dim intWeekdayOfFirst, intWeekdayOfEnd
	Dim strCloseWindow
	' v 1.01
	Dim strPassbackObject
	
	' Array of Day name headers
	Dim aryDays
	' Set up calendar array matrix (42 is the max dates displayed on our calendar)
	Dim aryCalendar(42)
	' Generic index counter
	Dim I
	' Generic string holder
	Dim strTemp
	Dim Item
	
	' SET UP VARIABLES
	
	inField = Trim (Request("inField"))
	inVariable = Trim (Request("inVariable"))
	inCallback = Trim (Request("inCallback"))
	
	' If all 3 passback methods are empty, generate an error.
	If Len (inField & inVariable & inCallback) < 1 Then
		strErrMsg = "You must choose a method to pass the selected date back to the calling page.  " & _
			"This is normally done by including the inField, inVariable, or inCallback variables as " & _
			"querystring arguments.  For example, to call the DatePicker, your URL should look like " & _
			"the following: &quot;http://ASPDatePicker.asp?inField=txtDate&quot;."
	End If
	
	' Populate inDayStyle--if invalid generate error
	If Len (Trim (Request("inDayStyle"))) > 0 Then
		If IsNumeric (Trim (Request("inDayStyle"))) Then
			inDayStyle = CInt (Trim (Request("inDayStyle")))
			If inDayStyle > 3 Or inDayStyle < 1 Then
				strErrMsg = "If passing a value in the inDayStyle parameter, it must be numeric and " & _
					"between 1 and 3, inclusive.  One shows days in M, T... format, two shows days in " & _
					"Mon, Tue... format, 3 shows days in Monday, Tuesday format."
			End If
		Else
			strErrMsg = "If passing a value in the inDayStyle parameter, it must be numeric and " & _
				"between 1 and 3, inclusive.  One shows days in M, T... format, two shows days in " & _
				"Mon, Tue... format, 3 shows days in Monday, Tuesday format."
		End If
	End If

	' Populate inBlnClose--if invalid error
	If Len (Trim (Request("inBlnClose"))) > 0 Then
		If Trim (Request("inBlnClose")) = "True" Or Trim (Request("inBlnClose")) = "False" Then
			inBlnClose = Trim (Request("inBlnClose"))
		Else
			strErrMsg = "If passing a value in the inBlnClose parameter, it must either be True or False. " & _
				"This parameter is optional.  It sets whether you want to close the DatePicker window " & _
				"after a date has been selected or not. "
		End If
	End If

	' Populate inBlnWeekends--if invalid error
	If Len (Trim (Request("inBlnWeekends"))) > 0 Then
		If Trim (Request("inBlnWeekends")) = "True" Or Trim (Request("inBlnWeekends")) = "False" Then
			inBlnWeekends = Trim (Request("inBlnWeekends"))
		Else
			strErrMsg = "If passing a value in the inBlnWeekends parameter, it must either be True or False. " & _
				"This parameter is optional.  It sets whether you want users to be able to select weekends or not."
		End If
	End If

	' v 1.01
	' Populate inBlnIframe--if invalid error
	If Len (Trim (Request("inBlnIframe"))) > 0 Then
		If Trim (Request("inBlnIframe")) = "True" Or Trim (Request("inBlnIframe")) = "False" Then
			inBlnIframe = Trim (Request("inBlnIframe"))
		Else
			strErrMsg = "If passing a value in the inBlnIframe parameter, it must either be True or False. " & _
				"This parameter is optional.  It sets whether you will be accessing the calendar from a pop-up window or an IFRAME tag."
		End If
	End If
	
	' Populate inDefault--if invalid error
	If Len (Trim (Request("inDefault"))) > 0 Then
		If IsDate (Trim (Request("inDefault"))) Then
			inDefault = Trim (Request("inDefault"))
			datCurrent = inDefault
		Else
			strErrMsg = "An invalid date or invalid format was passed in the inDefault parameter."
		End If
	End If
	
	' v1.02
	' Populate inFormat--if invalid generate error 
	If Len (Trim (Request("inFormat"))) > 0 Then
		If IsNumeric (Trim (Request("inFormat"))) Then
			inFormat = CInt (Trim (Request("inFormat")))
		Else
			strErrMsg = "If passing a value in the inFormat parameter, it must be numeric."
		End If
	Else
		inFormat = 0
	End If

	' v1.02
	' Populate inDelimiter
	If Len (Trim (Request("inDelimiter"))) > 0 Then
		inDelimiter = Mid (Trim (Request("inDelimiter")), 1, 1)
	End If
	
	' If there are any errors, end processing.
	If Err.number <> 0 Or Len (strErrMsg) > 0 Then
		strErrMsg = strErrMsg & "<BR>" & Err.Description 
		Response.Write "<DIV class='err'>" & strErrMsg & "</DIV>"
		Response.End 
	End If
	
	' Set up first and end of current month
	datFirstOfMonth = DateSerial (Year (datCurrent), Month (datCurrent), 1)
	datEndOfMonth = DateSerial (Year (datCurrent), Month (datCurrent) + 1, 0)
	
	intWeekdayOfFirst = Weekday (datFirstOfMonth)
	
	' Set up day name headers (default is 1)
	Select Case inDayStyle
		Case 2
			For I = 1 To 7
				strTemp = strTemp & UCase (WeekDayName (I, True)) & " "
			Next
			aryDays = Split (Trim (strTemp), " ")
		Case 3
			For I = 1 To 7
				strTemp = strTemp & WeekDayName (I) & " "
			Next
			aryDays = Split (Trim (strTemp), " ")
		Case Else
			For I = 1 To 7
				strTemp = strTemp & Left (WeekDayName (I), 1) & " "
			Next
			aryDays = Split (Trim (strTemp), " ")
	End Select
	
	' Put all preceeding days (if any) in array before month starts
	For I = 1 To (intWeekdayOfFirst - 1)
		Set aryCalendar(I) = New DP_Date
		aryCalendar(I).D_Date = DateAdd ("d", I - intWeekdayOfFirst, datFirstOfMonth)
		aryCalendar(I).D_Class = "offMonth"
		aryCalendar(I).D_Day = Day (aryCalendar(I).D_Date)
	Next

	' Put all days of month into array
	Do While Month (DateSerial (Year (datCurrent), Month (datCurrent), (I - intWeekdayOfFirst) + 1)) = Month (datCurrent)
		Set aryCalendar(I) = New DP_Date 
		aryCalendar(I).D_Date = DateAdd ("d", I - intWeekdayOfFirst, datFirstOfMonth)
		aryCalendar(I).D_Class = "onMonth"
		aryCalendar(I).D_Day = Day (aryCalendar(I).D_Date)
		I = I + 1
	Loop
	
	' Put any remaining days from the next month into the array to fill it
	Do While I < UBound (aryCalendar) + 1
		Set aryCalendar(I) = New DP_Date 
		aryCalendar(I).D_Date = DateAdd ("d", I - intWeekdayOfFirst, datFirstOfMonth)
		aryCalendar(I).D_Class = "offMonth"
		aryCalendar(I).D_Day = Day (aryCalendar(I).D_Date)
		I = I + 1
	Loop
	
	' Default action of the window is to close after a date has been selected.
	If inBlnClose = "False" Then
		strCloseWindow = ""
	Else
		strCloseWindow = "window.close();"
	End If
	
	' v 1.01
	' The root passback object is window.opener by default, but if in an IFRAME, it is window.parent
	If inBlnIframe = "True" Then
		strPassbackObject = "window.parent."
	Else
		strPassbackObject = "window.opener."
	End If

	Sub OutputCalendar
		Dim strClass, strJS
		Dim strMod ' V1.02
		
		Response.Write "<TR>"
		For Each Item In aryDays
			Response.Write "<TH>" & Item & "</TH>"
		Next
		Response.Write "</TR>"

		Response.Write "<TR>"
		For I = 1 To UBound (aryCalendar)
			' V1.02 - strMod is the DISPLAY date.  All back-end date working is
			' still handled according to the preferences of the server, but now
			' we can show the date in whatever format necessary to the user.
			Select Case inFormat
				Case 1
					strMod = Month (aryCalendar(I).D_Date) & "/" & Day (aryCalendar(I).D_Date) & "/" & Mid (Year (aryCalendar(I).D_Date), 3, 2)
				Case 2
					strMod = Day (aryCalendar(I).D_Date) & "/" & Month (aryCalendar(I).D_Date) & "/" & Year (aryCalendar(I).D_Date)
				Case 3
					strMod = Day (aryCalendar(I).D_Date) & "/" & Month (aryCalendar(I).D_Date) & "/" & Mid (Year (aryCalendar(I).D_Date), 3, 2)
				Case 4
					strMod = Year (aryCalendar(I).D_Date) & "/" & Month (aryCalendar(I).D_Date) & "/" & Day (aryCalendar(I).D_Date)
				Case 5
					strMod = Year (aryCalendar(I).D_Date) & "/" & Day (aryCalendar(I).D_Date) & "/" & Month (aryCalendar(I).D_Date)
				Case 6
					strMod = Mid (Year (aryCalendar(I).D_Date), 3, 2) & "/" & Day (aryCalendar(I).D_Date) & "/" & Month (aryCalendar(I).D_Date)
				Case 7
					strMod = Mid (Year (aryCalendar(I).D_Date), 3, 2) & "/" & Month (aryCalendar(I).D_Date) & "/" & Day (aryCalendar(I).D_Date)
				Case Else
					strMod = aryCalendar(I).D_Date
			End Select
			' v1.02 - Set delimiter if necessary
			If Not IsEmpty (inDelimiter) Then strMod = Replace (strMod, "/", inDelimiter)
			
			If Month (aryCalendar(I).D_Date) <> Month (datCurrent) Then
				strClass = "offMonth"
			Else
				strClass = "onMonth"
			End If
			If DateDiff ("d", aryCalendar(I).D_Date, Now) = 0 Then strClass = strClass & " selectedDate"
			strJS = "onMouseOver='doOver(this);' onMouseOut='doOut(this);'"
			If inField <> "" Then 
				strJS = strJS & " onClick='passbackField(""" & strMod & """);clickMe(this);'" ' V1.02
			ElseIf inVariable <> "" Then
				strJS = strJS & " onClick='passbackVariable(""" & strMod & """);clickMe(this);'" ' V1.02
			ElseIf inCallback <> "" Then
				strJS = strJS & " onClick='passbackCallback(""" & strMod & """);clickMe(this);'" ' V1.02
			End If
 			Response.Write "<TD class='" & strClass & "' " & strJS & ">" & aryCalendar(I).D_Day & "</TD>" & vbCrLf
			If I Mod 7 = 0 Then
				If Month (aryCalendar(I).D_Date) > Month (datCurrent) Then 
					Exit For
				Else
					Response.Write "</TR><TR>"
				End If
			End If
		Next
		Response.Write "</TR>"
	End Sub
%>


<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Select a Date With ASPDatePicker</title>
<link rel="stylesheet" type="text/css" href="ASPDatePicker.css">
<SCRIPT LANGUAGE="Javascript">
<!--
	var inField = "<%= inField %>";
	var inVariable = "<%= inVariable %>";
	var inCallback = "<%= inCallback %>";
	var inDefault = "<%= inDefault %>";
	var inDayStyle = "<%= inDayStyle %>";
	var inBlnClose = "<%= inBlnClose %>";
	var inBlnWeekends = "<%= inBlnWeekends %>";
	var inBlnSelected = "<%= inBlnSelected %>";
	var inBlnIframe = "<%= inBlnIframe %>";
	// v 1.02
	var inFormat = "<%= inFormat %>";
	var inDelimiter = "<%= inDelimiter %>";
	var clickedDate;
	
	var strURL = "&inField=" + inField + "&inVariable=" + inVariable + "&inCallback=" +
		inCallback + "&inDayStyle=" + inDayStyle + "&inBlnClose=" + inBlnClose +
		"&inBlnWeekends=" + inBlnWeekends + "&inBlnSelected=False&inBlnIframe=" + inBlnIframe +
		"&inFormat=" + inFormat + "&inDelimiter=" + inDelimiter;
	
	function previousMonth()
	{
		inDefault = "<%= DateAdd ("m", -1, datCurrent) %>";
		
		location.href = "ASPDatePicker.asp?inDefault=" + inDefault + strURL;
	}

	function previousYear()
	{
		inDefault = "<%= DateAdd ("yyyy", -1, datCurrent) %>";
		
		location.href = "ASPDatePicker.asp?inDefault=" + inDefault + strURL;
	}

	function nextMonth()
	{
		inDefault = "<%= DateAdd ("m", 1, datCurrent) %>";
		
		location.href = "ASPDatePicker.asp?inDefault=" + inDefault + strURL;
	}

	function nextYear()
	{
		inDefault = "<%= DateAdd ("yyyy", 1, datCurrent) %>";
		
		location.href = "ASPDatePicker.asp?inDefault=" + inDefault + strURL;
	}

	function doOver (tdCell)
	{
		tdCell.className += " onMonthOver";
	}

	function doOut (tdCell)
	{
		var aryClasses, i;
		
		aryClasses = tdCell.className.split(" ");
		for (i = 0; i < aryClasses.length; i++)
		{
			if (aryClasses[i] == "onMonthOver") aryClasses[i] = "";
		}
		tdCell.className = aryClasses.join(" ");
	}

	function clickMe (tdCell)
	{
		var aryClasses, i;
		
		if (clickedDate) {
			aryClasses = clickedDate.className.split(" ");
			for (i = 0; i < aryClasses.length; i++)
			{
				if (aryClasses[i] == "clickedDate") aryClasses[i] = "";
			}
			clickedDate.className = aryClasses.join(" ");
		}
		
		tdCell.className += " clickedDate";
		clickedDate = tdCell;
	}
		
	<% If inVariable <> "" Then %>
	function passbackVariable (strDate)
	{
		<%= strPassbackObject %><%= inVariable %> = strDate;
		<%= strCloseWindow %>
	}
	<% ElseIf inField <> "" Then %>
	function passbackField (strDate)
	{
		<%= strPassbackObject %>document.<%= inField %>.value = strDate;
		<%= strCloseWindow %>
	}
	<% ElseIf inCallback <> "" Then %>
	function passbackCallback (strDate)
	{
		<%= strPassbackObject %><%= inCallback %>(strDate);
		<%= strCloseWindow %>
	}
	<%End If %>
//-->
</SCRIPT>

</head>
<body>

<table border="0" cellpadding="0" cellspacing="0" bordercolor="white">
	<tr>
		<td colspan="7" class="titleBar" align="center" valign="bottom">
			<table border="0" cellpadding="0" cellspacing="0" align="center" width="100%">
				<tr>
					<td class="arrows" style="text-align: left;" valign="bottom">
						<img SRC="previousmonth.gif" WIDTH="11" HEIGHT="15" border="0" title="Previous Month" onClick="previousMonth();" style="cursor: hand;"><img SRC="previousyear.gif" WIDTH="9" HEIGHT="15" border="0" title="Previous Year" onClick="previousYear();" style="cursor: hand;">
					</td>
					<td class="titleText">
						<%= MonthName (Month (datCurrent)) %>&nbsp;
						<%= Year (datCurrent) %>
					</td>
					<td class="arrows" style="text-align: right;" valign="bottom">
						<img SRC="nextyear.gif" WIDTH="9" HEIGHT="15" border="0" title="Next Year" onClick="nextYear();" style="cursor: hand;"><img SRC="nextmonth.gif" WIDTH="11" HEIGHT="15" border="0" title="Next Month" onClick="nextMonth();" style="cursor: hand;">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<% OutputCalendar %>
</table>

</body>
</html>
