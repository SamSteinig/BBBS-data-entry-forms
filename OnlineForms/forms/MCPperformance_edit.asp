<!--#include file="../includes/NAD_BE.asp" -->

<% 

' Get Initial Values for fields to check for propagation

If Request("status") = "editOld" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmMCPPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & " AND Month=" & Int(Request("m"))
	Set GetInitialValues = Con.Execute(query)

	OpenMatchesCommunityBasedInitial = GetInitialValues("OpenMatchesCommunityBased")
	OpenMatchesSchoolBasedInitial = GetInitialValues("OpenMatchesSchoolBased")
	OpenMatchesOtherSiteBasedInitial = GetInitialValues("OpenMatchesOtherSiteBased")

	GetInitialValues.Close
	Set GetInitialValues = Nothing
	
End If

If Request("status") = "addNew" Then

	
' Check for duplicate records
	
	Set DupCon = Server.CreateObject("ADODB.Connection")
	DupCon.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT Count(*) As NumberOfEntries FROM tbl_frmMCPPerformance WHERE AgencyID = '" & Request("AgencyIDN") & "' and Year = " & Request("Year")	& " and Month = " & Request("Month")
	Set DuplicateRecord = DupCon.Execute(query)
	numberOfExisting = DuplicateRecord("NumberOfEntries")
	DuplicateRecord.Close
	Set DuplicateRecord = Nothing
	DupCon.Close
	Set DupCon = Nothing
	
	
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	
	If(numberOfExisting = 0) Then
		Set RST = Server.CreateObject("ADODB.Recordset")
		RST.Open "SELECT * FROM tbl_frmMCPPerformance", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = Request("AgencyIDN")
		RST("Year") = Request("year")
		RST("Month") = Request("month")
		
		RST("OpenMatchesCommunityBased") = Request("frmMCPPerformanceOpenMatchesCommunityBased")
		RST("OpenMatchesSchoolBased") = Request("frmMCPPerformanceOpenMatchesSchoolBased")
		RST("OpenMatchesOtherSiteBased") = Request("frmMCPPerformanceOpenMatchesOtherSiteBased")
	
		RST("NewMatchesCommunityBased") = Request("frmMCPPerformanceNewMatchesCommunityBased")
		RST("NewMatchesSchoolBased") = Request("frmMCPPerformanceNewMatchesSchoolBased")
		RST("NewMatchesSiteBasedNonSchool") = Request("frmMCPPerformanceNewMatchesSiteBasedNonSchool")
		
		RST("ClosedMatchesCommunityBased") = Request("frmMCPPerformanceClosedMatchesCommunityBased")
		RST("ClosedMatchesSchoolBased") = Request("frmMCPPerformanceClosedMatchesSchoolBased")
		RST("ClosedMatchesOtherSiteBased") = Request("frmMCPPerformanceClosedMatchesOtherSiteBased")
		
		RST("CreateDate") = Now
		RST.Update
		RST.Close
		Set RST = Nothing
		form = "MCPPerformance"
		modtype = "new"
		m = Request("month")
		%>
		<!--#include file="../includes/modify_stamp.asp"-->
		<%	
		Con.Close
		Set Con = Nothing
		say = "thanks"
	Else
		say = "previouslyEdited"
		Con.Close
		Set Con = Nothing
	End If
		
ElseIf Request("status") = "editSave" Then


	Set Con = Server.CreateObject("ADODB.Connection")
	Set RST = Server.CreateObject("ADODB.Recordset")
	Con.Open "BBBSAforms", "sa","12sist12"
	RST.Open "SELECT * FROM tbl_frmMCPPerformance WHERE agencyID='" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")) & " AND Month=" & Int(Request("month")), Con, 1, 3
	
	OpenMatchesCommunityBasedInitial = RST("OpenMatchesCommunityBased")		
	OpenMatchesSchoolBasedInitial = RST("OpenMatchesSchoolBased")		
	OpenMatchesOtherSiteBasedInitial = RST("OpenMatchesOtherSiteBased")	
	
	RST("OpenMatchesCommunityBased") = Request("frmMCPPerformanceOpenMatchesCommunityBased")
	RST("OpenMatchesSchoolBased") = Request("frmMCPPerformanceOpenMatchesSchoolBased")
	RST("OpenMatchesOtherSiteBased") = Request("frmMCPPerformanceOpenMatchesOtherSiteBased")

	RST("ClosedMatchesCommunityBased") = Request("frmMCPPerformanceClosedMatchesCommunityBased")
	RST("ClosedMatchesSchoolBased") = Request("frmMCPPerformanceClosedMatchesSchoolBased")
	RST("ClosedMatchesOtherSiteBased") = Request("frmMCPPerformanceClosedMatchesOtherSiteBased")
	
	RST("NewMatchesCommunityBased") = Request("frmMCPPerformanceNewMatchesCommunityBased")
	RST("NewMatchesSchoolBased") = Request("frmMCPPerformanceNewMatchesSchoolBased")
	RST("NewMatchesSiteBasedNonSchool") = Request("frmMCPPerformanceNewMatchesSiteBasedNonSchool")
	

		
	jMod = RST("MCPPerformanceID") %>
	
	
	
	
	<%
	RST.Update
	RST.Close
	Set RST = Nothing
	form = "MCPPerformance"
	modtype = "edit"
	m = Request("month")
	%>
	<!--#include file="../includes/modify_stamp.asp"-->
	<%	
	Con.Close
	Set Con = Nothing
	say = "thanks" %>
	
	<!-- Run Propagation Routine -->

	<%	
	
	set Con = Server.CreateObject("ADODB.Connection")
		Con.Open "BBBSAforms", "sa", "12sist12"

		CommunityDifference = int(trim(Request("frmMCPPerformanceOpenMatchesCommunityBased"))) - OpenMatchesCommunityBasedInitial
		sql_CommunityDifference = "sp_PropagatePerformanceChanges '" & Session("AgencyIDN") & "'," & "'" & right("00" & Int(Request("month")),2) & "/01/" & Int(Request("year")) & "','>','tbl_frmMCPPerformance','OpenMatchesCommunityBased','+','" & CommunityDifference & "'"

		SchoolDifference = int(trim(Request("frmMCPPerformanceOpenMatchesSchoolBased"))) - OpenMatchesSchoolBasedInitial
		sql_SchoolDifference = "sp_PropagatePerformanceChanges '" & Session("AgencyIDN") & "'," & "'" & right("00" & Int(Request("month")),2) & "/01/" & Int(Request("year")) & "','>','tbl_frmMCPPerformance','OpenMatchesSchoolBased','+','" & SchoolDifference & "'"

		OtherDifference = Request("frmMCPPerformanceOpenMatchesOtherSiteBased") - OpenMatchesOtherSiteBasedInitial
		sql_OtherDifference = "sp_PropagatePerformanceChanges '" & Session("AgencyIDN") & "'," & "'" & right("00" & Int(Request("month")),2) & "/01/" & Int(Request("year")) & "','>','tbl_frmMCPPerformance','OpenMatchesOtherSiteBased','+','" & OtherDifference & "'"

		
		if CommunityDifference <> 0 then		
		%><%=int(trim(Request("frmMCPPerformanceOpenMatchesCommunityBased")))%> - <%=OpenMatchesCommunityBasedInitial%> = <%=communitydifference%><%
			Set rs = Con.Execute(sql_CommunityDifference)
		end if
	
		if SchoolDifference <> 0 then		
			Set rs = Con.Execute(sql_SchoolDifference)
		end if

		if OtherDifference <> 0 then		
			Set rs = Con.Execute(sql_OtherDifference)
		end if

		
	Con.Close
	Set Con = Nothing
	
	
		
	
	
	
ElseIf Request("status") = "editOld" Then
	say = "edit"
Else
	say = "form"
End If
 %>


<% dim HelpId
HelpId = 0
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>MCP Grant Performance</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
<script language="javascript">
<!--	

function checkForIntegerCommas(valueToCheck)
{
	var myRegularExpression = /^[0-9]+(,[0-9]{3})*$/;  // Checks for integer with or without commas
	if(!(myRegularExpression.test(valueToCheck)))
	{
		alert("Please make sure you have entered a whole number with no spaces.\n Do not leave this field blank."); 
	} 
}

function addUpOpenComm()
{
	var box1 = Number(document.frmMCPPerformance.frmMCPPerformancePrevOpenComm.value)
	var box2 = Number(document.frmMCPPerformance.frmMCPPerformanceClosedMatchesCommunityBased.value)
	var box3 = Number(document.frmMCPPerformance.frmMCPPerformanceNewMatchesCommunityBased.value)	

	var boxtotal = box1 - box2 + box3
	document.frmMCPPerformance.frmMCPPerformanceOpenMatchesCommunityBased.value = boxtotal
	
}

function addUpOpenSchool()
{
	var box1 = Number(document.frmMCPPerformance.frmMCPPerformancePrevOpenSchool.value)
	var box2 = Number(document.frmMCPPerformance.frmMCPPerformanceClosedMatchesSchoolBased.value)
	var box3 = Number(document.frmMCPPerformance.frmMCPPerformanceNewMatchesSchoolBased.value)	

	var boxtotal = box1 - box2 + box3
	document.frmMCPPerformance.frmMCPPerformanceOpenMatchesSchoolBased.value = boxtotal
	
}


function addUpOpenNonSchool()
{
	var box1 = Number(document.frmMCPPerformance.frmMCPPerformancePrevOpenOther.value)
	var box2 = Number(document.frmMCPPerformance.frmMCPPerformanceClosedMatchesOtherSiteBased.value)
	var box3 = Number(document.frmMCPPerformance.frmMCPPerformanceNewMatchesSiteBasedNonSchool.value)	

	var boxtotal = box1 - box2 + box3
	document.frmMCPPerformance.frmMCPPerformanceOpenMatchesOtherSiteBased.value = boxtotal
	
}



function validateForm()
{	
	
	var onlyInteger = /^[0-9]+(,[0-9]{3})*$/;
	var PrevOpenComm = new Number(frmMCPPerformance.frmMCPPerformancePrevOpenComm.value)
	var CurCommOpen = new Number(frmMCPPerformance.frmMCPPerformanceOpenMatchesCommunityBased.value)
	var CurCommClosed = new Number(frmMCPPerformance.frmMCPPerformanceClosedMatchesCommunityBased.value)	
	var CurCommTotal = CurCommOpen + CurCommClosed
	
	var PrevOpenSchool = new Number(frmMCPPerformance.frmMCPPerformancePrevOpenSchool.value)
	var CurSchoolOpen = new Number(frmMCPPerformance.frmMCPPerformanceOpenMatchesSchoolBased.value)
	var CurSchoolClosed = new Number(frmMCPPerformance.frmMCPPerformanceClosedMatchesSchoolBased.value)	
	var CurSchoolTotal = CurSchoolOpen + CurSchoolClosed
	
	var PrevOpenOther = new Number(frmMCPPerformance.frmMCPPerformancePrevOpenOther.value)

	var CurOtherOpen = new Number(frmMCPPerformance.frmMCPPerformanceOpenMatchesOtherSiteBased.value)
	var CurOtherClosed = new Number(frmMCPPerformance.frmMCPPerformanceClosedMatchesOtherSiteBased.value)	
	var CurOtherTotal = CurOtherOpen + CurOtherClosed
	
		
	
//	if (CurCommTotal.valueOf() < PrevOpenComm)	
//		{alert( "ERROR:\n\nThe sum of your current OPEN PLUS your current CLOSED Community-Based matches ("+CurCommOpen+"+"+CurCommClosed+") must be greater than your previous month's OPEN Community-Based matches ("+PrevOpenComm+").  Please Correct and re-SAVE.");document.frmMCPPerformance.frmMCPPerformanceOpenMatchesCommunityBased.focus();}

//	else if (CurSchoolTotal.valueOf() < PrevOpenSchool)	
//		{alert( "ERROR:\n\nThe sum of your current OPEN PLUS your CLOSED School-Based matches ("+CurSchoolOpen+"+"+CurSchoolClosed+") must be greater than your previous month's OPEN School-Based matches ("+PrevOpenSchool+").  Please correct and re-SAVE.");document.frmMCPPerformance.frmMCPPerformanceOpenMatchesSchoolBased.focus();}

//	else if (CurOtherTotal.valueOf() < PrevOpenOther)	
//		{alert( "ERROR:\n\nThe sum of your current OPEN PLUS your CLOSED Other Site-Based matches ("+CurOtherOpen+"+"+CurOtherClosed+") must be greater than your previous month's OPEN Other Site-Based matches ("+PrevOpenOther+").  Please correct and re-SAVE.");document.frmMCPPerformance.frmMCPPerformanceOpenMatchesOtherSiteBased.focus();}

	
	
	if(document.frmMCPPerformance.frmMCPPerformanceOpenMatchesCommunityBased.value == "")
		{alert("Please complete all form fields");document.frmMCPPerformance.frmMCPPerformanceOpenMatchesCommunityBased.focus();}		
	else if(document.frmMCPPerformance.frmMCPPerformanceOpenMatchesSchoolBased.value == "")
		{alert("Please complete all form fields");document.frmMCPPerformance.frmMCPPerformanceOpenMatchesSchoolBased.focus();}
	else if(document.frmMCPPerformance.frmMCPPerformanceOpenMatchesOtherSiteBased.value == "")
		{alert("Please complete all form fields");document.frmMCPPerformance.frmMCPPerformanceOpenMatchesOtherSiteBased.focus();}
	
	else if(document.frmMCPPerformance.frmMCPPerformanceClosedMatchesCommunityBased.value == "")
		{alert("Please complete all form fields");document.frmMCPPerformance.frmMCPPerformanceClosedMatchesCommunityBased.focus();}
	else if(document.frmMCPPerformance.frmMCPPerformanceClosedMatchesSchoolBased.value == "")
		{alert("Please complete all form fields");document.frmMCPPerformance.frmMCPPerformanceClosedMatchesSchoolBased.focus();}
	else if(document.frmMCPPerformance.frmMCPPerformanceClosedMatchesOtherSiteBased.value == "")
		{alert("Please complete all form fields");document.frmMCPPerformance.frmMCPPerformanceClosedMatchesOtherSiteBased.focus();}
	
			
	else if(document.frmMCPPerformance.frmMCPPerformanceOpenMatchesCommunityBased.value < 0)
		{alert("Open Community Based Matches at the end of the month cannot be less than zero.");document.frmMCPPerformance.frmMCPPerformanceOpenMatchesCommunityBased.focus();}
	else if(!(onlyInteger.test(document.frmMCPPerformance.frmMCPPerformanceOpenMatchesCommunityBased.value)))
		{alert(document.frmMCPPerformance.frmMCPPerformanceOpenMatchesCommunityBased.value + " is an invalid number");document.frmMCPPerformance.frmMCPPerformanceOpenMatchesCommunityBased.focus();}

	else if(document.frmMCPPerformance.frmMCPPerformanceOpenMatchesSchoolBased.value < 0)
		{alert("Open School-Based Matches at the end of the month cannot be less than zero.");document.frmMCPPerformance.frmMCPPerformanceOpenMatchesSchoolBased.focus();}
	else if(!(onlyInteger.test(document.frmMCPPerformance.frmMCPPerformanceOpenMatchesSchoolBased.value)))
		{alert(document.frmMCPPerformance.frmMCPPerformanceOpenMatchesSchoolBased.value + " is an invalid number");document.frmMCPPerformance.frmMCPPerformanceOpenMatchesSchoolBased.focus();}

	else if(document.frmMCPPerformance.frmMCPPerformanceOpenMatchesOtherSiteBased.value < 0)
		{alert("Open Non-School Site-Based Matches at the end of the month cannot be less than zero.");document.frmMCPPerformance.frmMCPPerformanceOpenMatchesOtherSiteBased.focus();}
	else if(!(onlyInteger.test(document.frmMCPPerformance.frmMCPPerformanceOpenMatchesOtherSiteBased.value)))
		{alert(document.frmMCPPerformance.frmMCPPerformanceOpenMatchesOtherSiteBased.value + " is an invalid number");document.frmMCPPerformance.frmMCPPerformanceOpenMatchesOtherSiteBased.focus();}

	else if(document.frmMCPPerformance.frmMCPPerformanceClosedMatchesCommunityBased.value < 0)
		{alert("Closed Community Based matches cannot be less than zero.");document.frmMCPPerformance.frmMCPPerformanceClosedMatchesCommunityBased.focus();}
	else if(!(onlyInteger.test(document.frmMCPPerformance.frmMCPPerformanceClosedMatchesCommunityBased.value)))
		{alert(document.frmMCPPerformance.frmMCPPerformanceClosedMatchesCommunityBased.value + " is an invalid number");document.frmMCPPerformance.frmMCPPerformanceClosedMatchesCommunityBased.focus();}

	else if(document.frmMCPPerformance.frmMCPPerformanceClosedMatchesSchoolBased.value < 0)
		{alert("Closed School-Based matches cannot be less than zero.");document.frmMCPPerformance.frmMCPPerformanceClosedMatchesSchoolBased.focus();}
	else if(!(onlyInteger.test(document.frmMCPPerformance.frmMCPPerformanceClosedMatchesSchoolBased.value)))
		{alert(document.frmMCPPerformance.frmMCPPerformanceClosedMatchesSchoolBased.value + " is an invalid number");document.frmMCPPerformance.frmMCPPerformanceClosedMatchesSchoolBased.focus();}

	else if(document.frmMCPPerformance.frmMCPPerformanceClosedMatchesOtherSiteBased.value < 0)
		{alert("Closed Non-School Site-Based matches cannot be less than zero.");document.frmMCPPerformance.frmMCPPerformanceClosedMatchesOtherSiteBased.focus();}	
	else if(!(onlyInteger.test(document.frmMCPPerformance.frmMCPPerformanceClosedMatchesOtherSiteBased.value)))
		{alert(document.frmMCPPerformance.frmMCPPerformanceClosedMatchesOtherSiteBased.value + " is an invalid number");document.frmMCPPerformance.frmMCPPerformanceClosedMatchesOtherSiteBased.focus();}
		
	else
		document.frmMCPPerformance.submit();	
}		

function getNextElement (field) 
{
	var form = field.form;
  	for (var e = 0; e < form.elements.length; e++)
    	if (field == form.elements[e])
      	break;
  	return form.elements[++e % form.elements.length];
}

//-->	
</script>

<!-- Popup Window Script -->
<SCRIPT LANGUAGE = "JavaScript">

<!-- Begin
function NewWindow(mypage, myname, w, h) {
var winl = (screen.width - w) / 2;
var wint = (screen.height - h) / 2;
winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',resizable, scrollbars'
win = window.open(mypage, myname, winprops)
if (parseInt(navigator.appVersion) >= 4) { win.window.focus(); }
}

//  End -->

</SCRIPT>

<% If say="form" Then %>
	<body onLoad="addUpOpenComm(); addUpOpenSchool(); addUpOpenNonSchool(); ">
<% else %>
	<body>
<% end if %>

<% ' <!--#include file="../includes/top_nav_forms_monthly.inc"--><!-- include file has </head> and <body> tags --><br>      %>
<!--#include file="../includes/surveytitle.inc"-->


<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_football.jpg" alt="" width="220" height="477" border="0"></td> 
<td valign="top" align="left">

<% If say = "thanks" Then %>

<font class="formMain">
<br><br>
<strong>Thank you!</strong> Your information has been saved in the BBBS database.<br>
To choose another form, please select the form type from the choices above.
<br><br>
<i>Please note: These changes will not be reflected in the <strong>Agency Profile</strong> (in the My Agency Page and the Agency Directory) for 24 hours.</i>
</font>
<br>
<!--#include file="../includes/contact_info.inc"-->
<br>


<% ElseIf say <> "thanks" Then  %>


<form name="frmMCPPerformance" action="MCPPerformance_edit.asp" method="post"> <!-- onsubmit="return submitFormValidate(this)"> -->
<!--#include file="../includes/form_stamp.asp"-->

<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmMCPPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & " AND Month=" & Int(Request("m"))
	Set GetPerformance = Con.Execute(query)
	
		
	
 %>
<input type="hidden" name="status" value="editSave">
<% Else %>
<input type="hidden" name="status" value="addNew">
<%
End If
 %>
 
<%
If say = "previouslyEdited" Then
%>
<p class="formMain"><br>We're sorry, but this form was previously completed. To make changes please <a href="monthly.asp">reselect</a> the 
appropriate form and year and update the existing information.</p>
<%
Response.End
End If 
%> 
<br>



		<table width="550" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063" >
		<tr>
			<td colspan="7" class="formHeader">MCP GRANT PERFORMANCE - <%= MonthName(Request("m"), False) & " " & Request("y") %></td>
		</tr>

		<tr>
			<td colspan="7" class="formMain"><font color="#ff0000"><div align="center"><strong>Please Note: </strong>After entering your information, you <strong>must</strong> click on the "Save" button at the bottom of the form and wait for the "Thank You" screen or your changes will be lost.</div></font></td>
		</tr>
		
			<tr>
				<td>&nbsp;</td>
				<td align="center" valign="middle" class="formMain">Community Based</td>
				<td align="center" valign="middle" class="formMain">School Based</td>
				<td align="center" valign="middle" class="formMain">Non-School Site Based</td>
				
			</tr>
			<tr>
			
			<!-- Open Matches at the Beginning of the Month -->
			
			<% if y = "2005" and m = "1" then %>
			
			<% else %>
				<tr>
					<td align="center" valign="middle" class="formMain">OPEN/ACTIVE&nbsp;matches<br>on&nbsp;the&nbsp;<strong>FIRST</strong>&nbsp;day&nbsp;of<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
					<td align="center" valign="middle" class="formMain" bgcolor="#c0c0c0">
						<input type="text" class="formMain" size="5" value="<%=Request("PrevOpenComm")%>" readonly onFocus="addUpOpenComm();"><br><span class="formSubHead">calculated by system</span>
					</td>
					<td align="center" valign="middle" class="formMain" bgcolor="#c0c0c0">
						<input type="text" class="formMain" size="5" value="<%=Request("PrevOpenSchool")%>" readonly onFocus="addUpOpenSchool();"><br><span class="formSubHead">calculated by system</span>
					</td>
					<td align="center" valign="middle" class="formMain" bgcolor="#c0c0c0">
						<input type="text" class="formMain" size="5" value="<%=Request("PrevOpenOther")%>" readonly onFocus="addUpOpenNonSchool();"><br><span class="formSubHead">calculated by system</span>				
					</td>
	
				</tr>		
			<% end if %>

			<!-- Matches Closed During the Month -->
			

			<tr>
				<td align="center" valign="middle" class="formMain">Matches&nbsp;CLOSED&nbsp;during<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesCommunityBased") %><% Else %>0<% End If %>" name="frmMCPPerformanceClosedMatchesCommunityBased" tabindex="1" onchange="checkForIntegerCommas(this.value); addUpOpenComm();">
				</td>
				
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesSchoolBased") %><% Else %>0<% End If %>" name="frmMCPPerformanceClosedMatchesSchoolBased" tabindex="2" onchange="checkForIntegerCommas(this.value); addUpOpenSchool();">
				</td>
				
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesOtherSiteBased") %><% Else %>0<% End If %>" name="frmMCPPerformanceClosedMatchesOtherSiteBased" tabindex="3" onchange="checkForIntegerCommas(this.value); addUpOpenNonSchool();">
				</td>
				
				
			</tr>
	
			<!-- New Matches Opened During the Month -->		
			<tr>
			
			<% if y = "2005" and m = "1" then %>
				<td align="center" valign="middle" class="formMain"><b><font color="red">ONE-TIME Baseline Entry for January 2005:<br></font></b>Enter any MCP matches that existed<br><em>prior</em> to January 2005 <strong>PLUS</strong><br>any new MCP matches created <strong>DURING</strong> January 2005.</td>															
			<% else %>
				<td align="center" valign="middle" class="formMain"><br>NEW&nbsp;matches opened<br>during&nbsp;<b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>			
			<% end if %>
		

				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("NewMatchesCommunityBased") %><% Else %>0<% End If %>" name="frmMCPPerformanceNewMatchesCommunityBased" tabindex="7" onchange="checkForIntegerCommas(this.value); addUpOpenComm();">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("NewMatchesSchoolBased") %><% Else %>0<% End If %>" name="frmMCPPerformanceNewMatchesSchoolBased" tabindex="8" onchange="checkForIntegerCommas(this.value); addUpOpenSchool();">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("NewMatchesSiteBasedNonSchool") %><% Else %>0<% End If %>" name="frmMCPPerformanceNewMatchesSiteBasedNonSchool" tabindex="9" onchange="checkForIntegerCommas(this.value); addUpOpenNonSchool();">
				</td>

			</tr>					
			
			
			</tr>			
			
			<!-- Populate Initial Values of Open Matches for Propagation Comparison -->
			<input type="hidden" name="frmMCPPerformanceOpenMatchesCommunityBasedBegin" value="<%=OpenMatchesCommunityBasedBegin%>">			
			<input type="hidden" name="frmMCPPerformanceOpenMatchesSchoolBasedBegin" value="<%=OpenMatchesSchoolBasedBegin%>">						
			<input type="hidden" name="frmMCPPerformanceOpenMatchesOtherSiteBasedBegin" value="<%=OpenMatchesOtherSiteBasedBegin%>">
			
			<!-- Open Matches on the Last Day of the Month -->
			<tr>
			
				<td align="center" valign="middle" class="formMain">OPEN/ACTIVE&nbsp;matches<br>on&nbsp;the&nbsp;<strong>LAST</strong>&nbsp;day&nbsp;of<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				
				<td align="center" valign="middle" bgcolor="#c0c0c0" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesCommunityBased") %><% Else %>0<% End If %>" name="frmMCPPerformanceOpenMatchesCommunityBased" onFocus="addUpOpenComm();" onchange="addUpOpenComm();" readonly><br><span class="formSubHead">calculated by system</span>
				</td>
				
				<td align="center" valign="middle" bgcolor="#c0c0c0" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesSchoolBased") %><% Else %>0<% End If %>" name="frmMCPPerformanceOpenMatchesSchoolBased" onFocus="addUpOpenSchool();" onchange="addUpOpenSchool();" readonly><br><span class="formSubHead">calculated by system</span>				
				</td>
				
				<td align="center" valign="middle" bgcolor="#c0c0c0" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesOtherSiteBased") %><% Else %>0<% End If %>" name="frmMCPPerformanceOpenMatchesOtherSiteBased" onFocus="addUpOpenNonSchool();" onchange="addUpOpenNonSchool();" readonly><br><span class="formSubHead">calculated by system</span>
				</td>
				
			</tr>
			



<!-- ADD PREVIOUS MONTH'S MATCH FIELDS TO FORM FOR COMPARISON -->
<input type="hidden" class="formMain" size="5" maxlength="10" value="<%=Request("PrevOpenComm")%>" name="frmMCPPerformancePrevOpenComm" onchange="checkForIntegerCommas(this.value);">	
<input type="hidden" class="formMain" size="5" maxlength="10" value="<%=Request("PrevOpenSchool")%>" name="frmMCPPerformancePrevOpenSchool" onchange="checkForIntegerCommas(this.value);">	
<input type="hidden" class="formMain" size="5" maxlength="10" value="<%=Request("PrevOpenOther")%>" name="frmMCPPerformancePrevOpenOther" onchange="checkForIntegerCommas(this.value);">	

	<tr>
		<td colspan="7" class="formHeader">
		<input type="button" value="Save Form" class="formMainBold" onclick="validateForm(); addUpOpenComm(); addUpOpenSchool(); addUpOpenNonSchool(); return false;">
		</td>
	</tr>
	<tr>
		<td colspan="7"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
	</tr>
	</table>

</td>
</tr>
</table>

									
<% 
If say = "edit" Then
	GetPerformance.Close
	Set GetPerformance = Nothing

	Con.Close
	Set Con = Nothing


	
End If


 %>


</form>
<% End If %>
</body>
</html>
