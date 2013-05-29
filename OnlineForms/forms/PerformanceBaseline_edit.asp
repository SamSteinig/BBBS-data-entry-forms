<% 
If Request("status") = "addNew" Then

' Check for duplicate records
	
	Set DupCon = Server.CreateObject("ADODB.Connection")
	DupCon.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT Count(*) As NumberOfEntries FROM tbl_frmPerformanceBaseline WHERE AgencyID = '" & Request("AgencyIDN") & "' and Year = " & Request("Year")	
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
		RST.Open "SELECT * FROM tbl_frmPerformanceBaseline", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = Request("AgencyIDN")
		RST("Year") = Request("year")
'		RST("Month") = Request("month")
		RST("OpenMatchesCommunityBased") = Request("frmPerformanceOpenMatchesCommunityBased")
		RST("OpenMatchesSchoolBased") = Request("frmPerformanceOpenMatchesSchoolBased")
		RST("OpenMatchesOtherSiteBased") = Request("frmPerformanceOpenMatchesOtherSiteBased")
		RST("OpenMatchesGroupMentoring") = Request("frmPerformanceOpenMatchesGroupMentoring")
		RST("OpenMatchesSpecialProgramsMentoring") = Request("frmPerformanceOpenMatchesSpecialProgramsMentoring")
		RST("OpenMatchesSpecialProgramsNonMentoring") = Request("frmPerformanceOpenMatchesSpecialProgramsNonMentoring")
		RST("ClosedMatchesCommunityBased") = Request("frmPerformanceClosedMatchesCommunityBased")
		RST("ClosedMatchesSchoolBased") = Request("frmPerformanceClosedMatchesSchoolBased")
		RST("ClosedMatchesOtherSiteBased") = Request("frmPerformanceClosedMatchesOtherSiteBased")
		RST("ClosedMatchesGroupMentoring") = Request("frmPerformanceClosedMatchesGroupMentoring")
		RST("ClosedMatchesSpecialProgramsMentoring") = Request("frmPerformanceClosedMatchesSpecialProgramsMentoring")
		RST("ClosedMatchesSpecialProgramsNonMentoring") = Request("frmPerformanceClosedMatchesSpecialProgramsNonMentoring")
		RST("CreateDate") = Now
		RST.Update
		RST.Close
		Set RST = Nothing
		form = "PerformanceBaseline"
		modtype = "new"
'		m = Request("month")
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
	RST.Open "SELECT * FROM tbl_frmPerformanceBaseline WHERE agencyID='" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")), Con, 1, 3
	RST("OpenMatchesCommunityBased") = Request("frmPerformanceOpenMatchesCommunityBased")
	RST("OpenMatchesSchoolBased") = Request("frmPerformanceOpenMatchesSchoolBased")
	RST("OpenMatchesOtherSiteBased") = Request("frmPerformanceOpenMatchesOtherSiteBased")
	RST("OpenMatchesGroupMentoring") = Request("frmPerformanceOpenMatchesGroupMentoring")
	RST("OpenMatchesSpecialProgramsMentoring") = Request("frmPerformanceOpenMatchesSpecialProgramsMentoring")
	RST("OpenMatchesSpecialProgramsNonMentoring") = Request("frmPerformanceOpenMatchesSpecialProgramsNonMentoring")
	RST("ClosedMatchesCommunityBased") = Request("frmPerformanceClosedMatchesCommunityBased")
	RST("ClosedMatchesSchoolBased") = Request("frmPerformanceClosedMatchesSchoolBased")
	RST("ClosedMatchesOtherSiteBased") = Request("frmPerformanceClosedMatchesOtherSiteBased")
	RST("ClosedMatchesGroupMentoring") = Request("frmPerformanceClosedMatchesGroupMentoring")
	RST("ClosedMatchesSpecialProgramsMentoring") = Request("frmPerformanceClosedMatchesSpecialProgramsMentoring")
	RST("ClosedMatchesSpecialProgramsNonMentoring") = Request("frmPerformanceClosedMatchesSpecialProgramsNonMentoring")
	jMod = RST("PerformanceBaselineID")
	RST.Update
	RST.Close
	Set RST = Nothing
	form = "PerformanceBaseline"
	modtype = "edit"
'	m = Request("month")
	%>
	<!--#include file="../includes/modify_stamp.asp"-->
	<%	
	Con.Close
	Set Con = Nothing
	say = "thanks"
ElseIf Request("status") = "editOld" Then
	say = "edit"
Else
	say = "form"
End If
 %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Performance</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
<script language="javascript">
<!--	

function checkForIntegerCommas(valueToCheck)
{
	var myRegularExpression = /^[0-9]+(,[0-9]{3})*$/;  // Checks for integer with or without commas
	if(!(myRegularExpression.test(valueToCheck)))
	{
		alert("Please make sure you have entered a whole number with no spaces.\n We cannot process letters or words."); 
	} 
}

function validateForm()
{	
	var onlyInteger = /^[0-9]+(,[0-9]{3})*$/;
	
	if(document.frmPerformance.frmPerformanceOpenMatchesCommunityBased.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceOpenMatchesCommunityBased.focus();}
	else if(document.frmPerformance.frmPerformanceOpenMatchesSchoolBased.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceOpenMatchesSchoolBased.focus();}
	else if(document.frmPerformance.frmPerformanceOpenMatchesOtherSiteBased.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceOpenMatchesOtherSiteBased.focus();}
	else if(document.frmPerformance.frmPerformanceOpenMatchesGroupMentoring.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceOpenMatchesGroupMentoring.focus();}
	else if(document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsMentoring.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsMentoring.focus();}
	else if(document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsNonMentoring.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsNonMentoring.focus();}
	else if(document.frmPerformance.frmPerformanceClosedMatchesCommunityBased.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceClosedMatchesCommunityBased.focus();}
	else if(document.frmPerformance.frmPerformanceClosedMatchesSchoolBased.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceClosedMatchesSchoolBased.focus();}
	else if(document.frmPerformance.frmPerformanceClosedMatchesOtherSiteBased.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceClosedMatchesOtherSiteBased.focus();}
	else if(document.frmPerformance.frmPerformanceClosedMatchesGroupMentoring.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceClosedMatchesGroupMentoring.focus();}
	else if(document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsMentoring.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsMentoring.focus();}
	else if(document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsNonMentoring.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsNonMentoring.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceOpenMatchesCommunityBased.value)))
		{alert(document.frmPerformance.frmPerformanceOpenMatchesCommunityBased.value + " is an invalid number");document.frmPerformance.frmPerformanceOpenMatchesCommunityBased.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceOpenMatchesSchoolBased.value)))
		{alert(document.frmPerformance.frmPerformanceOpenMatchesSchoolBased.value + " is an invalid number");document.frmPerformance.frmPerformanceOpenMatchesSchoolBased.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceOpenMatchesOtherSiteBased.value)))
		{alert(document.frmPerformance.frmPerformanceOpenMatchesOtherSiteBased.value + " is an invalid number");document.frmPerformance.frmPerformanceOpenMatchesOtherSiteBased.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceOpenMatchesGroupMentoring.value)))
		{alert(document.frmPerformance.frmPerformanceOpenMatchesGroupMentoring.value + " is an invalid number");document.frmPerformance.frmPerformanceOpenMatchesGroupMentoring.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsMentoring.value)))
		{alert(document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsMentoring.value + " is an invalid number");document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsMentoring.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsNonMentoring.value)))
		{alert(document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsNonMentoring.value + " is an invalid number");document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsNonMentoring.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceClosedMatchesCommunityBased.value)))
		{alert(document.frmPerformance.frmPerformanceClosedMatchesCommunityBased.value + " is an invalid number");document.frmPerformance.frmPerformanceClosedMatchesCommunityBased.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceClosedMatchesSchoolBased.value)))
		{alert(document.frmPerformance.frmPerformanceClosedMatchesSchoolBased.value + " is an invalid number");document.frmPerformance.frmPerformanceClosedMatchesSchoolBased.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceClosedMatchesOtherSiteBased.value)))
		{alert(document.frmPerformance.frmPerformanceClosedMatchesOtherSiteBased.value + " is an invalid number");document.frmPerformance.frmPerformanceClosedMatchesOtherSiteBased.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceClosedMatchesGroupMentoring.value)))
		{alert(document.frmPerformance.frmPerformanceClosedMatchesGroupMentoring.value + " is an invalid number");document.frmPerformance.frmPerformanceClosedMatchesGroupMentoring.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsMentoring.value)))
		{alert(document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsMentoring.value + " is an invalid number");document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsMentoring.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsNonMentoring.value)))
		{alert(document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsNonMentoring.value + " is an invalid number");document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsNonMentoring.focus();}
	else
		document.frmPerformance.submit();	
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
	
<% '<!--#include file="../includes/top_nav_forms_yearly.inc"--><!-- include file has </head> and <body> tags --><br>     %>
<!--#include file="../includes/surveytitle.inc"-->
<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_fishing.jpg" alt="" width="220" height="477" border="0"></td>
<td width="100%" valign="top">
<br>

<% If say = "thanks" Then %>
<font class="formMain">
Thank you! Your information has been saved in the BBBS database.<br>
To choose another form, please select the form type from the choices above.
</font>
<br>
<!--#include file="../includes/contact_info.inc"-->
<br>

<% ElseIf say <> "thanks" Then  %>
<table width="400" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063">
<form name="frmPerformance" action="PerformanceBaseline_edit.asp" method="post"> <!-- onsubmit="return submitFormValidate(this)"> -->
<!--#include file="../includes/form_stamp.asp"-->

<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmPerformanceBaseline WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
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
<p class="formMain">We're sorry, but this form was previously completed. To make changes please <a href="monthly.asp">reselect</a> the 
appropriate form and year and update the existing information.</p>
<%
Response.End
End If 
%> 

		<tr>
			<td colspan="7" class="formHeader">END OF YEAR PERFORMANCE</td>
		</tr>
			<tr>
				<td align="center" valign="middle" class="formMain">&nbsp;</td>
				<td align="center" valign="middle" class="formMain">Community Based</td>
				<td align="center" valign="middle" class="formMain">School Based</td>
				<td align="center" valign="middle" class="formMain">Other Site Based</td>
				<td align="center" valign="middle" class="formMain">Group Mentoring</td>
				<td align="center" valign="middle" class="formMain">Special Programs: Mentoring</td>
				<td align="center" valign="middle" class="formMain">Special Programs: Non-Mentoring</td>
			</tr>
			<tr>

				<td align="center" valign="middle" class="formMain"><b>Opened/Active</b>&nbsp;Matches</td> 
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesCommunityBased") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesCommunityBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesSchoolBased") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesSchoolBased" tabindex="3" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesOtherSiteBased") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesOtherSiteBased" tabindex="5" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesGroupMentoring") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesGroupMentoring" tabindex="7" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"   class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesSpecialProgramsMentoring") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesSpecialProgramsMentoring" tabindex="9" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesSpecialProgramsNonMentoring") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesSpecialProgramsNonMentoring" tabindex="11" onchange="checkForIntegerCommas(this.value);">
				</td>
			</tr>
			<tr>

				<td align="center" valign="middle" class="formMain">Matches&nbsp;CLOSED</td> 

				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesCommunityBased") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesCommunityBased" tabindex="2" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesSchoolBased") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesSchoolBased" tabindex="4" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesOtherSiteBased") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesOtherSiteBased" tabindex="6" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesGroupMentoring") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesGroupMentoring" tabindex="8" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesSpecialProgramsMentoring") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesSpecialProgramsMentoring" tabindex="10" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesSpecialProgramsNonMentoring") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesSpecialProgramsNonMentoring" tabindex="12" onchange="checkForIntegerCommas(this.value);">
				</td>
			</tr>
			<tr>
				<td colspan="7" class="formHeader">
					<input type="button" value="Save Form" class="formMainBold" onclick="validateForm(); return false;">
				</td>
				
			<tr>
				<td colspan="7"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
			</tr>
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

</td>
</tr>
</table>
</body>
</html>
