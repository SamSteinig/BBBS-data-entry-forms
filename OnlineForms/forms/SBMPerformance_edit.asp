<!--#include file="../includes/NAD_BE.asp" -->

<% 

If Request("status") = "addNew" Then

	
	
' Check for duplicate records
	
	Set DupCon = Server.CreateObject("ADODB.Connection")
	DupCon.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT Count(*) As NumberOfEntries FROM tbl_frmSBMPerformance WHERE AgencyID = '" & Request("AgencyIDN") & "' and Year = " & Request("Year")	& " and Month = " & Request("Month")
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
		RST.Open "SELECT * FROM tbl_frmSBMPerformance", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = Request("AgencyIDN")
		RST("Year") = Request("year")
		RST("Month") = Request("month")
		
		RST("SBMVolunteersInEnrollmentProcess") = Request("frmSBMPerformanceSBMVolunteersInEnrollmentProcess")		
		RST("SBMAmountRaisedTowardsMatchPledge") = Request("frmSBMPerformanceSBMAmountRaisedTowardsMatchPledge")			

		
		RST("CreateDate") = Now
		RST.Update
		RST.Close
		Set RST = Nothing
		form = "SBMPerformance"
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
	RST.Open "SELECT * FROM tbl_frmSBMPerformance WHERE agencyID='" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")) & " AND Month=" & Int(Request("month")), Con, 1, 3

	RST("SBMVolunteersInEnrollmentProcess") = Request("frmSBMPerformanceSBMVolunteersInEnrollmentProcess")		
	RST("SBMAmountRaisedTowardsMatchPledge") = Request("frmSBMPerformanceSBMAmountRaisedTowardsMatchPledge")			
		
	jMod = RST("SBMPerformanceID") %>
	
	
	
	
	<%
	RST.Update
	RST.Close
	Set RST = Nothing
	form = "SBMPerformance"
	modtype = "edit"
	m = Request("month")
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


<% dim HelpId
HelpId = 0
%>



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Performance</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
<script language="javascript">
<!--	


//Field Validations

function checkForIntegerCommas(valueToCheck)
{
	var myRegularExpression = /^[0-9]+(,[0-9]{3})*$/;  // Checks for integer with or without commas
	if(!(myRegularExpression.test(valueToCheck)))
	{
		alert("Please make sure you have entered a whole number with no spaces.\n Do not leave this field blank."); 
	} 
}

function validateForm()
{	
	
	var onlyInteger = /^[0-9]+(,[0-9]{3})*$/;
	

	if(document.frmSBMPerformance.frmSBMPerformanceSBMVolunteersInEnrollmentProcess.value == "")
		{alert("Please complete all form fields");document.frmSBMPerformance.frmSBMPerformanceSBMVolunteersInEnrollmentProcess.focus();}		
	else if(document.frmSBMPerformance.frmSBMPerformanceSBMAmountRaisedTowardsMatchPledge.value == "")
		{alert("Please complete all form fields");document.frmSBMPerformance.frmSBMPerformanceSBMAmountRaisedTowardsMatchPledge.focus();}					
		
	else
		document.frmSBMPerformance.submit();	
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


<form name="frmSBMPerformance" action="SBMPerformance_edit.asp" method="post"> <!-- onsubmit="return submitFormValidate(this)"> -->
<!--#include file="../includes/form_stamp.asp"-->

<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmSBMPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & " AND Month=" & Int(Request("m"))
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
			<td colspan="2" class="formHeader">PERFORMANCE - SBM GRANT PROGRESS REPORT<BR><%= MonthName(Request("m"), False) & " " & Request("y") %></td>
		</tr>
		
		<tr>
			<td colspan="2" class="formMain"><font color="#ff0000"><div align="center"><strong>Please Note: </strong>After entering your information, you <strong>must</strong> click on the "Save" button at the bottom of the form and wait for the "Thank You" screen or your changes will be lost.</div></font></td>
		</tr>


<!--
			<tr>
				<td valign="middle" class="formMain">Number of Volunteers in Enrollment Process</td>
				<td valign="middle" class="formMain">
					<input type="text"  class="formMain"  size="5" maxlength="10" value="<% 'If say = "edit" Then %><% '= GetPerformance("SBMVolunteersInEnrollmentProcess") %><%' Else %>0<%' End If %>" name="frmSBMPerformanceSBMVolunteersInEnrollmentProcess">									
				</td>
			</tr>
			
-->			

			<tr>
				<td valign="middle" class="formMain">Amount Raised Towards Match Pledge</td>
				<td valign="middle" class="formMain">$
					<input type="text"  class="formMain"  size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBMAmountRaisedTowardsMatchPledge") %><% Else %>0<% End If %>" name="frmSBMPerformanceSBMAmountRaisedTowardsMatchPledge">				
					<input type="hidden" class="formMain"  size="5" maxlength="10" value="0" name="frmSBMPerformanceSBMVolunteersInEnrollmentProcess">														
				</td>			
			</tr>	

		<tr>
			<td colspan="2" class="formHeader">
				<input type="button" value="Save Form" class="formMainBold" onclick="validateForm(); return false;">
			</td>
		</tr>

		<tr>
			<td colspan="2"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
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

