<% 
If Request("status") = "addNew" Then

	' Check for duplicate records
	
	Set DupCon = Server.CreateObject("ADODB.Connection")
	DupCon.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT Count(*) As NumberOfEntries FROM tbl_frmGeneralInformation WHERE AgencyID = '" & Request("AgencyIDN") & "' and Year = " & Request("Year")	
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
		RST.Open "SELECT * FROM tbl_frmGeneralInformation", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = Request("AgencyIDN")
		RST("Year") = Request("year")
		RST("PopulationSCA") = Request("frmGeneralInformationPopulationSCA")
		RST("SchoolAgeSCA") = Request("frmGeneralInformationSchoolAgeSCA")
		RST("VolunteerInquiries") = Request("frmGeneralInformationVolunteerInquiries")
		' RST("VolunteerApplications") = Request("frmGeneralInformationVolunteerApplications")
		' RST("VolunteersAccepted") = Request("frmGeneralInformationVolunteersAccepted")
		RST("UnmatchedClientsOpen") = Request("frmGeneralInformationUnmatchedClientsOpen")
		RST("UnmatchedClientsForTheYear") = Request("frmGeneralInformationUnmatchedClientsForTheYear")
		RST("UnmatchedVolunteersOpen") = Request("frmGeneralInformationUnmatchedVolunteersOpen")
		RST("UnmatchedVolunteersForTheYear") = Request("frmGeneralInformationUnmatchedVolunteersForTheYear")
		RST("GroupVolunteersOpen") = Request("frmGeneralInformationGroupVolunteersOpen")
		RST("GroupVolunteersForTheYear") = Request("frmGeneralInformationGroupVolunteersForTheYear")
		RST("StrategicGrowthPlan") = Request("frmGeneralInformationStrategicGrowthPlan")
		RST("ChildrenBy2004") = Request("frmGeneralInformationChildrenBy2004")
		RST("SexualPreventionCurriculum") = Request("frmGeneralInformationSexualPreventionCurriculum")
		RST("TrainingMentoringOrganizations") = Request("frmGeneralInformationTrainingMentoringOrganizations")
		RST("TrainingPostMatch") = Request("frmGeneralInformationTrainingPostMatch")
		' RST("AfterSchoolMentoringProgram") = Request("frmGeneralInformationAfterSchoolMentoringProgram")
		' RST("ASMPHowManyChildren") = Request("frmGeneralInformationASMPHowManyChildren")
		RST("VolunteerInPersonInterviews") = Request("frmGeneralInformationVolunteerInPersonInterviews")
		RST("TotalVolunteersMatched") = Request("frmGeneralInformationTotalVolunteersMatched")
		RST("CreateDate") = Now
		RST.Update
		RST.Close
		Set RST = Nothing
		form = "GeneralInformation"
		modtype = "new"
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
	RST.Open "SELECT * FROM tbl_frmGeneralInformation WHERE agencyID='" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")), Con, 1, 3
	RST("PopulationSCA") = Request("frmGeneralInformationPopulationSCA")
	RST("SchoolAgeSCA") = Request("frmGeneralInformationSchoolAgeSCA")
	RST("VolunteerInquiries") = Request("frmGeneralInformationVolunteerInquiries")
	' RST("VolunteerApplications") = Request("frmGeneralInformationVolunteerApplications")
	' RST("VolunteersAccepted") = Request("frmGeneralInformationVolunteersAccepted")
	RST("UnmatchedClientsOpen") = Request("frmGeneralInformationUnmatchedClientsOpen")
	RST("UnmatchedClientsForTheYear") = Request("frmGeneralInformationUnmatchedClientsForTheYear")
	RST("UnmatchedVolunteersOpen") = Request("frmGeneralInformationUnmatchedVolunteersOpen")
	RST("UnmatchedVolunteersForTheYear") = Request("frmGeneralInformationUnmatchedVolunteersForTheYear")
	RST("GroupVolunteersOpen") = Request("frmGeneralInformationGroupVolunteersOpen")
	RST("GroupVolunteersForTheYear") = Request("frmGeneralInformationGroupVolunteersForTheYear")
	RST("StrategicGrowthPlan") = Request("frmGeneralInformationStrategicGrowthPlan")
	RST("ChildrenBy2004") = Request("frmGeneralInformationChildrenBy2004")
	RST("SexualPreventionCurriculum") = Request("frmGeneralInformationSexualPreventionCurriculum")
	RST("TrainingMentoringOrganizations") = Request("frmGeneralInformationTrainingMentoringOrganizations")
	RST("TrainingPostMatch") = Request("frmGeneralInformationTrainingPostMatch")
	' RST("AfterSchoolMentoringProgram") = Request("frmGeneralInformationAfterSchoolMentoringProgram")
	' RST("ASMPHowManyChildren") = Request("frmGeneralInformationASMPHowManyChildren")
	RST("VolunteerInPersonInterviews") = Request("frmGeneralInformationVolunteerInPersonInterviews")
	RST("TotalVolunteersMatched") = Request("frmGeneralInformationTotalVolunteersMatched")	
	jMod = RST("GeneralInformationID")
	RST.Update
	RST.Close
	Set RST = Nothing
	form = "GeneralInformation"
	modtype = "edit"
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

<!--#include file="../includes/session_stamp.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>General Information</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
	
<script language="javascript">
	<!--
		function checkForInteger(valueToCheck)
		{
			var myRegularExpression = /^[0-9]+(,[0-9]{3})*$/;  // Checks for integer with or without commas
			var replaceWhiteSpace = /\s/; // searches for any whitespace character
			var formField = valueToCheck; // passed in as parameter 1
			var newFormField = valueToCheck.replace(replaceWhiteSpace, ""); // remove any whitespace from the form entry and replace it with nothing
			var bContainsNonNumbers = myRegularExpression.test(newFormField); // check newFormField variable to see if it contains any nonnumeric character
			
			if(!bContainsNonNumbers)
			{
				alert("Please make sure you have entered a whole number.\n We cannot process letters or words."); 
			} 
		}

function changeForm()
{
	if(document.frmGeneralInformation.frmGeneralInformationASMPHowManyChildren.value != 0)
	{
		document.frmGeneralInformation.frmGeneralInformationAfterSchoolMentoringProgram[0].checked = true;
	}
}


var myRegularExpression1 = /^[0-9]+(,[0-9]{3})*$/;
		
function submitFormValidate(form)
{
	if(!(myRegularExpression1.test(form.frmGeneralInformationPopulationSCA.value)) || (form.frmGeneralInformationPopulationSCA.value == ""))
	{
		form.frmGeneralInformationPopulationSCA.focus();
		if(form.frmGeneralInformationPopulationSCA.value == "")
			alert("You must provide a value for all form fields");
		else
			alert((form.frmGeneralInformationPopulationSCA.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmGeneralInformationSchoolAgeSCA.value)) || (form.frmGeneralInformationSchoolAgeSCA.value == ""))	
	{
		form.frmGeneralInformationSchoolAgeSCA.focus();
		if(form.frmGeneralInformationSchoolAgeSCA.value == "")
			alert("You must provide a value for all form fields");
		else
			alert((form.frmGeneralInformationSchoolAgeSCA.value) + " is invalid.");
		return false;
	}

	/* 3 */
	else if(!(myRegularExpression1.test(form.frmGeneralInformationVolunteerInquiries.value)) || (form.frmGeneralInformationVolunteerInquiries.value == ""))	
	{
		form.frmGeneralInformationVolunteerInquiries.focus();
		if(form.frmGeneralInformationVolunteerInquiries.value == "")
			alert("You must provide a value for all form fields");
		else
			alert((form.frmGeneralInformationVolunteerInquiries.value) + " is invalid.");
		return false;
	}
	
	
	/* 4 */
		else if(!(myRegularExpression1.test(form.frmGeneralInformationVolunteerInPersonInterviews.value)) || (form.frmGeneralInformationVolunteerInPersonInterviews.value == ""))	
	{
		form.frmGeneralInformationVolunteerInPersonInterviews.focus();
		if(form.frmGeneralInformationVolunteerInPersonInterviews.value == "")
			alert("You must provide a value for all form fields");
		else
			alert((form.frmGeneralInformationVolunteerInPersonInterviews.value) + " is invalid.");
		return false;
	}
	
	
	/* 5 */
		else if(!(myRegularExpression1.test(form.frmGeneralInformationTotalVolunteersMatched.value)) || (form.frmGeneralInformationTotalVolunteersMatched.value == ""))	
	{
		form.frmGeneralInformationTotalVolunteersMatched.focus();
		if(form.frmGeneralInformationTotalVolunteersMatched.value == "")
			alert("You must provide a value for all form fields");
		else
			alert((form.frmGeneralInformationTotalVolunteersMatched.value) + " is invalid.");
		return false;
	}	


/*	else if(!(myRegularExpression1.test(form.frmGeneralInformationVolunteerApplications.value)) || (form.frmGeneralInformationVolunteerApplications.value == ""))	
	{
		form.frmGeneralInformationVolunteerApplications.focus();
		if(form.frmGeneralInformationVolunteerApplications.value == "")
			alert("You must provide a value for all form fields");
		else
			alert((form.frmGeneralInformationVolunteerApplications.value) + " is invalid.");
		return false;
	}
	
*/

/*	else if(!(myRegularExpression1.test(form.frmGeneralInformationVolunteersAccepted.value)) || (form.frmGeneralInformationVolunteersAccepted.value == ""))	
	{
		form.frmGeneralInformationVolunteersAccepted.focus();
		if(form.frmGeneralInformationVolunteersAccepted.value == "")
			alert("You must provide a value for all form fields");
		else
			alert((form.frmGeneralInformationVolunteersAccepted.value) + " is invalid.");
		return false;
	}
*/
	/* 6 */
	else if(!(myRegularExpression1.test(form.frmGeneralInformationUnmatchedClientsOpen.value)) || (form.frmGeneralInformationUnmatchedClientsOpen.value == ""))	
	{
		form.frmGeneralInformationUnmatchedClientsOpen.focus();
		if(form.frmGeneralInformationUnmatchedClientsOpen.value == "")
			alert("You must provide a value for all form fields");
		else
			alert((form.frmGeneralInformationUnmatchedClientsOpen.value) + " is invalid.");
		return false;
	}
	/* 7 */
	else if(!(myRegularExpression1.test(form.frmGeneralInformationUnmatchedClientsForTheYear.value)) || (form.frmGeneralInformationUnmatchedClientsForTheYear.value == ""))	
	{
		form.frmGeneralInformationUnmatchedClientsForTheYear.focus();
		if(form.frmGeneralInformationUnmatchedClientsForTheYear.value == "")
			alert("You must provide a value for all form fields");
		else
			alert((form.frmGeneralInformationUnmatchedClientsForTheYear.value) + " is invalid.");
		return false;
	}
	/* 8 */
	else if(!(myRegularExpression1.test(form.frmGeneralInformationUnmatchedVolunteersOpen.value)) || (form.frmGeneralInformationUnmatchedVolunteersOpen.value == ""))	
	{
		form.frmGeneralInformationUnmatchedVolunteersOpen.focus();
		if(form.frmGeneralInformationUnmatchedVolunteersOpen.value == "")
			alert("You must provide a value for all form fields");
		else
			alert((form.frmGeneralInformationUnmatchedVolunteersOpen.value) + " is invalid.");
		return false;
	}
	/* 9 */
	else if(!(myRegularExpression1.test(form.frmGeneralInformationUnmatchedVolunteersForTheYear.value)) || (form.frmGeneralInformationUnmatchedVolunteersForTheYear.value == ""))	
	{
		form.frmGeneralInformationUnmatchedVolunteersForTheYear.focus();
		if(form.frmGeneralInformationUnmatchedVolunteersForTheYear.value == "")
			alert("You must provide a value for all form fields");
		else
			alert((form.frmGeneralInformationUnmatchedVolunteersForTheYear.value) + " is invalid.");
		return false;
	}
	/* 10 */
	else if(!(myRegularExpression1.test(form.frmGeneralInformationGroupVolunteersOpen.value)) || (form.frmGeneralInformationGroupVolunteersOpen.value == ""))	
	{
		form.frmGeneralInformationGroupVolunteersOpen.focus();
		if(form.frmGeneralInformationGroupVolunteersOpen.value == "")
			alert("You must provide a value for all form fields");
		else
			alert((form.frmGeneralInformationGroupVolunteersOpen.value) + " is invalid.");
		return false;
	}
	/* 11 */
	else if(!(myRegularExpression1.test(form.frmGeneralInformationGroupVolunteersForTheYear.value)) || (form.frmGeneralInformationGroupVolunteersForTheYear.value == ""))	
	{
		form.frmGeneralInformationGroupVolunteersForTheYear.focus();
		if(form.frmGeneralInformationGroupVolunteersForTheYear.value == "")
			alert("You must provide a value for all form fields");
		else
			alert((form.frmGeneralInformationGroupVolunteersForTheYear.value) + " is invalid.");
		return false;
	}
	else if((form.frmGeneralInformationStrategicGrowthPlan[0].checked != true) && (form.frmGeneralInformationStrategicGrowthPlan[1].checked != true))
	{
		alert("Please make sure that you have selected the appropriate answer for question 6.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmGeneralInformationChildrenBy2004.value)) || (form.frmGeneralInformationChildrenBy2004.value == ""))	
	{
		form.frmGeneralInformationChildrenBy2004.focus();
		if(form.frmGeneralInformationChildrenBy2004.value == "")
			alert("You must provide a value for all form fields");
		else
			alert((form.frmGeneralInformationChildrenBy2004.value) + " is invalid.");
		return false;
	}
	else if((form.frmGeneralInformationSexualPreventionCurriculum[0].checked != true) && (form.frmGeneralInformationSexualPreventionCurriculum[1].checked != true))
	{
		alert("Please make sure that you have selected the appropriate answer for question 8.");
		return false;
	}
	else if((form.frmGeneralInformationTrainingMentoringOrganizations[0].checked != true) && (form.frmGeneralInformationTrainingMentoringOrganizations[1].checked != true))
	{
		alert("Please make sure that you have selected the appropriate answer for question 9.");
		return false;
	}
	else if((form.frmGeneralInformationTrainingPostMatch[0].checked != true) && (form.frmGeneralInformationTrainingPostMatch[1].checked != true))
	{
		alert("Please make sure that you have selected the appropriate answer for question 10.");
		return false;
	}

/*	else if((form.frmGeneralInformationAfterSchoolMentoringProgram[0].checked != true) && (form.frmGeneralInformationAfterSchoolMentoringProgram[1].checked != true))
	{
		alert("Please make sure that you have selected the appropriate answer for question 17.");
		return false;
	}


	else if(!(myRegularExpression1.test(form.frmGeneralInformationASMPHowManyChildren.value)) || (form.frmGeneralInformationASMPHowManyChildren.value == ""))	
	{
		form.frmGeneralInformationASMPHowManyChildren.focus();
		if(form.frmGeneralInformationASMPHowManyChildren.value == "")
			alert("You must provide a value for all form fields");
		else
			alert((form.frmGeneralInformationASMPHowManyChildren.value) + " is invalid.");
		return false;
	}
	else
	{
		return true;
	}
*/

}	
	// -->
</script>

<% '<!--#include file="../includes/top_nav_forms_yearly.inc"--><!-- include file has </head> and <body> tags --><br>      %>
<!--#include file="../includes/surveytitle.inc"-->

<table width=100% cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_fishing.jpg" alt="" width="220" height="477" border="0"></td>
<td valign="top">

<% If say = "thanks" Then %>

<font class="formMain">
<br><br>
<strong>Thank you!</strong> Your information has been saved in the BBBS database.<br>
To choose another form, please select the form type from the choices above.
</font>

<br>
<!--#include file="../includes/contact_info.inc"-->
<br>



<% ElseIf say <> "thanks" Then  %>
<form name="frmGeneralInformation" action="generalinformation_edit.asp" method="post" onsubmit="return submitFormValidate(this);">
<!--#include file="../includes/form_stamp.asp"-->

<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmGeneralInformation WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
	Set GetGeneralInformation = Con.Execute(query)
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
		<p class="formMain">We're sorry, but this form was previously completed. To make changes please <a href="yearly.asp">reselect</a> the 
		appropriate form and year and update the existing information.</p>
		<%
		Response.End
		End If 
		%>
			<br>
			<table border="1" cellspacing="0" cellpadding="3"  bordercolordark="#003063" width = "550">
				<tr> 
					<td colspan="3" align="center" valign="top" class="formSubhead">BBBS - <%= y %> Annual Agency Information (AAI)</td>
				</tr>
				<tr>
					<td colspan="3" class="formHeader">GENERAL INFORMATION</td>
				</tr>
				<tr>
					<td colspan="7" class="formMain"><font color="#ff0000"><div align="center"><strong>Please Note: </strong>After entering your information, you <strong>must</strong> click on the "Save Form" button at the bottom of the form and wait for the "Thank You" screen or your changes will be lost.</div></font></td>
				</tr>	
				
<!-- Question Number 1 -->
				<tr>
					<td align="right" valign="top" class="formMain">1.</td>
					<td align="left" valign="top" class="formMain">Population of your Service Community Area (SCA):<br>
					<i>Call your local library or visit <a href="http://www.census.gov" target="_blank">www.census.gov</a> for current Census figures</i></td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% If say = "edit" Then %><%=  GetGeneralInformation("PopulationSCA") %><% Else %>0<% End If %>" name="frmGeneralInformationPopulationSCA" onchange="checkForInteger(this.value);"></td>
				</tr>
				
<!-- Question Number 2 -->
				<tr>
					<td align="right" valign="top" class="formMain">2.</td>
					<td align="left" valign="top" class="formMain">Number of school age children (K-12) in SCA:<br>
					<i>Call your County Superintendent's Office</i></td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% If say = "edit" Then %><%= GetGeneralInformation("SchoolAgeSCA") %><% Else %>0<% End If %>" name="frmGeneralInformationSchoolAgeSCA" onchange="checkForInteger(this.value);"></td>
				</tr>
<!-- Question Number 3 -->
				<tr> 
					<td align="right" valign="top" class="formMain">3.</td>
					<td align="left" valign="top" class="formMain">What is the total number of volunteer inquiries you received?
					<br><em>A volunteer is considered to have inquired when he/she contacts the agency, expresses an interest in being a Big and provides basic contact information.  Contact includes web-based inquiries.</em>
					</td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% If say = "edit" Then %><%= GetGeneralInformation("VolunteerInquiries") %><% Else %>0<% End If %>" name="frmGeneralInformationVolunteerInquiries" onchange="checkForInteger(this.value);"></td>
				</tr>

<!-- Question Number 4 -->
				<tr>
					<td align="right" valign="top" class="formMain">4.</td>
					<td align="left" valign="top" class="formMain">What is the total number of volunteer in person interviews?</td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% If say = "edit" Then %><%= GetGeneralInformation("VolunteerInpersonInterviews") %><% Else %>0<% End If %>" name="frmGeneralInformationVolunteerInPersonInterviews" onchange="checkForInteger(this.value);"></td>
				</tr>
				

<!-- Question Number 5 -->
				<tr>
					<td align="right" valign="top" class="formMain">5.</td>
					<td align="left" valign="top" class="formMain">What is the total number of volunteers that were matched?</td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% If say = "edit" Then %><%= GetGeneralInformation("TotalVolunteersMatched") %><% Else %>0<% End If %>" name="frmGeneralInformationTotalVolunteersMatched" onchange="checkForInteger(this.value);"></td>
				</tr>				
				
				
<!-- Question Number 6 NOT USED
				<tr>
					<td align="right" valign="top" class="formMain">6.</td>
					<td align="left" valign="top" class="formMain">What is the total number of volunteer applications you received?</td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% ' If say = "edit" Then %><% ' = GetGeneralInformation("VolunteerApplications") %><% ' Else %>0<% ' End If %>" name="frmGeneralInformationVolunteerApplications" onchange="checkForInteger(this.value);"></td>
				</tr>
-->


<!-- Question Number 6 NOT USED
				<tr>
					<td align="right" valign="top" class="formMain">6.</td>
					<td align="left" valign="top" class="formMain">What is the total number of volunteers that were accepted?<br>
					<i>Reached the status of&nbsp;&nbsp;</i>&quot;<b>Ready to be Matched</b>&quot;</td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<%  ' If say = "edit" Then %><% '= GetGeneralInformation("VolunteersAccepted") %><% 'Else %>0<% 'End If %>" name="frmGeneralInformationVolunteersAccepted" onchange="checkForInteger(this.value);"></td>
				</tr> -->

<!-- Question Number 6 -->
				<tr>
					<td align="right" valign="top" class="formMain">6.</td>
					<td align="left" valign="top" class="formMain">Do you have a Strategic Growth Plan in place?</td>
					<td align="right" valign="top" class="formMain">Yes<input type="radio" value="1" class="formMain" name="frmGeneralInformationStrategicGrowthPlan"<% If (say = "edit") Then %><% If (GetGeneralInformation("StrategicGrowthPlan") = True) Then %> checked<% End If %><% End If %>>&nbsp;No<input type="radio" class="formMain" value="0" name="frmGeneralInformationStrategicGrowthPlan"<% If (say = "edit") Then %><% If (GetGeneralInformation("StrategicGrowthPlan") = False) Then %> checked<% End If %><% End If %>></td>
				</tr>
<!-- Question Number 7 -->
				<tr>
					<td align="right" valign="top" class="formMain">7.</td>
					<td align="left" valign="top" class="formMain">According to the Strategic Growth Plan, how many children do you plan to serve by 2004?</td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% If say = "edit" Then %><%= GetGeneralInformation("ChildrenBy2004") %><% Else %>0<% End If %>" name="frmGeneralInformationChildrenBy2004" onchange="checkForInteger(this.value);"></td>
				</tr>
<!-- Question Number 8 -->
				<tr>
					<td align="right" valign="top" class="formMain">8.</td>
					<td align="left" valign="top" class="formMain">Do you use EMPOWER or similar sexual-prevention curriculum?</td>
					<td align="right" valign="top" class="formMain">Yes<input type="radio" value="1" class="formMain" name="frmGeneralInformationSexualPreventionCurriculum"<% If (say = "edit") Then %><% If (GetGeneralInformation("SexualPreventionCurriculum") = True) Then %> checked<% End If %><% End If %>>&nbsp;No<input type="radio" class="formMain" value="0" name="frmGeneralInformationSexualPreventionCurriculum"<% If (say = "edit") Then %><% If (GetGeneralInformation("SexualPreventionCurriculum") = False) Then %> checked<% End If %><% End If %>></td>
				</tr>
<!-- Question Number 9 --> 
				<tr>
					<td align="right" valign="top" class="formMain">9.</td>
					<td align="left" valign="top" class="formMain">Do you provide training for other mentoring organizations?</td>
					<td align="right" valign="top" class="formMain">Yes<input type="radio" value="1" class="formMain" name="frmGeneralInformationTrainingMentoringOrganizations"<% If (say = "edit") Then %><% If (GetGeneralInformation("TrainingMentoringOrganizations") = True) Then %> checked<% End If %><% End If %>>&nbsp;No<input type="radio" class="formMain" value="0" name="frmGeneralInformationTrainingMentoringOrganizations"<% If (say = "edit") Then %><% If (GetGeneralInformation("TrainingMentoringOrganizations") = False) Then %> checked<% End If %><% End If %>></td>
				</tr>
<!-- Question Number 10 -->
				<tr>
					<td align="right" valign="top" class="formMain">10.</td>
					<td align="left" valign="top" class="formMain">Do you provide post-match training for your volunteers?</td>
					<td align="right" valign="top" class="formMain">Yes<input type="radio" value="1" class="formMain" name="frmGeneralInformationTrainingPostMatch"<% If (say = "edit") Then %><% If (GetGeneralInformation("TrainingPostMatch") = True) Then %> checked<% End If %><% End If %>>&nbsp;No<input type="radio" class="formMain" value="0" name="frmGeneralInformationTrainingPostMatch"<% If (say = "edit") Then %><% If (GetGeneralInformation("TrainingPostMatch") = False) Then %> checked<% End If %><% End If %>></td>
				</tr>

<!-- Question Number 17 NOT USED
				<tr>
					<td rowspan="2" align="right" valign="top" class="formMain">17.</td>
					<td align="left" valign="top" class="formMain">Do you have an After School Mentoring Program?</td>
					<td align="right" valign="top" class="formMain">Yes<input type="radio" value="1" class="formMain" name="frmGeneralInformationAfterSchoolMentoringProgram"<% ' If (say = "edit") Then %><% ' If (GetGeneralInformation("AfterSchoolMentoringProgram") = True) Then %> checked<% ' End If %><% ' End If %>>&nbsp;No<input value="0" type="radio" class="formMain" name="frmGeneralInformationAfterSchoolMentoringProgram"<% ' If (say = "edit") Then %><% ' If (GetGeneralInformation("AfterSchoolMentoringProgram") = False) Then %> checked<% ' End If %><% ' End If %> onclick="form.frmGeneralInformationASMPHowManyChildren.value=0"></td>
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain">If yes, how many children do you serve?</td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% ' If say = "edit" Then %><%' = GetGeneralInformation("ASMPHowManyChildren") %><% ' Else %>0<% ' End If %>" name="frmGeneralInformationASMPHowManyChildren" onblur="checkForInteger(this.value);changeForm()"></td>
				</tr> -->
				
				<tr>
				<td colspan="3" align="center" class="formMain"><em>Below please list all <strong>RTBM</strong> Clients and Volunteers, Open and Closed.<br>Categories are <strong>mutually exclusive</strong></em></td>
				</tr>

				
<!-- Question Number 11 -->
				<tr>
					<td align="right" valign="top" class="formMain">11.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Unmatched Clients (RTBM)<br>
					OPEN as of 12/31/<%= y %></td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% If say = "edit" Then %><%= GetGeneralInformation("UnmatchedClientsOpen") %><% Else %>0<% End If %>" name="frmGeneralInformationUnmatchedClientsOpen" onchange="checkForInteger(this.value);"></td>
				</tr>
<!-- Question Number 12 -->
				<tr>
					<td align="right" valign="top" class="formMain">12.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Unmatched Clients (RTBM)<br>
					CLOSED between 1/1/<%= y %>-12/31/<%= y %></td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% If say = "edit" Then %><%= GetGeneralInformation("UnmatchedClientsForTheYear") %><% Else %>0<% End If %>" name="frmGeneralInformationUnmatchedClientsForTheYear" onchange="checkForInteger(this.value);"></td>
				</tr>
<!-- Question Number 13 -->
				<tr>
					<td align="right" valign="top" class="formMain">13.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Unmatched Volunteers<br>
					OPEN as of 12/31/<%= y %></td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% If say = "edit" Then %><%= GetGeneralInformation("UnmatchedVolunteersOpen") %><% Else %>0<% End If %>" name="frmGeneralInformationUnmatchedVolunteersOpen" onchange="checkForInteger(this.value);"></td>
				</tr>
<!-- Question Number 14 -->
				<tr>
					<td align="right" valign="top" class="formMain">14.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Unmatched Volunteers<br>
					CLOSED between 1/1/<%= y %>-12/31/<%= y %></td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% If say = "edit" Then %><%= GetGeneralInformation("UnmatchedVolunteersForTheYear") %><% Else %>0<% End If %>" name="frmGeneralInformationUnmatchedVolunteersForTheYear" onchange="checkForInteger(this.value);"></td>
				</tr>
<!-- Question Number 15 -->
				<tr>
					<td align="right" valign="top" class="formMain">15.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Group Volunteers<br>
					OPEN as of 12/31/<%= y %></td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% If say = "edit" Then %><%= GetGeneralInformation("GroupVolunteersOpen") %><% Else %>0<% End If %>" name="frmGeneralInformationGroupVolunteersOpen" onchange="checkForInteger(this.value);"></td>
				</tr>
<!-- Question Number 16 -->
				<tr>
					<td align="right" valign="top" class="formMain">16.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Group Volunteers<br>
					CLOSED between 1/1/<%= y %>-12/31/<%= y %></td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% If say = "edit" Then %><%= GetGeneralInformation("GroupVolunteersForTheYear") %><% Else %>0<% End If %>" name="frmGeneralInformationGroupVolunteersForTheYear" onchange="checkForInteger(this.value);"></td>
				</tr>				
				
				
				<tr>
					<td colspan="3" class="formHeader"><input type="submit" value="Save Form" class="formMainBold"></td>
				</tr>
				<tr>
				<td colspan="3" align="center">
				<!--#include file="../includes/contact_info.inc"-->
				</td>
				</tr>
				</table>

			
				
				<br>

<br>
					
<% 
If say = "edit" Then
	GetGeneralInformation.Close
	Set GetGeneralInformation = Nothing
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

