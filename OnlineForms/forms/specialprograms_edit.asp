<% 
If Request("status") = "addNew" Then

' Check for duplicate records
	
	Set DupCon = Server.CreateObject("ADODB.Connection")
	DupCon.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT Count(*) As NumberOfEntries FROM tbl_frmSpecialPrograms WHERE AgencyID = '" & Request("AgencyIDN") & "' and Year = " & Request("Year")	
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
		RST.Open "SELECT * FROM tbl_frmSpecialPrograms", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = Request("AgencyIDN")
		RST("Year") = Request("year")
		RST("FiftyFiveOrOlderBigs") = Request("frmSpecialProgramsFiftyFiveOrOlderBigs")
		RST("AcademicAchievement") = Request("frmSpecialProgramsAcademicAchievement")
		RST("AfterSchoolMentoring") = Request("frmSpecialProgramsAfterSchoolMentoring")
		RST("AIDSPreventionIntervention") = Request("frmSpecialProgramsAIDSPreventionIntervention")
		RST("AlcoholAbusePreventionIntervention") = Request("frmSpecialProgramsAlcoholAbusePreventionIntervention")
		RST("Camping") = Request("frmSpecialProgramsCamping")
		RST("CharacterCounts") = Request("frmSpecialProgramsCharacterCounts")
		RST("ChildrenWithDisabilities") = Request("frmSpecialProgramsChildrenWithDisabilities")
		RST("CollegeStudentsAsBigs") = Request("frmSpecialProgramsCollegeStudentsAsBigs")
		RST("CommunityServiceProjects") = Request("frmSpecialProgramsCommunityServiceProjects")
		RST("DropOutPrevention") = Request("frmSpecialProgramsDropOutPrevention")
		RST("DrugAbusePreventionIntervention") = Request("frmSpecialProgramsDrugAbusePreventionIntervention")
		RST("EmergencyFinancialAssistance") = Request("frmSpecialProgramsEmergencyFinancialAssistance")
		RST("EmployabilityJobReadiness") = Request("frmSpecialProgramsEmployabilityJobReadiness")
		RST("FamilyCounseling") = Request("frmSpecialProgramsFamilyCounseling")
		RST("GroupActivities") = Request("frmSpecialProgramsGroupActivities")
		RST("HighSchoolStudentsAsBigs") = Request("frmSpecialProgramsHighSchoolStudentsAsBigs")
		RST("LifeSkillsLifeChoices") = Request("frmSpecialProgramsLifeSkillsLifeChoices")
		RST("ParentSupportGroups") = Request("frmSpecialProgramsParentSupportGroups")
		RST("PartnershipsCivicOrganizations") = Request("frmSpecialProgramsPartnershipsCivicOrganizations")
		RST("PartnershipsCollegesUniversities") = Request("frmSpecialProgramsPartnershipsCollegesUniversities")
		RST("PartnershipsCorporationsBusinesses") = Request("frmSpecialProgramsPartnershipsCorporationsBusinesses")
		RST("PartnershipsOtherYouthServingOrganizations") = Request("frmSpecialProgramsPartnershipsOtherYouthServingOrganizations")
		RST("PartnershipsReligiousOrganizations") = Request("frmSpecialProgramsPartnershipsReligiousOrganizations")
		RST("PregnancyPrevention") = Request("frmSpecialProgramsPregnancyPrevention")
		RST("Scholarships") = Request("frmSpecialProgramsScholarships")
		RST("SexualAbusePreventionInterventionEmpower") = Request("frmSpecialProgramsSexualAbusePreventionInterventionEmpower")
		RST("SexualAbusePreventionInterventionNOTEmpower") = Request("frmSpecialProgramsSexualAbusePreventionInterventionNOTEmpower")
		RST("SiteBasedMentoringSchoolBased") = Request("frmSpecialProgramsSiteBasedMentoringSchoolBased")
		RST("SiteBasedMentoringNOTSchoolBased") = Request("frmSpecialProgramsSiteBasedMentoringNOTSchoolBased")
		RST("TeenParenting") = Request("frmSpecialProgramsTeenParenting")
		RST("OtherText") = Request("frmSpecialProgramsOtherText")
		RST("Other") = Request("frmSpecialProgramsOther")
		RST("CreateDate") = Now
		RST.Update
		RST.Close
		Set RST = Nothing
		form = "SpecialPrograms"
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
	RST.Open "SELECT * FROM tbl_frmSpecialPrograms WHERE agencyID='" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")), Con, 1, 3
	RST("FiftyFiveOrOlderBigs") = Trim(Request("frmSpecialProgramsFiftyFiveOrOlderBigs"))
	RST("AcademicAchievement") = Trim(Request("frmSpecialProgramsAcademicAchievement"))
	RST("AfterSchoolMentoring") = Trim(Request("frmSpecialProgramsAfterSchoolMentoring"))
	RST("AIDSPreventionIntervention") = Trim(Request("frmSpecialProgramsAIDSPreventionIntervention"))
	RST("AlcoholAbusePreventionIntervention") = Trim(Request("frmSpecialProgramsAlcoholAbusePreventionIntervention"))
	RST("Camping") = Trim(Request("frmSpecialProgramsCamping"))
	RST("CharacterCounts") = Trim(Request("frmSpecialProgramsCharacterCounts"))
	RST("ChildrenWithDisabilities") = Trim(Request("frmSpecialProgramsChildrenWithDisabilities"))
	RST("CollegeStudentsAsBigs") = Trim(Request("frmSpecialProgramsCollegeStudentsAsBigs"))
	RST("CommunityServiceProjects") = Trim(Request("frmSpecialProgramsCommunityServiceProjects"))
	RST("DropOutPrevention") = Trim(Request("frmSpecialProgramsDropOutPrevention"))
	RST("DrugAbusePreventionIntervention") = Trim(Request("frmSpecialProgramsDrugAbusePreventionIntervention"))
	RST("EmergencyFinancialAssistance") = Trim(Request("frmSpecialProgramsEmergencyFinancialAssistance"))
	RST("EmployabilityJobReadiness") = Trim(Request("frmSpecialProgramsEmployabilityJobReadiness"))
	RST("FamilyCounseling") = Trim(Request("frmSpecialProgramsFamilyCounseling"))
	RST("GroupActivities") = Trim(Request("frmSpecialProgramsGroupActivities"))
	RST("HighSchoolStudentsAsBigs") = Trim(Request("frmSpecialProgramsHighSchoolStudentsAsBigs"))
	RST("LifeSkillsLifeChoices") = Trim(Request("frmSpecialProgramsLifeSkillsLifeChoices"))
	RST("ParentSupportGroups") = Trim(Request("frmSpecialProgramsParentSupportGroups"))
	RST("PartnershipsCivicOrganizations") = Trim(Request("frmSpecialProgramsPartnershipsCivicOrganizations"))
	RST("PartnershipsCollegesUniversities") = Trim(Request("frmSpecialProgramsPartnershipsCollegesUniversities"))
	RST("PartnershipsCorporationsBusinesses") = Trim(Request("frmSpecialProgramsPartnershipsCorporationsBusinesses"))
	RST("PartnershipsOtherYouthServingOrganizations") = Trim(Request("frmSpecialProgramsPartnershipsOtherYouthServingOrganizations"))
	RST("PartnershipsReligiousOrganizations") = Trim(Request("frmSpecialProgramsPartnershipsReligiousOrganizations"))
	RST("PregnancyPrevention") = Trim(Request("frmSpecialProgramsPregnancyPrevention"))
	RST("Scholarships") = Trim(Request("frmSpecialProgramsScholarships"))
	RST("SexualAbusePreventionInterventionEmpower") = Trim(Request("frmSpecialProgramsSexualAbusePreventionInterventionEmpower"))
	RST("SexualAbusePreventionInterventionNOTEmpower") = Trim(Request("frmSpecialProgramsSexualAbusePreventionInterventionNOTEmpower"))
	RST("SiteBasedMentoringSchoolBased") = Trim(Request("frmSpecialProgramsSiteBasedMentoringSchoolBased"))
	RST("SiteBasedMentoringNOTSchoolBased") = Trim(Request("frmSpecialProgramsSiteBasedMentoringNOTSchoolBased"))
	RST("TeenParenting") = Trim(Request("frmSpecialProgramsTeenParenting"))
	RST("OtherText") = Trim(Request("frmSpecialProgramsOtherText"))
	RST("Other") = Trim(Request("frmSpecialProgramsOther"))
	jMod = RST("SpecialProgramsID")
	RST.Update
	RST.Close
	Set RST = Nothing
	form = "SpecialPrograms"
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
	<title>Special Programs</title>
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



var myRegularExpression1 = /^[0-9]+(,[0-9]{3})*$/;  // Checks for integer with or without commas
	
function submitFormValidate(form)
{
	for(i = 0; i < form.length; i ++)
	{
		if(form.elements[i].value == "")
		{
			form.elements[i].focus();
			alert("You must provide a value for all form fields");
			return false;
		}
	}

	if(!(myRegularExpression1.test(form.frmSpecialProgramsFiftyFiveOrOlderBigs.value)))
	{
		form.frmSpecialProgramsFiftyFiveOrOlderBigs.focus();
		alert((form.frmSpecialProgramsFiftyFiveOrOlderBigs.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsGroupActivities.value)))	
	{
		form.frmSpecialProgramsGroupActivities.focus();
		alert((form.frmSpecialProgramsGroupActivities.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsAcademicAchievement.value)))	
	{
		form.frmSpecialProgramsAcademicAchievement.focus();
		alert((form.frmSpecialProgramsAcademicAchievement.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsHighSchoolStudentsAsBigs.value)))	
	{
		form.frmSpecialProgramsHighSchoolStudentsAsBigs.focus();
		alert((form.frmSpecialProgramsHighSchoolStudentsAsBigs.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsAfterSchoolMentoring.value)))	
	{
		form.frmSpecialProgramsAfterSchoolMentoring.focus();
		alert((form.frmSpecialProgramsAfterSchoolMentoring.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsLifeSkillsLifeChoices.value)))	
	{
		form.frmSpecialProgramsLifeSkillsLifeChoices.focus();
		alert((form.frmSpecialProgramsLifeSkillsLifeChoices.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsAIDSPreventionIntervention.value)))	
	{
		form.frmSpecialProgramsAIDSPreventionIntervention.focus();
		alert((form.frmSpecialProgramsAIDSPreventionIntervention.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsParentSupportGroups.value)))	
	{
		form.frmSpecialProgramsParentSupportGroups.focus();
		alert((form.frmSpecialProgramsParentSupportGroups.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsAlcoholAbusePreventionIntervention.value)))	
	{
		form.frmSpecialProgramsAlcoholAbusePreventionIntervention.focus();
		alert((form.frmSpecialProgramsAlcoholAbusePreventionIntervention.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsPartnershipsCivicOrganizations.value)))	
	{
		form.frmSpecialProgramsPartnershipsCivicOrganizations.focus();
		alert((form.frmSpecialProgramsPartnershipsCivicOrganizations.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsCamping.value)))	
	{
		form.frmSpecialProgramsCamping.focus();
		alert((form.frmSpecialProgramsCamping.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsPartnershipsCollegesUniversities.value)))	
	{
		form.frmSpecialProgramsPartnershipsCollegesUniversities.focus();
		alert((form.frmSpecialProgramsPartnershipsCollegesUniversities.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsCharacterCounts.value)))	
	{
		form.frmSpecialProgramsCharacterCounts.focus();
		alert((form.frmSpecialProgramsCharacterCounts.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsPartnershipsCorporationsBusinesses.value)))	
	{
		form.frmSpecialProgramsPartnershipsCorporationsBusinesses.focus();
		alert((form.frmSpecialProgramsPartnershipsCorporationsBusinesses.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsChildrenWithDisabilities.value)))	
	{
		form.frmSpecialProgramsChildrenWithDisabilities.focus();
		alert((form.frmSpecialProgramsChildrenWithDisabilities.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsPartnershipsOtherYouthServingOrganizations.value)))	
	{
		form.frmSpecialProgramsPartnershipsOtherYouthServingOrganizations.focus();
		alert((form.frmSpecialProgramsPartnershipsOtherYouthServingOrganizations.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsCollegeStudentsAsBigs.value)))	
	{
		form.frmSpecialProgramsCollegeStudentsAsBigs.focus();
		alert((form.frmSpecialProgramsCollegeStudentsAsBigs.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsPartnershipsReligiousOrganizations.value)))	
	{
		form.frmSpecialProgramsPartnershipsReligiousOrganizations.focus();
		alert((form.frmSpecialProgramsPartnershipsReligiousOrganizations.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsCommunityServiceProjects.value)))	
	{
		form.frmSpecialProgramsCommunityServiceProjects.focus();
		alert((form.frmSpecialProgramsCommunityServiceProjects.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsPregnancyPrevention.value)))	
	{
		form.frmSpecialProgramsPregnancyPrevention.focus();
		alert((form.frmSpecialProgramsPregnancyPrevention.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsDropOutPrevention.value)))	
	{
		form.frmSpecialProgramsDropOutPrevention.focus();
		alert((form.frmSpecialProgramsDropOutPrevention.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsScholarships.value)))	
	{
		form.frmSpecialProgramsScholarships.focus();
		alert((form.frmSpecialProgramsScholarships.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsDrugAbusePreventionIntervention.value)))	
	{
		form.frmSpecialProgramsDrugAbusePreventionIntervention.focus();
		alert((form.frmSpecialProgramsDrugAbusePreventionIntervention.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsSexualAbusePreventionInterventionEmpower.value)))	
	{
		form.frmSpecialProgramsSexualAbusePreventionInterventionEmpower.focus();
		alert((form.frmSpecialProgramsSexualAbusePreventionInterventionEmpower.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsSexualAbusePreventionInterventionNOTEmpower.value)))	
	{
		form.frmSpecialProgramsSexualAbusePreventionInterventionNOTEmpower.focus();
		alert((form.frmSpecialProgramsSexualAbusePreventionInterventionNOTEmpower.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsEmergencyFinancialAssistance.value)))	
	{
		form.frmSpecialProgramsEmergencyFinancialAssistance.focus();
		alert((form.frmSpecialProgramsEmergencyFinancialAssistance.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsSiteBasedMentoringSchoolBased.value)))	
	{
		form.frmSpecialProgramsSiteBasedMentoringSchoolBased.focus();
		alert((form.frmSpecialProgramsSiteBasedMentoringSchoolBased.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsSiteBasedMentoringNOTSchoolBased.value)))	
	{
		form.frmSpecialProgramsSiteBasedMentoringNOTSchoolBased.focus();
		alert((form.frmSpecialProgramsSiteBasedMentoringNOTSchoolBased.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsEmployabilityJobReadiness.value)))	
	{
		form.frmSpecialProgramsEmployabilityJobReadiness.focus();
		alert((form.frmSpecialProgramsEmployabilityJobReadiness.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsTeenParenting.value)))	
	{
		form.frmSpecialProgramsTeenParenting.focus();
		alert((form.frmSpecialProgramsTeenParenting.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsFamilyCounseling.value)))	
	{
		form.frmSpecialProgramsFamilyCounseling.focus();
		alert((form.frmSpecialProgramsFamilyCounseling.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialProgramsOther.value)))	
	{
		form.frmSpecialProgramsOther.focus();
		alert((form.frmSpecialProgramsOther.value) + " is invalid.");
		return false;
	}
	else if((form.frmSpecialProgramsOther.value != "0") && (form.frmSpecialProgramsOther.value != "") && (form.frmSpecialProgramsOther.value != "000"))
	{
		if((form.frmSpecialProgramsOtherText.value == "(Name Program)") || (form.frmSpecialProgramsOtherText.value == ""))
			{
			form.frmSpecialProgramsOtherText.focus();
			alert("Please Name Program");
			return false;
			}
		else
			{
			return true;
			}
	}
	else
	{
		return true;
	}
}
	// -->
</script>

<% '<!--#include file="../includes/top_nav_forms_yearly.inc"--><!-- include file has </head> and <body> tags --><br>     %>
<!--#include file="../includes/surveytitle.inc"-->

<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_football.jpg" alt="" width="220" height="477" border="0"></td>
<td width="100%" valign="top">
	
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

</div>
</center>

<% ElseIf say <> "thanks" Then  %>
<br>
<table width="640" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063">
<form name="frmSpecialPrograms" action="specialprograms_edit.asp" method="post" onsubmit="return submitFormValidate(this)">

<!--#include file="../includes/form_stamp.asp"-->
<center>
<% 
If say = "edit" Then

	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmSpecialPrograms WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("year"))
	Set GetSpecialPrograms = Con.Execute(query)
 %>
<input type="hidden" name="status" value="editSave">
<% Else %>
<input type="hidden" name="status" value="addNew">
<%
End If
 %>
<div align="center">
<center>

<%
If say = "previouslyEdited" Then
%>
<p class="formMain">We're sorry, but this form was previously completed. To make changes please <a href="yearly.asp">reselect</a> the 
appropriate form and year and update the existing information.</p>
<%
Response.End
End If 
%>

		<tr>
			<td colspan="4" align="center" valign="top" class="formSubhead">BBBS - <%= y %> Annual Agency Information (AAI)</td>
		</tr>
		<tr>
			<td colspan="4" class="formHeader">SPECIAL PROGRAMS</td>
		</tr>
		<tr>
			<td colspan="4" class="formMain"><font color="#ff0000"><div align="center"><strong>Please Note: </strong>After entering your information, you <strong>must</strong> click on the "Save Form" button at the bottom of the form and wait for the "Thank You" screen or your changes will be lost.</div></font></td>
		</tr>				
		<tr>
<!--			<td colspan="4" valign="top" class="formMain">If your agency operated any of the following <b>Special Programs</b> in this ADS year with six or more participants and: Specified staff members responsible for program design and /or implementation of;<br> Separate funding/budget or portion of core budget designated for the program or;<br> Guidelines or procedures unique to the program then <b>indicate quantity below</b>:</td>-->
			<td colspan="4" valign="top" class="formMain">If your agency operated any of the following <b>Special Programs</b> in this ADS year with six or more participants and: Specified staff members responsible for program design and /or implementation of;<br> Separate funding/budget or portion of core budget designated for the program or;<br> Guidelines or procedures unique to the program then <b><font color="#ff0000">indicate number of PARTICIPANTS below:</font></b></td>			
		</tr>
		<tr>
		<tr>
			<td align="left" valign="top" class="formMain">55 or Older as Bigs<br>(IntergenerationalProgram):</td>
			<td align="right" valign="top" class="formMain"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("FiftyFiveOrOlderBigs") %><% Else %>0<% End If %>" name="frmSpecialProgramsFiftyFiveOrOlderBigs" onchange="checkForInteger(this.value);"></td>
			<td align="left" valign="top" class="formMain">Group Activities:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("GroupActivities") %><% Else %>0<% End If %>" name="frmSpecialProgramsGroupActivities" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">Academic Achievement/Enrichment<br>(Literacy/Tutoring):</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("AcademicAchievement") %><% Else %>0<% End If %>" name="frmSpecialProgramsAcademicAchievement" onchange="checkForInteger(this.value);"></td>
			<td align="left" valign="top" class="formMain">High School Students as Bigs:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("HighSchoolStudentsAsBigs") %><% Else %>0<% End If %>" name="frmSpecialProgramsHighSchoolStudentsAsBigs" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">After School Mentoring:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("AfterSchoolMentoring") %><% Else %>0<% End If %>" name="frmSpecialProgramsAfterSchoolMentoring" onchange="checkForInteger(this.value);"></td>
			<td align="left" valign="top" class="formMain">Life Skills/Life Choices:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("LifeSkillsLifeChoices") %><% Else %>0<% End If %>" name="frmSpecialProgramsLifeSkillsLifeChoices" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">AIDS Prevention/Intervention:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("AIDSPreventionIntervention") %><% Else %>0<% End If %>" name="frmSpecialProgramsAIDSPreventionIntervention" onchange="checkForInteger(this.value);"></td>
			<td align="left" valign="top" class="formMain">Parent Support Groups:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("ParentSupportGroups") %><% Else %>0<% End If %>" name="frmSpecialProgramsParentSupportGroups" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">Alcohol Abuse Prevention/Intervention:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("AlcoholAbusePreventionIntervention") %><% Else %>0<% End If %>" name="frmSpecialProgramsAlcoholAbusePreventionIntervention" onchange="checkForInteger(this.value);"></td>
			<td align="left" valign="top" class="formMain">Partnerships: Civic Organizations:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("PartnershipsCivicOrganizations") %><% Else %>0<% End If %>" name="frmSpecialProgramsPartnershipsCivicOrganizations" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">Camping:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("Camping") %><% Else %>0<% End If %>" name="frmSpecialProgramsCamping" onchange="checkForInteger(this.value);"></td>
			<td align="left" valign="top" class="formMain">Partnerships: Colleges/Universities:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("PartnershipsCollegesUniversities") %><% Else %>0<% End If %>" name="frmSpecialProgramsPartnershipsCollegesUniversities" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">Character Counts:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("CharacterCounts") %><% Else %>0<% End If %>" name="frmSpecialProgramsCharacterCounts" onchange="checkForInteger(this.value);"></td>
			<td align="left" valign="top" class="formMain">Partnerships: Corporations/Businesses:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("PartnershipsCorporationsBusinesses") %><% Else %>0<% End If %>" name="frmSpecialProgramsPartnershipsCorporationsBusinesses" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">Children with Disabilities:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("ChildrenWithDisabilities") %><% Else %>0<% End If %>" name="frmSpecialProgramsChildrenWithDisabilities" onchange="checkForInteger(this.value);"></td>
			<td align="left" valign="top" class="formMain">Partnerships: Other Youth Serving Organizations:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("PartnershipsOtherYouthServingOrganizations") %><% Else %>0<% End If %>" name="frmSpecialProgramsPartnershipsOtherYouthServingOrganizations" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">College/University Students as Bigs:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("CollegeStudentsAsBigs") %><% Else %>0<% End If %>" name="frmSpecialProgramsCollegeStudentsAsBigs" onchange="checkForInteger(this.value);"></td>
			<td align="left" valign="top" class="formMain">Partnerships: Religious Organizations:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("PartnershipsReligiousOrganizations") %><% Else %>0<% End If %>" name="frmSpecialProgramsPartnershipsReligiousOrganizations" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">Community Service Projects:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("CommunityServiceProjects") %><% Else %>0<% End If %>" name="frmSpecialProgramsCommunityServiceProjects" onchange="checkForInteger(this.value);"></td>
			<td align="left" valign="top" class="formMain">Pregnancy Prevention:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("PregnancyPrevention") %><% Else %>0<% End If %>" name="frmSpecialProgramsPregnancyPrevention" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">Drop Out Prevention:</td>
			<td align="right"class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("DropOutPrevention") %><% Else %>0<% End If %>" name="frmSpecialProgramsDropOutPrevention" onchange="checkForInteger(this.value);"></td>
			<td align="left" valign="top" class="formMain">Scholarships:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("Scholarships") %><% Else %>0<% End If %>" name="frmSpecialProgramsScholarships" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td rowspan="2" align="left" valign="top" class="formMain">Drug Abuse Prevention/Intervention:</td>
			<td align="right" rowspan="2" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("DrugAbusePreventionIntervention") %><% Else %>0<% End If %>" name="frmSpecialProgramsDrugAbusePreventionIntervention" onchange="checkForInteger(this.value);"></td>
			<td rowspan="2" align="left" valign="top" class="formMain">Sexual Abuse Prevention/Intervention:</td>
			<td align="right" class="formMain" valign="top">EMPOWER:&nbsp;<input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("SexualAbusePreventionInterventionEmpower") %><% Else %>0<% End If %>" name="frmSpecialProgramsSexualAbusePreventionInterventionEmpower" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td align="right" class="formMain" valign="top"><b>NOT</b>&nbsp;EMPOWER:&nbsp;<input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("SexualAbusePreventionInterventionNOTEmpower") %><% Else %>0<% End If %>" name="frmSpecialProgramsSexualAbusePreventionInterventionNOTEmpower" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td rowspan="2" align="left" valign="top" class="formMain">Emergency Financial Assistance:</td>
			<td align="right" rowspan="2" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("EmergencyFinancialAssistance") %><% Else %>0<% End If %>" name="frmSpecialProgramsEmergencyFinancialAssistance" onchange="checkForInteger(this.value);"></td>
			<td rowspan="2" align="left" valign="top" class="formMain">Site-Based Mentoring:</td>
			<td align="right" class="formMain" valign="top">School Based: <input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("SiteBasedMentoringSchoolBased") %><% Else %>0<% End If %>" name="frmSpecialProgramsSiteBasedMentoringSchoolBased" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>			
			<td align="right" class="formMain" valign="top"><b>NOT</b> School Based: <input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("SiteBasedMentoringNOTSchoolBased") %><% Else %>0<% End If %>" name="frmSpecialProgramsSiteBasedMentoringNOTSchoolBased" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">Employability/Job Readiness<br>(School To Work):</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("EmployabilityJobReadiness") %><% Else %>0<% End If %>" name="frmSpecialProgramsEmployabilityJobReadiness" onchange="checkForInteger(this.value);"></td>
			<td align="left" valign="top" class="formMain">Teen Parenting:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("TeenParenting") %><% Else %>0<% End If %>" name="frmSpecialProgramsTeenParenting" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">Family Counseling:</td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("FamilyCounseling") %><% Else %>0<% End If %>" name="frmSpecialProgramsFamilyCounseling" onchange="checkForInteger(this.value);"></td>
			<td align="left" valign="top" class="formMain">Other: <input type="text" class="formMain" size="25" maxlength="50" value="<% If say = "edit" Then %><%= GetSpecialPrograms("OtherText") %><% Else %>(Name Program)<% End If %>" name="frmSpecialProgramsOtherText"></td>
			<td align="right" class="formMain" valign="top"><input type="text" class="formMain" size="5" maxlength="8" value="<% If say = "edit" Then %><%= GetSpecialPrograms("Other") %><% Else %>0<% End If %>" name="frmSpecialProgramsOther" onchange="checkForInteger(this.value);"></td>
	</tr>
		<tr>
		<td colspan="4" class="formHeader"><input type="submit" value="Save Form" class="formMainBold"></td>
		</tr>
		
		<tr>
		<td colspan="4"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
		</tr>
	</table>
</td>
</tr>	
</table>

<% 
If say = "edit" Then
	GetSpecialPrograms.Close
	Set GetSpecialPrograms = Nothing
	Con.Close
	Set Con = Nothing
End If
 %>
</form>
<% End If %>
<p></p>
<p></p>
</body>
</html>
