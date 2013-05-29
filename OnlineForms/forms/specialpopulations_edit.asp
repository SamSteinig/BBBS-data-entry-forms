<% 
If Request("status") = "addNew" Then

' Check for duplicate records
	
	Set DupCon = Server.CreateObject("ADODB.Connection")
	DupCon.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT Count(*) As NumberOfEntries FROM tbl_frmSpecialPopulations WHERE AgencyID = '" & Request("AgencyIDN") & "' and Year = " & Request("Year")	
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
		RST.Open "SELECT * FROM tbl_frmSpecialPopulations", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = Request("AgencyIDN")
		RST("Year") = Request("year")
		RST("AbusedNeglected") = Request("frmSpecialPopulationsAbusedNeglected")
		RST("AdjudicatedDelinquents") = Request("frmSpecialPopulationsAdjudicatedDelinquents")
		RST("AfterSchool") = Request("frmSpecialPopulationsAfterSchool")
		RST("AIDSAffected") = Request("frmSpecialPopulationsAIDSAffected")
		RST("DeafHearingImpaired") = Request("frmSpecialPopulationsDeafHearingImpaired")
		RST("DevelopmentallyDisabled") = Request("frmSpecialPopulationsDevelopmentallyDisabled")
		RST("FosterChildren") = Request("frmSpecialPopulationsFosterChildren")
		RST("Homeless") = Request("frmSpecialPopulationsHomeless")
		RST("IncarceratedParents") = Request("frmSpecialPopulationsIncarceratedParents")
		RST("Institutionalized") = Request("frmSpecialPopulationsInstitutionalized")
		RST("LearningDisabled") = Request("frmSpecialPopulationsLearningDisabled")
		RST("PhysicallyDisabled") = Request("frmSpecialPopulationsPhysicallyDisabled")
		RST("PregnantTeen") = Request("frmSpecialPopulationsPregnantTeen")
		RST("SchoolDropouts") = Request("frmSpecialPopulationsSchoolDropouts")
		RST("TeenParentsFemale") = Request("frmSpecialPopulationsTeenParentsFemale")
		RST("TeenParentsMale") = Request("frmSpecialPopulationsTeenParentsMale")
		RST("VisuallyImpaired") = Request("frmSpecialPopulationsVisuallyImpaired")
		RST("OtherType") = Request("frmSpecialPopulationsOtherType")
		RST("Other") = Request("frmSpecialPopulationsOther")
		RST("CreateDate") = Now
		RST.Update
		RST.Close
		Set RST = Nothing
		form = "SpecialPopulations"
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
	RST.Open "SELECT * FROM tbl_frmSpecialPopulations WHERE agencyID='" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")), Con, 1, 3
	RST("AbusedNeglected") = Request("frmSpecialPopulationsAbusedNeglected")
	RST("AdjudicatedDelinquents") = Request("frmSpecialPopulationsAdjudicatedDelinquents")
	RST("AfterSchool") = Request("frmSpecialPopulationsAfterSchool")
	RST("AIDSAffected") = Request("frmSpecialPopulationsAIDSAffected")
	RST("DeafHearingImpaired") = Request("frmSpecialPopulationsDeafHearingImpaired")
	RST("DevelopmentallyDisabled") = Request("frmSpecialPopulationsDevelopmentallyDisabled")
	RST("FosterChildren") = Request("frmSpecialPopulationsFosterChildren")
	RST("Homeless") = Request("frmSpecialPopulationsHomeless")
	RST("IncarceratedParents") = Request("frmSpecialPopulationsIncarceratedParents")
	RST("Institutionalized") = Request("frmSpecialPopulationsInstitutionalized")
	RST("LearningDisabled") = Request("frmSpecialPopulationsLearningDisabled")
	RST("PhysicallyDisabled") = Request("frmSpecialPopulationsPhysicallyDisabled")
	RST("PregnantTeen") = Request("frmSpecialPopulationsPregnantTeen")
	RST("SchoolDropouts") = Request("frmSpecialPopulationsSchoolDropouts")
	RST("TeenParentsFemale") = Request("frmSpecialPopulationsTeenParentsFemale")
	RST("TeenParentsMale") = Request("frmSpecialPopulationsTeenParentsMale")
	RST("VisuallyImpaired") = Request("frmSpecialPopulationsVisuallyImpaired")
	RST("OtherType") = Request("frmSpecialPopulationsOtherType")
	RST("Other") = Request("frmSpecialPopulationsOther")
	jMod = RST("SpecialPopulationsID")
	RST.Update
	RST.Close
	Set RST = Nothing
	form = "SpecialPopulations"
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
	<title>Special Populations</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
<script language="javascript">
	<!--
	
function checkForInteger(valueToCheck)
{
	var myRegularExpression = /^[0-9]+(,[0-9]{3})*$/;
	var replaceWhiteSpace = /\s/; // searches for any whitespace character
	var formField = valueToCheck; // passed in as parameter 1
	var newFormField = valueToCheck.replace(replaceWhiteSpace, ""); // remove any whitespace from the form entry and replace it with nothing
	var bContainsNonNumbers = myRegularExpression.test(newFormField); // check newFormField variable to see if it contains any nonnumeric character
	
	if(!bContainsNonNumbers)
	{
		alert("Please make sure you have entered a whole number.\n We cannot process letters or words."); 
	} 
}

var myRegularExpression1 = /^[0-9]+(,[0-9]{3})*$/;
	
function submitFormValidate(form)
{
	if(!(myRegularExpression1.test(form.frmSpecialPopulationsAbusedNeglected.value)) || (form.frmSpecialPopulationsAbusedNeglected.value == ""))
	{
		form.frmSpecialPopulationsAbusedNeglected.focus();
		if(form.frmSpecialPopulationsAbusedNeglected.value == "")
			alert("You must provide a value for all form fields");
		else
			alert((form.frmSpecialPopulationsAbusedNeglected.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialPopulationsInstitutionalized.value)) || (form.frmSpecialPopulationsInstitutionalized.value == ""))	
	{
		form.frmSpecialPopulationsInstitutionalized.focus();
		if(form.frmSpecialPopulationsInstitutionalized.value == "")
			alert("You must provide a value for all form fields");
		else
			alert(form.frmSpecialPopulationsInstitutionalized.value + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialPopulationsAdjudicatedDelinquents.value)) || (form.frmSpecialPopulationsAdjudicatedDelinquents.value == ""))	
	{
		form.frmSpecialPopulationsAdjudicatedDelinquents.focus();
		if(form.frmSpecialPopulationsAdjudicatedDelinquents.value == "")
			alert("You must provide a value for all form fields");
		else
			alert(form.frmSpecialPopulationsAdjudicatedDelinquents.value + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialPopulationsLearningDisabled.value)) || (form.frmSpecialPopulationsLearningDisabled.value == ""))	
	{
		form.frmSpecialPopulationsLearningDisabled.focus();
		if(form.frmSpecialPopulationsLearningDisabled.value == "")
			alert("You must provide a value for all form fields");
		else
			alert(form.frmSpecialPopulationsLearningDisabled.value + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialPopulationsAfterSchool.value)) || (form.frmSpecialPopulationsAfterSchool.value == ""))	
	{
		form.frmSpecialPopulationsAfterSchool.focus();
		if(form.frmSpecialPopulationsAfterSchool.value == "")
			alert("You must provide a value for all form fields");
		else
			alert(form.frmSpecialPopulationsAfterSchool.value + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialPopulationsPhysicallyDisabled.value)) || (form.frmSpecialPopulationsPhysicallyDisabled.value == ""))	
	{
		form.frmSpecialPopulationsPhysicallyDisabled.focus();
		if(form.frmSpecialPopulationsPhysicallyDisabled.value == "")
			alert("You must provide a value for all form fields");
		else
			alert(form.frmSpecialPopulationsPhysicallyDisabled.value + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialPopulationsAIDSAffected.value)) || (form.frmSpecialPopulationsAIDSAffected.value == ""))	
	{
		form.frmSpecialPopulationsAIDSAffected.focus();
		if(form.frmSpecialPopulationsAIDSAffected.value == "")
			alert("You must provide a value for all form fields");
		else
			alert(form.frmSpecialPopulationsAIDSAffected.value + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialPopulationsPregnantTeen.value)) || (form.frmSpecialPopulationsPregnantTeen.value == ""))	
	{
		form.frmSpecialPopulationsPregnantTeen.focus();
		if(form.frmSpecialPopulationsPregnantTeen.value == "")
			alert("You must provide a value for all form fields");
		else
			alert(form.frmSpecialPopulationsPregnantTeen.value + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialPopulationsDeafHearingImpaired.value)) || (form.frmSpecialPopulationsDeafHearingImpaired.value == ""))	
	{
		form.frmSpecialPopulationsDeafHearingImpaired.focus();
		if(form.frmSpecialPopulationsDeafHearingImpaired.value == "")
			alert("You must provide a value for all form fields");
		else
			alert(form.frmSpecialPopulationsDeafHearingImpaired.value + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialPopulationsSchoolDropouts.value)) || (form.frmSpecialPopulationsSchoolDropouts.value == ""))	
	{
		form.frmSpecialPopulationsSchoolDropouts.focus();
		if(form.frmSpecialPopulationsSchoolDropouts.value == "")
			alert("You must provide a value for all form fields");
		else
			alert(form.frmSpecialPopulationsSchoolDropouts.value + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialPopulationsDevelopmentallyDisabled.value)) || (form.frmSpecialPopulationsDevelopmentallyDisabled.value == ""))	
	{
		form.frmSpecialPopulationsDevelopmentallyDisabled.focus();
		if(form.frmSpecialPopulationsDevelopmentallyDisabled.value == "")
			alert("You must provide a value for all form fields");
		else
			alert(form.frmSpecialPopulationsDevelopmentallyDisabled.value + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialPopulationsTeenParentsFemale.value)) || (form.frmSpecialPopulationsTeenParentsFemale.value == ""))	
	{
		form.frmSpecialPopulationsTeenParentsFemale.focus();
		if(form.frmSpecialPopulationsTeenParentsFemale.value == "")
			alert("You must provide a value for all form fields");
		else
			alert(form.frmSpecialPopulationsTeenParentsFemale.value + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialPopulationsFosterChildren.value)) || (form.frmSpecialPopulationsFosterChildren.value == ""))	
	{
		form.frmSpecialPopulationsFosterChildren.focus();
		if(form.frmSpecialPopulationsFosterChildren.value == "")
			alert("You must provide a value for all form fields");
		else
			alert(form.frmSpecialPopulationsFosterChildren.value + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialPopulationsTeenParentsMale.value)) || (form.frmSpecialPopulationsTeenParentsMale.value == ""))	
	{
		form.frmSpecialPopulationsTeenParentsMale.focus();
		if(form.frmSpecialPopulationsTeenParentsMale.value == "")
			alert("You must provide a value for all form fields");
		else
			alert(form.frmSpecialPopulationsTeenParentsMale.value + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialPopulationsHomeless.value)) || (form.frmSpecialPopulationsHomeless.value == ""))	
	{
		form.frmSpecialPopulationsHomeless.focus();
		if(form.frmSpecialPopulationsHomeless.value == "")
			alert("You must provide a value for all form fields");
		else
			alert(form.frmSpecialPopulationsHomeless.value + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialPopulationsVisuallyImpaired.value)) || (form.frmSpecialPopulationsVisuallyImpaired.value == ""))	
	{
		form.frmSpecialPopulationsVisuallyImpaired.focus();
		if(form.frmSpecialPopulationsVisuallyImpaired.value == "")
			alert("You must provide a value for all form fields");
		else
			alert(form.frmSpecialPopulationsVisuallyImpaired.value + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialPopulationsIncarceratedParents.value)) || (form.frmSpecialPopulationsIncarceratedParents.value == ""))	
	{
		form.frmSpecialPopulationsIncarceratedParents.focus();
		if(form.frmSpecialPopulationsIncarceratedParents.value == "")
			alert("You must provide a value for all form fields");
		else
			alert(form.frmSpecialPopulationsIncarceratedParents.value + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmSpecialPopulationsOther.value)) || (form.frmSpecialPopulationsOther.value == ""))	
	{
		form.frmSpecialPopulationsOther.focus();
		if(form.frmSpecialPopulationsOther.value == "")
			alert("You must provide a value for all form fields");
		else
			alert(form.frmSpecialPopulationsOther.value + " is invalid.");
		return false;
	}
	else if((form.frmSpecialPopulationsOther.value != "0") && (form.frmSpecialPopulationsOther.value != "") && (form.frmSpecialPopulationsOther.value != "000"))
	{
		if((form.frmSpecialPopulationsOtherType.value == "(Name Population)") || (form.frmSpecialPopulationsOtherType.value == ""))
			{
			form.frmSpecialPopulationsOtherType.focus();
			alert("Please Name Population ");
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
<td width="220" valign="top"><img src="../includes/images/photos_baseball.jpg" alt="" width="220" height="477" border="0"></td>
<td width="100%" valign="top">
<br>	

<% If say = "thanks" Then %>

<font class="formMain"><br><br>
<strong>Thank you!</strong> Your information has been saved in the BBBS database.<br>
To choose another form, please select the form type from the choices above.
<br><br>
<i>Please note: These changes will not be reflected in the <strong>Agency Profile</strong> (in the My Agency Page and the Agency Directory) for 24 hours.</i>
</font>

<br>
<!--#include file="../includes/contact_info.inc"-->
<br>



<% ElseIf say <> "thanks" Then  %>
<table border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063" width="400">

<form name="frmSpecialPopulations" action="specialpopulations_edit.asp" method="post" onsubmit="return submitFormValidate(this)">
<!--#include file="../includes/form_stamp.asp"-->

<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmSpecialPopulations WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
	Set GetSpecialPopulations = Con.Execute(query)
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
			<td colspan="4" class="formHeader">SPECIAL POPULATIONS</td>
		</tr>
		
		<tr>
			<td colspan="4" class="formMain"><font color="#ff0000"><div align="center"><strong>Please Note: </strong>After entering your information, you <strong>must</strong> click on the "Save Form" button at the bottom of the form and wait for the "Thank You" screen or your changes will be lost.</div></font></td>
		</tr>				
<!-- Section 1 -->
		<tr>
			<td colspan="4" align="center" valign="top" class="formMain">If your agency served any of the following <b>Special Populations</b> in this<br>ADS year with a coordinated effort to reach a <b>Targeted</b> population<br>serving at least <b>six recipients</b>, then <b>indicate quantity below</b>:</td>
		</tr>
		<tr>
			<td class="formMain">Abused/Neglected:</td>
			<td class="formMain"><input type="text" class="formMain" size="6" value="<% If say = "edit" Then %><%= GetSpecialPopulations("AbusedNeglected") %><% Else %>0<% End If %>" name="frmSpecialPopulationsAbusedNeglected" onchange="checkForInteger(this.value);"></td>
			<td class="formMain">Institutionalized:</td>
			<td class="formMain"><input type="text" class="formMain" size="6" value="<% If say = "edit" Then %><%= GetSpecialPopulations("Institutionalized") %><% Else %>0<% End If %>" name="frmSpecialPopulationsInstitutionalized" onchange="checkForInteger(this.value);"></td>
		</tr>	
		<tr>
			<td class="formMain">Adjudicated Delinquents:</td>
			<td class="formMain"><input type="text" class="formMain" size="6" value="<% If say = "edit" Then %><%= GetSpecialPopulations("AdjudicatedDelinquents") %><% Else %>0<% End If %>" name="frmSpecialPopulationsAdjudicatedDelinquents" onchange="checkForInteger(this.value);"></td>
			<td class="formMain">Learning Disabled:
			<td class="formMain"><input type="text" class="formMain" size="6" value="<% If say = "edit" Then %><%= GetSpecialPopulations("LearningDisabled") %><% Else %>0<% End If %>" name="frmSpecialPopulationsLearningDisabled" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td class="formMain">After School (Latchkey):</td>
			<td class="formMain"><input type="text" class="formMain" size="6" value="<% If say = "edit" Then %><%= GetSpecialPopulations("AfterSchool") %><% Else %>0<% End If %>" name="frmSpecialPopulationsAfterSchool" onchange="checkForInteger(this.value);"></td>
			<td class="formMain">Physically Disabled:</td>
			<td class="formMain"><input type="text" class="formMain" size="6" value="<% If say = "edit" Then %><%= GetSpecialPopulations("PhysicallyDisabled") %><% Else %>0<% End If %>" name="frmSpecialPopulationsPhysicallyDisabled" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td class="formMain">AIDS Affected:</td>
			<td class="formMain"><input type="text" class="formMain" size="6" value="<% If say = "edit" Then %><%= GetSpecialPopulations("AIDSAffected") %><% Else %>0<% End If %>" name="frmSpecialPopulationsAIDSAffected" onchange="checkForInteger(this.value);"></td>
			<td class="formMain">Pregnant Teen:</td>
			<td class="formMain"><input type="text" class="formMain" size="6" value="<% If say = "edit" Then %><%= GetSpecialPopulations("PregnantTeen") %><% Else %>0<% End If %>" name="frmSpecialPopulationsPregnantTeen" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td class="formMain">Deaf & Hearing Impaired:</td>
			<td class="formMain"><input type="text" class="formMain" size="6" value="<% If say = "edit" Then %><%= GetSpecialPopulations("DeafHearingImpaired") %><% Else %>0<% End If %>" name="frmSpecialPopulationsDeafHearingImpaired" onchange="checkForInteger(this.value);"></td>
			<td class="formMain">School Dropouts:</td>
			<td class="formMain"><input type="text" class="formMain" size="6" value="<% If say = "edit" Then %><%= GetSpecialPopulations("SchoolDropouts") %><% Else %>0<% End If %>" name="frmSpecialPopulationsSchoolDropouts" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td class="formMain">Developmentally Disabled:</td>
			<td class="formMain"><input type="text" class="formMain" size="6" value="<% If say = "edit" Then %><%= GetSpecialPopulations("DevelopmentallyDisabled") %><% Else %>0<% End If %>" name="frmSpecialPopulationsDevelopmentallyDisabled" onchange="checkForInteger(this.value);"></td>
			<td class="formMain">Teen Parents (Female):</td>
			<td class="formMain"><input type="text" class="formMain" size="6" value="<% If say = "edit" Then %><%= GetSpecialPopulations("TeenParentsFemale") %><% Else %>0<% End If %>" name="frmSpecialPopulationsTeenParentsFemale" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td class="formMain">Foster Children:</td>
			<td class="formMain"><input type="text" class="formMain" size="6" value="<% If say = "edit" Then %><%= GetSpecialPopulations("FosterChildren") %><% Else %>0<% End If %>" name="frmSpecialPopulationsFosterChildren" onchange="checkForInteger(this.value);"></td>
			<td class="formMain">Teen Parents (Male):</td>
			<td class="formMain"><input type="text" class="formMain" size="6" value="<% If say = "edit" Then %><%= GetSpecialPopulations("TeenParentsMale") %><% Else %>0<% End If %>" name="frmSpecialPopulationsTeenParentsMale" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td class="formMain">Homeless:</td>
			<td class="formMain"><input type="text" class="formMain" size="6" value="<% If say = "edit" Then %><%= GetSpecialPopulations("Homeless") %><% Else %>0<% End If %>" name="frmSpecialPopulationsHomeless" onchange="checkForInteger(this.value);"></td>
			<td class="formMain">Visually Impaired:</td>
			<td class="formMain"><input type="text" class="formMain" size="6" value="<% If say = "edit" Then %><%= GetSpecialPopulations("VisuallyImpaired") %><% Else %>0<% End If %>" name="frmSpecialPopulationsVisuallyImpaired" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td class="formMain">Incarcerated Parents:</td>
			<td class="formMain"><input type="text" class="formMain" size="6" value="<% If say = "edit" Then %><%= GetSpecialPopulations("IncarceratedParents") %><% Else %>0<% End If %>" name="frmSpecialPopulationsIncarceratedParents" onchange="checkForInteger(this.value);"></td>
			<td class="formMain">Other: <input type="text" class="formMain" size="20" maxlength="50" value="<% If say = "edit" Then %><%= GetSpecialPopulations("OtherType") %><% Else %>(Name Population)<% End If %>" name="frmSpecialPopulationsOtherType"></td>
			<td class="formMain"><input type="text" class="formMain" size="6" value="<% If say = "edit" Then %><%= GetSpecialPopulations("Other") %><% Else %>0<% End If %>" name="frmSpecialPopulationsOther" onchange="checkForInteger(this.value);"></td>
		</tr>
		<tr>
			<td colspan="4" class="formHeader"><input type="submit" value="Save Form" class="formMainBold"></td>
		</tr>
		<tr>
			<td colspan="4"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
		</tr>
	</table>
	

<% 
If say = "edit" Then
	GetSpecialPopulations.Close
	Set GetSpecialPopulations = Nothing
	Con.Close
	Set Con = Nothing
End If
 %>



</form>
<% End If %>
<p></p>
<p></p>
</td>
</tr>
</table>
</body>
</html>
