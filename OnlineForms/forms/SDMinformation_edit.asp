<% 
If Request("status") = "addNew" Then

	' Check for duplicate records
	
	Set DupCon = Server.CreateObject("ADODB.Connection")
	DupCon.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT Count(*) As NumberOfEntries FROM tbl_frmSDMInformation WHERE AgencyID = '" & Request("AgencyIDN") & "' and Year = " & Request("Year")	
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
		RST.Open "SELECT * FROM tbl_frmSDMInformation", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = Request("AgencyIDN")
		RST("Year") = Request("year")
		RST("VolunteerInquiries") = Request("frmSDMInformationVolunteerInquiries")
		RST("VolunteerInPersonInterviews") = Request("frmSDMInformationVolunteerInPersonInterviews")
		RST("VolunteersMatched") = Request("frmSDMInformationVolunteersMatched")
		RST("VolunteerReMatchRate") = Request("frmSDMInformationVolunteerReMatchRate")
		RST("YouthInquiries") = Request("frmSDMInformationYouthInquiries")
		RST("YouthInPersonInterviews") = Request("frmSDMInformationYouthInPersonInterviews")
		RST("YouthsMatched") = Request("frmSDMInformationYouthsMatched")
		RST("CreateDate") = Now
		RST.Update
		RST.Close
		Set RST = Nothing
		form = "SDMInformation"
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
	RST.Open "SELECT * FROM tbl_frmSDMInformation WHERE agencyID='" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")), Con, 1, 3
	RST("AgencyID") = Request("AgencyIDN")
	RST("Year") = Request("year")
	RST("VolunteerInquiries") = Request("frmSDMInformationVolunteerInquiries")
	RST("VolunteerInPersonInterviews") = Request("frmSDMInformationVolunteerInPersonInterviews")
	RST("VolunteersMatched") = Request("frmSDMInformationVolunteersMatched")
	RST("VolunteerReMatchRate") = Request("frmSDMInformationVolunteerReMatchRate")
	RST("YouthInquiries") = Request("frmSDMInformationYouthInquiries")
	RST("YouthInPersonInterviews") = Request("frmSDMInformationYouthInPersonInterviews")
	RST("YouthsMatched") = Request("frmSDMInformationYouthsMatched")
	jMod = RST("SDMInformationID")
	RST.Update
	RST.Close
	Set RST = Nothing
	form = "SDMInformation"
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
	<title>SDM Information</title>
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
				alert("Please make sure you have entered a whole number. We cannot process letters, words or negative numbers."); 
			} 
		}
		
		
		function equalsLessThan101(valueToCheck)
		{
			if((valueToCheck < 101) && (valueToCheck >= 0))
			{
				return true;	
			}
			else
			{
				alert("Percentage cannot be greater than 100 or less than 0.");
				return false;
			}
		}		
		



var myRegularExpression1 = /^[0-9]+(,[0-9]{3})*$/;
		
function submitFormValidate(form)
{
	

	/* 1 */
	if(!(myRegularExpression1.test(form.frmSDMInformationVolunteerInquiries.value)) || (form.frmSDMInformationVolunteerInquiries.value == ""))	
	{
		form.frmSDMInformationVolunteerInquiries.focus();
		if(form.frmSDMInformationVolunteerInquiries.value == "")
			alert("You must provide a value for all form fields");
		else
			alert((form.frmSDMInformationVolunteerInquiries.value) + " is an invalid value for Volunteer Inquiries.");
		return false;
	}
	
	
	/* 2 */
		else if(!(myRegularExpression1.test(form.frmSDMInformationVolunteerInPersonInterviews.value)) || (form.frmSDMInformationVolunteerInPersonInterviews.value == ""))	
	{
		form.frmSDMInformationVolunteerInPersonInterviews.focus();
		if(form.frmSDMInformationVolunteerInPersonInterviews.value == "")
			alert("You must provide a value for all form fields");
		else
			alert("'" + (form.frmSDMInformationVolunteerInPersonInterviews.value) + "' is an invalid value for Volunteer In Person Interviews.");
		return false;
	}
	
	
	/* 3 */
		else if(!(myRegularExpression1.test(form.frmSDMInformationVolunteersMatched.value)) || (form.frmSDMInformationVolunteersMatched.value == ""))	
	{
		form.frmSDMInformationVolunteersMatched.focus();
		if(form.frmSDMInformationVolunteersMatched.value == "")
			alert("You must provide a value for all form fields");
		else
			alert("'" + (form.frmSDMInformationVolunteersMatched.value) + "' is an invalid value for Volunteers Matched.");
		return false;
	}	

	/* 4 */
	
		else if(!(myRegularExpression1.test(form.frmSDMInformationVolunteerRematchRate.value)) || (form.frmSDMInformationVolunteerRematchRate.value == ""))	
	{
		form.frmSDMInformationVolunteerRematchRate.focus();
		if(form.frmSDMInformationVolunteerRematchRate.value == "")
			alert("You must provide a value for all form fields");
		else
			alert("Volunteer Rematch Rate value is invalid.");
		return false;
	}	
	
	
	else if((form.frmSDMInformationVolunteerRematchRate.value > 100) || (form.frmSDMInformationVolunteerRematchRate.value < 0))	
	{
		form.frmSDMInformationVolunteerRematchRate.focus();
		alert("'" + (form.frmSDMInformationVolunteerRematchRate.value) + "' is invalid. Please enter a number between 0 and 100.");
		return false;
	}		
	

	/* 5 */
		else if(!(myRegularExpression1.test(form.frmSDMInformationYouthInquiries.value)) || (form.frmSDMInformationYouthInquiries.value == ""))	
	{
		form.frmSDMInformationYouthInquiries.focus();
		if(form.frmSDMInformationYouthInquiries.value == "")
			alert("You must provide a value for all form fields");
		else
			alert("'" + (form.frmSDMInformationYouthInquiries.value) + "' is an invalid value for Youth Inquiries.");
		return false;
	}	
	

	/* 6 */
		else if(!(myRegularExpression1.test(form.frmSDMInformationYouthInPersonInterviews.value)) || (form.frmSDMInformationYouthInPersonInterviews.value == ""))	
	{
		form.frmSDMInformationYouthInPersonInterviews.focus();
		if(form.frmSDMInformationYouthInPersonInterviews.value == "")
			alert("You must provide a value for all form fields");
		else
			alert("'" + (form.frmSDMInformationYouthInPersonInterviews.value) + "' is an invalid value for Youth In Person Interviews.");
		return false;
	}		
	
	/* 7 */
		else if(!(myRegularExpression1.test(form.frmSDMInformationYouthsMatched.value)) || (form.frmSDMInformationYouthsMatched.value == ""))	
	{
		form.frmSDMInformationYouthsMatched.focus();
		if(form.frmSDMInformationYouthsMatched.value == "")
			alert("You must provide a value for all form fields");
		else
			alert("'" + (form.frmSDMInformationYouthsMatched.value) + "' is an invalid value for Youth In Person Interviews.");
		return false;
	}				
	
	
	
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
<form name="frmSDMInformation" action="SDMInformation_edit.asp" method="post" onsubmit="return submitFormValidate(this);">
<!--#include file="../includes/form_stamp.asp"-->

<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmSDMInformation WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
	Set GetSDMInformation = Con.Execute(query)
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
					<td colspan="3" class="formHeader">SDM INFORMATION</td>
				</tr>
				<tr>
					<td colspan="7" class="formMain"><font color="#ff0000"><div align="center"><strong>Please Note: </strong>After entering your information, you <strong>must</strong> click on the "Save Form" button at the bottom of the form and wait for the "Thank You" screen or your changes will be lost.</div></font></td>
				</tr>	
				
				<tr>
					<td colspan="7" class="formHeaderSmall">VOLUNTEER</td>
				</tr>
				
				
<!-- Question Number 1 -->
				<tr>
					<td align="right" valign="top" class="formMain">1.</td>
					<td align="left" valign="top" class="formMain">What is the total number of volunteer inquiries you received in <%=y%>?<br><span class="formSubHead">A volunteer is considered to have inquired when he/she contacts the agency, expresses an interest in being a Big and provides basic contact information.  Contact includes web-based inquiries.</span></td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% If say = "edit" Then %><%=  GetSDMInformation("VolunteerInquiries") %><% Else %>0<% End If %>" name="frmSDMInformationVolunteerInquiries" onchange="checkForInteger(this.value);"></td>
				</tr>
				
<!-- Question Number 2 -->
				<tr>
					<td align="right" valign="top" class="formMain">2.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Volunteer In Person interviews in <%=y%>?</td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% If say = "edit" Then %><%= GetSDMInformation("VolunteerInPersonInterviews") %><% Else %>0<% End If %>" name="frmSDMInformationVolunteerInPersonInterviews" onchange="checkForInteger(this.value);"></td>
				</tr>
<!-- Question Number 3 -->
				<tr> 
					<td align="right" valign="top" class="formMain">3.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Volunteers that were matched in <%=y%>?</td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% If say = "edit" Then %><%= GetSDMInformation("VolunteersMatched") %><% Else %>0<% End If %>" name="frmSDMInformationVolunteersMatched" onchange="checkForInteger(this.value);"></td>
				</tr>

<!-- Question Number 4 -->
				<tr>
					<td align="right" valign="top" class="formMain">4.</td>
					<td align="left" valign="top" class="formMain">Volunteer Rematch Rate<br>
					<span class="formSubHead">
					If you collect this data, please report it.<hr>
					Volunteer Rematch Rate is calculated as follows: <br><br><em>&nbsp;&nbsp;&nbsp;&nbsp;Number of Volunteers Rematched with a New Child in <%=y%><br><strong>&nbsp;&nbsp;&nbsp;&nbsp;DIVIDED BY</strong><br>&nbsp;&nbsp;&nbsp;&nbsp;the Total Number of Closed Matches in <%=y%>.</em>
					</span>
					</td>
					<td align="right" valign="top"><input type="text" class="formMain" size="3" maxlength="3" value="<% If say = "edit" Then %><%= GetSDMInformation("VolunteerRematchRate") %><% Else %>0<% End If %>" name="frmSDMInformationVolunteerRematchRate" onchange="checkForInteger(this.value); equalsLessThan101(this.value);">%</td>
				</tr>
				
				<tr>
					<td colspan="7" class="formHeaderSmall">YOUTH</td>
				</tr>				
				

<!-- Question Number 5 -->
				<tr>
					<td align="right" valign="top" class="formMain">5.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Youth Inquiries you received in <%=y%>?<br>
					<span class="formSubHead">
					A Youth is considered to have inquired when his/her parent or guardian contacts the agency, expresses an interest in getting a Big and provides basic contact information.
					</span>
					</td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% If say = "edit" Then %><%= GetSDMInformation("YouthInquiries") %><% Else %>0<% End If %>" name="frmSDMInformationYouthInquiries" onchange="checkForInteger(this.value);"></td>
				</tr>				
				
				
<!-- Question Number 6 -->
				<tr>
					<td align="right" valign="top" class="formMain">6.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Youth In Person Interviews in <%=y%>?</td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% If say = "edit" Then %><%= GetSDMInformation("YouthInPersonInterviews") %><% Else %>0<% End If %>" name="frmSDMInformationYouthInPersonInterviews" onchange="checkForInteger(this.value);"></td>
				</tr>
<!-- Question Number 7 -->
				<tr>
					<td align="right" valign="top" class="formMain">7.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Youth who were matched in <%=y%>?</td>
					<td align="right" valign="top"><input type="text" class="formMain" size="5" maxlength="18" value="<% If say = "edit" Then %><%= GetSDMInformation("YouthsMatched") %><% Else %>0<% End If %>" name="frmSDMInformationYouthsMatched" onchange="checkForInteger(this.value);"></td>
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
	GetSDMInformation.Close
	Set GetSDMInformation = Nothing
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

