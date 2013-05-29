<% 
If Request("status") = "addNew" Then
	
	' Check for duplicate records
	
	Set DupCon = Server.CreateObject("ADODB.Connection")
	DupCon.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT Count(*) As NumberOfEntries FROM tbl_frmBudgetForecast WHERE AgencyID = '" & Request("AgencyIDN") & "' and Year = " & Request("Year")
response.write query
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
		RST.Open "SELECT * FROM tbl_frmBudgetForecast", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = Request("AgencyIDN")
		RST("Year") = Request("year")
		
		RST("TotalBudgetPrcnt") = Int(Request("frmBudgetForecastTotalBudgetPrcnt"))
		RST("BenefitsBudgetPrcnt") = Int(Request("frmBudgetForecastBenefitsBudgetPrcnt"))
		RST("MeritIncreasePrcnt") = Int(Request("frmBudgetForecastMeritincreasePrcnt"))
		
		
		RST("CreateDate") = Now
		RST.Update
		RST.Close
		Set RST = Nothing
		form = "BudgetForecast"
		modtype = "new"
%>
		<!--#include file="../includes/modify_stamp.asp"-->
<%	
		say = "thanks"
		Con.Close
		Set Con = Nothing
	Else
		say = "previouslyEdited"
		Con.Close
		Set Con = Nothing
	End If

	
ElseIf Request("status") = "editSave" Then


	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	Set RST = Server.CreateObject("ADODB.Recordset")
	RST.Open "SELECT * FROM tbl_frmBudgetForecast WHERE AgencyID = '" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")), Con, 1, 3

		
		RST("TotalBudgetPrcnt") = Int(Request("frmBudgetForecastTotalBudgetPrcnt"))
		RST("BenefitsBudgetPrcnt") = Int(Request("frmBudgetForecastBenefitsBudgetPrcnt"))
		RST("MeritIncreasePrcnt") = Int(Request("frmBudgetForecastMeritIncreasePrcnt"))
		
		
	jMod = RST("BudgetForecastID")
	RST.Update
	RST.Close
	Set RST = Nothing
	form = "BudgetForecast"
	modtype = "Edit"
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
	<title>Staffing Expense Forecast</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
	
<script language="JavaScript">	
function checkForWholeNumber(valueToCheck)
{
	var myRegularExpression = /^\d*$/;  // Checks for integer
	var replaceWhiteSpace = /\s/; // searches for any whitespace character
	var formField = valueToCheck; // passed in as parameter 1
	var newFormField = valueToCheck.replace(replaceWhiteSpace, ""); // remove any whitespace from the form entry and replace it with nothing
	var bContainsNonNumbers = myRegularExpression.test(newFormField); // check newFormField variable to see if it contains any nonnumeric character
	if(!bContainsNonNumbers)
	{
		alert("Please make sure you have entered a whole number.\n We cannot process letters or words."); 
	} 
}

function equalsLessThan101(valueToCheck)
{
	if((valueToCheck < 101)&& (valueToCheck >= 0))
	{
		return true;	
	}
	else
	{
		alert("Percentage cannot be greater than 100 or less than 0.");
		return false;
	}
}
var myRegularExpression3 = 	/^\d*$/;  // Checks for integer 

function submitFormValidate(form)
{
	if((!(myRegularExpression3.test(form.frmBudgetForecastTotalBudgetPrcnt.value))) || (form.frmBudgetForecastTotalBudgetPrcnt.value > 100) || (form.frmBudgetForecastTotalBudgetPrcnt.value < 0))
	{
		
		alert((form.frmBudgetForecastTotalBudgetPrcnt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		form.frmBudgetForecastTotalBudgetPrcnt.focus();
		return false;
	}
	else if(!(myRegularExpression3.test(form.frmBudgetForecastBenefitsBudgetPrcnt.value)))	
	{
		form.frmBudgetForecastBenefitsBudgetPrcnt.focus();
		alert((form.frmBudgetForecastBenefitsBudgetPrcnt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	
	else if(!(myRegularExpression3.test(form.frmBudgetForecastMeritIncreasePrcnt.value)))	
	{
		form.frmBudgetForecastMeritIncreasePrcnt.focus();
		alert((form.frmBudgetForecastMeritIncreasePrcnt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else
	{
		return true;
	}
}
</script>

</head>

<% '<!--#include file="../includes/top_nav_forms_yearly.inc"--><!-- include file has </head> and <body> tags --><br>     %>
<!--#include file="../includes/surveytitle.inc"-->

<table width="100%" cellpadding="0" cellspacing="0" border="0" ID="Table2">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_slinky.jpg" alt="" width="220" height="477" border="0"></td>
<td width="100%" valign="top">
<br>
<% If say = "thanks" Then %>
<font class="formMain"><br><br>
<strong>Thank you!</strong> Your information has been saved in the BBBS database.<br>
To choose another form, please select the form type from the choices above.
</font>
<br>
<!--#include file="../includes/contact_info.inc"-->
<br>

<% ElseIf say <> "thanks" Then  %>
<table border="1" cellspacing="0" cellpadding="1" width = "650" bordercolordark="#003063" ID="Table1">

<form name="frmBudgetForecast" action="BudgetForecast_edit.asp" method="post" onsubmit="return submitFormValidate(this)">
<!--#include file="../includes/form_stamp.asp"-->
<%End If %>

<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmBudgetForecast WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
	Set GetBudget = Con.Execute(query)
 %>
<input type="hidden" name="status" value="editSave" ID="Hidden3">
<% Else %>
<input type="hidden" name="status" value="addNew" ID="Hidden4">
<%
End If
%>

<%If say = "previouslyEdited" Then%>
<p class="formMain">We're sorry, but this form was previously completed. To make changes please <a href="yearly.asp">reselect</a> the 
appropriate form and year and update the existing information.</p>
<%
Response.End
End If 
%>


	<tr>
		<td colspan="2" align="center" class="formSubhead">BBBS - <%= y %> Annual Agency Information (AAI)</td>
	</tr>
	<tr>
		<td colspan="2" class="formHeader">Staffing Expense Forecast</td>
	</tr>
	<tr>
		<td colspan="15" align="center" valign="top" class="formMain"><font color="#ff0000">Please enter the following Staffing Expense Forecast information<br>Click "Save Form" when you have completed each column. Saved information will appear in a grid below.</td>
	</tr>
	
	
	
	<tr>
		<td class="formMain">For <%= y+1 %> budget year, what is your projected Average salary increase?(Salary not including Incentives)
		<br>
	  <br>(Example: 5% budget increase in salaries is planned for <%= y+1 %>, including new positions and salary increase for existing positions)</td>
		<td align="right" valign="top" class="formMain" class="formMain">
			<input type="text" size="3" maxlength="3" class="formMain" NAME="frmBudgetForecastTotalBudgetPrcnt" value="<% If say = "edit" Then %><%= GetBudget("TotalBudgetPrcnt") %><% Else %>0<% End If %>" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">%</td>
	</tr>
	
	<tr>
		<td class="formMain">What overall percent increase are you forecasting for Medical, Vision, Dental premiums in employee benefits cost?<br> 
		   <td align="right" valign="top" class="formMain" class="formMain">
			<input type="text" size="3" maxlength="3" class="formMain" NAME="frmBudgetForecastBenefitsBudgetPrcnt" value="<% If say = "edit" Then %><%= GetBudget("BenefitsBudgetPrcnt") %><% Else %>0<% End If %>" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">%</td>
	</tr>	
	
	<tr>
		<td class="formMain">What is your merit increase budget?<br> 
		   <td align="right" valign="top" class="formMain" class="formMain">
			<input type="text" size="3" maxlength="3" class="formMain" NAME="frmBudgetForecastMeritIncreasePrcnt" value="<% If say = "edit" Then %><%= GetBudget("MeritIncreasePrcnt") %><% Else %>0<% End If %>" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">%</td>
	</tr>	
	
	<tr>
	
		
		<td colspan="6" class="formHeader"><input type="submit" value="Save Form" class="formMainBold"></td>
	</tr>
	<tr>
		<td colspan="6" class="formMain" align="center"><!--#include file="../includes/contact_info.inc"--></td>
	</tr>

<% 
If say = "edit" Then
	GetBudget.Close
	Set GetBudget = Nothing
	Con.Close
	Set Con = Nothing
End If
 %>
</form>
</table>
</html>