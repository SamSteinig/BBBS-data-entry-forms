<% 
If Request("status") = "addNew" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Set RST = Server.CreateObject("ADODB.Recordset")
	Con.Open "BBBSAforms", "sa","12sist12"
	RST.Open "SELECT * FROM tbl_frmStaffTEMP", Con, 1, 3
	RST.AddNew
	RST("AgencyID") = Request("AgencyIDN")
	RST("Year") = Request("year")
	RST("ASPSessionID") = Session.SessionID
	RST("BirthYear") = Request("frmStaffBirthYear")
	RST("Position") = Request("frmStaffPosition")
	RST("Race") = Request("frmStaffRace")
	RST("Sex") = Request("frmStaffSex")
	RST("Time") = Request("frmStaffTime")
	RST("Education") = Request("frmStaffEducation")
	RST("MonthStart") = Request("frmStaffMonthStart")
	RST("YearStart") = Request("frmStaffYearStart")
	RST("MonthEnd") = Request("frmStaffMonthEnd")
	RST("HoursWeek") = Request("frmStaffHoursWeek")
	RST("YearlySalary") = FormatCurrency(Request("frmStaffYearlySalary"))
	RST("CreateDate") = Now
	RST.Update
	idn = RST("StaffID")
	form = "Staff"
	modtype = "new"
	Set RST = Nothing
	%>
	<!--include file="../includes/modify_stamp.asp"-->
	<%	
	Con.Close
	Set Con = Nothing
	say = "add"
ElseIf Request("status") = "saveAll" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmStaffTEMP WHERE AgencyID=" & Session("agencyidn") & " AND ASPSessionID=" & Session.SessionID
	Set GetStaff = Con.Execute(query)
	GetStaff.MoveFirst
	Do Until GetStaff.EOF
		Set RST = Server.CreateObject("ADODB.Recordset")
		RST.Open "SELECT * FROM tbl_frmStaff", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = GetStaff("AgencyID")
		RST("Year") = GetStaff("year")
		RST("BirthYear") = GetStaff("BirthYear")
		RST("Position") = GetStaff("Position")
		RST("Race") = GetStaff("Race")
		RST("Sex") = GetStaff("Sex")
		RST("Time") = GetStaff("Time")
		RST("Education") = GetStaff("Education")
		RST("MonthStart") = GetStaff("MonthStart")
		RST("YearStart") = GetStaff("YearStart")
		RST("MonthEnd") = GetStaff("MonthEnd")
		RST("HoursWeek") = GetStaff("HoursWeek")
		RST("YearlySalary") = FormatCurrency(GetStaff("YearlySalary"))
		RST("CreateDate") = Now
		RST.Update
		idn = RST("StaffID")
		Set RST = Nothing
		GetStaff.MoveNext
	Loop
	'GetStaff.MoveFirst
	'Do Until GetStaff.EOF
	

	'	GetStaff.MoveNext
	'Loop	
	GetStaff.Close
	Set GetStaff = Nothing
	%>
	<!--include file="../includes/modify_stamp.asp"-->
	<%	
	Con.Close
	Set Con = Nothing
	say = "thanks"
Else
	say = "form"
End If
 %>

<!--#include file="../includes/session_stamp.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Staff</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<script language="javascript">
	<!--
		function checkForInteger(valueToCheck)
		{
			
			/******************************************************************************
			This function checks to see if the user entered any nonnumeric character into the form field.
			It also removes single spaces and a few other whitespace-type characters.
			Note: This function only accepts whole numbers. It will not let you enter decimal points.
			******************************************************************************/

			var myRegularExpression = /\D/;  // contains any nonnumeric character???
			var replaceWhiteSpace = /\s/; // searches for any whitespace character
			var formField = valueToCheck; // passed in as parameter 1
			var newFormField = valueToCheck.replace(replaceWhiteSpace, ""); // remove any whitespace from the form entry and replace it with nothing
			var bContainsNonNumbers = myRegularExpression.test(newFormField); // check newFormField variable to see if it contains any nonnumeric character
			
			if(bContainsNonNumbers)
			{
				alert("Please make sure you have entered a whole number.\n We cannot process letters or words."); 
			} 
		}
		

		
	// -->
	
var myRegularExpression1 = /\D/;
	
function submitFormValidate(form)
{
	if(form.frmStaffBirthYear.options[0].selected)
	{
	form.frmStaffBirthYear.focus();
	alert("Please Enter Birth Year. This field is required.");
	return false;
	}
	else if(form.frmStaffPosition.options[0].selected)
	{
	form.frmStaffPosition.focus();
	alert("Please Enter Position. This field is required.");
	return false;
	}
	else if(form.frmStaffRace.options[0].selected)
	{
	form.frmStaffRace.focus();
	alert("Please Enter Race. This field is required.");
	return false;
	}
	else if((form.frmStaffSex[0].checked != true) && (form.frmStaffSex[1].checked != true))
	{
		alert("Please Enter Sex. This field is required.");
		return false;
	}
	else if((form.frmStaffTime[0].checked != true) && (form.frmStaffTime[1].checked != true))
	{
		alert("Please check Full Time or Part Time. This field is required.");
		return false;
	}
	else if(form.frmStaffEducation.options[0].selected)
	{
	form.frmStaffEducation.focus();
	alert("Please Enter Education. This field is required.");
	return false;
	}
	else if(form.frmStaffMonthStart.options[0].selected)
	{
	form.frmStaffMonthStart.focus();
	alert("Please Enter Month Start. This field is required.");
	return false;
	}
	else if(form.frmStaffYearStart.options[0].selected)
	{
	form.frmStaffYearStart.focus();
	alert("Please Enter Year Start. This field is required.");
	return false;
	}
	else if(form.frmStaffMonthEnd.options[0].selected)
	{
	form.frmStaffMonthEnd.focus();
	alert("Please Enter Month End. This field is required.");
	return false;
	}
	else if(myRegularExpression1.test(form.frmStaffHoursWeek.value))	
	{
		form.frmStaffHoursWeek.focus();
		alert((form.frmStaffHoursWeek.value) + " is invalid.");
		return false;
	}
	else if(myRegularExpression1.test(form.frmStaffYearlySalary.value))	
	{
		form.frmStaffYearlySalary.focus();
		alert((form.frmStaffYearlySalary.value) + " is invalid.");
		return false;
	}
	else if(form.frmStaffHoursWeek.value == "")	
	{
		form.frmStaffHoursWeek.focus();
		alert("This field is required. Please do not leave any fields blank.");
		return false;
	}
	else if(form.frmStaffYearlySalary.value == "")	
	{
		form.frmStaffYearlySalary.focus();
		alert("This field is required. Please do not leave any fields blank.");
		return false;
	}
	else
	{
		return true;
	}
}	
</script>

<!--#include file="../includes/top_nav_forms_yearly.inc"--><!-- include file has </head> and <body> tags --><br>     

<% If say = "thanks" Then %>
<center>
<div align="center">
<font class="formMain">
Thank you! Your information has been saved in the BBBS database.<br>
To choose another form, please select the form type from the choices above.
</font>
</div>
</center>

<% ElseIf say <> "thanks" Then  %>
<center>
<form name="frmStaff" action="staff_edit.asp?y=<%= Request("y") %>" method="post" onsubmit="return submitFormValidate(this)">
<!--#include file="../includes/form_stamp.asp"-->

<table border="1" cellpadding="2" cellspacing="0" bordercolordark="003063">
	<tr>
		<td colspan="2" align="center" class="formSubhead">BBBS - <%= y %> Annual Agency Information (AAI)</td>
	</tr>
	<tr>
		<td colspan="2" class="formHeader">STAFF</td>
	</tr>
	<tr>
		<td colspan="2" class="formSubhead">Please enter all staff members into the fields below. Click "save" when you have<br>completed each. If any saved information needs to be changed on the lines below,<br>simply click "delete" on that individual row and re-enter the correct information.</td>
	</tr>
	<tr>
		<td class="formMain">Birth Year:</td>
		<td align="right" valign="top" class="formMain" class="formMain">
			<select size="1" class="formMain" name="frmStaffBirthYear">
			<option value="bad" class="formMainCentered">(please select)</option>
			<% 
			birthYear = Year(Now) + 1
			Do Until birthYear = 1900
				birthYear = (birthYear - 1)
			 %>
			<option value="<%= birthYear %>" class="formMainCentered"><%= birthYear %></option>
			<% 
			Loop			
			%>
			</select>
		</td>
	</tr>
<% 
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
 %>
	<tr>
		<td class="formMain">Position:</td>
		<td align="right" valign="top" class="formMain">
			<select size="1" class="formMain" name="frmStaffPosition">
				<option value="bad" class="formMain">(please select)</option>
<% 
query = "SELECT code,position FROM tbl_StaffPosition ORDER BY code"
Set GetCode = Con.Execute(query)
Do Until GetCode.EOF
 %>
				<option value="<%= GetCode("code") %>" class="formMain"><%= GetCode("position") %></option>
<% 
	GetCode.MoveNext
Loop
GetCode.Close
Set GetCode = Nothing
 %>
			</select>
		</td>
	</tr>
	<tr>
		<td class="formMain">Race:</td>
		<td align="right" valign="top" class="formMain">
			<select size="1" class="formMain" name="frmStaffRace">
				<option value="bad" class="formMain">(please select)</option>
<% 
query = "SELECT code,race FROM tbl_StaffRace ORDER BY code"
Set GetCode = Con.Execute(query)
Do Until GetCode.EOF
 %>
				<option value="<%= GetCode("code") %>" class="formMain"><%= GetCode("race") %></option>
<% 
	GetCode.MoveNext
Loop
GetCode.Close
Set GetCode = Nothing
 %>
			</select>	
		</td>
	</tr>
	<tr>
		<td class="formMain">Sex:</td>
		<td align="right" valign="top" class="formMain">			
			<input type="radio" name="frmStaffSex" value="M">M
			<input type="radio" name="frmStaffSex" value="F">F
		</td>
	</tr>
	<tr>
		<td class="formMain">Time:</td>
		<td align="right" valign="top" class="formMain">
			<input type="radio" name="frmStaffTime" value="FT">Full Time
			<input type="radio" name="frmStaffTime" value="PT">Part Time
		</td>
	</tr>
	<tr>
		<td class="formMain">Education:</td>
		<td align="right" valign="top" class="formMain">
			<select size="1" class="formMain" name="frmStaffEducation">
				<option value="bad" class="formMain">(please select)</option>
<% 
query = "SELECT code,education FROM tbl_StaffEducation ORDER BY code"
Set GetCode = Con.Execute(query)
Do Until GetCode.EOF
 %>
				<option value="<%= GetCode("code") %>" class="formMain"><%= GetCode("education") %></option>
<% 
	GetCode.MoveNext
Loop
GetCode.Close
Set GetCode = Nothing
 %>
			</select>	
		</td>
	</tr>
<% 
Con.Close
Set Con = Nothing
 %>
	<tr>
		<td class="formMain">Month Start:</td>
		<td align="right" valign="top" class="formMain">
			<select size="1" class="formMain" name="frmStaffMonthStart">
			<option value="bad" class="formMain">(please select)</option>
			<option value="1" class="formMain">January</option>
			<option value="2" class="formMain">February</option>
			<option value="3" class="formMain">March</option>
			<option value="4" class="formMain">April</option>
			<option value="5" class="formMain">May</option>
			<option value="6" class="formMain">June</option>
			<option value="7" class="formMain">July</option>
			<option value="8" class="formMain">August</option>
			<option value="9" class="formMain">September</option>
			<option value="10" class="formMain">October</option>
			<option value="11" class="formMain">November</option>
			<option value="12" class="formMain">December</option>
			</select>
		</td>
	</tr>
	<tr>
		<td class="formMain">Year Start:</td>
		<td align="right" valign="top" class="formMain">
			<select size="1" class="formMain" name="frmStaffYearStart">
			<option value="bad" class="formMain">(please select)</option>
			<% 
			yearStart = Year(Now) + 1
			Do Until yearStart = 1960
				yearStart = (yearStart - 1)
			 %>
			<option value="<%= yearStart %>" class="formMain"><%= yearStart %></option>
			<% 
			Loop			
			%>
			</select>
		</td>
	</tr>
	<tr>
		<td class="formMain">Month End:</td>
		<td align="right" valign="top" class="formMain">
			<select size="1" class="formMain" name="frmStaffMonthEnd">
			<option value="bad" class="formMain">(please select)</option>
			<option value="0" class="formMain">Still Employed</option>
			<option value="1" class="formMain">January</option>
			<option value="2" class="formMain">February</option>
			<option value="3" class="formMain">March</option>
			<option value="4" class="formMain">April</option>
			<option value="5" class="formMain">May</option>
			<option value="6" class="formMain">June</option>
			<option value="7" class="formMain">July</option>
			<option value="8" class="formMain">August</option>
			<option value="9" class="formMain">September</option>
			<option value="10" class="formMain">October</option>
			<option value="11" class="formMain">November</option>
			<option value="12" class="formMain">December</option>
			</select>
		</td>
	</tr>
	<tr>
		<td class="formMain">Hours per Week:</td>
		<td align="right" valign="top" class="formMain"><input type="text" size="2" maxlength="3" value="00" class="formMainRightJ" name="frmStaffHoursWeek" onblur="checkForInteger(this.value)"></td>
	</tr>
	<tr>
		<td class="formMain">Yearly Salary:</td>
		<td align="right" valign="top" class="formMain">&nbsp;$&nbsp;<input type="text" size="7" maxlength="10" value="0000" class="formMainRightJ" name="frmStaffYearlySalary" onblur="checkForInteger(this.value)"></td>
	</tr>
	<tr>
		<td colspan="2" class="formHeader"><input type="submit" value="Add Staff Member" class="formMainBold"></td>
	</tr>

</table>
<input type="hidden" name="status" value="addNew">
</form>
<% End If %>


<% 
If say = "add" Then
 %>
<form action="staff_edit.asp" method="post">
<!--#include file="../includes/form_stamp.asp"-->
<input type="hidden" name="status" value="saveAll">

<!-- RESULTS TABLE STARTS HERE -->
		<table border="1" cellpadding="2" cellspacing="0" bordercolordark="003063">
<!-- first row of table headers -->
			<tr>
				<td rowspan="2" class="formHeaderSmall">#</td>
				<td class="formHeaderSmall">Birth Year:</td>
				<td class="formHeaderSmall">Position:</td>
				<td class="formHeaderSmall">Race:</td>
				<td class="formHeaderSmall">Sex:</td>
				<td class="formHeaderSmall">Time:</td>
				<td class="formHeaderSmall">Education:</td>
				<td rowspan="2" class="formHeaderSmall">Edit/Delete Row</td>
			</tr>
<!-- second row of table headers -->
			<tr>
				<td class="formHeaderSmall">Month Start:</td>
				<td class="formHeaderSmall">Year Start:</td>
				<td class="formHeaderSmall">Month End:</td>
				<td class="formHeaderSmall">Hrs/Wk:</td>
				<td colspan="2" class="formHeaderSmall">Yearly Salary:</td>
				
			</tr>
					<%
						ct = 1
						Set Con = Server.CreateObject("ADODB.Connection")
						Con.Open "BBBSAforms", "sa","12sist12"
						query = "SELECT * FROM tbl_frmStaffTEMP WHERE AgencyID=" & Session("agencyidn") & " AND ASPSessionID=" & Session.SessionID
						Set GetStaff = Con.Execute(query)
						GetStaff.MoveFirst
						Do Until GetStaff.EOF
					 %>
<!-- first row of results -->
			<tr>
				<td rowspan="2" class="formMain"><%= ct %></td>
				<td class="formMain" align="center"><%= GetStaff("BirthYear") %></td>
					<% 
					query = "SELECT position FROM tbl_StaffPosition WHERE code=" & Int(GetStaff("position"))
					Set GetCode = Con.Execute(query)
					 %>
					<td class="formMain"><%= GetCode("position") %></td>
					<% 
					GetCode.Close
					Set GetCode = Nothing
					 %>
					<% 
					query = "SELECT race FROM tbl_StaffRace WHERE code=" & Int(GetStaff("race"))
					Set GetCode = Con.Execute(query)
					 %>
				<td class="formMain"><%= GetCode("race") %></td>
					<% 
					GetCode.Close
					Set GetCode = Nothing
					 %>
				<td class="formMain" align="center"><%= UCase(GetStaff("sex")) %></td>
				<td class="formMain" align="center"><%= UCase(GetStaff("time")) %></td>
					<% 
					query = "SELECT education FROM tbl_StaffEducation WHERE code=" & Int(GetStaff("education"))
					Set GetCode = Con.Execute(query)
					 %>
				<td class="formMain"><%= GetCode("education") %></td>
				<td align="right" class="formMain">Edit Row</td>				
			</tr>	
					<% 
					GetCode.Close
					Set GetCode = Nothing
					 %>
			
			<tr>
				<td class="formMain" align="center"><%= MonthName(GetStaff("Monthstart")) %></td>
				<td class="formMain" align="center"><%= GetStaff("yearstart") %></td>
				<td class="formMain" align="center"><% If GetStaff("monthend") = 0 Then %>Still Employed<% Else %><%= MonthName(GetStaff("monthend")) %><% End If %></td>
				<td class="formMain" align="center"><%= GetStaff("hoursweek") %></td>
				<td colspan="2" class="formMainRightJ"><%= FormatCurrency(GetStaff("yearlysalary")) %></td>
				<td align="right" class="formMain">Delete Row</td>
			</tr>
			<tr>
                <td colspan="8" class="formHeader"><img src="../images/spacer.gif" width="1" height="5" alt="" border="0"></td>
       		</tr>

					<% 
							GetStaff.MoveNext
							ct = ct + 1
						Loop
						GetStaff.Close
						Set GetStaff = Nothing
						Con.Close
						Set Con = Nothing
					 %>
			<tr>
				<td colspan="8" class="formHeader"><input type="submit" value="Save Form" class="formMainBold"></td>
			</tr>
		</table>
<P>
</form>
<% 
End If
 %>

</center>
</body>
</html>
