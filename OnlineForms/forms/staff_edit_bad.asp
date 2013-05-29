
<!--#include file="../includes/session_stamp.asp"-->

<% 

' check to see if user has rights to view and edit this form
if not Session("staffFormAccess") then %> 

	<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

		<html>
		<head>
			<title>Untitled</title>
		<link rel="stylesheet" type="text/css" href="../../bbbsa.css">
</head>
		
		<body>
			<p align="center"><br><br><b>You do not have access to view this form.<br><br><br>
			<a href="javascript: history.back()">back</a></p>
		</body>
		</html> <%
	
	response.end
end if


If Request("status") = "addNew" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Set RST = Server.CreateObject("ADODB.Recordset")
	Con.Open "BBBSAforms", "sa","12sist12"
	RST.Open "SELECT * FROM tbl_frmStaff", Con, 1, 3
	RST.AddNew
	RST("AgencyID") = Request("AgencyIDN")
	RST("Year") = Request("year")
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
	Set RST = Nothing
	form = "Staff"
	modtype = "new"
	%>
	<!--#include file="../includes/modify_stamp.asp"-->
	<%	
	Con.Close
	Set Con = Nothing
	say = "add"
ElseIf Request("status") = "deleteRow" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Set RST = Server.CreateObject("ADODB.Recordset")
	Con.Open "BBBSAforms", "sa","12sist12"
	RST.Open "SELECT * FROM tbl_frmStaff WHERE StaffID=" & Int(Request("row")), Con, 1, 3
	jMod = RST("StaffID")
	RST.Delete
	RST.Update
	Set RST = Nothing
	form = "Staff"
	modtype = "delete"
	%>
	<!--#include file="../includes/modify_stamp.asp"-->
	<%	
	Con.Close
	Set Con = Nothing
	say = "delete"
ElseIf Request("status") = "editRow" Then
	say = "edit"
ElseIf Request("status") = "editSave" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Set RST = Server.CreateObject("ADODB.Recordset")
	Con.Open "BBBSAforms", "sa","12sist12"
	RST.Open "SELECT * FROM tbl_frmStaff WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & " AND StaffID=" & Int(Request("row")), Con, 1, 3
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
	jMod = RST("StaffID")
	RST.Update
	Set RST = Nothing
	form = "Staff"
	modtype = "edit"
	%>
	<!--#include file="../includes/modify_stamp.asp"-->
	<%	
	Con.Close
	Set Con = Nothing
	say = "add"
ElseIf Request("status") = "done" Then
	say = "thanks"
ElseIf Request("status") = "newStaff" Then
	say = "form"
Else
	say = "form"
End If
 %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Staff</title>
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
		alert("Please enter gender. This field is required.");
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
	else if(form.frmStaffHoursWeek.value == "")	
	{
		form.frmStaffHoursWeek.focus();
		alert("This field is required. Please do not leave any fields blank.");
		return false;
	}
	else if(form.frmStaffHoursWeek.value == "0")	
	{
		form.frmStaffHoursWeek.focus();
			alert("Hours per Week cannot equal zero.");
		return false;
	}
	else if(form.frmStaffHoursWeek.value == "")
	{	
		alert("This field is required. Please do not leave any fields blank.");
		return false;
	}
	else if(form.frmStaffYearlySalary.value == "")	
	{
		form.frmStaffYearlySalary.focus();
		alert("Yearly Salary is required. Please do not leave any fields blank.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmStaffYearlySalary.value)))	
	{
		form.frmStaffYearlySalary.focus();
		alert((form.frmStaffYearlySalary.value) + " is invalid.");
		return false;
	}
	else
	{
		return true;
	}
}	
// -->
</script>

<!--#include file="../includes/top_nav_forms_yearly.inc"--><!-- include file has </head> and <body> tags --><br>     

<% If say = "thanks" Then %>
<center>
<div align="center">
<font class="formMain">
Thank you! Your information has been saved in the BBBS database.<br>
To choose another form, please select the form type from the choices above.
</font>

<br>
<!--#include file="../includes/contact_info.inc"-->
<br>

</div>
</center>

<% ElseIf say <> "thanks" Then  %>
<center>
<form name="frmStaff" action="staff_edit.asp?y=<%= Request("y") %>" method="post" onsubmit="return submitFormValidate(this)">
<!--#include file="../includes/form_stamp.asp"-->
<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmStaff WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & " AND StaffID=" & Int(Request("row"))
	Set GetStaff = Con.Execute(query)
 %>
<input type="hidden" name="status" value="editSave">
<input type="hidden" name="row" value="<%= Request("row") %>">
<% Else %>
<input type="hidden" name="status" value="addNew">
<%
End If
 %>

<table border="1" cellpadding="2" cellspacing="0" bordercolordark="003063">
	<tr>
		<td colspan="2" align="center" class="formSubhead">BBBS - <%= y %> Annual Agency Information (AAI)</td>
	</tr>
	<tr>
		<td colspan="2" class="formHeader">STAFF</td>
	</tr>
	<tr>
		<td colspan="2" align="center" valign="top" class="formMain">Please enter the following information on each staff member into the fields below.<br>Click "Add Staff Member" when you have completed each. Saved information will appear in a grid below.</td>
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
			<option value="<%= birthYear %>" selected<% End If %><% End If %> class="formMainCentered" <% If say = "edit" Then %><% If birthYear = GetStaff("BirthYear") Then %>><%= birthYear %></option>
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
				<option value="<%= GetCode("code") %>" selected<% End If %><% End If %> class="formMain" <% If say = "edit" Then %><% If GetCode("code") = GetStaff("Position") Then %>><%= GetCode("position") %></option>
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
				<option value="<%= GetCode("code") %>" selected<% End If %><% End If %> class="formMain" <% If say = "edit" Then %><% If GetCode("code") = GetStaff("Race") Then %>><%= GetCode("race") %></option>
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
		<td class="formMain">Gender:</td>
		<td align="right" valign="top" class="formMain">			
			<input type="radio" name="frmStaffSex" value="M" checked<% End If %><% End If %> <% If say = "edit" Then %><% If Trim(GetStaff("Sex")) = "M" Then %>>M
			<input type="radio" name="frmStaffSex" value="F" checked<% End If %><% End If %> <% If say = "edit" Then %><% If Trim(GetStaff("Sex")) = "F" Then %>>F
		</td>
	</tr>
	<tr>
		<td class="formMain">Time:</td>
		<td align="right" valign="top" class="formMain">
			<input type="radio" name="frmStaffTime" value="FT" checked<% End If %><% End If %> <% If say = "edit" Then %><% If Trim(GetStaff("Time")) = "FT" Then %>>Full Time
			<input type="radio" name="frmStaffTime" value="PT" checked<% End If %><% End If %> <% If say = "edit" Then %><% If Trim(GetStaff("Time")) = "PT" Then %>>Part Time
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
				<option value="<%= GetCode("code") %>" selected<% End If %><% End If %> class="formMain" <% If say = "edit" Then %><% If GetCode("code") = GetStaff("Education") Then %>><%= GetCode("education") %></option>
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
			<option value="1" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthStart") = 1 Then %>>January</option>
			<option value="2" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthStart") = 2 Then %>>February</option>
			<option value="3" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthStart") = 3 Then %>>March</option>
			<option value="4" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthStart") = 4 Then %>>April</option>
			<option value="5" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthStart") = 5 Then %>>May</option>
			<option value="6" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthStart") = 6 Then %>>June</option>
			<option value="7" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthStart") = 7 Then %>>July</option>
			<option value="8" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthStart") = 8 Then %>>August</option>
			<option value="9" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthStart") = 9 Then %>>September</option>
			<option value="10" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthStart") = 10 Then %>>October</option>
			<option value="11" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthStart") = 11 Then %>>November</option>
			<option value="12" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthStart") = 12 Then %>>December</option>
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
			<option value="<%= yearStart %>" selected<% End If %><% End If %> class="formMain" <% If say = "edit" Then %><% If GetStaff("YearStart") = yearStart Then %>><%= yearStart %></option>
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
			<option value="0" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 0 Then %>>Still Employed</option>
			<option value="1" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 1 Then %>>January</option>
			<option value="2" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 2 Then %>>February</option>
			<option value="3" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 3 Then %>>March</option>
			<option value="4" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 4 Then %>>April</option>
			<option value="5" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 5 Then %>>May</option>
			<option value="6" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 6 Then %>>June</option>
			<option value="7" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 7 Then %>>July</option>
			<option value="8" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 8 Then %>>August</option>
			<option value="9" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 9 Then %>>September</option>
			<option value="10" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 10 Then %>>October</option>
			<option value="11" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 11 Then %>>November</option>
			<option value="12" class="formMain" selected<% End If %><% End If %> <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 12 Then %>>December</option>
			</select>
		</td>
	</tr>
	<tr>
		<td class="formMain">Hours per Week:</td>
		<td align="right" valign="top" class="formMain"><input type="text" size="4" value="<% If say = "edit" Then %><%= GetStaff("HoursWeek") %><% Else  %>0<% End If %>" class="formMain" name="frmStaffHoursWeek" onchange="checkForInteger(this.value)"></td>
	</tr>
	<tr>
		<td class="formMain">Yearly Salary:</td>
		<td align="right" valign="top" class="formMain">&nbsp;$&nbsp;<input type="text" size="7" maxlength="10" value="<% If say = "edit" Then %><%= GetStaff("YearlySalary") %><% Else  %>0<% End If %>" class="formMain" name="frmStaffYearlySalary" onchange="checkForInteger(this.value)"></td>
	</tr>
	<% If say = "edit" Then %>
	<tr>
		<td colspan="2" class="formHeader"><input type="submit" value="Edit Staff Member" class="formMainBold"></td>
	</tr>
	<% Else %>
	<tr>
		<td colspan="2" class="formHeader"><input type="submit" value="Add Staff Member" class="formMainBold"></td>
	</tr>
	<% End If %>
</table>
</form>
<% End If %>


<% 
If say <> "thanks" Then
 %>
<script language="JavaScript">
<!-- 
	function confirmDelete(row)
	{
		if (confirm("Are you sure you want to delete this record?"))
		{
			location.href = "staff_edit.asp?status=deleteRow&row=" + row + "&y=<%= Request("y") %>";
			// alert("Record deleted.");
		}
		else
		{
		return false;
		}
	}		
// -->
</script>	
<!-- RESULTS TABLE STARTS HERE -->
		<table border="1" cellpadding="2" cellspacing="0" bordercolordark="003063">
<!-- first row of table headers -->
			<tr>
				<td colspan="8" align="center" valign="top" class="formMain">If any of the following information needs to be changed or removed,<br> simply click "Edit Record" or "Delete Record" for that individual and re-enter their information.<br> When all staff members have been added, click "Done" to submit this form.</td>
			</tr>
			<tr>
				<td rowspan="2" class="formHeaderSmall">#</td>
				<td class="formHeaderSmall">Birth Year:</td>
				<td class="formHeaderSmall">Position:</td>
				<td class="formHeaderSmall">Race:</td>
				<td class="formHeaderSmall">Sex:</td>
				<td class="formHeaderSmall">Time:</td>
				<td class="formHeaderSmall">Education:</td>
				<td rowspan="2" class="formHeaderSmall">Edit/Delete</td>
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
						query = "SELECT * FROM tbl_frmStaff WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
						Set GetStaff = Con.Execute(query)
						If GetStaff.EOF OR GetStaff.BOF Then
					%>
					<tr>
	                <td colspan="8" class="formMainBold">No Staff Members To List</td>
    		   		</tr>
					<%
						Else
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
				<td rowspan="2" align="right" class="formMain"><a href="staff_edit.asp?status=editRow&row=<%= GetStaff("StaffID") %>&y=<%= Request("y") %>">Edit Record</a><br><a href="#" onclick="confirmDelete(<%= GetStaff("StaffID") %>); return false;">Delete Record</a></td>				
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
					End If
					 %>
<form name="frmStaff" action="staff_edit.asp?status=done" method="post">
			<tr>
                <td colspan="8" class="formHeader"><input type="submit" value="Done" class="formMainBold"></td>
       		</tr>
</form>
		</table><br>
<!--#include file="../includes/contact_info.inc"-->
<br>
<P>
<% 
End If
 %>
 
</div>
</center>

</body>
</html>
