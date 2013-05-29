<!--#include file="../includes/session_stamp.asp"-->
<% 

' check to see if user has rights to view and edit this form
if not Session("staffFormAccess") then %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>Untitled</title>
</head>
<body>
    <p align="center">
        <br>
        <br>
        <b>You do not have access to view this form.<br>
            <br>
            <br>
            <a href="javascript: history.back()">back</a></p>
</body>
</html>
<%
	
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
	RST("EverABig") = Request("frmStaffEverABig")
	RST("Position") = Request("frmStaffPosition")
	RST("Race") = Request("frmStaffRace")
	RST("Sex") = Request("frmStaffSex")
	RST("Time") = Request("frmStaffTime")
	RST("Education") = Request("frmStaffEducation")
	RST("MonthStart") = Request("frmStaffMonthStart")
	RST("YearStart") = Request("frmStaffYearStart")
	RST("YearsInNetwork") = Request("frmStaffYearsInNetwork")
	RST("MonthEnd") = Request("frmStaffTerminated")
	RST("HoursWeek") = Int(Request("frmStaffHoursWeek"))
	RST("YearlySalary") = FormatCurrency(Request("frmStaffYearlySalary"))
	RST("CreateDate") = Now
	RST("EmployeeName") = Request("frmStaffEmployeeName")
	If Len(Request("frmStaffPositionStartDate")) <> 0 Then 
		RST("PositionStartDate") = Request("frmStaffPositionStartDate")
	Else
		RST("PositionStartDate") = NULL
	End If
	RST("BaseSalary") = FormatCurrency(Request("frmStaffBaseSalary"))
	RST("BonusSalary") = FormatCurrency(Request("frmStaffBonusSalary"))
	RST("FTE") = Request("frmStaffFTE")
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
	display = "showSummary"
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
	display = "showSummary"
ElseIf Request("status") = "editRow" Then
	say = "edit"
ElseIf Request("status") = "editSave" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Set RST = Server.CreateObject("ADODB.Recordset")
	Con.Open "BBBSAforms", "sa","12sist12"
	RST.Open "SELECT * FROM tbl_frmStaff WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & " AND StaffID=" & Int(Request("row")), Con, 1, 3
	RST("BirthYear") = Request("frmStaffBirthYear")
	RST("EverABig") = Request("frmStaffEverABig")	
	RST("Position") = Request("frmStaffPosition")
	RST("Race") = Request("frmStaffRace")
	RST("Sex") = Request("frmStaffSex")
	RST("Time") = Request("frmStaffTime")
	RST("Education") = Request("frmStaffEducation")
	RST("MonthStart") = Request("frmStaffMonthStart")
	RST("YearStart") = Request("frmStaffYearStart")
	RST("YearsInNetwork") = Request("frmStaffYearsInNetwork")
	RST("MonthEnd") = Request("frmStaffTerminated")
	RST("HoursWeek") = Request("frmStaffHoursWeek")
	RST("YearlySalary") = FormatCurrency(Request("frmStaffYearlySalary"))
	RST("EmployeeName") = Request("frmStaffEmployeeName")
	If Len(Request("frmStaffPositionStartDate")) <> 0 Then 
		RST("PositionStartDate") = Request("frmStaffPositionStartDate")
	Else
		RST("PositionStartDate") = NULL
	End If
	RST("BaseSalary") = FormatCurrency(Request("frmStaffBaseSalary"))
	RST("BonusSalary") = FormatCurrency(Request("frmStaffBonusSalary"))
	If LEN(Request("frmStaffFTE")) = 0 Then
		RST("FTE") = NULL
	Else
		RST("FTE") = Request("frmStaffFTE")
	End If
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
	display = "showSummary"
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
function NewWindow(mypage, myname, w, h)
{
	var winl = (screen.width - w) / 2;
	var wint = (screen.height - h) / 2;
	winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',resizable, scrollbars'
	win = window.open(mypage, myname, winprops)
	if (parseInt(navigator.appVersion) >= 4) { win.window.focus(); }
}

// Data Validation code
function chkdate(objName) {
var strDatestyle = "US"; //United States date style
//var strDatestyle = "EU";  //European date style
var strDate;
var strDateArray;
var strDay;
var strMonth;
var strYear;
var intday;
var intMonth;
var intYear;
var booFound = false;
var datefield = objName;
var strSeparatorArray = new Array("-"," ","/",".");
var intElementNr;
var err = 0;
var strMonthArray = new Array(12);
strMonthArray[0] = "Jan";
strMonthArray[1] = "Feb";
strMonthArray[2] = "Mar";
strMonthArray[3] = "Apr";
strMonthArray[4] = "May";
strMonthArray[5] = "Jun";
strMonthArray[6] = "Jul";
strMonthArray[7] = "Aug";
strMonthArray[8] = "Sep";
strMonthArray[9] = "Oct";
strMonthArray[10] = "Nov";
strMonthArray[11] = "Dec";
strDate = datefield.value;
if (strDate.length < 1) {
return true;
}
for (intElementNr = 0; intElementNr < strSeparatorArray.length; intElementNr++) {
if (strDate.indexOf(strSeparatorArray[intElementNr]) != -1) {
strDateArray = strDate.split(strSeparatorArray[intElementNr]);
if (strDateArray.length != 3) {
err = 1;
return false;
}
else {
strDay = strDateArray[0];
strMonth = strDateArray[1];
strYear = strDateArray[2];
}
booFound = true;
   }
}
if (booFound == false) {
if (strDate.length>5) {
strDay = strDate.substr(0, 2);
strMonth = strDate.substr(2, 2);
strYear = strDate.substr(4);
   }
}
if (strYear.length == 2) {
strYear = '20' + strYear;
}
// US style
if (strDatestyle == "US") {
strTemp = strDay;
strDay = strMonth;
strMonth = strTemp;
}
intday = parseInt(strDay, 10);
if (isNaN(intday)) {
err = 2;
return false;
}
intMonth = parseInt(strMonth, 10);
if (isNaN(intMonth)) {
for (i = 0;i<12;i++) {
if (strMonth.toUpperCase() == strMonthArray[i].toUpperCase()) {
intMonth = i+1;
strMonth = strMonthArray[i];
i = 12;
   }
}
if (isNaN(intMonth)) {
err = 3;
return false;
   }
}
intYear = parseInt(strYear, 10);
if (isNaN(intYear)) {
err = 4;
return false;
}
if (intMonth>12 || intMonth<1) {
err = 5;
return false;
}
if ((intMonth == 1 || intMonth == 3 || intMonth == 5 || intMonth == 7 || intMonth == 8 || intMonth == 10 || intMonth == 12) && (intday > 31 || intday < 1)) {
err = 6;
return false;
}
if ((intMonth == 4 || intMonth == 6 || intMonth == 9 || intMonth == 11) && (intday > 30 || intday < 1)) {
err = 7;
return false;
}
if (intMonth == 2) {
if (intday < 1) {
err = 8;
return false;
}
if (LeapYear(intYear) == true) {
if (intday > 29) {
err = 9;
return false;
}
}
else {
if (intday > 28) {
err = 10;
return false;
}
}
}
/*if (strDatestyle == "US") {
datefield.value = strMonthArray[intMonth-1] + " " + intday+" " + strYear;
}
else {
datefield.value = intday + " " + strMonthArray[intMonth-1] + " " + strYear;
}
*/
return true;
}
function LeapYear(intYear) {
if (intYear % 100 == 0) {
if (intYear % 400 == 0) { return true; }
}
else {
if ((intYear % 4) == 0) { return true; }
}
return false;
}
function doDateCheck(from, to) {
if (Date.parse(from.value) <= Date.parse(to.value)) {
alert("The dates are valid.");
}
else {
if (from.value == "" || to.value == "") 
alert("Both dates must be entered.");
else 
alert("To date must occur after the from date.");
   }
}

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

function checkForNumeric(valueToCheck)
{
	var myRegularExpression = /^\d*\.?\d*$/;  	// contains an int, double or float
	var replaceWhiteSpace = /\s/; // searches for any whitespace character
	var formField = valueToCheck; // passed in as parameter 1
	var newFormField = valueToCheck.replace(replaceWhiteSpace, ""); // remove any whitespace from the form entry and replace it with nothing
	var bContainsNonNumbers = myRegularExpression.test(newFormField); // check newFormField variable to see if it contains any nonnumeric character
	
	if(!bContainsNonNumbers)
	{
		alert("Please make sure you have entered a number.\n We cannot process letters or words."); 
	} 
}
function addSalary(form)
{
	var box1 = Number(form.frmStaffBaseSalary.value)
	var box2 = Number(form.frmStaffBonusSalary.value)
	
	var boxtotal = box1 + box2
	form.frmStaffYearlySalary.value = boxtotal
}

function enableFTE(form)
{
	form.frmStaffFTE.readonly = false;
}
function disableFTE(form)
{
	form.frmStaffFTE.value = 1.0
	form.frmStaffFTE.readonly = true;
}


function enableTerminated(form)
{
	form.frmStaffTerminated.readonly = false;
}
function disableTerminated(form)
{
	form.frmStaffTerminated.value = 1.0
	form.frmStaffTerminated.readonly = true;
}
/*function changeFTE(form)
{
	if(form.frmStaffTime[0].checked == true)
	{
		form.frmStaffFTE.value = 1;
		document.form.frmStaffFTE.disabled = true;
	}
	if(form.frmStaffTime[1].checked == true)
	{
		form.frmStaffFTE.value = 2;
		form.frmStaffFTE.disabled = false;
	}
}*/
		

var myRegularExpression1 = /^[0-9]+(,[0-9]{3})*$/;  // Checks for integer with or without commas
var myRegularExpression2 = /^\d*\.?\d*$/;  	// contains an int, double or float
var myRegularExpression3 = 	/^\d*$/;  // Checks for integer 
	
function submitFormValidate(form)
{
	var myDateString = form.frmStaffPositionStartDate.value;//12/21/2006 saf
	if(form.frmStaffBirthYear.options[0].selected)
	{
	    form.frmStaffBirthYear.focus();
	    alert("Please Enter Birth Year. This field is required.");
	    return false;
	}
	else if((form.frmStaffSex[0].checked != true) && (form.frmStaffSex[1].checked != true))
	{
		alert("Please enter gender. This field is required.");
		return false;
	}
	else if(form.frmStaffRace.options[0].selected)
	{
	    form.frmStaffRace.focus();
	    alert("Please Enter Race. This field is required.");
	    return false;
	}
	else if(form.frmStaffEducation.options[0].selected)
	{
	    form.frmStaffEducation.focus();
	    alert("Please Enter Education. This field is required.");
	    return false;
	}
	else if(form.frmStaffPosition.options[0].selected)
	{
	    form.frmStaffPosition.focus();
	    alert("Please Enter Position. This field is required.");
	    return false;
	}
    //Added 12/21/2006 saf to ensure Position Date is valid
	//else if (!chkdate(form.frmStaffPositionStartDate.value) || form.frmStaffPositionStartDate.value == "")
	else if (form.frmStaffPositionStartDate.value == "") 
	{ 
		form.frmStaffPositionStartDate.focus();
	    alert("Please Enter the Position Start date (MM/DD/YYYY). This field is required."); 
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
	
	else if(new Number(form.frmStaffHoursWeek.value) < 35 && form.frmStaffTime[0].checked == true || form.frmStaffHoursWeek.value > 40)
		{
		alert("Hours per week for a Full-Time employee must be Between 35 and 40.");
		return false;
	}
	else if(new Number(form.frmStaffHoursWeek.value) > 34 && form.frmStaffTime[1].checked == true)
	{
		alert("Hours per week for a Part-Time employee must be less than 35.");
		return false;
	}
	else if(form.frmStaffMonthStart.options[0].selected)
	{
	    form.frmStaffMonthStart.focus();
	    alert("Please Enter Month Hired. This field is required.");
	    return false;
	}
	else if(form.frmStaffYearStart.options[0].selected)
	{
	    form.frmStaffYearStart.focus();
	    alert("Please Enter Year Start. This field is required.");
	    return false;
	}
	else if(form.frmStaffYearsInNetwork.options[0].selected)
	{
	    form.frmStaffYearsInNetwork.focus();
	    alert("Please Enter Total Years in BBBS Network. This field is required.");
	    return false;
	}
	else if((form.frmStaffTime[0].checked != true) && (form.frmStaffTime[1].checked != true))
	{
		alert("Please check Full Time or Part Time. This field is required.");
		return false;
	}
	else if((form.frmStaffTime[1].checked == true) && (form.frmStaffFTE.value.length < 1))
	{
		form.frmStaffFTE.focus();
		alert("This field is required. Please do not leave any fields blank.");
		return false;
	}
	
	else if(form.frmStaffTerminated.options[0].selected)
	{
		form.frmStaffTerminated.focus();
		alert("Please Enter Employment Status. This field is required.");
		return false;
	}
	else if(form.frmStaffEmployeeName.value == "")	
	{
		form.frmStaffEmployeeName.focus();
		alert("Please Enter Employee Name. This field is required.");
		return false;
	}

	else if(form.frmStaffBaseSalary.value == "0" || form.frmStaffBaseSalary.value == "")	
	{
		form.frmStaffBaseSalary.focus();
		alert("Please Enter Base Salary. This field is required.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmStaffBaseSalary.value)))	
	{
		form.frmStaffBaseSalary.focus();
		alert((form.frmStaffBaseSalary.value) + " is invalid.");
		return false;
	}
	
	else if(form.frmStaffBonusSalary.value == "")	
	{
		form.frmStaffBonusSalary.focus();
		alert("Please Enter Bonus/Incentive Salary. If employee does not have bonus - enter zero(0).");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmStaffBonusSalary.value)))	
	{
		form.frmStaffBonusSalary.focus();
		alert((form.frmStaffBonusSalary.value) + " is invalid.");
		return false;
	}

	/*else if((Number(form.frmStaffBaseSalary.value)+ Number(form.frmStaffBonusSalary.value)) < 10000 && form.frmStaffTime[0].checked == true && !(form.frmStaffPosition.options.value == 7 || form.frmStaffPosition.options.value == 90 || form.frmStaffPosition.options.value == 99))
	{
		form.frmStaffBaseSalary.focus();	
		alert("Yearly salary for a Full-Time employee must be at least $10,000.");
		return false;
	}*/
	
	else
	{
		// This insert compares position with base salary ranges (validation) //slp 06/16/2008
		if((form.frmStaffPosition.options[1].selected) && (Number(form.frmStaffBaseSalary.value)<35000 ||  Number(form.frmStaffBaseSalary.value)>190000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[2].selected) && (Number(form.frmStaffBaseSalary.value)<35000 ||  Number(form.frmStaffBaseSalary.value)>190000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[3].selected) && (Number(form.frmStaffBaseSalary.value)<40000 ||  Number(form.frmStaffBaseSalary.value)>125000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[4].selected) && (Number(form.frmStaffBaseSalary.value)<30000 ||  Number(form.frmStaffBaseSalary.value)>130000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[5].selected) && (Number(form.frmStaffBaseSalary.value)<22000 ||  Number(form.frmStaffBaseSalary.value)>70000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[6].selected) && (Number(form.frmStaffBaseSalary.value)<11000 ||  Number(form.frmStaffBaseSalary.value)>60000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[7].selected) && (Number(form.frmStaffBaseSalary.value)<30000 ||  Number(form.frmStaffBaseSalary.value)>80000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[10].selected) && (Number(form.frmStaffBaseSalary.value)<35000 ||  Number(form.frmStaffBaseSalary.value)>136000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[11].selected) && (Number(form.frmStaffBaseSalary.value)<30000 ||  Number(form.frmStaffBaseSalary.value)>68000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[12].selected) && (Number(form.frmStaffBaseSalary.value)<18000 ||  Number(form.frmStaffBaseSalary.value)>45000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[14].selected) && (Number(form.frmStaffBaseSalary.value)<18000 ||  Number(form.frmStaffBaseSalary.value)>45000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[15].selected) && (Number(form.frmStaffBaseSalary.value)<22000 ||  Number(form.frmStaffBaseSalary.value)>55000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[16].selected) && (Number(form.frmStaffBaseSalary.value)<13000 ||  Number(form.frmStaffBaseSalary.value)>60000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[17].selected) && (Number(form.frmStaffBaseSalary.value)<30000 ||  Number(form.frmStaffBaseSalary.value)>50000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[18].selected) && (Number(form.frmStaffBaseSalary.value)<30000 ||  Number(form.frmStaffBaseSalary.value)>55000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[19].selected) && (Number(form.frmStaffBaseSalary.value)<30000 ||  Number(form.frmStaffBaseSalary.value)>55000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[20].selected) && (Number(form.frmStaffBaseSalary.value)<40000 ||  Number(form.frmStaffBaseSalary.value)>150000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[22].selected) && (Number(form.frmStaffBaseSalary.value)<30000 ||  Number(form.frmStaffBaseSalary.value)>85000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[23].selected) && (Number(form.frmStaffBaseSalary.value)<30000 ||  Number(form.frmStaffBaseSalary.value)>80000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[24].selected) && (Number(form.frmStaffBaseSalary.value)<30000 ||  Number(form.frmStaffBaseSalary.value)>80000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[25].selected) && (Number(form.frmStaffBaseSalary.value)<40000 ||  Number(form.frmStaffBaseSalary.value)>95000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[30].selected) && (Number(form.frmStaffBaseSalary.value)<35000 ||  Number(form.frmStaffBaseSalary.value)>65000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[32].selected) && (Number(form.frmStaffBaseSalary.value)<17000 ||  Number(form.frmStaffBaseSalary.value)>55000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[40].selected) && (Number(form.frmStaffBaseSalary.value)<35000 ||  Number(form.frmStaffBaseSalary.value)>95000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[41].selected) && (Number(form.frmStaffBaseSalary.value)<36000 ||  Number(form.frmStaffBaseSalary.value)>60000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[42].selected) && (Number(form.frmStaffBaseSalary.value)<17000 ||  Number(form.frmStaffBaseSalary.value)>55000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[50].selected) && (Number(form.frmStaffBaseSalary.value)<30000 ||  Number(form.frmStaffBaseSalary.value)>80000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[51].selected) && (Number(form.frmStaffBaseSalary.value)<35000 ||  Number(form.frmStaffBaseSalary.value)>50000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[52].selected) && (Number(form.frmStaffBaseSalary.value)<16000 ||  Number(form.frmStaffBaseSalary.value)>60000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[60].selected) && (Number(form.frmStaffBaseSalary.value)<45000 ||  Number(form.frmStaffBaseSalary.value)>80000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[61].selected) && (Number(form.frmStaffBaseSalary.value)<40000 ||  Number(form.frmStaffBaseSalary.value)>65000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[62].selected) && (Number(form.frmStaffBaseSalary.value)<12000 ||  Number(form.frmStaffBaseSalary.value)>70000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[70].selected) && (Number(form.frmStaffBaseSalary.value)<45000 ||  Number(form.frmStaffBaseSalary.value)>65000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[71].selected) && (Number(form.frmStaffBaseSalary.value)<40000 ||  Number(form.frmStaffBaseSalary.value)>65000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[72].selected) && (Number(form.frmStaffBaseSalary.value)<10000 ||  Number(form.frmStaffBaseSalary.value)>30000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[80].selected) && (Number(form.frmStaffBaseSalary.value)<20000 ||  Number(form.frmStaffBaseSalary.value)>60000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[81].selected) && (Number(form.frmStaffBaseSalary.value)<25000 ||  Number(form.frmStaffBaseSalary.value)>80000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[90].selected) && (Number(form.frmStaffBaseSalary.value)<10000 ||  Number(form.frmStaffBaseSalary.value)>100000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[99].selected) && (Number(form.frmStaffBaseSalary.value)<3000 ||  Number(form.frmStaffBaseSalary.value)>20000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[73].selected) && (Number(form.frmStaffBaseSalary.value)<30000 ||  Number(form.frmStaffBaseSalary.value)>60000))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		else if((form.frmStaffPosition.options[100].selected) && (Number(form.frmStaffBaseSalary.value)= 0 ))
		{
		    form.frmStaffPosition.focus();
		    return confirm("The salary you entered ( $" + form.frmStaffBaseSalary.value + " ) for selected position is outside of normal range. If this number is correct click OK, if not click CANCEL to re enter the salary.");
		}
		// End of comparing position with base salary ranges (validation)
		else
		{
		    var Base = Number(form.frmStaffBaseSalary.value);
		    var Bonus = Number(form.frmStaffBonusSalary.value);
		    var Yearly = Base + Bonus;
		    form.frmStaffYearlySalary.value = Yearly;
		    return true;
		}
	}
}	
// -->
    </script>

    <script language="javascript" type="text/javascript" src="datetimepicker.js"></script>

    </script>
    <% '<!--#include file="../includes/top_nav_forms_yearly.inc"--><!-- include file has </head> and <body> tags --><br>     %>
    <!--#include file="../includes/surveytitle.inc"-->
    <% If display = "showSummary" Then %>
    <% Response.Redirect("Staff_complete.asp?y=" & Request("y") & "&id=" & Request("AgencyIDN")) %><%End If%>
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
            <td width="220" valign="top">
                <img src="../includes/images/photos_wheelbarrow.jpg" alt="" width="220" height="477"
                    border="0">
            </td>
            <td valign="top">
                <br>
                <% If say = "thanks" Then %>
                <font class="formMain">
                    <br>
                    <br>
                    <strong>Thank you!</strong> Your information has been saved in the BBBS database.<br>
                    To choose another form, please select the form type from the choices above.
                    <br>
                    <br>
                    <i>Please note: These changes will not be reflected in the <strong>Agency Profile</strong>
                        (in the My Agency Page and the Agency Directory) for 24 hours.</i> </font>
                <br>
                <!--#include file="../includes/contact_info.inc"-->
                <br>
                <% ElseIf say <> "thanks" Then  %>
                <form name="frmStaff" action="staff_edit.asp?y=<%= Request("y") %>" method="post"
                onsubmit="return submitFormValidate(this)">
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
                <% If display <> "showSummary" Then %>
                <table border="1" cellpadding="2" cellspacing="0" bordercolordark="003063" width="640">
                    <tr>
                        <td colspan="2" align="center" class="formSubhead">
                            BBBS -
                            <%= y %>
                            Annual Agency Information (AAI)
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" class="formHeader">
                            STAFF
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" align="center" valign="top" class="formMain">
                            <font color="#ff0000">Please enter the following information on each staff member into
                                the fields below.<br>
                                Click "Save This Entry" when you have completed each. Saved information will appear
                                in a grid below.</font>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3" class="formMain">
                            <font color="#ff0000">
                                <div align="center">
                                    If you need help with understanding the topic, please click on
                                        <img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0">
                                        next to the topic of your interest.<br><strong>
                                        The data you enter should reflect your employee data on 6/30/09.</strong>
                            </font>
                        </td>
                    </tr>
                    <%'added 12/20/2006 saf to capture Employee name %>
                    <tr>
                        <td class="formMain">
                            <a href="../helpfiles/StaffFormHelp.asp?HelpID=EmployeeName" onclick="NewWindow(this.href,'name','600','450','yes');return false;">
                                <img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
                            Employee Name:
                        </td>
                        <td align="right" valign="top" class="formMain">
                            <% if say = "add" or say = "form" or say = "delete" then %>
                            <input type="text" size="70" maxlength="70" value="<% If say = "edit" Then %><%= GetStaff("EmployeeName") %><% Else  %><% End If %>"
                                class="formMain" name="frmStaffEmployeeName">
                            <% else %>
                            <% if say = "edit" and not isnull(GetStaff("EmployeeName")) then %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<% end if %><input
                                type="text" size="60" maxlength="600" value="<% If say = "edit" Then %><%= GetStaff("EmployeeName") %><% Else  %><% End If %>"
                                class="formMain" name="frmStaffEmployeeName">
                            <% end if %>
                        </td>
                    </tr>
                    <tr>
                        <td class="formMain">
                            <a href="../helpfiles/StaffFormHelp.asp?HelpID=BirthYear" onclick="NewWindow(this.href,'name','600','450','yes');return false;">
                                <img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
                            Birth Year:
                        </td>
                        <td align="right" valign="top" class="formMain" class="formMain">
                            <select size="1" class="formMain" name="frmStaffBirthYear">
                                <option value="bad" class="formMainCentered">(please select)</option>
                                <% 
			birthYear = Year(Now) - 16
			Do Until birthYear = 1900
				birthYear = (birthYear - 1)
                                %>
                                <option value="<%= birthYear %>" <% If say = "edit" Then %><% If birthYear = GetStaff("BirthYear") Then %>
                                    selected<% End If %><% End If %> class="formMainCentered">
                                    <%= birthYear %></option>
                                <% 
			Loop			
                                %>
                            </select>
                        </td>
                    </tr>
                    <tr>
                        <td class="formMain">
                            <a href="../helpfiles/StaffFormHelp.asp?HelpID=Sex" onclick="NewWindow(this.href,'name','600','450','yes');return false;">
                                <img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
                            Gender:
                        </td>
                        <td align="right" valign="top" class="formMain">
                            <input type="radio" name="frmStaffSex" value="M" <% If say = "edit" Then %><% If Trim(GetStaff("Sex")) = "M" Then %>
                                checked<% End If %><% End If %>>M
                            <input type="radio" name="frmStaffSex" value="F" <% If say = "edit" Then %><% If Trim(GetStaff("Sex")) = "F" Then %>
                                checked<% End If %><% End If %>>F
                        </td>
                    </tr>
                    <!---
	<tr>
		<td class="formMain">
			<a href="../helpfiles/StaffFormHelp.asp?HelpID=EverABig" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
				Ever A Big?</td>
		<td align="right" valign="top" class="formMain">			
			<input type="radio" name="frmStaffEverABig" value="1"<% If say = "edit" Then %><% If Trim(GetStaff("EverABig")) = "1" Then %> checked<% End If %><% End If %>>Yes
			<input type="radio" name="frmStaffEverABig" value="0"<% If say = "edit" Then %><% If Trim(GetStaff("EverABig")) = "0" Then %> checked<% End If %><% End If %>>No
		</td>
	</tr>	--->
                    <!--	
<% 
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
 %>
	<tr>
		<td class="formMain">
			<a href="../helpfiles/StaffFormHelp.asp?HelpID=Position" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
				Position: (Click <a href="../helpfiles/StaffFormHelp.asp?HelpID=Position" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp; for Definitions)</td>
		<td align="right" valign="top" class="formMain">
			<select size="1" class="formMain" name="frmStaffPosition">
				<option value="bad" class="formMain">(please select)</option>
<% 
query = "SELECT code,position, grouping FROM tbl_StaffPosition ORDER BY grouping, code"
Set GetCode = Con.Execute(query)
Do Until GetCode.EOF
 %>
				<option value="<%= GetCode("code") %>"<% If say = "edit" Then %><% If GetCode("code") = GetStaff("Position") Then %> selected<% End If %><% End If %> class="formMain"><% if GetCode("Code") < 10 then %>0<%end if%><%= GetCode("code")%>&nbsp;-&nbsp;<%=GetCode("Grouping")%>&nbsp;-&nbsp;<%= GetCode("position") %></option>
<% 
	GetCode.MoveNext
Loop
GetCode.Close
Set GetCode = Nothing
 %>
			</select>
		</td>
	</tr>-->
                    <tr>
                        <td class="formMain">
                            <a href="../helpfiles/StaffFormHelp.asp?HelpID=Race" onclick="NewWindow(this.href,'name','600','450','yes');return false;">
                                <img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
                            Ethnicity/Race:
                        </td>
                        <td align="right" valign="top" class="formMain">
                            <select size="1" class="formMain" name="frmStaffRace">
                                <option value="bad" class="formMain">(please select)</option>
                                <% 
query = "SELECT code,race FROM tbl_StaffRace ORDER BY code"
Set GetCode = Con.Execute(query)
Do Until GetCode.EOF
                                %>
                                <option value="<%= GetCode("code") %>" <% If say = "edit" Then %><% If GetCode("code") = GetStaff("Race") Then %>
                                    selected<% End If %><% End If %> class="formMain">
                                    <%=GetCode("code")%>&nbsp;-&nbsp;<%= GetCode("race") %></option>
                                <% 
	GetCode.MoveNext
Loop
GetCode.Close
Set GetCode = Nothing
                                %>
                            </select>
                        </td>
                    </tr>
                    <!---
	<tr>
		<td class="formMain">
			<a href="../helpfiles/StaffFormHelp.asp?HelpID=Sex" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
				Gender:</td>
		<td align="right" valign="top" class="formMain">			
			<input type="radio" name="frmStaffSex" value="M"<% If say = "edit" Then %><% If Trim(GetStaff("Sex")) = "M" Then %> checked<% End If %><% End If %>>M
			<input type="radio" name="frmStaffSex" value="F"<% If say = "edit" Then %><% If Trim(GetStaff("Sex")) = "F" Then %> checked<% End If %><% End If %>>F
		</td>
	</tr>
	--->
                    <tr>
                        <td class="formMain">
                            <a href="../helpfiles/StaffFormHelp.asp?HelpID=Education" onclick="NewWindow(this.href,'name','600','450','yes');return false;">
                                <img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
                            Education:
                        </td>
                        <td align="right" valign="top" class="formMain">
                            <select size="1" class="formMain" name="frmStaffEducation">
                                <option value="bad" class="formMain">(please select)</option>
                                <% 
query = "SELECT code,education FROM tbl_StaffEducation ORDER BY code"
Set GetCode = Con.Execute(query)
Do Until GetCode.EOF
                                %>
                                <option value="<%= GetCode("code") %>" <% If say = "edit" Then %><% If GetCode("code") = GetStaff("Education") Then %>
                                    selected<% End If %><% End If %> class="formMain">
                                    <%=GetCode("code")%>&nbsp;-&nbsp;<%= GetCode("education") %></option>
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
                    <% 
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
                    %>
                    <tr>
                        <td class="formMain">
                            <a href="../helpfiles/StaffFormHelp.asp?HelpID=Position" onclick="NewWindow(this.href,'name','600','450','yes');return false;">
                                <img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
                            Position: (Click <a href="../helpfiles/StaffFormHelp.asp?HelpID=Position" onclick="NewWindow(this.href,'name','600','450','yes');return false;">
                                <img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
                            for Definitions)
                        </td>
                        <td align="right" valign="top" class="formMain">
                            <select size="1" class="formMain" name="frmStaffPosition">
                                <option value="bad" class="formMain">(please select)</option>
                                <% 
query = "SELECT code,position, grouping FROM tbl_StaffPosition ORDER BY grouping, code"
Set GetCode = Con.Execute(query)
Do Until GetCode.EOF
                                %>
                                <option value="<%= GetCode("code") %>" <% If say = "edit" Then %><% If GetCode("code") = GetStaff("Position") Then %>
                                    selected<% End If %><% End If %> class="formMain">
                                    <% if GetCode("Code") < 10 then %>0<%end if%><%= GetCode("code")%>&nbsp;-&nbsp;<%=GetCode("Grouping")%>&nbsp;-&nbsp;<%= GetCode("position") %></option>
                                <% 
	GetCode.MoveNext
Loop
GetCode.Close
Set GetCode = Nothing
                                %>
                            </select>
                        </td>
                    </tr>
                    <!---
	<tr>
		<td class="formMain">
			<a href="../helpfiles/StaffFormHelp.asp?HelpID=MonthStart" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
				Month Hired @ Agency:</td>
		<td align="right" valign="top" class="formMain">
			<select size="1" class="formMain" name="frmStaffMonthStart">
			<option value="bad" class="formMain">(please select)</option>
			<option value="1" class="formMain"<% If say = "edit" Then %><% If GetStaff("MonthStart") = 1 Then %> selected<% End If %><% End If %>>January</option>
			<option value="2" class="formMain"<% If say = "edit" Then %><% If GetStaff("MonthStart") = 2 Then %> selected<% End If %><% End If %>>February</option>
			<option value="3" class="formMain"<% If say = "edit" Then %><% If GetStaff("MonthStart") = 3 Then %> selected<% End If %><% End If %>>March</option>
			<option value="4" class="formMain"<% If say = "edit" Then %><% If GetStaff("MonthStart") = 4 Then %> selected<% End If %><% End If %>>April</option>
			<option value="5" class="formMain"<% If say = "edit" Then %><% If GetStaff("MonthStart") = 5 Then %> selected<% End If %><% End If %>>May</option>
			<option value="6" class="formMain"<% If say = "edit" Then %><% If GetStaff("MonthStart") = 6 Then %> selected<% End If %><% End If %>>June</option>
			<option value="7" class="formMain"<% If say = "edit" Then %><% If GetStaff("MonthStart") = 7 Then %> selected<% End If %><% End If %>>July</option>
			<option value="8" class="formMain"<% If say = "edit" Then %><% If GetStaff("MonthStart") = 8 Then %> selected<% End If %><% End If %>>August</option>
			<option value="9" class="formMain"<% If say = "edit" Then %><% If GetStaff("MonthStart") = 9 Then %> selected<% End If %><% End If %>>September</option>
			<option value="10" class="formMain"<% If say = "edit" Then %><% If GetStaff("MonthStart") = 10 Then %> selected<% End If %><% End If %>>October</option>
			<option value="11" class="formMain"<% If say = "edit" Then %><% If GetStaff("MonthStart") = 11 Then %> selected<% End If %><% End If %>>November</option>
			<option value="12" class="formMain"<% If say = "edit" Then %><% If GetStaff("MonthStart") = 12 Then %> selected<% End If %><% End If %>>December</option>
			</select>
		</td>
	</tr>
	<tr>
		<td class="formMain">
			<a href="../helpfiles/StaffFormHelp.asp?HelpID=YearStart" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
				Year Hired @ Agency:</td>
		<td align="right" valign="top" class="formMain">
			<select size="1" class="formMain" name="frmStaffYearStart">
			<option value="bad" class="formMain">(please select)</option>
			<% 
			yearStart = Year(Now)+1
			Do Until yearStart = 1960
				'yearStart = (yearStart)
				yearStart = (yearStart - 1) 'change to be able to enter 2007 year during 2007 year
			 %>
			<option value="<%= yearStart %>"<% If say = "edit" Then %><% If GetStaff("YearStart") = yearStart Then %> selected<% End If %><% End If %> class="formMain"><%= yearStart %></option>
			<% 
			Loop			
			%>
			</select>
		</td>
	</tr>
	--->
                    <!--
	<tr>
		<td class="formMain">
			<a href="../helpfiles/StaffFormHelp.asp?HelpID=YearsInNetwork" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
				Years of Service in the BBBS<img src="../images/whatsnew.gif" alt="" width="20" height="25"><br>Network: </td>
		<td align="right" valign="top" class="formMain">
			<select size="1" class="formMain" name="frmStaffYearsInNetwork">
			<option value="bad" class="formMain">(please select)</option>
			<% 
			yearStart = -1
			Do Until yearStart = 50
				yearStart = (yearStart + 1)
			 %>
			<option value="<%= yearStart %>"<% If say = "edit" Then %><% If GetStaff("YearsInNetwork") = yearStart Then %> selected<% End If %><% End If %> class="formMain"><%= yearStart %></option>
			<% 
			Loop			
			%>
			</select>
		</td>
	</tr>
--->
                    <!--
<tr> 
		<td class="formMain">
			<a href="../helpfiles/StaffFormHelp.asp?HelpID=Time" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
				Full-time/Part-time:<img src="../images/whatsnew.gif" alt="" width="20" height="25"></td> 
 		<td align="right" valign="top" class="formMain"> 
			<input type="radio" name="frmStaffTime" value="FT"<% If say = "edit" Then %><% If Trim(GetStaff("Time")) = "FT" Then %> checked<% End If %><% End If %> onclick="disableFTE(this.form)">Full-Time 
			<input type="radio" name="frmStaffTime" value="PT"<% If say = "edit" Then %><% If Trim(GetStaff("Time")) = "PT" Then %> checked<% End If %><% End If %> onclick="enableFTE(this.form)">Part-Time
			&nbsp;&nbsp;FTE&nbsp;<input type="text" size="4" value="<% If say = "edit" Then %><%= GetStaff("FTE") %><% Else  %>1<% End If %>" class="formMain" name="frmStaffFTE" onchange="checkForNumeric(this.value)">
		</td>
	</tr>
-->
                    <!--
	<tr>
		<td class="formMain">
			<a href="../helpfiles/StaffFormHelp.asp?HelpID=HoursWeek" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
				Hours per Week:</td>
		<td align="right" valign="top" class="formMain"><input type="text" size="4" value="<% If say = "edit" Then %><%= GetStaff("HoursWeek") %><% Else  %>0<% End If %>" class="formMain" name="frmStaffHoursWeek" onchange="checkForInteger(this.value)"></td>
	</tr>
		--->
                    <tr>
                        <td class="formMain">
                            <a href="../helpfiles/StaffFormHelp.asp?HelpID=PositionStartDate" onclick="NewWindow(this.href,'name','600','450','yes');return false;">
                                <img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
                            Date Started Current Position:<br>
                            (MM-DD-YYYY):
                        </td>
                        <td align="right" valign="top" class="formMain">
                            <% If say = "edit" Then 
				   Dim PosStart
				   If Not ISNULL(GetStaff("PositionStartDate")) Then 
				       PosStart = FormatDateTime(GetStaff("PositionStartDate"), 2) 'vbShortDate
				   Else
				       PosStart = ""
				   End If
                            %>
                            <input id="demo1" type="text" size="10" maxlength="10" value="<%=PosStart%>" class="formMain"
                                name="frmStaffPositionStartDate"><a href="javascript:NewCal('demo1','mmddyyyy')"><img
                                    src="../images/cal.gif" width="16" height="16" border="0" alt="Pick a date"></a>
                            <% Else  %>
                            <input id="demo1" type="text" size="10" maxlength="10" value="" class="formMain"
                                name="frmStaffPositionStartDate"><a href="javascript:NewCal('demo1','mmddyyyy')"><img
                                    src="../images/cal.gif" width="16" height="16" border="0" alt="Pick a date"></a>
                            <% End If %>
                        </td>
                    </tr>
                    <tr>
                        <td class="formMain">
                            <a href="../helpfiles/StaffFormHelp.asp?HelpID=HoursWeek" onclick="NewWindow(this.href,'name','600','450','yes');return false;">
                                <img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
                            Hours per Week:
                        </td>
                        <td align="right" valign="top" class="formMain">
                            <input type="text" size="4" value="<% If say = "edit" Then %><%= GetStaff("HoursWeek") %><% Else  %>0<% End If %>"
                                class="formMain" name="frmStaffHoursWeek" onchange="checkForInteger(this.value)">
                        </td>
                    </tr>
                    <tr>
                        <td class="formMain">
                            <a href="../helpfiles/StaffFormHelp.asp?HelpID=MonthStart" onclick="NewWindow(this.href,'name','600','450','yes');return false;">
                                <img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
                            Month Hired @ Agency:
                        </td>
                        <td align="right" valign="top" class="formMain">
                            <select size="1" class="formMain" name="frmStaffMonthStart">
                                <option value="bad" class="formMain">(please select)</option>
                                <option value="1" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthStart") = 1 Then %>
                                    selected<% End If %><% End If %>>January</option>
                                <option value="2" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthStart") = 2 Then %>
                                    selected<% End If %><% End If %>>February</option>
                                <option value="3" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthStart") = 3 Then %>
                                    selected<% End If %><% End If %>>March</option>
                                <option value="4" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthStart") = 4 Then %>
                                    selected<% End If %><% End If %>>April</option>
                                <option value="5" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthStart") = 5 Then %>
                                    selected<% End If %><% End If %>>May</option>
                                <option value="6" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthStart") = 6 Then %>
                                    selected<% End If %><% End If %>>June</option>
                                <option value="7" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthStart") = 7 Then %>
                                    selected<% End If %><% End If %>>July</option>
                                <option value="8" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthStart") = 8 Then %>
                                    selected<% End If %><% End If %>>August</option>
                                <option value="9" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthStart") = 9 Then %>
                                    selected<% End If %><% End If %>>September</option>
                                <option value="10" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthStart") = 10 Then %>
                                    selected<% End If %><% End If %>>October</option>
                                <option value="11" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthStart") = 11 Then %>
                                    selected<% End If %><% End If %>>November</option>
                                <option value="12" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthStart") = 12 Then %>
                                    selected<% End If %><% End If %>>December</option>
                            </select>
                        </td>
                    </tr>
                    <tr>
                        <td class="formMain">
                            <a href="../helpfiles/StaffFormHelp.asp?HelpID=YearStart" onclick="NewWindow(this.href,'name','600','450','yes');return false;">
                                <img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
                            Year Hired @ Agency:
                        </td>
                        <td align="right" valign="top" class="formMain">
                            <select size="1" class="formMain" name="frmStaffYearStart">
                                <option value="bad" class="formMain">(please select)</option>
                                <% 
			yearStart = Year(Now)+1
			Do Until yearStart = 1960
				'yearStart = (yearStart)
				yearStart = (yearStart - 1) 'change to be able to enter 2007 year during 2007 year
                                %>
                                <option value="<%= yearStart %>" <% If say = "edit" Then %><% If GetStaff("YearStart") = yearStart Then %>
                                    selected<% End If %><% End If %> class="formMain">
                                    <%= yearStart %></option>
                                <% 
			Loop			
                                %>
                            </select>
                        </td>
                    </tr>
                    <tr>
                        <td class="formMain">
                            <a href="../helpfiles/StaffFormHelp.asp?HelpID=YearsInNetwork" onclick="NewWindow(this.href,'name','600','450','yes');return false;">
                                <img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
                            Years of Service in the BBBS Network:
                        </td>
                        <td align="right" valign="top" class="formMain">
                            <select size="1" class="formMain" name="frmStaffYearsInNetwork">
                                <option value="bad" class="formMain">(please select)</option>
                                <% 
			yearStart = -1
			Do Until yearStart = 50
				yearStart = (yearStart + 1)
                                %>
                                <option value="<%= yearStart %>" <% If say = "edit" Then %><% If GetStaff("YearsInNetwork") = yearStart Then %>
                                    selected<% End If %><% End If %> class="formMain">
                                    <%= yearStart %></option>
                                <% 
			Loop			
                                %>
                            </select>
                        </td>
                    </tr>
                    <tr>
                        <td class="formMain">
                            <a href="../helpfiles/StaffFormHelp.asp?HelpID=Time" onclick="NewWindow(this.href,'name','600','450','yes');return false;">
                                <img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
                            Employment Status:
                        </td>
                        <td align="left" valign="top" class="formMain">
                            <!--	<input type="radio" name="frmStaffTime" value="FT"<% If say = "edit" Then %><% If Trim(GetStaff("Time")) = "FT" Then %> checked<% End If %><% End If %> onclick="disableFTE(this.form)">Full-Time 
			<input type="radio" name="frmStaffTime" value="PT"<% If say = "edit" Then %><% If Trim(GetStaff("Time")) = "PT" Then %> checked<% End If %><% End If %> onclick="enableFTE(this.form)">Part-Time
			&nbsp;&nbsp;FTE:&nbsp;<input type="text" size="4" value="<% If say = "edit" Then %><%= GetStaff("FTE") %><% Else  %>1<% End If %>" class="formMain" name="frmStaffFTE" onchange="checkForNumeric(this.value)"><br><br>-->
                            <!--<a href="../helpfiles/StaffFormHelp.asp?HelpID=Terminated" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;&nbsp;-->
                            &nbsp;If no longer employed, select termination month,<br>
                            &nbsp;Else select <b>"Still Employed"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                <select size="1" class="formMain" name="frmStaffTerminated" id="Select1">
                                    <option value="bad" class="formMain">(please select)</option>
                                    <option value="0" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 0 Then %>
                                        selected<% End If %><% End If %>>Still Employed</option>
                                    <option value="1" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 1 Then %>
                                        selected<% End If %><% End If %>>January</option>
                                    <option value="2" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 2 Then %>
                                        selected<% End If %><% End If %>>February</option>
                                    <option value="3" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 3 Then %>
                                        selected<% End If %><% End If %>>March</option>
                                    <option value="4" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 4 Then %>
                                        selected<% End If %><% End If %>>April</option>
                                    <option value="5" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 5 Then %>
                                        selected<% End If %><% End If %>>May</option>
                                    <option value="6" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 6 Then %>
                                        selected<% End If %><% End If %>>June</option>
                                    <option value="7" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 7 Then %>
                                        selected<% End If %><% End If %>>July</option>
                                    <option value="8" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 8 Then %>
                                        selected<% End If %><% End If %>>August</option>
                                    <option value="9" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 9 Then %>
                                        selected<% End If %><% End If %>>September</option>
                                    <option value="10" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 10 Then %>
                                        selected<% End If %><% End If %>>October</option>
                                    <option value="11" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 11 Then %>
                                        selected<% End If %><% End If %>>November</option>
                                    <option value="12" class="formMain" <% If say = "edit" Then %><% If GetStaff("MonthEnd") = 12 Then %>
                                        selected<% End If %><% End If %>>December</option>
                                </select>
                        </td>
                    </tr>
                    <!--
<%'added 12/20/2006 saf to capture Employee name %>
	<tr>
		<td class="formMain">
			<a href="../helpfiles/StaffFormHelp.asp?HelpID=EmployeeName" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
				Employee Name:</td>
		<td align="right" valign="top" class="formMain">
		<% if say = "add" or say = "form" or say = "delete" then %>
			<input type="text" size="70" maxlength="70" value="<% If say = "edit" Then %><%= GetStaff("EmployeeName") %><% Else  %><% End If %>" class="formMain" name="frmStaffEmployeeName">
		<% else %>
			<% if say = "edit" and not isnull(GetStaff("EmployeeName")) then %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<% end if %><input type="text" size="60" maxlength="600" value="<% If say = "edit" Then %><%= GetStaff("EmployeeName") %><% Else  %><% End If %>" class="formMain" name="frmStaffEmployeeName">
		<% end if %>
		</td>
	</tr>
--->
                    <tr>
                        <td class="formMain">
                            <a href="../helpfiles/StaffFormHelp.asp?HelpID=EverABig" onclick="NewWindow(this.href,'name','600','450','yes');return false;">
                                <img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
                            Ever A Big?
                        </td>
                        <td align="right" valign="top" class="formMain">
                            <input type="radio" name="frmStaffEverABig" value="1" <% If say = "edit" Then %><% If Trim(GetStaff("EverABig")) = "1" Then %>
                                checked<% End If %><% End If %>>Yes
                            <input type="radio" name="frmStaffEverABig" value="0" <% If say = "edit" Then %><% If Trim(GetStaff("EverABig")) = "0" Then %>
                                checked<% End If %><% End If %>>No
                        </td>
                    </tr>
                    <tr>
                        <td class="formMain">
                            <a href="../helpfiles/StaffFormHelp.asp?HelpID=YearlySalary" onclick="NewWindow(this.href,'name','600','450','yes');return false;">
                                <img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
                            Compensation<br>
                            (Salary + Bonus/Incentives)
                        </td>
                        <td align="right" valign="top" class="formMain">
                            <% if say = "add" or say = "form" or say = "delete" then %>
                            &nbsp;Base Salary&nbsp;&nbsp;$&nbsp;<input type="text" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetStaff("BaseSalary") %><% Else  %>0<% End If %>"
                                class="formMain" name="frmStaffBaseSalary" onchange="checkForNumeric(this.value); addSalary(this.form);">
                            &nbsp;+&nbsp;Bonus&nbsp;&nbsp;$&nbsp;<input type="text" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetStaff("BonusSalary") %><% Else  %>0<% End If %>"
                                class="formMain" name="frmStaffBonusSalary" onchange="checkForNumeric(this.value); addSalary(this.form);">
                            &nbsp;=&nbsp;Total&nbsp;&nbsp;$&nbsp;<input readonly="true" type="text" size="5"
                                maxlength="10" value="<% If say = "edit" Then %><%= GetStaff("YearlySalary") %><% Else  %>0<% End If %>"
                                class="formMain" name="frmStaffYearlySalary" onchange="checkForNumeric(this.value)">
                            <% else %>
                            <% 'if say = "edit" and not isnull(GetStaff("BaseSalary")) then %>&nbsp;Base Salary&nbsp;&nbsp;$&nbsp;<% 'end if %><input
                                type="text" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetStaff("BaseSalary") %><% Else  %>0<% End If %>"
                                class="formMain" name="frmStaffBaseSalary" onchange="checkForNumeric(this.value);addSalary(this.form);">
                            <% 'if say = "edit" and not isnull(GetStaff("BonusSalary")) then %>&nbsp;+&nbsp;Bonus&nbsp;&nbsp;$&nbsp;<% 'end if %><input
                                type="text" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetStaff("BonusSalary") %><% Else  %>0<% End If %>"
                                class="formMain" name="frmStaffBonusSalary" onchange="checkForNumeric(this.value);addSalary(this.form);">
                            <% 'if say = "edit" and not isnull(GetStaff("YearlySalary")) then %>&nbsp;=&nbsp;Total&nbsp;&nbsp;$&nbsp;<% 'end if %><input
                                readonly="true" type="text" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetStaff("YearlySalary") %><% Else  %>0<% End If %>"
                                class="formMain" name="frmStaffYearlySalary" onchange="checkForNumeric(this.value)">
                            <% end if %>
                        </td>
                    </tr>
                    <!--<tr>
		<td class="formMain">
			<a href="../helpfiles/StaffFormHelp.asp?HelpID=SalaryPriorYear" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
				Compensation <br>(Salary + Bonus/Incentives)</td>
		<td align="right" valign="top" class="formMain">
		<% if say = "add" or say = "form" or say = "delete" then %>
			Base Salary <input type="text" size="7" maxlength = "10">
			Bonus/Incentives<input type="text" size="7" maxlength = "10">
			Total<input type="text" size="7" maxlength = "10">
			
			
			<!--<input type="text" size="7" maxlength="10" value="<% If say = "edit" Then %><%= GetStaff("YearlySalary") %><% Else  %>0<% End If %>" class="formMain" name="frmStaffYearlySalary" onchange="checkForInteger(this.value)">
		<% else %>
			<% if say = "edit" and not isnull(GetStaff("SalaryPriorYear")) then %><i><strong><%=y-1%>&nbsp;Salary:&nbsp;<%= formatcurrency(GetStaff("SalaryPriorYear"))%></strong></i>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$&nbsp;<% end if %><input type="text" size="7" maxlength="10" value="<% If say = "edit" Then %><%= GetStaff("YearlySalary") %><% Else  %>0<% End If %>" class="formMain" name="frmStaffYearlySalary" onchange="checkForInteger(this.value)">
		<% end if %>
		</td>
	</tr>-->
                    <% If say = "edit" Then %>
                    <tr>
                        <td colspan="2" class="formHeader">
                            <input type="submit" value="Save Staff Member" class="formMainBold">
                        </td>
                    </tr>
                    <% Else %>
                    <tr>
                        <td colspan="2" class="formHeader">
                            <input type="submit" value="Save This Entry" class="formMainBold">
                        </td>
                    </tr>
                    <% End If %>
                </table>
                </form>
                <% End If %>
                <% End If %>
                <% If say <> "thanks" Then %>
                <% If display = "showSummary" Then %>
                <!-- 
<script language="JavaScript">

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
 -->
                </script>
                <!-- RESULTS TABLE STARTS HERE -->
                <table border="1" cellpadding="2" cellspacing="0" width="640" bordercolordark="003063">
                    <!-- first row of table headers -->
                    <tr>
                        <td colspan="8" align="center" valign="top" class="formMain">
                            If any of the following information needs to be changed,<br>
                            simply click "Edit Record" for that individual and re-enter their information.<br>
                            When all staff members have been added, click "Finish" to submit this form.
                        </td>
                    </tr>
                    <tr>
                        <td rowspan="3" class="formHeaderSmall">
                            #
                        </td>
                        <td class="formHeaderSmall">
                            Birth Year:
                        </td>
                        <td class="formHeaderSmall">
                            Position:
                        </td>
                        <td class="formHeaderSmall">
                            Race:
                        </td>
                        <td class="formHeaderSmall">
                            Sex:
                        </td>
                        <td class="formHeaderSmall" colspan="2">
                            Compensation<br>
                            (Base Salary)
                        </td>
                        <td rowspan="3" class="formHeaderSmall">
                            Edit
                        </td>
                    </tr>
                    <!-- second row of table headers -->
                    <tr>
                        <td class="formHeaderSmall">
                            Hired (M/Y):
                        </td>
                        <td class="formHeaderSmall">
                            Position Start:
                        </td>
                        <td class="formHeaderSmall">
                            Status:
                        </td>
                        <td class="formHeaderSmall">
                            Hrs/Wk:
                        </td>
                        <td class="formHeaderSmall" colspan="2">
                            Compensation<br>
                            (Bonus/Incentives)
                        </td>
                    </tr>
                    <!-- third row of table headers -->
                    <tr>
                        <td class="formHeaderSmall">
                            Years in BBBS:
                        </td>
                        <td class="formHeaderSmall">
                            Employee Name:
                        </td>
                        <td class="formHeaderSmall">
                            Education:
                        </td>
                        <td class="formHeaderSmall">
                            Ever a BIG:
                        </td>
                        <td class="formHeaderSmall" colspan="2">
                            Total Compensation<br>
                            (Salary+Bonus/Incentives)
                        </td>
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
                        <td colspan="8" class="formMainBold">
                            No Staff Members To List
                        </td>
                    </tr>
                    <%
						Else
						GetStaff.MoveFirst
						Do Until GetStaff.EOF
                    %>
                    <!-- first row of results -->
                    <tr>
                        <td rowspan="3" class="formMain">
                            <%= ct %>
                        </td>
                        <td class="formMain" align="center">
                            <%= GetStaff("BirthYear") %>
                        </td>
                        <% 
					query = "SELECT position FROM tbl_StaffPosition WHERE code=" & Int(GetStaff("position"))
					Set GetCode = Con.Execute(query)
                        %>
                        <td class="formMain" align="center">
                            <% If GetCode.EOF OR GetCode.BOF Then %><i>Unlisted</i><% else %>
                            <%= GetCode("position") %><% end if %>
                        </td>
                        <% 
					GetCode.Close
					Set GetCode = Nothing
                        %>
                        <% 
					query = "SELECT race FROM tbl_StaffRace WHERE code=" & Int(GetStaff("race"))
					Set GetCode = Con.Execute(query)
                        %>
                        <td class="formMain" align="center">
                            <%= GetCode("race") %>
                        </td>
                        <% 
					GetCode.Close
					Set GetCode = Nothing
                        %>
                        <td class="formMain" align="center">
                            <%= UCase(GetStaff("sex")) %>
                        </td>
                        <!--<td colspan="2" class="formMain" align="center">&nbsp;<%= (GetStaff("EmployeeName")) %>&nbsp;</td>-->
                        <td colspan="2" class="formMainRightJ">
                            <%If IsNull(GetStaff("basesalary")) Then%><%= FormatCurrency(GetStaff("yearlysalary")) %><%Else%><%= FormatCurrency(GetStaff("basesalary")) %><%End If%>&nbsp;
                        </td>
                        <td rowspan="3" align="right" class="formMain">
                            <a href="staff_edit.asp?status=editRow&row=<%= GetStaff("StaffID") %>&y=<%= Request("y") %>">
                                Edit Record</a><br>
                            <br>
                            <a href="#" onclick="confirmDelete(<%= GetStaff("StaffID") %>); return false;"></a>
                        </td>
                    </tr>
                    <tr>
                        <td class="formMain" align="center">
                            <%= (GetStaff("Monthstart")) & "/" & GetStaff("yearstart") %>
                        </td>
                        <td class="formMain" align="center">
                            &nbsp;<% If not ISNULL(GetStaff("PositionStartDate")) Then Response.Write(FormatDateTime(GetStaff("PositionStartDate"),2)) %>&nbsp;
                        </td>
                        <td class="formMain" align="center">
                            <% If GetStaff("MonthEnd") = 0 Then %>Still Employed<% Else %><%= MonthName(GetStaff("MonthEnd")) %><% End If %>
                        </td>
                        <td class="formMain" align="center">
                            <%If IsNull(GetStaff("hoursweek")) Then%>N/A<%Else%><%= GetStaff("hoursweek")%><%End If%>
                        </td>
                        <td colspan="2" class="formMainRightJ">
                            <%If IsNull(GetStaff("bonussalary")) Then%>0<%Else%><%= FormatCurrency(GetStaff("bonussalary"))%><%End If%>&nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td class="formMain" align="center">
                            <%If IsNull(GetStaff("YearsInNetwork")) Then%>N/A<%Else%><%= (GetStaff("YearsInNetwork"))%><%End If%>
                        </td>
                        <td class="formMain" align="center">
                            &nbsp;<%If IsNull(GetStaff("EmployeeName")) Then%>N/A<%Else%><%= (GetStaff("EmployeeName"))%><%End If%>&nbsp;
                        </td>
                        <% query = "SELECT education FROM tbl_StaffEducation WHERE code=" & Int(GetStaff("Education"))
					Set GetCode = Con.Execute(query)
                        %>
                        <td class="formMain" align="center">
                            <%= GetCode("Education") %>
                        </td>
                        <% 
					GetCode.Close
					Set GetCode = Nothing
                        %>
                        <td class="formMain" align="center">
                            <%If IsNull(GetStaff("EverABig")) Then%>1<%Else%><%= (GetStaff("EverABig")) %><%End If%>
                        </td>
                        <td colspan="2" class="formMainRightJ">
                            <%= FormatCurrency(GetStaff("yearlysalary")) %>&nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td colspan="8" class="formHeader">
                            <img src="../images/spacer.gif" width="1" height="5" alt="" border="0">
                        </td>
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
                        <td colspan="8" class="formHeader">
                            <input type="submit" value="Finish" class="formMainBold">
                        </td>
                    </tr>
                    <tr>
                        <td colspan="8">
                            <div align="center">
                                <!--#include file="../includes/contact_info.inc"-->
                            </div>
                        </td>
                    </tr>
                    </form>
                </table>
                <br>
                <br>
                <p>
                    <% End If %>
                    <% End If %>
            </td>
        </tr>
    </table>
</body>
</html>
