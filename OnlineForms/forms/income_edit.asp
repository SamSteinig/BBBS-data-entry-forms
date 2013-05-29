
<% 
If Request("status") = "addNew" Then	
	' Check for duplicate records
	
	Set DupCon = Server.CreateObject("ADODB.Connection")
	DupCon.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT Count(*) As NumberOfEntries FROM tbl_frmIncome WHERE AgencyID = '" & Request("AgencyIDN") & "' and Year = " & Request("Year")	
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
		RST.Open "SELECT * FROM tbl_frmIncome", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = Request("AgencyIDN")
		RST("Year") = Request("year")
			tt = Int(Request("frmIncomeUnitedWay")) _ 
			+ Int(Request("frmIncomeFederalGovernmentFunding"))_
			+ Int(Request("frmIncomeStateGovernmentFunding"))_			
			+ Int(Request("frmIncomeLocalGovernmentFunding"))_						
			+ Int(Request("frmIncomeFoundationGrants")) _
			+ Int(Request("frmIncomeCorporateGifts")) _
			+ Int(Request("frmIncomeBBBSAGrants"))_
			+ Int(Request("frmIncomeOnlineDonations"))_
			+ Int(Request("frmIncomeRMM"))_
			+ Int(Request("frmIncomeIndividualGivingBoard")) _
			+ Int(Request("frmIncomeIndividualGivingNonBoard")) _						
			+ Int(Request("frmIncomeBowlForKidsSake")) _			
			+ Int(Request("frmIncomeDinnerAuctions"))_
			+ Int(Request("frmIncomeGolf"))_			
			+ Int(Request("frmIncomeBingo"))_			
			+ Int(Request("frmIncomeRaffle"))_			
			+ Int(Request("frmIncomeCarsForKidsSake")) _
			+ Int(Request("frmIncomeOtherSpecialEvents"))_			
			+ Int(Request("frmIncomeDividendsInterest")) _
			+ Int(Request("frmIncomeOtherFunding"))
		RST("UnitedWay") = FormatCurrency(Request("frmIncomeUnitedWay"))
		RST("FederalGovernmentFunding") = FormatCurrency(Request("frmIncomeFederalGovernmentFunding"))
		RST("StateGovernmentFunding") = FormatCurrency(Request("frmIncomeStateGovernmentFunding"))		
		RST("LocalGovernmentFunding") = FormatCurrency(Request("frmIncomeLocalGovernmentFunding"))				
		RST("FoundationGrants") = FormatCurrency(Request("frmIncomeFoundationGrants"))
		RST("CorporateGifts") = FormatCurrency(Request("frmIncomeCorporateGifts"))
		RST("BBBSAGrants") = FormatCurrency(Request("frmIncomeBBBSAGrants"))	
		RST("OnlineDonations") = FormatCurrency(Request("frmIncomeOnlineDonations"))				
		RST("RMM") = FormatCurrency(Request("frmIncomeRMM"))			
		RST("IndividualGivingBoard") = FormatCurrency(Request("frmIncomeIndividualGivingBoard"))
		RST("IndividualGivingNonBoard") = FormatCurrency(Request("frmIncomeIndividualGivingNonBoard"))				
		RST("BowlForKidsSake") = FormatCurrency(Request("frmIncomeBowlForKidsSake"))
		RST("DinnerAuctions") = FormatCurrency(Request("frmIncomeDinnerAuctions"))		
		RST("Golf") = FormatCurrency(Request("frmIncomeGolf"))				
		RST("Bingo") = FormatCurrency(Request("frmIncomeBingo"))		
		RST("Raffle") = FormatCurrency(Request("frmIncomeRaffle"))						
		RST("CarsForKidsSake") = FormatCurrency(Request("frmIncomeCarsForKidsSake"))
		RST("OtherSpecialEvents") = FormatCurrency(Request("frmIncomeOtherSpecialEvents"))		
		RST("DividendsInterest") = FormatCurrency(Request("frmIncomeDividendsInterest"))
		RST("OtherFunding") = FormatCurrency(Request("frmIncomeOtherFunding"))
		RST("OtherFundingType") = Request("frmIncomeOtherFundingType")
		RST("NonMentoringIncome") = FormatCurrency(Request("frmIncomeNonMentoringIncome"))				
		RST("Total") = FormatCurrency(tt)
		RST("CreateDate") = Now
		RST.Update
		RST.Close
		Set RST = Nothing
		form = "Income"
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
	RST.Open "SELECT * FROM tbl_frmIncome WHERE agencyID='" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")), Con, 1, 3
		RST("Year") = Request("year")
			tt = Int(Request("frmIncomeUnitedWay")) _ 
			+ Int(Request("frmIncomeFederalGovernmentFunding"))_
			+ Int(Request("frmIncomeStateGovernmentFunding"))_			
			+ Int(Request("frmIncomeLocalGovernmentFunding"))_						
			+ Int(Request("frmIncomeFoundationGrants")) _
			+ Int(Request("frmIncomeCorporateGifts")) _
			+ Int(Request("frmIncomeBBBSAGrants"))_
			+ Int(Request("frmIncomeOnlineDonations"))_
			+ Int(Request("frmIncomeRMM"))_
			+ Int(Request("frmIncomeIndividualGivingBoard")) _
			+ Int(Request("frmIncomeIndividualGivingNonBoard")) _						
			+ Int(Request("frmIncomeBowlForKidsSake")) _			
			+ Int(Request("frmIncomeDinnerAuctions"))_
			+ Int(Request("frmIncomeGolf"))_			
			+ Int(Request("frmIncomeBingo"))_			
			+ Int(Request("frmIncomeRaffle"))_			
			+ Int(Request("frmIncomeCarsForKidsSake")) _
			+ Int(Request("frmIncomeOtherSpecialEvents"))_			
			+ Int(Request("frmIncomeDividendsInterest")) _
			+ Int(Request("frmIncomeOtherFunding"))
		RST("UnitedWay") = FormatCurrency(Request("frmIncomeUnitedWay"))
		RST("FederalGovernmentFunding") = FormatCurrency(Request("frmIncomeFederalGovernmentFunding"))
		RST("StateGovernmentFunding") = FormatCurrency(Request("frmIncomeStateGovernmentFunding"))		
		RST("LocalGovernmentFunding") = FormatCurrency(Request("frmIncomeLocalGovernmentFunding"))				
		RST("FoundationGrants") = FormatCurrency(Request("frmIncomeFoundationGrants"))
		RST("CorporateGifts") = FormatCurrency(Request("frmIncomeCorporateGifts"))
		RST("BBBSAGrants") = FormatCurrency(Request("frmIncomeBBBSAGrants"))	
		RST("OnlineDonations") = FormatCurrency(Request("frmIncomeOnlineDonations"))				
		RST("RMM") = FormatCurrency(Request("frmIncomeRMM"))			
		RST("IndividualGivingBoard") = FormatCurrency(Request("frmIncomeIndividualGivingBoard"))
		RST("IndividualGivingNonBoard") = FormatCurrency(Request("frmIncomeIndividualGivingNonBoard"))				
		RST("BowlForKidsSake") = FormatCurrency(Request("frmIncomeBowlForKidsSake"))
		RST("DinnerAuctions") = FormatCurrency(Request("frmIncomeDinnerAuctions"))		
		RST("Golf") = FormatCurrency(Request("frmIncomeGolf"))				
		RST("Bingo") = FormatCurrency(Request("frmIncomeBingo"))		
		RST("Raffle") = FormatCurrency(Request("frmIncomeRaffle"))						
		RST("CarsForKidsSake") = FormatCurrency(Request("frmIncomeCarsForKidsSake"))
		RST("OtherSpecialEvents") = FormatCurrency(Request("frmIncomeOtherSpecialEvents"))		
		RST("DividendsInterest") = FormatCurrency(Request("frmIncomeDividendsInterest"))
		RST("OtherFunding") = FormatCurrency(Request("frmIncomeOtherFunding"))
		RST("OtherFundingType") = Request("frmIncomeOtherFundingType")
		RST("NonMentoringIncome") = FormatCurrency(Request("frmIncomeNonMentoringIncome"))				
		RST("Total") = FormatCurrency(tt)
	jMod = RST("IncomeID")
	RST.Update
	RST.Close
	Set RST = Nothing
	form = "Income"
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
	<title>Income</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
<script language="JavaScript">
<!--

function addEmUp() {
	var box1 = Number(document.frmIncome.frmIncomeUnitedWay.value)
	var box2 = Number(document.frmIncome.frmIncomeFederalGovernmentFunding.value)
	var box3 = Number(document.frmIncome.frmIncomeStateGovernmentFunding.value)	
	var box4 = Number(document.frmIncome.frmIncomeLocalGovernmentFunding.value)		
	var box5 = Number(document.frmIncome.frmIncomeFoundationGrants.value)
	var box6 = Number(document.frmIncome.frmIncomeCorporateGifts.value)
	var box7 = Number(document.frmIncome.frmIncomeBBBSAGrants.value)	
	var box8 = Number(document.frmIncome.frmIncomeOnlineDonations.value)	
	var box9 = Number(document.frmIncome.frmIncomeRMM.value)
	var box10 = Number(document.frmIncome.frmIncomeIndividualGivingBoard.value)
	var box11 = Number(document.frmIncome.frmIncomeIndividualGivingNonBoard.value)		
	var box12 = Number(document.frmIncome.frmIncomeBowlForKidsSake.value)
	var box13 = Number(document.frmIncome.frmIncomeDinnerAuctions.value)	
	var box14 = Number(document.frmIncome.frmIncomeGolf.value)	
	var box15 = Number(document.frmIncome.frmIncomeBingo.value)	
	var box16 = Number(document.frmIncome.frmIncomeRaffle.value)	
	var box17 = Number(document.frmIncome.frmIncomeCarsForKidsSake.value)
	var box18 = Number(document.frmIncome.frmIncomeOtherSpecialEvents.value)
	var box19 = Number(document.frmIncome.frmIncomeDividendsInterest.value)
	var box20 = Number(document.frmIncome.frmIncomeOtherFunding.value)
	var boxtotal = box1 + box2 + box3 + box4 + box5 + box6 + box7 + box8 + box9 + box10 + box11 + box12 + box13 + box14 + box15 + box16 + box17 + box18 + box19 + box20
	document.frmIncome.frmIncomeTotal.value = boxtotal
}

function noChange()
	{
	alert("This will add automatically. Do not edit this field.");
	addEmUp();
	}

// -->
</script>

<script language="javascript">
<!--
function checkForInteger(valueToCheck)
{

	var myRegularExpression = /^[0-9]+(,[0-9]{3})*$/;  // contains any nonnumeric character???
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
	if(!(myRegularExpression1.test(form.frmIncomeUnitedWay.value)) || (form.frmIncomeUnitedWay.value == ""))
	{
		form.frmIncomeUnitedWay.focus();
		if(form.frmIncomeUnitedWay.value == "")
			alert("Please provide a value for United Way");
		else
			alert((form.frmIncomeUnitedWay.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmIncomeFederalGovernmentFunding.value)) || (form.frmIncomeFederalGovernmentFunding.value == ""))	
	{
		form.frmIncomeFederalGovernmentFunding.focus();
		if(form.frmIncomeFederalGovernmentFunding.value == "")
			alert("Please provide a value for Federal Government Funding");
		else
			alert((form.frmIncomeFederalGovernmentFunding.value) + " is an invalid value for Federal Government Funding");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmIncomeStateGovernmentFunding.value)) || (form.frmIncomeStateGovernmentFunding.value == ""))	
	{
		form.frmIncomeStateGovernmentFunding.focus();
		if(form.frmIncomeStateGovernmentFunding.value == "")
			alert("Please provide a value for State Government Funding");
		else
			alert((form.frmIncomeStateGovernmentFunding.value) + " is an invalid value for State Government Funding.");
		return false;
	}	
	else if(!(myRegularExpression1.test(form.frmIncomeLocalGovernmentFunding.value)) || (form.frmIncomeLocalGovernmentFunding.value == ""))	
	{
		form.frmIncomeLocalGovernmentFunding.focus();
		if(form.frmIncomeLocalGovernmentFunding.value == "")
			alert("Please provide a value for Local Government Funding");
		else
			alert((form.frmIncomeLocalGovernmentFunding.value) + " is an invalid value for Local Government Funding.");
		return false;
	}		
	else if(!(myRegularExpression1.test(form.frmIncomeFoundationGrants.value)) || (form.frmIncomeFoundationGrants.value == ""))	
	{
		form.frmIncomeFoundationGrants.focus();
		if(form.frmIncomeFoundationGrants.value == "")
			alert("Please provide a value for Foundation Grants");
		else
			alert((form.frmIncomeFoundationGrants.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmIncomeCorporateGifts.value)) || (form.frmIncomeCorporateGifts.value == ""))	
	{
		form.frmIncomeCorporateGifts.focus();
		if(form.frmIncomeCorporateGifts.value == "")
			alert("Please provide a value for Corporate Gifts");
		else
			alert((form.frmIncomeCorporateGifts.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmIncomeBBBSAGrants.value)) || (form.frmIncomeBBBSAGrants.value == ""))	
	{
		form.frmIncomeBBBSAGrants.focus();
		if(form.frmIncomeBBBSAGrants.value == "")
			alert("Please provide a value for BBBSA Grants");
		else
			alert((form.frmIncomeBBBSAGrants.value) + " is invalid.");
		return false;
	}	
	else if(!(myRegularExpression1.test(form.frmIncomeOnlineDonations.value)) || (form.frmIncomeOnlineDonations.value == ""))	
	{
		form.frmIncomeOnlineDonations.focus();
		if(form.frmIncomeOnlineDonations.value == "")
			alert("Please provide a value for Online Donations");
		else
			alert((form.frmIncomeOnlineDonations.value) + " is invalid.");
		return false;
	}		
	else if(!(myRegularExpression1.test(form.frmIncomeRMM.value)) || (form.frmIncomeRMM.value == ""))	
	{
		form.frmIncomeRMM.focus();
		if(form.frmIncomeRMM.value == "")
			alert("Please provide a value for RMM");
		else
			alert((form.frmIncomeRMM.value) + " is invalid.");
		return false;
	}	
	
	else if(!(myRegularExpression1.test(form.frmIncomeIndividualGivingBoard.value)) || (form.frmIncomeIndividualGivingBoard.value == ""))	
	{
		form.frmIncomeIndividualGivingBoard.focus();
		if(form.frmIncomeIndividualGivingBoard.value == "")
			alert("Please provide a value for Individual Giving - Board");
		else
			alert((form.frmIncomeIndividualGivingBoard.value) + " is invalid.");
		return false;
	}
	
	else if(!(myRegularExpression1.test(form.frmIncomeIndividualGivingNonBoard.value)) || (form.frmIncomeIndividualGivingNonBoard.value == ""))	
	{
		form.frmIncomeIndividualGivingNonBoard.focus();
		if(form.frmIncomeIndividualGivingNonBoard.value == "")
			alert("Please provide a value for Individual Giving - Board");
		else
			alert((form.frmIncomeIndividualGivingNonBoard.value) + " is invalid.");
		return false;
	}	
	
	else if(!(myRegularExpression1.test(form.frmIncomeBowlForKidsSake.value)) || (form.frmIncomeBowlForKidsSake.value == ""))	
	{
		form.frmIncomeBowlForKidsSake.focus();
		if(form.frmIncomeBowlForKidsSake.value == "")
			alert("Please provide a value for Bowl For Kids Sake");
		else
			alert((form.frmIncomeBowlForKidsSake.value) + " is invalid.");
		return false;
	}
	
	else if(!(myRegularExpression1.test(form.frmIncomeDinnerAuctions.value)) || (form.frmIncomeDinnerAuctions.value == ""))	
	{
		form.frmIncomeDinnerAuctions.focus();
		if(form.frmIncomeDinnerAuctions.value == "")
			alert("Please provide a value for Dinner / Auctions");
		else
			alert((form.frmIncomeDinnerAuctions.value) + " is invalid.");
		return false;
	}	
	

	else if(!(myRegularExpression1.test(form.frmIncomeGolf.value)) || (form.frmIncomeGolf.value == ""))	
	{
		form.frmIncomeGolf.focus();
		if(form.frmIncomeGolf.value == "")
			alert("Please provide a value for Golf");
		else
			alert((form.frmIncomeGolf.value) + " is invalid.");
		return false;
	}
	


	else if(!(myRegularExpression1.test(form.frmIncomeBingo.value)) || (form.frmIncomeBingo.value == ""))	
	{
		form.frmIncomeBingo.focus();
		if(form.frmIncomeBingo.value == "")
			alert("Please provide a value for Bingo");
		else
			alert((form.frmIncomeBingo.value) + " is invalid.");
		return false;
	}	
	
	else if(!(myRegularExpression1.test(form.frmIncomeRaffle.value)) || (form.frmIncomeRaffle.value == ""))	
	{
		form.frmIncomeRaffle.focus();
		if(form.frmIncomeRaffle.value == "")
			alert("Please provide a value for Raffle");
		else
			alert((form.frmIncomeRaffle.value) + " is invalid.");
		return false;
	}		
	
	
	else if(!(myRegularExpression1.test(form.frmIncomeCarsForKidsSake.value)) || (form.frmIncomeCarsForKidsSake.value == ""))	
	{
		form.frmIncomeCarsForKidsSake.focus();
		if(form.frmIncomeCarsForKidsSake.value == "")
			alert("Please provide a value for Cars for Kids' Sake");
		else
			alert((form.frmIncomeCarsForKidsSake.value) + " is invalid.");
		return false;
	}
	
	else if(!(myRegularExpression1.test(form.frmIncomeOtherSpecialEvents.value)) || (form.frmIncomeOtherSpecialEvents.value == ""))	
	{
		form.frmIncomeOtherSpecialEvents.focus();
		if(form.frmIncomeOtherSpecialEvents.value == "")
			alert("Please provide a value for Other Special Events");
		else
			alert((form.frmIncomeOtherSpecialEvents.value) + " is invalid.");
		return false;
	}

	
	else if(!(myRegularExpression1.test(form.frmIncomeDividendsInterest.value)) || (form.frmIncomeDividendsInterest.value == ""))	
	{
		form.frmIncomeDividendsInterest.focus();
		if(form.frmIncomeDividendsInterest.value == "")
			alert("Please provide a value for Dividends & Interest");
		else
			alert((form.frmIncomeDividendsInterest.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmIncomeOtherFunding.value)) || (form.frmIncomeOtherFunding.value == ""))	
	{
		form.frmIncomeOtherFunding.focus();
		if(form.frmIncomeOtherFunding.value == "")
			alert("Please provide a value for Other Funding");
		else
			alert((form.frmIncomeOtherFunding.value) + " is invalid.");
		return false;
	}
	else if(((form.frmIncomeOtherFunding.value != "0") && (form.frmIncomeOtherFunding.value != "")) && ((form.frmIncomeOtherFundingType.value == "") || (form.frmIncomeOtherFundingType.value == "")))
	{
		form.frmIncomeOtherFundingType.focus();
		alert("You indicated a dollar amount for Other Funding, but no Funding Type.  Please List Type of Funding.");
		return false;
	}
	
	else if((form.frmIncomeOtherFunding.value == "0") && (form.frmIncomeOtherFundingType.value != ""))
	{
		form.frmIncomeOtherFunding.focus();
		alert("You've listed an Other Funding Type but no dollar amount.  Please provide an amount for Other Funding, or clear out the Other Funding Type field.");
		return false;
	}
	
	
	else if(!(myRegularExpression1.test(form.frmIncomeNonMentoringIncome.value)) || (form.frmIncomeNonMentoringIncome.value == ""))	
	{
		form.frmNonMentoringIncome.focus();
		if(form.frmIncomeNonMentoringIncome.value == "")
			alert("Please provide a value for Non-Mentoring Income");
		else
			alert((form.frmIncomeNonMentoringIncome.value) + " is invalid.");
		return false;
	}		
	
	else if(Number(document.frmIncome.frmIncomeNonMentoringIncome.value) > Number(document.frmIncome.frmIncomeTotal.value))
	{
		form.frmIncomeNonMentoringIncome.focus();
		alert("Non-Mentoring Income cannot be greater than Total Income");
		return false;
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


<table width=100% cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_fishing.jpg" alt="" width="220" height="477" border="0"></td>

<td valign="top">
<br>

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

<table border="1" cellpadding="2" cellspacing="0" width="400" bordercolordark="003063">
<form name="frmIncome" action="income_edit.asp" method="post" onsubmit="return submitFormValidate(this)">
<!--#include file="../includes/form_stamp.asp"-->
<center>
<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmIncome WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
	Set GetIncome = Con.Execute(query)
 %>
<input type="hidden" name="status" value="editSave">
<% Else %>
<input type="hidden" name="status" value="addNew">
<%
End If
 %>

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
		<td colspan="2" align="center" valign="top" class="formSubhead">BBBS - <%= y %> Annual Agency Information (AAI)</td>
	</tr>
	<tr>
		<td colspan="2" class="formHeader">Revenue</td>
	</tr>
	
	<tr>
		<td colspan="2" class="formMain"><font color="#ff0000"><div align="center"><strong>Please Note: </strong>After entering your information, you <strong>must</strong> click on the "Save Form" button at the bottom of the form and wait for the "Thank You" screen or your changes will be lost.</div></font></td>
	</tr>		
	
	<tr>
		<td colspan="2" align="center" valign="top" class="formMain"><span class="formSubHead">List below the amount of your total <b>Gross Revenue</b> for each category. <br>For fundraisers such as BFKS use <b>NET</b> amount raised.<br>(e.g. total amount raised less <b>DIRECT</b> expense)</span></td>
	</tr>
	<tr>
		<td class="formMain">United Way</td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("UnitedWay") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" class="formMain" name="frmIncomeUnitedWay" onblur="addEmUp();"></td>
	</tr>
	<tr>
		<td class="formMain">Federal Government Funding</td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("FederalGovernmentFunding") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" class="formMain" name="frmIncomeFederalGovernmentFunding" onblur="addEmUp();"></td>
	</tr>
	<tr>
		<td class="formMain">State Government Funding</td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("StateGovernmentFunding") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" class="formMain" name="frmIncomeStateGovernmentFunding" onblur="addEmUp();"></td>
	</tr>	
	<tr>
		<td class="formMain">Local Government Funding</td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("LocalGovernmentFunding") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" class="formMain" name="frmIncomeLocalGovernmentFunding" onblur="addEmUp();"></td>
	</tr>		
	<tr>
		<td class="formMain">Foundation Grants</td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("FoundationGrants") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" class="formMain" name="frmIncomeFoundationGrants" onblur="addEmUp();"></td>
	</tr>
	<tr>
		<td class="formMain">Corporate Gifts</td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("CorporateGifts") %><% Else %>0<% End If %>"  onchange="checkForInteger(this.value);" class="formMain" name="frmIncomeCorporateGifts" onblur="addEmUp();"></td>
	</tr>	
	<tr>
		<td class="formMain">BBBSA Grants</td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("BBBSAGrants") %><% Else %>0<% End If %>"  onchange="checkForInteger(this.value);" class="formMain" name="frmIncomeBBBSAGrants" onblur="addEmUp();"></td>
	</tr>	
	<tr>
		<td class="formMain">Online Donations <span class="formSubHead"> (through BBBSA)</span></td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("OnlineDonations") %><% Else %>0<% End If %>"  onchange="checkForInteger(this.value);" class="formMain" name="frmIncomeOnlineDonations" onblur="addEmUp();"></td>
	</tr>		
	<tr>
		<td class="formMain">RMM <span class="formSubHead"> (Raising More Money)</span></td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("RMM") %><% Else %>0<% End If %>"  onchange="checkForInteger(this.value);" class="formMain" name="frmIncomeRMM" onblur="addEmUp();"></td>
	</tr>			
	<tr>
		<td class="formMain">Individual Giving - Board <span class="formSubHead">(excluding RMM)</span></td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("IndividualGivingBoard") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" class="formMain" name="frmIncomeIndividualGivingBoard" onblur="addEmUp();"></td>
	</tr>
	
	<tr>
		<td class="formMain">Individual Giving - Non Board <span class="formSubHead">(excluding RMM)</span></td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("IndividualGivingNonBoard") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" class="formMain" name="frmIncomeIndividualGivingNonBoard" onblur="addEmUp();"></td>
	</tr>		
	
	<tr>
		<td colspan="2" class="formHeaderSmall"><strong><div align="center">Special Events</div></strong></td>
	</tr>
	
	
	<tr>
		<td class="formMain">Bowl For Kids' Sake <span class="formSubHead">(BFKS)</span></td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("BowlForKidsSake") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" class="formMain" name="frmIncomeBowlForKidsSake" onblur="addEmUp();"></td>
	</tr>
	
	<tr>
		<td class="formMain">Dinner / Auctions</td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("DinnerAuctions") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" class="formMain" name="frmIncomeDinnerAuctions" onblur="addEmUp();"></td>
	</tr>
	
	<tr>
		<td class="formMain">Golf</td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("Golf") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" class="formMain" name="frmIncomeGolf" onblur="addEmUp();"></td>
	</tr>	
	
	<tr>
		<td class="formMain">Bingo</td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("Bingo") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" class="formMain" name="frmIncomeBingo" onblur="addEmUp();"></td>
	</tr>		
	
	<tr>
		<td class="formMain">Raffle</td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("Raffle") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" class="formMain" name="frmIncomeRaffle" onblur="addEmUp();"></td>
	</tr>		
	
	<tr>
		<td class="formMain">Cars For Kids' Sake <span class="formSubHead">(CFKS)</span></td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("CarsForKidsSake") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" class="formMain" name="frmIncomeCarsForKidsSake" onblur="addEmUp();"></td>
	</tr>
	
	<tr>
		<td class="formMain">Other Special Events <span class="formSubHead">(Total)</span></td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("OtherSpecialEvents") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" class="formMain" name="frmIncomeOtherSpecialEvents" onblur="addEmUp();"></td>
	</tr>	
	
	<tr>
		<td class="formHeaderSmall" colspan="2">&nbsp;</td>
	</tr>

	
	<tr>
		<td class="formMain">Dividends &amp; Interest</td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("DividendsInterest") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" class="formMain" name="frmIncomeDividendsInterest" onblur="addEmUp();"></td>
	</tr>
	<tr>
		<td class="formMain">Other Funding&nbsp;&nbsp;<span class="formSubHead">(please describe funding type below)</span><br><input type="text" size="25" maxlength="50" value="<% If say = "edit" Then %><%= GetIncome("OtherFundingType") %><% Else %><% End If %>" class="formMain" name="frmIncomeOtherFundingType" onblur="addEmUp();"></td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("OtherFunding") %><% Else %>0<% End If %>" class="formMain" name="frmIncomeOtherFunding" onblur="addEmUp();" onchange="checkForInteger(this.value);"></td>
	</tr>
	<tr>
		<td class="formMainRightJ">TOTAL&nbsp;&#61;&nbsp;</td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("Total") %><% Else %>0<% End If %>" onchange="noChange();" class="formMain" name="frmIncomeTotal" readonly></td>
	</tr>
	<tr>
		<td class="formHeaderSmall" colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="formMain">Of this total, how much income is targeted for <strong>Non-Mentoring</strong> Programs?</td>
		<td class="formMain">$&nbsp;<input type="text" class="formMain" size="8" maxlength="25" value="<% If say = "edit" Then %><%= GetIncome("NonMentoringIncome") %><% Else %>0<% End If %>" class="formMain" name="frmIncomeNonMentoringIncome" onblur="addEmUp();" onchange="checkForInteger(this.value);"></td>		
	</tr>
	
	<tr>
		<td colspan="2" class="formHeader"><input type="submit" value="Save Form" class="formMainBold"></td>
	</tr>
	<tr>
		<td colspan="2"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
	</tr>
	
</table><% 
If say = "edit" Then
	GetIncome.Close
	Set GetIncome = Nothing
	Con.Close
	Set Con = Nothing
End If
 %>

</center>

</td>
</tr>
</table>


</form>
<% End If %>

</body>
</html>
