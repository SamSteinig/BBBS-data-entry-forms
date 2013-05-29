
<% 
If Request("status") = "addNew" Then
' Check for duplicate records
	
	Set DupCon = Server.CreateObject("ADODB.Connection")
	DupCon.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT Count(*) As NumberOfEntries FROM tbl_frmBoardMembers WHERE AgencyID = '" & Request("AgencyIDN") & "' and Year = " & Request("Year")	
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
		RST.Open "SELECT * FROM tbl_frmBoardMembers", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = Request("AgencyIDN")
		RST("Year") = Request("year")
		RST("NumberBoardMembers") = Request("frmBoardMembersNumberBoardMembers")
		RST("TermLimitsYears") = Request("frmBoardMembersTermLimitsYears")
		RST("AverageTenureYears") = Request("frmBoardMembersAverageTenureYears")
		RST("AverageTenureMonths") = Request("frmBoardMembersAverageTenureMonths")
		If Request("frmBoardMembersStandingCommitteesPersonnel") = "Yes" Then
			RST("StandingCommitteesPersonnel") = True
		Else
			RST("StandingCommitteesPersonnel") = False
		End If
		If Request("frmBoardMembersStandingCommitteesProgram") = "Yes" Then
			RST("StandingCommitteesProgram") = True
		Else
			RST("StandingCommitteesProgram") = False
		End If
		If Request("frmBoardMembersStandingCommitteesExecutive") = "Yes" Then
			RST("StandingCommitteesExecutive") = True
		Else
			RST("StandingCommitteesExecutive") = False
		End If
		If Request("frmBoardMembersStandingCommitteesFundDevelopment") = "Yes" Then
			RST("StandingCommitteesFundDevelopment") = True
		Else
			RST("StandingCommitteesFundDevelopment") = False
		End If
		If Request("frmBoardMembersStandingCommitteesFinance") = "Yes" Then
			RST("StandingCommitteesFinance") = True
		Else
			RST("StandingCommitteesFinance") = False
		End If
		If Request("frmBoardMembersStandingCommitteesPublicRelations") = "Yes" Then
			RST("StandingCommitteesPublicRelations") = True
		Else
			RST("StandingCommitteesPublicRelations") = False
		End If
		If Request("frmBoardMembersStandingCommitteesStrategicPlanning") = "Yes" Then
			RST("StandingCommitteesStrategicPlanning") = True
		Else
			RST("StandingCommitteesStrategicPlanning") = False
		End If
		If Request("frmBoardMembersStandingCommitteesBoardDevelopment") = "Yes" Then
			RST("StandingCommitteesBoardDevelopment") = True
		Else
			RST("StandingCommitteesBoardDevelopment") = False
		End If
		If Request("frmBoardMembersStandingCommitteesVolunteerRecruitment") = "Yes" Then
			RST("StandingCommitteesVolunteerRecruitment") = True
		Else
			RST("StandingCommitteesVolunteerRecruitment") = False
		End If
		If Request("frmBoardMembersStandingCommitteesOther") = "Yes" Then
			RST("StandingCommitteesOther") = True
		Else
			RST("StandingCommitteesOther") = False
		End If
		RST("StandingCommitteesOtherText") = Request("frmBoardMembersStandingCommitteesOtherText")
		RST("FemaleWhite") = Request("frmBoardMembersFemaleWhite")
		RST("FemaleBlack") = Request("frmBoardMembersFemaleBlack")
		RST("FemaleHispanic") = Request("frmBoardMembersFemaleHispanic")
		RST("FemaleAsian") = Request("frmBoardMembersFemaleAsian")
		RST("FemaleIslander") = Request("frmBoardMembersFemaleIslander")
		RST("FemaleNative") = Request("frmBoardMembersFemaleNative")
		RST("FemaleMulti") = Request("frmBoardMembersFemaleMulti")
		RST("FemaleUnknown") = Request("frmBoardMembersFemaleUnknown")
		RST("MaleWhite") = Request("frmBoardMembersMaleWhite")
		RST("MaleBlack") = Request("frmBoardMembersMaleBlack")
		RST("MaleHispanic") = Request("frmBoardMembersMaleHispanic")
		RST("MaleAsian") = Request("frmBoardMembersMaleAsian")
		RST("MaleIslander") = Request("frmBoardMembersMaleIslander")
		RST("MaleNative") = Request("frmBoardMembersMaleNative")
		RST("MaleMulti") = Request("frmBoardMembersMaleMulti")
		RST("MaleUnknown") = Request("frmBoardMembersMaleUnknown")
		If Request("frmBoardMembersFrequency") = "Monthly" Then
			RST("FrequencyMonthly") = True
			RST("FrequencyTwoMonths") = False
			RST("FrequencyQuarterly") = False
			RST("FrequencyOther") = False
		ElseIf Request("frmBoardMembersFrequency") = "TwoMonths" Then
			RST("FrequencyMonthly") = False
			RST("FrequencyTwoMonths") = True
			RST("FrequencyQuarterly") = False
			RST("FrequencyOther") = False
		ElseIf Request("frmBoardMembersFrequency") = "Quarterly" Then
			RST("FrequencyMonthly") = False
			RST("FrequencyTwoMonths") = False
			RST("FrequencyQuarterly") = True
			RST("FrequencyOther") = False
		ElseIf Request("frmBoardMembersFrequency") = "Other" Then
			RST("FrequencyMonthly") = False
			RST("FrequencyTwoMonths") = False
			RST("FrequencyQuarterly") = False
			RST("FrequencyOther") = True
		End If
		RST("FrequencyOtherText") = Request("frmBoardMembersFrequencyOtherText")
		If Request("frmBoardMembersMoney") = "Minimum" Then
			RST("MoneyMinimum") = True
			RST("MoneyInKind") = False
			RST("MoneyNotExpected") = False
			RST("MoneyNoPolicy") = False
		ElseIf Request("frmBoardMembersMoney") = "InKind" Then
			RST("MoneyMinimum") = False
			RST("MoneyInKind") = True
			RST("MoneyNotExpected") = False
			RST("MoneyNoPolicy") = False
		ElseIf Request("frmBoardMembersMoney") = "NotExpected" Then
			RST("MoneyMinimum") = False
			RST("MoneyInKind") = False
			RST("MoneyNotExpected") = True
			RST("MoneyNoPolicy") = False
		ElseIf Request("frmBoardMembersMoney") = "NoPolicy" Then
			RST("MoneyMinimum") = False
			RST("MoneyInKind") = False
			RST("MoneyNotExpected") = False
			RST("MoneyNoPolicy") = True
		End If
		RST("MoneyMinimumAmount") = FormatCurrency(Request("frmBoardMembersMoneyMinimumAmount"))
		RST("YearlyContribution") = FormatCurrency(Request("frmBoardMembersYearlyContribution"))
		RST("SkillsFinance") = Request("frmBoardMembersSkillsFinance")
		RST("SkillsLegal") = Request("frmBoardMembersSkillsLegal")
		RST("SkillsPublicRelations") = Request("frmBoardMembersSkillsPublicRelations")
		RST("SkillsHumanServicesPractitioner") = Request("frmBoardMembersSkillsHumanServicesPractitioner")
		RST("SkillsHumanServicesAdministrator") = Request("frmBoardMembersSkillsHumanServicesAdministrator")
		RST("SkillsFullTimeStudent") = Request("frmBoardMembersSkillsFullTimeStudent")
		RST("SkillsHumanResources") = Request("frmBoardMembersSkillsHumanResources")
		RST("SkillsCorporateCEO") = Request("frmBoardMembersSkillsCorporateCEO")
		RST("SkillsOtherCorporateOfficer") = Request("frmBoardMembersSkillsOtherCorporateOfficer")
		RST("SkillsInsurance") = Request("frmBoardMembersSkillsInsurance")
		RST("SkillsSmallBusiness") = Request("frmBoardMembersSkillsSmallBusiness")
		RST("SkillsBig") = Request("frmBoardMembersSkillsBig")
		RST("SkillsParentLittle") = Request("frmBoardMembersSkillsParentLittle")
		RST("SkillsLittle") = Request("frmBoardMembersSkillsLittle")
		RST("SkillsLocalGovernment") = Request("frmBoardMembersSkillsLocalGovernment")
		RST("SkillsOther") = Request("frmBoardMembersSkillsOther")
		RST("SkillsUnknown") = Request("frmBoardMembersSkillsUnknown")
		RST("CreateDate") = Now
		RST.Update
		RST.Close
		Set RST = Nothing
		form = "BoardMembers"
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
	RST.Open "SELECT * FROM tbl_frmBoardMembers WHERE agencyID='" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")), Con, 1, 3
	RST("NumberBoardMembers") = Request("frmBoardMembersNumberBoardMembers")
	RST("TermLimitsYears") = Request("frmBoardMembersTermLimitsYears")
	RST("AverageTenureYears") = Request("frmBoardMembersAverageTenureYears")
	RST("AverageTenureMonths") = Request("frmBoardMembersAverageTenureMonths")
	If Request("frmBoardMembersStandingCommitteesPersonnel") = "Yes" Then
		RST("StandingCommitteesPersonnel") = True
	Else
		RST("StandingCommitteesPersonnel") = False
	End If
	If Request("frmBoardMembersStandingCommitteesProgram") = "Yes" Then
		RST("StandingCommitteesProgram") = True
	Else
		RST("StandingCommitteesProgram") = False
	End If
	If Request("frmBoardMembersStandingCommitteesExecutive") = "Yes" Then
		RST("StandingCommitteesExecutive") = True
	Else
		RST("StandingCommitteesExecutive") = False
	End If
	If Request("frmBoardMembersStandingCommitteesFundDevelopment") = "Yes" Then
		RST("StandingCommitteesFundDevelopment") = True
	Else
		RST("StandingCommitteesFundDevelopment") = False
	End If
	If Request("frmBoardMembersStandingCommitteesFinance") = "Yes" Then
		RST("StandingCommitteesFinance") = True
	Else
		RST("StandingCommitteesFinance") = False
	End If
	If Request("frmBoardMembersStandingCommitteesPublicRelations") = "Yes" Then
		RST("StandingCommitteesPublicRelations") = True
	Else
		RST("StandingCommitteesPublicRelations") = False
	End If
	If Request("frmBoardMembersStandingCommitteesStrategicPlanning") = "Yes" Then
		RST("StandingCommitteesStrategicPlanning") = True
	Else
		RST("StandingCommitteesStrategicPlanning") = False
	End If
	If Request("frmBoardMembersStandingCommitteesBoardDevelopment") = "Yes" Then
		RST("StandingCommitteesBoardDevelopment") = True
	Else
		RST("StandingCommitteesBoardDevelopment") = False
	End If
	If Request("frmBoardMembersStandingCommitteesVolunteerRecruitment") = "Yes" Then
		RST("StandingCommitteesVolunteerRecruitment") = True
	Else
		RST("StandingCommitteesVolunteerRecruitment") = False
	End If
	If Request("frmBoardMembersStandingCommitteesOther") = "Yes" Then
		RST("StandingCommitteesOther") = True
	Else
		RST("StandingCommitteesOther") = False
	End If
	RST("StandingCommitteesOtherText") = Request("frmBoardMembersStandingCommitteesOtherText")
	RST("FemaleWhite") = Request("frmBoardMembersFemaleWhite")
	RST("FemaleBlack") = Request("frmBoardMembersFemaleBlack")
	RST("FemaleHispanic") = Request("frmBoardMembersFemaleHispanic")
	RST("FemaleAsian") = Request("frmBoardMembersFemaleAsian")
	RST("FemaleIslander") = Request("frmBoardMembersFemaleIslander")
	RST("FemaleNative") = Request("frmBoardMembersFemaleNative")
	RST("FemaleMulti") = Request("frmBoardMembersFemaleMulti")
	RST("FemaleUnknown") = Request("frmBoardMembersFemaleUnknown")
	RST("MaleWhite") = Request("frmBoardMembersMaleWhite")
	RST("MaleBlack") = Request("frmBoardMembersMaleBlack")
	RST("MaleHispanic") = Request("frmBoardMembersMaleHispanic")
	RST("MaleAsian") = Request("frmBoardMembersMaleAsian")
	RST("MaleIslander") = Request("frmBoardMembersMaleIslander")
	RST("MaleNative") = Request("frmBoardMembersMaleNative")
	RST("MaleMulti") = Request("frmBoardMembersMaleMulti")
	RST("MaleUnknown") = Request("frmBoardMembersMaleUnknown")
	If Request("frmBoardMembersFrequency") = "Monthly" Then
		RST("FrequencyMonthly") = True
		RST("FrequencyTwoMonths") = False
		RST("FrequencyQuarterly") = False
		RST("FrequencyOther") = False
	ElseIf Request("frmBoardMembersFrequency") = "TwoMonths" Then
		RST("FrequencyMonthly") = False
		RST("FrequencyTwoMonths") = True
		RST("FrequencyQuarterly") = False
		RST("FrequencyOther") = False
	ElseIf Request("frmBoardMembersFrequency") = "Quarterly" Then
		RST("FrequencyMonthly") = False
		RST("FrequencyTwoMonths") = False
		RST("FrequencyQuarterly") = True
		RST("FrequencyOther") = False
	ElseIf Request("frmBoardMembersFrequency") = "Other" Then
		RST("FrequencyMonthly") = False
		RST("FrequencyTwoMonths") = False
		RST("FrequencyQuarterly") = False
		RST("FrequencyOther") = True
	End If
	RST("FrequencyOtherText") = Request("frmBoardMembersFrequencyOtherText")
	If Request("frmBoardMembersMoney") = "Minimum" Then
		RST("MoneyMinimum") = True
		RST("MoneyInKind") = False
		RST("MoneyNotExpected") = False
		RST("MoneyNoPolicy") = False
	ElseIf Request("frmBoardMembersMoney") = "InKind" Then
		RST("MoneyMinimum") = False
		RST("MoneyInKind") = True
		RST("MoneyNotExpected") = False
		RST("MoneyNoPolicy") = False
	ElseIf Request("frmBoardMembersMoney") = "NotExpected" Then
		RST("MoneyMinimum") = False
		RST("MoneyInKind") = False
		RST("MoneyNotExpected") = True
		RST("MoneyNoPolicy") = False
	ElseIf Request("frmBoardMembersMoney") = "NoPolicy" Then
		RST("MoneyMinimum") = False
		RST("MoneyInKind") = False
		RST("MoneyNotExpected") = False
		RST("MoneyNoPolicy") = True
	End If
	RST("MoneyMinimumAmount") = FormatCurrency(Request("frmBoardMembersMoneyMinimumAmount"))
	RST("YearlyContribution") = FormatCurrency(Request("frmBoardMembersYearlyContribution"))
	RST("SkillsFinance") = Request("frmBoardMembersSkillsFinance")
	RST("SkillsLegal") = Request("frmBoardMembersSkillsLegal")
	RST("SkillsPublicRelations") = Request("frmBoardMembersSkillsPublicRelations")
	RST("SkillsHumanServicesPractitioner") = Request("frmBoardMembersSkillsHumanServicesPractitioner")
	RST("SkillsHumanServicesAdministrator") = Request("frmBoardMembersSkillsHumanServicesAdministrator")
	RST("SkillsFullTimeStudent") = Request("frmBoardMembersSkillsFullTimeStudent")
	RST("SkillsHumanResources") = Request("frmBoardMembersSkillsHumanResources")
	RST("SkillsCorporateCEO") = Request("frmBoardMembersSkillsCorporateCEO")
	RST("SkillsOtherCorporateOfficer") = Request("frmBoardMembersSkillsOtherCorporateOfficer")
	RST("SkillsInsurance") = Request("frmBoardMembersSkillsInsurance")
	RST("SkillsSmallBusiness") = Request("frmBoardMembersSkillsSmallBusiness")
	RST("SkillsBig") = Request("frmBoardMembersSkillsBig")
	RST("SkillsParentLittle") = Request("frmBoardMembersSkillsParentLittle")
	RST("SkillsLittle") = Request("frmBoardMembersSkillsLittle")
	RST("SkillsLocalGovernment") = Request("frmBoardMembersSkillsLocalGovernment")
	RST("SkillsOther") = Request("frmBoardMembersSkillsOther")
	RST("SkillsUnknown") = Request("frmBoardMembersSkillsUnknown")
	jMod = RST("BoardMembersID")
	RST.Update
	RST.Close
	Set RST = Nothing
	form = "BoardMembers"
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
	<title>Board Members</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
	
<script language="javascript">
	<!--

	
function checkForInteger(valueToCheck)
{
	var myRegularExpression = /^[0-9]+(,[0-9]{3})*$/;  // Checks for integer with or without commas
	if(!(myRegularExpression.test(valueToCheck)))
	{
		alert("Please make sure you have entered a whole number with no spaces.\n We cannot process letters or words."); 
	} 
}

function changeForm1()
{
	if((document.frmBoardMembers.frmBoardMembersStandingCommitteesOtherText.value != "") && (document.frmBoardMembers.frmBoardMembersStandingCommitteesOtherText.value != "(Name Here)"))
	{
		document.frmBoardMembers.frmBoardMembersStandingCommitteesOther.checked = true;
	}
}

function changeForm2()
{
	if((document.frmBoardMembers.frmBoardMembersFrequencyOtherText.value != "") && (document.frmBoardMembers.frmBoardMembersFrequencyOtherText.value != "(Please Enter)"))
	{
		document.frmBoardMembers.frmBoardMembersFrequency[3].checked = true;
	}
}

function changeForm3()
{
	if((document.frmBoardMembers.frmBoardMembersMoneyMinimumAmount.value != "") && (document.frmBoardMembers.frmBoardMembersMoneyMinimumAmount.value != "0"))
	{
		document.frmBoardMembers.frmBoardMembersMoney[0].checked = true;
	}
}



var myRegularExpression1 = /^[0-9]+(,[0-9]{3})*$/;
	
function submitFormValidate(form)
{
var F1 = form.frmBoardMembersFemaleWhite.value.replace(",",""); // remove any comma from the form entry and replace it with nothing
var F2 = form.frmBoardMembersFemaleBlack.value.replace(",","");
var F3 = form.frmBoardMembersFemaleHispanic.value.replace(",","");
var F4 = form.frmBoardMembersFemaleAsian.value.replace(",","");
var F5 = form.frmBoardMembersFemaleIslander.value.replace(",","");
var F6 = form.frmBoardMembersFemaleNative.value.replace(",","");
var F7 = form.frmBoardMembersFemaleMulti.value.replace(",","");
var F8 = form.frmBoardMembersFemaleUnknown.value.replace(",","");
var M1 = form.frmBoardMembersMaleWhite.value.replace(",","");
var M2 = form.frmBoardMembersMaleBlack.value.replace(",","");
var M3 = form.frmBoardMembersMaleHispanic.value.replace(",","");
var M4 = form.frmBoardMembersMaleAsian.value.replace(",","");
var M5 = form.frmBoardMembersMaleIslander.value.replace(",","");
var M6 = form.frmBoardMembersMaleNative.value.replace(",","");
var M7 = form.frmBoardMembersMaleMulti.value.replace(",","");
var M8 = form.frmBoardMembersMaleUnknown.value.replace(",","");
var Total = form.frmBoardMembersNumberBoardMembers.value.replace(",","");

	if(!(myRegularExpression1.test(form.frmBoardMembersNumberBoardMembers.value)) || (form.frmBoardMembersNumberBoardMembers.value == ""))
	{
		form.frmBoardMembersNumberBoardMembers.focus();
		if(form.frmBoardMembersNumberBoardMembers.value == "")
			alert("Number of Board Members is invalid");
		else
			alert((form.frmBoardMembersNumberBoardMembers.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersTermLimitsYears.value)) || (form.frmBoardMembersTermLimitsYears.value == ""))
	{
		form.frmBoardMembersTermLimitsYears.focus();
		if(form.frmBoardMembersTermLimitsYears.value == "")
			alert("Question 2 is invalid");
		else
			alert((form.frmBoardMembersTermLimitsYears.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersAverageTenureYears.value)) || (form.frmBoardMembersAverageTenureYears.value == ""))
	{
		form.frmBoardMembersAverageTenureYears.focus();
		if(form.frmBoardMembersAverageTenureYears.value == "")
			alert("Tenure Years is invalid");
		else
			alert((form.frmBoardMembersAverageTenureYears.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersAverageTenureMonths.value)) || (form.frmBoardMembersAverageTenureMonths.value == ""))
	{
		form.frmBoardMembersAverageTenureMonths.focus();
		if((form.frmBoardMembersAverageTenureMonths.value == ""))
			alert("Tenure Months is invalid");
		else
			alert((form.frmBoardMembersAverageTenureMonths.value) + " is invalid.");
		return false;
	}
	else if((form.frmBoardMembersStandingCommitteesOther.checked == true) && ((form.frmBoardMembersStandingCommitteesOtherText.value == "(Name Here)") || (form.frmBoardMembersStandingCommitteesOtherText.value == "")))
	{
		form.frmBoardMembersStandingCommitteesOtherText.focus();
		alert("Please name standing committee.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersFemaleWhite.value)) || (form.frmBoardMembersFemaleWhite.value == ""))
	{
		form.frmBoardMembersFemaleWhite.focus();
		if(form.frmBoardMembersFemaleWhite.value == "")
			alert("Female Board Members cannot have an empty field");
		else
			alert((form.frmBoardMembersFemaleWhite.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersFemaleBlack.value)) || (form.frmBoardMembersFemaleBlack.value == ""))
	{
		form.frmBoardMembersFemaleBlack.focus();
		if(form.frmBoardMembersFemaleBlack.value == "")
			alert("Female Board Members cannot have an empty field");
		else
			alert((form.frmBoardMembersFemaleBlack.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersFemaleHispanic.value)) || (form.frmBoardMembersFemaleHispanic.value == ""))
	{
		form.frmBoardMembersFemaleHispanic.focus();
		if(form.frmBoardMembersFemaleHispanic.value == "")
			alert("Female Board Members cannot have an empty field");
		else
			alert((form.frmBoardMembersFemaleHispanic.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersFemaleAsian.value)) || (form.frmBoardMembersFemaleAsian.value == ""))
	{
		form.frmBoardMembersFemaleAsian.focus();
		if(form.frmBoardMembersFemaleAsian.value == "")
			alert("Female Board Members cannot have an empty field");
		else
			alert((form.frmBoardMembersFemaleAsian.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersFemaleIslander.value)) || (form.frmBoardMembersFemaleIslander.value == ""))
	{
		form.frmBoardMembersFemaleIslander.focus();
		if(form.frmBoardMembersFemaleIslander.value == "")
			alert("Female Board Members cannot have an empty field");
		else
			alert((form.frmBoardMembersFemaleIslander.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersFemaleNative.value)) || (form.frmBoardMembersFemaleNative.value == ""))
	{
		form.frmBoardMembersFemaleNative.focus();
		if(form.frmBoardMembersFemaleNative.value == "")
			alert("Female Board Members cannot have an empty field");
		else
			alert((form.frmBoardMembersFemaleNative.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersFemaleMulti.value)) || (form.frmBoardMembersFemaleMulti.value == ""))
	{
		form.frmBoardMembersFemaleMulti.focus();
		if(form.frmBoardMembersFemaleMulti.value == "")
			alert("Female Board Members cannot have an empty field");
		else
			alert((form.frmBoardMembersFemaleMulti.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersFemaleUnknown.value)) || (form.frmBoardMembersFemaleUnknown.value == ""))
	{
		form.frmBoardMembersFemaleUnknown.focus();
		if(form.frmBoardMembersFemaleUnknown.value == "")
			alert("Female Board Members cannot have an empty field");
		else
			alert((form.frmBoardMembersFemaleUnknown.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersMaleWhite.value)) || (form.frmBoardMembersMaleWhite.value == ""))
	{
		form.frmBoardMembersMaleWhite.focus();
		if(form.frmBoardMembersMaleWhite.value == "")
			alert("Male Board Members cannot have an empty field");
		else
			alert((form.frmBoardMembersMaleWhite.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersMaleBlack.value)) || (form.frmBoardMembersMaleBlack.value == ""))
	{
		form.frmBoardMembersMaleBlack.focus();
		if(form.frmBoardMembersMaleBlack.value == "")
			alert("Male Board Members cannot have an empty field");
		else
			alert((form.frmBoardMembersMaleBlack.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersMaleHispanic.value)) || (form.frmBoardMembersMaleHispanic.value == ""))
	{
		form.frmBoardMembersMaleHispanic.focus();
		if(form.frmBoardMembersMaleHispanic.value == "")
			alert("Male Board Members cannot have an empty field");
		else
			alert((form.frmBoardMembersMaleHispanic.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersMaleAsian.value)) || (form.frmBoardMembersMaleAsian.value == ""))
	{
		form.frmBoardMembersMaleAsian.focus();
		if(form.frmBoardMembersMaleAsian.value == "")
			alert("Male Board Members cannot have an empty field");
		else
			alert((form.frmBoardMembersMaleAsian.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersMaleIslander.value)) || (form.frmBoardMembersMaleIslander.value == ""))
	{
		form.frmBoardMembersMaleIslander.focus();
		if(form.frmBoardMembersMaleIslander.value == "")
			alert("Male Board Members cannot have an empty field");
		else
			alert((form.frmBoardMembersMaleIslander.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersMaleNative.value)) || (form.frmBoardMembersMaleNative.value == ""))
	{
		form.frmBoardMembersMaleNative.focus();
		if(form.frmBoardMembersMaleNative.value == "")
			alert("Male Board Members cannot have an empty field");
		else
			alert((form.frmBoardMembersMaleNative.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersMaleMulti.value)) || (form.frmBoardMembersMaleMulti.value == ""))
	{
		form.frmBoardMembersMaleMulti.focus();
		if(form.frmBoardMembersMaleMulti.value == "")
			alert("Male Board Members cannot have an empty field");
		else
			alert((form.frmBoardMembersMaleMulti.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersMaleUnknown.value)) || (form.frmBoardMembersMaleUnknown.value == ""))
	{
		form.frmBoardMembersMaleUnknown.focus();
		if(form.frmBoardMembersMaleUnknown.value == "")
			alert("Male Board Members cannot have an empty field");
		else
			alert((form.frmBoardMembersMaleUnknown.value) + " is invalid.");
		return false;
	}
	else if((parseInt(F1) + parseInt(F2) + parseInt(F3) + parseInt(F4) + parseInt(F5) + parseInt(F6) + parseInt(F7) + parseInt(F8) + parseInt(M1) + parseInt(M2) + parseInt(M3) + parseInt(M4) + parseInt(M5) + parseInt(M6) + parseInt(M7) + parseInt(M8)) != parseInt(Total))
	{
		form.frmBoardMembersNumberBoardMembers.focus();
		alert("Number of Board Members must equal the total of questions 5 and 6.");
		return false;
	}
	else if((form.frmBoardMembersFrequency[0].checked != true) && (form.frmBoardMembersFrequency[1].checked != true) && (form.frmBoardMembersFrequency[2].checked != true) && (form.frmBoardMembersFrequency[3].checked != true))
	{
		alert("Please make sure that you have selected the appropriate answer for question 7.");
		return false;
	}
	else if((form.frmBoardMembersFrequency[3].checked == true) && ((form.frmBoardMembersFrequencyOtherText.value == "") || (form.frmBoardMembersFrequencyOtherText.value == "(Please Enter)")))
	{
		form.frmBoardMembersFrequencyOtherText.focus();
		alert("Please enter frequency of board meetings.");
		return false;
	}
	else if((form.frmBoardMembersMoney[0].checked != true) && (form.frmBoardMembersMoney[1].checked != true) && (form.frmBoardMembersMoney[2].checked != true) && (form.frmBoardMembersMoney[3].checked != true))
	{
		alert("Please make sure that you have selected the appropriate answer for question 8.");
		return false;
	}
	else if((form.frmBoardMembersMoney[0].checked == true) && ((form.frmBoardMembersMoneyMinimumAmount.value == "") || (form.frmBoardMembersMoneyMinimumAmount.value == "0")))
	{
		form.frmBoardMembersMoneyMinimumAmount.focus();
		alert("Please enter amount of financial committment.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersMoneyMinimumAmount.value)))
	{
		form.frmBoardMembersMoneyMinimumAmount.focus();
		alert((form.frmBoardMembersMoneyMinimumAmount.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersYearlyContribution.value)) || (form.frmBoardMembersYearlyContribution.value == ""))
	{
		form.frmBoardMembersYearlyContribution.focus();
		if(form.frmBoardMembersYearlyContribution.value == "")
			alert("Board Members Contribution cannot be an empty field");
		else
			alert((form.frmBoardMembersYearlyContribution.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersSkillsFinance.value)) || (form.frmBoardMembersSkillsFinance.value == ""))
	{
		form.frmBoardMembersSkillsFinance.focus();
		if(form.frmBoardMembersSkillsFinance.value == "")
			alert("Question 10 cannot contain an empty field");
		else
			alert((form.frmBoardMembersSkillsFinance.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersSkillsLegal.value)) || (form.frmBoardMembersSkillsLegal.value == ""))
	{
		form.frmBoardMembersSkillsLegal.focus();
		if(form.frmBoardMembersSkillsLegal.value == "")
			alert("Question 10 cannot contain an empty field");
		else
			alert((form.frmBoardMembersSkillsLegal.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersSkillsPublicRelations.value)) || (form.frmBoardMembersSkillsPublicRelations.value == ""))
	{
		form.frmBoardMembersSkillsPublicRelations.focus();
		if(form.frmBoardMembersSkillsPublicRelations.value == "")
			alert("Question 10 cannot contain an empty field");
		else
			alert((form.frmBoardMembersSkillsPublicRelations.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersSkillsHumanServicesPractitioner.value)) || (form.frmBoardMembersSkillsHumanServicesPractitioner.value == ""))
	{
		form.frmBoardMembersSkillsHumanServicesPractitioner.focus();
		if(form.frmBoardMembersSkillsHumanServicesPractitioner.value == "")
			alert("Question 10 cannot contain an empty field");
		else
			alert((form.frmBoardMembersSkillsHumanServicesPractitioner.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersSkillsHumanServicesAdministrator.value)) || (form.frmBoardMembersSkillsHumanServicesAdministrator.value == ""))
	{
		form.frmBoardMembersSkillsHumanServicesAdministrator.focus();
		if(form.frmBoardMembersSkillsHumanServicesAdministrator.value == "")
			alert("Question 10 cannot contain an empty field");
		else
			alert((form.frmBoardMembersSkillsHumanServicesAdministrator.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersSkillsFullTimeStudent.value)) || (form.frmBoardMembersSkillsFullTimeStudent.value == ""))
	{
		form.frmBoardMembersSkillsFullTimeStudent.focus();
		if(form.frmBoardMembersSkillsFullTimeStudent.value == "")
			alert("Question 10 cannot contain an empty field");
		else
			alert((form.frmBoardMembersSkillsFullTimeStudent.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersSkillsHumanResources.value)) || (form.frmBoardMembersSkillsHumanResources.value == ""))
	{
		form.frmBoardMembersSkillsHumanResources.focus();
		if(form.frmBoardMembersSkillsHumanResources.value == "")
			alert("Question 10 cannot contain an empty field");
		else
			alert((form.frmBoardMembersSkillsHumanResources.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersSkillsCorporateCEO.value)) || (form.frmBoardMembersSkillsCorporateCEO.value == ""))
	{
		form.frmBoardMembersSkillsCorporateCEO.focus();
		if(form.frmBoardMembersSkillsCorporateCEO.value == "")
			alert("Question 10 cannot contain an empty field");
		else
			alert((form.frmBoardMembersSkillsCorporateCEO.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersSkillsOtherCorporateOfficer.value)) || (form.frmBoardMembersSkillsOtherCorporateOfficer.value == ""))
	{
		form.frmBoardMembersSkillsOtherCorporateOfficer.focus();
		if(form.frmBoardMembersSkillsOtherCorporateOfficer.value == "")
			alert("Question 10 cannot contain an empty field");
		else
			alert((form.frmBoardMembersSkillsOtherCorporateOfficer.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersSkillsInsurance.value)) || (form.frmBoardMembersSkillsInsurance.value == ""))
	{
		form.frmBoardMembersSkillsInsurance.focus();
		if(form.frmBoardMembersSkillsInsurance.value == "")
			alert("Question 10 cannot contain an empty field");
		else
			alert((form.frmBoardMembersSkillsInsurance.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersSkillsSmallBusiness.value)) || (form.frmBoardMembersSkillsSmallBusiness.value == ""))
	{
		form.frmBoardMembersSkillsSmallBusiness.focus();
		if(form.frmBoardMembersSkillsSmallBusiness.value == "")
			alert("Question 10 cannot contain an empty field");
		else
			alert((form.frmBoardMembersSkillsSmallBusiness.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersSkillsBig.value)) || (form.frmBoardMembersSkillsBig.value == ""))
	{
		form.frmBoardMembersSkillsBig.focus();
		if(form.frmBoardMembersSkillsBig.value == "")
			alert("Question 10 cannot contain an empty field");
		else
			alert((form.frmBoardMembersSkillsBig.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersSkillsParentLittle.value)) || (form.frmBoardMembersSkillsParentLittle.value == ""))
	{
		form.frmBoardMembersSkillsParentLittle.focus();
		if(form.frmBoardMembersSkillsParentLittle.value == "")
			alert("Question 10 cannot contain an empty field");
		else
			alert((form.frmBoardMembersSkillsParentLittle.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersSkillsLittle.value)) || (form.frmBoardMembersSkillsLittle.value == ""))
	{
		form.frmBoardMembersSkillsLittle.focus();
		if(form.frmBoardMembersSkillsLittle.value == "")
			alert("Question 10 cannot contain an empty field");
		else
			alert((form.frmBoardMembersSkillsLittle.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersSkillsLocalGovernment.value)) || (form.frmBoardMembersSkillsLocalGovernment.value == ""))
	{
		form.frmBoardMembersSkillsLocalGovernment.focus();
		if(form.frmBoardMembersSkillsLocalGovernment.value == "")
			alert("Question 10 cannot contain an empty field");
		else
			alert((form.frmBoardMembersSkillsLocalGovernment.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersSkillsOther.value)) || (form.frmBoardMembersSkillsOther.value == ""))
	{
		form.frmBoardMembersSkillsOther.focus();
		if(form.frmBoardMembersSkillsOther.value == "")
			alert("Question 10 cannot contain an empty field");
		else
			alert((form.frmBoardMembersSkillsOther.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBoardMembersSkillsUnknown.value)) || (form.frmBoardMembersSkillsUnknown.value == ""))
	{
		form.frmBoardMembersSkillsUnknown.focus();
		if(form.frmBoardMembersSkillsUnknown.value == "")
			alert("Question 10 cannot contain an empty field");
		else
			alert((form.frmBoardMembersSkillsUnknown.value) + " is invalid.");
		return false;
	}
	else
	{
		return true;
	}
}	
//-->
</script>

<% '<!--#include file="../includes/top_nav_forms_yearly.inc"--><!-- include file has </head> and <body> tags --><br>%>
<!--#include file="../includes/surveytitle.inc"-->

<table width="100%" cellpadding="0" cellspacing="0" border="0">
<td width="220" valign="top"><img src="../includes/images/photos_slinky.jpg" alt="" width="220" height="477" border="0"></td>
<td width="100%" valign="top">
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
<table width="640" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063">
<form name="frmBoardMembers" action="BoardMembers_edit.asp" method="post" onsubmit="return submitFormValidate(this);">
<!--#include file="../includes/form_stamp.asp"-->
<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmBoardMembers WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
	Set GetBoardMembers = Con.Execute(query)
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
			
				<tr> 
					<td colspan="3" align="center" valign="top" class="formSubhead">BBBS - <%= y %> Annual Agency Information (AAI)</td>
				</tr>
				<tr>
					<td colspan="3" class="formHeader">BOARD MEMBERS</td>
				</tr>
				<tr>
					<td colspan="3" class="formMain"><font color="#ff0000"><div align="center"><strong>Please Note: </strong>After entering your information, you <strong>must</strong> click on the "Save" button at the bottom of the form and wait for the "Thank You" screen or your changes will be lost.</div></font></td>
				</tr>				
<!-- Question Number 1 -->
				<tr>
					<td align="left" valign="top" class="formMain">1.</td>
					<td align="left" valign="top" class="formMain">Number of Board Members as of 12/31:</td>
					<td align="right" valign="top"><input type="text" class="formMain" size="18" value="<% If say = "edit" Then %><%= GetBoardMembers("NumberBoardMembers") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersNumberBoardMembers"></td>
				</tr>
<!-- Question Number 2 -->
				<tr> 
					<td align="left" valign="top" class="formMain">2.</td>
					<td align="left" valign="top" class="formMain">If you have term limits for Board Members, enter number of years:</td>
					<td align="right" valign="top"><input type="text" size="18" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("TermLimitsYears") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersTermLimitsYears"></td>
				</tr>
<!-- Question Number 3 -->
				<tr>
					<td align="left" valign="top" class="formMain">3.</td>
					<!-- the reason there are non breaking spaces in between each word in this field is that in netscape it makes the column expand -->
					<td align="left" valign="top" class="formMain" >What&nbsp;is&nbsp;the&nbsp;average&nbsp;tenure&nbsp;of&nbsp;your&nbsp;board&nbsp;members?</td>
					<td align="right" valign="top" class="formMain">
					&nbsp;Years&nbsp;&nbsp;&nbsp;<input type="text" class="formMain" size="3" value="<% If say = "edit" Then %><%= GetBoardMembers("AverageTenureYears") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersAverageTenureYears"><br>
					&nbsp;Months&nbsp;<input type="text" class="formMain" size="3" value="<% If say = "edit" Then %><%= GetBoardMembers("AverageTenureMonths") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersAverageTenureMonths"></td>
				</tr>
<!-- Question Number 4 -->
				<tr>
					<td align="left" valign="top" class="formMain">4.</td>
					<td colspan="2" align="left" valign="top" class="formMain">What standing committees do you have? (Please check all that apply.)<br>
<!-- nested table -->
						<table width="640" border="0" cellspacing="3" cellpadding="3">
							<tr>
								<td align="left" valign="top" class="formMain">1 <input type="checkbox"<% If (say = "edit") Then %><% If (GetBoardMembers("StandingCommitteesPersonnel") = True) Then %> checked<% End If %><% End If %> value="Yes" class="formMain" name="frmBoardMembersStandingCommitteesPersonnel">Personnel</td>
								<td align="left" valign="top" class="formMain">2 <input type="checkbox"<% If (say = "edit") Then %><% If (GetBoardMembers("StandingCommitteesProgram") = True) Then %> checked<% End If %><% End If %> value="Yes" class="formMain" name="frmBoardMembersStandingCommitteesProgram">Program</td>
								<td align="left" valign="top" class="formMain">3 <input type="checkbox"<% If (say = "edit") Then %><% If (GetBoardMembers("StandingCommitteesExecutive") = True) Then %> checked<% End If %><% End If %> value="Yes" class="formMain" name="frmBoardMembersStandingCommitteesExecutive">Executive</td>
								<td align="left" valign="top" class="formMain">4 <input type="checkbox"<% If (say = "edit") Then %><% If (GetBoardMembers("StandingCommitteesFundDevelopment") = True) Then %> checked<% End If %><% End If %> value="Yes" class="formMain" name="frmBoardMembersStandingCommitteesFundDevelopment">Fund Development</td>
							</tr>
							<tr>
								<td align="left" valign="top" class="formMain">5 <input type="checkbox"<% If (say = "edit") Then %><% If (GetBoardMembers("StandingCommitteesFinance") = True) Then %> checked<% End If %><% End If %> value="Yes" class="formMain" name="frmBoardMembersStandingCommitteesFinance">Finance</td>
								<td align="left" valign="top" class="formMain">6 <input type="checkbox"<% If (say = "edit") Then %><% If (GetBoardMembers("StandingCommitteesPublicRelations") = True) Then %> checked<% End If %><% End If %> value="Yes" class="formMain" name="frmBoardMembersStandingCommitteesPublicRelations">Public Relations</td>
								<td align="left" valign="top" class="formMain">7 <input type="checkbox"<% If (say = "edit") Then %><% If (GetBoardMembers("StandingCommitteesStrategicPlanning") = True) Then %> checked<% End If %><% End If %> value="Yes" class="formMain" name="frmBoardMembersStandingCommitteesStrategicPlanning">Strategic Planning</td>
								<td align="left" valign="top" class="formMain">8 <input type="checkbox"<% If (say = "edit") Then %><% If (GetBoardMembers("StandingCommitteesBoardDevelopment") = True) Then %> checked<% End If %><% End If %> value="Yes" class="formMain" name="frmBoardMembersStandingCommitteesBoardDevelopment">Board Development</td>
							</tr>
							<tr>
								<td align="left" valign="top" class="formMain">9 <input type="checkbox"<% If (say = "edit") Then %><% If (GetBoardMembers("StandingCommitteesVolunteerRecruitment") = True) Then %> checked<% End If %><% End If %> value="Yes" name="frmBoardMembersStandingCommitteesVolunteerRecruitment">Volunteer Recruitment</td>
								<td colspan="3" align="left" valign="top" class="formMain">10 <input type="checkbox"<% If (say = "edit") Then %><% If (GetBoardMembers("StandingCommitteesOther") = True) Then %> checked<% End If %><% End If %> value="Yes" class="formMain" name="frmBoardMembersStandingCommitteesOther">Other (Name): <input type="text" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("StandingCommitteesOtherText") %><% Else %>(Name Here)<% End If %>" name="frmBoardMembersStandingCommitteesOtherText" onblur="changeForm1();"></td>
							</tr>
						</table>
					</td>		
				</tr>
<!-- Question Number 5 -->
				<tr>
					<td align="left" valign="top" class="formMain">5.</td>
					<td colspan="2" align="left" valign="top" class="formMain">Enter the number of FEMALE board members by ethnicity below:<br>
<!-- nested table -->
						<table width="640" border="0" cellspacing="3" cellpadding="3" align="center">
							<tr>
								<td align="left" valign="top" class="formMain">White (Not Hispanic)<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("FemaleWhite") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersFemaleWhite"></td>
								<td align="left" valign="top" class="formMain">Black<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("FemaleBlack") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersFemaleBlack"></td>
								<td align="left" valign="top" class="formMain">Hispanic<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("FemaleHispanic") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersFemaleHispanic"></td>
								<td align="left" valign="top" class="formMain">Asian<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("FemaleAsian") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersFemaleAsian"></td>
								
							</tr>
							<tr>
								<td align="left" valign="top" class="formMain">Pacific Islander<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("FemaleIslander") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersFemaleIslander"></td>
								<td align="left" valign="top" class="formMain">Native American<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("FemaleNative") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersFemaleNative"></td>
								<td align="left" valign="top" class="formMain">Multi-Racial<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("FemaleMulti") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersFemaleMulti"></td>
								<td align="left" valign="top" class="formMain">Unknown<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("FemaleUnknown") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersFemaleUnknown"></td>
							</tr>
						</table> 
					</td>		
				</tr>
<!-- Question Number 6 -->
				<tr>
					<td align="left" valign="top" class="formMain">6.</td>
					<td colspan="2" align="left" valign="top" class="formMain">Enter the number of MALE board members by ethnicity below:<br>
<!-- nested table -->
						<table width="640" border="0" cellspacing="3" cellpadding="3" align="center">
							<tr>
								<td align="left" valign="top" class="formMain">White (Not Hispanic)<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("MaleWhite") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersMaleWhite"></td>
								<td align="left" valign="top" class="formMain">Black<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("MaleBlack") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersMaleBlack"></td>
								<td align="left" valign="top" class="formMain">Hispanic<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("MaleHispanic") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersMaleHispanic"></td>
								<td align="left" valign="top" class="formMain">Asian<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("MaleAsian") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersMaleAsian"></td>
							</tr>
							<tr>
								<td align="left" valign="top" class="formMain">Pacific Islander<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("MaleIslander") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersMaleIslander"></td>
								<td align="left" valign="top" class="formMain">Native American<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("MaleNative") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersMaleNative"></td>
								<td align="left" valign="top" class="formMain">Multi-Racial<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("MaleMulti") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersMaleMulti"></td>
								<td align="left" valign="top" class="formMain">Unknown<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("MaleUnknown") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersMaleUnknown"></td>
							</tr>
						</table>
					</td>		
				</tr>
<!-- Question Number 7 -->
				<tr>
					<td align="left" valign="top" class="formMain">7.</td>
					<td colspan="2" align="left" valign="top" class="formMain">Frequency of board meetings.<br>
<!-- nested table -->
						<table width="640" border="0" cellspacing="3" cellpadding="3" align="center">
							<tr>
								<td align="center" valign="bottom" class="formMain"><input type="radio"<% If (say = "edit") Then %><% If (GetBoardMembers("FrequencyMonthly") = True) Then %> checked<% End If %><% End If %> value="Monthly" class="formMain" name="frmBoardMembersFrequency" onclick="form.frmBoardMembersFrequencyOtherText.value = '(Please Enter)'">Monthly</td>
								<td align="center" valign="bottom" class="formMain"><input type="radio"<% If (say = "edit") Then %><% If (GetBoardMembers("FrequencyTwoMonths") = True) Then %> checked<% End If %><% End If %> value="TwoMonths" class="formMain" name="frmBoardMembersFrequency" onclick="form.frmBoardMembersFrequencyOtherText.value = '(Please Enter)'">Every 2 Months</td>
								<td align="center" valign="bottom" class="formMain"><input type="radio"<% If (say = "edit") Then %><% If (GetBoardMembers("FrequencyQuarterly") = True) Then %> checked<% End If %><% End If %> value="Quarterly" class="formMain" name="frmBoardMembersFrequency" onclick="form.frmBoardMembersFrequencyOtherText.value = '(Please Enter)'">Quarterly</td>
								<td align="center" valign="bottom" class="formMain"><input type="radio"<% If (say = "edit") Then %><% If (GetBoardMembers("FrequencyOther") = True) Then %> checked<% End If %><% End If %> value="Other" class="formMain" name="frmBoardMembersFrequency">Other: <input type="text" size="18" class="formMain" name="frmBoardMembersFrequencyOtherText" value="<% If say = "edit" Then %><%= GetBoardMembers("FrequencyOtherText") %><% Else %>(Please Enter)<% End If %>" onblur="changeForm2();"></td>
							</tr> 
						</table> 
					</td>		
				</tr>
<!-- Question Number 8 --> 
				<tr> 
					<td align="left" valign="top" class="formMain">8.</td>
					<td colspan="3" align="left" valign="top" class="formMain">Are all board members expected to:<br> 
						a. <input type="radio"<% If (say = "edit") Then %><% If (GetBoardMembers("MoneyMinimum") = True) Then %> checked<% End If %><% End If %> value="Minimum" class="formMain" name="frmBoardMembersMoney"> Make a minimum annual financial commitment? If checked, indicate the amount:&nbsp;&nbsp;&#36;&nbsp;<input type="text" size="18" value="<% If say = "edit" Then %><%= GetBoardMembers("MoneyMinimumAmount") %><% Else %>0<% End If %>" onblur="checkForInteger(this.value);changeForm3();" class="formMain" name="frmBoardMembersMoneyMinimumAmount"><br>
						b. <input type="radio"<% If (say = "edit") Then %><% If (GetBoardMembers("MoneyInKind") = True) Then %> checked<% End If %><% End If %> value="InKind" class="formMain" name="frmBoardMembersMoney" onclick="form.frmBoardMembersMoneyMinimumAmount.value = 0"> Make either a monetary or in-kind contribution - no specified amount<br>
						c. <input type="radio"<% If (say = "edit") Then %><% If (GetBoardMembers("MoneyNotExpected") = True) Then %> checked<% End If %><% End If %> value="NotExpected" class="formMain" name="frmBoardMembersMoney" onclick="form.frmBoardMembersMoneyMinimumAmount.value = 0"> Not expected, but encouraged<br>
						d. <input type="radio"<% If (say = "edit") Then %><% If (GetBoardMembers("MoneyNoPolicy") = True) Then %> checked<% End If %><% End If %> value="NoPolicy" class="formMain" name="frmBoardMembersMoney" onclick="form.frmBoardMembersMoneyMinimumAmount.value = 0"> No policy or expectations
					</td>
				</tr>
<!-- Question Number 9 -->
				<tr> 
					<td align="left" valign="top" class="formMain">9.</td>
					<td align="left" valign="top" class="formMain">How much money did your board members contribute this past year?</td>
					<td align="right" valign="top" class="formMain">&nbsp;&nbsp;&#36;&nbsp;<input type="text" size="18" value="<% If say = "edit" Then %><%= GetBoardMembers("YearlyContribution") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" class="formMain" name="frmBoardMembersYearlyContribution"></td>
				</tr>	
<!-- Question Number 10 -->
				<tr>
					<td align="left" valign="top" class="formMain">10.</td>
					<td colspan="2" align="left" valign="top" class="formMain">Please enter the number of board members by professional skills/expertise below.<br>
<!-- nested table -->
						<table width="640" border="0" cellspacing="1" cellpadding="1" align="center">
							<tr>
								<td align="left" valign="top" class="formMain">Finance/Accounting/Banking<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("SkillsFinance") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersSkillsFinance"></td>
								<td align="left" valign="top" class="formMain">Legal<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("SkillsLegal") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersSkillsLegal"></td>
								<td align="left" valign="top" class="formMain">Public Relations<sup>1</sup><br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("SkillsPublicRelations") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersSkillsPublicRelations"></td>
								<td colspan="2" align="left" valign="top" class="formMain">Human Services Practitioner<sup>2</sup><br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("SkillsHumanServicesPractitioner") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersSkillsHumanServicesPractitioner"></td>								
							</tr>
							<tr>
								<td align="left" valign="top" class="formMain">Human Services Administrator<sup>2</sup><br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("SkillsHumanServicesAdministrator") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersSkillsHumanServicesAdministrator"></td>
								<td align="left" valign="top" class="formMain">Full Time College/H.S. Student<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("SkillsFulLTimeStudent") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersSkillsFullTimeStudent"></td>
								<td align="left" valign="top" class="formMain">Human Resources<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("SkillsHumanResources") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersSkillsHumanResources"></td>
								<td colspan="2" align="left" valign="top" class="formMain">Corporate CEO<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("SkillsCorporateCEO") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersSkillsCorporateCEO"></td>
							</tr>
							<tr>
								<td align="left" valign="top" class="formMain">Other Corporate Officer<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("SkillsOtherCorporateOfficer") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersSkillsOtherCorporateOfficer"></td>
								<td align="left" valign="top" class="formMain">Insurance/Sales<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("SkillsInsurance") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersSkillsInsurance"></td>
								<td align="left" valign="top" class="formMain">Small Business Owner<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("SkillsSmallBusiness") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersSkillsSmallBusiness"></td><br>
								<td colspan="2" align="left" valign="top" class="formMain">Big<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("SkillsBig") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersSkillsBig"></td>								
							</tr>
							<tr>
								<td align="left" valign="top" class="formMain">Parent of Little<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("SkillsParentLittle") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersSkillsParentLittle"></td>
								<td align="left" valign="top" class="formMain">Little<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("SkillsLittle") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersSkillsLittle"></td>
								<td align="left" valign="top" class="formMain">Local Government<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("SkillsLocalGovernment") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersSkillsLocalGovernment"></td>
								<td align="left" valign="top" class="formMain">Other<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("SkillsOther") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersSkillsOther"></td>
								<td align="left" valign="top" class="formMain">Unknown<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("SkillsUnknown") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersSkillsUnknown"></td>
							</tr>
							
						</table> 
					</td>		
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td align="left" valign="top" class="formSubHead" colspan="2">(1) Public Relations Includes: Marketing, Communications, Graphic Design<br>(2) Human Services Includes: Teacher, Psychologist, Social Worker</td>
				</tr>
				<tr>
					<td colspan="3" class="formHeader"><input type="submit" value="Save Form" class="formMainBold"></td>
				</tr>
				<tr>
					<td colspan="3"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
				</tr>
			</table>	
			
<% 
If say = "edit" Then
	GetBoardMembers.Close
	Set GetBoardMembers = Nothing
	Con.Close
	Set Con = Nothing
End If
 %>


</form>
<% End If %>
	<p>&nbsp;</p>
	<p>&nbsp;</p>   	
</td>
</tr>
</table>
</body>
</html>
