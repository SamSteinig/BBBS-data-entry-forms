
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
		
			ft = Int(Request("frmBoardMembersFemaleWhite")) _
			+ Int(Request("frmBoardMembersFemaleBlack")) _
			+ Int(Request("frmBoardMembersFemaleHispanic")) _
			+ Int(Request("frmBoardMembersFemaleAsian")) _
			+ Int(Request("frmBoardMembersFemaleIslander")) _
			+ Int(Request("frmBoardMembersFemaleNative")) _
			+ Int(Request("frmBoardMembersFemaleMulti")) _			
			+ Int(Request("frmBoardMembersFemaleUnknown"))		
		
		RST("FemaleWhite") = Request("frmBoardMembersFemaleWhite")
		RST("FemaleBlack") = Request("frmBoardMembersFemaleBlack")
		RST("FemaleHispanic") = Request("frmBoardMembersFemaleHispanic")
		RST("FemaleAsian") = Request("frmBoardMembersFemaleAsian")
		RST("FemaleIslander") = Request("frmBoardMembersFemaleIslander")
		RST("FemaleNative") = Request("frmBoardMembersFemaleNative")
		RST("FemaleMulti") = Request("frmBoardMembersFemaleMulti")
		RST("FemaleUnknown") = Request("frmBoardMembersFemaleUnknown")
		
		RST("FemaleTotal") = ft
		
			mt = Int(Request("frmBoardMembersMaleWhite")) _
			+ Int(Request("frmBoardMembersMaleBlack")) _
			+ Int(Request("frmBoardMembersMaleHispanic")) _
			+ Int(Request("frmBoardMembersMaleAsian")) _
			+ Int(Request("frmBoardMembersMaleIslander")) _
			+ Int(Request("frmBoardMembersMaleNative")) _
			+ Int(Request("frmBoardMembersMaleMulti")) _			
			+ Int(Request("frmBoardMembersMaleUnknown"))			
		
		RST("MaleWhite") = Request("frmBoardMembersMaleWhite")
		RST("MaleBlack") = Request("frmBoardMembersMaleBlack")
		RST("MaleHispanic") = Request("frmBoardMembersMaleHispanic")
		RST("MaleAsian") = Request("frmBoardMembersMaleAsian")
		RST("MaleIslander") = Request("frmBoardMembersMaleIslander")
		RST("MaleNative") = Request("frmBoardMembersMaleNative")
		RST("MaleMulti") = Request("frmBoardMembersMaleMulti")
		RST("MaleUnknown") = Request("frmBoardMembersMaleUnknown")
		
		RST("MaleTotal") = mt
		
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

		If Request("frmBoardMembersBoardDonatingPolicy") = "Yes" Then
			RST("BoardDonatingPolicy") = True
		Else
			RST("BoardDonatingPolicy") = False
		End If
		RST("MinimumBoardDonation") = formatcurrency(request("frmBoardMembersMinimumBoardDonation"))
		RST("BoardBigs") = Request("frmBoardMembersBoardBigs")
	    RST("BoardDonationPrcnt") = Request("frmBoardDonationPrcnt")
		RST("BoardConnectedPrcnt") = Request("frmBoardConnectedPrcnt")
		RST("AvgDonationBoardMember") = Request("frmAvgDonationBoardMember")
		If Request("frmBoardDevelopmentPlan") = "Yes" Then
			RST("BoardDevelopmentPlan") = True
		Else
			RST("BoardDevelopmentPlan") = False
		End If		
		If Request("frmAssessmentDone") = "Yes" Then
			RST("AssessmentDone") = True
		Else
			RST("AssessmentDone") = False
		End If
		
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
		
			ft = Int(Request("frmBoardMembersFemaleWhite")) _
			+ Int(Request("frmBoardMembersFemaleBlack")) _
			+ Int(Request("frmBoardMembersFemaleHispanic")) _
			+ Int(Request("frmBoardMembersFemaleAsian")) _
			+ Int(Request("frmBoardMembersFemaleIslander")) _
			+ Int(Request("frmBoardMembersFemaleNative")) _
			+ Int(Request("frmBoardMembersFemaleMulti")) _			
			+ Int(Request("frmBoardMembersFemaleUnknown"))		
		
		RST("FemaleWhite") = Request("frmBoardMembersFemaleWhite")
		RST("FemaleBlack") = Request("frmBoardMembersFemaleBlack")
		RST("FemaleHispanic") = Request("frmBoardMembersFemaleHispanic")
		RST("FemaleAsian") = Request("frmBoardMembersFemaleAsian")
		RST("FemaleIslander") = Request("frmBoardMembersFemaleIslander")
		RST("FemaleNative") = Request("frmBoardMembersFemaleNative")
		RST("FemaleMulti") = Request("frmBoardMembersFemaleMulti")
		RST("FemaleUnknown") = Request("frmBoardMembersFemaleUnknown")

		RST("FemaleTotal") = ft
		
			mt = Int(Request("frmBoardMembersMaleWhite")) _
			+ Int(Request("frmBoardMembersMaleBlack")) _
			+ Int(Request("frmBoardMembersMaleHispanic")) _
			+ Int(Request("frmBoardMembersMaleAsian")) _
			+ Int(Request("frmBoardMembersMaleIslander")) _
			+ Int(Request("frmBoardMembersMaleNative")) _
			+ Int(Request("frmBoardMembersMaleMulti")) _			
			+ Int(Request("frmBoardMembersMaleUnknown"))			
		
		
		RST("MaleWhite") = Request("frmBoardMembersMaleWhite")
		RST("MaleBlack") = Request("frmBoardMembersMaleBlack")
		RST("MaleHispanic") = Request("frmBoardMembersMaleHispanic")
		RST("MaleAsian") = Request("frmBoardMembersMaleAsian")
		RST("MaleIslander") = Request("frmBoardMembersMaleIslander")
		RST("MaleNative") = Request("frmBoardMembersMaleNative")
		RST("MaleMulti") = Request("frmBoardMembersMaleMulti")
		RST("MaleUnknown") = Request("frmBoardMembersMaleUnknown")
		
		RST("MaleTotal") = mt
		
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

		If Request("frmBoardMembersBoardDonatingPolicy") = "Yes" Then
			RST("BoardDonatingPolicy") = True
		Else
			RST("BoardDonatingPolicy") = False
		End If
		RST("MinimumBoardDonation") = formatcurrency(request("frmBoardMembersMinimumBoardDonation"))
		RST("BoardBigs") = Request("frmBoardMembersBoardBigs")	
		RST("BoardDonationPrcnt") = Request("frmBoardDonationPrcnt")
		RST("BoardConnectedPrcnt") = Request("frmBoardConnectedPrcnt")
		RST("AvgDonationBoardMember") = Request("frmAvgDonationBoardMember")
		If Request("frmBoardDevelopmentPlan") = "Yes" Then
			RST("BoardDevelopmentPlan") = True
		Else
			RST("BoardDevelopmentPlan") = False
		End If
			If Request("frmAssessmentDone") = "Yes" Then
			RST("AssessmentDone") = True
		Else
			RST("AssessmentDone") = False
		End If 				
	
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
function NewWindow(mypage, myname, w, h)
{
	var winl = (screen.width - w) / 2;
	var wint = (screen.height - h) / 2;
	winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',resizable, scrollbars'
	win = window.open(mypage, myname, winprops)
	if (parseInt(navigator.appVersion) >= 4) { win.window.focus(); }
}
	
function checkForInteger(valueToCheck)
{
	var myRegularExpression = /^[0-9]+(,[0-9]{3})*$/;  // Checks for integer with or without commas
	if(!(myRegularExpression.test(valueToCheck)))
	{
		alert("Please make sure you have entered a whole number with no spaces.\n We cannot process letters or words."); 
	} 
}



function changeForm2()
{
	if((document.frmBoardMembers.frmBoardMembersFrequencyOtherText.value != "") && (document.frmBoardMembers.frmBoardMembersFrequencyOtherText.value != "(Please Enter)"))
	{
		document.frmBoardMembers.frmBoardMembersFrequency[3].checked = true;
	}
}

function AddUpFemaleBoard() 
{
	var box1 = Number(document.frmBoardMembers.frmBoardMembersFemaleWhite.value)
	var box2 = Number(document.frmBoardMembers.frmBoardMembersFemaleBlack.value)
	var box3 = Number(document.frmBoardMembers.frmBoardMembersFemaleHispanic.value)	
	var box4 = Number(document.frmBoardMembers.frmBoardMembersFemaleAsian.value)
	var box5 = Number(document.frmBoardMembers.frmBoardMembersFemaleIslander.value)	
	var box6 = Number(document.frmBoardMembers.frmBoardMembersFemaleNative.value)		
	var box7 = Number(document.frmBoardMembers.frmBoardMembersFemaleMulti.value)			
	var box8 = Number(document.frmBoardMembers.frmBoardMembersFemaleUnknown.value)			
	var boxtotal = box1 + box2 + box3 + box4 + box5 + box6 + box7 + box8
	document.frmBoardMembers.frmBoardMembersFemaleTotal.value = boxtotal
}


function AddUpMaleBoard() 
{
	var box1 = Number(document.frmBoardMembers.frmBoardMembersMaleWhite.value)
	var box2 = Number(document.frmBoardMembers.frmBoardMembersMaleBlack.value)
	var box3 = Number(document.frmBoardMembers.frmBoardMembersMaleHispanic.value)	
	var box4 = Number(document.frmBoardMembers.frmBoardMembersMaleAsian.value)
	var box5 = Number(document.frmBoardMembers.frmBoardMembersMaleIslander.value)	
	var box6 = Number(document.frmBoardMembers.frmBoardMembersMaleNative.value)		
	var box7 = Number(document.frmBoardMembers.frmBoardMembersMaleMulti.value)			
	var box8 = Number(document.frmBoardMembers.frmBoardMembersMaleUnknown.value)			
	var boxtotal = box1 + box2 + box3 + box4 + box5 + box6 + box7 + box8
	document.frmBoardMembers.frmBoardMembersMaleTotal.value = boxtotal
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
	else if(!(myRegularExpression1.test(form.frmBoardMembersBoardBigs.value)) || (form.frmBoardMembersBoardBigs.value == ""))
	{
		form.frmBoardMembersBoardBigs.focus();
		if(form.frmBoardMembersBoardBigs.value == "")
			alert("Number of Board Members who are or have been a Big must not be left blank.");
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
		alert("Number of Board Members must equal the total of questions 2 and 3.");
		return false;
	}
	
	
	else if(new Number(form.frmBoardMembersBoardBigs.value) > new Number(form.frmBoardMembersNumberBoardMembers.value))
	{
		form.frmBoardMembersBoardBigs.focus();	
		alert("The number of Board Members who are Bigs (" + new Number(form.frmBoardMembersBoardBigs.value) + ") cannot be greater than the total number of Board Members (" + new Number(form.frmBoardMembersNumberBoardMembers.value) + ").");
		return false;
	}		

	/*else if((form.frmBoardMembersBoardDonatingPolicy[0].checked == true)&&(form.frmBoardMembersMinimumBoardDonation.value == 0))
	{
		form.frmBoardMembersMinimumBoardDonation.focus();	
		alert("You selected that you have 100% Board Donating Policy, but did not enter Minimum Donation Amount. Please enter required data or select that you do NOT have such policy.");
		return false;
	}*/
	
/*	else if((form.frmBoardMembersBoardDonatingPolicy[1].checked == true)&&(form.frmBoardMembersMinimumBoardDonation.value > 0))
	{
		form.frmBoardMembersMinimumBoardDonation.focus();	
		alert("You selected that you do NOT have 100% Board Donating Policy, but entered Minimum Donation Amount. Please enter 0 a or select that you have such policy.");
		return false;
	}*/


//Validation for Question 4--

	
  else if(!(myRegularExpression1.test(form.frmBoardDonationPrcnt.value))) 
	{
		form.frmBoardDonationPrcnt.focus();
		alert((form.frmBoardDonationPrcnt.value.value) + " is invalid. Please enter a whole number between 0 and 100.");
		
			
	return false;
	}
	
//Validation for Question 5--

	
  else if(!(myRegularExpression1.test(form.frmBoardConnectedPrcnt.value))) 
	{
		form.frmBoardConnectedPrcnt.focus();
		alert((form.frmBoardConnectedPrcnt.value.value) + " is invalid. Please enter a whole number between 0 and 100.");
		
			
	return false;
	}
	
	
//Validation for Question 6--

	
  else if(!(myRegularExpression1.test(form.frmAvgDonationBoardMember.value))) 
	{
		form.frmAvgDonationBoardMember.focus();
		alert((form.frmAvgDonationBoardMember.value.value) + " is invalid. Please enter a whole number.");
		
			
	return false;
	}
	
// Validation for Question 7 -- 		
	
	else if((form.frmBoardDevelopmentPlan[0].checked != true) && (form.frmBoardDevelopmentPlan[1].checked != true))
	{
		form.frmBoardDevelopmentPlan.focus();	
		alert("This field is required");
		return false;
	}	
	
// Validation for Question 8 -- 		
	
	else if((form.frmAssessmentDone[0].checked != true) && (form.frmAssessmentDone[1].checked != true))
	{
		form.frmAssessmentDone.focus();	
		alert("This field is required");
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
				<tr>
					<td colspan="3" class="formMain"><font color="#ff0000"><div align="center">If you need help with understanding the topic, please click on <img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"> next to the topic of your interest.<br><strong>The data you enter should reflect your employee data on 6/30/09.</strong></font></td>
				</tr>			
<!-- Question Number 1 -->
				<tr>
					<td align="left" valign="top" class="formMain">1.</td>
					<td align="left" valign="top" class="formMain"><a href="../helpfiles/surveyhelp.asp?HelpID=GoverningBoardMembers" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a> Number of Board Members:</td>
					<td align="right" valign="top"><input type="text" class="formMain" size="18" value="<% If say = "edit" Then %><%= GetBoardMembers("NumberBoardMembers") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersNumberBoardMembers"></td>
				</tr>



<!-- Question Number 2 -->
				<tr>
					<td align="left" valign="top" class="formMain">2.</td>
					<td colspan="2" align="left" valign="top" class="formMain">Enter the number of FEMALE board members by ethnicity below:<br>
<!-- nested table -->
						<table width="640" border="0" cellspacing="3" cellpadding="3" align="center">
							<tr>
								<td align="left" valign="top" class="formMain">White (Not Hispanic or Latino)<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("FemaleWhite") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" onFocus="AddUpFemaleBoard();" name="frmBoardMembersFemaleWhite"></td>
								<td align="left" valign="top" class="formMain">Black or African American (Not Hispanic or Latino)<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("FemaleBlack") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);"  onFocus="AddUpFemaleBoard();" name="frmBoardMembersFemaleBlack"></td>
								<td align="left" valign="top" class="formMain">Hispanic or Latino<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("FemaleHispanic") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);"  onFocus="AddUpFemaleBoard();" name="frmBoardMembersFemaleHispanic"></td>
								<td align="left" valign="top" class="formMain">Asian (Not Hispanic or Latino)<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("FemaleAsian") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);"  onFocus="AddUpFemaleBoard();" name="frmBoardMembersFemaleAsian"></td>
								
							</tr>

							<tr>
								<td align="left" valign="top" class="formMain">Native Hawaiian or Other Pacific Islander(Not Hispanic or Latino)<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("FemaleIslander") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);"  onFocus="AddUpFemaleBoard();" name="frmBoardMembersFemaleIslander"></td>
								<td align="left" valign="top" class="formMain">American Indian or Alaska Native (Not Hispanic or Latino)<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("FemaleNative") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);"  onFocus="AddUpFemaleBoard();" name="frmBoardMembersFemaleNative"></td>
								<td align="left" valign="top" class="formMain">Two or More Races (Not Hispanic or Latino)<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("FemaleMulti") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);"  onFocus="AddUpFemaleBoard();" name="frmBoardMembersFemaleMulti"></td>
								<td align="left" valign="top" class="formMain">Race missing or Unknown<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("FemaleUnknown") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);"  onFocus="AddUpFemaleBoard();" name="frmBoardMembersFemaleUnknown"></td>
							</tr>
							
							<tr>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td align="left" valign="top" class="formMain" bgcolor="#c0c0c0"><strong>Total:</strong><span class="formSubHead"><br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("FemaleTotal") %><%Else%>0<% End If %>" name="frmBoardMembersFemaleTotal" readonly onfocus="AddUpFemaleBoard();"><br><strong>calculated</strong></span></td>																								
							</tr>							
						</table> 
					</td>		
				</tr>
<!-- Question Number 3 -->
				<tr>
					<td align="left" valign="top" class="formMain">3.</td>
					<td colspan="2" align="left" valign="top" class="formMain">Enter the number of MALE board members by ethnicity below:<br>
<!-- nested table -->
						<table width="640" border="0" cellspacing="3" cellpadding="3" align="center">
							<tr>
								<td align="left" valign="top" class="formMain">White (Not Hispanic or Latino)<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("MaleWhite") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" onFocus="AddUpMaleBoard();" name="frmBoardMembersMaleWhite"></td>
								<td align="left" valign="top" class="formMain">Black or African American (Not Hispanic or Latino)<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("MaleBlack") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);"  onFocus="AddUpMaleBoard();" name="frmBoardMembersMaleBlack"></td>
								<td align="left" valign="top" class="formMain">Hispanic or Latino<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("MaleHispanic") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);"  onFocus="AddUpMaleBoard();" name="frmBoardMembersMaleHispanic"></td>
								<td align="left" valign="top" class="formMain">Asian (Not Hispanic or Latino)<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("MaleAsian") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);"  onFocus="AddUpMaleBoard();" name="frmBoardMembersMaleAsian"></td>
							</tr>
							<tr>
								<td align="left" valign="top" class="formMain">Native Hawaiian or Other Pacific Islander(Not Hispanic or Latino)<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("MaleIslander") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);"  onFocus="AddUpMaleBoard();" name="frmBoardMembersMaleIslander"></td>
								<td align="left" valign="top" class="formMain">American Indian or Alaska Native (Not Hispanic or Latino)<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("MaleNative") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);"  onFocus="AddUpMaleBoard();" name="frmBoardMembersMaleNative"></td>
								<td align="left" valign="top" class="formMain">Two or More Races (Not Hispanic or Latino)<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("MaleMulti") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);"  onFocus="AddUpMaleBoard();" name="frmBoardMembersMaleMulti"></td>
								<td align="left" valign="top" class="formMain">Race missing or Unknown<br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("MaleUnknown") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);"  onFocus="AddUpMaleBoard();" name="frmBoardMembersMaleUnknown"></td>
							</tr>
							
							<tr>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td align="left" valign="top" class="formMain" bgcolor="#c0c0c0"><strong>Total:</strong><span class="formSubHead"><br><input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("MaleTotal") %><%Else%>0<% End If %>" name="frmBoardMembersMaleTotal"  onFocus="AddUpMaleBoard();" readonly><br><strong>calculated</strong></span></td>								
							</tr>
														
						</table>
					</td>		
				</tr>
<!--Question 4 is not being asked from 2008 onwards
<!-- Question Number 4 --> 
			<!--	<tr> 
					<td align="left" valign="top" class="formMain">4.</td>
					<td colspan="3" align="left" valign="top" class="formMain">Do you have a policy of 100% Board donating?<br> 
						<input type="radio"<% If (say = "edit") Then %><% If (GetBoardMembers("BoardDonatingPolicy") = True) Then %> checked<% End If %><% End If %> value="Yes" class="formMain" name="frmBoardMembersBoardDonatingPolicy"> Yes <br>
						<input type="radio"<% If (say = "edit") Then %><% If (GetBoardMembers("BoardDonatingPolicy") = False) Then %> checked<% End If %><% End If %> value="No" class="formMain" name="frmBoardMembersBoardDonatingPolicy"> No <br>
						If Yes, Minimum Amount: $<input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("MinimumBoardDonation") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersMinimumBoardDonation">
						&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=Board100donate" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>
					</td>
				</tr>
				
				
				
<! New 4 Questions are being asked from 2008 on wards
				
<!-- Question Number 4 -->
         <tr> 
			<td align="left" valign="middle" class="formMain">4.</td>
			<td colspan="3" align="left" valign="top" class="formMain"> <a href="../helpfiles/StaffFormHelp.asp?HelpID=Percent of your Board Members are Donationg to your Agency" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;What Percent of your Board Members are personally donating to your Agency?&nbsp; 
						<input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("BoardDonationPrcnt")%><% Else %>0<% End If %>" onchange="checkForInteger(this.value);equalsLessThan101(this.value);" name="frmBoardDonationPrcnt">%</td>

        </tr>
				
				
			
<!-- Question Number 5 -->
         <tr> 
			<td align="left" valign="middle" class="formMain">5.</td>
			<td colspan="3" align="left" valign="top" class="formMain"> <a href="../helpfiles/StaffFormHelp.asp?HelpID=Percent of your Board has connected the agency" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;What Percent of your Board has connected the agency to potential Corporate and Individual Donors?&nbsp; 
						<input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("BoardConnectedPrcnt")%><% Else %>0<% End If %>" onchange="checkForInteger(this.value);equalsLessThan101(this.value);" name="frmBoardConnectedPrcnt">%</td>

        </tr>
							
				
<!-- Question Number 6 -->
         <tr> 
			<td align="left" valign="middle" class="formMain">6.</td>
			<td colspan="3" align="left" valign="top" class="formMain"> <a href="../helpfiles/StaffFormHelp.asp?HelpID=Average Donation by Board Member" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;What is the Average Annual Donation by Board Member?(Do not include the value of In-Kind services)$&nbsp; 
						<input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("AvgDonationBoardMember")%><% Else %>0<% End If %>" onchange="checkForInteger(this.value);"name="frmAvgDonationBoardMember">

        </tr>			
				
<!-- Question Number 7--> 
		<tr> 
		<td align="left" valign="top" class="formMain">7.</td>
		<td colspan="3" align="left" valign="top" class="formMain">Does your Agency have an updated Board Development Plan in Place?<br> 
						<input type="radio"<% If (say = "edit") Then %><% If (GetBoardMembers("BoardDevelopmentPlan") = True) Then %> checked<% End If %><% End If %> value="Yes" class="formMain" name="frmBoardDevelopmentPlan"> Yes <br>
						<input type="radio"<% If (say = "edit") Then %><% If (GetBoardMembers("BoardDevelopmentPlan") = False) Then %> checked<% End If %><% End If %> value="No" class="formMain" name="frmBoardDevelopmentPlan"> No <br> 
				</tr>
				
<!-- Question Number 8 --> 							
 	<tr> 
		<td align="left" valign="top" class="formMain">8.</td>
		<td colspan="3" align="left" valign="top" class="formMain">Has your board done an assessment in the last year?<br> 
						<input type="radio"<% If (say = "edit") Then %><% If (GetBoardMembers("AssessmentDone") = True) Then %> checked<% End If %><% End If %> value="Yes" class="formMain" name="frmAssessmentDone"> Yes <br>
						<input type="radio"<% If (say = "edit") Then %><% If (GetBoardMembers("AssessmentDone") = False) Then %> checked<% End If %><% End If %> value="No" class="formMain" name="frmAssessmentDone"> No <br> 
				</tr>				
												
<!-- Question Number 9 --> 
				<tr> 
					<td align="left" valign="middle" class="formMain">9.</td>
					<td colspan="3" align="left" valign="top" class="formMain">Number of Board Members as of June 30th of current AAI year who are or have been a BIG:&nbsp; 
						<input type="text" size="5" class="formMain" value="<% If say = "edit" Then %><%= GetBoardMembers("BoardBigs") %><% Else %>0<% End If %>" onchange="checkForInteger(this.value);" name="frmBoardMembersBoardBigs">
					</td>
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
