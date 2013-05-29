<!--#include file="../includes/NAD_BE.asp" -->

<% 

If Request("status") = "addNew" Then

	
	
' Check for duplicate records
	
	Set DupCon = Server.CreateObject("ADODB.Connection")
	DupCon.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT Count(*) As NumberOfEntries FROM tbl_frmFinancePerformance WHERE AgencyID = '" & Request("AgencyIDN") & "' and Year = " & Request("Year")	& " and Month = " & Request("Month")
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
		RST.Open "SELECT * FROM tbl_frmFinancePerformance", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = Request("AgencyIDN")
		RST("Year") = Request("year")
		RST("Month") = Request("month")
		tt = Int(Request("frmFinancePerformanceUnitedWay")) _ 
		+ Int(Request("frmFinancePerformanceGovFederalFunding"))_
		+ Int(Request("frmFinancePerformanceGovStateFunding"))_			
		+ Int(Request("frmFinancePerformanceGovLocalFunding"))_						
		+ Int(Request("frmFinancePerformanceFoundationGrants")) _
		+ Int(Request("frmFinancePerformanceCorporateGifts")) _
		+ Int(Request("frmFinancePerformanceBBBSAGrants"))_
		+ Int(Request("frmFinancePerformanceIndividualGiving"))_
		+ Int(Request("frmFinancePerformanceDividendsInterest"))_
		+ Int(Request("frmFinancePerformanceOther"))
		
		RST("UnitedWay") = Request("frmFinancePerformanceUnitedWay")		
		RST("GovFederalFunding") = Request("frmFinancePerformanceGovFederalFunding")		
		RST("GovStateFunding") = Request("frmFinancePerformanceGovStateFunding")					
		RST("GovLocalFunding") = Request("frmFinancePerformanceGovLocalFunding")					
		RST("FoundationGrants") = Request("frmFinancePerformanceFoundationGrants")	
		RST("CorporateGifts") = Request("frmFinancePerformanceCorporateGifts")			
		RST("BBBSAGrants") = Request("frmFinancePerformanceBBBSAGrants")		
		RST("IndividualGiving") = Request("frmFinancePerformanceIndividualGiving")			
		RST("DividendsInterest") = Request("frmFinancePerformanceDividendsInterest")					
		RST("Other") = Request("frmFinancePerformanceOther")				
		RST("Total") = Request("frmFinancePerformanceOther")						
		RST("TotalAmountRMM") = Request("frmFinancePerformanceTotalAmountRMM")			
		RST("TotalAmountBFKS") = Request("frmFinancePerformanceTotalAmountBFKS")			
		RST("TotalOperatingExpense") = Request("frmFinancePerformanceTotalOperatingExpense")		
										
		
		RST("CreateDate") = Now
		RST.Update
		RST.Close
		Set RST = Nothing
		form = "FinancePerformance"
		modtype = "new"	
		
		
		m = Request("month")
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
	RST.Open "SELECT * FROM tbl_frmFinancePerformance WHERE agencyID='" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")) & " AND Month=" & Int(Request("month")), Con, 1, 3
	
	tt = Int(Request("frmFinancePerformanceUnitedWay")) _ 
	+ Int(Request("frmFinancePerformanceGovFederalFunding"))_
	+ Int(Request("frmFinancePerformanceGovStateFunding"))_			
	+ Int(Request("frmFinancePerformanceGovLocalFunding"))_						
	+ Int(Request("frmFinancePerformanceFoundationGrants")) _
	+ Int(Request("frmFinancePerformanceCorporateGifts")) _
	+ Int(Request("frmFinancePerformanceBBBSAGrants"))_
	+ Int(Request("frmFinancePerformanceIndividualGiving"))_
	+ Int(Request("frmFinancePerformanceDividendsInterest"))_
	+ Int(Request("frmFinancePerformanceOther"))

	RST("UnitedWay") = Request("frmFinancePerformanceUnitedWay")		
	RST("GovFederalFunding") = Request("frmFinancePerformanceGovFederalFunding")		
	RST("GovStateFunding") = Request("frmFinancePerformanceGovStateFunding")					
	RST("GovLocalFunding") = Request("frmFinancePerformanceGovLocalFunding")					
	RST("FoundationGrants") = Request("frmFinancePerformanceFoundationGrants")	
	RST("CorporateGifts") = Request("frmFinancePerformanceCorporateGifts")			
	RST("BBBSAGrants") = Request("frmFinancePerformanceBBBSAGrants")		
	RST("IndividualGiving") = Request("frmFinancePerformanceIndividualGiving")			
	RST("DividendsInterest") = Request("frmFinancePerformanceDividendsInterest")					
	RST("Other") = Request("frmFinancePerformanceOther")				
	RST("Total") = Request("frmFinancePerformanceOther")						
	RST("TotalAmountRMM") = Request("frmFinancePerformanceTotalAmountRMM")			
	RST("TotalAmountBFKS") = Request("frmFinancePerformanceTotalAmountBFKS")			
	RST("TotalOperatingExpense") = Request("frmFinancePerformanceTotalOperatingExpense")		
	
	jMod = RST("FinancePerformanceID") %>
	
	
	
	
	<%
	RST.Update
	RST.Close
	Set RST = Nothing
	form = "FinancePerformance"
	modtype = "edit"
	m = Request("month")
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


<% dim HelpId
HelpId = 0
%>



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Monthly Revenue / Expense</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
<script language="javascript">
<!--	


function addEmUp() {
	var box1 = Number(document.frmFinancePerformance.frmFinancePerformanceUnitedWay.value)
	var box2 = Number(document.frmFinancePerformance.frmFinancePerformanceGovFederalFunding.value)
	var box3 = Number(document.frmFinancePerformance.frmFinancePerformanceGovStateFunding.value)	
	var box4 = Number(document.frmFinancePerformance.frmFinancePerformanceGovLocalFunding.value)		
	var box5 = Number(document.frmFinancePerformance.frmFinancePerformanceFoundationGrants.value)
	var box6 = Number(document.frmFinancePerformance.frmFinancePerformanceCorporateGifts.value)
	var box7 = Number(document.frmFinancePerformance.frmFinancePerformanceBBBSAGrants.value)	
	var box8 = Number(document.frmFinancePerformance.frmFinancePerformanceIndividualGiving.value)	
	var box9 = Number(document.frmFinancePerformance.frmFinancePerformanceDividendsInterest.value)
	var box10 = Number(document.frmFinancePerformance.frmFinancePerformanceOther.value)
	var boxtotal = box1 + box2 + box3 + box4 + box5 + box6 + box7 + box8 + box9 + box10
	document.frmFinancePerformance.frmFinancePerformanceTotal.value = boxtotal
}


function noChange()
	{
	alert("This will add automatically. Do not edit this field.");
	addEmUp();
	}

function addUpBreakouts() {
	var box1 = Number(document.frmFinancePerformance.frmFinancePerformanceTotalAmountBFKS.value)
	var box2 = Number(document.frmFinancePerformance.frmFinancePerformanceTotalAmountRMM.value)
	var boxtotal = box1 + box2
	if (boxtotal > Number(document.frmFinancePerformance.frmFinancePerformanceTotal.value))
	{
		alert("Total Amounts of BFKS and RMM Cannot Exceed Total Revenue.");
	}
}


//Field Validations

function checkForIntegerCommas(valueToCheck)
{
	var myRegularExpression = /^[0-9]+(,[0-9]{3})*$/;  // Checks for integer with or without commas
	if(!(myRegularExpression.test(valueToCheck)))
	{
		alert("Please make sure you have entered a whole, non-negative number with no spaces, decimal points or commas.\nDo not leave this field blank."); 
	} 
}

function validateForm()
{	
	
	var onlyInteger = /^[0-9]+(,[0-9]{3})*$/;
	
	if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceUnitedWay.value)))
	{
		alert("Error - United Way Field.\nPlease make sure you have entered a whole, non-negative number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceUnitedWay.focus();				
	} 

	else	
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceGovFederalFunding.value)))
	{
		alert("Error - Federal Funding Field.\nPlease make sure you have entered a whole, non-negative number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceGovFederalFunding.focus();		
	}

	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceGovStateFunding.value)))
	{
		alert("Error - State Funding Field.\nPlease make sure you have entered a whole, non-negative number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceGovStateFunding.focus();
	}	

	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceGovLocalFunding.value)))
	{
		alert("Error - Local Funding Field.\nPlease make sure you have entered a whole, non-negative number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceGovLocalFunding.focus();
	}	

	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceFoundationGrants.value)))
	{
		alert("Error - Foundation Grants Field.\nPlease make sure you have entered a whole, non-negative number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceFoundationGrants.focus();
	}	

	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceCorporateGifts.value)))
	{
		alert("Error - Corporate Gifts Field.\nPlease make sure you have entered a whole, non-negative number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceCorporateGifts.focus();
	}

	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceBBBSAGrants.value)))
	{
		alert("Error - BBBSA (Pass-Through) Grants Field.\nPlease make sure you have entered a whole, non-negative number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceBBBSAGrants.focus();
	}	
	
	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceIndividualGiving.value)))
	{
		alert("Error - Individual Giving Field.\nPlease make sure you have entered a whole, non-negative number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceIndividualGiving.focus();
	}		
	
	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceDividendsInterest.value)))
	{
		alert("Error - Dividends and Interest Field.\nPlease make sure you have entered a whole, non-negative number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceDividendsInterest.focus();
	}	
	
	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceOther.value)))
	{
		alert("Error - Other Field.\nPlease make sure you have entered a whole, non-negative number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceOther.focus();
	}		
	
	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceTotalAmountBFKS.value)))
	{
		alert("Error - BFKS Field.\nPlease make sure you have entered a whole, non-negative number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceTotalAmountBFKS.focus();
	}	

	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceTotalAmountRMM.value)))
	{
		alert("Error - RMM Field.\nPlease make sure you have entered a whole, non-negative number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceTotalAmountRMM.focus();
	}			
	
	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceTotalOperatingExpense.value)))
	{
		alert("Error - Total Operating Expense Field.\nPlease make sure you have entered a whole, non-negative number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceTotalOperatingExpense.focus();
	}			


	else	
		document.frmFinancePerformance.submit();	
}		


function getNextElement (field) 
{
	var form = field.form;
  	for (var e = 0; e < form.elements.length; e++)
    	if (field == form.elements[e])
      	break;
  	return form.elements[++e % form.elements.length];
}


//-->	
</script>

<!-- Popup Window Script -->
<SCRIPT LANGUAGE = "JavaScript">

<!-- Begin
function NewWindow(mypage, myname, w, h) {
var winl = (screen.width - w) / 2;
var wint = (screen.height - h) / 2;
winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',resizable, scrollbars'
win = window.open(mypage, myname, winprops)
if (parseInt(navigator.appVersion) >= 4) { win.window.focus(); }
}

//  End -->

</SCRIPT>

	
<% ' <!--#include file="../includes/top_nav_forms_monthly.inc"--><!-- include file has </head> and <body> tags --><br>      %>
<!--#include file="../includes/surveytitle.inc"-->


<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_football.jpg" alt="" width="220" height="477" border="0"></td> 
<td valign="top" align="left">

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


<form name="frmFinancePerformance" action="FinancePerformance_edit.asp" method="post"> <!-- onsubmit="return submitFormValidate(this)"> -->
<!--#include file="../includes/form_stamp.asp"-->

<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmFinancePerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & " AND Month=" & Int(Request("m"))
	Set GetPerformance = Con.Execute(query)
	
		
	
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
<p class="formMain"><br>We're sorry, but this form was previously completed. To make changes please <a href="monthly.asp">reselect</a> the 
appropriate form and year and update the existing information.</p>
<%
Response.End
End If 
%> 




<br>
		<table width="550" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063" >
		<tr>
			<td colspan="2" class="formHeader">Monthly Revenue / Expense<BR><%= MonthName(Request("m"), False) & " " & Request("y") %></td>
		</tr>
		
		<tr>
			<td colspan="2" class="formMain"><font color="#ff0000"><div align="center"><strong>Please Note: </strong>After entering your information, you <strong>must</strong> click on the "Save" button at the bottom of the form and wait for the "Thank You" screen or your changes will be lost.</div></font></td>
		</tr>


			<tr>
				<td valign="middle" align="center" class="formMainBold" colspan="2">REVENUE</td>
			</tr>
<!-- Populate Breakdown Fields Wth Zeros -->
					<input type="hidden"  class="formMain"  size="5" maxlength="10" value="0" name="frmFinancePerformanceUnitedWay"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">				
					<input type="hidden"  class="formMain"  size="5" maxlength="10" value="0" name="frmFinancePerformanceGovFederalFunding"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">								
					<input type="hidden"  class="formMain"  size="5" maxlength="10" value="0" name="frmFinancePerformanceGovStateFunding"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">								
					<input type="hidden"  class="formMain"  size="5" maxlength="10" value="0" name="frmFinancePerformanceGovLocalFunding"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">								
					<input type="hidden"  class="formMain"  size="5" maxlength="10" value="0" name="frmFinancePerformanceFoundationGrants"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">								
					<input type="hidden"  class="formMain"  size="5" maxlength="10" value="0" name="frmFinancePerformanceCorporateGifts"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">								
					<input type="hidden"  class="formMain"  size="5" maxlength="10" value="0" name="frmFinancePerformanceBBBSAGrants"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">								
					<input type="hidden"  class="formMain"  size="5" maxlength="10" value="0" name="frmFinancePerformanceIndividualGiving"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">								
					<input type="hidden"  class="formMain"  size="5" maxlength="10" value="0" name="frmFinancePerformanceDividendsInterest"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">
<!-- User-entered -->
			<tr>
				<td valign="middle" class="formMain">Total Revenue&nbsp;booked&nbsp;for&nbsp;the&nbsp;Month&nbsp;of&nbsp;<%= MonthName(Request("m"), False) & " " & Request("y") %>&nbsp;&nbsp;<br><i>NOTE: Detailed revenue breakdown questions will be asked later in 2005.</i></td>
				<td valign="middle" class="formMain">$				
					<input type="text"  class="formMain"  size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Other") %><% Else %>0<% End If %>" name="frmFinancePerformanceOther"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">
				</td>
			</tr>
					<input type="hidden"  class="formMain"  size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Total") %><% Else %>0<% End If %>" name="frmFinancePerformanceTotal" onchange="noChange();" readonly>			
					<input type="hidden" class="formMain"  size="5" maxlength="10" value="0" name="frmFinancePerformanceTotalAmountBFKS"  onchange="checkForIntegerCommas(this.value);" onblur="addUpBreakouts();">
					<input type="hidden" class="formMain"  size="5" maxlength="10" value="0" name="frmFinancePerformanceTotalAmountRMM"  onchange="checkForIntegerCommas(this.value);" onblur="addUpBreakouts();">
			
		
			<tr>
				<td valign="middle" align="center" class="formMainBold" colspan="2">EXPENSE</td>
			</tr>
			
			<tr>
				<td valign="middle" class="formMain">Total Operating Expense<br>(should not include expense directly related to fundraising events)</td>
				<td valign="middle" class="formMain">$				
					<input type="text"  class="formMain"  size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("TotalOperatingExpense") %><% Else %>0<% End If %>" name="frmFinancePerformanceTotalOperatingExpense"  onchange="checkForIntegerCommas(this.value);" >
				</td>				
			</tr>	
			

		<tr>
			<td colspan="2" class="formHeader">
				<input type="button" value="Save Form" class="formMainBold" onclick="validateForm(); return false;">
			</td>
		</tr>

		<tr>
			<td colspan="2"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
		</tr>
		</table>

</td>
</tr>
</table>

									
<% 
If say = "edit" Then
	GetPerformance.Close
	Set GetPerformance = Nothing

	Con.Close
	Set Con = Nothing
	
End If


 %>


</form>
<% End If %>
</body>
</html>


