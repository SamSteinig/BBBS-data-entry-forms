<!--#include file="../includes/NAD_BE.asp" -->

<% 
Dim RequestStatus 
RequestStatus = Request("status")

If Requeststatus = "" Then 'Check for data, the quarterly page allows skipping quarters
	Set DupCon = Server.CreateObject("ADODB.Connection")
	DupCon.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT Count(*) As NumberOfEntries FROM tbl_frmQuarterlyCashflow WHERE AgencyID = '" & Request("id") & "' and Year = " & Request("y")	& " and Quarter = " & Request("q")
	Set DuplicateRecord = DupCon.Execute(query)
	numberOfExisting = DuplicateRecord("NumberOfEntries")
	DuplicateRecord.Close
	Set DuplicateRecord = Nothing
	DupCon.Close
	Set DupCon = Nothing
	
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	
	If(numberOfExisting <> 0) Then
		Requeststatus = "editOld"
	End If	    
End If

If Requeststatus = "addNew" Then
' Check for duplicate records
	
	Set DupCon = Server.CreateObject("ADODB.Connection")
	DupCon.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT Count(*) As NumberOfEntries FROM tbl_frmQuarterlyCashflow WHERE AgencyID = '" & Request("id") & "' and Year = " & Request("y")	& " and Quarter = " & Request("q")
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
		RST.Open "SELECT * FROM tbl_frmQuarterlyCashflow", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = Request("id")
		RST("Year") = Request("y")
		RST("Quarter") = Request("q")

		RST("CashAndInvestments") = Request("CashAndInvestments")										
		RST("Receivables") = Request("Receivables")										
		RST("AllOtherAssets") = Request("AllOtherAssets")										
		RST("TotalAssets") = Request("TotalAssets")										
		RST("CurrentLiabilities") = Request("CurrentLiabilities")										
		RST("LongTermLiabilities") = Request("LongTermLiabilities")										
		RST("TotalLiabilities") = Request("TotalLiabilities")										
		RST("NetAssets") = Request("NetAssets")										
		RST("LiabilitiesAndNetAssets") = Request("LiabilitiesAndNetAssets")										
		
		RST("CreateDate") = Now
		RST.Update
		RST.Close
		Set RST = Nothing
		form = "QuarterlyCashflow"
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

ElseIf Requeststatus = "editSave" Then

	Set Con = Server.CreateObject("ADODB.Connection")
	Set RST = Server.CreateObject("ADODB.Recordset")
	Con.Open "BBBSAforms", "sa","12sist12"
	RST.Open "SELECT * FROM tbl_frmQuarterlyCashflow WHERE agencyID='" & Request("id") & "' AND Year=" & Int(Request("y")) & " AND Quarter=" & Int(Request("q")), Con, 1, 3
	
	RST("CashAndInvestments") = Request("CashAndInvestments")										
	RST("Receivables") = Request("Receivables")										
	RST("AllOtherAssets") = Request("AllOtherAssets")										
	RST("TotalAssets") = Request("TotalAssets")										
	RST("CurrentLiabilities") = Request("CurrentLiabilities")										
	RST("LongTermLiabilities") = Request("LongTermLiabilities")										
	RST("TotalLiabilities") = Request("TotalLiabilities")										
	RST("NetAssets") = Request("NetAssets")										
	RST("LiabilitiesAndNetAssets") = Request("LiabilitiesAndNetAssets")										
	
	jMod = RST("QuarterlyCashflowID") %>
	
	<%
	RST.Update
	RST.Close
	Set RST = Nothing
	form = "QuarterlyCashflow"
	modtype = "edit"
	m = Request("q")
	%>
	<!--#include file="../includes/modify_stamp.asp"-->
	<%	
	Con.Close
	Set Con = Nothing
	say = "thanks"
ElseIf Requeststatus = "editOld" Then
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
	<title>Quarterly Balance Sheet</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
<script language="javascript">
<!--	

function addEmUp() {
	var box1 = Number(document.frmQuarterlyCashflow.CashAndInvestments.value)
	var box2 = Number(document.frmQuarterlyCashflow.Receivables.value)
	var box3 = Number(document.frmQuarterlyCashflow.AllOtherAssets.value)	
	var boxTA = Number(box1 + box2 + box3)
	document.frmQuarterlyCashflow.TotalAssets.value = boxTA

	var box4 = Number(document.frmQuarterlyCashflow.CurrentLiabilities.value)		
	var box5 = Number(document.frmQuarterlyCashflow.LongTermLiabilities.value)
	var boxTL = Number(box4 + box5)
	document.frmQuarterlyCashflow.TotalLiabilities.value = boxTL
	
	var boxNA = Number(boxTA - boxTL)
	document.frmQuarterlyCashflow.NetAssets.value = boxNA
	var boxLnNA = Number(boxTL + boxNA)
	document.frmQuarterlyCashflow.LiabilitiesAndNetAssets.value = boxLnNA
}


function TotalNet() {
}	


function noChange()
	{
	alert("This will add automatically. Do not edit this field.");
	addEmUp();
	}

function addUpBreakouts() {
}

function addEventBreakouts() {
}

//Field Validations


function checkForIntegerCommas(valueToCheck)
{
//	var myRegularExpression = /^[0-9]+([0-9]{3})*$/;  // Checks for integer with or without commas
//	var myRegularExpression = /\d+$/;  // Checks for integer with or without commas
	var myRegularExpression = /^[-]?[0-9]+([0-9]{3})*$/;  // Checks for integer with or without commas	
	if(!(myRegularExpression.test(valueToCheck)))
	{
		alert("Please make sure you have entered a whole number with no spaces, decimal points or commas.\nDo not leave this field blank."); 
	} 
}

function validateForm()
{	
	
//	var onlyInteger = /^[0-9]+([0-9]{3})*$/;
	var onlyInteger = /\d+$/;

		
	var box1 = Number(document.frmQuarterlyCashflow.CashAndInvestments.value)
	var box2 = Number(document.frmQuarterlyCashflow.Receivables.value)
	var box3 = Number(document.frmQuarterlyCashflow.AllOtherAssets.value)	
	var box4 = Number(document.frmQuarterlyCashflow.CurrentLiabilities.value)		
	var box5 = Number(document.frmQuarterlyCashflow.LongTermLiabilities.value)
	
	
//	if (!(onlyInteger.test(document.frmQuarterlyCashflow.frmQuarterlyCashflowUnitedWay.value)) || document.frmQuarterlyCashflow.frmQuarterlyCashflowUnitedWay.value =="")

	if (!(onlyInteger.test(document.frmQuarterlyCashflow.CashAndInvestments.value)) || document.frmQuarterlyCashflow.CashAndInvestments.value =="")
	{
		alert("Error - Cash and Investments Field.\nPlease make sure you have entered a whole number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmQuarterlyCashflow.CashAndInvestments.focus();				
	} 

	else	
		if (!(onlyInteger.test(document.frmQuarterlyCashflow.Receivables.value)) || document.frmQuarterlyCashflow.Receivables.value =="")
	{
		alert("Error - Receivables Field.\nPlease make sure you have entered a whole number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmQuarterlyCashflow.Receivables.focus();		
	}

	else
		if (!(onlyInteger.test(document.frmQuarterlyCashflow.AllOtherAssets.value)) || document.frmQuarterlyCashflow.AllOtherAssets.value =="")
	{
		alert("Error - All Other Assets Field.\nPlease make sure you have entered a whole number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmQuarterlyCashflow.AllOtherAssets.focus();
	}	

	else
		if (!(onlyInteger.test(document.frmQuarterlyCashflow.CurrentLiabilities.value)) || document.frmQuarterlyCashflow.CurrentLiabilities.value =="")
	{
		alert("Error - Current Liabilities Field.\nPlease make sure you have entered a whole number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmQuarterlyCashflow.CurrentLiabilities.focus();
	}	

	else
		if (!(onlyInteger.test(document.frmQuarterlyCashflow.LongTermLiabilities.value)) || document.frmQuarterlyCashflow.LongTermLiabilities.value =="")
	{
		alert("Error - Long-Term Liabilities Field.\nPlease make sure you have entered a whole number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmQuarterlyCashflow.LongTermLiabilities.focus();
	}	

	else	
		document.frmQuarterlyCashflow.submit();	
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
    <tr><td width="220" valign="top"><img src="../includes/images/photos_football.jpg" alt="" width="220" height="477" border="0"></td> 
		<td valign="top" align="left" class="formMain">


<!-- Check Form Status -->

<% 	
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
query = "SELECT * FROM tbl_FormStatus WHERE FormName='Finance'"
Set GetFormStatus = Con.Execute(query)
%>	

<% if (GetFormStatus("Status").Value) = "Down" then %>
	<p><br><br><br>
	<i><font color="red"><b>
	<%= (GetFormStatus("Message").Value) %>
	</p></i></font></b>

<% else %>	





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


<% ElseIf say <> "thanks" Then  
%>


<form name="frmQuarterlyCashflow" action="QuarterlyCashflow_edit.asp" method="post"> <!-- onsubmit="return submitFormValidate(this)"> -->
<!--#include file="../includes/form_stamp.asp"-->

<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmQuarterlyCashflow WHERE AgencyID=" & Request("id") & " AND Year=" & Int(Request("y")) & " AND Quarter=" & Int(Request("q"))
	Set GetPerformance = Con.Execute(query)
 %>
<input type="hidden" name="status" value="editSave">
<% Else %>
<input type="hidden" name="status" value="addNew">
<%
End If
 %>
<input type="hidden" name="y" value=<%=Request("y")%>>
<input type="hidden" name="q" value=<%=Request("q")%>>
<input type="hidden" name="id" value=<%=Request("id")%>>
 
<%
If say = "previouslyEdited" Then
%>
<p class="formMain"><br>We're sorry, but this form was previously completed. To make changes please <a href="monthly.asp">reselect</a> the 
appropriate form and year and update the existing information.</p>
<%
Response.End
End If 

Dim strAsOf 

Select Case Request("q")
     Case "1"
         strAsOf = "As Of March 31,"
     Case "2"
         strAsOf = "As Of June 30,"
     Case "3"
         strAsOf = "As Of September 30,"
     Case "4"
         strAsOf = "As Of December 31,"
End Select

%> 

<br>
	<table width="550" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063" >
		<tr><td colspan="2" class="formHeader">BALANCE SHEET<BR><%= strAsOf & " " & Request("y") %></td></tr>
		<tr><td colspan="2" class="formMain"><font color="#ff0000"><div align="center"><strong>Please Note: </strong>After entering your information, you <strong>must</strong> click on the "Save" button at the bottom of the form and wait for the "Thank You" screen or your changes will be lost.</div></font></td></tr>
		<tr><td valign="middle" align="center" class="formHeaderMedium" colspan="2">&nbsp;</td></tr>
		<tr><td valign="middle" align="center" class="formMain" colspan="2"><p><i>Please report Cashflow information you booked for Q<%= Request("q") & " " & Request("y") %></u>. </p><p>Click on the <img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"> next to each line item for a detailed explanation.</p></i></td></tr>			

		<tr><td valign="middle" class="formMain"><a href="../helpfiles/surveyhelp.asp?HelpID=qfp_Cash_and_Investments" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Cash and Investments</td>
			<td valign="middle" class="formMain">$<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CashAndInvestments") %><% Else %>0<% End If %>" name="CashAndInvestments"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();"></td></tr>	
		<tr><td valign="middle" class="formMain"><a href="../helpfiles/surveyhelp.asp?HelpID=qfp_Receivables" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Receivables</td>
			<td valign="middle" class="formMain">$<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Receivables") %><% Else %>0<% End If %>" name="Receivables"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();"></td></tr>	
		<tr><td valign="middle" class="formMain"><a href="../helpfiles/surveyhelp.asp?HelpID=qfp_AllOtherAssets" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;All Other Assets</td>
			<td valign="middle" class="formMain">$<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("AllOtherAssets") %><% Else %>0<% End If %>" name="AllOtherAssets"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();"></td></tr>				
		<tr><td valign="middle" class="formMain"><a href="../helpfiles/surveyhelp.asp?HelpID=qfp_TotalAssets" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>Total Assets</td>
			<td valign="middle" class="formMain" bgcolor="#c0c0c0">$<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("TotalAssets") %><% Else %>0<% End If %>" name="TotalAssets" onchange="noChange();" readonly><span class="formSubHead">&nbsp;&nbsp;&nbsp;Calculated</span></td></tr>					

		<tr><td colspan=2>&nbsp;</td>
		
		<tr><td valign="middle" class="formMain"><a href="../helpfiles/surveyhelp.asp?HelpID=qfp_CurrentLiabilities" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Current Liabilities</td>
			<td valign="middle" class="formMain">$<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CurrentLiabilities") %><% Else %>0<% End If %>" name="CurrentLiabilities"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();"></td></tr>	
		<tr><td valign="middle" class="formMain"><a href="../helpfiles/surveyhelp.asp?HelpID=qfp_LongTermLiabilities" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Long-Term Liabilities</td>
			<td valign="middle" class="formMain">$<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("LongTermLiabilities") %><% Else %>0<% End If %>" name="LongTermLiabilities"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();"></td></tr>	
		<tr><td valign="middle" class="formMain"><a href="../helpfiles/surveyhelp.asp?HelpID=qfp_TotalLiabilities" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>Total Liabilities</td>
			<td valign="middle" class="formMain" bgcolor="#c0c0c0">$<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("TotalLiabilities") %><% Else %>0<% End If %>" name="TotalLiabilities" onchange="noChange();" readonly><span class="formSubHead">&nbsp;&nbsp;&nbsp;Calculated</span></td></tr>					

		<tr><td colspan=2>&nbsp;</td>
		
		<tr><td valign="middle" class="formMain"><a href="../helpfiles/surveyhelp.asp?HelpID=qfp_NetAssets" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>Net Assets</td>
			<td valign="middle" class="formMain" bgcolor="#c0c0c0">$<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("NetAssets") %><% Else %>0<% End If %>" name="NetAssets" onchange="noChange();" readonly><span class="formSubHead">&nbsp;&nbsp;&nbsp;Calculated</span></td></tr>					
		<tr><td valign="middle" class="formMain"><a href="../helpfiles/surveyhelp.asp?HelpID=qfp_LiabilitiesAndNetAssets" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>Liabilities and Net Assets</td>
			<td valign="middle" class="formMain" bgcolor="#c0c0c0">$<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("LiabilitiesAndNetAssets") %><% Else %>0<% End If %>" name="LiabilitiesAndNetAssets" onchange="noChange();" readonly><span class="formSubHead">&nbsp;&nbsp;&nbsp;Calculated</span></td></tr>					

		<tr><td colspan=2>&nbsp;</td>
		
		<tr><td colspan="2" class="formHeader"><input type="button" value="Save Form" class="formMainBold" onclick="validateForm(); return false;"  onclick="TotalNet();" ></td></tr>

		<tr><td colspan="2"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td></tr>
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

<% End If %>
</body>
</html>


