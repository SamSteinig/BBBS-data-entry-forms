<%@LANGUAGE="VBSCRIPT"%>

<!--#include file="../../../Connections/NAD_BE.asp" -->
<% if Request("Agency_ID") = "" then %>
<!--#INCLUDE FILE="../security/passwordpro/check_user_inc.asp"-->
<% else 
	Dim ID
	if Request("Agency_ID") = 9999 then ID = 0 else ID = Request("Agency_ID") end if
	Session("AGENCY_ID") = ID
	Session("ACCESS_LEVEL") = 4
	Session("PASSWORDACCESS") = "Yes"
end if %>

<%
dim templateURL
templateURL="http://agencyconnection.bbbs.org/site/lookup.asp?c=9dJGKRNqFmG&b=2574379"
'templateURL="http://agencyconnection.bbbs.org/site/lookup.asp?c=9dJGKRNqFmG&b=1812579"
%>
<!--#INCLUDE FILE="../../../media/inc/kinterawrapper.asp"-->
<% Response.Write(LeftContent(myStr))%>

<% 

if Request("SQLHID") <> "" then

	Dim sqlStr, Submited
	Submited = "yes"
	sqlStr = Request("SQLHID")	
	
	set Edit = Server.CreateObject("ADODB.Connection")
	Edit.ConnectionString = ConnStr
	Edit.Open
	Edit.Execute sqlStr
	Edit.Close
	
	AID = request("AgencyID")
	CompYear = Request("Year")
	AgencyName = Request("AgencyName")
	AgencyCity = Request("AgencyCity")
	AgencyState = Request("AgencyState")
	AgencyAddress = Request("AgencyAddress")
	AgencyZip = Request("AgencyZip")
	DateSubmitted = Request("DateSubmitted")
	FiscalYearEnded = Request("FiscalYearEnded")
	PrepBy = Request("PrepBy")
	PrepByPhone = Request("PrepByPhone")
	TotalExpenditures = Request("TotalExpenditures")
	PriorYearFeesPaidToBBBSA = Request("PriorYearFeesPaidToBBBSA")
	PriorYearCapitalPurchases = Request("PriorYearCapitalPurchases")
	PriorYearDepreciation = Request("PriorYearDepreciation")
	PriorYearFundraisingExpenses = Request("PriorYearFundraisingExpenses")
	AnnualDiscount = Request("AnnualDiscount")
	
	TotalDeductions = Request("TotalDeductions")
	AdjustedExpenditures = Request("AdjustedExpenditures")
	FeeCalc1 = Request("FeeCalc1")
	FeeCalc2 = Request("FeeCalc2")
	FeeCalc3 = Request("FeeCalc3")
	FeeCalc4 = Request("FeeCalc4")
	TotalFeeCalc = Request("TotalFeeCalc")
	TotalAffiliationFeesDue = Request("TotalAffiliationFeesDue")
	
	Dim Message
	Message = " " & Year(Now()) & " AFFFILIATION FEE CALCULATION FORM from the Agency Connection Web site " & vbCrlf & vbCrLf
	Message = Message & "Agency Name:       " & AgencyName & " with ID " & AID & vbCrLf
	Message = Message & "Submition Date:    " & DateSubmitted & vbCrLf
	Message = Message & "City, State, ZIP:  " & AgencyCity & ", " & AgencyState & ", " & AgencyZip & vbCrLf & vbCrLf
	Message = Message & "Fiscal Year Ended: " & FiscalYearEnded & vbCrLf
	Message = Message & "Form Prepared By:  " & PrepBy & vbCrLf
	Message = Message & "Preparer's Phone:  " & PrepByPhone & vbCrLf
	Message = Message & vbCrLf & Chr(13) & Chr(10)
	Message = Message & " (A) Total Expenditures:                                           " & TotalExpenditures & " (A)" & vbCrLf
	Message = Message & " (B) Less: Prior Year Fees Paid to BBBSA:                   " & PriorYearFeesPaidToBBBSA & " (B)" & vbCrLf
	Message = Message & " (C) Less: Prior Year Capital Purchases:                    " & PriorYearCapitalPurchases & " (C)" & vbCrLf
	Message = Message & " (D) Less: Prior Year Depreciation:                         " & PriorYearDepreciation & " (D)" & vbCrLf
	Message = Message & " (E) Less: Prior Year Fundraising Expenses:                 " & PriorYearFundraisingExpenses & " (E)" & vbCrLf
	Message = Message & vbCrLf & Chr(13) & Chr(10)
	Message = Message & " (F) TOTAL DEDUCTIONS (B + C + D + E):                             " & TotalDeductions & " (F)" & vbCrLf
	Message = Message & " (G) ADJUSTED EXPENDITURES (A - F):                                " & AdjustedExpenditures & " (G)" & vbCrLf & vbCrLf
	Message = Message & " (H) 3.80% of first $100,000 of Adjusted Expenditures:      " & FeeCalc1 & " (H)" & vbCrLf
	Message = Message & " (I) 2.25% of the next $100,000 of Adjusted Expenditures:   " & FeeCalc2 & " (I)" & vbCrLf
	Message = Message & " (J) 1.00% of the next $300,000 of Adjusted Expenditures:   " & FeeCalc3 & " (J)" & vbCrLf
	Message = Message & " (K) 0.50% of the remaining Adjusted Expenditures:          " & FeeCalc4 & " (K)" & vbCrLf
	Message = Message & vbCrLf & Chr(13) & Chr(10)
	Message = Message & " (L) CALCULATED FEE (H + I + J + K):                               " & TotalFeeCalc & " (L)" & vbCrLf
	Message = Message & " (M) Less: Annual Discount %:                                      " & AnnualDiscount & " (M)" & vbCrLf & vbCrLf
	Message = Message & " (N) TOTAL AFFILIATION FEES DUE (L - M):                           " & TotalAffiliationFeesDue & " (N)" & vbCrLf
	
	Dim oMessage
	Set oMessage = Server.CreateObject("CDONTS.NewMail")
	Const CdoBodyFormatHTML = 0 ' Send HTML Mail
	Const CdoBodyFormatText = 1 ' Send Text Mail
	
	Const CdoMailFormatMime = 0 ' Send mail in MIME format.
	Const CdoMailFormatText = 1 ' Send mail in uninterrupted plain text (default value).
	
	With oMessage
		.To = "afffees@bbbs.org"
		.Bcc = "sergey_pinchuk@hotmail.com"
		.From = "AffilFees@AgencyConnection.org"
		.Subject = AgencyName & " with ID " & AID & " submited Fee Calculation Form"
		.BodyFormat = CdoBodyFormatText ' CdoBodyFromatHTML
		.MailFormat = CdoMailFormatMime ' CdoMailFormatText
		.Body = Message
		.Send
	End with
	Set oMessage = Nothing

	viewMode = "view"
	'AID = Request.QueryString("AgencyID")
	
	'Response.Redirect "feeform_confirm.asp?updated=1&CompYear="&CompYear&"&AgencyID="&AID&"PrepBy="&PrepBy&"&FiscalYear="&FiscalYear&
	
	'Response.Expires = 0
	'Response.Buffer = True
	'Response.AddHeader "Refresh"	



else
	Response.buffer = True 
	Response.ExpiresAbsolute = Now() - 1 
	Response.Expires = 0 
	Response.CacheControl = "no-cache"
	
	set Conn = Server.CreateObject("ADODB.Connection")
	Conn.ConnectionString = ConnStr
	Conn.Open
	
	Dim CheckFee, SQL, SelfAssessSearchYear, AID, CompYear, FeeSubmitted, AgencyName, AgencyCity, AgencyState, AgencyAddress, AgencyZip, PrepBy, PrepByPhone, TotalExpenditures, PaidtoBBBSA, CapitalPurchases, Depriciation, FundRaising, AnnualDiscount, viewMode
	
'	AID = 2'Request("AgencyID")
	AID = Session("AGENCY_ID")
	if Request("year") <> "" then CompYear = Request("year") else CompYear = 2009 end if
	'CompYear = Year(Now()) - 1
	
	query = "SELECT Fee_Calculation_Form_Submited FROM tbl_frmMinCompliance " &_
			   "WHERE Compliance_Year=" & Int(CompYear) & " and FK_Agency_ID =" & AID
	
	Set FeeForm = Conn.Execute(query)
	CheckFee = FeeForm("Fee_Calculation_Form_Submited")
	FeeForm.Close
	Set FeeForm = Nothing
	
		SQL = "SELECT m.FK_Agency_ID, d.AgencyName, d.AgencyCity, d.AgencyState, d.AgencyMailAddress, d.AgencyMailZip, " &_
		"m.Fee_Form_Submited_By, m.Fee_SubmitedBy_Phone, m.Fee_Total_Expenditures, m.Fee_Paid_To_BBBSA, m.Fee_Capital_Purchases, m.Fee_Depriciation, m.Fee_Fund_Raising_Expenses, m.Fee_Discount_Percent " &_
		"FROM tbl_frmMinCompliance m inner join tblDemogs d on m.FK_Agency_ID = d.AgencyID " &_
		"WHERE m.Compliance_Year=" & Int(CompYear) & " and FK_Agency_ID =" & AID
		
		set AgencyFees = Server.CreateObject("ADODB.Recordset")
		AgencyFees.ActiveConnection = ConnStr
		AgencyFees.Source = SQL

		AgencyFees.CursorType = 0
		AgencyFees.CursorLocation = 2
		AgencyFees.Open()
		
		AgencyName = AgencyFees.Fields.Item("AgencyName").Value
		AgencyCity = AgencyFees.Fields.Item("AgencyCity").Value
		AgencyState = AgencyFees.Fields.Item("AgencyState").Value
		AgencyAddress = AgencyFees.Fields.Item("AgencyMailAddress").Value
		AgencyZip = AgencyFees.Fields.Item("AgencyMailZip").Value
	
	if (CheckFee) then
		viewMode = "view"
		
		PrepBy = AgencyFees.Fields.Item("Fee_Form_Submited_By").Value
		PrepByPhone = AgencyFees.Fields.Item("Fee_SubmitedBy_Phone").Value
		TotalExpenditures = AgencyFees.Fields.Item("Fee_Total_Expenditures").Value
		PriorYearFeesPaidToBBBSA = AgencyFees.Fields.Item("Fee_Paid_To_BBBSA").Value
		PriorYearCapitalPurchases = AgencyFees.Fields.Item("Fee_Capital_Purchases").Value
		PriorYearDepreciation = AgencyFees.Fields.Item("Fee_Depriciation").Value
		PriorYearFundraisingExpenses = AgencyFees.Fields.Item("Fee_Fund_Raising_Expenses").Value
		AnnualDiscount = AgencyFees.Fields.Item("Fee_Discount_Percent").Value
		
	else
		viewMode = "edit"
	end if

	AgencyFees.Close()
	set AgencyFees = nothing

	'response.Write "ID:" & AID
	'response.Write "CompYear: " & CompYear
	'response.Write "Fee: " & viewMode
	'response.Write "SQL: " & SQL
	'response.End
end if
%>

<%
'dim referrer
'referrer = "../MyAgency/POEMainEdit.asp"
'dim pageID
'pageID = "POESat"
'dim subID
'subID = ""

Dim ReadOnly, AccessLevel
'AgencyId = Session("AGENCY_ID")
'Session("AgencyIDN") = AgencyID
AccessLevel = Session("ACCESS_LEVEL")
ReadOnly = true
if (AccessLevel = "4") or (AccessLevel = "1") then
  ReadOnly = false
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
  <title>BBBS :: My Agency</title>
  <link rel="STYLESHEET" type="text/css" href="../../../surveys/OnlineForms/includes/bbbsa_forms.css">
  
  <SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
  
  if(window.event + "" == "undefined") event = null;
  function HM_f_PopUp(){return false};
  function HM_f_PopDown(){return false};
  popUp = HM_f_PopUp;
  popDown = HM_f_PopDown;
  
  </SCRIPT>
  
  <SCRIPT LANGUAGE="JavaScript1.2" SRC="../media/scripts/HM_Loader.js" TYPE='text/javascript'></SCRIPT>
  
  <SCRIPT LANGUAGE="JavaScript1.2" SRC="../media/scripts/tool_tip.js" TYPE='text/javascript'></SCRIPT>

<!-- Popup Window Script -->
<SCRIPT LANGUAGE = "JavaScript">

function PrintThisWindow() {
	window.print()
	}
	
   //Begin
function NewWindow(mypage, myname, w, h) {
var winl = (screen.width - w) / 2;
var wint = (screen.height - h) / 2;
winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',resizable, scrollbars'
win = window.open(mypage, myname, winprops)
if (parseInt(navigator.appVersion) >= 4) { win.window.focus(); }
}

  //  End -->


  // ======================
  // FORM VALIDATION SCRIPT
  // ======================
  function ValidateForm(f)
  {
	var sqlUpdate, PrepBy, PrepByPhone, FiscalYear, LA, LB, LC, LD, LE, LF, LG, AID, Year, TotalD;
	var currentDate = new Date();
	AID = FeeCalculator.AgencyID.value;
	Year = (currentDate.getFullYear())-1;
	PrepBy = FeeCalculator.PrepBy.value;
	PrepByPhone = FeeCalculator.PrepByPhone.value;
	FiscalYear = FeeCalculator.FiscalYearEnded.value;
	LA = (FeeCalculator.TotalExpenditures.value) * 1;
	LB = (FeeCalculator.PriorYearFeesPaidToBBBSA.value) * 1;
	LC = (FeeCalculator.PriorYearCapitalPurchases.value) * 1;
	LD = (FeeCalculator.PriorYearDepreciation.value) * 1;
	LE = (FeeCalculator.PriorYearFundraisingExpenses.value) * 1;
	LF = (FeeCalculator.TotalDeductions.value) * 1;
	LM = document.getElementById('AnnualDiscount').options[document.getElementById('AnnualDiscount').selectedIndex].value;
	LN = (FeeCalculator.TotalAffiliationFeesDue.value) * 1;
	
	TotalD = (LB*1 + LC*1 + LD*1 + LE*1);

	if (LF != TotalD)
	{
		alert("Total Deductions (Line F) is not equal to the sum of all deductions.");
		return false;
	}
	else if (FiscalYear == '')
	{
		alert("Please specify Fiscal Year Ended!");
		return false;
	}
	else if (PrepBy == "")
	{
		alert("Please enter the name of person who prepared this form");
		return false;
	}
	else if (PrepByPhone == "")
	{
		alert("Please provide contact phone number of person who prepared this form");
		return false;
	}
	else if (LF > LA)
	{
		alert("You cannot have Total Deductions greater than Total Expenditures");
		return false;
	}
	else if (LM == 'selected')
	{
		alert("You did not select any discount. If you don't wish to take advantage of the discount, select None in Line (M)");
		return false;
	}
	else
	{
		var dd = currentDate.getDate();
		var mm = currentDate.getMonth();
		mm = mm + 1;
		var yy = currentDate.getFullYear();
		currentDate = mm+"/"+dd+"/"+yy;
		//currentDate = new Date(currentDate);
		sqlUpdate = "UPDATE tbl_frmMinCompliance SET Fee_Calculation_Form_Submited = 1, Fee_Form_FiscalYear = " + FiscalYear + ", Fee_Form_Submited_By = '" + PrepBy + "', Fee_SubmitedBy_Phone = '" + PrepByPhone + "', Fee_Form_Submited_Date = '" + currentDate + "', Fee_Total_Expenditures = " + LA + ", Fee_Paid_To_BBBSA = " + LB + ", Fee_Capital_Purchases = " + LC + ", Fee_Depriciation = " + LD + ", Fee_Fund_Raising_Expenses = " + LE + ", Fee_Discount_Percent = " + LM + ", Fee_Total_Due = " + LN + " WHERE FK_Agency_ID = " + AID + " and Compliance_Year = " + Year;
		document.getElementById("SQLHID").value = sqlUpdate;
		document.getElementById("Year").value = Year;
		return true;
	}
  }
  // ======================
  // FEE CALCULATION SCRIPT
  // ======================
  function CalculateFees(valueToCheck)
  {
	var myRegularExpression = /\d+$/;  // Checks for integer with or without commas
	if(!(myRegularExpression.test(valueToCheck)))
	{
		alert("Please make sure you have entered a whole number with no spaces, decimal points or commas.\nDo not leave this field blank.");
	}
	function formatAsMoney(mnt)  //converts value into money format
	{
		mnt -= 0;
		mnt = (Math.round(mnt*100))/100;
		return (mnt == Math.floor(mnt)) ? mnt + '.00' : ( (mnt*10 == Math.floor(mnt*10)) ? mnt + '0' : mnt);
	}
	
	//var FeeCalcForm
	var LineA = FeeCalculator.TotalExpenditures.value
	var LineB = FeeCalculator.PriorYearFeesPaidToBBBSA.value
	var LineC = FeeCalculator.PriorYearCapitalPurchases.value
	var LineD = FeeCalculator.PriorYearDepreciation.value
	var LineE = FeeCalculator.PriorYearFundraisingExpenses.value
	var LineF = FeeCalculator.TotalDeductions.value
	var LineG = FeeCalculator.AdjustedExpenditures.value
	var LineH = FeeCalculator.FeeCalc1.value
	var LineI = FeeCalculator.FeeCalc2.value
	var LineJ = FeeCalculator.FeeCalc3.value
	var LineK = FeeCalculator.FeeCalc4.value
	var LineL = FeeCalculator.TotalFeeCalc.value
	var LineM = document.getElementById('AnnualDiscount').options[document.getElementById('AnnualDiscount').selectedIndex].value;
	//var LineM = Math.round(FeeCalculator.AnnualDiscount.value)
	var LineN = FeeCalculator.TotalAffiliationFeesDue.value
	
	// Lines A - E - EXPENDITURES & DEDUCTIONS
	FeeCalculator.TotalExpenditures.value = formatAsMoney(LineA)
	FeeCalculator.PriorYearFeesPaidToBBBSA.value = formatAsMoney(LineB)
	FeeCalculator.PriorYearCapitalPurchases.value = formatAsMoney(LineC)
	FeeCalculator.PriorYearDepreciation.value = formatAsMoney(LineD)
	FeeCalculator.PriorYearFundraisingExpenses.value = formatAsMoney(LineE)
	
	// Line F - TOTAL DEDUCTIONS
	FeeCalculator.TotalDeductions.value = formatAsMoney(LineB*1 + LineC*1 + LineD*1 + LineE*1)
	LineF = formatAsMoney(FeeCalculator.TotalDeductions.value)
	
	// Line G - ADJUSTED EXPENDITURES
	FeeCalculator.AdjustedExpenditures.value = formatAsMoney(LineA*1 - LineF*1)
	LineG = formatAsMoney(FeeCalculator.AdjustedExpenditures.value)
	
	// Lines H - L - FEE CALCULATIONS
	var FeeCalc1
	var FeeCalc2
	var FeeCalc3
	var FeeCalc4
	
	if (LineG < 0)
	{
		FeeCalculator.FeeCalc1.value = 0.00
		FeeCalculator.FeeCalc2.value = 0.00
		FeeCalculator.FeeCalc3.value = 0.00
		FeeCalculator.FeeCalc4.value = 0.00
		
		LineH = formatAsMoney(FeeCalculator.FeeCalc1.value)
		LineI = formatAsMoney(FeeCalculator.FeeCalc2.value)
		LineJ = formatAsMoney(FeeCalculator.FeeCalc3.value)
		LineK = formatAsMoney(FeeCalculator.FeeCalc4.value)
	}
	
	if (LineG >= 0 && LineG <= 100000)
	{
		FeeCalculator.FeeCalc1.value = formatAsMoney(LineG * .038)
		FeeCalculator.FeeCalc2.value = 0.00
		FeeCalculator.FeeCalc3.value = 0.00
		FeeCalculator.FeeCalc4.value = 0.00
		
		LineH = formatAsMoney(FeeCalculator.FeeCalc1.value)
		LineI = formatAsMoney(FeeCalculator.FeeCalc2.value)
		LineJ = formatAsMoney(FeeCalculator.FeeCalc3.value)
		LineK = formatAsMoney(FeeCalculator.FeeCalc4.value)
	}
	
	if (LineG > 100000 && LineG <= 200000)
	{
		FeeCalculator.FeeCalc1.value = formatAsMoney(100000 * .038)
		FeeCalculator.FeeCalc2.value = formatAsMoney((LineG - 100000) * .0225)
		FeeCalculator.FeeCalc3.value = 0.00
		FeeCalculator.FeeCalc4.value = 0.00
		
		LineH = formatAsMoney(FeeCalculator.FeeCalc1.value)
		LineI = formatAsMoney(FeeCalculator.FeeCalc2.value)
		LineJ = formatAsMoney(FeeCalculator.FeeCalc3.value)
		LineK = formatAsMoney(FeeCalculator.FeeCalc4.value)
	}
	
	if (LineG > 200000 && LineG <= 500000)
	{
		FeeCalculator.FeeCalc1.value = formatAsMoney(100000 * .0380)
		FeeCalculator.FeeCalc2.value = formatAsMoney(100000 * .0225)
		FeeCalculator.FeeCalc3.value = formatAsMoney((LineG - 200000) * .0100)
		FeeCalculator.FeeCalc4.value = 0.00
		
		LineH = formatAsMoney(FeeCalculator.FeeCalc1.value)
		LineI = formatAsMoney(FeeCalculator.FeeCalc2.value)
		LineJ = formatAsMoney(FeeCalculator.FeeCalc3.value)
		LineK = formatAsMoney(FeeCalculator.FeeCalc4.value)
	}
	
	if (LineG > 500000)
	{
		FeeCalculator.FeeCalc1.value = formatAsMoney(100000 * .0380)
		FeeCalculator.FeeCalc2.value = formatAsMoney(100000 * .0225)
		FeeCalculator.FeeCalc3.value = formatAsMoney(300000 * .0100)
		FeeCalculator.FeeCalc4.value = formatAsMoney((LineG - 500000) * .0050)
		
		LineH = formatAsMoney(FeeCalculator.FeeCalc1.value)
		LineI = formatAsMoney(FeeCalculator.FeeCalc2.value)
		LineJ = formatAsMoney(FeeCalculator.FeeCalc3.value)
		LineK = formatAsMoney(FeeCalculator.FeeCalc4.value)
	}
	
	FeeCalculator.TotalFeeCalc.value = formatAsMoney(LineH*1 + LineI*1 + LineJ*1 + LineK*1)
	LineL = FeeCalculator.TotalFeeCalc.value
	
	// Line M - ANNUAL DISCOUNT
	//var CurrentDate = FeeCalculator.DateSubmitted.value
	var AnnualDiscount
	AnnualDiscount = formatAsMoney(LineL * LineM)
	//FeeCalculator.AnnualDiscount.value = Math.round(AnnualDiscount)
	//LineM = Math.round(FeeCalculator.AnnualDiscount.value)
	
	// Line N - TOTAL AFFILIATION FEES DUE
	if (LineM == 'selected')
	{
		FeeCalculator.TotalAffiliationFeesDue.value = "Select (M)"
		LineN = FeeCalculator.TotalAffiliationFeesDue.value
	}
	else
	{
		FeeCalculator.TotalAffiliationFeesDue.value = formatAsMoney(LineL - AnnualDiscount)
		LineN = FeeCalculator.TotalAffiliationFeesDue.value
	}
	
  }
  //-->
</script>
  
  <LINK rel=STYLESHEET href = "../media/scripts/bbbsa.css" Type = "text/css">
  
  <!--#include file="../../../media/inc/mouseover.inc"-->
  

</head>

<BODY>



<!-- main table -->


<TABLE WIDTH="100%" BORDER="0" CELLPADDING="0" CELLSPACING="0" ID="Table2">
  <TR>


    <!-- main content cell -->
    <TD VALIGN="TOP" WIDTH="88%">
  
      
      <TABLE WIDTH="100%" BORDER="0" CELLPADDING="0" CELLSPACING="0" ID="Table4">

        <tr>
			<td VALIGN="TOP">
			<TABLE WIDTH="100%" BORDER="0" bgcolor="#DADADA" CELLPADDING="5" CELLSPACING="5" ID="Table3">
				<tr>
					<TD VALIGN="TOP" WIDTH="100%" align="center">
						<%if ((ReadOnly = true)) then%>
						<p class="formMain"><br>We're sorry, but you can only edit fee information using the <i>Full Access</i> or <i>Limited Access</i> password.</p>
						<%elseif ((viewMode = "view") AND (Submited <> "yes") AND (ReadOnly = false)) then%>
						<p class="formMain"><strong>We're sorry, but your data has already been submitted.</strong> <br>If you wish to make adjustments to your affiliation fees, please <a href="AffiliationCalculationForm.xls" class="cool3" target="_blank">download</a> the Excel version of the Fee Calculation form, fill it out manually, and either mail it to the National office or email it to <a href="mailto:Peter.Paiman@bbbs.org" class="cool3">Peter.Paiman@bbbs.org</a></p>
						<%elseif ((viewMode = "view") AND (Submited = "yes") AND (ReadOnly = false)) then%>
						<p class="formMain"><b>Thank you for submitting your Fee Calculation Data!</b></p>
						<%else%>
						<p class="formMain"><STRONG>Please Note: </STRONG>After entering your information, you <STRONG>must</STRONG> click on the "Submit Form" button at the bottom of the form 
						and wait for the "Thank You" screen or your changes will be lost. After you see the "Thank You" screen, you'll be able to PRINT the form.</p>
						<p class="formMain">If you wish to submit a paper Fee Calculation form, click <a href="AffiliationFeeCalculationForm.xls" class="cool3" target="_blank">HERE</a> to download the form.
						<%end if%>
						<%if ReadOnly = true then viewMode = "view"%>
					</td>
				</tr>
	</TABLE>
			</td>
        </tr>
        <TR>
          <TD VALIGN="TOP">
          <span class="bigtext">&nbsp;</span>
			<% if viewMode = "view" then%>
				<br>
				<!--<input type="button" Value="Print This Form" onClick="PrintThisWindow()">-->
				<A href="http://hostedagencies.bbbs.org/myagency/feeform_print.asp?AID=<%=AID%>&AgencyName=<%=AgencyName%>&AgencyAddress=<%=AgencyAddress%>&AgencyCity=<%=AgencyCity%>&AgencyState=<%=AgencyState%>&AgencyZip=<%=AgencyZip%>&PrepBy=<%=PrepBy%>&PrepByPhone=<%=PrepByPhone%>&TotalExpenditures=<%=TotalExpenditures%>&PriorYearFeesPaidToBBBSA=<%=PriorYearFeesPaidToBBBSA%>&PriorYearCapitalPurchases=<%=PriorYearCapitalPurchases%>&PriorYearDepreciation=<%=PriorYearDepreciation%>&PriorYearFundraisingExpenses=<%=PriorYearFundraisingExpenses%>&AnnualDiscount=<%=AnnualDiscount%>&TotalDeductions=<%=TotalDeductions%>&AdjustedExpenditures=<%=AdjustedExpenditures%>&FeeCalc1=<%=FeeCalc1%>&FeeCalc2=<%=FeeCalc2%>&FeeCalc3=<%=FeeCalc3%>&FeeCalc4=<%=FeeCalc4%>&TotalFeeCalc=<%=TotalFeeCalc%>&TotalAffiliationFeesDue=<%=TotalAffiliationFeesDue%>&FiscalYearEnded=<%=FiscalYearEnded%>" onclick="NewWindow(this.href,'name','700','500','yes');return false;">PRINT THIS FORM</a>
				<br>
			<% end if %>
<form action="feeform.asp?AgencyID=<%=AID%>" method="POST" name="FeeCalculator" id="FeeCalculator" onsubmit="return ValidateForm(this)">
	<table width="100%" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063" ID="Table5">
	


		<tr>
			<td colspan="7" class="formHeader"><%= Year(Date())+1%> FEE CALCULATION FORM</td>
		</tr>

		<tr>
			<td colspan="3" class="formMain">Your fee calculation is based on total "Adjusted Expenditures" for your agency's <u>most recent certified audit report</u>.  This form must be completed by all agencies and returned no later than Fabruary 29, <%= Year(Date())+1%>.</td>
		</tr>
		<tr>
		<td colspan="2" class="formMain" nowrap><strong>Agency Name</strong> <input type="text" name="AgencyName" size="40" readonly value="<%=AgencyName%>" ID="Text1"></td>
		<td colspan="1" class="formMain" nowrap><strong>Agency ID#</strong> <input type="text" name="AgencyID" size="4" readonly value="<%= AID %>" ID="Text2"></td>
		</tr>
		<tr>
		<td colspan="2" class="formMain" nowrap><strong>Agency Address</strong> <input type="text" name="AgencyAddress" size="40" readonly value="<%=AgencyAddress%>" ID="Text3"></td>
		<td colspan="1" class="formMain" nowrap><strong>Date</strong> <input type="text" name="DateSubmitted" value="<% =Date() %>" size="10" readonly ID="Text4"></td>
		</tr>
		<tr>
		<td colspan="2" class="formMain" nowrap><strong>City</strong> <input type="text" name="AgencyCity" readonly value="<%=AgencyCity%>" ID="Text5"> <strong>State</strong> <select name="AgencyState" readonly ID="Select1">
<option value="<%=AgencyState%>" selected><%=AgencyState%></option><!--
            <option value="AK">AK</option>
            <option value="AL">AL</option>
            <option value="AR">AR</option>
            <option value="AZ">AZ</option>
            <option value="CA">CA</option>
            <option value="CO">CO</option>
            <option value="CT">CT</option>
            <option value="DC">DC</option>
            <option value="DE">DE</option>
            <option value="FL">FL</option>
            <option value="GA">GA</option>
            <option value="HI">HI</option>
            <option value="IA">IA</option>
            <option value="ID">ID</option>
            <option value="IL">IL</option>
            <option value="IN">IN</option>
            <option value="KS">KS</option>
            <option value="KY">KY</option>
            <option value="LA">LA</option>
            <option value="MA">MA</option>
            <option value="MD">MD</option>
            <option value="ME">ME</option>
            <option value="MI">MI</option>
            <option value="MN">MN</option>
            <option value="MO">MO</option>
            <option value="MS">MS</option>
            <option value="MT">MT</option>
            <option value="NC">NC</option>
            <option value="ND">ND</option>
            <option value="NE">NE</option>
            <option value="NH">NH</option>
            <option value="NJ">NJ</option>
            <option value="NM">NM</option>
            <option value="NV">NV</option>
            <option value="NY">NY</option>
            <option value="OH">OH</option>
            <option value="OK">OK</option>
            <option value="OR">OR</option>
            <option value="PA">PA</option>
            <option value="PR">PR</option>
            <option value="RI">RI</option>
            <option value="SC">SC</option>
            <option value="SD">SD</option>
            <option value="TN">TN</option>
            <option value="TX">TX</option>
            <option value="UT">UT</option>
            <option value="VA">VA</option>
            <option value="VI">VI</option>
            <option value="VT">VT</option>
            <option value="WA">WA</option>
            <option value="WI">WI</option>
            <option value="WV">WV</option>
            <option value="WY">WY</option>
			-->
</select> <strong>Zip Code</strong> <input type="text" name="AgencyZip" size="7" maxlength="5" readonly value="<%=AgencyZip%>" ID="Text6"></td>
		<td colspan="1" class="formMain" nowrap><strong>Fiscal Year Ended</strong> <select name="FiscalYearEnded" tabindex="1" ID="Select2">
	<option value="" SELECTED>----</option>
	<option value="2006">2006</option>
	<option value="2007">2007</option>
	<option value="2008">2008</option>
	<option value="2009">2009</option>
	<option value="2010">2010</option>
	<option value="2011">2011</option>
	<option value="2012">2012</option>
	<option value="2013">2013</option>
</select></td>
		</tr>
		<tr>
		<td colspan="2" class="formMain" nowrap><strong>Form Prepared By</strong> <input type="text" name="PrepBy" size="30" maxlength="30" tabindex="2" <%if (viewMode = "view") then%>readonly<%end if%> value="<%if (viewMode = "view") then%><%=PrepBy%><%end if%>" ID="Text7"></td>
		<td colspan="1" class="formMain" nowrap><strong>Phone</strong> <input type="text" name="PrepByPhone" size="15" maxlength="15" tabindex="3" <%if (viewMode = "view") then%>readonly<%end if%> value="<%if (viewMode = "view") then%><%=PrepByPhone%><%end if%>" ID="Text20"></td>
		</tr>
		<tr>
		<td colspan="3" class="formMain">In acccordance with your affiliation agreement with Big Brothers Big Sisters of America, please calculate your agency's affiliation fees for the period <u>January 1, <%=Year(Date())+1%> through December 31, <%=Year(Date())+1%>.</u></td>
		</tr>
		<tr>
		<td colspan="3" class="formMain"><font color="#ff0000"><strong>Please note:</strong> Please write your Agency Identification Number on all fee payments.</font></td>
		</tr>
		<tr>
		<td class="formMain" nowrap><strong>(A) TOTAL EXPENDITURES</strong> (Last Completed Fiscal Year) <a href="../surveys/OnlineForms/helpfiles/feecalculation_form_help.asp?HelpID=total_expenditures" onclick="NewWindow(this.href,'name','500','275','yes');return false;"><img src="../surveys/OnlineForms/images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
		<td class="formMain" nowrap>&nbsp;</td>
		<td class="formMain" nowrap>$<input type="text" name="TotalExpenditures" size="12" maxlength="11" <%if (viewMode = "view") then%>readonly<%end if%> value="<%if (viewMode = "view") then%><%=TotalExpenditures%><%else%>0<%end if%>" onChange="CalculateFees(this.value);" tabindex="4" ID="Text8">(A)</td>
		</tr>
		<tr>
		<td class="formMain" nowrap>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(B) Less: Prior Year Fees Paid to BBBSA* <a href="../surveys/OnlineForms/helpfiles/feecalculation_form_help.asp?HelpID=prior_year_fees_paid_to_bbbsa" onclick="NewWindow(this.href,'name','500','275','yes');return false;"><img src="../surveys/OnlineForms/images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
		<td class="formMain" nowrap>$<input type="text" name="PriorYearFeesPaidToBBBSA" size="12" maxlength="11" <%if (viewMode = "view") then%>readonly<%end if%> value="<%if (viewMode = "view") then%><%=PriorYearFeesPaidToBBBSA%><%else%>0<%end if%>" onChange="CalculateFees(this.value);" tabindex="5" ID="Text9">(B)</td>
		<td rowspan="4" class="formMain" nowrap>&nbsp;</td>
		</tr>
		<tr>
		<td class="formMain" nowrap>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(C) Less: Prior Year Capital Purchases* <a href="../surveys/OnlineForms/helpfiles/feecalculation_form_help.asp?HelpID=prior_year_capital_purchases" onclick="NewWindow(this.href,'name','500','275','yes');return false;"><img src="../surveys/OnlineForms/images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
		<td class="formMain" nowrap>$<input type="text" name="PriorYearCapitalPurchases" size="12" maxlength="11" <%if (viewMode = "view") then%>readonly<%end if%> value="<%if (viewMode = "view") then%><%=PriorYearCapitalPurchases%><%else%>0<%end if%>" onChange="CalculateFees(this.value);" tabindex="6" ID="Text10">(C)</td>
		</tr>
		<tr>
		<td class="formMain" nowrap>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(D) Less: Prior Year Depreciation* <a href="../surveys/OnlineForms/helpfiles/feecalculation_form_help.asp?HelpID=prior_year_depreciation" onclick="NewWindow(this.href,'name','500','275','yes');return false;"><img src="../surveys/OnlineForms/images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
		<td class="formMain" nowrap>$<input type="text" name="PriorYearDepreciation" size="12" maxlength="11" <%if (viewMode = "view") then%>readonly<%end if%> value="<%if (viewMode = "view") then%><%=PriorYearDepreciation%><%else%>0<%end if%>" onChange="CalculateFees(this.value);" tabindex="7" ID="Text11">(D)</td>
		</tr>
		<tr>
		<td class="formMain" nowrap>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(E) Less: Prior Year Fundraising Expenses* <a href="../surveys/OnlineForms/helpfiles/feecalculation_form_help.asp?HelpID=prior_year_fundraising_expenses" onclick="NewWindow(this.href,'name','500','275','yes');return false;"><img src="../surveys/OnlineForms/images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
		<td class="formMain" nowrap>$<input type="text" name="PriorYearFundraisingExpenses" size="12" maxlength="11" <%if (viewMode = "view") then%>readonly<%end if%> value="<%if (viewMode = "view") then%><%=PriorYearFundraisingExpenses%><%else%>0<%end if%>" onChange="CalculateFees(this.value);" tabindex="8" ID="Text12">(E)</td>
		</tr>
		<tr>
		<td class="formMain" nowrap><strong>(F) TOTAL DEDUCTIONS</strong> (B + C + D + E)</td>
		<td class="formMain" nowrap bgcolor="#c0c0c0">&nbsp;</td>
		<td class="formMain" nowrap bgcolor="#c0c0c0">$<input type="text" name="TotalDeductions" size="12" value="0" readonly ID="Text13">(F)</td>
		</tr>
		<tr>
		<td class="formMain" nowrap><strong>(G) ADJUSTED EXPENDITURES</strong> (A - F)</td>
		<td class="formMain" nowrap bgcolor="#c0c0c0">&nbsp;</td>
		<td class="formMain" nowrap bgcolor="#c0c0c0">$<input type="text" name="AdjustedExpenditures" size="12" value="0" readonly ID="Text14">(G)</td>
		</tr>
		<tr bgcolor="#660099">
		<td colspan="3" class="formMain" nowrap><strong><font color="#ffffff">FEE CALCULATION:</font></strong></td>
		</tr>
		<tr>
		<td class="formMain" nowrap>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(H) 3.80% of first $100,000 of Adjusted Expenditures</td>
		<td class="formMain" nowrap bgcolor="#c0c0c0">$<input type="text" name="FeeCalc1" size="12" value="0" readonly ID="Text15">(H)</td>
		<td class="formMain" nowrap bgcolor="#c0c0c0"><font color="#eeeeee">&nbsp;&nbsp;&nbsp;Max. $3800.00</font></td>
		</tr>
		<tr>
		<td class="formMain" nowrap>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(I) 2.25% of the next $100,000 of Adjusted Expenditures</td>
		<td class="formMain" nowrap bgcolor="#c0c0c0">$<input type="text" name="FeeCalc2" size="12" value="0" readonly ID="Text16">(I)</td>
		<td class="formMain" nowrap bgcolor="#c0c0c0"><font color="#eeeeee">&nbsp;&nbsp;&nbsp;Max. $2250.00</font></td>
		</tr>
		<tr>
		<td class="formMain" nowrap>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(J) 1.00% of the next $300,000 of Adjusted Expenditures</td>
		<td class="formMain" nowrap bgcolor="#c0c0c0">$<input type="text" name="FeeCalc3" size="12" value="0" readonly ID="Text17">(J)</td>
		<td class="formMain" nowrap bgcolor="#c0c0c0"><font color="#eeeeee">&nbsp;&nbsp;&nbsp;Max. $3000.00</font></td>
		</tr>
		<tr>
		<td class="formMain" nowrap>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(K) 0.50% of the remaining Adjusted Expenditures</td>
		<td class="formMain" nowrap bgcolor="#c0c0c0">$<input type="text" name="FeeCalc4" size="12" value="0" readonly ID="Text18">(K)</td>
		<td class="formMain" nowrap bgcolor="#c0c0c0">&nbsp;</td>
		</tr>
		<tr>
		<td class="formMain" nowrap><strong>(L) CALCULATED FEE</strong> (H + I + J + K)</td>
		<td class="formMain" nowrap bgcolor="#c0c0c0">&nbsp;</td>
		<td class="formMain" nowrap bgcolor="#c0c0c0">$<input type="text" name="TotalFeeCalc" size="12" value="0" readonly ID="Text19">(L)</td>
		</tr>
		<TR>
			<td class="formMain" nowrap>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(M) Less: Annual Discount**</td>
			<td class="formMain" nowrap bgcolor="#c0c0c0">&nbsp;</td>
			<td class="formMain" nowrap bgcolor="#c0c0c0">
				<%if (viewMode = "view") then %>
				<select name="AnnualDiscount" CLASS="formMain" onChange="CalculateFees(this.value);" readonly tabindex="8" ID="Select4">
					<option value="<%=AnnualDiscount%>" selected><%=AnnualDiscount%></option>
				</select>(M)
				<%else%>
				<select name="AnnualDiscount" CLASS="formMain" onChange="CalculateFees(this.value);" tabindex="9" ID="AnnualDiscount">
					<option value="selected">Select Percent</option>
					<option value="0.050">5.0 %</option>
					<option value="0.045">4.5 %</option>
					<option value="0.040">4.0 %</option>
					<option value="0.000">None</option>
				</select>(M)
				<%end if%>
			</TD>
		</TR>
		<!--<tr>
		<td class="formMain" nowrap>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(M) Less: Annual Discount**</td>
		<td class="formMain" nowrap bgcolor="#c0c0c0">&nbsp;</td>
		<td class="formMain" nowrap bgcolor="#c0c0c0">$<input type="text" name="AnnualDiscount" size="12" value="0" readonly ID="Text20">(M)</td>
		</tr>-->
		<tr>
		<td class="formMain" nowrap><strong>(N) TOTAL AFFILIATION FEES DUE</strong> (L - M)</td>
		<td class="formMain" nowrap bgcolor="#c0c0c0">&nbsp;</td>
		<td class="formMain" nowrap bgcolor="#c0c0c0">$<input type="text" name="TotalAffiliationFeesDue" size="12" value="0" readonly ID="Text21">(N)</td>
		</tr>
		<tr>
		<td colspan="3" class="formMain"><strong>*</strong> All deductions should be substantiated and will be verified against a copy of your last audit, which must be submitted to BBBSA after the audit has been finalized.<br>
        <strong>** ANNUAL DISCOUNTS:</strong><br>
			<table ID="Table6">
				<tr>
				<td class="formMain">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
				<td class="formMain">Full payment by January 31 2009</td>
				<td class="formMain">.......</td>
				<td class="formMain">5.0% of Calculated Fee</td>
				</tr>
				<tr>
				<td class="formMain">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
				<td class="formMain">Full payment by February 28 2009</td>
				<td class="formMain">.......</td>
				<td class="formMain">4.5% of Calculated Fee</td>
				</tr>
				<tr>
				<td class="formMain">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
				<td class="formMain">Full payment by March 31 2009</td>
				<td class="formMain">.......</td>
				<td class="formMain">4.0% of Calculated Fee</td>
				</tr>
			</table>
        </td>
		</tr>
		<tr>
			<td colspan="3" class="formMain"><strong>***</strong> If you have questions about this form, please contact Peter Paiman at 215-665-7732</td>
		</tr>
		<tr align="center" bgcolor="#660099">
		<td colspan="3" class="formMain" nowrap>
		<%if (viewMode <> "view") then%>
		<input type="submit" name="SubmitForm" value="Print and Submit Form" tabindex="10" ID="Submit1">
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="reset" name="ResetForm" value="Reset Form" ID="Reset1">
		<input type="hidden" name="SQLHID" value="" id="SQLHID">
		<input type="hidden" name="Year" value="" id="Year">
		<%end if%>
		</td>
		</tr>
	</table>
</form>
<script language="javascript">
<!--
document.FeeCalculator.FiscalYearEnded.focus()
//-->
</script>
<%if (viewMode = "view") then%>
<script language="javascript">
<!--
CalculateFees(1);
//-->
</script>
<%end if%>

		  </span>
          
          </TD>
        
        
        </TR>
        
      </TABLE>          

    </TD>
  </TR>
</TABLE>          
      <!-- end search table -->
      
      
      
    </TD>
    <!-- end main content cell -->
    
    <!-- spacer cell -->
    <TD VALIGN="TOP" WIDTH="1%"><IMG SRC="../media/images/spacer.gif" WIDTH="5" HEIGHT="1" BORDER="0"></TD>    
  </TR>
</TABLE>

</body>

</html>

<%Response.Write(RightContent(myStr))%>
