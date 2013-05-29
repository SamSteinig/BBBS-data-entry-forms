<!--#include file="../includes/NAD_BE.asp" -->

<% 
Server.ScriptTimeout=900

'''''''	Month variable for Balance Sheet & Expenses Section 7/25/2009 saf '''''''
	Dim iMonth
	
	If Request("month") <> "" then 
		iMonth = CINT(Request("month"))
	Else
		iMonth = CINT(Request("m"))
	End If	
	m = iMonth
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
		RST.Open "SELECT top 1 * FROM tbl_frmFinancePerformance", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = Request("AgencyIDN")
		RST("Year") = Request("year")
		RST("Month") = iMonth
		
		RST("UnitedWay") = Request("frmFinancePerformanceUnitedWay")		
		RST("GovFederalFunding") = Request("frmFinancePerformanceGovFederalFunding")		
		RST("GovStateFunding") = Request("frmFinancePerformanceGovStateFunding")					
		RST("GovLocalFunding") = Request("frmFinancePerformanceGovLocalFunding")					
		RST("FoundationGrants") = Request("frmFinancePerformanceFoundationGrants")	
		RST("CorporateGifts") = Request("frmFinancePerformanceCorporateGifts")			
		RST("BBBSAGrants") = Request("frmFinancePerformanceBBBSAGrants")		
		RST("IndividualGiving") = Request("frmFinancePerformanceIndividualGiving")		
		RST("EventsTotal") = Request("frmFinancePerformanceEventsTotal")				
		RST("EventsIndiv") = Request("frmFinancePerformanceEventsIndiv")
		RST("EventsCorp") = Request("frmFinancePerformanceEventsCorp")				
		RST("DividendsInterest") = Request("frmFinancePerformanceDividendsInterest")					
		RST("Other") = Request("frmFinancePerformanceOther")	
		RST("TotalGross") = Request("frmFinancePerformanceTotalGross")		
		RST("SpecEventExp") = Request("frmFinancePerformanceSpecEventExp")				
		RST("Total") = Request("frmFinancePerformanceTotal")						
		RST("TotalAmountRMM") = Request("frmFinancePerformanceTotalAmountRMM")			
		RST("TotalAmountBFKS") = Request("frmFinancePerformanceTotalAmountBFKS")
		
		If iMonth = 12 then
			RST("TotalOperatingExpense") = Request("frmExpensesTotal")
		else
			RST("TotalOperatingExpense") = Request("frmFinancePerformanceTotalOperatingExpense")
		End If
		
		RST("CreateDate") = Now
		RST.Update
		RST.Close
		Set RST = Nothing
		form = "FinancePerformance"
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
	RST.Open "SELECT * FROM tbl_frmFinancePerformance WHERE agencyID='" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")) & " AND Month=" & Int(iMonth), Con, 1, 3

	RST("UnitedWay") = Request("frmFinancePerformanceUnitedWay")		
	RST("GovFederalFunding") = Request("frmFinancePerformanceGovFederalFunding")		
	RST("GovStateFunding") = Request("frmFinancePerformanceGovStateFunding")					
	RST("GovLocalFunding") = Request("frmFinancePerformanceGovLocalFunding")					
	RST("FoundationGrants") = Request("frmFinancePerformanceFoundationGrants")	
	RST("CorporateGifts") = Request("frmFinancePerformanceCorporateGifts")			
	RST("BBBSAGrants") = Request("frmFinancePerformanceBBBSAGrants")		
	RST("IndividualGiving") = Request("frmFinancePerformanceIndividualGiving")		
	RST("EventsTotal") = Request("frmFinancePerformanceEventsTotal")				
	RST("EventsIndiv") = Request("frmFinancePerformanceEventsIndiv")
	RST("EventsCorp") = Request("frmFinancePerformanceEventsCorp")						
	RST("DividendsInterest") = Request("frmFinancePerformanceDividendsInterest")					
	RST("Other") = Request("frmFinancePerformanceOther")				
	RST("TotalGross") = Request("frmFinancePerformanceTotalGross")						
	RST("SpecEventExp") = Request("frmFinancePerformanceSpecEventExp")											
	RST("Total") = Request("frmFinancePerformanceTotal")						
	RST("TotalAmountRMM") = Request("frmFinancePerformanceTotalAmountRMM")			
	RST("TotalAmountBFKS") = Request("frmFinancePerformanceTotalAmountBFKS")			

	If iMonth = 12 then
		RST("TotalOperatingExpense") = Request("frmExpensesTotal")
	else
		RST("TotalOperatingExpense") = Request("frmFinancePerformanceTotalOperatingExpense")
	End If
	
	jMod = RST("FinancePerformanceID") 
	
	RST.Update
	RST.Close
	Set RST = Nothing
	form = "FinancePerformance"
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

'''''''	Balance Sheet & Expenses Section 3/23/2009 saf '''''''
	Dim rsBalanceSheet, strSQL, CashInvestments, Receivables, AllOtherAssets, TotalAssets, STLiabilities, LTLiabilities
	Dim TotalLiabilities, Surplus_NetAssets, TotalLiabilitiesAndNetAssets, iQtr, rsExpenses

	If Request("status") = "addNew" or Request("status") = "editSave" Then
		Select case iMonth 
		    case 12 'Update Annual Expenses section
				Set DupCon = Server.CreateObject("ADODB.Connection")
				DupCon.Open "BBBSAforms", "sa","12sist12"
				strSQL = "SELECT Count(*) FROM tbl_frmExpenses WHERE AgencyID='" & Session("AgencyIDN") & "' AND [Year]=" & Int(Request("year")) 
				Set rsExpenses = DupCon.Execute(strSQL)
				numberOfExisting = rsExpenses(0)
				rsExpenses.Close
				DupCon.CLose
				Set rsExpenses = Nothing
				Set DUpCon = Nothing

		    If numberOfExisting = 0 then 'add

				Set Con = Server.CreateObject("ADODB.Connection")
				Con.Open "BBBSAforms", "sa","12sist12"
				Set RST = Server.CreateObject("ADODB.Recordset")
				RST.Open "SELECT top 1 * FROM tbl_frmExpenses", Con, 1, 3
				RST.AddNew
				RST("AgencyID") = Request("AgencyIDN")
				RST("Year") = Request("year")
		
				tt = Int(Request("frmExpensesSalariesWages")) _
					+ Int(Request("frmExpensesEmployeeBenefits")) _
					+ Int(Request("frmExpensesInsurance")) _
					+ Int(Request("frmExpensesRent")) _
					+ Int(Request("frmExpensesOther"))

				RST("SalariesWages") = FormatCurrency(Request("frmExpensesSalariesWages"))
				RST("EmployeeBenefits") = FormatCurrency(Request("frmExpensesEmployeeBenefits"))
				RST("Insurance") = FormatCurrency(Request("frmExpensesInsurance"))
				RST("RentOccupancy") = FormatCurrency(Request("frmExpensesRent"))
				RST("Other") = FormatCurrency(Request("frmExpensesOther"))
				RST("Total") = FormatCurrency(tt)
		
				RST("Administration") = Request("frmExpensesAdministration")
				RST("Program") = Request("frmExpensesProgram")
				RST("FundRaising") = Request("frmExpensesFundRaising")
		
				CategoryTotal = Int(Request("frmExpensesAdministration")) _
								  + Int(Request("frmExpensesProgram")) _
								  + Int(Request("frmExpensesFundRaising")) 

				RST("CategoryTotal") = CategoryTotal

				RST("CreateDate") = Now
				RST.Update
				RST.Close
				Set RST = Nothing
				
				form = "Expenses"
				modtype = "new(m)"

				Set RST = Server.CreateObject("ADODB.Recordset")
				RST.Open "SELECT top 1 ExpensesID FROM tbl_frmExpenses WHERE AgencyID = '" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")), Con, 1, 3	
				jMod = RST("ExpensesID")
				RST.Close
				Set RST = Nothing
Response.Write("jMod " & jMod)
Response.Flush
				%>
				<!--#include file="../includes/modify_stamp.asp"-->
				<%	
				Con.Close
				Set Con = Nothing
				say = "thanks"

			Else 'update

				Set Con = Server.CreateObject("ADODB.Connection")
				Con.Open "BBBSAforms", "sa","12sist12"
				Set RST = Server.CreateObject("ADODB.Recordset")
				RST.Open "SELECT * FROM tbl_frmExpenses WHERE AgencyID = '" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")), Con, 1, 3
					
				tt = Int(Request("frmExpensesSalariesWages")) _
						+ Int(Request("frmExpensesEmployeeBenefits")) _
						+ Int(Request("frmExpensesInsurance")) _
						+ Int(Request("frmExpensesRent")) _
						+ Int(Request("frmExpensesOther"))

				RST("SalariesWages") = FormatCurrency(Request("frmExpensesSalariesWages"))
				RST("EmployeeBenefits") = FormatCurrency(Request("frmExpensesEmployeeBenefits"))
				RST("Insurance") = FormatCurrency(Request("frmExpensesInsurance"))
				RST("RentOccupancy") = FormatCurrency(Request("frmExpensesRent"))
				RST("Other") = FormatCurrency(Request("frmExpensesOther"))
				RST("Total") = FormatCurrency(tt)	
				
				RST("Administration") = Request("frmExpensesAdministration")
				RST("Program") = Request("frmExpensesProgram")
				RST("FundRaising") = Request("frmExpensesFundRaising")

				CategoryTotal = Int(Request("frmExpensesAdministration")) _
								  + Int(Request("frmExpensesProgram")) _
								  + Int(Request("frmExpensesFundRaising")) 

				RST("CategoryTotal") = CategoryTotal
								
				jMod = RST("ExpensesID")
				RST.Update
				RST.Close
				Set RST = Nothing
				form = "Expenses"
				modtype = "edit"
				%>
				<!--#include file="../includes/modify_stamp.asp"-->
				<%	
				Con.Close
				Set Con = Nothing
				say = "thanks"
			End If
				
		End Select
		
		Select case iMonth 
		    case 3, 6, 9, 12
		    
				iQtr = iMonth \ 3					
				Set Con = Server.CreateObject("ADODB.Connection")
				Con.Open "BBBSAforms", "sa","12sist12"
				strSQL = "SELECT Count(*) FROM tbl_frmBalanceSheet WHERE AgencyID=" & Session("AgencyIDN") & " AND Yr=" & Int(Request("year")) & " AND Mth=" & Int(iMonth)
				Set rsBalanceSheet = Con.Execute(strSQL)

				'Check for Existing records
				If rsBalanceSheet(0) = 0 then 'add
					strSQL = "INSERT INTO tbl_frmBalanceSheet (AgencyID, Yr, Qtr, Mth, CashInvestments, Receivables, AllOtherAssets, TotalAssets, ShortTermLiabilities, " & _
							 "                                 LongTermLiabilities, TotalLiabilities, Surplus_NetAssets, TotalLiabNetAssets)" & _
							 "VALUES ("  & Session("AgencyIDN") & _
							         "," & Request("year") & _							 
							         "," & iQTR & _							 
							         "," & iMonth & _							 
							         "," & Request("CashInvestments") & _
							         "," & Request("Receivables") & _
							         "," & Request("AllOtherAssets") & _
							         "," & Request("TotalAssets") & _
							         "," & Request("STLiabilities") & _
							         "," & Request("LTLiabilities") & _
							         "," & Request("TotalLiabilities") & _
							         "," & Request("Surplus_NetAssets") & _
							         "," & Request("TotalLiabilitiesAndNetAssets") & ")"
							 	
					Set rsBalanceSheet = Con.Execute(strSQL)
					modtype = "new(m)"
				Else 'update
					strSQL = "UPDATE tbl_frmBalanceSheet " & _
							 " set CashInvestments = "  & Request("CashInvestments") & _
							 " ,Receivables = " & Request("Receivables") & _
							 " ,AllOtherAssets = " & Request("AllOtherAssets") & _
							 " ,TotalAssets = " & Request("TotalAssets") & _
							 " ,ShortTermLiabilities = " & Request("STLiabilities") & _
							 " ,LongTermLiabilities = " & Request("LTLiabilities") & _
							 " ,TotalLiabilities = " & Request("TotalLiabilities") & _
							 " ,Surplus_NetAssets = " & Request("Surplus_NetAssets") & _
							 " ,TotalLiabNetAssets = " & Request("TotalLiabilitiesAndNetAssets") & _
							 " where Yr = " & Request("year") & _
							 "   and Mth = " & iMonth & _
							 "   and AgencyID = " & Session("AgencyIDN")
							  
							 	
					Set rsBalanceSheet = Con.Execute(strSQL)
										
					modtype = "edit"
				End If

				'Add Record to Modify Log
				Set RST = Server.CreateObject("ADODB.Recordset")
				RST.Open "SELECT top 1 BalanceSheetID FROM tbl_frmBalanceSheet WHERE AgencyID = '" & Request("AgencyIDN") & "' AND Yr=" & Int(Request("year")) & " and Qtr = " & iQTR, Con, 1, 3	
				jMod = RST("BalanceSheetID")
				form = "BalanceSheet"
				RST.Close
				Set RST = Nothing
				' %>
				<!--#include file="../includes/modify_stamp.asp"-->
				<%	
				Con.Close
				Set Con = Nothing
				say = "thanks"
				
			case else 'Do Nothing
		End Select
	End If
'''''''''''''''''''''''''''''''''''''''''''''''''''
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
	var box11 = Number(document.frmFinancePerformance.frmFinancePerformanceEventsTotal.value)
	var boxtotal = box1 + box2 + box3 + box4 + box5 + box6 + box7 + box8 + box9 + box10 + box11
	document.frmFinancePerformance.frmFinancePerformanceTotalGross.value = boxtotal
	
}


function TotalNet() {
	var box1 = Number(document.frmFinancePerformance.frmFinancePerformanceTotalGross.value)
	var box2 = Number(document.frmFinancePerformance.frmFinancePerformanceSpecEventExp.value)	
	var boxtotal = box1 - box2
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

function addEventBreakouts() {
	var box1 = Number(document.frmFinancePerformance.frmFinancePerformanceEventsIndiv.value)
	var box2 = Number(document.frmFinancePerformance.frmFinancePerformanceEventsCorp.value)
	var boxtotal = box1 + box2
	if (boxtotal > Number(document.frmFinancePerformance.frmFinancePerformanceEventsTotal.value))	
	{
		alert("Amounts of Individual and Corporate Events Cannot Exceed Total Event Value.");
	}
	
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
	var box11 = Number(document.frmFinancePerformance.frmFinancePerformanceEventsTotal.value)
	var boxtotal = box1 + box2 + box3 + box4 + box5 + box6 + box7 + box8 + box9 + box10 + box11
	document.frmFinancePerformance.frmFinancePerformanceTotalGross.value = boxtotal	
	
	
	
	var Netbox1 = Number(document.frmFinancePerformance.frmFinancePerformanceTotalGross.value)
	var Netbox2 = Number(document.frmFinancePerformance.frmFinancePerformanceSpecEventExp.value)	
	var Netboxtotal = Netbox1 - Netbox2
	document.frmFinancePerformance.frmFinancePerformanceTotal.value = Netboxtotal		
	
	
	
	
//	if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceUnitedWay.value)) || document.frmFinancePerformance.frmFinancePerformanceUnitedWay.value =="")

	if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceUnitedWay.value)) || document.frmFinancePerformance.frmFinancePerformanceUnitedWay.value =="")
	{
		alert("Error - United Way Field.\nPlease make sure you have entered a whole number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceUnitedWay.focus();				
	} 

	else	
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceGovFederalFunding.value)) || document.frmFinancePerformance.frmFinancePerformanceGovFederalFunding.value =="")
	{
		alert("Error - Federal Funding Field.\nPlease make sure you have entered a whole number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceGovFederalFunding.focus();		
	}

	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceGovStateFunding.value)) || document.frmFinancePerformance.frmFinancePerformanceGovStateFunding.value =="")
	{
		alert("Error - State Funding Field.\nPlease make sure you have entered a whole number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceGovStateFunding.focus();
	}	

	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceGovLocalFunding.value)) || document.frmFinancePerformance.frmFinancePerformanceGovLocalFunding.value =="")
	{
		alert("Error - Local Funding Field.\nPlease make sure you have entered a whole number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceGovLocalFunding.focus();
	}	

	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceFoundationGrants.value)) || document.frmFinancePerformance.frmFinancePerformanceFoundationGrants.value =="")
	{
		alert("Error - Foundation Grants Field.\nPlease make sure you have entered a whole number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceFoundationGrants.focus();
	}	

	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceCorporateGifts.value)) || document.frmFinancePerformance.frmFinancePerformanceCorporateGifts.value =="")
	{
		alert("Error - Corporate Gifts Field.\nPlease make sure you have entered a whole number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceCorporateGifts.focus();
	}

	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceBBBSAGrants.value)) || document.frmFinancePerformance.frmFinancePerformanceBBBSAGrants.value =="")
	{
		alert("Error - BBBSA (Pass-Through) Grants Field.\nPlease make sure you have entered a whole number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceBBBSAGrants.focus();
	}	
	
	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceIndividualGiving.value)) || document.frmFinancePerformance.frmFinancePerformanceIndividualGiving.value =="")
	{
		alert("Error - Individual Giving Field.\nPlease make sure you have entered a whole number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceIndividualGiving.focus();
	}		
	
	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceDividendsInterest.value)) || document.frmFinancePerformance.frmFinancePerformanceDividendsInterest.value =="")
	{
		alert("Error - Dividends and Interest Field.\nPlease make sure you have entered a whole number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceDividendsInterest.focus();
	}	
	
	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceOther.value)) || document.frmFinancePerformance.frmFinancePerformanceOther.value =="")
	{
		alert("Error - Other Field.\nPlease make sure you have entered a whole number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceOther.focus();
	}		
	
	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceTotalAmountBFKS.value)) || document.frmFinancePerformance.frmFinancePerformanceTotalAmountBFKS.value =="")
	{
		alert("Error - BFKS Field.\nPlease make sure you have entered a whole number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceTotalAmountBFKS.focus();
	}	

	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceTotalAmountRMM.value)) || document.frmFinancePerformance.frmFinancePerformanceTotalAmountRMM.value =="")
	{
		alert("Error - RMM Field.\nPlease make sure you have entered a whole number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceTotalAmountRMM.focus();
	}			

	else
		if( Number(document.frmFinancePerformance.frmFinancePerformanceEventsIndiv.value) + Number(document.frmFinancePerformance.frmFinancePerformanceEventsCorp.value) > Number(document.frmFinancePerformance.frmFinancePerformanceEventsTotal.value) )
	{
		alert("Amounts of Individual and Corporate Events Cannot Exceed Total Event Value.")
		document.frmFinancePerformance.frmFinancePerformanceEventsTotal.focus();
	}

<% ''Validate quarterly data
  Select Case Request("m") 
	Case 3, 6, 9, 12  
%>
	else
		if( Number(document.frmFinancePerformance.TotalLiabilitiesAndNetAssets.value) != Number(document.frmFinancePerformance.TotalAssets.value) )
	{
		alert("Balance sheet must balance. Total Assets should equal total Liabilities + NetAssets.")
		document.frmFinancePerformance.CashInvestments.focus();
	}

	else
		if( Number(document.frmFinancePerformance.frmExpensesCategoryTotal.value) != 100)
	{
		alert("Expense Breakdown by Category must equal 100%")
		document.frmFinancePerformance.frmExpensesCategoryTotal.focus();
	}

<%	End Select

 ''Validate Annual Expenses data
  Select Case Request("m") 
	Case 12  
%>
	else
		if( Number(document.frmFinancePerformance.TotalLiabilitiesAndNetAssets.value) != Number(document.frmFinancePerformance.TotalAssets.value) )
	{
		alert("Balance sheet must balance. Total Assets should equal total Liabilities + NetAssets.")
		document.frmFinancePerformance.CashInvestments.focus();
	}

<%  Case Else %>
	else
		if (!(onlyInteger.test(document.frmFinancePerformance.frmFinancePerformanceTotalOperatingExpense.value)) || document.frmFinancePerformance.frmFinancePerformanceTotalOperatingExpense.value =="")
	{
		alert("Error - Total Operating Expense Field.\nPlease make sure you have entered a whole number with no spaces, decimal points or commas. Do not leave this field blank."); 
		document.frmFinancePerformance.frmFinancePerformanceTotalOperatingExpense.focus();
	}			
<% End Select %>
		
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

<%Select case Request("m") 
    case 3, 6, 9, 12 'Add %>
		<script language="javascript">
			<!--	

			function addUpBalanceSheet() {
				var CashInvestments = Number(document.frmFinancePerformance.CashInvestments.value)
				var Receivables = Number(document.frmFinancePerformance.Receivables.value)
				var AllOtherAssets = Number(document.frmFinancePerformance.AllOtherAssets.value)	
				var TotalAssets = CashInvestments + Receivables + AllOtherAssets
				var STLiabilities = Number(document.frmFinancePerformance.STLiabilities.value)
				var LTLiabilities = Number(document.frmFinancePerformance.LTLiabilities.value)
				var TotalLiabilities = 	STLiabilities + LTLiabilities
				var Surplus_NetAssets = Number(document.frmFinancePerformance.Surplus_NetAssets.value)	
				var TotalLiabilitiesAndNetAssets = TotalLiabilities + Surplus_NetAssets

				document.frmFinancePerformance.TotalAssets.value = TotalAssets
				document.frmFinancePerformance.TotalLiabilities.value = TotalLiabilities
				document.frmFinancePerformance.TotalLiabilitiesAndNetAssets.value = TotalLiabilitiesAndNetAssets
			}

<%    If Request("m") = 12 Then %>

			function addExpensesUp() {
				var box1 = Number(document.frmFinancePerformance.frmExpensesSalariesWages.value)
				var box2 = Number(document.frmFinancePerformance.frmExpensesEmployeeBenefits.value)
				var box3 = Number(document.frmFinancePerformance.frmExpensesInsurance.value)	
				var box4 = Number(document.frmFinancePerformance.frmExpensesOther.value)
				var box5 = Number(document.frmFinancePerformance.frmExpensesRent.value)
				var boxtotal = box1 + box2 + box3 + box4 + box5
				document.frmFinancePerformance.frmExpensesTotal.value = boxtotal
			}
			
			function AddUpAssets()
			{
				var box1 = Number(document.frmFinancePerformance.frmExpensesCashInvestments.value)
				var box2 = Number(document.frmFinancePerformance.frmExpensesReceivables.value)
				var box3 = Number(document.frmFinancePerformance.frmExpensesAllOtherAssets.value)	

				var boxtotal = box1 + box2 + box3
				document.frmFinancePerformance.frmExpensesTotalAssets.value = boxtotal
			}

			function AddUpLiabilities()
			{
				//var box1 = Number(document.frmExpenses.frmExpensesLiabilities.value)
				var box2 = Number(document.frmFinancePerformance.frmExpensesSurplus_NetAssets.value)
				var box3 = Number(document.frmFinancePerformance.frmExpensesLiabilitiesShort.value)
				var box4 = Number(document.frmFinancePerformance.frmExpensesLiabilitiesLong.value)

				var boxtotal = box3 + box4
				document.frmFinancePerformance.frmExpensesLiabilities.value = boxtotal
				boxtotal = boxtotal + box2
				document.frmFinancePerformance.frmExpensesTotalLiabNetAssets.value = boxtotal
			}

			function addUpCategory()
			{
				var box1 = Number(document.frmFinancePerformance.frmExpensesAdministration.value)
				var box2 = Number(document.frmFinancePerformance.frmExpensesProgram.value)
				var box3 = Number(document.frmFinancePerformance.frmExpensesFundRaising.value)	

				var boxtotal = box1 + box2 + box3
				document.frmFinancePerformance.frmExpensesCategoryTotal.value = boxtotal
				
			}

			function addUpExpenseMentNonMent()
			{
				var box1 = Number(document.frmFinancePerformance.frmExpensesTotalExpenseMentoring.value)
				var box2 = Number(document.frmFinancePerformance.frmExpensesTotalExpenseNonMentoring.value)

				var boxtotal = box1 + box2
				document.frmFinancePerformance.frmExpensesExpensesMentNonMentTotal.value = boxtotal
				
			}

			function addUpFTEProgram()
			{
				var box1 = Number(document.frmFinancePerformance.frmExpensesFTECommunity.value)
				var box2 = Number(document.frmFinancePerformance.frmExpensesFTESchool.value)
				var box3 = Number(document.frmFinancePerformance.frmExpensesFTESite.value)	

				var boxtotal = box1 + box2 + box3
				document.frmFinancePerformance.frmExpensesFTEProgramTotal.value = boxtotal
				
			}

			function addUpFTEFunction()
			{
				var box1 = Number(document.frmFinancePerformance.frmExpensesFTECustomerRelations.value)
				var box2 = Number(document.frmFinancePerformance.frmExpensesFTEEnrollmentMatching.value)
				var box3 = Number(document.frmFinancePerformance.frmExpensesFTEMatchSupport.value)	

				var boxtotal = box1 + box2 + box3
				document.frmFinancePerformance.frmExpensesFTEFunctionTotal.value = boxtotal
				
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
				if((valueToCheck < 101) && (valueToCheck >= 0))
				{
					return true;	
				}
				else
				{
					alert("Percentage cannot be greater than 100 or less than 0.");
					return false;
				}
			}

	<%End If%>
			//-->
		</script>    

<%    Case Else 
		'Add no Balance Sheet 
  End Select
%>

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
			<td colspan="6" class="formHeader">Monthly Revenue / Expense<BR><%= MonthName(Request("m"), False) & " " & Request("y") %></td>
		</tr>
		
		<tr>
			<td colspan="6" class="formMain"><font color="#ff0000"><div align="center"><strong>Please Note: </strong>After entering your information, you <strong>must</strong> click on the "Save" button at the bottom of the form and wait for the "Thank You" screen or your changes will be lost.</div></font></td>
		</tr>


			<tr>
				<td valign="middle" align="center" class="formHeaderMedium" colspan="62">REVENUE</td>
			</tr>
			
			<tr>
				<td valign="middle" align="center" class="formMain" colspan="6"><p><i>Please report revenue you booked for <u><%= MonthName(Request("m"), False) & " " & Request("y") %></u> according to SOURCE of the funds. Special Events revenue should be broken out by source (e.g. individuals pledges, or corporate gifts). In addition to reporting BFKS and RMM revenue by SOURCE, please provide a total of their monthly revenue from all sources.</p><p>Click on the <img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"> next to each line item for a detailed explanation.</p></i></td>
			</tr>			

			<tr>
				<td valign="middle" class="formMain" colspan="5">
				<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_united_way" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;United Way
				</td>
				<td valign="middle" class="formMain">$
					<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("UnitedWay") %><% Else %>0<% End If %>" name="frmFinancePerformanceUnitedWay"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">				
				</td>			
			</tr>	
			
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">
				<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_gov_federal_funding" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Government - Federal Funding
				</td>
				<td valign="middle" class="formMain">$				
					<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("GovFederalFunding") %><% Else %>0<% End If %>" name="frmFinancePerformanceGovFederalFunding"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">								
				</td>					
			</tr>	
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">
				<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_gov_state_funding" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Government - State Funding
				</td>
				<td valign="middle" class="formMain">$				
					<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("GovStateFunding") %><% Else %>0<% End If %>" name="frmFinancePerformanceGovStateFunding"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">								
				</td>					
			</tr>				

			<tr>
				<td valign="middle" class="formMain" colspan="5">
				<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_gov_local_funding" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Government - Local Funding
				</td>
				<td valign="middle" class="formMain">$				
					<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("GovLocalFunding") %><% Else %>0<% End If %>" name="frmFinancePerformanceGovLocalFunding"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">								
				</td>					
			</tr>	
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">
				<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_foundations_grants" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Foundations - Grants
				</td>
				<td valign="middle" class="formMain">$				
					<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("FoundationGrants") %><% Else %>0<% End If %>" name="frmFinancePerformanceFoundationGrants"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">								
				</td>					
			</tr>				
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">
				<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_corporations" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Corporations - Non-event Donations
				</td>
				<td valign="middle" class="formMain">$				
					<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CorporateGifts") %><% Else %>0<% End If %>" name="frmFinancePerformanceCorporateGifts"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">								
				</td>
			</tr>				
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">
				<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_bbbsa_grants" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;BBBSA (Pass-Through) Grants
				</td>
				<td valign="middle" class="formMain">$				
					<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("BBBSAGrants") %><% Else %>0<% End If %>" name="frmFinancePerformanceBBBSAGrants"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">								
				</td>				
			</tr>		
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">
				<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_individual_giving" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Individual Giving (Non-Event)
				</td>
				<td valign="middle" class="formMain">$				
					<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("IndividualGiving") %><% Else %>0<% End If %>" name="frmFinancePerformanceIndividualGiving"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">								
				</td>				
			</tr>	
			
			<tr>
				<td valign="middle" class="formMain" colspan="5"><a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_events" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Events</td>
				<td valign="middle" class="formMain">$
					<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("EventsTotal") %><% Else %>0<% End If %>" name="frmFinancePerformanceEventsTotal"  onchange="checkForIntegerCommas(this.value);" >				
				</td>				
			</tr>
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">
				&nbsp;&nbsp;&nbsp;&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_events_individual" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;of Total Events, Portion From Individuals
				</td>
				<td valign="middle" class="formMain">$
					<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("EventsIndiv") %><% Else %>0<% End If %>" name="frmFinancePerformanceEventsIndiv"  onchange="checkForIntegerCommas(this.value);"  onblur="addEventBreakouts();">								
				</td>				
			</tr>				
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">
				&nbsp;&nbsp;&nbsp;&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_events_corporations" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;of Total Events, Portion From Corporations</td>
				<td valign="middle" class="formMain">$
					<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("EventsCorp") %><% Else %>0<% End If %>" name="frmFinancePerformanceEventsCorp"  onchange="checkForIntegerCommas(this.value);"  onblur="addEventBreakouts();">				
				</td>				
			</tr>											
							
			<tr>
				<td valign="middle" class="formMain" colspan="5">
				<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_events_dividends_interest" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Dividends and Interest
				</td>
				<td valign="middle" class="formMain">$				
					<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("DividendsInterest") %><% Else %>0<% End If %>" name="frmFinancePerformanceDividendsInterest"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">
				</td>								
			</tr>
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">
				<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_other" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Other
				</td>
				<td valign="middle" class="formMain">$				
					<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Other") %><% Else %>0<% End If %>" name="frmFinancePerformanceOther"  onchange="checkForIntegerCommas(this.value);" onblur="addEmUp();">
				</td>												
			</tr>		
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">Total Gross Revenue</td>
				<td valign="middle" class="formMain" bgcolor="#c0c0c0">$				
					<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("TotalGross") %><% Else %>0<% End If %>" name="frmFinancePerformanceTotalGross" onchange="noChange();" readonly>
					<span class="formSubHead">&nbsp;&nbsp;&nbsp;Calculated</span>
				</td>												
			</tr>					
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">
				<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_event_expenses" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Total Direct <strong>Expenses</strong> From Special Event Fundraising
				</td>
				<td valign="middle" class="formMain">$
					<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SpecEventExp") %><% Else %>0<% End If %>" name="frmFinancePerformanceSpecEventExp"  onchange="checkForIntegerCommas(this.value);"  onblur="TotalNet();">								
				</td>				
			</tr>				
			
			<tr>
				<td valign="middle" class="formMain" colspan="5"><strong>TOTAL NET REVENUE</strong></td>
				<td valign="middle" class="formMain" bgcolor="#c0c0c0">$				
					<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Total") %><% Else %>0<% End If %>" name="frmFinancePerformanceTotal" onchange="noChange();"  onblur="TotalNet();" readonly>
					<span class="formSubHead">&nbsp;&nbsp;&nbsp;Calculated</span>
				</td>												
			</tr>								

			<tr>
				<td valign="middle" class="formMain" colspan="5">
				&nbsp;&nbsp;&nbsp;&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_event_BFKS" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Of Total, Net Amount Raised Through BFKS</td>
				<td valign="middle" class="formMain">$				
					<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("TotalAmountBFKS") %><% Else %>0<% End If %>" name="frmFinancePerformanceTotalAmountBFKS"  onchange="checkForIntegerCommas(this.value);" onblur="addUpBreakouts();">
				</td>								
			</tr>		
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">
				&nbsp;&nbsp;&nbsp;&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_event_RMM" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Of Total, Net Amount Raised Through RMM</td>
				<td valign="middle" class="formMain">$				
					<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("TotalAmountRMM") %><% Else %>0<% End If %>" name="frmFinancePerformanceTotalAmountRMM"  onchange="checkForIntegerCommas(this.value);" onblur="addUpBreakouts();">
				</td>								
			</tr>			
			
<%''''' Expenses section added 4/12/2009 saf
	Select case iMonth
	    Case 12
			'Get data for year, month from tbl_frmExpenses
			Set Con = Server.CreateObject("ADODB.Connection")
			Con.Open "BBBSAforms", "sa","12sist12"
			query = "SELECT * FROM tbl_frmExpenses WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
			Set GetExpenses = Con.Execute(query)

%>
			<tr>
					<td colspan="6" class="formHeaderSmall"><strong>EXPENSES</strong></td>
			</tr>
			<tr>
				<td colspan="6" class="formSubhead" align="center">Do Not Include Direct Expenses From Fundraising - No Cents</td>
			</tr>
			<tr>
				<td class="formMain" width="24%">Salaries and Wages:</td>
				<td class="formMainRightJ" width="24%">$&nbsp;<input type="text" class="formMain" size="10" maxlength="25" value="<% If say = "edit" Then %><%= GetExpenses("SalariesWages") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesSalariesWages" onFocus="addExpensesUp();" onchange="checkForInteger(this.value);"></td>
				<td colspan="2" width="4%">&nbsp;</td>		
				<td class="formMain" width="24%">Employee Benefits:</td>
				<td class="formMainRightJ" width="24%">$&nbsp;<input type="text" class="formMain" size="10" maxlength="25" value="<% If say = "edit" Then %><%= GetExpenses("EmployeeBenefits") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesEmployeeBenefits" onFocus="addExpensesUp();" onchange="checkForInteger(this.value);"></td>

			</tr>
			<tr>
				<td class="formMain" width="24%">Insurance:</td>				
				<td class="formMainRightJ" width="24%">$&nbsp;<input type="text" class="formMain" size="10" maxlength="25" value="<% If say = "edit" Then %><%= GetExpenses("Insurance") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesInsurance" onFocus="addExpensesUp();" onchange="checkForInteger(this.value);"></td>
				<td colspan="2" width="4%">&nbsp;</td>
				<td class="formMain" width="24%">Other:</td>								
				<td class="formMainRightJ" width="24%">$&nbsp;<input type="text" class="formMain" size="10" maxlength="25" value="<% If say = "edit" Then %><%= GetExpenses("Other") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesOther" onFocus="addExpensesUp();" onchange="checkForInteger(this.value);"></td>										

			</tr>
			<tr>
				<td class="formMain" width="24%">Rent/Occupancy:</td>				
				<td class="formMainRightJ" width="24%">$&nbsp;<input type="text" class="formMain" size="10" maxlength="25" value="<% If say = "edit" Then %><%= GetExpenses("RentOccupancy") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesRent" onFocus="addExpensesUp();" onchange="checkForInteger(this.value);"></td>
				<td class="formMain" colspan="2">&nbsp;</td>
				<td class="formMain" width="24%" bgcolor="#c0c0c0"><strong>Total Operating Expenses</strong></td>
				<td class="formMainRightJ" width="24%" bgcolor="#c0c0c0">$&nbsp;<input type="text" class="formMain" size="10" maxlength="25" value="<% If say = "edit" Then %><%= GetExpenses("Total") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesTotal" onFocus="addExpensesUp();" onchange="addExpensesUp();" readonly><br><span class="formSubHead">calculated by system</span></td>		
			</tr>

			<tr>
				<td class="formHeaderSmall" colspan="6">EXPENSE BREAKDOWN BY CATEGORY<br>(Enter whole numbers only)</td>
			</tr>

			<tr>
				<td class="formMain" colspan="5">Program:<br>
				<font class="formSubhead"><i>Including time spent supervising program staff</i></font></td>
				<td class="formMain" valign="top" align="left" colspan="1">&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" size="4" maxlength="18" value="<% If say = "edit" Then %><%= GetExpenses("Program") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesProgram" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value); addUpCategory();" onFocus="addUpCategory();">&nbsp;&#37;&nbsp;</td>
			</tr>
			<tr>
				<td class="formMain" colspan="5">Fundraising:</td>
				<td class="formMain" valign="top" align="left" colspan="1">&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" size="4" maxlength="18" value="<% If say = "edit" Then %><%= GetExpenses("FundRaising") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesFundRaising" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value); addUpCategory();" onFocus="addUpCategory();">&nbsp;&#37;&nbsp;</td>
			</tr>	

			<tr>
				<td class="formMain" colspan="5">Administration:<br>
				<font class="formSubhead"><i>If any administration expenses are related to program or fundraising then include those expenses in program or fundraising when calculating percentages.</i></font></td>
				<td class="formMain" valign="top" align="left" colspan="1">&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" size="4" maxlength="18" value="<% If say = "edit" Then %><%= GetExpenses("Administration") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesAdministration" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value); addUpCategory(); " onFocus="addUpCategory();">&nbsp;&#37;&nbsp;</td>
			</tr>			
				
			<tr>
				<td class="formMain" colspan="5" bgcolor="#c0c0c0"><strong>TOTAL</strong></td>
				<td class="formMain" valign="top" align="left" bgcolor="#c0c0c0">&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" size="4" maxlength="4" value="<% If say = "edit" Then %><%= GetExpenses("CategoryTotal") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesCategoryTotal" onFocus="addUpCategory();" onchange="addUpCategory();" readonly>&nbsp;&#37;&nbsp;<br><span class="formSubHead">calculated by system</span></td>
			</tr>	
							
<%	Case Else %>
			<tr>
				<td valign="middle" align="center" class="formHeaderMedium" colspan="6">EXPENSE</td>
			</tr>
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">Total Operating Expense<br>(should not include expense directly related to fundraising events)</td>
				<td valign="middle" class="formMain">$				
					<input type="text"  class="formMain"  size="10" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("TotalOperatingExpense") %><% Else %>0<% End If %>" name="frmFinancePerformanceTotalOperatingExpense"  onchange="checkForIntegerCommas(this.value);" >
				</td>				
			</tr>	
<%
	End Select
'''''''''''''''''''''''''''''''''''''''''''''''''
''''' Balance Sheet section added 3/23/2009 saf
	Select case iMonth
	    Case 3, 6, 9, 12
			'Get data for year, month from tbl_frmBalanceSheet
			Set Con = Server.CreateObject("ADODB.Connection")
			Con.Open "BBBSAforms", "sa","12sist12"
			strSQL = "SELECT * FROM tbl_frmBalanceSheet WHERE AgencyID=" & Session("AgencyIDN") & " AND Yr=" & Int(Request("Year")) & " AND Mth=" & Int(Request("Month"))

			Set rsBalanceSheet = Con.Execute(strSQL)

			If not rsBalanceSheet.EOF Then
				CashInvestments = rsBalanceSheet("CashInvestments")
				Receivables = rsBalanceSheet("Receivables")
				AllOtherAssets = rsBalanceSheet("AllOtherAssets")
				TotalAssets = rsBalanceSheet("TotalAssets")
				STLiabilities = rsBalanceSheet("ShortTermLiabilities")
				LTLiabilities = rsBalanceSheet("LongTermLiabilities")
				TotalLiabilities = rsBalanceSheet("TotalLiabilities")
				Surplus_NetAssets = rsBalanceSheet("Surplus_NetAssets")
				TotalLiabilitiesAndNetAssets = rsBalanceSheet("TotalLiabNetAssets")
			Else
				CashInvestments = 0
				Receivables = 0
				AllOtherAssets = 0
				TotalAssets = 0
				STLiabilities = 0
				LTLiabilities = 0
				TotalLiabilities = 0
				Surplus_NetAssets = 0
				TotalLiabilitiesAndNetAssets = 0
			End If
			
			'Display Data	
			'Section header
			'response.write "<script language='javascript'>NewWindow(this.href,'name','500','250','yes');return false;</script>"
			Response.Write("<tr><td valign='middle' align='center' class='formHeaderMedium' colspan=6>BALANCE SHEET<br>(as of " & MonthName(Request("m")) & " 31, " & Request("y") & "</td></tr>")
			
			'Data
			Response.Write("<tr><td valign='middle' class='formMain' colspan='5'>")'Help file code added 3/31/2009 saf''''''
			%> 
			<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_cash_investments" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
			<%''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			Response.Write("Cash Investments</td>")
			Response.Write("    <td valign='middle' class='formMain'>$ <input type='text'  class='formMain'  size=10 maxlength=10 value=" & CashInvestments & " name='CashInvestments'  onchange='checkForIntegerCommas(this.value);' onblur='addUpBalanceSheet();'></td></tr>")					

			Response.Write("<tr><td valign='middle' class='formMain' colspan='5'>")'Help file code added 3/31/2009 saf''''''
			%> 
			<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_raoca" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
			<%''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			Response.Write("Receivables and Other Current Assets</td>")
			Response.Write("    <td valign='middle' class='formMain'>$ <input type='text'  class='formMain'  size=10 maxlength=10 value=" & Receivables & " name='Receivables'  onchange='checkForIntegerCommas(this.value);' onblur='addUpBalanceSheet();'></td></tr>")					

			Response.Write("<tr><td valign='middle' class='formMain' colspan='5'>")'Help file code added 3/31/2009 saf''''''
			%> 
			<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_non_current_assets" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
			<%''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			Response.Write("Non Current Assets</td>")
			Response.Write("    <td valign='middle' class='formMain'>$ <input type='text'  class='formMain'  size=10 maxlength=10 value=" & AllOtherAssets & " name='AllOtherAssets'  onchange='checkForIntegerCommas(this.value);' onblur='addUpBalanceSheet();'></td></tr>")					

			Response.Write("<tr><td valign='middle' class='formMain' colspan='5'>")'Help file code added 3/31/2009 saf''''''
			%> 
			<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_total_assets" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
			<%''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			Response.Write("Total Assets</td>")
			Response.Write("    <td valign='middle' class='formMain' bgcolor='#c0c0c0'>$ <input type='text'  class='formMain'  size=10 maxlength=10 value=" & TotalAssets & " name='TotalAssets'  onchange='noChange();' readonly><span class='formSubHead'>&nbsp;&nbsp;&nbsp;<BR>Calculated</span></td></tr>")					

			Response.Write("<tr><td valign='middle' class='formMain' colspan='5'>")'Help file code added 3/31/2009 saf''''''
			%> 
			<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_current_libilities" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
			<%''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			Response.Write("Current Liabilities</td>")
			Response.Write("    <td valign='middle' class='formMain'>$ <input type='text'  class='formMain'  size=10 maxlength=10 value=" & STLiabilities & " name='STLiabilities'  onchange='checkForIntegerCommas(this.value);' onblur='addUpBalanceSheet();'></td></tr>")					

			Response.Write("<tr><td valign='middle' class='formMain' colspan='5'>")'Help file code added 3/31/2009 saf''''''
			%> 
			<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_non_current_libilities" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
			<%''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			Response.Write("Non Current Liabilities</td>")
			Response.Write("    <td valign='middle' class='formMain'>$ <input type='text'  class='formMain'  size=10 maxlength=10 value=" & LTLiabilities & " name='LTLiabilities'  onchange='checkForIntegerCommas(this.value);' onblur='addUpBalanceSheet();'></td></tr>")					

			Response.Write("<tr><td valign='middle' class='formMain' colspan='5'>")'Help file code added 3/31/2009 saf''''''
			%> 
			<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_total_libilities" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
			<%''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			Response.Write("Total Liabilities</td>")
			Response.Write("    <td valign='middle' class='formMain' bgcolor='#c0c0c0'>$ <input type='text'  class='formMain'  size=10 maxlength=10 value=" & TotalLiabilities & " name='TotalLiabilities'  onchange='noChange();' readonly><span class='formSubHead'>&nbsp;&nbsp;&nbsp;<BR>Calculated</span></td></tr>")					

			Response.Write("<tr><td valign='middle' class='formMain' colspan='5'>")'Help file code added 3/31/2009 saf''''''
			%> 
			<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_net_assets" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
			<%''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			Response.Write("Net Assets</td>")
			Response.Write("    <td valign='middle' class='formMain'>$ <input type='text'  class='formMain'  size=10 maxlength=10 value=" & Surplus_NetAssets & " name='Surplus_NetAssets'  onchange='checkForIntegerCommas(this.value);' onblur='addUpBalanceSheet();'></td></tr>")					

			Response.Write("<tr><td valign='middle' class='formMain' colspan='5'>")'Help file code added 3/31/2009 saf''''''
			%> 
			<a href="../helpfiles/surveyhelp.asp?HelpID=finance_performance_total_libilities_and_net_assets" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;
			<%''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			Response.Write("Total Liabilities and Net Assets</td>")
			Response.Write("    <td valign='middle' class='formMain' bgcolor='#c0c0c0'>$ <input type='text'  class='formMain'  size=10 maxlength=10 value=" & TotalLiabilitiesAndNetAssets & " name='TotalLiabilitiesAndNetAssets'  onchange='noChange();' readonly><span class='formSubHead'>&nbsp;&nbsp;&nbsp;<BR>Calculated</span></td></tr>")					

	    Case else
	End Select
'''''''''''''''''''''''''''''''''''''''''''''''''
%>

		<tr>
			<td colspan="6" class="formHeader">
				<input type="button" value="Save Form" class="formMainBold" onclick="validateForm(); return false;"  onclick="TotalNet();" >
			</td>
		</tr>

		<tr>
			<td colspan="6"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
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


<% End If %>
</body>
</html>


