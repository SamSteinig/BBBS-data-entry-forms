<% 
If Request("status") = "addNew" Then
	
	' Check for duplicate records
	
	Set DupCon = Server.CreateObject("ADODB.Connection")
	DupCon.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT Count(*) As NumberOfEntries FROM tbl_frmExpenses WHERE AgencyID = '" & Request("AgencyIDN") & "' and Year = " & Request("Year")	
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
		RST.Open "SELECT * FROM tbl_frmExpenses", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = Request("AgencyIDN")
		RST("Year") = Request("year")
		
			tt = Int(Request("frmExpensesSalariesWages")) _
			+ Int(Request("frmExpensesEmployeeBenefits")) _
			+ Int(Request("frmExpensesInsurance")) _			
			+ Int(Request("frmExpensesOther"))
			
		
		RST("SalariesWages") = FormatCurrency(Request("frmExpensesSalariesWages"))
		RST("EmployeeBenefits") = FormatCurrency(Request("frmExpensesEmployeeBenefits"))
		RST("Insurance") = FormatCurrency(Request("frmExpensesInsurance"))
		RST("Other") = FormatCurrency(Request("frmExpensesOther"))
		RST("Total") = FormatCurrency(tt)
		
			TotalAssets = Int(Request("frmExpensesCashInvestments")) _
			+ Int(Request("frmExpensesReceivables")) _
			+ Int(Request("frmExpensesAllOtherAssets")) 		
		
		RST("CashInvestments") = FormatCurrency(Request("frmExpensesCashInvestments"))
		RST("Receivables") = FormatCurrency(Request("frmExpensesReceivables"))
		RST("AllOtherAssets") = FormatCurrency(Request("frmExpensesAllOtherAssets"))
		RST("TotalAssets") = FormatCurrency(TotalAssets)
		
			TotalLiabNetAssets = Int(Request("frmExpensesLiabilities")) _
			+ Int(Request("frmExpensesSurplus_NetAssets"))
			
		RST("Liabilities") = FormatCurrency(Request("frmExpensesLiabilities"))		
		RST("Surplus_NetAssets") = FormatCurrency(Request("frmExpensesSurplus_NetAssets"))		
		RST("TotalLiabNetAssets") = FormatCurrency(TotalLiabNetAssets)
		
		
		RST("Administration") = Request("frmExpensesAdministration")
		RST("Program") = Request("frmExpensesProgram")
		RST("FundRaising") = Request("frmExpensesFundRaising")
		
			CategoryTotal = Int(Request("frmExpensesAdministration")) _
						  + Int(Request("frmExpensesProgram")) _
						  + Int(Request("frmExpensesFundRaising")) 

			RST("CategoryTotal") = CategoryTotal
			
		RST("TotalExpenseMentoring") = Request("frmExpensesTotalExpenseMentoring")
		RST("TotalExpenseNonMentoring") = Request("frmExpensesTotalExpenseNonMentoring")		
		
			ExpensesMentNonMentTotal = Int(Request("frmExpensesTotalExpenseMentoring"))_
									+ Int(Request("frmExpensesTotalExpenseNonMentoring"))
			RST("ExpensesMentNonMentTotal") = ExpensesMentNonMentTotal			
		
			
		RST("FTECommunity") = Request("frmExpensesFTECommunity")		
		RST("FTESchool") = Request("frmExpensesFTESchool")
		RST("FTESite") = Request("frmExpensesFTESite")		
		
			FTEProgramTotal = Int(Request("frmExpensesFTECommunity")) _
							+ Int(Request("frmExpensesFTESchool")) _
							+ Int(Request("frmExpensesFTESite")) 							
			RST("FTEProgramTotal") = FTEProgramTotal
		
		RST("FTECustomerRelations") = Request("frmExpensesFTECustomerRelations")		
		RST("FTEEnrollmentMatching") = Request("frmExpensesFTEEnrollmentMatching")		
		RST("FTEMatchSupport") = Request("frmExpensesFTEMatchSupport")	
		
			FTEFunctionTotal = Int(Request("frmExpensesFTECustomerRelations")) _
							 + Int(Request("frmExpensesFTEEnrollmentMatching")) _
							 + Int(Request("frmExpensesFTEMatchSupport"))
			RST("FTEFunctionTotal") = FTEFunctionTotal	
		
		
		RST("BenMedFullEmployee") = Request("frmExpensesBenMedFullEmployee")		
		RST("BenMedFullFamily") = Request("frmExpensesBenMedFullFamily")				
		RST("BenMedPartEmployee") = Request("frmExpensesBenMedPartEmployee")			
		RST("BenMedPartFamily") = Request("frmExpensesBenMedPartFamily")						
		RST("BenDentFullEmployee") = Request("frmExpensesBenDentFullEmployee")	
		RST("BenDentFullFamily") = Request("frmExpensesBenDentFullFamily")
		RST("BenDentPartEmployee") = Request("frmExpensesBenDentPartEmployee")
		RST("BenDentPartFamily") = Request("frmExpensesBenDentPartFamily")
		RST("DisInsShortTermFull") = Request("frmExpensesDisInsShortTermFull")						
		RST("DisInsShortTermPart") = Request("frmExpensesDisInsShortTermPart")	
		RST("DisInsLongTermFull") = Request("frmExpensesDisInsLongTermFull")									
		RST("DisInsLongTermPart") = Request("frmExpensesDisInsLongTermPart")		
		RST("EAPFull") = Request("frmExpensesEAPFull")
		RST("EAPPart") = Request("frmExpensesEAPPart")
		RST("FlexFull") = Request("frmExpensesFlexFull")
		RST("FlexPart") = Request("frmExpensesFlexPart")
		RST("HealthClubFull") = Request("frmExpensesHealthClubFull")
		RST("HealthClubPart") = Request("frmExpensesHealthClubPart")
		RST("LifeInsuranceFull") = Request("frmExpensesLifeInsuranceFull")
		RST("LifeInsurancePart") = Request("frmExpensesLifeInsurancePart")
		RST("TimeOffFull") = Request("frmExpensesTimeOffFull")
		RST("TimeOffPart") = Request("frmExpensesTimeOffPart")
		RST("TimeOffSickFull") = Request("frmExpensesTimeOffSickFull")		
		RST("TimeOffSickPart") = Request("frmExpensesTimeOffSickPart")
		RST("TimeOffVacFull") = Request("frmExpensesTimeOffVacFull")
		RST("TimeOffVacPart") = Request("frmExpensesTimeOffVacPart")
		RST("ProfDuesFull") = Request("frmExpensesProfDuesFull")		
		RST("ProfDuesFull") = Request("frmExpensesProfDuesFull")				
		RST("ProfDuesPart") = Request("frmExpensesProfDuesPart")				
		RST("RetirementFull") = Request("frmExpensesRetirementFull")			
		RST("RetirementPart") = Request("frmExpensesRetirementPart")					
		RST("TelecommFull") = Request("frmExpensesTelecommFull")							
		RST("TelecommPart") = Request("frmExpensesTelecommPart")			
		RST("TuitionFull") = Request("frmExpensesTuitionFull")			
		RST("TuitionPart") = Request("frmExpensesTuitionPart")				
		
		RST("CreateDate") = Now
		RST.Update
		RST.Close
		Set RST = Nothing
		form = "Expenses"
		modtype = "new"
		%>
		<!--#include file="../includes/modify_stamp.asp"-->
		<%	
		say = "thanks"
		Con.Close
		Set Con = Nothing
	Else
		say = "previouslyEdited"
		Con.Close
		Set Con = Nothing
	End If

	
ElseIf Request("status") = "editSave" Then


	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	Set RST = Server.CreateObject("ADODB.Recordset")
	RST.Open "SELECT * FROM tbl_frmExpenses WHERE AgencyID = '" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")), Con, 1, 3
		
			tt = Int(Request("frmExpensesSalariesWages")) _
			+ Int(Request("frmExpensesEmployeeBenefits")) _
			+ Int(Request("frmExpensesInsurance")) _			
			+ Int(Request("frmExpensesOther"))

		RST("SalariesWages") = FormatCurrency(Request("frmExpensesSalariesWages"))
		RST("EmployeeBenefits") = FormatCurrency(Request("frmExpensesEmployeeBenefits"))
		RST("Insurance") = FormatCurrency(Request("frmExpensesInsurance"))
		RST("Other") = FormatCurrency(Request("frmExpensesOther"))
		RST("Total") = FormatCurrency(tt)
		
			TotalAssets = Int(Request("frmExpensesCashInvestments")) _
			+ Int(Request("frmExpensesReceivables")) _
			+ Int(Request("frmExpensesAllOtherAssets")) _		
		
		RST("CashInvestments") = FormatCurrency(Request("frmExpensesCashInvestments"))
		RST("Receivables") = FormatCurrency(Request("frmExpensesReceivables"))
		RST("AllOtherAssets") = FormatCurrency(Request("frmExpensesAllOtherAssets"))
		RST("TotalAssets") = FormatCurrency(TotalAssets)
		
			TotalLiabNetAssets = Int(Request("frmExpensesLiabilities")) _
			+ Int(Request("frmExpensesSurplus_NetAssets"))
			
		RST("Liabilities") = FormatCurrency(Request("frmExpensesLiabilities"))		
		RST("Surplus_NetAssets") = FormatCurrency(Request("frmExpensesSurplus_NetAssets"))		
		RST("TotalLiabNetAssets") = FormatCurrency(TotalLiabNetAssets)		

		RST("Administration") = Request("frmExpensesAdministration")
		RST("Program") = Request("frmExpensesProgram")
		RST("FundRaising") = Request("frmExpensesFundRaising")

			CategoryTotal = Int(Request("frmExpensesAdministration")) _
						  + Int(Request("frmExpensesProgram")) _
						  + Int(Request("frmExpensesFundRaising")) 

			RST("CategoryTotal") = CategoryTotal

		RST("TotalExpenseMentoring") = Request("frmExpensesTotalExpenseMentoring")
		RST("TotalExpenseNonMentoring") = Request("frmExpensesTotalExpenseNonMentoring")		
		
			ExpensesMentNonMentTotal = Int(Request("frmExpensesTotalExpenseMentoring"))_
									+ Int(Request("frmExpensesTotalExpenseNonMentoring"))
			RST("ExpensesMentNonMentTotal") = ExpensesMentNonMentTotal									
									
									
		RST("FTECommunity") = Request("frmExpensesFTECommunity")		
		RST("FTESchool") = Request("frmExpensesFTESchool")
		RST("FTESite") = Request("frmExpensesFTESite")		
		
			FTEProgramTotal = Int(Request("frmExpensesFTECommunity")) _
							+ Int(Request("frmExpensesFTESchool")) _
							+ Int(Request("frmExpensesFTESite")) 							
			RST("FTEProgramTotal") = FTEProgramTotal
		
		
		RST("FTECustomerRelations") = Request("frmExpensesFTECustomerRelations")		
		RST("FTEEnrollmentMatching") = Request("frmExpensesFTEEnrollmentMatching")		
		RST("FTEMatchSupport") = Request("frmExpensesFTEMatchSupport")
		
			FTEFunctionTotal = Int(Request("frmExpensesFTECustomerRelations")) _
							 + Int(Request("frmExpensesFTEEnrollmentMatching")) _
							 + Int(Request("frmExpensesFTEMatchSupport"))
			RST("FTEFunctionTotal") = FTEFunctionTotal

		
		RST("BenMedFullEmployee") = Request("frmExpensesBenMedFullEmployee")		
		RST("BenMedFullFamily") = Request("frmExpensesBenMedFullFamily")				
		RST("BenMedPartEmployee") = Request("frmExpensesBenMedPartEmployee")			
		RST("BenMedPartFamily") = Request("frmExpensesBenMedPartFamily")						
		RST("BenDentFullEmployee") = Request("frmExpensesBenDentFullEmployee")	
		RST("BenDentFullFamily") = Request("frmExpensesBenDentFullFamily")		
		RST("BenDentPartEmployee") = Request("frmExpensesBenDentPartEmployee")				
		RST("BenDentPartFamily") = Request("frmExpensesBenDentPartFamily")						
		RST("DisInsShortTermFull") = Request("frmExpensesDisInsShortTermFull")						
		RST("DisInsShortTermPart") = Request("frmExpensesDisInsShortTermPart")	
		RST("DisInsLongTermFull") = Request("frmExpensesDisInsLongTermFull")										
		RST("DisInsLongTermPart") = Request("frmExpensesDisInsLongTermPart")		
		RST("EAPFull") = Request("frmExpensesEAPFull")
		RST("EAPPart") = Request("frmExpensesEAPPart")
		RST("FlexFull") = Request("frmExpensesFlexFull")
		RST("FlexPart") = Request("frmExpensesFlexPart")
		RST("HealthClubFull") = Request("frmExpensesHealthClubFull")
		RST("HealthClubPart") = Request("frmExpensesHealthClubPart")
		RST("LifeInsuranceFull") = Request("frmExpensesLifeInsuranceFull")
		RST("LifeInsurancePart") = Request("frmExpensesLifeInsurancePart")
		RST("TimeOffFull") = Request("frmExpensesTimeOffFull")
		RST("TimeOffPart") = Request("frmExpensesTimeOffPart")
		RST("TimeOffSickFull") = Request("frmExpensesTimeOffSickFull")		
		RST("TimeOffSickPart") = Request("frmExpensesTimeOffSickPart")
		RST("TimeOffVacFull") = Request("frmExpensesTimeOffVacFull")
		RST("TimeOffVacPart") = Request("frmExpensesTimeOffVacPart")
		RST("ProfDuesFull") = Request("frmExpensesProfDuesFull")		
		RST("ProfDuesFull") = Request("frmExpensesProfDuesFull")				
		RST("ProfDuesPart") = Request("frmExpensesProfDuesPart")				
		RST("RetirementFull") = Request("frmExpensesRetirementFull")			
		RST("RetirementPart") = Request("frmExpensesRetirementPart")					
		RST("TelecommFull") = Request("frmExpensesTelecommFull")							
		RST("TelecommPart") = Request("frmExpensesTelecommPart")			
		RST("TuitionFull") = Request("frmExpensesTuitionFull")			
		RST("TuitionPart") = Request("frmExpensesTuitionPart")
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
	<title>Expenses</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
<script language="JavaScript">
<!--

function addEmUp() 
{
	var box1 = Number(document.frmExpenses.frmExpensesSalariesWages.value)
	var box2 = Number(document.frmExpenses.frmExpensesEmployeeBenefits.value)
	var box3 = Number(document.frmExpenses.frmExpensesInsurance.value)	
	var box4 = Number(document.frmExpenses.frmExpensesOther.value)
	var boxtotal = box1 + box2 + box3 + box4
	document.frmExpenses.frmExpensesTotal.value = boxtotal
}


function addUpAssets()
{
	var box1 = Number(document.frmExpenses.frmExpensesCashInvestments.value)
	var box2 = Number(document.frmExpenses.frmExpensesReceivables.value)
	var box3 = Number(document.frmExpenses.frmExpensesAllOtherAssets.value)	

	var boxtotal = box1 + box2 + box3
	document.frmExpenses.frmExpensesTotalAssets.value = boxtotal
	
}

function addUpLiabilities()
{
	var box1 = Number(document.frmExpenses.frmExpensesLiabilities.value)
	var box2 = Number(document.frmExpenses.frmExpensesSurplus_NetAssets.value)

	var boxtotal = box1 + box2
	document.frmExpenses.frmExpensesTotalLiabNetAssets.value = boxtotal
	
}




function addUpCategory()
{
	var box1 = Number(document.frmExpenses.frmExpensesAdministration.value)
	var box2 = Number(document.frmExpenses.frmExpensesProgram.value)
	var box3 = Number(document.frmExpenses.frmExpensesFundRaising.value)	

	var boxtotal = box1 + box2 + box3
	document.frmExpenses.frmExpensesCategoryTotal.value = boxtotal
	
}

function addUpExpenseMentNonMent()
{
	var box1 = Number(document.frmExpenses.frmExpensesTotalExpenseMentoring.value)
	var box2 = Number(document.frmExpenses.frmExpensesTotalExpenseNonMentoring.value)

	var boxtotal = box1 + box2
	document.frmExpenses.frmExpensesExpensesMentNonMentTotal.value = boxtotal
	
}

function addUpFTEProgram()
{
	var box1 = Number(document.frmExpenses.frmExpensesFTECommunity.value)
	var box2 = Number(document.frmExpenses.frmExpensesFTESchool.value)
	var box3 = Number(document.frmExpenses.frmExpensesFTESite.value)	

	var boxtotal = box1 + box2 + box3
	document.frmExpenses.frmExpensesFTEProgramTotal.value = boxtotal
	
}

function addUpFTEFunction()
{
	var box1 = Number(document.frmExpenses.frmExpensesFTECustomerRelations.value)
	var box2 = Number(document.frmExpenses.frmExpensesFTEEnrollmentMatching.value)
	var box3 = Number(document.frmExpenses.frmExpensesFTEMatchSupport.value)	

	var boxtotal = box1 + box2 + box3
	document.frmExpenses.frmExpensesFTEFunctionTotal.value = boxtotal
	
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


var myRegularExpression1 = /^[0-9]+(,[0-9]{3})*$/;  // Checks for integer with or without commas
var myRegularExpression2 = /^\d*\.?\d*$/;  	// contains an int, double or float
var myRegularExpression3 = 	/^\d*$/;  // Checks for integer 
	
function submitFormValidate(form)
{

//	for(i=0; i < form.elements.length; i++)
//	{
//		if(form.elements[i].value == "")
//		{	
//			form.elements[i].focus();
//			alert("You must enter a value for all fields");
//			return false;
//		}
//	}


	//Check Operating Expenses Total
	var TotalOperatingExpenses = Number(document.frmExpenses.frmExpensesTotal.value)
	if (TotalOperatingExpenses.valueOf() == 0)
	{
		alert("Total Operating Expenses must be greater than zero.");
		return false;
	}
	
	// Add Up Totals for Category
	var box1 = Number(document.frmExpenses.frmExpensesAdministration.value)
	var box2 = Number(document.frmExpenses.frmExpensesProgram.value)
	var box3 = Number(document.frmExpenses.frmExpensesFundRaising.value)
	var CategoryTotal = box1 + box2 + box3
	document.frmExpenses.frmExpensesCategoryTotal.value = CategoryTotal	

	
	
	var value1 = new Number(form.frmExpensesAdministration.value);
	var value2 = new Number(form.frmExpensesProgram.value);
	var value3 = new Number(form.frmExpensesFundRaising.value);
	var valueTotal = new Number(value1 + value2 + value3);
	var OneHundred = new Number(100);
	
	if (CategoryTotal.valueOf() != 100)
	{
		alert("Total for Expense Breakdown by CATEGORY should be 100%.  Current total is " +CategoryTotal+"%");
		return false;
	}	
	
	// Add Up Totals for Mentoring / Non-Mentoring	
	
	var box1 = Number(document.frmExpenses.frmExpensesTotalExpenseMentoring.value)
	var box2 = Number(document.frmExpenses.frmExpensesTotalExpenseNonMentoring.value)
	var ExpensesMentNonMentTotal = box1 + box2
	document.frmExpenses.frmExpensesExpensesMentNonMentTotal.value = ExpensesMentNonMentTotal	

	
	var value1 = new Number(form.frmExpensesTotalExpenseMentoring.value);
	var value2 = new Number(form.frmExpensesTotalExpenseNonMentoring.value);
	var MentNonMentTotal = new Number(value1 + value2);
	var OneHundred = new Number(100);
	
	if (ExpensesMentNonMentTotal.valueOf() != 100)
	{
		alert("Total for Expense Breakdown by Mentoring / Non-Mentoring should be 100%.  Current total is " +ExpensesMentNonMentTotal+"%");
		return false;
	}	
	
	// Add Up Totals for FTE Program	
	
	var box1 = Number(document.frmExpenses.frmExpensesFTECommunity.value)
	var box2 = Number(document.frmExpenses.frmExpensesFTESchool.value)
	var box3 = Number(document.frmExpenses.frmExpensesFTESite.value)	
	var FTEProgramTotal = box1 + box2 + box3
	document.frmExpenses.frmExpensesFTEProgramTotal.value = FTEProgramTotal	

	
	var value1 = new Number(form.frmExpensesFTECommunity.value);
	var value2 = new Number(form.frmExpensesFTESchool.value);
	var value3 = new Number(form.frmExpensesFTESite.value);
	var FTEProgramTotal = new Number(value1 + value2 + value3);
	var OneHundred = new Number(100);
	
	if (FTEProgramTotal.valueOf() != 100)
	{
		alert("Total percentage of Mentoring Program FTEs should be 100%.  Current total is " + FTEProgramTotal +"%");
		return false;
	}
	
	// Add Up Totals for FTE Function	
	
	var box1 = Number(document.frmExpenses.frmExpensesFTECustomerRelations.value)
	var box2 = Number(document.frmExpenses.frmExpensesFTEEnrollmentMatching.value)
	var box3 = Number(document.frmExpenses.frmExpensesFTEMatchSupport.value)	
	var FTEFunctionTotal = box1 + box2 + box3
	document.frmExpenses.frmExpensesFTEFunctionTotal.value = FTEFunctionTotal	

	
	var value1 = new Number(form.frmExpensesFTECustomerRelations.value);
	var value2 = new Number(form.frmExpensesFTEEnrollmentMatching.value);
	var value3 = new Number(form.frmExpensesFTEMatchSupport.value);
	var FTEFunctionTotal = new Number(value1 + value2 + value3);
	var OneHundred = new Number(100);
	
	if (FTEFunctionTotal.valueOf() != 100 && FTEFunctionTotal.valueOf() != 0)
	{
		alert("Total percentage of Mentoring Program FTEs should be 100%.  Current total is " + FTEFunctionTotal +"%");
		return false;
	}
	
	
		
	else if(!(myRegularExpression1.test(form.frmExpensesSalariesWages.value)))
	{
		form.frmExpensesSalariesWages.focus();
		alert((form.frmExpensesSalariesWages.value) + " is invalid.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmExpensesEmployeeBenefits.value)))	
	{
		form.frmExpensesEmployeeBenefits.focus();
		alert((form.frmExpensesEmployeeBenefits.value) + " is invalid.");
		return false;
	}


	else if(!(myRegularExpression1.test(form.frmExpensesInsurance.value)))	
	{
		form.frmExpensesInsurance.focus();
		alert((form.frmExpensesInsurance.value) + " is invalid.");
		return false;
	}		
	

	else if(!(myRegularExpression1.test(form.frmExpensesOther.value)))	
	{
		form.frmExpensesOther.focus();
		alert((form.frmExpensesOther.value) + " is invalid.");
		return false;
	}
	
	else if(!(myRegularExpression1.test(form.frmExpensesCashInvestments.value)))	
	{
		form.frmExpensesCashInvestments.focus();
		alert((form.frmExpensesCashInvestments.value) + " is invalid.");
		return false;
	}	
	
	else if(!(myRegularExpression1.test(form.frmExpensesReceivables.value)))	
	{
		form.frmExpensesReceivables.focus();
		alert((form.frmExpensesReceivables.value) + " is invalid.");
		return false;
	}		
	
	else if(!(myRegularExpression1.test(form.frmExpensesAllOtherAssets.value)))	
	{
		form.frmExpensesAllOtherAssets.focus();
		alert((form.frmExpensesAllOtherAssets.value) + " is invalid.");
		return false;
	}		
	
	else if(!(myRegularExpression1.test(form.frmExpensesLiabilities.value)))	
	{
		form.frmExpensesLiabilities.focus();
		alert((form.frmExpensesLiabilities.value) + " is invalid.");
		return false;
	}		
	
	else if(!(myRegularExpression1.test(form.frmExpensesSurplus_NetAssets.value)))	
	{
		form.frmExpensesSurplus_NetAssets.focus();
		alert((form.frmExpensesSurplus_NetAssets.value) + " is invalid.");
		return false;
	}		
	
	
	
	else if(!myRegularExpression3.test(form.frmExpensesAdministration.value))	
	{
		form.frmExpensesAdministration.focus();
		alert((form.frmExpensesAdministration.value) + " is invalid.");
		return false;
	}
	else if(!myRegularExpression3.test(form.frmExpensesProgram.value))	
	{
		form.frmExpensesProgram.focus();
		alert((form.frmExpensesProgram.value) + " is invalid.");
		return false;
	}
	else if(!myRegularExpression3.test(form.frmExpensesFundRaising.value))	
	{
		form.frmExpensesFundRaising.focus();
		alert((form.frmExpensesFundRaising.value) + " is invalid.");
		return false;
	}
	
						
	
	else if(!(myRegularExpression1.test(form.frmExpensesBenMedFullEmployee.value)))	
	{
		form.frmExpensesBenMedFullEmployee.focus();
		alert((form.frmExpensesBenMedFullEmployee.value) + " is invalid.");
		return false;
	}							
	
	else if(!(myRegularExpression1.test(form.frmExpensesBenMedFullFamily.value)))	
	{
		form.frmExpensesBenMedFullFamily.focus();
		alert((form.frmExpensesBenMedFullFamily.value) + " is invalid.");
		return false;
	}							
	
	else if(!(myRegularExpression1.test(form.frmExpensesBenMedPartEmployee.value)))	
	{
		form.frmExpensesBenMedPartEmployee.focus();
		alert((form.frmExpensesBenMedPartEmployee.value) + " is invalid.");
		return false;
	}							
	
	else if(!(myRegularExpression1.test(form.frmExpensesBenMedPartFamily.value)))	
	{
		form.frmExpensesBenMedPartFamily.focus();
		alert((form.frmExpensesBenMedPartFamily.value) + " is invalid.");
		return false;
	}							
	
	else if(!(myRegularExpression1.test(form.frmExpensesBenDentFullEmployee.value)))	
	{
		form.frmExpensesBenDentFullEmployee.focus();
		alert((form.frmExpensesBenDentFullEmployee.value) + " is invalid.");
		return false;
	}									
	
	else if(!(myRegularExpression1.test(form.frmExpensesBenDentFullFamily.value)))	
	{
		form.frmExpensesBenDentFullFamily.focus();
		alert((form.frmExpensesBenDentFullFamily.value) + " is invalid.");
		return false;
	}									
	
	else if(!(myRegularExpression1.test(form.frmExpensesBenDentPartEmployee.value)))	
	{
		form.frmExpensesBenDentPartEmployee.focus();
		alert((form.frmExpensesBenDentPartEmployee.value) + " is invalid.");
		return false;
	}											
	
	else if(!(myRegularExpression1.test(form.frmExpensesBenDentPartFamily.value)))	
	{
		form.frmExpensesBenDentPartFamily.focus();
		alert((form.frmExpensesBenDentPartFamily.value) + " is invalid.");
		return false;
	}											
	
	
	
	else if((form.frmExpensesAdministration.value > 100) || (form.frmExpensesAdministration.value < 0))	
	{
		form.frmExpensesAdministration.focus();
		alert((form.frmExpensesAdministration.value) + " is invalid. Please enter a number between 0 and 100.");
		return false;
	}
	else if((form.frmExpensesProgram.value > 100) || (form.frmExpensesProgram.value < 0))	
	{
		form.frmExpensesProgram.focus();
		alert((form.frmExpensesProgram.value) + " is invalid. Please enter a number between 0 and 100.");
		return false;
	}
	else if((form.frmExpensesFundRaising.value > 100) || (form.frmExpensesFundRaising.value < 0))	
	{
		form.frmExpensesFundRaising.focus();
		alert((form.frmExpensesFundRaising.value) + " is invalid. Please enter a number between 0 and 100.");
		return false;
	}
	

		
	
	else if((form.frmExpensesBenMedFullEmployee.value > 100) || (form.frmExpensesBenMedFullEmployee.value < 0))	
	{
		form.frmExpensesBenMedFullEmployee.focus();
		alert((form.frmExpensesBenMedFullEmployee.value) + " is invalid. Please enter a number between 0 and 100.");
		return false;
	}				
	
	else if((form.frmExpensesBenMedFullFamily.value > 100) || (form.frmExpensesBenMedFullFamily.value < 0))	
	{
		form.frmExpensesBenMedFullFamily.focus();
		alert((form.frmExpensesBenMedFullFamily.value) + " is invalid. Please enter a number between 0 and 100.");
		return false;
	}				
	
	else if((form.frmExpensesBenMedPartEmployee.value > 100) || (form.frmExpensesBenMedPartEmployee.value < 0))	
	{
		form.frmExpensesBenMedPartEmployee.focus();
		alert((form.frmExpensesBenMedPartEmployee.value) + " is invalid. Please enter a number between 0 and 100.");
		return false;
	}					
	
	else if((form.frmExpensesBenMedPartFamily.value > 100) || (form.frmExpensesBenMedPartFamily.value < 0))	
	{
		form.frmExpensesBenMedPartFamily.focus();
		alert((form.frmExpensesBenMedPartFamily.value) + " is invalid. Please enter a number between 0 and 100.");
		return false;
	}		
	
	else if((form.frmExpensesBenDentFullEmployee.value > 100) || (form.frmExpensesBenDentFullEmployee.value < 0))	
	{
		form.frmExpensesBenDentFullEmployee.focus();
		alert((form.frmExpensesBenDentFullEmployee.value) + " is invalid. Please enter a number between 0 and 100.");
		return false;
	}		
	
	else if((form.frmExpensesBenDentFullFamily.value > 100) || (form.frmExpensesBenDentFullFamily.value < 0))	
	{
		form.frmExpensesBenDentFullFamily.focus();
		alert((form.frmExpensesBenDentFullFamily.value) + " is invalid. Please enter a number between 0 and 100.");
		return false;
	}		
	
	else if((form.frmExpensesBenDentPartEmployee.value > 100) || (form.frmExpensesBenDentPartEmployee.value < 0))	
	{
		form.frmExpensesBenDentPartEmployee.focus();
		alert((form.frmExpensesBenDentPartEmployee.value) + " is invalid. Please enter a number between 0 and 100.");
		return false;
	}								
	
	else if((form.frmExpensesBenDentPartFamily.value > 100) || (form.frmExpensesBenDentPartFamily.value < 0))	
	{
		form.frmExpensesBenDentPartFamily.focus();
		alert((form.frmExpensesBenDentPartFamily.value) + " is invalid. Please enter a number between 0 and 100.");
		return false;
	}									
	
	
	
	
	
	else if(valueTotal.valueOf() != OneHundred.valueOf())
	{
		form.frmExpensesAdministration.focus();
		alert("Total for EXPENSE BREAKDOWN BY CATEGORY should equal 100%. Current total is " + valueTotal + "%");
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

<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_slinky.jpg" alt="" width="220" height="477" border="0"></td>
<td width="100%" valign="top">
<br>

<% If say = "thanks" Then %>
<font class="formMain"><br><br>
<strong>Thank you!</strong> Your information has been saved in the BBBS database.<br>
To choose another form, please select the form type from the choices above.
</font>
<br>
<!--#include file="../includes/contact_info.inc"-->
<br>




<% ElseIf say <> "thanks" Then  %>
<table border="1" cellspacing="0" cellpadding="2" width = "600" bordercolordark="#003063">
<form name="frmExpenses" action="expenses_edit.asp" method="post" onsubmit="return submitFormValidate(this)">
<!--#include file="../includes/form_stamp.asp"-->

<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmExpenses WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
	Set GetExpenses = Con.Execute(query)
 %>
<input type="hidden" name="status" value="editSave">
<% Else %>
<input type="hidden" name="status" value="addNew">
<%
End If
 
If say = "previouslyEdited" Then
%>
<p class="formMain">We're sorry, but this form was previously completed. To make changes please <a href="yearly.asp">reselect</a> the 
appropriate form and year and update the existing information.</p>
<%
Response.End
End If 
%>
 
 
 

	<tr>
		<td colspan="6" align="center" class="formSubhead">BBBS - <%= y %> Annual Agency Information (AAI)</td>
	</tr>
	<tr>
		<td colspan="6" class="formHeader">FINANCES</td>
	</tr>
	<tr>
		<td colspan="6" class="formMain"><font color="#ff0000"><div align="center"><strong>Please Note: </strong>After entering your information, you <strong>must</strong> click on the "Save Form" button at the bottom of the form and wait for the "Thank You" screen or your changes will be lost.</div></font></td>
	</tr>
	<tr>
		<td colspan="6" class="formMainCentered"><strong>EXPENSES</strong></td>
	</tr>
	<tr>
		<td colspan="6" class="formSubhead" align="center">Do Not Include Direct Expenses From Fundraising - No Cents</td>
	</tr>
	<tr>
		<td class="formMain" width="24%">Salaries and Wages:</td>
		<td class="formMainRightJ" width="24%">$&nbsp;<input type="text" class="formMain" size="10" maxlength="25" value="<% If say = "edit" Then %><%= GetExpenses("SalariesWages") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesSalariesWages" onFocus="addEmUp();" onchange="checkForInteger(this.value);"></td>
		<td colspan="2" width="4%">&nbsp;</td>		
		<td class="formMain" width="24%">Employee Benefits:</td>
		<td class="formMainRightJ" width="24%">$&nbsp;<input type="text" class="formMain" size="10" maxlength="25" value="<% If say = "edit" Then %><%= GetExpenses("EmployeeBenefits") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesEmployeeBenefits" onFocus="addEmUp();" onchange="checkForInteger(this.value);"></td>		
	</tr>
	<tr>
		<td class="formMain" width="24%">Liability Insurance:</td>				
		<td class="formMainRightJ" width="24%">$&nbsp;<input type="text" class="formMain" size="10" maxlength="25" value="<% If say = "edit" Then %><%= GetExpenses("Insurance") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesInsurance" onFocus="addEmUp();" onchange="checkForInteger(this.value);"></td>				
		<td colspan="2" width="4%">&nbsp;</td>		
		<td class="formMain" width="24%">Other:</td>								
		<td class="formMainRightJ" width="24%">$&nbsp;<input type="text" class="formMain" size="10" maxlength="25" value="<% If say = "edit" Then %><%= GetExpenses("Other") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesOther" onFocus="addEmUp();" onchange="checkForInteger(this.value);"></td>										
	</tr>

	<tr>
		<td class="formMain" width="24%">&nbsp;</td>
		<td class="formMainRightJ" width="24%">$&nbsp;</td>
		<td colspan="2" width="4%">&nbsp;</td>		
		<td class="formMain" width="24%" bgcolor="#c0c0c0"><strong>Total Operating Expenses</strong></td>
		<td class="formMainRightJ" width="24%" bgcolor="#c0c0c0">$&nbsp;<input type="text" class="formMain" size="10" maxlength="25" value="<% If say = "edit" Then %><%= GetExpenses("Total") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesTotal" onFocus="addEmUp();" onchange="addEmUp();" readonly><br><span class="formSubHead">calculated by system</span></td>		
	</tr>
	<tr>
	
	
		<tr>
			<td colspan="6" class="formMainCentered"><strong>BALANCE SHEET<br>(as of December 31, <%=Y%>)</strong></td>
		</tr>				
		
	
		<TR>
			<TD class="formMain" colspan="6"><strong><div align="center">Assets</div></strong></TD>
		</TR>
		<tr>
			<td class="formMain" width="80%" colspan="5">Cash/Investments:</td>
			<input type="text" class="formMain" size="10" maxlength="25" value="<% If say = "edit" Then %><%= GetExpenses("CashInvestments") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesCashInvestments" onFocus="AddUpAssets();" onchange="checkForInteger(this.value);"></td>			
		</tr>
		<tr>
			<td class="formMain" width="80%" colspan="5">Receivables:</td>
			<input type="text" class="formMain" size="10" maxlength="25" value="<% If say = "edit" Then %><%= GetExpenses("Receivables") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesReceivables" onFocus="AddUpAssets();" onchange="checkForInteger(this.value);"></td>						
		</tr>		
		<tr>
			<td class="formMain" width="80%" colspan="5">All Other Assets:</td>
			<input type="text" class="formMain" size="10" maxlength="25" value="<% If say = "edit" Then %><%= GetExpenses("AllOtherAssets") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesAllOtherAssets" onFocus="AddUpAssets();" onchange="checkForInteger(this.value);"></td>									
		</tr>
		<tr>
			<td class="formMain" width="80%" colspan="5"  bgcolor="#c0c0c0"><strong>TOTAL ASSETS:</strong></td>
			<td class="formMainRightJ" width="24%" bgcolor="#c0c0c0">$&nbsp;<input type="text" class="formMain" size="10" maxlength="25" value="<% If say = "edit" Then %><%= GetExpenses("TotalAssets") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesTotalAssets" onFocus="addEmUp();" onchange="addUpAssets();" readonly><br><span class="formSubHead">calculated by system</span></td>					
		</tr>	
		<TR>
			<TD class="formMain" colspan="6"><strong><div align="center">Liabilities and Net Assets</div></strong></TD>
		</TR>
		<tr>
			<td class="formMain" width="80%" colspan="5">Total Liabilities:</td>
			<input type="text" class="formMain" size="10" maxlength="25" value="<% If say = "edit" Then %><%= GetExpenses("Liabilities") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesLiabilities" onFocus="AddUpLiabilities();" onchange="checkForInteger(this.value);"></td>						
		</tr>		
		<tr>
			<td class="formMain" width="80%" colspan="5">Surplus/Net Assets:</td>
			<input type="text" class="formMain" size="10" maxlength="25" value="<% If say = "edit" Then %><%= GetExpenses("Surplus_NetAssets") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesSurplus_NetAssets" onFocus="AddUpLiabilities();" onchange="checkForInteger(this.value);"></td>			
		</tr>		
		<tr>
			<td class="formMain" width="80%" colspan="5" bgcolor="#c0c0c0"><strong>TOTAL LIABILITIES AND NET ASSETS:</strong></td>
			<td class="formMainRightJ" width="24%" bgcolor="#c0c0c0">$&nbsp;<input type="text" class="formMain" size="10" maxlength="25" value="<% If say = "edit" Then %><%= GetExpenses("TotalLiabNetAssets") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesTotalLiabNetAssets" onFocus="addUpLiabilities();" onchange="addUpLiabilities();" readonly><br><span class="formSubHead">calculated by system</span></td>								
			<td class="formMain" width="20%" align="right" colspan="2" bgcolor="#c0c0c0"><%=FormatCurrency(GetExpenses("TotalLiabNetAssets"))%></td>							
		</tr>				
	
	
	
	
	
	
	
	
	
	
		<td colspan="6">
			<table width="100%" border="1" bordercolordark="#003063" cellspacing="0" cellpadding="2">
			
			<tr>
				<td class="formHeaderSmall" colspan="6">EXPENSE BREAKDOWN BY CATEGORY<br>(Enter whole numbers only)</td>
			</tr>
			<tr>
				<td class="formMain" colspan="4">Administration:<br>
				<font class="formSubhead"><i>If any administration expenses are related to program or fundraising then include those expenses in program or fundraising when calculating percentages.</i></font></td>
				<td class="formMain" valign="top" align="right" colspan="2"><input type="text" size="4" maxlength="18" value="<% If say = "edit" Then %><%= GetExpenses("Administration") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesAdministration" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value); addUpCategory(); " onFocus="addUpCategory();">&nbsp;&#37;&nbsp;</td>
			</tr>
			<tr>
				<td class="formMain" colspan="4">Program:<br>
				<font class="formSubhead"><i>Including time spent supervising program staff</i></font></td>
				<td class="formMain" valign="top" align="right" colspan="2"><input type="text" size="4" maxlength="18" value="<% If say = "edit" Then %><%= GetExpenses("Program") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesProgram" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value); addUpCategory();" onFocus="addUpCategory();">&nbsp;&#37;&nbsp;</td>
			</tr>
			<tr>
				<td class="formMain" colspan="4">Fundraising:</td>
				<td class="formMain" valign="top" align="right" colspan="2"><input type="text" size="4" maxlength="18" value="<% If say = "edit" Then %><%= GetExpenses("FundRaising") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesFundRaising" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value); addUpCategory();" onFocus="addUpCategory();">&nbsp;&#37;&nbsp;</td>
			</tr>		
			<tr>
				<td class="formMain" colspan="4" bgcolor="#c0c0c0"><strong>TOTAL</strong></td>
				<td class="formMain" valign="top" align="right" bgcolor="#c0c0c0"><input type="text" size="4" maxlength="4" value="<% If say = "edit" Then %><%= GetExpenses("CategoryTotal") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesCategoryTotal" onFocus="addUpCategory();" onchange="addUpCategory();" readonly>&nbsp;&#37;&nbsp;<br><span class="formSubHead">calculated by system</span></td>
			</tr>					
			
			<tr>
				<td class="formHeaderSmall" colspan="6" align="center">EXPENSE BREAKDOWN BY FUNCTION<br>(enter whole numbers only)</td>
			</tr>
			
			<tr>
				<td class="formMain" colspan="5" align="left"><em>How much total expense goes toward (must equal 100%)</em></td>
			</tr>			
			<tr>
				<td class="formMain" width="80%" colspan="4">Mentoring</td>
				<td class="formMain" align="right" width="20%"><input type="text" size="4" maxlength="18" value="<% If say = "edit" Then %><%= GetExpenses("TotalExpenseMentoring") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesTotalExpenseMentoring" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value); addUpExpenseMentNonMent(); " onFocus="addUpExpenseMentNonMent();">&nbsp;&#37;&nbsp;</td>
			</tr>
			<tr>
				<td class="formMain" width="80%" colspan="4">Non-Mentoring:</td>
				<td class="formMain" align="right" width="20%"><input type="text" size="4" maxlength="18" value="<% If say = "edit" Then %><%= GetExpenses("TotalExpenseNonMentoring") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesTotalExpenseNonMentoring" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value); addUpExpenseMentNonMent(); " onFocus="addUpExpenseMentNonMent();">&nbsp;&#37;&nbsp;</td>	
			</tr>
			<tr>
				<td class="formMain" colspan="4" bgcolor="#c0c0c0"><strong>TOTAL</strong></td>
				<td class="formMain" valign="top" align="right" bgcolor="#c0c0c0"><input type="text" size="4" maxlength="4" value="<% If say = "edit" Then %><%= GetExpenses("ExpensesMentNonMentTotal") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesExpensesMentNonMentTotal" onFocus="addUpExpenseMentNonMent();" onchange="addUpExpenseMentNonMent();" readonly>&nbsp;&#37;&nbsp;<br><span class="formSubHead">calculated by system</span></td>
			</tr>				
			
			<tr>
				<td class="formMain" colspan="6" align="left"><em>Estimate the percent of Mentoring Program *FTEs (Full Time Employees) that go toward the following PROGRAMS (must equal 100%)</em></td>
			</tr>			
			<tr>
				<td class="formMain" width="80%" colspan="4">Community:</td>
				<td class="formMain" width="20%" align="right"><input type="text" size="4" maxlength="18" value="<% If say = "edit" Then %><%= GetExpenses("FTECommunity") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesFTECommunity" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value); addUpFTEProgram(); " onFocus="addUpFTEProgram();">&nbsp;&#37;&nbsp;</td>	
			</tr>
			<tr>
				<td class="formMain" width="80%" colspan="4">School-Based:</td>
				<td class="formMain" width="20%" align="right"><input type="text" size="4" maxlength="18" value="<% If say = "edit" Then %><%= GetExpenses("FTESchool") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesFTESchool" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value); addUpFTEProgram(); " onFocus="addUpFTEProgram();">&nbsp;&#37;&nbsp;</td>	
			</tr>
			<tr>
				<td class="formMain" width="80%" colspan="4">Site-Based:</td>
				<td class="formMain" width="20%" align="right"><input type="text" size="4" maxlength="18" value="<% If say = "edit" Then %><%= GetExpenses("FTESite") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesFTESite" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value); addUpFTEProgram(); " onFocus="addUpFTEProgram();">&nbsp;&#37;&nbsp;</td>	
			</tr>
			<tr>
				<td class="formMain" colspan="4" bgcolor="#c0c0c0"><strong>TOTAL</strong></td>
				<td class="formMain" valign="top" align="right" bgcolor="#c0c0c0"><input type="text" size="4" maxlength="4" value="<% If say = "edit" Then %><%= GetExpenses("FTEProgramTotal") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesFTEProgramTotal" onFocus="addUpFTEProgram();" onchange="addUpFTEProgram();" readonly>&nbsp;&#37;&nbsp;<br><span class="formSubHead">calculated by system</span></td>
			</tr>							
			
													


			<tr>
				<td class="formMain" colspan="5" align="left"><em>Estimate the percent of Mentoring Program *FTEs (Full Time Employees) that go toward the following FUNCTIONS (must equal 100%)</em><br><i><span class="formSubHead">NOTE: IF YOU ARE NOT IMPLEMENTING SDM, PLEASE LEAVE ZEROES IN THESE FIELDS.</span></i></td>
			</tr>			
			<tr>
				<td class="formMain" width="80%" colspan="4">Customer Relations:</td>
				<td class="formMain" width="20%" align="right"><input type="text" size="4" maxlength="18" value="<% If say = "edit" Then %><%= GetExpenses("FTECustomerRelations") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesFTECustomerRelations" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value); addUpFTEFunction(); " onFocus="addUpFTEFunction();">&nbsp;&#37;&nbsp;</td>					
			</tr>
			<tr>
				<td class="formMain" width="80%" colspan="4">Enrollment / Matching:</td>
				<td class="formMain" width="20%" align="right"><input type="text" size="4" maxlength="18" value="<% If say = "edit" Then %><%= GetExpenses("FTEEnrollmentMatching") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesFTEEnrollmentMatching" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value); addUpFTEFunction(); " onFocus="addUpFTEFunction();">&nbsp;&#37;&nbsp;</td>									
			</tr>
			<tr>
				<td class="formMain" width="80%" colspan="4">Match Support:</td>
				<td class="formMain" width="20%" align="right"><input type="text" size="4" maxlength="18" value="<% If say = "edit" Then %><%= GetExpenses("FTEMatchSupport") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesFTEMatchSupport" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value); addUpFTEFunction(); " onFocus="addUpFTEFunction();">&nbsp;&#37;&nbsp;</td>													
			</tr>
			<tr>
				<td class="formMain" colspan="4" bgcolor="#c0c0c0"><strong>TOTAL</strong></td>
				<td class="formMain" valign="top" align="right" bgcolor="#c0c0c0"><input type="text" size="4" maxlength="4" value="<% If say = "edit" Then %><%= GetExpenses("FTEFunctionTotal") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesFTEFunctionTotal" onFocus="addUpFTEFunction();" onchange="addUpFTEFunction();" readonly>&nbsp;&#37;&nbsp;<br><span class="formSubHead">calculated by system</span></td>
			</tr>			
			
			<tr>
				<td colspan="6" class="formHeaderSmall">BENEFITS - MEDICAL<BR>% paid by BBBS for Employee and Employee's Family</td>
			</tr>
			
			<tr>
				<td class="formMain">&nbsp;</td>
				<td class="formMain" align="center" colspan="2">Full Time</td>
				<td class="formMain" align="center" colspan="2">Part Time</td>				
			</tr>
			
			<tr>
				<td class="formMain" align="center">Medical</td>

				<td class="formMain" valign="top" align="center">
				For Employee<br>
				<input type="text" size="3" maxlength="3" value="<% If say = "edit" Then %><%= GetExpenses("BenMedFullEmployee") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesBenMedFullEmployee" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">&nbsp;&#37;&nbsp;
				</td>
				
				<td class="formMain" valign="top" align="center">
				For Family<br>
				<input type="text" size="3" maxlength="3" value="<% If say = "edit" Then %><%= GetExpenses("BenMedFullFamily") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesBenMedFullFamily" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">&nbsp;&#37;&nbsp;
				</td>
				
				<td class="formMain" valign="top" align="center">
				For Employee<br>
				<input type="text" size="3" maxlength="3" value="<% If say = "edit" Then %><%= GetExpenses("BenMedPartEmployee") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesBenMedPartEmployee" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">&nbsp;&#37;&nbsp;
				</td>		
				
				<td class="formMain" valign="top" align="center">
				For Family<br>
				<input type="text" size="3" maxlength="3" value="<% If say = "edit" Then %><%= GetExpenses("BenMedPartFamily") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesBenMedPartFamily" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);" >&nbsp;&#37;&nbsp;
				</td>						
			
			</tr>
			
			<tr>
				<td class="formMain" align="center">Dental</td>

				<td class="formMain" valign="top" align="center">
				For Employee<br>
				<input type="text" size="3" maxlength="3" value="<% If say = "edit" Then %><%= GetExpenses("BenDentFullEmployee") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesBenDentFullEmployee" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">&nbsp;&#37;&nbsp;
				</td>
				
				<td class="formMain" valign="top" align="center">
				For Family<br>
				<input type="text" size="3" maxlength="3" value="<% If say = "edit" Then %><%= GetExpenses("BenDentFullFamily") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesBenDentFullFamily" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">&nbsp;&#37;&nbsp;
				</td>
				
				<td class="formMain" valign="top" align="center">
				For Employee<br>
				<input type="text" size="3" maxlength="3" value="<% If say = "edit" Then %><%= GetExpenses("BenDentPartEmployee") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesBenDentPartEmployee" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">&nbsp;&#37;&nbsp;
				</td>		
				
				<td class="formMain" valign="top" align="center">
				For Family<br>
				<input type="text" size="3" maxlength="3" value="<% If say = "edit" Then %><%= GetExpenses("BenDentPartFamily") %><% Else %>0<% End If %>" class="formMain" name="frmExpensesBenDentPartFamily" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);" >&nbsp;&#37;&nbsp;
				</td>						
			
			</tr>		
			
			<tr>
				<td colspan="6" class="formHeaderSmall">BENEFITS - NON-MEDICAL<br>(check all that apply)</td>
			</tr>	
			
			<tr>
				<td class="formMain" colspan="3">&nbsp;</td>
				<td class="formMain" align="center">Full Time</td>
				<td class="formMain" align="center">Part Time</td>

			</tr>			
			
			<tr>
				<td class="formMain" colspan="3">Disability Insurance SHORT Term</td>
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesDisInsShortTermFull" value="1"  <% If say = "edit" Then %><% if GetExpenses("DisInsShortTermFull")=true then%>checked<% end if %><% end if %>></td>
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesDisInsShortTermPart" value="1"  <% If say = "edit" Then %><% if GetExpenses("DisInsShortTermPart")=true then%>checked<% end if %><% end if %>></td>				
			</tr>
			
			<tr>
				<td class="formMain" colspan="3">Disability Insurance LONG Term</td>				
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesDisInsLongTermFull" value="1"  <% If say = "edit" Then %><% if GetExpenses("DisInsLongTermFull")=true then%>checked<% end if %><% end if %>></td>
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesDisInsLongTermPart" value="1"  <% If say = "edit" Then %><% if GetExpenses("DisInsLongTermPart")=true then%>checked<% end if %><% end if %>></td>								
			</tr>
			
			<tr>
				<td class="formMain" colspan="3">EAP: Employee Assistance Programs</td>
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesEAPFull" value="1"  <% If say = "edit" Then %><% if GetExpenses("EAPFull")=true then%>checked<% end if %><% end if %>></td>				
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesEAPPart" value="1"  <% If say = "edit" Then %><% if GetExpenses("EAPPart")=true then%>checked<% end if %><% end if %>></td>								
			</tr>
			
			<tr>
				<td class="formMain" colspan="3">"Flex" Pre-Tax Plan (medical, dependent)</td>			
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesFlexFull" value="1"  <% If say = "edit" Then %><% if GetExpenses("FlexFull")=true then%>checked<% end if %><% end if %>></td>				
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesFlexPart" value="1"  <% If say = "edit" Then %><% if GetExpenses("FlexPart")=true then%>checked<% end if %><% end if %>></td>								
			</tr>			
			
			<tr>
				<td class="formMain" colspan="3">Health Club</td>			
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesHealthClubFull" value="1"  <% If say = "edit" Then %><% if GetExpenses("HealthClubFull")=true then%>checked<% end if %><% end if %>></td>				
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesHealthClubPart" value="1"  <% If say = "edit" Then %><% if GetExpenses("HealthClubPart")=true then%>checked<% end if %><% end if %>></td>								
			</tr>		
			
			<tr>
				<td class="formMain" colspan="3">Life Insurance</td>			
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesLifeInsuranceFull" value="1"  <% If say = "edit" Then %><% if GetExpenses("LifeInsuranceFull")=true then%>checked<% end if %><% end if %>></td>				
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesLifeInsurancePart" value="1"  <% If say = "edit" Then %><% if GetExpenses("LifeInsurancePart")=true then%>checked<% end if %><% end if %>></td>								
			</tr>								
			
			<tr>
				<td class="formMain" colspan="3">Paid Time Off (Floating Holidays, Personal)</td>			
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesTimeOffFull" value="1"  <% If say = "edit" Then %><% if GetExpenses("TimeOffFull")=true then%>checked<% end if %><% end if %>></td>				
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesTimeOffPart" value="1"  <% If say = "edit" Then %><% if GetExpenses("TimeOffPart")=true then%>checked<% end if %><% end if %>></td>								
			</tr>		
			
			<tr>
				<td class="formMain" colspan="3">Paid Time Off (Sick Time)</td>			
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesTimeOffSickFull" value="1"  <% If say = "edit" Then %><% if GetExpenses("TimeOffSickFull")=true then%>checked<% end if %><% end if %>></td>				
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesTimeOffSickPart" value="1"  <% If say = "edit" Then %><% if GetExpenses("TimeOffSickPart")=true then%>checked<% end if %><% end if %>></td>								
			</tr>													
					
			<tr>
				<td class="formMain" colspan="3">Paid Time Off (Vacation)</td>			
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesTimeOffVacFull" value="1"  <% If say = "edit" Then %><% if GetExpenses("TimeOffVacFull")=true then%>checked<% end if %><% end if %>></td>				
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesTimeOffVacPart" value="1"  <% If say = "edit" Then %><% if GetExpenses("TimeOffVacPart")=true then%>checked<% end if %><% end if %>></td>								
			</tr>	
			
			<tr>
				<td class="formMain" colspan="3">Professional Dues, Conferences, etc.</td>			
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesProfDuesFull" value="1"  <% If say = "edit" Then %><% if GetExpenses("ProfDuesFull")=true then%>checked<% end if %><% end if %>></td>				
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesProfDuesPart" value="1"  <% If say = "edit" Then %><% if GetExpenses("ProfDuesPart")=true then%>checked<% end if %><% end if %>></td>								
			</tr>			
			
			<tr>
				<td class="formMain" colspan="3">Retirement</td>			
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesRetirementFull" value="1"  <% If say = "edit" Then %><% if GetExpenses("RetirementFull")=true then%>checked<% end if %><% end if %>></td>				
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesRetirementPart" value="1"  <% If say = "edit" Then %><% if GetExpenses("RetirementPart")=true then%>checked<% end if %><% end if %>></td>								
			</tr>						
			
			<tr>
				<td class="formMain" colspan="3">Telecommuting</td>			
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesTelecommFull" value="1"  <% If say = "edit" Then %><% if GetExpenses("TelecommFull")=true then%>checked<% end if %><% end if %>></td>				
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesTelecommPart" value="1"  <% If say = "edit" Then %><% if GetExpenses("TelecommPart")=true then%>checked<% end if %><% end if %>></td>								
			</tr>						
			
			<tr>
				<td class="formMain" colspan="3">Tuition</td>			
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesTuitionFull" value="1"  <% If say = "edit" Then %><% if GetExpenses("TuitionFull")=true then%>checked<% end if %><% end if %>></td>				
				<td class="formMain" align="center"><input type="checkbox" class="formMain" name="frmExpensesTuitionPart" value="1"  <% If say = "edit" Then %><% if GetExpenses("TuitionPart")=true then%>checked<% end if %><% end if %>></td>								
			</tr>									
						
			</table>
		</td>
	
	</tr>
		

	<tr>
		<td colspan="6" class="formHeader"><input type="submit" value="Save Form" class="formMainBold"></td>
	</tr>
	<tr>
		<td colspan="6" class="formMain" align="center"><!--#include file="../includes/contact_info.inc"--></td>
	</tr>
</table>
<% 
If say = "edit" Then
	GetExpenses.Close
	Set GetExpenses = Nothing
	Con.Close
	Set Con = Nothing
End If
 %>

</form>


<% End If %>
</td>
</tr>
</table>
</body>
</html>
