<!--#include file="../includes/NAD_BE.asp" -->

<% 

If Request("status") = "addNew" Then

	
	
' Check for duplicate records
	
	Set DupCon = Server.CreateObject("ADODB.Connection")
	DupCon.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT Count(*) As NumberOfEntries FROM tbl_frmSDMPerformance WHERE AgencyID = '" & Request("AgencyIDN") & "' and Year = " & Request("Year")	& " and Month = " & Request("Month")
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
		RST.Open "SELECT * FROM tbl_frmSDMPerformance", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = Request("AgencyIDN")
		RST("Year") = Request("year")
		RST("Month") = Request("month")
		RST("AverageMatchLengthCB") = Request("frmSDMPerformanceAverageMatchLengthCB")
		RST("AverageMatchLengthSB") = Request("frmSDMPerformanceAverageMatchLengthSB")
		RST("AverageMatchLengthOSB") = Request("frmSDMPerformanceAverageMatchLengthOSB")
		
		RST("ProcTim_Vol_InquiryToInterview_Number_Comm") = Request("frmSDMPerformanceProcTim_Vol_InquiryToInterview_Number_Comm")
		RST("ProcTim_Vol_InquiryToInterview_AveDays_Comm") = Request("frmSDMPerformanceProcTim_Vol_InquiryToInterview_AveDays_Comm")
		RST("ProcTim_Vol_InquiryToInterview_Number_School") = Request("frmSDMPerformanceProcTim_Vol_InquiryToInterview_Number_School")
		RST("ProcTim_Vol_InquiryToInterview_AveDays_School") = Request("frmSDMPerformanceProcTim_Vol_InquiryToInterview_AveDays_School")
		RST("ProcTim_Vol_InquiryToInterview_Number_Other") = Request("frmSDMPerformanceProcTim_Vol_InquiryToInterview_Number_Other")
		RST("ProcTim_Vol_InquiryToInterview_AveDays_Other") = Request("frmSDMPerformanceProcTim_Vol_InquiryToInterview_AveDays_Other")
		RST("ProcTim_Vol_InterviewToMatched_Number_Comm") = Request("frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_Comm")
		RST("ProcTim_Vol_InterviewToMatched_AveDays_Comm") = Request("frmSDMPerformanceProcTim_Vol_InterviewToMatched_AveDays_Comm")
		RST("ProcTim_Vol_InterviewToMatched_Number_School") = Request("frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_School")
		RST("ProcTim_Vol_InterviewToMatched_AveDays_School") = Request("frmSDMPerformanceProcTim_Vol_InterviewToMatched_AveDays_School")
		RST("ProcTim_Vol_InterviewToMatched_Number_Other") = Request("frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_Other")
		RST("ProcTim_Vol_InterviewToMatched_AveDays_Other") = Request("frmSDMPerformanceProcTim_Vol_InterviewToMatched_AveDays_Other")
		RST("ProcTim_Youth_InquiryToInterview_Number_Comm") = Request("frmSDMPerformanceProcTim_Youth_InquiryToInterview_Number_Comm")
		RST("ProcTim_Youth_InquiryToInterview_AveDays_Comm") = Request("frmSDMPerformanceProcTim_Youth_InquiryToInterview_AveDays_Comm")
		RST("ProcTim_Youth_InquiryToInterview_Number_School") = Request("frmSDMPerformanceProcTim_Youth_InquiryToInterview_Number_School")
		RST("ProcTim_Youth_InquiryToInterview_AveDays_School") = Request("frmSDMPerformanceProcTim_Youth_InquiryToInterview_AveDays_School")
		RST("ProcTim_Youth_InquiryToInterview_Number_Other") = Request("frmSDMPerformanceProcTim_Youth_InquiryToInterview_Number_Other")
		RST("ProcTim_Youth_InquiryToInterview_AveDays_Other") = Request("frmSDMPerformanceProcTim_Youth_InquiryToInterview_AveDays_Other")
		RST("ProcTim_Youth_InterviewToMatched_Number_Comm") = Request("frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_Comm")
		RST("ProcTim_Youth_InterviewToMatched_AveDays_Comm") = Request("frmSDMPerformanceProcTim_Youth_InterviewToMatched_AveDays_Comm")
		RST("ProcTim_Youth_InterviewToMatched_Number_School") = Request("frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_School")
		RST("ProcTim_Youth_InterviewToMatched_AveDays_School") = Request("frmSDMPerformanceProcTim_Youth_InterviewToMatched_AveDays_School")
		RST("ProcTim_Youth_InterviewToMatched_Number_Other") = Request("frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_Other")
		RST("ProcTim_Youth_InterviewToMatched_AveDays_Other") = Request("frmSDMPerformanceProcTim_Youth_InterviewToMatched_AveDays_Other")
		RST("Freq_Under3Months_Comm") = Request("frmSDMPerformanceFreq_Under3Months_Comm")
		RST("Freq_Under3Months_School") = Request("frmSDMPerformanceFreq_Under3Months_School")
		RST("Freq_Under3Months_Other") = Request("frmSDMPerformanceFreq_Under3Months_Other")		
		RST("Freq_3To6Months_Comm") = Request("frmSDMPerformanceFreq_3To6Months_Comm")
		RST("Freq_3To6Months_School") = Request("frmSDMPerformanceFreq_3To6Months_School")
		RST("Freq_3To6Months_Other") = Request("frmSDMPerformanceFreq_3To6Months_Other")		
		RST("Freq_7To9Months_Comm") = Request("frmSDMPerformanceFreq_7To9Months_Comm")
		RST("Freq_7To9Months_School") = Request("frmSDMPerformanceFreq_7To9Months_School")
		RST("Freq_7To9Months_Other") = Request("frmSDMPerformanceFreq_7To9Months_Other")		
		RST("Freq_10To12Months_Comm") = Request("frmSDMPerformanceFreq_10To12Months_Comm")
		RST("Freq_10To12Months_School") = Request("frmSDMPerformanceFreq_10To12Months_School")
		RST("Freq_10To12Months_Other") = Request("frmSDMPerformanceFreq_10To12Months_Other")		
		RST("Freq_13To23Months_Comm") = Request("frmSDMPerformanceFreq_13To23Months_Comm")
		RST("Freq_13To23Months_School") = Request("frmSDMPerformanceFreq_13To23Months_School")
		RST("Freq_13To23Months_Other") = Request("frmSDMPerformanceFreq_13To23Months_Other")		
		RST("Freq_24OrMoreMonths_Comm") = Request("frmSDMPerformanceFreq_24OrMoreMonths_Comm")
		RST("Freq_24OrMoreMonths_School") = Request("frmSDMPerformanceFreq_24OrMoreMonths_School")
		RST("Freq_24OrMoreMonths_Other") = Request("frmSDMPerformanceFreq_24OrMoreMonths_Other")		
		RST("Volunteers_ReMatchedCB") = Request("frmSDMPerformanceVolunteers_ReMatchedCB")		
		RST("Volunteers_ReMatchedSB") = Request("frmSDMPerformanceVolunteers_ReMatchedSB")		
		RST("Volunteers_ReMatchedOSB") = Request("frmSDMPerformanceVolunteers_ReMatchedOSB")		
		RST("CBNumberClosedPrematurely") = Request("frmSDMPerformanceCBNumberClosedPrematurely")
		RST("SBNumberClosedPrematurely") = Request("frmSDMPerformanceSBNumberClosedPrematurely")				
		RST("CBChildParentStatusChange") = Request("frmSDMPerformanceCBChildParentStatusChange")
		RST("CBVolunteerStatusChange") = Request("frmSDMPerformanceCBVolunteerStatusChange")
		RST("CBChildParentDissatisfaction") = Request("frmSDMPerformanceCBChildParentDissatisfaction")
		RST("CBVolunteerDissatisfaction") = Request("frmSDMPerformanceCBVolunteerDissatisfaction")

		RST("CBSuccessfulMatches") = Request("frmSDMPerformanceCBSuccessfulMatches")		
		RST("SBSuccessfulMatches") = Request("frmSDMPerformanceSBSuccessfulMatches")				
		RST("OSBSuccessfulMatches") = Request("frmSDMPerformanceOSBSuccessfulMatches")						
	
	
		RST("SBChildParentStatusChange") = Request("frmSDMPerformanceSBChildParentStatusChange")
		RST("SBVolunteerStatusChange") = Request("frmSDMPerformanceSBVolunteerStatusChange")
		RST("SBChildParentDissatisfaction") = Request("frmSDMPerformanceSBChildParentDissatisfaction")
		RST("SBVolunteerDissatisfaction") = Request("frmSDMPerformanceSBVolunteerDissatisfaction")
		RST("CBTotalOpened6MonthsAgo") = Request("frmSDMPerformanceCBTotalOpened6MonthsAgo")
		RST("CBNumberStillOpen") = Request("frmSDMPerformanceCBNumberStillOpen")
		RST("EnrollmentSatAvgScore") = Request("frmSDMPerformanceEnrollmentSatAvgScore")
		RST("EnrollmentSatCount") = Request("frmSDMPerformanceEnrollmentSatCount")
		RST("MatchSatAvgScore") = Request("frmSDMPerformanceMatchSatAvgScore")
		RST("MatchSatCount") = Request("frmSDMPerformanceMatchSatCount")
		RST("CBPOEAggregateScore") = Request("frmSDMPerformanceCBPOEAggregateScore")
		RST("CBPOECount") = Request("frmSDMPerformanceCBPOECount")
		RST("SBPOEAggregateScore") = Request("frmSDMPerformanceSBPOEAggregateScore")
		RST("SBPOECount") = Request("frmSDMPerformanceSBPOECount")
		
		RST("YieldRate_Vol_Inquiries_CB") = Request("frmSDMPerformanceYieldRate_Vol_Inquiries_CB")
		RST("YieldRate_Vol_Inquiries_SB") = Request("frmSDMPerformanceYieldRate_Vol_Inquiries_SB")
		RST("YieldRate_Vol_Inquiries_OSB") = Request("frmSDMPerformanceYieldRate_Vol_Inquiries_OSB")
			
		RST("YieldRate_Youth_Inquiries_CB") = Request("frmSDMPerformanceYieldRate_Youth_Inquiries_CB")
		RST("YieldRate_Youth_Inquiries_SB") = Request("frmSDMPerformanceYieldRate_Youth_Inquiries_SB")
		RST("YieldRate_Youth_Inquiries_OSB") = Request("frmSDMPerformanceYieldRate_Youth_Inquiries_OSB")
		
		RST("OSBNumberClosedPrematurely") = Request("frmSDMPerformanceOSBNumberClosedPrematurely")
		RST("OSBChildParentStatusChange") = Request("frmSDMPerformanceOSBChildParentStatusChange")
		RST("OSBVolunteerStatusChange") = Request("frmSDMPerformanceOSBVolunteerStatusChange")
		RST("OSBChildParentDissatisfaction") = Request("frmSDMPerformanceOSBChildParentDissatisfaction")
		RST("OSBVolunteerDissatisfaction") = Request("frmSDMPerformanceOSBVolunteerDissatisfaction")
		RST("SBTotalOpened6MonthsAgo") = Request("frmSDMPerformanceSBTotalOpened6MonthsAgo")
		RST("SBNumberStillOpen") = Request("frmSDMPerformanceSBNumberStillOpen")
		RST("OSBTotalOpened6MonthsAgo") = Request("frmSDMPerformanceOSBTotalOpened6MonthsAgo")
		RST("OSBNumberStillOpen") = Request("frmSDMPerformanceOSBNumberStillOpen")																	
		RST("OSBPOEAggregateScore") = Request("frmSDMPerformanceOSBPOEAggregateScore")
		RST("OSBPOECount") = Request("frmSDMPerformanceOSBPOECount")		

		
		RST("CreateDate") = Now
		RST.Update
		RST.Close
		Set RST = Nothing
		form = "SDMPerformance"
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
	RST.Open "SELECT * FROM tbl_frmSDMPerformance WHERE agencyID='" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")) & " AND Month=" & Int(Request("month")), Con, 1, 3
	RST("AverageMatchLengthCB") = Request("frmSDMPerformanceAverageMatchLengthCB")
	RST("AverageMatchLengthSB") = Request("frmSDMPerformanceAverageMatchLengthSB")
	RST("AverageMatchLengthOSB") = Request("frmSDMPerformanceAverageMatchLengthOSB")	
	
	RST("ProcTim_Vol_InquiryToInterview_Number_Comm") = Request("frmSDMPerformanceProcTim_Vol_InquiryToInterview_Number_Comm")
	RST("ProcTim_Vol_InquiryToInterview_AveDays_Comm") = Request("frmSDMPerformanceProcTim_Vol_InquiryToInterview_AveDays_Comm")
	RST("ProcTim_Vol_InquiryToInterview_Number_School") = Request("frmSDMPerformanceProcTim_Vol_InquiryToInterview_Number_School")
	RST("ProcTim_Vol_InquiryToInterview_AveDays_School") = Request("frmSDMPerformanceProcTim_Vol_InquiryToInterview_AveDays_School")
	RST("ProcTim_Vol_InquiryToInterview_Number_Other") = Request("frmSDMPerformanceProcTim_Vol_InquiryToInterview_Number_Other")
	RST("ProcTim_Vol_InquiryToInterview_AveDays_Other") = Request("frmSDMPerformanceProcTim_Vol_InquiryToInterview_AveDays_Other")
	RST("ProcTim_Vol_InterviewToMatched_Number_Comm") = Request("frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_Comm")
	RST("ProcTim_Vol_InterviewToMatched_AveDays_Comm") = Request("frmSDMPerformanceProcTim_Vol_InterviewToMatched_AveDays_Comm")
	RST("ProcTim_Vol_InterviewToMatched_Number_School") = Request("frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_School")
	RST("ProcTim_Vol_InterviewToMatched_AveDays_School") = Request("frmSDMPerformanceProcTim_Vol_InterviewToMatched_AveDays_School")
	RST("ProcTim_Vol_InterviewToMatched_Number_Other") = Request("frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_Other")
	RST("ProcTim_Vol_InterviewToMatched_AveDays_Other") = Request("frmSDMPerformanceProcTim_Vol_InterviewToMatched_AveDays_Other")
	RST("ProcTim_Youth_InquiryToInterview_Number_Comm") = Request("frmSDMPerformanceProcTim_Youth_InquiryToInterview_Number_Comm")
	RST("ProcTim_Youth_InquiryToInterview_AveDays_Comm") = Request("frmSDMPerformanceProcTim_Youth_InquiryToInterview_AveDays_Comm")
	RST("ProcTim_Youth_InquiryToInterview_Number_School") = Request("frmSDMPerformanceProcTim_Youth_InquiryToInterview_Number_School")
	RST("ProcTim_Youth_InquiryToInterview_AveDays_School") = Request("frmSDMPerformanceProcTim_Youth_InquiryToInterview_AveDays_School")
	RST("ProcTim_Youth_InquiryToInterview_Number_Other") = Request("frmSDMPerformanceProcTim_Youth_InquiryToInterview_Number_Other")
	RST("ProcTim_Youth_InquiryToInterview_AveDays_Other") = Request("frmSDMPerformanceProcTim_Youth_InquiryToInterview_AveDays_Other")
	RST("ProcTim_Youth_InterviewToMatched_Number_Comm") = Request("frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_Comm")
	RST("ProcTim_Youth_InterviewToMatched_AveDays_Comm") = Request("frmSDMPerformanceProcTim_Youth_InterviewToMatched_AveDays_Comm")
	RST("ProcTim_Youth_InterviewToMatched_Number_School") = Request("frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_School")
	RST("ProcTim_Youth_InterviewToMatched_AveDays_School") = Request("frmSDMPerformanceProcTim_Youth_InterviewToMatched_AveDays_School")
	RST("ProcTim_Youth_InterviewToMatched_Number_Other") = Request("frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_Other")
	RST("ProcTim_Youth_InterviewToMatched_AveDays_Other") = Request("frmSDMPerformanceProcTim_Youth_InterviewToMatched_AveDays_Other")
	RST("Freq_Under3Months_Comm") = Request("frmSDMPerformanceFreq_Under3Months_Comm")
	RST("Freq_Under3Months_School") = Request("frmSDMPerformanceFreq_Under3Months_School")
	RST("Freq_Under3Months_Other") = Request("frmSDMPerformanceFreq_Under3Months_Other")		
	RST("Freq_3To6Months_Comm") = Request("frmSDMPerformanceFreq_3To6Months_Comm")
	RST("Freq_3To6Months_School") = Request("frmSDMPerformanceFreq_3To6Months_School")
	RST("Freq_3To6Months_Other") = Request("frmSDMPerformanceFreq_3To6Months_Other")		
	RST("Freq_7To9Months_Comm") = Request("frmSDMPerformanceFreq_7To9Months_Comm")
	RST("Freq_7To9Months_School") = Request("frmSDMPerformanceFreq_7To9Months_School")
	RST("Freq_7To9Months_Other") = Request("frmSDMPerformanceFreq_7To9Months_Other")		
	RST("Freq_10To12Months_Comm") = Request("frmSDMPerformanceFreq_10To12Months_Comm")
	RST("Freq_10To12Months_School") = Request("frmSDMPerformanceFreq_10To12Months_School")
	RST("Freq_10To12Months_Other") = Request("frmSDMPerformanceFreq_10To12Months_Other")		
	RST("Freq_13To23Months_Comm") = Request("frmSDMPerformanceFreq_13To23Months_Comm")
	RST("Freq_13To23Months_School") = Request("frmSDMPerformanceFreq_13To23Months_School")
	RST("Freq_13To23Months_Other") = Request("frmSDMPerformanceFreq_13To23Months_Other")		
	RST("Freq_24OrMoreMonths_Comm") = Request("frmSDMPerformanceFreq_24OrMoreMonths_Comm")
	RST("Freq_24OrMoreMonths_School") = Request("frmSDMPerformanceFreq_24OrMoreMonths_School")
	RST("Freq_24OrMoreMonths_Other") = Request("frmSDMPerformanceFreq_24OrMoreMonths_Other")
	RST("Volunteers_ReMatchedCB") = Request("frmSDMPerformanceVolunteers_ReMatchedCB")		
	RST("Volunteers_ReMatchedSB") = Request("frmSDMPerformanceVolunteers_ReMatchedSB")		
	RST("Volunteers_ReMatchedOSB") = Request("frmSDMPerformanceVolunteers_ReMatchedOSB")
	
	RST("CBNumberClosedPrematurely") = Request("frmSDMPerformanceCBNumberClosedPrematurely")
	RST("SBNumberClosedPrematurely") = Request("frmSDMPerformanceSBNumberClosedPrematurely")						
	RST("CBChildParentStatusChange") = Request("frmSDMPerformanceCBChildParentStatusChange")
	RST("CBVolunteerStatusChange") = Request("frmSDMPerformanceCBVolunteerStatusChange")
	RST("CBChildParentDissatisfaction") = Request("frmSDMPerformanceCBChildParentDissatisfaction")
	RST("CBVolunteerDissatisfaction") = Request("frmSDMPerformanceCBVolunteerDissatisfaction")
	RST("SBChildParentStatusChange") = Request("frmSDMPerformanceSBChildParentStatusChange")
	RST("SBVolunteerStatusChange") = Request("frmSDMPerformanceSBVolunteerStatusChange")
	RST("SBChildParentDissatisfaction") = Request("frmSDMPerformanceSBChildParentDissatisfaction")
	RST("SBVolunteerDissatisfaction") = Request("frmSDMPerformanceSBVolunteerDissatisfaction")

	RST("CBSuccessfulMatches") = Request("frmSDMPerformanceCBSuccessfulMatches")		
	RST("SBSuccessfulMatches") = Request("frmSDMPerformanceSBSuccessfulMatches")				
	RST("OSBSuccessfulMatches") = Request("frmSDMPerformanceOSBSuccessfulMatches")						

	
	RST("CBTotalOpened6MonthsAgo") = Request("frmSDMPerformanceCBTotalOpened6MonthsAgo")
	RST("CBNumberStillOpen") = Request("frmSDMPerformanceCBNumberStillOpen")
	RST("EnrollmentSatAvgScore") = Request("frmSDMPerformanceEnrollmentSatAvgScore")
	RST("EnrollmentSatCount") = Request("frmSDMPerformanceEnrollmentSatCount")
	RST("MatchSatAvgScore") = Request("frmSDMPerformanceMatchSatAvgScore")
	RST("MatchSatCount") = Request("frmSDMPerformanceMatchSatCount")
	RST("CBPOEAggregateScore") = Request("frmSDMPerformanceCBPOEAggregateScore")
	RST("CBPOECount") = Request("frmSDMPerformanceCBPOECount")
	RST("SBPOEAggregateScore") = Request("frmSDMPerformanceSBPOEAggregateScore")	
	RST("SBPOECount") = Request("frmSDMPerformanceSBPOECount")			
	RST("YieldRate_Vol_Inquiries_CB") = Request("frmSDMPerformanceYieldRate_Vol_Inquiries_CB")
	RST("YieldRate_Vol_Inquiries_SB") = Request("frmSDMPerformanceYieldRate_Vol_Inquiries_SB")
	RST("YieldRate_Vol_Inquiries_OSB") = Request("frmSDMPerformanceYieldRate_Vol_Inquiries_OSB")

	RST("YieldRate_Youth_Inquiries_CB") = Request("frmSDMPerformanceYieldRate_Youth_Inquiries_CB")
	RST("YieldRate_Youth_Inquiries_SB") = Request("frmSDMPerformanceYieldRate_Youth_Inquiries_SB")
	RST("YieldRate_Youth_Inquiries_OSB") = Request("frmSDMPerformanceYieldRate_Youth_Inquiries_OSB")

	RST("OSBNumberClosedPrematurely") = Request("frmSDMPerformanceOSBNumberClosedPrematurely")
	RST("OSBChildParentStatusChange") = Request("frmSDMPerformanceOSBChildParentStatusChange")
	RST("OSBVolunteerStatusChange") = Request("frmSDMPerformanceOSBVolunteerStatusChange")
	RST("OSBChildParentDissatisfaction") = Request("frmSDMPerformanceOSBChildParentDissatisfaction")
	RST("OSBVolunteerDissatisfaction") = Request("frmSDMPerformanceOSBVolunteerDissatisfaction")
	RST("SBTotalOpened6MonthsAgo") = Request("frmSDMPerformanceSBTotalOpened6MonthsAgo")
	RST("SBNumberStillOpen") = Request("frmSDMPerformanceSBNumberStillOpen")
	RST("OSBTotalOpened6MonthsAgo") = Request("frmSDMPerformanceOSBTotalOpened6MonthsAgo")
	RST("OSBNumberStillOpen") = Request("frmSDMPerformanceOSBNumberStillOpen")																	
	RST("OSBPOEAggregateScore") = Request("frmSDMPerformanceOSBPOEAggregateScore")
	RST("OSBPOECount") = Request("frmSDMPerformanceOSBPOECount")		
		
	jMod = RST("SDMPerformanceID") %>
	
	
	
	
	<%
	RST.Update
	RST.Close
	Set RST = Nothing
	form = "SDMPerformance"
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
	<title>Performance</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
<script language="javascript">
<!--	


//Field Validations

function checkForIntegerCommas(valueToCheck)
{
	var myRegularExpression = /^[0-9]+(,[0-9]{3})*$/;  // Checks for integer with or without commas
	if(!(myRegularExpression.test(valueToCheck)))
	{
		alert("Please make sure you have entered a whole number with no spaces.\n Do not leave this field blank."); 
	} 
}

function validateForm()
{	
	
	var onlyInteger = /^[0-9]+(,[0-9]{3})*$/;
	
	//Average Match Length
	if(document.frmSDMPerformance.frmSDMPerformanceAverageMatchLengthCB.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceAverageMatchLengthCB.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceAverageMatchLengthSB.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceAverageMatchLengthSB.focus();}					
	else if(document.frmSDMPerformance.frmSDMPerformanceAverageMatchLengthOSB.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceAverageMatchLengthOSB.focus();}							


	//Yield Rate
	else if(document.frmSDMPerformance.frmSDMPerformanceYieldRate_Vol_Inquiries_CB.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceYieldRate_Vol_Inquiries_CB.focus();}							
	else if(document.frmSDMPerformance.frmSDMPerformanceYieldRate_Vol_Inquiries_SB.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceYieldRate_Vol_Inquiries_SB.focus();}											
	else if(document.frmSDMPerformance.frmSDMPerformanceYieldRate_Vol_Inquiries_OSB.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceYieldRate_Vol_Inquiries_OSB.focus();}

	else if(document.frmSDMPerformance.frmSDMPerformanceYieldRate_Youth_Inquiries_CB.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceYieldRate_Youth_Inquiries_CB.focus();}			
	else if(document.frmSDMPerformance.frmSDMPerformanceYieldRate_Youth_Inquiries_SB.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceYieldRate_Youth_Inquiries_SB.focus();}					
	else if(document.frmSDMPerformance.frmSDMPerformanceYieldRate_Youth_Inquiries_OSB.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceYieldRate_Youth_Inquiries_OSB.focus();}			

	//Volunteer Inquiry to Interview	
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InquiryToInterview_Number_Comm.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InquiryToInterview_Number_Comm.focus();}													
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InquiryToInterview_Number_School.value == "")	
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InquiryToInterview_Number_School.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InquiryToInterview_Number_Other.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InquiryToInterview_Number_Other.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InquiryToInterview_AveDays_Comm.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InquiryToInterview_AveDays_Comm.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InquiryToInterview_AveDays_School.value == "")	
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InquiryToInterview_AveDays_School.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InquiryToInterview_AveDays_Other.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InquiryToInterview_AveDays_Other.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_Comm.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_Comm.focus();}				

	// Volunteer Interview to Matched
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_Comm.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_Comm.focus();}						
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_School.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_School.focus();}						
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_Other.value == "")	
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_Other.focus();}				
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InterviewToMatched_AveDays_Comm.value == "")	
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InterviewToMatched_AveDays_Comm.focus();}				
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InterviewToMatched_AveDays_School.value == "")		
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InterviewToMatched_AveDays_School.focus();}				
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InterviewToMatched_AveDays_Other.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Vol_InterviewToMatched_AveDays_Other.focus();}		

	//Youth Inquiry to Interview	
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InquiryToInterview_Number_Comm.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InquiryToInterview_Number_Comm.focus();}													
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InquiryToInterview_Number_School.value == "")	
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InquiryToInterview_Number_School.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InquiryToInterview_Number_Other.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InquiryToInterview_Number_Other.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InquiryToInterview_AveDays_Comm.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InquiryToInterview_AveDays_Comm.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InquiryToInterview_AveDays_School.value == "")	
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InquiryToInterview_AveDays_School.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InquiryToInterview_AveDays_Other.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InquiryToInterview_AveDays_Other.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_Comm.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_Comm.focus();}				

	// Youth Interview to Matched
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_Comm.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_Comm.focus();}						
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_School.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_School.focus();}						
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_Other.value == "")	
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_Other.focus();}				
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InterviewToMatched_AveDays_Comm.value == "")	
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InterviewToMatched_AveDays_Comm.focus();}				
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InterviewToMatched_AveDays_School.value == "")		
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InterviewToMatched_AveDays_School.focus();}				
	else if(document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InterviewToMatched_AveDays_Other.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceProcTim_Youth_InterviewToMatched_AveDays_Other.focus();}		

	
	//Frequency of Match Closure
	else if(document.frmSDMPerformance.frmSDMPerformanceFreq_Under3Months_Comm.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceFreq_Under3Months_Comm.focus();}			
	else if(document.frmSDMPerformance.frmSDMPerformanceFreq_Under3Months_School.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceFreq_Under3Months_School.focus();}					
	else if(document.frmSDMPerformance.frmSDMPerformanceFreq_Under3Months_Other.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceFreq_Under3Months_Other.focus();}					
	else if(document.frmSDMPerformance.frmSDMPerformanceFreq_3To6Months_Comm.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceFreq_3To6Months_Comm.focus();}					
	else if(document.frmSDMPerformance.frmSDMPerformanceFreq_3To6Months_School.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceFreq_3To6Months_School.focus();}			
	else if(document.frmSDMPerformance.frmSDMPerformanceFreq_3To6Months_Other.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceFreq_3To6Months_Other.focus();}	
	else if(document.frmSDMPerformance.frmSDMPerformanceFreq_7To9Months_Comm.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceFreq_7To9Months_Comm.focus();}
	else if(document.frmSDMPerformance.frmSDMPerformanceFreq_7To9Months_School.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceFreq_7To9Months_School.focus();}					
	else if(document.frmSDMPerformance.frmSDMPerformanceFreq_7To9Months_Other.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceFreq_7To9Months_Other.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceFreq_10To12Months_Comm.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceFreq_10To12Months_Comm.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceFreq_10To12Months_School.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceFreq_10To12Months_School.focus();}				
	else if(document.frmSDMPerformance.frmSDMPerformanceFreq_10To12Months_Other.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceFreq_10To12Months_Other.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceFreq_13To23Months_Comm.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceFreq_13To23Months_Comm.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceFreq_13To23Months_School.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceFreq_13To23Months_School.focus();}				
	else if(document.frmSDMPerformance.frmSDMPerformanceFreq_13To23Months_Other.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceFreq_13To23Months_Other.focus();}			
	else if(document.frmSDMPerformance.frmSDMPerformanceFreq_24OrMoreMonths_Comm.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceFreq_24OrMoreMonths_Comm.focus();}					
	else if(document.frmSDMPerformance.frmSDMPerformanceFreq_24OrMoreMonths_School.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceFreq_24OrMoreMonths_School.focus();}			
	else if(document.frmSDMPerformance.frmSDMPerformanceFreq_24OrMoreMonths_Other.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceFreq_24OrMoreMonths_Other.focus();}			

	
	//Volunteers Rematched
	else if(document.frmSDMPerformance.frmSDMPerformanceVolunteers_ReMatchedCB.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceVolunteers_ReMatchedCB.focus();}			
	else if(document.frmSDMPerformance.frmSDMPerformanceVolunteers_ReMatchedSB.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceVolunteers_ReMatchedSB.focus();}			
	else if(document.frmSDMPerformance.frmSDMPerformanceVolunteers_ReMatchedOSB.value == "")			
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceVolunteers_ReMatchedOSB.focus();}					
		
	//Premature Closure
	else if(document.frmSDMPerformance.frmSDMPerformanceCBNumberClosedPrematurely.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceCBNumberClosedPrematurely.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceSBNumberClosedPrematurely.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceSBNumberClosedPrematurely.focus();}				
	else if(document.frmSDMPerformance.frmSDMPerformanceOSBNumberClosedPrematurely.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceOSBNumberClosedPrematurely.focus();}

		
		
	//Closure Codes		
	else if(document.frmSDMPerformance.frmSDMPerformanceCBChildParentStatusChange.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceCBChildParentStatusChange.focus();}				
	else if(document.frmSDMPerformance.frmSDMPerformanceSBChildParentStatusChange.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceSBChildParentStatusChange.focus();}				
	else if(document.frmSDMPerformance.frmSDMPerformanceOSBChildParentStatusChange.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceOSBChildParentStatusChange.focus();}						
		
	else if(document.frmSDMPerformance.frmSDMPerformanceCBVolunteerStatusChange.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceCBVolunteerStatusChange.focus();}				
	else if(document.frmSDMPerformance.frmSDMPerformanceSBVolunteerStatusChange.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceSBVolunteerStatusChange.focus();}				
	else if(document.frmSDMPerformance.frmSDMPerformanceOSBVolunteerStatusChange.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceOSBVolunteerStatusChange.focus();}								
	
	else if(document.frmSDMPerformance.frmSDMPerformanceCBChildParentDissatisfaction.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceCBChildParentDissatisfaction.focus();}						
	else if(document.frmSDMPerformance.frmSDMPerformanceSBChildParentDissatisfaction.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceSBChildParentDissatisfaction.focus();}			
	else if(document.frmSDMPerformance.frmSDMPerformanceOSBChildParentDissatisfaction.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceOSBChildParentDissatisfaction.focus();}			
	
	else if(document.frmSDMPerformance.frmSDMPerformanceCBVolunteerDissatisfaction.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceCBVolunteerDissatisfaction.focus();}					
	else if(document.frmSDMPerformance.frmSDMPerformanceSBVolunteerDissatisfaction.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceSBVolunteerDissatisfaction.focus();}							
	else if(document.frmSDMPerformance.frmSDMPerformanceOSBVolunteerDissatisfaction.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceOSBVolunteerDissatisfaction.focus();}							

	else if(document.frmSDMPerformance.frmSDMPerformanceCBSuccessfulMatches.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceCBSuccessfulMatches.focus();}					
	else if(document.frmSDMPerformance.frmSDMPerformanceSBSuccessfulMatches.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceSBSuccessfulMatches.focus();}							
	else if(document.frmSDMPerformance.frmSDMPerformanceOSBSuccessfulMatches.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceOSBSuccessfulMatches.focus();}									
		
		
		
	//6-Month Retention		
	else if(document.frmSDMPerformance.frmSDMPerformanceCBTotalOpened6MonthsAgo.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceCBTotalOpened6MonthsAgo.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceCBNumberStillOpen.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceCBNumberStillOpen.focus();}	
	else if(document.frmSDMPerformance.frmSDMPerformanceSBTotalOpened6MonthsAgo.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceSBTotalOpened6MonthsAgo.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceSBNumberStillOpen.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceSBNumberStillOpen.focus();}			
	else if(document.frmSDMPerformance.frmSDMPerformanceOSBTotalOpened6MonthsAgo.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceOSBTotalOpened6MonthsAgo.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceOSBNumberStillOpen.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceOSBNumberStillOpen.focus();}					
		
	
	//Customer Satisfaction										
	else if(document.frmSDMPerformance.frmSDMPerformanceEnrollmentSatAvgScore.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceEnrollmentSatAvgScore.focus();}												
	else if(document.frmSDMPerformance.frmSDMPerformanceEnrollmentSatAvgScore.value > 5 || frmSDMPerformance.frmSDMPerformanceEnrollmentSatAvgScore.value < 0 && (document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 3 || document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 6 || document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 9 || document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 12))
		{alert("Enrollment Satisfaction Average Score Must be between 0 and 5");document.frmSDMPerformance.frmSDMPerformanceEnrollmentSatAvgScore.focus();}					
		
	else if(document.frmSDMPerformance.frmSDMPerformanceEnrollmentSatCount.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceEnrollmentSatCount.focus();}														
	
	else if(document.frmSDMPerformance.frmSDMPerformanceMatchSatAvgScore.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceMatchSatAvgScore.focus();}			
	else if(document.frmSDMPerformance.frmSDMPerformanceMatchSatAvgScore.value > 5 || frmSDMPerformance.frmSDMPerformanceMatchSatAvgScore.value < 0 && (document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 3 || document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 6 || document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 9 || document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 12))
		{alert("Match Satisfaction Average Score must be between 0 and 5");document.frmSDMPerformance.frmSDMPerformanceMatchSatAvgScore.focus();}					
	
	else if(document.frmSDMPerformance.frmSDMPerformanceMatchSatCount.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceMatchSatCount.focus();}			

	//POE
	else if(document.frmSDMPerformance.frmSDMPerformanceCBPOEAggregateScore.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceCBPOEAggregateScore.focus();}	
	else if(document.frmSDMPerformance.frmSDMPerformanceCBPOEAggregateScore.value > 5 || frmSDMPerformance.frmSDMPerformanceCBPOEAggregateScore.value < 0 && (document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 3 || document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 6 || document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 9 || document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 12))
		{alert("POE Community Based Aggregate Score must be between 0 and 5");document.frmSDMPerformance.frmSDMPerformanceCBPOEAggregateScore.focus();}			

	else if(document.frmSDMPerformance.frmSDMPerformanceSBPOEAggregateScore.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceSBPOEAggregateScore.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceSBPOEAggregateScore.value > 5 || document.frmSDMPerformance.frmSDMPerformanceSBPOEAggregateScore.value < 0 && (document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 3 || document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 6 || document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 9 || document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 12))
		{alert("POE School Based Aggregate Score must be between 0 and 5");document.frmSDMPerformance.frmSDMPerformanceSBPOEAggregateScore.focus();}				
	
	else if(document.frmSDMPerformance.frmSDMPerformanceOSBPOEAggregateScore.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceOSBPOEAggregateScore.focus();}		
	else if(document.frmSDMPerformance.frmSDMPerformanceOSBPOEAggregateScore.value > 5 || document.frmSDMPerformance.frmSDMPerformanceOSBPOEAggregateScore.value < 0 && (document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 3 || document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 6 || document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 9 || document.frmSDMPerformance.frmSDMPerformanceMonthValue.value == 12))
		{alert("POE Non-School Site Based Aggregate Score must be between 0 and 5");document.frmSDMPerformance.frmSDMPerformanceOSBPOEAggregateScore.focus();}				

	else if(document.frmSDMPerformance.frmSDMPerformanceCBPOECount.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceCBPOECount.focus();}	
	else if(document.frmSDMPerformance.frmSDMPerformanceSBPOECount.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceSBPOECount.focus();}											
	else if(document.frmSDMPerformance.frmSDMPerformanceOSBPOECount.value == "")
		{alert("Please complete all form fields");document.frmSDMPerformance.frmSDMPerformanceOSBPOECount.focus();}											
	
		
	else
		document.frmSDMPerformance.submit();	
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


<form name="frmSDMPerformance" action="SDMPerformance_edit.asp" method="post"> <!-- onsubmit="return submitFormValidate(this)"> -->
<!--#include file="../includes/form_stamp.asp"-->

<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmSDMPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & " AND Month=" & Int(Request("m"))
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
		<table width="650" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063" >
		<tr>
			<td colspan="7" class="formHeader">SDM METRICS COMPONENTS - <%= MonthName(Request("m"), False) & " " & Request("y") %></td>
		</tr>
		
		<tr>
			<td colspan="7" class="formMain"><font color="#ff0000"><div align="center"><strong>Please Note: </strong>After entering your information, you <strong>must</strong> click on the "Save" button at the bottom of the form and wait for the "Thank You" screen or your changes will be lost.</div></font></td>
		</tr>





<!-- SDM Metrics -->
			<tr>
				<td colspan="7" class="formMain" align="center"><strong>IF YOU DO NOT HAVE SDM DATA FOR A PARTICULAR CATEGORY, JUST ENTER A "0" (ZERO)</strong></td>
			</tr>
	<!-- Ignore old questions if reporting year >= 2006 -->
	<% if y < 2006 then %>			
			<TR>
				<TD colspan="7" class="formHeaderMedium">AVERAGE MATCH LENGTH</TD>
			</TR>		
			
			<tr>
				<td>&nbsp;</td>
				<td colspan="2" class="formMain" align="center">Community-Based</td>
				<td colspan="2" class="formMain" align="center">School-Based</td>
				<td colspan="2" class="formMain" align="center">Non-School<br>Site-Based</td>
			</tr>				

			<tr>
				<td align="center" valign="middle" class="formMain">
					AVERAGE&nbsp;LENGTH&nbsp;(In&nbsp;Months)<br>&nbsp;of&nbsp;Matches&nbsp;Closed&nbsp;during<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b>&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_aml" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<td align="center" colspan = "2" valign="middle" class="formMain"> 
						<input type="text"  class="formMain"  size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("AverageMatchLengthCB") %><% Else %>0<% End If %>" name="frmSDMPerformanceAverageMatchLengthCB" tabindex="8" >					
					</td>
					
					<td align="center" colspan = "2" valign="middle" class="formMain">
						<input type="text"  class="formMain"size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("AverageMatchLengthSB") %><% Else %>0<% End If %>" name="frmSDMPerformanceAverageMatchLengthSB" tabindex="8" >					
					</td>
	
					<td align="center" colspan = "2" valign="middle" class="formMain">
						<input type="text"  class="formMain"size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("AverageMatchLengthOSB") %><% Else %>0<% End If %>" name="frmSDMPerformanceAverageMatchLengthOSB" tabindex="8" >					
					</td>		
									
			</tr>			
		
		
		<% else %>
		
					<!-- Pre-Populate old fields with Zeros -->
					<input type="hidden"  value="0" name="frmSDMPerformanceAverageMatchLengthCB">					
					<input type="hidden"  value="0" name="frmSDMPerformanceAverageMatchLengthSB">					
					<input type="hidden"  value="0" name="frmSDMPerformanceAverageMatchLengthOSB">					
		
		<% end if %>
			
			

		<% if y < 2006 then %>	
		<!-- Only Display Yield Heading for older components (prior to 2006) -->	
			<TR>
				<TD colspan="7" class="formHeaderMedium">YIELD AND PROCESSING TIME</TD>
			</TR>	
				
		<% end if %>
			
			<tr>
				<td colspan="7" <% if y < 2006 then %>class="formMain"<%else%>class="formHeaderMedium"<%end if%> align="center"><strong>Volunteer</strong></td>
				<!-- <td colspan="4" class="formMain" align="center"><strong>Child</strong></td>-->
			</tr>		
			
			<tr>
				<td>&nbsp;</td>
				<td colspan="2" class="formMain" align="center" <%if y>=2006 then%> width="100"<%end if%>>Community-Based</td>
				<td colspan="2" class="formMain" align="center" <%if y>=2006 then%> width="100"<%end if%>>School-Based</td>
				<td colspan="2" class="formMain" align="center" <%if y>=2006 then%> width="100"<%end if%>>Non-School<br>Site-Based</td>

		
			</tr>
			
			<tr>
				<td>&nbsp;</td>
				<!-- Only include 'average days' fields for records prior to 2006 -->
				<% if y < 2006 then %>
					<td class="formMain" align="center">Number of Individuals</td>
					<td class="formMain" align="center">Average Days</td>
				<% else %>
					<td class="formMain" colspan="2" align="center" width="100">Number of Individuals</td>				
				<% end if %>
				
				<% if y < 2006 then %>
					<td class="formMain" align="center">Number of Individuals</td>
					<td class="formMain" align="center">Average Days</td>	
				<% else %>
					<td class="formMain" align="center" colspan="2" width="100">Number of Individuals</td>				
				<% end if %>
				
				<% if y < 2006 then %>
					<td class="formMain" align="center">Number of Individuals</td>
					<td class="formMain" align="center">Average Days</td>
				<% else %>
					<td class="formMain" align="center" colspan="2" width="100">Number of Individuals</td>				
				<% end if %>
				
			</tr>
			
			<tr>
				<td class="formMain" align="left">Volunteer Inquiries</td>
				
				<% if y < 2006 then %>
					<td class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Inquiries_CB") %><% Else %>0<% End If %>" name="frmSDMPerformanceYieldRate_Vol_Inquiries_CB" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_vol_yield_inq" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>				
					<td class="formMain" bgcolor="#c0c0c0">&nbsp;</td>	
				<% else %>
					<td class="formMain" align="center" colspan="2" width="100"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Inquiries_CB") %><% Else %>0<% End If %>" name="frmSDMPerformanceYieldRate_Vol_Inquiries_CB" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_vol_yield_inq" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>								
				<% end if %>		
				
				<% if y < 2006 then %>
					<td class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Inquiries_SB") %><% Else %>0<% End If %>" name="frmSDMPerformanceYieldRate_Vol_Inquiries_SB" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_vol_yield_inq" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>								
					<td class="formMain" bgcolor="#c0c0c0">&nbsp;</td>	
				<% else %>
					<td class="formMain" align="center" colspan="2" width="100"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Inquiries_SB") %><% Else %>0<% End If %>" name="frmSDMPerformanceYieldRate_Vol_Inquiries_SB" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_vol_yield_inq" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>												
				<% end if %>			
				
				<% if y < 2006 then %>
					<td class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Inquiries_OSB") %><% Else %>0<% End If %>" name="frmSDMPerformanceYieldRate_Vol_Inquiries_OSB" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_vol_yield_inq" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>												
					<td class="formMain" bgcolor="#c0c0c0">&nbsp;</td>				
				<% else %>
					<td class="formMain" align="center" colspan="2" width="100"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Inquiries_OSB") %><% Else %>0<% End If %>" name="frmSDMPerformanceYieldRate_Vol_Inquiries_OSB" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_vol_yield_inq" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>																
				<% end if %>
				
			</tr>	
			

			
			<tr>
				<td class="formMain">Volunteer Interviews</td>
				
				<% if y < 2006 then %>
					<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_Number_Comm") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Vol_InquiryToInterview_Number_Comm" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_vol_inq_intNUM_CB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_AveDays_Comm") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Vol_InquiryToInterview_AveDays_Comm" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_vol_inq_intAVG_CB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<% else %>
					<td class="formMain" colspan="2" align="center" width="100"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" colspan="2" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_Number_Comm") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Vol_InquiryToInterview_Number_Comm" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_vol_inq_intNUM_CB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceProcTim_Vol_InquiryToInterview_AveDays_Comm" onchange="checkForIntegerCommas(this.value);">
				<% end if %>
				
				<% if y < 2006 then %>
					<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_Number_School") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Vol_InquiryToInterview_Number_School" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_vol_inq_intNUM_SB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_AveDays_School") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Vol_InquiryToInterview_AveDays_School" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_vol_inq_intAVG_SB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<% else %>
					<td class="formMain" colspan="2" align="center" width="100"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_Number_School") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Vol_InquiryToInterview_Number_School" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_vol_inq_intNUM_SB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>				
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceProcTim_Vol_InquiryToInterview_AveDays_School" onchange="checkForIntegerCommas(this.value);">
				<% end if %>
				
				<% if y < 2006 then %>
					<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_Number_Other") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Vol_InquiryToInterview_Number_Other" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_vol_inq_intNUM_OSB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_AveDays_Other") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Vol_InquiryToInterview_AveDays_Other" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_vol_inq_intAVG_OSB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>				
				<% else %>
					<td class="formMain" colspan="2" align="center" width="100"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_Number_Other") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Vol_InquiryToInterview_Number_Other" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_vol_inq_intNUM_OSB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceProcTim_Vol_InquiryToInterview_AveDays_Other" onchange="checkForIntegerCommas(this.value);">
				<% end if %>
			</tr>

		<% if y < 2006 then %>
			<tr>
				<td class="formMain">Volunteer Interview <strong>to Matched</strong></td>
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InterviewToMatched_Number_Comm") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_Comm" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_vol_int_matchNUM_CB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InterviewToMatched_AveDays_Comm") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Vol_InterviewToMatched_AveDays_Comm" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_vol_int_matchAVG_CB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>

				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InterviewToMatched_Number_School") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_School" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_vol_int_matchNUM_SB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InterviewToMatched_AveDays_School") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Vol_InterviewToMatched_AveDays_School" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_vol_int_matchAVG_SB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>

				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InterviewToMatched_Number_Other") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_Other" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_vol_int_matchNUM_OSB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InterviewToMatched_AveDays_Other") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Vol_InterviewToMatched_AveDays_Other" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_vol_int_matchAVG_OSB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>

			</tr>
		
		<% else %>
			<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_Comm" onchange="checkForIntegerCommas(this.value);">
			<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceProcTim_Vol_InterviewToMatched_AveDays_Comm" onchange="checkForIntegerCommas(this.value);">
			<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_School" onchange="checkForIntegerCommas(this.value);">
			<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceProcTim_Vol_InterviewToMatched_AveDays_School" onchange="checkForIntegerCommas(this.value);">
			<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceProcTim_Vol_InterviewToMatched_Number_Other" onchange="checkForIntegerCommas(this.value);">
			<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceProcTim_Vol_InterviewToMatched_AveDays_Other" onchange="checkForIntegerCommas(this.value);">
		<% end if %>
			
			
			<tr>
				<td colspan="7" <% if y < 2006 then %>class="formMain"<%else%>class="formHeaderMedium"<%end if%> align="center"><strong>Child</strong></td>
			</tr>		
			
			<tr>
				<td>&nbsp;</td>
				<td colspan="2" class="formMain" align="center" <%if y>=2006 then%> width="100"<%end if%>>Community-Based</td>
				<td colspan="2" class="formMain" align="center" <%if y>=2006 then%> width="100"<%end if%>>School-Based</td>
				<td colspan="2" class="formMain" align="center" <%if y>=2006 then%> width="100"<%end if%>>Non-School<br>Site-Based</td>
			</tr>
			
			<tr>
				<td>&nbsp;</td>
				<% if y < 2006 then %>
					<td class="formMain">Number of Individuals</td>
					<td class="formMain">Average Days</td>
				<% else %>
					<td class="formMain" align="center" colspan="2" width="100">Number of Individuals</td>				
				<% end if %>

				<% if y < 2006 then %>
					<td class="formMain">Number of Individuals</td>
					<td class="formMain">Average Days</td>	
				<% else %>
					<td class="formMain" align="center" colspan="2" width="100">Number of Individuals</td>				
				<% end if %>

				<% if y < 2006 then %>
					<td class="formMain">Number of Individuals</td>
					<td class="formMain">Average Days</td>	
				<% else %>
					<td class="formMain" align="center" colspan="2" width="100">Number of Individuals</td>				
				<% end if %>
								
			</tr>
			

			
			<tr>
				<td class="formMain" align="left">Child Inquiries</td>
				<% if y < 2006 then %>
					<td class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Inquiries_CB") %><% Else %>0<% End If %>" name="frmSDMPerformanceYieldRate_Youth_Inquiries_CB" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_child_yield_inq" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>				
					<td class="formMain" bgcolor="#c0c0c0">&nbsp;</td>				
				<% else %>
					<td class="formMain" align="center" colspan="2" width="100"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Inquiries_CB") %><% Else %>0<% End If %>" name="frmSDMPerformanceYieldRate_Youth_Inquiries_CB" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_child_yield_inq" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>								
				<% end if %>
				
				<% if y < 2006 then %>
					<td class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Inquiries_SB") %><% Else %>0<% End If %>" name="frmSDMPerformanceYieldRate_Youth_Inquiries_SB" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_child_yield_inq" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>								
					<td class="formMain" bgcolor="#c0c0c0">&nbsp;</td>				
				<% else %>
					<td class="formMain" align="center" colspan="2" width="100"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Inquiries_SB") %><% Else %>0<% End If %>" name="frmSDMPerformanceYieldRate_Youth_Inquiries_SB" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_child_yield_inq" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>												
				<% end if %>
				
				<% if y < 2006 then %>
					<td class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Inquiries_OSB") %><% Else %>0<% End If %>" name="frmSDMPerformanceYieldRate_Youth_Inquiries_OSB" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_child_yield_inq" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>												
					<td class="formMain" bgcolor="#c0c0c0">&nbsp;</td>				
				<% else %>
					<td class="formMain" align="center" colspan="2" width="100"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Inquiries_OSB") %><% Else %>0<% End If %>" name="frmSDMPerformanceYieldRate_Youth_Inquiries_OSB" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_child_yield_inq" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>																
				<% end if %>
			</tr>			
			
			
			<tr>
				<td class="formMain">Child Inquiry Interviews</td>
				<% if y < 2006 then %>
					<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_Number_Comm") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Youth_InquiryToInterview_Number_Comm" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_child_inq_intNUM_CB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_AveDays_Comm") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Youth_InquiryToInterview_AveDays_Comm" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_child_inq_intAVG_CB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<% else %>
					<td class="formMain" align="center" colspan="2" width="100"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_Number_Comm") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Youth_InquiryToInterview_Number_Comm" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_child_inq_intNUM_CB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceProcTim_Youth_InquiryToInterview_AveDays_Comm" onchange="checkForIntegerCommas(this.value);">
				<% end if %>
				
				<% if y < 2006 then %>
					<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_Number_School") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Youth_InquiryToInterview_Number_School" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_child_inq_intNUM_SB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_AveDays_School") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Youth_InquiryToInterview_AveDays_School" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_child_inq_intAVG_SB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<% else %>
					<td class="formMain" align="center" colspan="2" width="100"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_Number_School") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Youth_InquiryToInterview_Number_School" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_child_inq_intNUM_SB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceProcTim_Youth_InquiryToInterview_AveDays_School" onchange="checkForIntegerCommas(this.value);">
				<% end if %>			
				
				<% if y < 2006 then %>	
					<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_Number_Other") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Youth_InquiryToInterview_Number_Other" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_child_inq_intNUM_OSB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_AveDays_Other") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Youth_InquiryToInterview_AveDays_Other" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_child_inq_intAVG_OSB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<% else %>
					<td class="formMain" align="center" colspan="2" width="100"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_Number_Other") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Youth_InquiryToInterview_Number_Other" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_child_inq_intNUM_OSB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceProcTim_Youth_InquiryToInterview_AveDays_Other" onchange="checkForIntegerCommas(this.value);">
				<% end if %>
				
			</tr>

			<% if y < 2006 then %>
				<tr>
					<td class="formMain">Child Interview <strong>to Matched</strong></td>
	
					<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InterviewToMatched_Number_Comm") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_Comm" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_child_int_matchNUM_CB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InterviewToMatched_AveDays_Comm") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Youth_InterviewToMatched_AveDays_Comm" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_child_int_matchAVG_CB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
	
					<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InterviewToMatched_Number_School") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_School" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_child_int_matchNUM_SB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InterviewToMatched_AveDays_School") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Youth_InterviewToMatched_AveDays_School" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_child_int_matchAVG_SB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					
					<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InterviewToMatched_Number_Other") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_Other" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_child_int_matchNUM_OSB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InterviewToMatched_AveDays_Other") %><% Else %>0<% End If %>" name="frmSDMPerformanceProcTim_Youth_InterviewToMatched_AveDays_Other" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_proc_child_int_matchAVG_OSB" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>				
					
				</tr>
			<% else %>
				<input type="hidden" colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_Comm" onchange="checkForIntegerCommas(this.value);">
				<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceProcTim_Youth_InterviewToMatched_AveDays_Comm" onchange="checkForIntegerCommas(this.value);">
				<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_School" onchange="checkForIntegerCommas(this.value);">
				<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceProcTim_Youth_InterviewToMatched_AveDays_School" onchange="checkForIntegerCommas(this.value);">
				<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceProcTim_Youth_InterviewToMatched_Number_Other" onchange="checkForIntegerCommas(this.value);">
				<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceProcTim_Youth_InterviewToMatched_AveDays_Other" onchange="checkForIntegerCommas(this.value);">
			<% end if %>
			
			
			
			<% if y < 2006 then %>
				<tr>
					<TD colspan="7" class="formHeaderMedium">NUMBER OF MATCH CLOSURES&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_freq_match_closures" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmark_purplesmall.gif" alt="" width="15" height="16" border="0"></a></TD>	
				</tr>
				
				<tr>		
					<td>&nbsp;</td>
					<td class="formMain" colspan="2" align="center">Community-Based</td>
					<td class="formMain" colspan="2" align="center">School-Based</td>			
					<td class="formMain" colspan="2" align="center">Non-School<br>Site-Based</td>		
				</tr>
				
				<tr>
					<td class="formMain">Less Than 3 Months</td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_Under3Months_Comm") %><% Else %>0<% End If %>" name="frmSDMPerformanceFreq_Under3Months_Comm" onchange="checkForIntegerCommas(this.value);"></td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_Under3Months_School") %><% Else %>0<% End If %>" name="frmSDMPerformanceFreq_Under3Months_School" onchange="checkForIntegerCommas(this.value);"></td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_Under3Months_Other") %><% Else %>0<% End If %>" name="frmSDMPerformanceFreq_Under3Months_Other" onchange="checkForIntegerCommas(this.value);"></td>				
				</tr>
				
				<tr>
					<td class="formMain">3-6 Months</td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_3To6Months_Comm") %><% Else %>0<% End If %>" name="frmSDMPerformanceFreq_3To6Months_Comm" onchange="checkForIntegerCommas(this.value);"></td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_3To6Months_School") %><% Else %>0<% End If %>" name="frmSDMPerformanceFreq_3To6Months_School" onchange="checkForIntegerCommas(this.value);"></td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_3To6Months_Other") %><% Else %>0<% End If %>" name="frmSDMPerformanceFreq_3To6Months_Other" onchange="checkForIntegerCommas(this.value);"></td>				
				</tr>		
				
				<tr>
					<td class="formMain">7-9 Months</td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_7To9Months_Comm") %><% Else %>0<% End If %>" name="frmSDMPerformanceFreq_7To9Months_Comm" onchange="checkForIntegerCommas(this.value);"></td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_7To9Months_School") %><% Else %>0<% End If %>" name="frmSDMPerformanceFreq_7To9Months_School" onchange="checkForIntegerCommas(this.value);"></td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_7To9Months_Other") %><% Else %>0<% End If %>" name="frmSDMPerformanceFreq_7To9Months_Other" onchange="checkForIntegerCommas(this.value);"></td>				
				</tr>		
				
				<tr>
					<td class="formMain">10-12 Months</td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_10To12Months_Comm") %><% Else %>0<% End If %>" name="frmSDMPerformanceFreq_10To12Months_Comm" onchange="checkForIntegerCommas(this.value);"></td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_10To12Months_School") %><% Else %>0<% End If %>" name="frmSDMPerformanceFreq_10To12Months_School" onchange="checkForIntegerCommas(this.value);"></td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_10To12Months_Other") %><% Else %>0<% End If %>" name="frmSDMPerformanceFreq_10To12Months_Other" onchange="checkForIntegerCommas(this.value);"></td>				
				</tr>
				
				<tr>
					<td class="formMain">13-23 Months</td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_13To23Months_Comm") %><% Else %>0<% End If %>" name="frmSDMPerformanceFreq_13To23Months_Comm" onchange="checkForIntegerCommas(this.value);"></td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_13To23Months_School") %><% Else %>0<% End If %>" name="frmSDMPerformanceFreq_13To23Months_School" onchange="checkForIntegerCommas(this.value);"></td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_13To23Months_Other") %><% Else %>0<% End If %>" name="frmSDMPerformanceFreq_13To23Months_Other" onchange="checkForIntegerCommas(this.value);"></td>				
				</tr>
				
				<tr>
					<td class="formMain">24 or More Months</td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_24OrMoreMonths_Comm") %><% Else %>0<% End If %>" name="frmSDMPerformanceFreq_24OrMoreMonths_Comm" onchange="checkForIntegerCommas(this.value);"></td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_24OrMoreMonths_School") %><% Else %>0<% End If %>" name="frmSDMPerformanceFreq_24OrMoreMonths_School" onchange="checkForIntegerCommas(this.value);"></td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_24OrMoreMonths_Other") %><% Else %>0<% End If %>" name="frmSDMPerformanceFreq_24OrMoreMonths_Other" onchange="checkForIntegerCommas(this.value);"></td>				
				</tr>
				
			<% else %>
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceFreq_Under3Months_Comm" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceFreq_Under3Months_School" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceFreq_Under3Months_Other" onchange="checkForIntegerCommas(this.value);">				
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceFreq_3To6Months_Comm" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceFreq_3To6Months_School" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceFreq_3To6Months_Other" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceFreq_7To9Months_Comm" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceFreq_7To9Months_School" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceFreq_7To9Months_Other" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceFreq_10To12Months_Comm" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceFreq_10To12Months_School" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceFreq_10To12Months_Other" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceFreq_13To23Months_Comm" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceFreq_13To23Months_School" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceFreq_13To23Months_Other" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceFreq_24OrMoreMonths_Comm" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceFreq_24OrMoreMonths_School" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceFreq_24OrMoreMonths_Other" onchange="checkForIntegerCommas(this.value);">
			<% end if %>
			
			
			<% if y < 2006 then %>
				<tr>
					<TD colspan="7" class="formHeaderMedium">VOLUNTEERS RE-MATCHED</TD>	
				</tr>		
				
				<tr>
					<td>&nbsp;</td>
					<td class="formMain" colspan="2" align="center">Community-Based</td>
					<td class="formMain" colspan="2" align="center">School-Based</td>				
					<td class="formMain" colspan="2" align="center">Non-School<br>Site-Based</td>				
				</tr>															
				
				<tr>
					<td class="formMain">Volunteers Re-Matched&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_vol_rematched" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Volunteers_ReMatchedCB") %><% Else %>0<% End If %>" name="frmSDMPerformanceVolunteers_ReMatchedCB" onchange="checkForIntegerCommas(this.value);"></td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Volunteers_ReMatchedSB") %><% Else %>0<% End If %>" name="frmSDMPerformanceVolunteers_ReMatchedSB" onchange="checkForIntegerCommas(this.value);"></td>				
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Volunteers_ReMatchedOSB") %><% Else %>0<% End If %>" name="frmSDMPerformanceVolunteers_ReMatchedOSB" onchange="checkForIntegerCommas(this.value);"></td>								
				</tr>
			
			<% else %>
				<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceVolunteers_ReMatchedCB" onchange="checkForIntegerCommas(this.value);">
				<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceVolunteers_ReMatchedSB" onchange="checkForIntegerCommas(this.value);">
				<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceVolunteers_ReMatchedOSB" onchange="checkForIntegerCommas(this.value);">
			<% end if %>
			
			<% if y < 2006 then %>
				
				<tr>
					<TD colspan="7" class="formHeaderMedium">PREMATURE CLOSURE</TD>
				</tr>
				
				<TR>
					<TD colspan="1">&nbsp;</TD>
					<TD colspan="2" class="formMain" align="center">Community-Based</TD>
					<TD colspan="2" class="formMain" align="center">School-Based</TD>
					<TD colspan="2" class="formMain" align="center">Non-School<br>Site-Based</TD>				
		
				</TR>
				
				<tr>
					<td colspan="1" class="formMain">Number of Matches that Closed Prematurely&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_premature_closure" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>			
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBNumberClosedPrematurely") %><% Else %>0<% End If %>" name="frmSDMPerformanceCBNumberClosedPrematurely" onchange="checkForIntegerCommas(this.value);"></td>				
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBNumberClosedPrematurely") %><% Else %>0<% End If %>" name="frmSDMPerformanceSBNumberClosedPrematurely" onchange="checkForIntegerCommas(this.value);"></td>								
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBNumberClosedPrematurely") %><% Else %>0<% End If %>" name="frmSDMPerformanceOSBNumberClosedPrematurely" onchange="checkForIntegerCommas(this.value);"></td>												
				</tr>
				
			<% else %>

				<input type="hidden" colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceCBNumberClosedPrematurely" onchange="checkForIntegerCommas(this.value);">
				<input type="hidden" colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceSBNumberClosedPrematurely" onchange="checkForIntegerCommas(this.value);">
				<input type="hidden" colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceOSBNumberClosedPrematurely" onchange="checkForIntegerCommas(this.value);">

			<% end if %>

			
			<% if y < 2006 then %>
			
				<tr>
					<TD colspan="7" class="formHeaderMedium">CLOSURE CODES</TD>	
				</tr>
				
				<TR>
					<TD colspan="1">&nbsp;</TD>
					<TD colspan="2" class="formMain" align="center">Community-Based</TD>
					<TD colspan="2" class="formMain" align="center">School-Based</TD>
					<TD colspan="2" class="formMain" align="center">Non-School<br>Site-Based</TD>				
				</TR>
					
				<tr>
					<td colspan="1" class="formMain">Child/Parent Status Change&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_closure_cpstatuschange" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBChildParentStatusChange") %><% Else %>0<% End If %>" name="frmSDMPerformanceCBChildParentStatusChange" onchange="checkForIntegerCommas(this.value);"></td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBChildParentStatusChange") %><% Else %>0<% End If %>" name="frmSDMPerformanceSBChildParentStatusChange" onchange="checkForIntegerCommas(this.value);"></td>				
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBChildParentStatusChange") %><% Else %>0<% End If %>" name="frmSDMPerformanceOSBChildParentStatusChange" onchange="checkForIntegerCommas(this.value);"></td>								
	
				</tr>
				
				<tr>
					<td colspan="1" class="formMain">Volunteer Status Change&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_closure_volstatuschange" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>	
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBVolunteerStatusChange") %><% Else %>0<% End If %>" name="frmSDMPerformanceCBVolunteerStatusChange" onchange="checkForIntegerCommas(this.value);"></td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBVolunteerStatusChange") %><% Else %>0<% End If %>" name="frmSDMPerformanceSBVolunteerStatusChange" onchange="checkForIntegerCommas(this.value);"></td>				
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBVolunteerStatusChange") %><% Else %>0<% End If %>" name="frmSDMPerformanceOSBVolunteerStatusChange" onchange="checkForIntegerCommas(this.value);"></td>								
				</tr>
				
				<tr>
					<td colspan="1" class="formMain">Child/Parent Dissatisfaction&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_closure_cpdissatisfaction" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>	
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBChildParentDissatisfaction") %><% Else %>0<% End If %>" name="frmSDMPerformanceCBChildParentDissatisfaction" onchange="checkForIntegerCommas(this.value);"></td>				
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBChildParentDissatisfaction") %><% Else %>0<% End If %>" name="frmSDMPerformanceSBChildParentDissatisfaction" onchange="checkForIntegerCommas(this.value);"></td>								
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBChildParentDissatisfaction") %><% Else %>0<% End If %>" name="frmSDMPerformanceOSBChildParentDissatisfaction" onchange="checkForIntegerCommas(this.value);"></td>												
				</tr>
				
				<tr>
	
					<td colspan="1" class="formMain">Volunteer Dissatisfaction&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_closure_voldissatisfaction" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBVolunteerDissatisfaction") %><% Else %>0<% End If %>" name="frmSDMPerformanceCBVolunteerDissatisfaction" onchange="checkForIntegerCommas(this.value);"></td>								
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBVolunteerDissatisfaction") %><% Else %>0<% End If %>" name="frmSDMPerformanceSBVolunteerDissatisfaction" onchange="checkForIntegerCommas(this.value);"></td>												
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBVolunteerDissatisfaction") %><% Else %>0<% End If %>" name="frmSDMPerformanceOSBVolunteerDissatisfaction" onchange="checkForIntegerCommas(this.value);"></td>																
				</tr>
				
			<tr>
	
					<td colspan="1" class="formMain">Successful Matches&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_successfulmatches" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBSuccessfulMatches") %><% Else %>0<% End If %>" name="frmSDMPerformanceCBSuccessfulMatches" onchange="checkForIntegerCommas(this.value);"></td>								
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBSuccessfulMatches") %><% Else %>0<% End If %>" name="frmSDMPerformanceSBSuccessfulMatches" onchange="checkForIntegerCommas(this.value);"></td>												
					<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBSuccessfulMatches") %><% Else %>0<% End If %>" name="frmSDMPerformanceOSBSuccessfulMatches" onchange="checkForIntegerCommas(this.value);"></td>																
			</tr>
			
		<% else %>
		
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceCBChildParentStatusChange" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceSBChildParentStatusChange" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceOSBChildParentStatusChange" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceCBVolunteerStatusChange" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceSBVolunteerStatusChange" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceOSBVolunteerStatusChange" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceCBChildParentDissatisfaction" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceSBChildParentDissatisfaction" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceOSBChildParentDissatisfaction" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceCBVolunteerDissatisfaction" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceSBVolunteerDissatisfaction" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceOSBVolunteerDissatisfaction" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceCBSuccessfulMatches" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceSBSuccessfulMatches" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceOSBSuccessfulMatches" onchange="checkForIntegerCommas(this.value);">
		<% end if %>
		
		
		<% if y < 2006 then %>	
			<tr>
				<TD colspan="7" class="formHeaderMedium">6-Month Retention</TD>	
			</tr>
			
			<TR>
				<TD colspan="1">&nbsp;</TD>
				<TD colspan="2" class="formMain" align="center">Community-Based</TD>
				<TD colspan="2" class="formMain" align="center">School-Based</TD>
				<TD colspan="2" class="formMain" align="center">Non-School<br>Site-Based</TD>				
			</TR>					
			
			
			<!-- Calculate Six Months Prior -->
			<% dim SixMonthsAgo
			SixMonthsAgo = m-6
			if SixmonthsAgo = -1 then
				SixMonthsAgo = 11
			else
				if SixMonthsAgo = -2 then
					SixMonthsAgo = 10
				else
					if SixMonthsAgo = -3 then
						SixMonthsAgo = 9
					else
						if SixMonthsAgo = -4 then
							SixMonthsAgo = 8
						else
							if SixMonthsAgo = -5 then
								SixMonthsAgo = 7
							else
								if SixMonthsAgo = 0 then
									SixMonthsAgo = 12
								end if
							end if
						end if
					end if 
				end if
			end if %>
			
			
			
			<tr>
			<td colspan="1" class="formMain">Number of <b>New </b>Matches Made in <strong><%=MonthName(SixMonthsAgo)%></strong>&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm_ret_new_matches_6months_ago&SixMonthsAgo=<%=SixMonthsAgo%>" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>				
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBTotalOpened6MonthsAgo") %><% Else %>0<% End If %>" name="frmSDMPerformanceCBTotalOpened6MonthsAgo" onchange="checkForIntegerCommas(this.value);"></td>								
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBTotalOpened6MonthsAgo") %><% Else %>0<% End If %>" name="frmSDMPerformanceSBTotalOpened6MonthsAgo" onchange="checkForIntegerCommas(this.value);"></td>												
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBTotalOpened6MonthsAgo") %><% Else %>0<% End If %>" name="frmSDMPerformanceOSBTotalOpened6MonthsAgo" onchange="checkForIntegerCommas(this.value);"></td>												
			</tr>	
			
			<tr>
				<td colspan="1" class="formMain">Number of These Matches That CLOSED Before the end of <strong><%=MonthName(m)%></strong> <strong><%=y%></strong>&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm16&SixMonthsAgo=<%=SixMonthsAgo%>&Now=<%=m%>" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBNumberStillOpen") %><% Else %>0<% End If %>" name="frmSDMPerformanceCBNumberStillOpen" onchange="checkForIntegerCommas(this.value);"></td>												
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBNumberStillOpen") %><% Else %>0<% End If %>" name="frmSDMPerformanceSBNumberStillOpen" onchange="checkForIntegerCommas(this.value);"></td>																
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBNumberStillOpen") %><% Else %>0<% End If %>" name="frmSDMPerformanceOSBNumberStillOpen" onchange="checkForIntegerCommas(this.value);"></td>														
			</tr>
			
		<% else %>
			<input type="hidden" colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceCBTotalOpened6MonthsAgo" onchange="checkForIntegerCommas(this.value);">
			<input type="hidden" colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceSBTotalOpened6MonthsAgo" onchange="checkForIntegerCommas(this.value);">
			<input type="hidden" colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceOSBTotalOpened6MonthsAgo" onchange="checkForIntegerCommas(this.value);">
			<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceCBNumberStillOpen" onchange="checkForIntegerCommas(this.value);">
			<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceSBNumberStillOpen" onchange="checkForIntegerCommas(this.value);">
			<input type="hidden"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="0" name="frmSDMPerformanceOSBNumberStillOpen" onchange="checkForIntegerCommas(this.value);">
		<% end if %>
			
			<!-- Ask Customer Satisfaction and POE Questions Quarterly -->
			
			<% if (m=3 or m=6 or m=9 or m=12) and y < 2005 then %>
			
				<tr>
					<TD colspan="7" class="formHeaderMedium">Customer Satisfaction</TD>	
				</tr>	
				
				<tr>
					<td colspan="1" class="formMain">Enrollment Satisfaction Average Score&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=cust_sat_1" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<td class="formMain" colspan="6" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("EnrollmentSatAvgScore") %><% Else %>0<% End If %>" name="frmSDMPerformanceEnrollmentSatAvgScore"></td>				
				</tr>				
				
				<tr>
					<td colspan="1" class="formMain">Enrollment Satisfaction Count&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=cust_sat_2" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<td class="formMain" colspan="6" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("EnrollmentSatCount") %><% Else %>0<% End If %>" name="frmSDMPerformanceEnrollmentSatCount" onchange="checkForIntegerCommas(this.value);"></td>								
				</tr>
	
				<tr>
					<td colspan="1" class="formMain">Match Satisfaction Average Score&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=cust_sat_3" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<td class="formMain" colspan="6" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("MatchSatAvgScore") %><% Else %>0<% End If %>" name="frmSDMPerformanceMatchSatAvgScore"></td>								
				</tr>			
				
				<tr>
					<td colspan="1" class="formMain">Match Satisfaction Count&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=cust_sat_4" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<td class="formMain" colspan="6" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("MatchSatCount") %><% Else %>0<% End If %>" name="frmSDMPerformanceMatchSatCount" onchange="checkForIntegerCommas(this.value);"></td>								
				</tr>	
				
				<tr>
					<TD colspan="7" class="formHeaderMedium">POE</TD>	
				</tr>								
				
				<tr>
					<td colspan="1">&nbsp;</td>
					<td colspan="2" class="formMain" align="center">Community-Based</td>
					<td colspan="2" class="formMain" align="center">School-Based</td>
					<td colspan="2" class="formMain" align="center">Non-School<br>Site-Based</td>				
				</tr>
				
				<tr>
					<td colspan="1" class="formMain">POE Aggregate Score&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=poe_aggregate" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
					<td colspan="2" class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBPOEAggregateScore") %><% Else %>0<% End If %>" name="frmSDMPerformanceCBPOEAggregateScore"></td>
					<td colspan="2" class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBPOEAggregateScore") %><% Else %>0<% End If %>" name="frmSDMPerformanceSBPOEAggregateScore"></td>				
					<td colspan="2" class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBPOEAggregateScore") %><% Else %>0<% End If %>" name="frmSDMPerformanceOSBPOEAggregateScore"></td>								
				</tr>
				
				<tr>
					<td colspan="1" class="formMain">POE Count&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=poe_count" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>	
					<td colspan="2" class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBPOECount") %><% Else %>0<% End If %>" name="frmSDMPerformanceCBPOECount" onchange="checkForIntegerCommas(this.value);"></td>
					<td colspan="2" class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBPOECount") %><% Else %>0<% End If %>" name="frmSDMPerformanceSBPOECount" onchange="checkForIntegerCommas(this.value);"></td>								
					<td colspan="2" class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBPOECount") %><% Else %>0<% End If %>" name="frmSDMPerformanceOSBPOECount" onchange="checkForIntegerCommas(this.value);"></td>												
				</tr>
				
			<% else %>
				<% if y < 2005 then %>
					<tr>
						<td colspan="7" class="formMain" align="center"><em><strong>Customer Satisfaction and POE Questions are answered Quarterly<br>(March, June, September, and December)</strong></em></td>
					</tr>
				<% else %>
					<tr>
						<td colspan="7" class="formMain" align="center"><em><strong>Starting in 2005, POE and Customer Satisfaction questions are no longer answered using this form.  Use the online POE and Customer Satisfaction Forms found <a href="http://agencies.bbbsa.org/myagency/POESat.asp">here</a></strong></em></td>
					</tr>				
				<% end if %>
				<!-- Prepopulate Fields to eliminate nulls and pass validation -->

				<input type="hidden"  value="0" name="frmSDMPerformanceEnrollmentSatAvgScore" onchange="checkForIntegerCommas(this.value);">
				<input type="hidden"  value="0" name="frmSDMPerformanceEnrollmentSatCount" onchange="checkForIntegerCommas(this.value);">
				<input type="hidden"  value="0" name="frmSDMPerformanceMatchSatAvgScore" onchange="checkForIntegerCommas(this.value);">
				<input type="hidden"  value="0" name="frmSDMPerformanceMatchSatCount" onchange="checkForIntegerCommas(this.value);">
				<input type="hidden"  value="0" name="frmSDMPerformanceCBPOEAggregateScore" onchange="checkForIntegerCommas(this.value);">
				<input type="hidden"  value="0" name="frmSDMPerformanceSBPOEAggregateScore" onchange="checkForIntegerCommas(this.value);">	
				<input type="hidden"  value="0" name="frmSDMPerformanceOSBPOEAggregateScore" onchange="checkForIntegerCommas(this.value);">
				<input type="hidden"  value="0" name="frmSDMPerformanceCBPOECount" onchange="checkForIntegerCommas(this.value);">
				<input type="hidden"  value="0" name="frmSDMPerformanceSBPOECount" onchange="checkForIntegerCommas(this.value);">							
				<input type="hidden"  value="0" name="frmSDMPerformanceOSBPOECount" onchange="checkForIntegerCommas(this.value);">										
				<input type="hidden" value=<%=m%> name="frmSDMPerformanceMonthValue" onchange="checkForIntegerCommas(this.value);">				
			<% end if %>
			



		<tr>
				<td colspan="7" class="formHeader">

				<input type="hidden" value=<%=m%> name="frmSDMPerformanceMonthValue" onchange="checkForIntegerCommas(this.value);">
				<input type="button" value="Save Form" class="formMainBold" onclick="validateForm(); return false;">


				</td>
			</tr>
			<tr>
			<td colspan="7"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
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

