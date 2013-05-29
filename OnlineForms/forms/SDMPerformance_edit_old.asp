<!--#include file="../includes/NAD_BE.asp" -->

<% 

' Check for SBM Agency
DIM SBMAgency
Set SBMCon = Server.CreateObject("ADODB.Connection")
SBMCon.Open "BBBSAforms","sa","12sist12"
query = "SELECT SBM FROM tbl_AgencyInfo WHERE AgencyID = '" & Session("AgencyIDN") & "' and SBM = -1  " 
Set SBMQuery = SBMCon.Execute(query)
if (SBMquery.eof) then
	SBMAgency = 0
else
	SBMAgency = 1
End if

' Check for SDM Agency
DIM SDMPilot
Set SDMCon = Server.CreateObject("ADODB.Connection")
SDMCon.Open "BBBSAForms","sa","12sist12"
query = "SELECT SDMPilot FROM tbl_AgencyInfo WHERE AgencyID = '" & Session("AgencyIDN") & "' and SDMPilot = -1 "
Set SDMQuery = SDMCon.Execute(query)
if (SDMquery.eof) then 
	SDMPilot = 0
else
	SDMPilot = 1
End if
	
SBMQuery.Close
Set SBMQuery = Nothing
SBMCon.Close
Set SBMCon = Nothing

SDMQuery.Close
Set SDMQuery = Nothing
SDMCon.Close
Set SDMCon = Nothing

' Check for Faith-Based / Incarcerated Agency
Dim FBIAgency
Set FBICon = Server.CreateObject("ADODB.Connection")
FBICon.Open "BBBSAforms","sa","12sist12"
query = "SELECT FBI FROM tbl_AgencyInfo WHERE AgencyID = '" & Session("AgencyIDN") & "' and FBI = -1 "
Set FBIQuery = FBICon.Execute(query)
if (FBIQuery.eof) then
	FBIAgency = 0
else
	FBIAgency = 1
End If

FBIQuery.Close
Set FBIQuery = Nothing
FBICon.Close
Set FBICon = Nothing





If Request("status") = "addNew" Then

	
	
' Check for duplicate records
	
	Set DupCon = Server.CreateObject("ADODB.Connection")
	DupCon.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT Count(*) As NumberOfEntries FROM tbl_frmPerformance WHERE AgencyID = '" & Request("AgencyIDN") & "' and Year = " & Request("Year")	& " and Month = " & Request("Month")
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
		RST.Open "SELECT * FROM tbl_frmPerformance", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = Request("AgencyIDN")
		RST("Year") = Request("year")
		RST("Month") = Request("month")
		RST("OpenMatchesCommunityBased") = Request("frmPerformanceOpenMatchesCommunityBased")
		RST("OpenMatchesSchoolBased") = Request("frmPerformanceOpenMatchesSchoolBased")
		RST("OpenMatchesOtherSiteBased") = Request("frmPerformanceOpenMatchesOtherSiteBased")
		RST("OpenMatchesGroupMentoring") = Request("frmPerformanceOpenMatchesGroupMentoring")
		RST("OpenMatchesSpecialProgramsMentoring") = Request("frmPerformanceOpenMatchesSpecialProgramsMentoring")
		RST("OpenMatchesSpecialProgramsNonMentoring") = Request("frmPerformanceOpenMatchesSpecialProgramsNonMentoring")
		RST("ClosedMatchesCommunityBased") = Request("frmPerformanceClosedMatchesCommunityBased")
		RST("ClosedMatchesSchoolBased") = Request("frmPerformanceClosedMatchesSchoolBased")
		RST("ClosedMatchesOtherSiteBased") = Request("frmPerformanceClosedMatchesOtherSiteBased")
		RST("ClosedMatchesGroupMentoring") = Request("frmPerformanceClosedMatchesGroupMentoring")
		RST("ClosedMatchesSpecialProgramsMentoring") = Request("frmPerformanceClosedMatchesSpecialProgramsMentoring")
		RST("ClosedMatchesSpecialProgramsNonMentoring") = Request("frmPerformanceClosedMatchesSpecialProgramsNonMentoring")
		
		RST("NewMatchesCommunityBased") = Request("frmPerformanceNewMatchesCommunityBased")
		RST("NewMatchesSchoolBased") = Request("frmPerformanceNewMatchesSchoolBased")
		RST("NewMatchesSiteBasedNonSchool") = Request("frmPerformanceNewMatchesSiteBasedNonSchool")
		RST("NewMatchesGroupMentoring") = Request("frmPerformanceNewMatchesGroupMentoring")		
		RST("NewMatchesSpecialProgramsMentoring") = Request("frmPerformanceNewMatchesSpecialProgramsMentoring")				
		RST("NewMatchesSpecialProgramsNonMentoring") = Request("frmPerformanceNewMatchesSpecialProgramsNonMentoring")			
		
		RST("AverageMatchLengthCB") = Request("frmPerformanceAverageMatchLengthCB")
		RST("AverageMatchLengthSB") = Request("frmPerformanceAverageMatchLengthSB")
		RST("AverageMatchLengthOSB") = Request("frmPerformanceAverageMatchLengthOSB")
		RST("Revenue") = Request("frmPerformanceRevenue")
		RST("AlphaCommunityBased") = Request("frmPerformanceAlphaCommunityBased")
		RST("AlphaSchoolBased") = Request("frmPerformanceAlphaSchoolBased")
		RST("AlphaOtherSiteBased") = Request("frmPerformanceAlphaOtherSiteBased")
		RST("AlphaNotPartnering") = Request("frmPerformanceAlphaNotPartnering")
		RST("Alphainterest") = Request("frmPerformanceAlphainterest")		
		RST("LionsCommunityBased") = Request("frmPerformanceLionsCommunityBased")
		RST("LionsSchoolBased") = Request("frmPerformanceLionsSchoolBased")
		RST("LionsOtherSiteBased") = Request("frmPerformanceLionsOtherSiteBased")
		RST("LionsNotPartnering") = Request("frmPerformanceLionsNotPartnering")
		RST("Lionsinterest") = Request("frmPerformanceLionsinterest")	
		RST("RotaryCommunityBased") = Request("frmPerformanceRotaryCommunityBased")
		RST("RotarySchoolBased") = Request("frmPerformanceRotarySchoolBased")
		RST("RotaryOtherSiteBased") = Request("frmPerformanceRotaryOtherSiteBased")
		RST("RotaryNotPartnering") = Request("frmPerformanceRotaryNotPartnering")
		RST("RotaryInterest") = Request("frmPerformanceRotaryInterest")	
		RST("KiwanisCommunityBased") = Request("frmPerformanceKiwanisCommunityBased")
		RST("KiwanisSchoolBased") = Request("frmPerformanceKiwanisSchoolBased")
		RST("KiwanisOtherSiteBased") = Request("frmPerformanceKiwanisOtherSiteBased")
		RST("KiwanisNotPartnering") = Request("frmPerformanceKiwanisNotPartnering")
		RST("KiwanisInterest") = Request("frmPerformanceKiwanisInterest")	
		RST("OptimistCommunityBased") = Request("frmPerformanceOptimistCommunityBased")
		RST("OptimistSchoolBased") = Request("frmPerformanceOptimistSchoolBased")
		RST("OptimistOtherSiteBased") = Request("frmPerformanceOptimistOtherSiteBased")
		RST("OptimistNotPartnering") = Request("frmPerformanceOptimistNotPartnering")
		RST("OptimistInterest") = Request("frmPerformanceOptimistInterest")		
		RST("AARPCommunityBased") = Request("frmPerformanceAARPCommunityBased")
		RST("AARPSchoolBased") = Request("frmPerformanceAARPSchoolBased")
		RST("AARPOtherSiteBased") = Request("frmPerformanceAARPOtherSiteBased")
		RST("AARPNotPartnering") = Request("frmPerformanceAARPNotPartnering")
		RST("AARPInterest") = Request("frmPerformanceAARPInterest")		
		RST("AlphaRating") = Request("frmPerformanceAlphaRating")				
		RST("LionsRating") = Request("frmPerformanceLionsRating")			
		RST("RotaryRating") = Request("frmPerformanceRotaryRating")			
		RST("KiwanisRating") = Request("frmPerformanceKiwanisRating")			
		RST("OptimistRating") = Request("frmPerformanceOptimistRating")			
		RST("AARPRating") = Request("frmPerformanceAARPRating")			
		RST("AlphaFunding") = Request("frmPerformanceAlphaFunding")			
		RST("AlphaProgramInitiative") = Request("frmPerformanceAlphaProgramInitiative")			
		RST("AlphaLeadershipInvolvement") = Request("frmPerformanceAlphaLeadershipInvolvement")		
		RST("AlphaUndergradChapterName") = Request("frmPerformanceAlphaUndergradChapterName")		
		RST("AlphaUndergradChapterCity") = Request("frmPerformanceAlphaUndergradChapterCity")
		RST("AlphaUndergradChapterState") = Request("frmPerformanceAlphaUndergradChapterState")
		RST("AlphaAlumniChapterName") = Request("frmPerformanceAlphaAlumniChapterName")		
		RST("AlphaAlumniChapterCity") = Request("frmPerformanceAlphaAlumniChapterCity")
		RST("AlphaAlumniChapterState") = Request("frmPerformanceAlphaAlumniChapterState")	
		
		' SDM Metrics Fields
		
		RST("YieldRate_Vol_Inquiries") = Request("frmPerformanceYieldRate_Vol_Inquiries")
		RST("YieldRate_Vol_Interviews") = Request("frmPerformanceYieldRate_Vol_Interviews")
		RST("YieldRate_Vol_Matched") = Request("frmPerformanceYieldRate_Vol_Matched")
		RST("YieldRate_Youth_Inquiries") = Request("frmPerformanceYieldRate_Youth_Inquiries")
		RST("YieldRate_Youth_Interviews") = Request("frmPerformanceYieldRate_Youth_Interviews")
		RST("YieldRate_Youth_Matched") = Request("frmPerformanceYieldRate_Youth_Matched")
		RST("ProcTim_Vol_InquiryToInterview_Number_Comm") = Request("frmPerformanceProcTim_Vol_InquiryToInterview_Number_Comm")
		RST("ProcTim_Vol_InquiryToInterview_AveDays_Comm") = Request("frmPerformanceProcTim_Vol_InquiryToInterview_AveDays_Comm")
		RST("ProcTim_Vol_InquiryToInterview_Number_School") = Request("frmPerformanceProcTim_Vol_InquiryToInterview_Number_School")
		RST("ProcTim_Vol_InquiryToInterview_AveDays_School") = Request("frmPerformanceProcTim_Vol_InquiryToInterview_AveDays_School")
		RST("ProcTim_Vol_InquiryToInterview_Number_Other") = Request("frmPerformanceProcTim_Vol_InquiryToInterview_Number_Other")
		RST("ProcTim_Vol_InquiryToInterview_AveDays_Other") = Request("frmPerformanceProcTim_Vol_InquiryToInterview_AveDays_Other")
		RST("ProcTim_Vol_InterviewToMatched_Number_Comm") = Request("frmPerformanceProcTim_Vol_InterviewToMatched_Number_Comm")
		RST("ProcTim_Vol_InterviewToMatched_AveDays_Comm") = Request("frmPerformanceProcTim_Vol_InterviewToMatched_AveDays_Comm")
		RST("ProcTim_Vol_InterviewToMatched_Number_School") = Request("frmPerformanceProcTim_Vol_InterviewToMatched_Number_School")
		RST("ProcTim_Vol_InterviewToMatched_AveDays_School") = Request("frmPerformanceProcTim_Vol_InterviewToMatched_AveDays_School")
		RST("ProcTim_Vol_InterviewToMatched_Number_Other") = Request("frmPerformanceProcTim_Vol_InterviewToMatched_Number_Other")
		RST("ProcTim_Vol_InterviewToMatched_AveDays_Other") = Request("frmPerformanceProcTim_Vol_InterviewToMatched_AveDays_Other")
		RST("ProcTim_Youth_InquiryToInterview_Number_Comm") = Request("frmPerformanceProcTim_Youth_InquiryToInterview_Number_Comm")
		RST("ProcTim_Youth_InquiryToInterview_AveDays_Comm") = Request("frmPerformanceProcTim_Youth_InquiryToInterview_AveDays_Comm")
		RST("ProcTim_Youth_InquiryToInterview_Number_School") = Request("frmPerformanceProcTim_Youth_InquiryToInterview_Number_School")
		RST("ProcTim_Youth_InquiryToInterview_AveDays_School") = Request("frmPerformanceProcTim_Youth_InquiryToInterview_AveDays_School")
		RST("ProcTim_Youth_InquiryToInterview_Number_Other") = Request("frmPerformanceProcTim_Youth_InquiryToInterview_Number_Other")
		RST("ProcTim_Youth_InquiryToInterview_AveDays_Other") = Request("frmPerformanceProcTim_Youth_InquiryToInterview_AveDays_Other")
		RST("ProcTim_Youth_InterviewToMatched_Number_Comm") = Request("frmPerformanceProcTim_Youth_InterviewToMatched_Number_Comm")
		RST("ProcTim_Youth_InterviewToMatched_AveDays_Comm") = Request("frmPerformanceProcTim_Youth_InterviewToMatched_AveDays_Comm")
		RST("ProcTim_Youth_InterviewToMatched_Number_School") = Request("frmPerformanceProcTim_Youth_InterviewToMatched_Number_School")
		RST("ProcTim_Youth_InterviewToMatched_AveDays_School") = Request("frmPerformanceProcTim_Youth_InterviewToMatched_AveDays_School")
		RST("ProcTim_Youth_InterviewToMatched_Number_Other") = Request("frmPerformanceProcTim_Youth_InterviewToMatched_Number_Other")
		RST("ProcTim_Youth_InterviewToMatched_AveDays_Other") = Request("frmPerformanceProcTim_Youth_InterviewToMatched_AveDays_Other")

		RST("Freq_Under3Months_Comm") = Request("frmPerformanceFreq_Under3Months_Comm")
		RST("Freq_Under3Months_School") = Request("frmPerformanceFreq_Under3Months_School")
		RST("Freq_Under3Months_Other") = Request("frmPerformanceFreq_Under3Months_Other")		
		RST("Freq_3To5Months_Comm") = Request("frmPerformanceFreq_3To5Months_Comm")
		RST("Freq_3To5Months_School") = Request("frmPerformanceFreq_3To5Months_School")
		RST("Freq_3To5Months_Other") = Request("frmPerformanceFreq_3To5Months_Other")		
		RST("Freq_6To8Months_Comm") = Request("frmPerformanceFreq_6To8Months_Comm")
		RST("Freq_6To8Months_School") = Request("frmPerformanceFreq_6To8Months_School")
		RST("Freq_6To8Months_Other") = Request("frmPerformanceFreq_6To8Months_Other")		
		RST("Freq_9To11Months_Comm") = Request("frmPerformanceFreq_9To11Months_Comm")
		RST("Freq_9To11Months_School") = Request("frmPerformanceFreq_9To11Months_School")
		RST("Freq_9To11Months_Other") = Request("frmPerformanceFreq_9To11Months_Other")		
		RST("Freq_12To23Months_Comm") = Request("frmPerformanceFreq_12To23Months_Comm")
		RST("Freq_12To23Months_School") = Request("frmPerformanceFreq_12To23Months_School")
		RST("Freq_12To23Months_Other") = Request("frmPerformanceFreq_12To23Months_Other")		
		RST("Freq_24OrMoreMonths_Comm") = Request("frmPerformanceFreq_24OrMoreMonths_Comm")
		RST("Freq_24OrMoreMonths_School") = Request("frmPerformanceFreq_24OrMoreMonths_School")
		RST("Freq_24OrMoreMonths_Other") = Request("frmPerformanceFreq_24OrMoreMonths_Other")		
		RST("Volunteers_ReMatched") = Request("frmPerformanceVolunteers_ReMatched")
		RST("POE_Confidence_Number_Comm") = Request("frmPerformancePOE_Confidence_Number_Comm")
		RST("POE_Confidence_Ave_Comm") = Request("frmPerformancePOE_Confidence_Ave_Comm")
		RST("POE_Confidence_Number_School") = Request("frmPerformancePOE_Confidence_Number_School")
		RST("POE_Confidence_Ave_School") = Request("frmPerformancePOE_Confidence_Ave_School")
		RST("POE_Competence_Number_Comm") = Request("frmPerformancePOE_Competence_Number_Comm")
		RST("POE_Competence_Ave_Comm") = Request("frmPerformancePOE_Competence_Ave_Comm")
		RST("POE_Competence_Number_School") = Request("frmPerformancePOE_Competence_Number_School")
		RST("POE_Competence_Ave_School") = Request("frmPerformancePOE_Competence_Ave_School")
		RST("POE_Caring_Number_Comm") = Request("frmPerformancePOE_Caring_Number_Comm")
		RST("POE_Caring_Ave_Comm") = Request("frmPerformancePOE_Caring_Ave_Comm")
		RST("POE_Caring_Number_School") = Request("frmPerformancePOE_Caring_Number_School")
		RST("POE_Caring_Ave_School") = Request("frmPerformancePOE_Caring_Ave_School")
		RST("VolSat_PostEnrollment_Number") = Request("frmPerformanceVolSat_PostEnrollment_Number")
		RST("VolSat_PostEnrollment_Ave") = Request("frmPerformanceVolSat_PostEnrollment_Ave")
		RST("VolSat_SatQuest_Number") = Request("frmPerformanceVolSat_SatQuest_Number")
		RST("VolSat_SatQuest_Ave") = Request("frmPerformanceVolSat_SatQuest_Ave")

		RST("CBNumberClosedPrematurely") = Request("frmPerformanceCBNumberClosedPrematurely")
		RST("SBNumberClosedPrematurely") = Request("frmPerformanceSBNumberClosedPrematurely")				
		RST("CBChildParentStatusChange") = Request("frmPerformanceCBChildParentStatusChange")
		RST("CBVolunteerStatusChange") = Request("frmPerformanceCBVolunteerStatusChange")
		RST("CBChildParentDissatisfaction") = Request("frmPerformanceCBChildParentDissatisfaction")
		RST("CBVolunteerDissatisfaction") = Request("frmPerformanceCBVolunteerDissatisfaction")
		RST("SBChildParentStatusChange") = Request("frmPerformanceSBChildParentStatusChange")
		RST("SBVolunteerStatusChange") = Request("frmPerformanceSBVolunteerStatusChange")
		RST("SBChildParentDissatisfaction") = Request("frmPerformanceSBChildParentDissatisfaction")
		RST("SBVolunteerDissatisfaction") = Request("frmPerformanceSBVolunteerDissatisfaction")
		RST("CBTotalOpened6MonthsAgo") = Request("frmPerformanceCBTotalOpened6MonthsAgo")
		RST("CBNumberStillOpen") = Request("frmPerformanceCBNumberStillOpen")
		RST("EnrollmentSatAvgScore") = Request("frmPerformanceEnrollmentSatAvgScore")
		RST("EnrollmentSatCount") = Request("frmPerformanceEnrollmentSatCount")
		RST("MatchSatAvgScore") = Request("frmPerformanceMatchSatAvgScore")
		RST("MatchSatCount") = Request("frmPerformanceMatchSatCount")
		RST("CBPOEAggregateScore") = Request("frmPerformanceCBPOEAggregateScore")
		RST("CBPOECount") = Request("frmPerformanceCBPOECount")
		RST("SBPOEAggregateScore") = Request("frmPerformanceSBPOEAggregateScore")
		RST("SBPOECount") = Request("frmPerformanceSBPOECount")
		
		
		' Additional SDM fields
		
		RST("YieldRate_Vol_Inquiries_CB") = Request("frmPerformanceYieldRate_Vol_Inquiries_CB")
		RST("YieldRate_Vol_Inquiries_SB") = Request("frmPerformanceYieldRate_Vol_Inquiries_SB")
		RST("YieldRate_Vol_Inquiries_OSB") = Request("frmPerformanceYieldRate_Vol_Inquiries_OSB")
		RST("YieldRate_Vol_Interviews_CB") = Request("frmPerformanceYieldRate_Vol_Interviews_CB")
		RST("YieldRate_Vol_Interviews_SB") = Request("frmPerformanceYieldRate_Vol_Interviews_SB")		
		RST("YieldRate_Vol_Interviews_OSB") = Request("frmPerformanceYieldRate_Vol_Interviews_OSB")		
		RST("YieldRate_Vol_Matched_CB") = Request("frmPerformanceYieldRate_Vol_Matched_CB")
		RST("YieldRate_Vol_Matched_SB") = Request("frmPerformanceYieldRate_Vol_Matched_SB")		
		RST("YieldRate_Vol_Matched_OSB") = Request("frmPerformanceYieldRate_Vol_Matched_OSB")		
		RST("YieldRate_Youth_Inquiries_CB") = Request("frmPerformanceYieldRate_Youth_Inquiries_CB")
		RST("YieldRate_Youth_Inquiries_SB") = Request("frmPerformanceYieldRate_Youth_Inquiries_SB")
		RST("YieldRate_Youth_Inquiries_OSB") = Request("frmPerformanceYieldRate_Youth_Inquiries_OSB")
		RST("YieldRate_Youth_Interviews_CB") = Request("frmPerformanceYieldRate_Youth_Interviews_CB")
		RST("YieldRate_Youth_Interviews_SB") = Request("frmPerformanceYieldRate_Youth_Interviews_SB")		
		RST("YieldRate_Youth_Interviews_OSB") = Request("frmPerformanceYieldRate_Youth_Interviews_OSB")		
		RST("YieldRate_Youth_Matched_CB") = Request("frmPerformanceYieldRate_Youth_Matched_CB")
		RST("YieldRate_Youth_Matched_SB") = Request("frmPerformanceYieldRate_Youth_Matched_SB")		
		RST("YieldRate_Youth_Matched_OSB") = Request("frmPerformanceYieldRate_Youth_Matched_OSB")	
		RST("OSBNumberClosedPrematurely") = Request("frmPerformanceOSBNumberClosedPrematurely")
		RST("OSBChildParentStatusChange") = Request("frmPerformanceOSBChildParentStatusChange")
		RST("OSBVolunteerStatusChange") = Request("frmPerformanceOSBVolunteerStatusChange")
		RST("OSBChildParentDissatisfaction") = Request("frmPerformanceOSBChildParentDissatisfaction")
		RST("OSBVolunteerDissatisfaction") = Request("frmPerformanceOSBVolunteerDissatisfaction")
		RST("SBTotalOpened6MonthsAgo") = Request("frmPerformanceSBTotalOpened6MonthsAgo")
		RST("SBNumberStillOpen") = Request("frmPerformanceSBNumberStillOpen")
		RST("OSBTotalOpened6MonthsAgo") = Request("frmPerformanceOSBTotalOpened6MonthsAgo")
		RST("OSBNumberStillOpen") = Request("frmPerformanceOSBNumberStillOpen")																	
		RST("OSBPOEAggregateScore") = Request("frmPerformanceOSBPOEAggregateScore")
		RST("OSBPOECount") = Request("frmPerformanceOSBPOECount")		
		
		
		
		
		' RTBM Fields
		
		If Int(Request("month"))=12 then
			RST("RTBM_UnmatchedChildren") = Request("frmPerformanceRTBM_UnmatchedChildren")
			RST("RTBM_UnmatchedVolunteers") = Request("frmPerformanceRTBM_UnmatchedVolunteers")
		End If
	
		' SBM Fields				
		
		If (Int(Request("month"))=6 or Int(Request("month"))=12) And SBMAgency = 1  Then 			
			RST("SBMVolunteersInEnrollmentProcess") = Request("frmPerformanceSBMVolunteersInEnrollmentProcess")
			RST("SBMAmountRaisedTowardsMatchPledge") = Request("frmPerformanceSBMAmountRaisedTowardsMatchPledge")
		End If

		If FBIAgency = 1 then		
			RST("CBIandFB") = Request("frmPerformanceCBIandFB")
			RST("CBInotFB") = Request("frmPerformanceCBInotFB")
			RST("CBFBnotI") = Request("frmPerformanceCBFBnotI")		
			RST("SBIandFB") = Request("frmPerformanceSBIandFB")
			RST("SBInotFB") = Request("frmPerformanceSBInotFB")
			RST("SBFBnotI") = Request("frmPerformanceSBFBnotI")				
			RST("OSBIandFB") = Request("frmPerformanceOSBIandFB")
			RST("OSBInotFB") = Request("frmPerformanceOSBInotFB")
			RST("OSBFBnotI") = Request("frmPerformanceOSBFBnotI")		
			RST("GMIandFB") = Request("frmPerformanceGMIandFB")
			RST("GMInotFB") = Request("frmPerformanceGMInotFB")
			RST("GMFBnotI") = Request("frmPerformanceGMFBnotI")				
			RST("SPMIandFB") = Request("frmPerformanceSPMIandFB")
			RST("SPMInotFB") = Request("frmPerformanceSPMInotFB")
			RST("SPMFBnotI") = Request("frmPerformanceSPMFBnotI")				
			RST("SPNMIandFB") = Request("frmPerformanceSPNMIandFB")
			RST("SPNMInotFB") = Request("frmPerformanceSPNMInotFB")
			RST("SPNMFBnotI") = Request("frmPerformanceSPNMFBnotI")			
		End If	
		
		RST("CreateDate") = Now
		RST.Update
		RST.Close
		Set RST = Nothing
		form = "Performance"
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
	RST.Open "SELECT * FROM tbl_frmPerformance WHERE agencyID='" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")) & " AND Month=" & Int(Request("month")), Con, 1, 3
	RST("OpenMatchesCommunityBased") = Request("frmPerformanceOpenMatchesCommunityBased")
	RST("OpenMatchesSchoolBased") = Request("frmPerformanceOpenMatchesSchoolBased")
	RST("OpenMatchesOtherSiteBased") = Request("frmPerformanceOpenMatchesOtherSiteBased")
	RST("OpenMatchesGroupMentoring") = Request("frmPerformanceOpenMatchesGroupMentoring")
	RST("OpenMatchesSpecialProgramsMentoring") = Request("frmPerformanceOpenMatchesSpecialProgramsMentoring")
	RST("OpenMatchesSpecialProgramsNonMentoring") = Request("frmPerformanceOpenMatchesSpecialProgramsNonMentoring")
	RST("ClosedMatchesCommunityBased") = Request("frmPerformanceClosedMatchesCommunityBased")
	RST("ClosedMatchesSchoolBased") = Request("frmPerformanceClosedMatchesSchoolBased")
	RST("ClosedMatchesOtherSiteBased") = Request("frmPerformanceClosedMatchesOtherSiteBased")
	RST("ClosedMatchesGroupMentoring") = Request("frmPerformanceClosedMatchesGroupMentoring")
	RST("ClosedMatchesSpecialProgramsMentoring") = Request("frmPerformanceClosedMatchesSpecialProgramsMentoring")
	RST("ClosedMatchesSpecialProgramsNonMentoring") = Request("frmPerformanceClosedMatchesSpecialProgramsNonMentoring")
	
	RST("NewMatchesCommunityBased") = Request("frmPerformanceNewMatchesCommunityBased")
	RST("NewMatchesSchoolBased") = Request("frmPerformanceNewMatchesSchoolBased")
	RST("NewMatchesSiteBasedNonSchool") = Request("frmPerformanceNewMatchesSiteBasedNonSchool")
	RST("NewMatchesGroupMentoring") = Request("frmPerformanceNewMatchesGroupMentoring")		
	RST("NewMatchesSpecialProgramsMentoring") = Request("frmPerformanceNewMatchesSpecialProgramsMentoring")				
	RST("NewMatchesSpecialProgramsNonMentoring") = Request("frmPerformanceNewMatchesSpecialProgramsNonMentoring")	
	
	RST("AverageMatchLengthCB") = Request("frmPerformanceAverageMatchLengthCB")
	RST("AverageMatchLengthSB") = Request("frmPerformanceAverageMatchLengthSB")
	RST("AverageMatchLengthOSB") = Request("frmPerformanceAverageMatchLengthOSB")	
	RST("Revenue") = Request("frmPerformanceRevenue")
	RST("AlphaCommunityBased") = Request("frmPerformanceAlphaCommunityBased")
	RST("AlphaSchoolBased") = Request("frmPerformanceAlphaSchoolBased")
	RST("AlphaOtherSiteBased") = Request("frmPerformanceAlphaOtherSiteBased")
	RST("AlphaNotPartnering") = Request("frmPerformanceAlphaNotPartnering")
	RST("Alphainterest") = Request("frmPerformanceAlphainterest")		
	RST("LionsCommunityBased") = Request("frmPerformanceLionsCommunityBased")
	RST("LionsSchoolBased") = Request("frmPerformanceLionsSchoolBased")
	RST("LionsOtherSiteBased") = Request("frmPerformanceLionsOtherSiteBased")
	RST("LionsNotPartnering") = Request("frmPerformanceLionsNotPartnering")
	RST("Lionsinterest") = Request("frmPerformanceLionsinterest")	
	RST("RotaryCommunityBased") = Request("frmPerformanceRotaryCommunityBased")
	RST("RotarySchoolBased") = Request("frmPerformanceRotarySchoolBased")
	RST("RotaryOtherSiteBased") = Request("frmPerformanceRotaryOtherSiteBased")
	RST("RotaryNotPartnering") = Request("frmPerformanceRotaryNotPartnering")
	RST("RotaryInterest") = Request("frmPerformanceRotaryInterest")	
	RST("KiwanisCommunityBased") = Request("frmPerformanceKiwanisCommunityBased")
	RST("KiwanisSchoolBased") = Request("frmPerformanceKiwanisSchoolBased")
	RST("KiwanisOtherSiteBased") = Request("frmPerformanceKiwanisOtherSiteBased")
	RST("KiwanisNotPartnering") = Request("frmPerformanceKiwanisNotPartnering")
	RST("KiwanisInterest") = Request("frmPerformanceKiwanisInterest")	
	RST("OptimistCommunityBased") = Request("frmPerformanceOptimistCommunityBased")
	RST("OptimistSchoolBased") = Request("frmPerformanceOptimistSchoolBased")
	RST("OptimistOtherSiteBased") = Request("frmPerformanceOptimistOtherSiteBased")
	RST("OptimistNotPartnering") = Request("frmPerformanceOptimistNotPartnering")
	RST("OptimistInterest") = Request("frmPerformanceOptimistInterest")		
	RST("AARPCommunityBased") = Request("frmPerformanceAARPCommunityBased")
	RST("AARPSchoolBased") = Request("frmPerformanceAARPSchoolBased")
	RST("AARPOtherSiteBased") = Request("frmPerformanceAARPOtherSiteBased")
	RST("AARPNotPartnering") = Request("frmPerformanceAARPNotPartnering")
	RST("AARPInterest") = Request("frmPerformanceAARPInterest")		
	RST("AlphaRating") = Request("frmPerformanceAlphaRating")				
	RST("LionsRating") = Request("frmPerformanceLionsRating")			
	RST("RotaryRating") = Request("frmPerformanceRotaryRating")			
	RST("KiwanisRating") = Request("frmPerformanceKiwanisRating")			
	RST("OptimistRating") = Request("frmPerformanceOptimistRating")			
	RST("AARPRating") = Request("frmPerformanceAARPRating")			
	RST("AlphaFunding") = Request("frmPerformanceAlphaFunding")			
	RST("AlphaProgramInitiative") = Request("frmPerformanceAlphaProgramInitiative")			
	RST("AlphaLeadershipInvolvement") = Request("frmPerformanceAlphaLeadershipInvolvement")		
	RST("AlphaUndergradChapterName") = Request("frmPerformanceAlphaUndergradChapterName")		
	RST("AlphaUndergradChapterCity") = Request("frmPerformanceAlphaUndergradChapterCity")
	RST("AlphaUndergradChapterState") = Request("frmPerformanceAlphaUndergradChapterState")
	RST("AlphaAlumniChapterName") = Request("frmPerformanceAlphaAlumniChapterName")		
	RST("AlphaAlumniChapterCity") = Request("frmPerformanceAlphaAlumniChapterCity")
	RST("AlphaAlumniChapterState") = Request("frmPerformanceAlphaAlumniChapterState")			
	
	' SDM Metrics Fields
	
	RST("YieldRate_Vol_Inquiries") = Request("frmPerformanceYieldRate_Vol_Inquiries")
	RST("YieldRate_Vol_Interviews") = Request("frmPerformanceYieldRate_Vol_Interviews")
	RST("YieldRate_Vol_Matched") = Request("frmPerformanceYieldRate_Vol_Matched")
	RST("YieldRate_Youth_Inquiries") = Request("frmPerformanceYieldRate_Youth_Inquiries")
	RST("YieldRate_Youth_Interviews") = Request("frmPerformanceYieldRate_Youth_Interviews")
	RST("YieldRate_Youth_Matched") = Request("frmPerformanceYieldRate_Youth_Matched")
	RST("ProcTim_Vol_InquiryToInterview_Number_Comm") = Request("frmPerformanceProcTim_Vol_InquiryToInterview_Number_Comm")
	RST("ProcTim_Vol_InquiryToInterview_AveDays_Comm") = Request("frmPerformanceProcTim_Vol_InquiryToInterview_AveDays_Comm")
	RST("ProcTim_Vol_InquiryToInterview_Number_School") = Request("frmPerformanceProcTim_Vol_InquiryToInterview_Number_School")
	RST("ProcTim_Vol_InquiryToInterview_AveDays_School") = Request("frmPerformanceProcTim_Vol_InquiryToInterview_AveDays_School")
	RST("ProcTim_Vol_InquiryToInterview_Number_Other") = Request("frmPerformanceProcTim_Vol_InquiryToInterview_Number_Other")
	RST("ProcTim_Vol_InquiryToInterview_AveDays_Other") = Request("frmPerformanceProcTim_Vol_InquiryToInterview_AveDays_Other")
	RST("ProcTim_Vol_InterviewToMatched_Number_Comm") = Request("frmPerformanceProcTim_Vol_InterviewToMatched_Number_Comm")
	RST("ProcTim_Vol_InterviewToMatched_AveDays_Comm") = Request("frmPerformanceProcTim_Vol_InterviewToMatched_AveDays_Comm")
	RST("ProcTim_Vol_InterviewToMatched_Number_School") = Request("frmPerformanceProcTim_Vol_InterviewToMatched_Number_School")
	RST("ProcTim_Vol_InterviewToMatched_AveDays_School") = Request("frmPerformanceProcTim_Vol_InterviewToMatched_AveDays_School")
	RST("ProcTim_Vol_InterviewToMatched_Number_Other") = Request("frmPerformanceProcTim_Vol_InterviewToMatched_Number_Other")
	RST("ProcTim_Vol_InterviewToMatched_AveDays_Other") = Request("frmPerformanceProcTim_Vol_InterviewToMatched_AveDays_Other")
	RST("ProcTim_Youth_InquiryToInterview_Number_Comm") = Request("frmPerformanceProcTim_Youth_InquiryToInterview_Number_Comm")
	RST("ProcTim_Youth_InquiryToInterview_AveDays_Comm") = Request("frmPerformanceProcTim_Youth_InquiryToInterview_AveDays_Comm")
	RST("ProcTim_Youth_InquiryToInterview_Number_School") = Request("frmPerformanceProcTim_Youth_InquiryToInterview_Number_School")
	RST("ProcTim_Youth_InquiryToInterview_AveDays_School") = Request("frmPerformanceProcTim_Youth_InquiryToInterview_AveDays_School")
	RST("ProcTim_Youth_InquiryToInterview_Number_Other") = Request("frmPerformanceProcTim_Youth_InquiryToInterview_Number_Other")
	RST("ProcTim_Youth_InquiryToInterview_AveDays_Other") = Request("frmPerformanceProcTim_Youth_InquiryToInterview_AveDays_Other")
	RST("ProcTim_Youth_InterviewToMatched_Number_Comm") = Request("frmPerformanceProcTim_Youth_InterviewToMatched_Number_Comm")
	RST("ProcTim_Youth_InterviewToMatched_AveDays_Comm") = Request("frmPerformanceProcTim_Youth_InterviewToMatched_AveDays_Comm")
	RST("ProcTim_Youth_InterviewToMatched_Number_School") = Request("frmPerformanceProcTim_Youth_InterviewToMatched_Number_School")
	RST("ProcTim_Youth_InterviewToMatched_AveDays_School") = Request("frmPerformanceProcTim_Youth_InterviewToMatched_AveDays_School")
	RST("ProcTim_Youth_InterviewToMatched_Number_Other") = Request("frmPerformanceProcTim_Youth_InterviewToMatched_Number_Other")
	RST("ProcTim_Youth_InterviewToMatched_AveDays_Other") = Request("frmPerformanceProcTim_Youth_InterviewToMatched_AveDays_Other")
	RST("Freq_Under3Months_Comm") = Request("frmPerformanceFreq_Under3Months_Comm")
	RST("Freq_Under3Months_School") = Request("frmPerformanceFreq_Under3Months_School")
	RST("Freq_Under3Months_Other") = Request("frmPerformanceFreq_Under3Months_Other")		
	RST("Freq_3To5Months_Comm") = Request("frmPerformanceFreq_3To5Months_Comm")
	RST("Freq_3To5Months_School") = Request("frmPerformanceFreq_3To5Months_School")
	RST("Freq_3To5Months_Other") = Request("frmPerformanceFreq_3To5Months_Other")		
	RST("Freq_6To8Months_Comm") = Request("frmPerformanceFreq_6To8Months_Comm")
	RST("Freq_6To8Months_School") = Request("frmPerformanceFreq_6To8Months_School")
	RST("Freq_6To8Months_Other") = Request("frmPerformanceFreq_6To8Months_Other")		
	RST("Freq_9To11Months_Comm") = Request("frmPerformanceFreq_9To11Months_Comm")
	RST("Freq_9To11Months_School") = Request("frmPerformanceFreq_9To11Months_School")
	RST("Freq_9To11Months_Other") = Request("frmPerformanceFreq_9To11Months_Other")		
	RST("Freq_12To23Months_Comm") = Request("frmPerformanceFreq_12To23Months_Comm")
	RST("Freq_12To23Months_School") = Request("frmPerformanceFreq_12To23Months_School")
	RST("Freq_12To23Months_Other") = Request("frmPerformanceFreq_12To23Months_Other")		
	RST("Freq_24OrMoreMonths_Comm") = Request("frmPerformanceFreq_24OrMoreMonths_Comm")
	RST("Freq_24OrMoreMonths_School") = Request("frmPerformanceFreq_24OrMoreMonths_School")
	RST("Freq_24OrMoreMonths_Other") = Request("frmPerformanceFreq_24OrMoreMonths_Other")
	RST("Volunteers_ReMatched") = Request("frmPerformanceVolunteers_ReMatched")
	RST("POE_Confidence_Number_Comm") = Request("frmPerformancePOE_Confidence_Number_Comm")
	RST("POE_Confidence_Ave_Comm") = Request("frmPerformancePOE_Confidence_Ave_Comm")
	RST("POE_Confidence_Number_School") = Request("frmPerformancePOE_Confidence_Number_School")
	RST("POE_Confidence_Ave_School") = Request("frmPerformancePOE_Confidence_Ave_School")
	RST("POE_Competence_Number_Comm") = Request("frmPerformancePOE_Competence_Number_Comm")
	RST("POE_Competence_Ave_Comm") = Request("frmPerformancePOE_Competence_Ave_Comm")
	RST("POE_Competence_Number_School") = Request("frmPerformancePOE_Competence_Number_School")
	RST("POE_Competence_Ave_School") = Request("frmPerformancePOE_Competence_Ave_School")
	RST("POE_Caring_Number_Comm") = Request("frmPerformancePOE_Caring_Number_Comm")
	RST("POE_Caring_Ave_Comm") = Request("frmPerformancePOE_Caring_Ave_Comm")
	RST("POE_Caring_Number_School") = Request("frmPerformancePOE_Caring_Number_School")
	RST("POE_Caring_Ave_School") = Request("frmPerformancePOE_Caring_Ave_School")
	RST("VolSat_PostEnrollment_Number") = Request("frmPerformanceVolSat_PostEnrollment_Number")
	RST("VolSat_PostEnrollment_Ave") = Request("frmPerformanceVolSat_PostEnrollment_Ave")
	RST("VolSat_SatQuest_Number") = Request("frmPerformanceVolSat_SatQuest_Number")
	RST("VolSat_SatQuest_Ave") = Request("frmPerformanceVolSat_SatQuest_Ave")
	
	RST("CBNumberClosedPrematurely") = Request("frmPerformanceCBNumberClosedPrematurely")
	RST("SBNumberClosedPrematurely") = Request("frmPerformanceSBNumberClosedPrematurely")						
	RST("CBChildParentStatusChange") = Request("frmPerformanceCBChildParentStatusChange")
	RST("CBVolunteerStatusChange") = Request("frmPerformanceCBVolunteerStatusChange")
	RST("CBChildParentDissatisfaction") = Request("frmPerformanceCBChildParentDissatisfaction")
	RST("CBVolunteerDissatisfaction") = Request("frmPerformanceCBVolunteerDissatisfaction")
	RST("SBChildParentStatusChange") = Request("frmPerformanceSBChildParentStatusChange")
	RST("SBVolunteerStatusChange") = Request("frmPerformanceSBVolunteerStatusChange")
	RST("SBChildParentDissatisfaction") = Request("frmPerformanceSBChildParentDissatisfaction")
	RST("SBVolunteerDissatisfaction") = Request("frmPerformanceSBVolunteerDissatisfaction")
	RST("CBTotalOpened6MonthsAgo") = Request("frmPerformanceCBTotalOpened6MonthsAgo")
	RST("CBNumberStillOpen") = Request("frmPerformanceCBNumberStillOpen")
	RST("EnrollmentSatAvgScore") = Request("frmPerformanceEnrollmentSatAvgScore")
	RST("EnrollmentSatCount") = Request("frmPerformanceEnrollmentSatCount")
	RST("MatchSatAvgScore") = Request("frmPerformanceMatchSatAvgScore")
	RST("MatchSatCount") = Request("frmPerformanceMatchSatCount")
	RST("CBPOEAggregateScore") = Request("frmPerformanceCBPOEAggregateScore")
	RST("CBPOECount") = Request("frmPerformanceCBPOECount")
	RST("SBPOEAggregateScore") = Request("frmPerformanceSBPOEAggregateScore")	
	RST("SBPOECount") = Request("frmPerformanceSBPOECount")			
	

		' Additional SDM fields
		
		RST("YieldRate_Vol_Inquiries_CB") = Request("frmPerformanceYieldRate_Vol_Inquiries_CB")
		RST("YieldRate_Vol_Inquiries_SB") = Request("frmPerformanceYieldRate_Vol_Inquiries_SB")
		RST("YieldRate_Vol_Inquiries_OSB") = Request("frmPerformanceYieldRate_Vol_Inquiries_OSB")
		RST("YieldRate_Vol_Interviews_CB") = Request("frmPerformanceYieldRate_Vol_Interviews_CB")
		RST("YieldRate_Vol_Interviews_SB") = Request("frmPerformanceYieldRate_Vol_Interviews_SB")		
		RST("YieldRate_Vol_Interviews_OSB") = Request("frmPerformanceYieldRate_Vol_Interviews_OSB")		
		RST("YieldRate_Vol_Matched_CB") = Request("frmPerformanceYieldRate_Vol_Matched_CB")
		RST("YieldRate_Vol_Matched_SB") = Request("frmPerformanceYieldRate_Vol_Matched_SB")		
		RST("YieldRate_Vol_Matched_OSB") = Request("frmPerformanceYieldRate_Vol_Matched_OSB")		
		RST("YieldRate_Youth_Inquiries_CB") = Request("frmPerformanceYieldRate_Youth_Inquiries_CB")
		RST("YieldRate_Youth_Inquiries_SB") = Request("frmPerformanceYieldRate_Youth_Inquiries_SB")
		RST("YieldRate_Youth_Inquiries_OSB") = Request("frmPerformanceYieldRate_Youth_Inquiries_OSB")
		RST("YieldRate_Youth_Interviews_CB") = Request("frmPerformanceYieldRate_Youth_Interviews_CB")
		RST("YieldRate_Youth_Interviews_SB") = Request("frmPerformanceYieldRate_Youth_Interviews_SB")		
		RST("YieldRate_Youth_Interviews_OSB") = Request("frmPerformanceYieldRate_Youth_Interviews_OSB")		
		RST("YieldRate_Youth_Matched_CB") = Request("frmPerformanceYieldRate_Youth_Matched_CB")
		RST("YieldRate_Youth_Matched_SB") = Request("frmPerformanceYieldRate_Youth_Matched_SB")		
		RST("YieldRate_Youth_Matched_OSB") = Request("frmPerformanceYieldRate_Youth_Matched_OSB")	
		RST("OSBNumberClosedPrematurely") = Request("frmPerformanceOSBNumberClosedPrematurely")
		RST("OSBChildParentStatusChange") = Request("frmPerformanceOSBChildParentStatusChange")
		RST("OSBVolunteerStatusChange") = Request("frmPerformanceOSBVolunteerStatusChange")
		RST("OSBChildParentDissatisfaction") = Request("frmPerformanceOSBChildParentDissatisfaction")
		RST("OSBVolunteerDissatisfaction") = Request("frmPerformanceOSBVolunteerDissatisfaction")
		RST("SBTotalOpened6MonthsAgo") = Request("frmPerformanceSBTotalOpened6MonthsAgo")
		RST("SBNumberStillOpen") = Request("frmPerformanceSBNumberStillOpen")
		RST("OSBTotalOpened6MonthsAgo") = Request("frmPerformanceOSBTotalOpened6MonthsAgo")
		RST("OSBNumberStillOpen") = Request("frmPerformanceOSBNumberStillOpen")																	
		RST("OSBPOEAggregateScore") = Request("frmPerformanceOSBPOEAggregateScore")
		RST("OSBPOECount") = Request("frmPerformanceOSBPOECount")		
		
	
	
	
	' RTBM Fields
		
	If Int(Request("month"))=12 then
		RST("RTBM_UnmatchedChildren") = Request("frmPerformanceRTBM_UnmatchedChildren")
		RST("RTBM_UnmatchedVolunteers") = Request("frmPerformanceRTBM_UnmatchedVolunteers")
	End If
	
	
	' SBM Fields
	
	If (Int(Request("month"))=6 or Int(Request("month"))=12) And SBMAgency = 1 Then
		RST("SBMVolunteersInEnrollmentProcess") = Request("frmPerformanceSBMVolunteersInEnrollmentProcess")
		RST("SBMAmountRaisedTowardsMatchPledge") = Request("frmPerformanceSBMAmountRaisedTowardsMatchPledge")
	End If
	
	If FBIAgency = 1 Then
	
		RST("CBIandFB") = Request("frmPerformanceCBIandFB")
		RST("CBInotFB") = Request("frmPerformanceCBInotFB")
		RST("CBFBnotI") = Request("frmPerformanceCBFBnotI")		
		RST("SBIandFB") = Request("frmPerformanceSBIandFB")
		RST("SBInotFB") = Request("frmPerformanceSBInotFB")
		RST("SBFBnotI") = Request("frmPerformanceSBFBnotI")				
		RST("OSBIandFB") = Request("frmPerformanceOSBIandFB")
		RST("OSBInotFB") = Request("frmPerformanceOSBInotFB")
		RST("OSBFBnotI") = Request("frmPerformanceOSBFBnotI")		
		RST("GMIandFB") = Request("frmPerformanceGMIandFB")
		RST("GMInotFB") = Request("frmPerformanceGMInotFB")
		RST("GMFBnotI") = Request("frmPerformanceGMFBnotI")				
		RST("SPMIandFB") = Request("frmPerformanceSPMIandFB")
		RST("SPMInotFB") = Request("frmPerformanceSPMInotFB")
		RST("SPMFBnotI") = Request("frmPerformanceSPMFBnotI")				
		RST("SPNMIandFB") = Request("frmPerformanceSPNMIandFB")
		RST("SPNMInotFB") = Request("frmPerformanceSPNMInotFB")
		RST("SPNMFBnotI") = Request("frmPerformanceSPNMFBnotI")				
		
	End If
		
	jMod = RST("PerformanceID") %>
	
	
	
	
	<%
	RST.Update
	RST.Close
	Set RST = Nothing
	form = "Performance"
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

		
	if(document.frmPerformance.frmPerformanceCBNumberClosedPrematurely.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceCBNumberClosedPrematurely.focus();}		
	else if(document.frmPerformance.frmPerformanceSBNumberClosedPrematurely.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceSBNumberClosedPrematurely.focus();}				
		
	else if(document.frmPerformance.frmPerformanceCBChildParentStatusChange.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceCBChildParentStatusChange.focus();}				
		
	else if(document.frmPerformance.frmPerformanceCBChildParentStatusChange.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceCBChildParentStatusChange.focus();}				
	else if(document.frmPerformance.frmPerformanceSBChildParentStatusChange.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceSBChildParentStatusChange.focus();}						
	else if(document.frmPerformance.frmPerformanceCBChildParentDissatisfaction.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceCBChildParentDissatisfaction.focus();}						
	else if(document.frmPerformance.frmPerformanceSBChildParentDissatisfaction.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceSBChildParentDissatisfaction.focus();}			
	else if(document.frmPerformance.frmPerformanceCBVolunteerDissatisfaction.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceCBVolunteerDissatisfaction.focus();}					
	else if(document.frmPerformance.frmPerformanceSBVolunteerDissatisfaction.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceSBVolunteerDissatisfaction.focus();}							
	else if(document.frmPerformance.frmPerformanceCBTotalOpened6MonthsAgo.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceCBTotalOpened6MonthsAgo.focus();}		
	else if(document.frmPerformance.frmPerformanceCBNumberStillOpen.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceCBNumberStillOpen.focus();}										
	else if(document.frmPerformance.frmPerformanceEnrollmentSatAvgScore.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceEnrollmentSatAvgScore.focus();}												
	else if(document.frmPerformance.frmPerformanceEnrollmentSatCount.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceEnrollmentSatCount.focus();}														
	else if(document.frmPerformance.frmPerformanceMatchSatAvgScore.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceMatchSatAvgScore.focus();}			
	else if(document.frmPerformance.frmPerformanceMatchSatCount.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceMatchSatCount.focus();}			
	else if(document.frmPerformance.frmPerformanceCBPOEAggregateScore.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceCBPOEAggregateScore.focus();}	
	else if(document.frmPerformance.frmPerformanceSBPOEAggregateScore.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceSBPOEAggregateScore.focus();}							
	else if(document.frmPerformance.frmPerformanceCBPOECount.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceCBPOECount.focus();}	
	else if(document.frmPerformance.frmPerformanceSBPOECount.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceSBPOECount.focus();}											
		
	else
		document.frmPerformance.submit();	
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


<form name="frmPerformance" action="SDMPerformance_edit.asp" method="post"> <!-- onsubmit="return submitFormValidate(this)"> -->
<!--#include file="../includes/form_stamp.asp"-->

<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y")) & " AND Month=" & Int(Request("m"))
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
			<td colspan="7" class="formHeader">PERFORMANCE - <%= MonthName(Request("m"), False) & " " & Request("y") %></td>
		</tr>
		
		<tr>
			<td colspan="7" class="formMain"><font color="#ff0000"><div align="center"><strong>Please Note: </strong>After entering your information, you <strong>must</strong> click on the "Save" button at the bottom of the form and wait for the "Thank You" screen or your changes will be lost.</div></font></td>
		</tr>




<% if (SDMPilot = 1 and ( y > 2002 or (y = 2002 and m >= 7))) or y > 2003 then %>

<!-- Prepopulate Core Business Fields -->

<input type="hidden"  value="<%= GetPerformance("OpenMatchesCommunityBased") %>" name="frmPerformanceOpenMatchesCommunityBased">
<input type="hidden"  value="<%= GetPerformance("OpenMatchesSchoolBased") %>" name="frmPerformanceOpenMatchesSchoolBased">
<input type="hidden"  value="<%= GetPerformance("OpenMatchesOtherSiteBased") %>" name="frmPerformanceOpenMatchesOtherSiteBased">
<input type="hidden"  value="<%= GetPerformance("OpenMatchesGroupMentoring") %>" name="frmPerformanceOpenMatchesGroupMentoring">
<input type="hidden"  value="<%= GetPerformance("OpenMatchesSpecialProgramsMentoring") %>" name="frmPerformanceOpenMatchesSpecialProgramsMentoring">
<input type="hidden"  value="<%= GetPerformance("OpenMatchesSpecialProgramsNonMentoring") %>" name="frmPerformanceOpenMatchesSpecialProgramsNonMentoring">

<input type="hidden"  value="<%= GetPerformance("NewMatchesCommunityBased") %>" name="frmPerformanceNewMatchesCommunityBased">
<input type="hidden"  value="<%= GetPerformance("NewMatchesSchoolBased") %>" name="frmPerformanceNewMatchesSchoolBased">
<input type="hidden"  value="<%= GetPerformance("NewMatchesSiteBasedNonSchool") %>" name="frmPerformanceNewMatchesSiteBasedNonSchool">
<input type="hidden"  value="<%= GetPerformance("NewMatchesGroupMentoring") %>" name="frmPerformanceNewMatchesGroupMentoring">
<input type="hidden"  value="<%= GetPerformance("NewMatchesSpecialProgramsMentoring") %>" name="frmPerformanceNewMatchesSpecialProgramsMentoring">
<input type="hidden"  value="<%= GetPerformance("NewMatchesSpecialProgramsNonMentoring") %>" name="frmPerformanceNewMatchesSpecialProgramsNonMentoring">

<input type="hidden"  value="<%= GetPerformance("ClosedMatchesCommunityBased") %>" name="frmPerformanceClosedMatchesCommunityBased">
<input type="hidden"  value="<%= GetPerformance("ClosedMatchesSchoolBased") %>" name="frmPerformanceClosedMatchesSchoolBased">
<input type="hidden"  value="<%= GetPerformance("ClosedMatchesOtherSiteBased") %>" name="frmPerformanceClosedMatchesOtherSiteBased">
<input type="hidden"  value="<%= GetPerformance("ClosedMatchesGroupMentoring") %>" name="frmPerformanceClosedMatchesGroupMentoring">
<input type="hidden"  value="<%= GetPerformance("ClosedMatchesSpecialProgramsMentoring") %>" name="frmPerformanceClosedMatchesSpecialProgramsMentoring">
<input type="hidden"  value="<%= GetPerformance("ClosedMatchesSpecialProgramsNonMentoring") %>" name="frmPerformanceClosedMatchesSpecialProgramsNonMentoring">
<input type="hidden"  value="<%= GetPerformance("AverageMatchLengthCB") %>" name="frmPerformanceAverageMatchLengthCB">
<input type="hidden"  value="<%= GetPerformance("AverageMatchLengthSB") %>" name="frmPerformanceAverageMatchLengthSB">
<input type="hidden"  value="<%= GetPerformance("AverageMatchLengthOSB") %>" name="frmPerformanceAverageMatchLengthOSB">
<input type="hidden"  value="<%= GetPerformance("Revenue") %>" name="frmPerformanceRevenue">
<input type="hidden"  value="<%= GetPerformance("RTBM_UnmatchedChildren") %>" name="frmPerformanceRTBM_UnmatchedChildren">
<input type="hidden"  value="<%= GetPerformance("RTBM_UnmatchedVolunteers") %>" name="frmPerformanceRTBM_UnmatchedVolunteers">


<!-- Prepopulate SBM Fields -->
<input type="hidden"  value="<%= GetPerformance("SBMVolunteersInEnrollmentProcess") %>" name="frmPerformanceSBMVolunteersInEnrollmentProcess">
<input type="hidden"  value="<%= GetPerformance("SBMAmountRaisedTowardsMatchPledge") %>" name="frmPerformanceSBMAmountRaisedTowardsMatchPledge">


<!-- Prepopulate Faith-Based / Children with Incarcerated Parents Fields -->
<input type="hidden"  value="<%= GetPerformance("CBIandFB") %>" name="frmPerformanceCBIandFB">				
<input type="hidden"  value="<%= GetPerformance("SBIandFB") %>" name="frmPerformanceSBIandFB">
<input type="hidden"  value="<%= GetPerformance("OSBIandFB")%>" name="frmPerformanceOSBIandFB">				
<input type="hidden"  value="<%= GetPerformance("CBInotFB") %>" name="frmPerformanceCBInotFB">				
<input type="hidden"  value="<%= GetPerformance("SBInotFB") %>" name="frmPerformanceSBInotFB">				
<input type="hidden"  value="<%= GetPerformance("OSBInotFB")%>" name="frmPerformanceOSBInotFB">				
<input type="hidden"  value="<%= GetPerformance("CBFBnotI") %>"  name="frmPerformanceCBFBnotI">				
<input type="hidden"  value="<%= GetPerformance("SBFBnotI") %>"  name="frmPerformanceSBFBnotI">				
<input type="hidden"  value="<%= GetPerformance("OSBFBnotI")%>" name="frmPerformanceOSBFBnotI">				
	
<!-- Prepopulate Partnership Fields -->
<input type="hidden" value="<%=GetPerformance("AlphaCommunityBased")%>" name="frmPerformanceAlphaCommunityBased">	
<input type="hidden" value="<%=GetPerformance("AlphaSchoolBased")%>" name="frmPerformanceAlphaSchoolBased">		
<input type="hidden" value="<%=GetPerformance("AlphaOtherSiteBased")%>" name="frmPerformanceAlphaOtherSiteBased">	
<input type="hidden" value="<%=GetPerformance("AlphaNotPartnering")%>" name="frmperformanceAlphaNotPartnering">	
<input type="hidden" value="<%=GetPerformance("AlphaInterest")%>" name="frmperformanceAlphaInterest">	

<input type="hidden" value="<%=GetPerformance("LionsCommunityBased")%>" name="frmPerformanceLionsCommunityBased">	
<input type="hidden" value="<%=GetPerformance("LionsSchoolBased")%>" name="frmPerformanceLionsSchoolBased">		
<input type="hidden" value="<%=GetPerformance("LionsOtherSiteBased")%>" name="frmPerformanceLionsOtherSiteBased">		
<input type="hidden" value="<%=GetPerformance("LionsNotPartnering")%>" name="frmperformanceLionsNotPartnering">	
<input type="hidden" value="<%=GetPerformance("LionsInterest")%>" name="frmperformanceLionsInterest">	

<input type="hidden" value="<%=GetPerformance("RotaryCommunityBased")%>" name="frmPerformanceRotaryCommunityBased">	
<input type="hidden" value="<%=GetPerformance("RotarySchoolBased")%>" name="frmPerformanceRotarySchoolBased">		
<input type="hidden" value="<%=GetPerformance("RotaryOtherSiteBased")%>" name="frmPerformanceRotaryOtherSiteBased">		
<input type="hidden" value="<%=GetPerformance("RotaryNotPartnering")%>" name="frmperformanceRotaryNotPartnering">	
<input type="hidden" value="<%=GetPerformance("RotaryInterest")%>" name="frmperformanceRotaryInterest">	

<input type="hidden" value="<%=GetPerformance("KiwanisCommunityBased")%>" name="frmPerformanceKiwanisCommunityBased">	
<input type="hidden" value="<%=GetPerformance("KiwanisSchoolBased")%>" name="frmPerformanceKiwanisSchoolBased">		
<input type="hidden" value="<%=GetPerformance("KiwanisOtherSiteBased")%>" name="frmPerformanceKiwanisOtherSiteBased">		

<input type="hidden" value="<%=GetPerformance("KiwanisNotPartnering")%>" name="frmperformanceKiwanisNotPartnering">	
<input type="hidden" value="<%=GetPerformance("KiwanisInterest")%>" name="frmperformanceKiwanisInterest">		

<input type="hidden" value="<%=GetPerformance("OptimistCommunityBased")%>" name="frmPerformanceOptimistCommunityBased">	
<input type="hidden" value="<%=GetPerformance("OptimistSchoolBased")%>" name="frmPerformanceOptimistSchoolBased">		
<input type="hidden" value="<%=GetPerformance("OptimistOtherSiteBased")%>" name="frmPerformanceOptimistOtherSiteBased">		
<input type="hidden" value="<%=GetPerformance("OptimistNotPartnering")%>" name="frmperformanceOptimistNotPartnering">	
<input type="hidden" value="<%=GetPerformance("OptimistInterest")%>" name="frmperformanceOptimistInterest">		

<input type="hidden" value="<%=GetPerformance("AARPCommunityBased")%>" name="frmPerformanceAARPCommunityBased">	
<input type="hidden" value="<%=GetPerformance("AARPSchoolBased")%>" name="frmPerformanceAARPSchoolBased">		
<input type="hidden" value="<%=GetPerformance("AARPOtherSiteBased")%>" name="frmPerformanceAARPOtherSiteBased">		
<input type="hidden" value="<%=GetPerformance("AARPNotPartnering")%>" name="frmperformanceAARPNotPartnering">	
<input type="hidden" value="<%=GetPerformance("AARPInterest")%>" name="frmperformanceAARPInterest">		

<input type="hidden" value="<%=GetPerformance("AlphaRating")%>" name="frmperformanceAlphaRating">
<input type="hidden" value="<%=GetPerformance("LionsRating")%>" name="frmperformanceLionsRating">
<input type="hidden" value="<%=GetPerformance("RotaryRating")%>" name="frmperformanceRotaryRating">			
<input type="hidden" value="<%=GetPerformance("KiwanisRating")%>" name="frmperformanceKiwanisRating">
<input type="hidden" value="<%=GetPerformance("OptimistRating")%>" name="frmperformanceOptimistRating">				
<input type="hidden" value="<%=GetPerformance("AARPRating")%>" name="frmperformanceAARPRating">		

<input type="hidden" value="<%=GetPerformance("AlphaFunding")%>" name="frmperformanceAlphaFunding">		
<input type="hidden" value="<%=GetPerformance("AlphaProgramInitiative")%>" name="frmperformanceAlphaProgramInitiative">			
<input type="hidden" value="<%=GetPerformance("AlphaLeadershipInvolvement")%>" name="frmperformanceAlphaLeadershipInvolvement">				

<input type="hidden" value="<%=GetPerformance("AlphaUndergradChapterName")%> " name="frmperformanceAlphaUndergradChapterName">		
<input type="hidden" value="<%=GetPerformance("AlphaUndergradChapterCity")%> " name="frmperformanceAlphaUndergradChapterCity">			
<input type="hidden" value="<%=GetPerformance("AlphaUndergradChapterState")%> " name="frmperformanceAlphaUndergradChapterState">			

<input type="hidden" value="<%=GetPerformance("AlphaAlumniChapterName")%>" name="frmperformanceAlphaAlumniChapterName">		
<input type="hidden" value="<%=GetPerformance("AlphaAlumniChapterCity")%>"  name="frmperformanceAlphaAlumniChapterCity">			
<input type="hidden" value="<%=GetPerformance("AlphaAlumniChapterState")%>" name="frmperformanceAlphaAlumniChapterState">


<!-- SDM Metrics -->
			<tr>
				<td colspan="7">&nbsp;</td>
			</tr>
			<tr>				
				<td colspan="7" class="formHeader">SDM METRICS</td>
			</tr>
			
			<TR>
				<TD colspan="7" class="formHeaderMedium">YIELD RATE DATA</TD>
			</TR>
			
			<tr>
				<td>&nbsp;</td>
				<td class="formMain" colspan="6" align="center"><strong>Volunteer</strong></td>			
			</tr>
			
			<tr>
				<td>&nbsp;</td>
				<td class="formMain" colspan="2" align="left">Community-Based</td>
				<td class="formMain" colspan="2" align="left">School-Based</td>				
				<td class="formMain" colspan="2" align="left">Non-School Site-Based</td>				
			</tr>
			
			<tr>
				<td class="formMain" align="left">Number of Inquiries</td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Inquiries_CB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Inquiries_CB" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm1" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>				
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Inquiries_SB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Inquiries_SB" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm1" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>								
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Inquiries_OSB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Inquiries_OSB" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm1" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>												
			</tr>
			
			<tr>
				<td class="formMain" align="left">Number of In-Person Interviews&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm3" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Interviews_CB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Interviews_CB" onchange="checkForIntegerCommas(this.value);"></td>								
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Interviews_SB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Interviews_SB" onchange="checkForIntegerCommas(this.value);"></td>												
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Interviews_OSB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Interviews_OSB" onchange="checkForIntegerCommas(this.value);"></td>												
			</tr>
			
			<tr>
				<td class="formMain" align="left">Number Matched</td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Matched_CB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Matched_CB" onchange="checkForIntegerCommas(this.value);"></td>												
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Matched_SB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Matched_SB" onchange="checkForIntegerCommas(this.value);"></td>				
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Matched_OSB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Matched_OSB" onchange="checkForIntegerCommas(this.value);"></td>								
			</tr>			
			
			
			<tr>
				<td>&nbsp;</td>
				<td class="formMain" colspan="6" align="center"><strong>Youth</strong></td>			
			</tr>
			
			<tr>
				<td>&nbsp;</td>
				<td class="formMain" colspan="2" align="left">Community-Based</td>
				<td class="formMain" colspan="2" align="left">School-Based</td>				
				<td class="formMain" colspan="2" align="left">Non-School Site-Based</td>				
			</tr>
			
			<tr>
				<td class="formMain" align="left">Number of Inquiries</td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Inquiries_CB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Inquiries_CB" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm2" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>				
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Inquiries_SB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Inquiries_SB" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm2" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>								
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Inquiries_OSB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Inquiries_OSB" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm2" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>												
			</tr>
			
			<tr>
				<td class="formMain" align="left">Number of In-Person Interviews&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm3" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Interviews_CB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Interviews_CB" onchange="checkForIntegerCommas(this.value);"></td>								
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Interviews_SB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Interviews_SB" onchange="checkForIntegerCommas(this.value);"></td>												
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Interviews_OSB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Interviews_OSB" onchange="checkForIntegerCommas(this.value);"></td>												
			</tr>
			
			<tr>
				<td class="formMain" align="left">Number Matched</td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Matched_CB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Matched_CB" onchange="checkForIntegerCommas(this.value);"></td>												
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Matched_SB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Matched_SB" onchange="checkForIntegerCommas(this.value);"></td>				
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Matched_OSB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Matched_OSB" onchange="checkForIntegerCommas(this.value);"></td>								
			</tr>				
			
			
			
			<!-- Old SDM Yield Rate Fields - Zero them out -->		
			<input type="hidden" name="frmPerformanceYieldRate_Vol_Inquiries" value="0">
			<input type="hidden" name="frmPerformanceYieldRate_Youth_Inquiries" value="0">
			<input type="hidden" name="frmPerformanceYieldRate_Vol_Interviews" value="0">
			<input type="hidden" name="frmPerformanceYieldRate_Youth_Interviews" value="0">
			<input type="hidden" name="frmPerformanceYieldRate_Vol_Matched" value="0">
			<input type="hidden" name="frmPerformanceYieldRate_Youth_Matched" value="0">

			
			<TR>
				<TD colspan="7" class="formHeaderMedium">PROCESSING TIME</TD>
			</TR>		
			
			<tr>
				<td>&nbsp;</td>
				<td colspan="6" class="formMain" align="center"><strong>Volunteer</strong></td>
				<!-- <td colspan="4" class="formMain" align="center"><strong>Parent / Youth</strong></td>-->
			</tr>		
			
			<tr>
				<td>&nbsp;</td>
				<td colspan="2" class="formMain">Community-Based</td>
				<td colspan="2" class="formMain">School-Based</td>
				<td colspan="2" class="formMain">Other Site-Based</td>

		
			</tr>
			
			<tr>
				<td>&nbsp;</td>
				<td class="formMain">Number of Individuals</td>
				<td class="formMain">Average Days</td>
				<td class="formMain">Number of Individuals</td>
				<td class="formMain">Average Days</td>	
				<td class="formMain">Number of Individuals</td>
				<td class="formMain">Average Days</td>
			</tr>
			
			<tr>
				<td class="formMain">Inquiry to Interview&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm5" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_Number_Comm") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InquiryToInterview_Number_Comm" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_AveDays_Comm") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InquiryToInterview_AveDays_Comm" onchange="checkForIntegerCommas(this.value);"></td>

				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_Number_School") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InquiryToInterview_Number_School" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_AveDays_School") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InquiryToInterview_AveDays_School" onchange="checkForIntegerCommas(this.value);"></td>

				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_Number_Other") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InquiryToInterview_Number_Other" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_AveDays_Other") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InquiryToInterview_AveDays_Other" onchange="checkForIntegerCommas(this.value);"></td>				
				
			</tr>

			<tr>
				<td class="formMain">Interview to Matched&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm6" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InterviewToMatched_Number_Comm") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InterviewToMatched_Number_Comm" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InterviewToMatched_AveDays_Comm") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InterviewToMatched_AveDays_Comm" onchange="checkForIntegerCommas(this.value);"></td>

				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InterviewToMatched_Number_School") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InterviewToMatched_Number_School" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InterviewToMatched_AveDays_School") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InterviewToMatched_AveDays_School" onchange="checkForIntegerCommas(this.value);"></td>

				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InterviewToMatched_Number_Other") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InterviewToMatched_Number_Other" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InterviewToMatched_AveDays_Other") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InterviewToMatched_AveDays_Other" onchange="checkForIntegerCommas(this.value);"></td>

			</tr>
			
			
			<tr>
				<td>&nbsp;</td>
				<td colspan="6" class="formMain" align="center"><strong>Parent / Youth</strong></td>
			</tr>		
			
			<tr>
				<td>&nbsp;</td>
				<td colspan="2" class="formMain">Community-Based</td>
				<td colspan="2" class="formMain">School-Based</td>
				<td colspan="2" class="formMain">Other Site-Based</td>
			</tr>
			
			<tr>
				<td>&nbsp;</td>
				<td class="formMain">Number of Individuals</td>
				<td class="formMain">Average Days</td>
				<td class="formMain">Number of Individuals</td>
				<td class="formMain">Average Days</td>	
				<td class="formMain">Number of Individuals</td>
				<td class="formMain">Average Days</td>
			</tr>
			
			<tr>
				<td class="formMain">Inquiry to Interview&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm5" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_Number_Comm") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InquiryToInterview_Number_Comm" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_AveDays_Comm") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InquiryToInterview_AveDays_Comm" onchange="checkForIntegerCommas(this.value);"></td>

				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_Number_School") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InquiryToInterview_Number_School" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_AveDays_School") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InquiryToInterview_AveDays_School" onchange="checkForIntegerCommas(this.value);"></td>
				
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_Number_Other") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InquiryToInterview_Number_Other" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_AveDays_Other") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InquiryToInterview_AveDays_Other" onchange="checkForIntegerCommas(this.value);"></td>
				
				
			</tr>

			<tr>
				<td class="formMain">Interview to Matched&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm6" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>

				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InterviewToMatched_Number_Comm") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InterviewToMatched_Number_Comm" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InterviewToMatched_AveDays_Comm") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InterviewToMatched_AveDays_Comm" onchange="checkForIntegerCommas(this.value);"></td>

				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InterviewToMatched_Number_School") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InterviewToMatched_Number_School" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InterviewToMatched_AveDays_School") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InterviewToMatched_AveDays_School" onchange="checkForIntegerCommas(this.value);"></td>
				
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InterviewToMatched_Number_Other") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InterviewToMatched_Number_Other" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InterviewToMatched_AveDays_Other") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InterviewToMatched_AveDays_Other" onchange="checkForIntegerCommas(this.value);"></td>				
				
			</tr>
			

			
			<tr>
				<TD colspan="7" class="formHeaderMedium">NUMBER OF MATCH CLOSURES&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm7" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmark_purplesmall.gif" alt="" width="15" height="16" border="0"></a></TD>	
			</tr>
			
			<tr>		
				<td>&nbsp;</td>
				<td class="formMain" colspan="2" align="center"><strong>Community-Based</strong></td>
				<td class="formMain" colspan="2" align="center"><strong>School-Based</strong></td>			
				<td class="formMain" colspan="2" align="center"><strong>Non-School Site-Based</strong></td>		
			</tr>
			
			<tr>
				<td class="formMain">Less Than 3 Months</td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_Under3Months_Comm") %><% Else %>0<% End If %>" name="frmPerformanceFreq_Under3Months_Comm" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_Under3Months_School") %><% Else %>0<% End If %>" name="frmPerformanceFreq_Under3Months_School" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_Under3Months_Other") %><% Else %>0<% End If %>" name="frmPerformanceFreq_Under3Months_Other" onchange="checkForIntegerCommas(this.value);"></td>				
			</tr>
			
			<tr>
				<td class="formMain">3-5 Months</td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_3To5Months_Comm") %><% Else %>0<% End If %>" name="frmPerformanceFreq_3To5Months_Comm" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_3To5Months_School") %><% Else %>0<% End If %>" name="frmPerformanceFreq_3To5Months_School" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_3To5Months_Other") %><% Else %>0<% End If %>" name="frmPerformanceFreq_3To5Months_Other" onchange="checkForIntegerCommas(this.value);"></td>				
			</tr>		
			
			<tr>
				<td class="formMain">6-8 Months</td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_6To8Months_Comm") %><% Else %>0<% End If %>" name="frmPerformanceFreq_6To8Months_Comm" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_6To8Months_School") %><% Else %>0<% End If %>" name="frmPerformanceFreq_6To8Months_School" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_6To8Months_Other") %><% Else %>0<% End If %>" name="frmPerformanceFreq_6To8Months_Other" onchange="checkForIntegerCommas(this.value);"></td>				
			</tr>		
			
			<tr>
				<td class="formMain">9-11 Months</td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_9To11Months_Comm") %><% Else %>0<% End If %>" name="frmPerformanceFreq_9To11Months_Comm" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_9To11Months_School") %><% Else %>0<% End If %>" name="frmPerformanceFreq_9To11Months_School" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_9To11Months_Other") %><% Else %>0<% End If %>" name="frmPerformanceFreq_9To11Months_Other" onchange="checkForIntegerCommas(this.value);"></td>				
			</tr>
			
			<tr>
				<td class="formMain">12-23 Months</td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_12To23Months_Comm") %><% Else %>0<% End If %>" name="frmPerformanceFreq_12To23Months_Comm" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_12To23Months_School") %><% Else %>0<% End If %>" name="frmPerformanceFreq_12To23Months_School" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_12To23Months_Other") %><% Else %>0<% End If %>" name="frmPerformanceFreq_12To23Months_Other" onchange="checkForIntegerCommas(this.value);"></td>				
			</tr>
			
			<tr>
				<td class="formMain">24 or More Months</td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_24OrMoreMonths_Comm") %><% Else %>0<% End If %>" name="frmPerformanceFreq_24OrMoreMonths_Comm" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_24OrMoreMonths_School") %><% Else %>0<% End If %>" name="frmPerformanceFreq_24OrMoreMonths_School" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Freq_24OrMoreMonths_Other") %><% Else %>0<% End If %>" name="frmPerformanceFreq_24OrMoreMonths_Other" onchange="checkForIntegerCommas(this.value);"></td>				
			</tr>
			
			<tr>
				<TD colspan="7" class="formHeaderMedium">VOLUNTEERS RE-MATCHED</TD>	
			</tr>														
			
			<tr>
				<td class="formMain">Volunteers Re-Matched&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm8" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<td class="formMain" colspan="7" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("Volunteers_ReMatched") %><% Else %>0<% End If %>" name="frmPerformanceVolunteers_ReMatched" onchange="checkForIntegerCommas(this.value);"></td>
			</tr>
			
<!-- Additional SDM Fields for October 2003 and beyond -->

<% if (y=2003 and m >= 10) or (y > 2003) then %>

			<tr>
				<TD colspan="7" class="formHeaderMedium">PREMATURE CLOSURE</TD>
			</tr>
			
			<TR>
				<TD colspan="1">&nbsp;</TD>
				<TD colspan="2" class="formMain" align="center"><b>Community-Based</b></TD>
				<TD colspan="2" class="formMain" align="center"><b>School-Based</b></TD>
				<TD colspan="2" class="formMain" align="center"><b>Non-School Site-Based</b></TD>				
	
			</TR>
			
			<tr>
				<td colspan="1" class="formMain">Number of Matches that Closed Prematurely</td>			
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBNumberClosedPrematurely") %><% Else %>0<% End If %>" name="frmPerformanceCBNumberClosedPrematurely" onchange="checkForIntegerCommas(this.value);"></td>				
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBNumberClosedPrematurely") %><% Else %>0<% End If %>" name="frmPerformanceSBNumberClosedPrematurely" onchange="checkForIntegerCommas(this.value);"></td>								
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBNumberClosedPrematurely") %><% Else %>0<% End If %>" name="frmPerformanceOSBNumberClosedPrematurely" onchange="checkForIntegerCommas(this.value);"></td>												
			</tr>
			
			<tr>
				<TD colspan="7" class="formHeaderMedium">CLOSE CODES</TD>	
			</tr>
			
			<TR>
				<TD colspan="1">&nbsp;</TD>
				<TD colspan="2" class="formMain" align="center"><b>Community-Based</b></TD>
				<TD colspan="2" class="formMain" align="center"><b>School-Based</b></TD>
				<TD colspan="2" class="formMain" align="center"><b>Non-School Site-Based</b></TD>				
			</TR>
				
			<tr>
				<td colspan="1" class="formMain">Child/Parent Status Change&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm11" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBChildParentStatusChange") %><% Else %>0<% End If %>" name="frmPerformanceCBChildParentStatusChange" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBChildParentStatusChange") %><% Else %>0<% End If %>" name="frmPerformanceSBChildParentStatusChange" onchange="checkForIntegerCommas(this.value);"></td>				
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBChildParentStatusChange") %><% Else %>0<% End If %>" name="frmPerformanceOSBChildParentStatusChange" onchange="checkForIntegerCommas(this.value);"></td>								

			</tr>
			
			<tr>
				<td colspan="1" class="formMain">Volunteer Status Change&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm12" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>	
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBVolunteerStatusChange") %><% Else %>0<% End If %>" name="frmPerformanceCBVolunteerStatusChange" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBVolunteerStatusChange") %><% Else %>0<% End If %>" name="frmPerformanceSBVolunteerStatusChange" onchange="checkForIntegerCommas(this.value);"></td>				
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBVolunteerStatusChange") %><% Else %>0<% End If %>" name="frmPerformanceOSBVolunteerStatusChange" onchange="checkForIntegerCommas(this.value);"></td>								
			</tr>
			
			<tr>
				<td colspan="1" class="formMain">Child/Parent Dissatisfaction&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm13" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>	
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBChildParentDissatisfaction") %><% Else %>0<% End If %>" name="frmPerformanceCBChildParentDissatisfaction" onchange="checkForIntegerCommas(this.value);"></td>				
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBChildParentDissatisfaction") %><% Else %>0<% End If %>" name="frmPerformanceSBChildParentDissatisfaction" onchange="checkForIntegerCommas(this.value);"></td>								
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBChildParentDissatisfaction") %><% Else %>0<% End If %>" name="frmPerformanceOSBChildParentDissatisfaction" onchange="checkForIntegerCommas(this.value);"></td>												
			</tr>
			
			<tr>

				<td colspan="1" class="formMain">Volunteer Dissatisfaction&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm14" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBVolunteerDissatisfaction") %><% Else %>0<% End If %>" name="frmPerformanceCBVolunteerDissatisfaction" onchange="checkForIntegerCommas(this.value);"></td>								
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBVolunteerDissatisfaction") %><% Else %>0<% End If %>" name="frmPerformanceSBVolunteerDissatisfaction" onchange="checkForIntegerCommas(this.value);"></td>												
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBVolunteerDissatisfaction") %><% Else %>0<% End If %>" name="frmPerformanceOSBVolunteerDissatisfaction" onchange="checkForIntegerCommas(this.value);"></td>																
			</tr>
			
			<tr>
				<TD colspan="7" class="formHeaderMedium">6-Month Retention</TD>	
			</tr>
			
			<TR>
				<TD colspan="1">&nbsp;</TD>
				<TD colspan="2" class="formMain" align="center"><b>Community-Based</b></TD>
				<TD colspan="2" class="formMain" align="center"><b>School-Based</b></TD>
				<TD colspan="2" class="formMain" align="center"><b>Non-School Site-Based</b></TD>				
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
							end if
						end if
					end if 
				end if
			end if %>
			
			
			
			<tr>
				<td colspan="1" class="formMain">New Matches Made 6 Months Ago&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm15&SixMonthsAgo=<%=SixMonthsAgo%>" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>				
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBTotalOpened6MonthsAgo") %><% Else %>0<% End If %>" name="frmPerformanceCBTotalOpened6MonthsAgo" onchange="checkForIntegerCommas(this.value);"></td>								
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBTotalOpened6MonthsAgo") %><% Else %>0<% End If %>" name="frmPerformanceSBTotalOpened6MonthsAgo" onchange="checkForIntegerCommas(this.value);"></td>												
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBTotalOpened6MonthsAgo") %><% Else %>0<% End If %>" name="frmPerformanceOSBTotalOpened6MonthsAgo" onchange="checkForIntegerCommas(this.value);"></td>												
			</tr>	
			
			<tr>
				<td colspan="1" class="formMain">Number Still Open Now&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm16&SixMonthsAgo=<%=SixMonthsAgo%>&Now=<%=m%>" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBNumberStillOpen") %><% Else %>0<% End If %>" name="frmPerformanceCBNumberStillOpen" onchange="checkForIntegerCommas(this.value);"></td>												
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBNumberStillOpen") %><% Else %>0<% End If %>" name="frmPerformanceSBNumberStillOpen" onchange="checkForIntegerCommas(this.value);"></td>																
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBNumberStillOpen") %><% Else %>0<% End If %>" name="frmPerformanceOSBNumberStillOpen" onchange="checkForIntegerCommas(this.value);"></td>														
			</tr>
			
			<tr>
				<TD colspan="7" class="formHeaderMedium">Customer Satisfaction</TD>	
			</tr>	
			
			<tr>
				<td colspan="1" class="formMain">Enrollment Satisfaction Average Score</td>
				<td class="formMain" colspan="6" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("EnrollmentSatAvgScore") %><% Else %>0<% End If %>" name="frmPerformanceEnrollmentSatAvgScore" onchange="checkForIntegerCommas(this.value);"></td>				
			</tr>				
			
			<tr>
				<td colspan="1" class="formMain">Enrollment Satisfaction Count</td>
				<td class="formMain" colspan="6" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("EnrollmentSatCount") %><% Else %>0<% End If %>" name="frmPerformanceEnrollmentSatCount" onchange="checkForIntegerCommas(this.value);"></td>								
			</tr>

			<tr>
				<td colspan="1" class="formMain">Match Satisfaction Average Score</td>
				<td class="formMain" colspan="6" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("MatchSatAvgScore") %><% Else %>0<% End If %>" name="frmPerformanceMatchSatAvgScore" onchange="checkForIntegerCommas(this.value);"></td>								
			</tr>			
			
			<tr>
				<td colspan="1" class="formMain">Match Satisfaction Count</td>
				<td class="formMain" colspan="6" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("MatchSatCount") %><% Else %>0<% End If %>" name="frmPerformanceMatchSatCount" onchange="checkForIntegerCommas(this.value);"></td>								
			</tr>	
			
			<tr>
				<TD colspan="7" class="formHeaderMedium">POE</TD>	
			</tr>								
			
			<tr>
				<td colspan="1">&nbsp;</td>
				<td colspan="2" class="formMain" align="center"><b>Community-Based</b></td>
				<td colspan="2" class="formMain" align="center"><b>School-Based</b></td>
				<td colspan="2" class="formMain" align="center"><b>Non-School Site-Based</b></td>				
			</tr>
			
			<tr>
				<td colspan="1" class="formMain">POE Aggregate Score</td>
				<td colspan="2" class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBPOEAggregateScore") %><% Else %>0<% End If %>" name="frmPerformanceCBPOEAggregateScore" onchange="checkForIntegerCommas(this.value);"></td>
				<td colspan="2" class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBPOEAggregateScore") %><% Else %>0<% End If %>" name="frmPerformanceSBPOEAggregateScore" onchange="checkForIntegerCommas(this.value);"></td>				
				<td colspan="2" class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBPOEAggregateScore") %><% Else %>0<% End If %>" name="frmPerformanceOSBPOEAggregateScore" onchange="checkForIntegerCommas(this.value);"></td>								
			</tr>
			
			<tr>
				<td colspan="1" class="formMain">POE Count</td>	
				<td colspan="2" class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBPOECount") %><% Else %>0<% End If %>" name="frmPerformanceCBPOECount" onchange="checkForIntegerCommas(this.value);"></td>
				<td colspan="2" class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBPOECount") %><% Else %>0<% End If %>" name="frmPerformanceSBPOECount" onchange="checkForIntegerCommas(this.value);"></td>								
				<td colspan="2" class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBPOECount") %><% Else %>0<% End If %>" name="frmPerformanceSBPOECount" onchange="checkForIntegerCommas(this.value);"></td>												
			</tr>
			
<% else %>

			<!-- Put zero values to prevent nulls and not trigger validation errors -->
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBNumberClosedPrematurely" onchange="checkForIntegerCommas(this.value);">		
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceSBNumberClosedPrematurely" onchange="checkForIntegerCommas(this.value);">					
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBChildParentStatusChange" onchange="checkForIntegerCommas(this.value);">								
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceSBChildParentStatusChange" onchange="checkForIntegerCommas(this.value);">											
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBVolunteerStatusChange" onchange="checkForIntegerCommas(this.value);">														
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceSBVolunteerStatusChange" onchange="checkForIntegerCommas(this.value);">																	
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBChildParentDissatisfaction" onchange="checkForIntegerCommas(this.value);">					
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceSBChildParentDissatisfaction" onchange="checkForIntegerCommas(this.value);">								
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceSBVolunteerDissatisfaction" onchange="checkForIntegerCommas(this.value);">					
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBVolunteerDissatisfaction" onchange="checkForIntegerCommas(this.value);">								
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBTotalOpened6MonthsAgo" onchange="checkForIntegerCommas(this.value);">											
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBNumberStillOpen" onchange="checkForIntegerCommas(this.value);">														
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceEnrollmentSatAvgScore" onchange="checkForIntegerCommas(this.value);">					
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceEnrollmentSatCount" onchange="checkForIntegerCommas(this.value);">	
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceMatchSatAvgScore" onchange="checkForIntegerCommas(this.value);">				
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceMatchSatCount" onchange="checkForIntegerCommas(this.value);">				
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBPOEAggregateScore" onchange="checkForIntegerCommas(this.value);">				
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBPOECount" onchange="checkForIntegerCommas(this.value);">					
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceSBPOEAggregateScore" onchange="checkForIntegerCommas(this.value);">								
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceSBPOECount" onchange="checkForIntegerCommas(this.value);">												
			
			
			
			

<% end if %>			

		

<% else %>		

			<!-- Put zero values to prevent nulls and not trigger validation errors -->
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBNumberClosedPrematurely" onchange="checkForIntegerCommas(this.value);">		
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceSBNumberClosedPrematurely" onchange="checkForIntegerCommas(this.value);">					
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBChildParentStatusChange" onchange="checkForIntegerCommas(this.value);">								
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceSBChildParentStatusChange" onchange="checkForIntegerCommas(this.value);">											
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBVolunteerStatusChange" onchange="checkForIntegerCommas(this.value);">														
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceSBVolunteerStatusChange" onchange="checkForIntegerCommas(this.value);">																	
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBChildParentDissatisfaction" onchange="checkForIntegerCommas(this.value);">					
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceSBChildParentDissatisfaction" onchange="checkForIntegerCommas(this.value);">								
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceSBVolunteerDissatisfaction" onchange="checkForIntegerCommas(this.value);">					
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBVolunteerDissatisfaction" onchange="checkForIntegerCommas(this.value);">								
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBTotalOpened6MonthsAgo" onchange="checkForIntegerCommas(this.value);">											
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBNumberStillOpen" onchange="checkForIntegerCommas(this.value);">														
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceEnrollmentSatAvgScore" onchange="checkForIntegerCommas(this.value);">					
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceEnrollmentSatCount" onchange="checkForIntegerCommas(this.value);">	
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceMatchSatAvgScore" onchange="checkForIntegerCommas(this.value);">				
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceMatchSatCount" onchange="checkForIntegerCommas(this.value);">				
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBPOEAggregateScore" onchange="checkForIntegerCommas(this.value);">				
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBPOECount" onchange="checkForIntegerCommas(this.value);">					
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceSBPOEAggregateScore" onchange="checkForIntegerCommas(this.value);">								
			<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceSBPOECount" onchange="checkForIntegerCommas(this.value);">											


<% end if %>
			


<!-- ADD PREVIOUS MONTH'S MATCH FIELDS TO FORM FOR COMPARISON -->
<input type="hidden" class="formMain" size="5" maxlength="10" value="<%=Request("PrevOpenComm")%>" name="frmPerformancePrevOpenComm" onchange="checkForIntegerCommas(this.value);">	
<input type="hidden" class="formMain" size="5" maxlength="10" value="<%=Request("PrevOpenSchool")%>" name="frmPerformancePrevOpenSchool" onchange="checkForIntegerCommas(this.value);">	
<input type="hidden" class="formMain" size="5" maxlength="10" value="<%=Request("PrevOpenOther")%>" name="frmPerformancePrevOpenOther" onchange="checkForIntegerCommas(this.value);">	
<input type="hidden" class="formMain" size="5" maxlength="10" value="<%=Request("PrevOpenGroup")%>" name="frmPerformancePrevOpenGroup" onchange="checkForIntegerCommas(this.value);">	
<input type="hidden" class="formMain" size="5" maxlength="10" value="<%=Request("PrevOpenSpecMent")%>" name="frmPerformancePrevOpenSpecMent" onchange="checkForIntegerCommas(this.value);">	
<input type="hidden" class="formMain" size="5" maxlength="10" value="<%=Request("PrevOpenSpecNonMent")%>" name="frmPerformancePrevOpenSpecNonMent" onchange="checkForIntegerCommas(this.value);">	



		<tr>
				<td colspan="7" class="formHeader">


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

