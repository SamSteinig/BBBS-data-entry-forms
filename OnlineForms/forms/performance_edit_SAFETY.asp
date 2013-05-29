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
		
		RST("TransferCommunityBased") = Request("frmPerformanceTransferCommunityBased")
		RST("TransferSchoolBased") = Request("frmPerformanceTransferSchoolBased")
		RST("TransferOtherSiteBased") = Request("frmPerformanceTransferOtherSiteBased")
		RST("TransferGroupMentoring") = Request("frmPerformanceTransferGroupMentoring")
		RST("TransferSpecialProgramsMentoring") = Request("frmPerformanceTransferSpecialProgramsMentoring")
		RST("TransferSpecialProgramsNonMentoring") = Request("frmPerformanceTransferSpecialProgramsNonMentoring")
		
		
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

	RST("TransferCommunityBased") = Request("frmPerformanceTransferCommunityBased")
	RST("TransferSchoolBased") = Request("frmPerformanceTransferSchoolBased")
	RST("TransferOtherSiteBased") = Request("frmPerformanceTransferOtherSiteBased")
	RST("TransferGroupMentoring") = Request("frmPerformanceTransferGroupMentoring")
	RST("TransferSpecialProgramsMentoring") = Request("frmPerformanceTransferSpecialProgramsMentoring")
	RST("TransferSpecialProgramsNonMentoring") = Request("frmPerformanceTransferSpecialProgramsNonMentoring")


	
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
	say = "thanks" %>
	
	
	
	
	
	
		
	
	
	
<% ElseIf Request("status") = "editOld" Then
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

function addUpOpenComm()
{
	var box1 = Number(document.frmPerformance.frmPerformancePrevOpenComm.value)
	var box2 = Number(document.frmPerformance.frmPerformanceClosedMatchesCommunityBased.value)
	var box3 = Number(document.frmPerformance.frmPerformanceNewMatchesCommunityBased.value)	
	var box4 = Number(document.frmPerformance.frmPerformanceTransferCommunityBased.value)

	var boxtotal = box1 - box2 + box3 + box4
	document.frmPerformance.frmPerformanceOpenMatchesCommunityBased.value = boxtotal
	
}

function addUpOpenSchool()
{
	var box1 = Number(document.frmPerformance.frmPerformancePrevOpenSchool.value)
	var box2 = Number(document.frmPerformance.frmPerformanceClosedMatchesSchoolBased.value)
	var box3 = Number(document.frmPerformance.frmPerformanceNewMatchesSchoolBased.value)	
	var box4 = Number(document.frmPerformance.frmPerformanceTransferSchoolBased.value)	

	var boxtotal = box1 - box2 + box3 + box4
	document.frmPerformance.frmPerformanceOpenMatchesSchoolBased.value = boxtotal
	
}


function addUpOpenNonSchool()
{
	var box1 = Number(document.frmPerformance.frmPerformancePrevOpenOther.value)
	var box2 = Number(document.frmPerformance.frmPerformanceClosedMatchesOtherSiteBased.value)
	var box3 = Number(document.frmPerformance.frmPerformanceNewMatchesSiteBasedNonSchool.value)	
	var box4 = Number(document.frmPerformance.frmPerformanceTransferOtherSiteBased.value)	

	var boxtotal = box1 - box2 + box3 + box4
	document.frmPerformance.frmPerformanceOpenMatchesOtherSiteBased.value = boxtotal
	
}

function addUpOpenGroup()
{
	var box1 = Number(document.frmPerformance.frmPerformancePrevOpenGroup.value)
	var box2 = Number(document.frmPerformance.frmPerformanceClosedMatchesGroupMentoring.value)
	var box3 = Number(document.frmPerformance.frmPerformanceNewMatchesGroupMentoring.value)	
	var box4 = Number(document.frmPerformance.frmPerformanceTransferGroupMentoring.value)		

	var boxtotal = box1 - box2 + box3 + box4
	document.frmPerformance.frmPerformanceOpenMatchesGroupMentoring.value = boxtotal
	
}

function addUpOpenSpecMent()
{
	var box1 = Number(document.frmPerformance.frmPerformancePrevOpenSpecMent.value)
	var box2 = Number(document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsMentoring.value)
	var box3 = Number(document.frmPerformance.frmPerformanceNewMatchesSpecialProgramsMentoring.value)	
	var box4 = Number(document.frmPerformance.frmPerformanceTransferSpecialProgramsMentoring.value)			

	var boxtotal = box1 - box2 + box3 + box4
	document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsMentoring.value = boxtotal
	
}

function addUpOpenSpecNonMent()
{
	var box1 = Number(document.frmPerformance.frmPerformancePrevOpenSpecNonMent.value)
	var box2 = Number(document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsNonMentoring.value)
	var box3 = Number(document.frmPerformance.frmPerformanceNewMatchesSpecialProgramsNonMentoring.value)	
	var box4 = Number(document.frmPerformance.frmPerformanceTransferSpecialProgramsNonMentoring.value)				
	var boxtotal = box1 - box2 + box3 + box4

	document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsNonMentoring.value = boxtotal
	
}






function validateForm()
{	
	
	var onlyInteger = /^[0-9]+(,[0-9]{3})*$/;
	var PrevOpenComm = new Number(frmPerformance.frmPerformancePrevOpenComm.value)
	var CurCommOpen = new Number(frmPerformance.frmPerformanceOpenMatchesCommunityBased.value)
	var CurCommClosed = new Number(frmPerformance.frmPerformanceClosedMatchesCommunityBased.value)	
	var CurCommTotal = CurCommOpen + CurCommClosed
	
	var PrevOpenSchool = new Number(frmPerformance.frmPerformancePrevOpenSchool.value)
	var CurSchoolOpen = new Number(frmPerformance.frmPerformanceOpenMatchesSchoolBased.value)
	var CurSchoolClosed = new Number(frmPerformance.frmPerformanceClosedMatchesSchoolBased.value)	
	var CurSchoolTotal = CurSchoolOpen + CurSchoolClosed
	
	var PrevOpenOther = new Number(frmPerformance.frmPerformancePrevOpenOther.value)

	var CurOtherOpen = new Number(frmPerformance.frmPerformanceOpenMatchesOtherSiteBased.value)
	var CurOtherClosed = new Number(frmPerformance.frmPerformanceClosedMatchesOtherSiteBased.value)	
	var CurOtherTotal = CurOtherOpen + CurOtherClosed
	
	var PrevOpenGroup = new Number(frmPerformance.frmPerformancePrevOpenGroup.value)
	var CurGroupOpen = new Number(frmPerformance.frmPerformanceOpenMatchesGroupMentoring.value)
	var CurGroupClosed = new Number(frmPerformance.frmPerformanceClosedMatchesGroupMentoring.value)
	var CurGroupTotal = CurGroupOpen + CurGroupClosed
	
	var	PrevOpenSpecMent = new Number(frmPerformance.frmPerformancePrevOpenSpecMent.value)
	var CurSpecMentOpen = new Number(frmPerformance.frmPerformanceOpenMatchesSpecialProgramsMentoring.value)
	var CurSpecMentClosed = new Number(frmPerformance.frmPerformanceClosedMatchesSpecialProgramsMentoring.value)
	var CurSpecMentTotal = CurSpecMentOpen + CurSpecMentClosed
	
	var PrevOpenSpecNonMent = new Number(frmPerformance.frmPerformancePrevOpenSpecNonMent.value)
	var CurSpecNonMentOpen = new Number(frmPerformance.frmPerformanceOpenMatchesSpecialProgramsNonMentoring.value)
	var CurSpecNonMentClosed = new Number(frmPerformance.frmPerformanceClosedMatchesSpecialProgramsNonMentoring.value)
	var CurSpecNonMentTotal = CurSpecNonMentOpen + CurSpecNonMentClosed
	
	var Transfer1 = new Number(frmPerformance.frmPerformanceTransferCommunityBased.value)
	var Transfer2 = new Number(frmPerformance.frmPerformanceTransferSchoolBased.value)	
	var Transfer3 = new Number(frmPerformance.frmPerformanceTransferOtherSiteBased.value)			
	var Transfer4 = new Number(frmPerformance.frmPerformanceTransferGroupMentoring.value)
	var Transfer5 = new Number(frmPerformance.frmPerformanceTransferSpecialProgramsMentoring.value)	
	var Transfer6 = new Number(frmPerformance.frmPerformanceTransferSpecialProgramsNonMentoring.value)		
	var TransferTotal = Transfer1 + Transfer2 + Transfer3 + Transfer4 + Transfer5 + Transfer6


	
	
//	if (CurCommTotal.valueOf() < PrevOpenComm)	
//		{alert( "ERROR:\n\nThe sum of your current OPEN PLUS your current CLOSED Community-Based matches ("+CurCommOpen+"+"+CurCommClosed+") must be greater than your previous month's OPEN Community-Based matches ("+PrevOpenComm+").  Please Correct and re-SAVE.");document.frmPerformance.frmPerformanceOpenMatchesCommunityBased.focus();}

//	else if (CurSchoolTotal.valueOf() < PrevOpenSchool)	
//		{alert( "ERROR:\n\nThe sum of your current OPEN PLUS your CLOSED School-Based matches ("+CurSchoolOpen+"+"+CurSchoolClosed+") must be greater than your previous month's OPEN School-Based matches ("+PrevOpenSchool+").  Please correct and re-SAVE.");document.frmPerformance.frmPerformanceOpenMatchesSchoolBased.focus();}

//	else if (CurOtherTotal.valueOf() < PrevOpenOther)	
//		{alert( "ERROR:\n\nThe sum of your current OPEN PLUS your CLOSED Other Site-Based matches ("+CurOtherOpen+"+"+CurOtherClosed+") must be greater than your previous month's OPEN Other Site-Based matches ("+PrevOpenOther+").  Please correct and re-SAVE.");document.frmPerformance.frmPerformanceOpenMatchesOtherSiteBased.focus();}

	if(document.frmPerformance.frmPerformanceRTBM_UnmatchedChildren.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceRTBM_UnmatchedChildren.focus();}						
		
	else if(document.frmPerformance.frmPerformanceRTBM_UnmatchedVolunteers.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceRTBM_UnmatchedVolunteers.focus();}								

	else if(document.frmPerformance.frmPerformanceRevenue.value == "")
		{alert("Revenue must not be blank");document.frmPerformance.frmPerformanceRevenue.focus();}										
		
	else if(document.frmPerformance.frmPerformanceRevenue.value == 0)
		{alert("Revenue must not be zero.  If you have no revenue for this month, put in 1 dollar to indicate that you did not skip this field");document.frmPerformance.frmPerformanceRevenue.focus();}												
		
	else if(document.frmPerformance.frmPerformanceRevenue.value < 0)
		{alert("Revenue must not be a negative number.");document.frmPerformance.frmPerformanceRevenue.focus();}														
	else if(document.frmPerformance.frmPerformanceSBMVolunteersInEnrollmentProcess.value == "")
		{alert("SBM Volunteers in Enrollment Process must not be blank");document.frmPerformance.frmPerformanceSBMVolunteersInEnrollmentProcess.focus();}
	else if(document.frmPerformance.frmPerformanceSBMAmountRaisedTowardsMatchPledge.value == "")
		{alert("SBM Amount Raised Towards Match Pledge must not be blank");document.frmPerformance.frmPerformanceSBMAmountRaisedTowardsMatchPledge.focus();}		
		
	else if(document.frmPerformance.frmPerformanceAverageMatchLengthCB.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceAverageMatchLengthCB.focus();}				
	else if(document.frmPerformance.frmPerformanceAverageMatchLengthSB.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceAverageMatchLengthSB.focus();}						
	else if(document.frmPerformance.frmPerformanceAverageMatchLengthOSB.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceAverageMatchLengthOSB.focus();}						
	else if(document.frmPerformance.frmPerformanceOpenMatchesCommunityBased.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceOpenMatchesCommunityBased.focus();}		
	else if(document.frmPerformance.frmPerformanceOpenMatchesSchoolBased.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceOpenMatchesSchoolBased.focus();}
	else if(document.frmPerformance.frmPerformanceOpenMatchesOtherSiteBased.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceOpenMatchesOtherSiteBased.focus();}
	else if(document.frmPerformance.frmPerformanceOpenMatchesGroupMentoring.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceOpenMatchesGroupMentoring.focus();}
	else if(document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsMentoring.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsMentoring.focus();}
	else if(document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsNonMentoring.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsNonMentoring.focus();}
	else if(document.frmPerformance.frmPerformanceClosedMatchesCommunityBased.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceClosedMatchesCommunityBased.focus();}
	else if(document.frmPerformance.frmPerformanceClosedMatchesSchoolBased.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceClosedMatchesSchoolBased.focus();}
	else if(document.frmPerformance.frmPerformanceClosedMatchesOtherSiteBased.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceClosedMatchesOtherSiteBased.focus();}
	else if(document.frmPerformance.frmPerformanceClosedMatchesGroupMentoring.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceClosedMatchesGroupMentoring.focus();}
	else if(document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsMentoring.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsMentoring.focus();}
	else if(document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsNonMentoring.value == "")
		{alert("Please complete all form fields");document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsNonMentoring.focus();}

		
		
		
	else if(document.frmPerformance.frmPerformanceCBNumberClosedPrematurely.value == "")
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
		
	else if(document.frmPerformance.frmPerformanceOpenMatchesCommunityBased.value < 0)
		{alert("Open Community Based Matches at the end of the month cannot be less than zero.");document.frmPerformance.frmPerformanceOpenMatchesCommunityBased.focus();}
				
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceOpenMatchesCommunityBased.value)))
		{alert(document.frmPerformance.frmPerformanceOpenMatchesCommunityBased.value + " is an invalid number");document.frmPerformance.frmPerformanceOpenMatchesCommunityBased.focus();}
		
	else if(document.frmPerformance.frmPerformanceOpenMatchesSchoolBased.value < 0)
		{alert("Open School-Based Matches at the end of the month cannot be less than zero.");document.frmPerformance.frmPerformanceOpenMatchesSchoolBased.focus();}

	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceOpenMatchesSchoolBased.value)))
		{alert(document.frmPerformance.frmPerformanceOpenMatchesSchoolBased.value + " is an invalid number");document.frmPerformance.frmPerformanceOpenMatchesSchoolBased.focus();}

	else if(document.frmPerformance.frmPerformanceOpenMatchesOtherSiteBased.value < 0)
		{alert("Open Non-School Site-Based Matches at the end of the month cannot be less than zero.");document.frmPerformance.frmPerformanceOpenMatchesOtherSiteBased.focus();}
		
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceOpenMatchesOtherSiteBased.value)))
		{alert(document.frmPerformance.frmPerformanceOpenMatchesOtherSiteBased.value + " is an invalid number");document.frmPerformance.frmPerformanceOpenMatchesOtherSiteBased.focus();}

	else if(document.frmPerformance.frmPerformanceOpenMatchesGroupMentoring.value < 0)
		{alert("Open Group Mentoring Matches at the end of the month cannot be less than zero.");document.frmPerformance.frmPerformanceOpenMatchesGroupMentoring.focus();}
		
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceOpenMatchesGroupMentoring.value)))
		{alert(document.frmPerformance.frmPerformanceOpenMatchesGroupMentoring.value + " is an invalid number");document.frmPerformance.frmPerformanceOpenMatchesGroupMentoring.focus();}

	else if(document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsMentoring.value < 0)
		{alert("Open Special Programs Mentoring Matches at the end of the month cannot be less than zero.");document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsMentoring.focus();}

	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsMentoring.value)))
		{alert(document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsMentoring.value + " is an invalid number");document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsMentoring.focus();}

	else if(document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsNonMentoring.value < 0)
		{alert("Open Special Programs Non-Mentoring Matches at the end of the month cannot be less than zero.");document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsNonMentoring.focus();}

	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsNonMentoring.value)))
		{alert(document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsNonMentoring.value + " is an invalid number");document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsNonMentoring.focus();}

	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceClosedMatchesCommunityBased.value)))
		{alert(document.frmPerformance.frmPerformanceClosedMatchesCommunityBased.value + " is an invalid number");document.frmPerformance.frmPerformanceClosedMatchesCommunityBased.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceClosedMatchesSchoolBased.value)))
		{alert(document.frmPerformance.frmPerformanceClosedMatchesSchoolBased.value + " is an invalid number");document.frmPerformance.frmPerformanceClosedMatchesSchoolBased.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceClosedMatchesOtherSiteBased.value)))
		{alert(document.frmPerformance.frmPerformanceClosedMatchesOtherSiteBased.value + " is an invalid number");document.frmPerformance.frmPerformanceClosedMatchesOtherSiteBased.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceClosedMatchesGroupMentoring.value)))
		{alert(document.frmPerformance.frmPerformanceClosedMatchesGroupMentoring.value + " is an invalid number");document.frmPerformance.frmPerformanceClosedMatchesGroupMentoring.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsMentoring.value)))
		{alert(document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsMentoring.value + " is an invalid number");document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsMentoring.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsNonMentoring.value)))
		{alert(document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsNonMentoring.value + " is an invalid number");document.frmPerformance.frmPerformanceClosedMatchesSpecialProgramsNonMentoring.focus();}
		
	else if(TransferTotal.valueOf() != 0)
		{alert("Total of all transfers must add up to ZERO.")}	
		
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

<% If say="form" Then %>
	<body onLoad="addUpOpenComm(); addUpOpenSchool(); addUpOpenNonSchool(); addUpOpenGroup(); addUpOpenSpecMent(); addUpOpenSpecNonMent();">
<% else %>
	<body>
<% end if %>

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


<form name="frmPerformance" action="Performance_edit.asp" method="post"> <!-- onsubmit="return submitFormValidate(this)"> -->
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
			<td colspan="7" class="formHeader">PERFORMANCE - CORE BUSINESS - <%= MonthName(Request("m"), False) & " " & Request("y") %></td>
		</tr>

		<tr>
			<td colspan="7" class="formMain"><font color="#ff0000"><div align="center"><strong>Please Note: </strong>After entering your information, you <strong>must</strong> click on the "Save" button at the bottom of the form and wait for the "Thank You" screen or your changes will be lost.</div></font></td>
		</tr>
		
			<tr>
				<td>&nbsp;</td>
				<td align="center" valign="middle" class="formMain">Community Based</td>
				<td align="center" valign="middle" class="formMain">School Based</td>
				<td align="center" valign="middle" class="formMain">Non-School Site Based</td>
				
					<td align="center" valign="middle" class="formMain">Group Mentoring</td>
					<td align="center" valign="middle" class="formMain">Special Programs: Mentoring</td>
					<td align="center" valign="middle" class="formMain">Special Programs: Non-Mentoring</td>
				
			</tr>
			<tr>
			
			<!-- Open Matches at the Beginning of the Month -->
			<tr>
				<td align="center" valign="middle" class="formMain">OPEN/ACTIVE&nbsp;matches<br>on&nbsp;the&nbsp;<strong>FIRST</strong>&nbsp;day&nbsp;of<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				<td align="center" valign="middle" class="formMain" bgcolor="#c0c0c0">
					<input type="text" class="formMain" size="5" value="<%=Request("PrevOpenComm")%>" readonly onFocus="addUpOpenComm();"><br><span class="formSubHead">calculated by system</span>
				</td>
				<td align="center" valign="middle" class="formMain" bgcolor="#c0c0c0">
					<input type="text" class="formMain" size="5" value="<%=Request("PrevOpenSchool")%>" readonly onFocus="addUpOpenSchool();"><br><span class="formSubHead">calculated by system</span>
				</td>
				<td align="center" valign="middle" class="formMain" bgcolor="#c0c0c0">
					<input type="text" class="formMain" size="5" value="<%=Request("PrevOpenOther")%>" readonly "addUpOpenNonSchool();"><br><span class="formSubHead">calculated by system</span>				
				</td>

				<% if Y < 2004 then %>
					<td align="left" valign="middle" class="formMain" bgcolor="#c0c0c0">
						<input type="text" class="formMain" size="5" value="<%=Request("PrevOpenGroup")%>" readonly onFocus="addUpOpenGroup();"><br><span class="formSubHead">calculated by system</span>			
					</td>
					<td align="left" valign="middle" class="formMain" bgcolor="#c0c0c0">
						<input type="text" class="formMain" size="5" value="<%=Request("PrevOpenSpecMent")%>" readonly onFocus="addUpOpenSpecMent();"><br><span class="formSubHead">calculated by system</span>
					</td>
					<td align="left" valign="middle" class="formMain" bgcolor="#c0c0c0">
						<input type="text" class="formMain" size="5" value="<%=Request("PrevOpenSpecNonMent")%>" readonly onFocus="addUpOpenSpecNonMent();"><br><span class="formSubHead">calculated by system</span>					
					</td>
					
				<% else %>
					<td align="center" class="formMain" colspan="3"><span class="formSubHead">No Longer Required</span></td>				
				<% end if %>							
				
				
			</tr>		

			<!-- Matches Closed During the Month -->
			<tr>
				<td align="center" valign="middle" class="formMain">Matches&nbsp;CLOSED&nbsp;during<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesCommunityBased") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesCommunityBased" tabindex="1" onchange="checkForIntegerCommas(this.value); addUpOpenComm();">
				</td>
				
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesSchoolBased") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesSchoolBased" tabindex="2" onchange="checkForIntegerCommas(this.value); addUpOpenSchool();">
				</td>
				
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesOtherSiteBased") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesOtherSiteBased" tabindex="3" onchange="checkForIntegerCommas(this.value); addUpOpenNonSchool();">
				</td>
				
				<% if y < 2004 then %>			
				
					<td align="center" valign="middle" class="formMain">
						<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesGroupMentoring") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesGroupMentoring" tabindex="4" onchange="checkForIntegerCommas(this.value); addUpOpenGroup();">
					</td>
					
					<td align="center" valign="middle" class="formMain">
						<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesSpecialProgramsMentoring") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesSpecialProgramsMentoring" tabindex="5" onchange="checkForIntegerCommas(this.value); addUpOpenSpecMent();">
					</td>
					
					<td align="center" valign="middle" class="formMain">
						<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesSpecialProgramsNonMentoring") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesSpecialProgramsNonMentoring" tabindex="6" onchange="checkForIntegerCommas(this.value); addUpOpenSpecNonMent();">
					</td>
					
				<% else %>
					<!-- Prepopulate the fields -->
					<td align="center" class="formMain" colspan="3"><span class="formSubHead">No Longer Required</span></td>				
					<input type="hidden"  value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesGroupMentoring") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesGroupMentoring">					
					<input type="hidden"  value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesSpecialProgramsMentoring") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesSpecialProgramsMentoring">										
					<input type="hidden"  value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesSpecialProgramsNonMentoring") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesSpecialProgramsNonMentoring">				
				<% end if %>
				
			</tr>
	
			<!-- New Matches Opened During the Month -->		
			<tr>
		
				<td align="center" valign="middle" class="formMain">NEW&nbsp;matches opened<br>during&nbsp;<b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("NewMatchesCommunityBased") %><% Else %>0<% End If %>" name="frmPerformanceNewMatchesCommunityBased" tabindex="7" onchange="checkForIntegerCommas(this.value); addUpOpenComm();">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("NewMatchesSchoolBased") %><% Else %>0<% End If %>" name="frmPerformanceNewMatchesSchoolBased" tabindex="8" onchange="checkForIntegerCommas(this.value); addUpOpenSchool();">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("NewMatchesSiteBasedNonSchool") %><% Else %>0<% End If %>" name="frmPerformanceNewMatchesSiteBasedNonSchool" tabindex="9" onchange="checkForIntegerCommas(this.value); addUpOpenNonSchool();">
				</td>

				<% if Y < 2004 then %>
					<td align="left" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("NewMatchesGroupMentoring") %><% Else %>0<% End If %>" name="frmPerformanceNewMatchesGroupMentoring" tabindex="10" onchange="checkForIntegerCommas(this.value); addUpOpenGroup();">
					</td>
					<td align="left" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("NewMatchesSpecialProgramsMentoring") %><% Else %>0<% End If %>" name="frmPerformanceNewMatchesSpecialProgramsMentoring" tabindex="11" onchange="checkForIntegerCommas(this.value); addUpOpenSpecMent();">
					</td>
					<td align="left" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("NewMatchesSpecialProgramsNonMentoring") %><% Else %>0<% End If %>" name="frmPerformanceNewMatchesSpecialProgramsNonMentoring" tabindex="12" onchange="checkForIntegerCommas(this.value); addUpOpenSpecNonMent()">
					</td>
				<% else %>
					<!-- Prepopulate the Fields -->
					<td align="center" class="formMain" colspan="3"><span class="formSubHead">No Longer Required</span></td>				
					<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("NewMatchesGroupMentoring") %><% Else %>0<% End If %>" name="frmPerformanceNewMatchesGroupMentoring">					
					<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("NewMatchesSpecialProgramsMentoring") %><% Else %>0<% End If %>" name="frmPerformanceNewMatchesSpecialProgramsMentoring">					
					<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("NewMatchesSpecialProgramsNonMentoring") %><% Else %>0<% End If %>" name="frmPerformanceNewMatchesSpecialProgramsNonMentoring">					
				<% end if %>				
				

			
			</tr>					
			
			
			<!-- Match Transfers -->
			<tr>
				<td align="center" valign="middle" class="formMain">Transfer Matches</td>			
			
			<td align="center" valign="middle" class="formMain">
				<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("TransferCommunityBased") %><% Else %>0<% End If %>" name="frmPerformanceTransferCommunityBased" onchange="addUpOpenComm()">
			</td>		
			
			<td align="center" valign="middle" class="formMain">
				<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("TransferSchoolBased") %><% Else %>0<% End If %>" name="frmPerformanceTransferSchoolBased" onchange="addUpOpenSchool()">
			</td>					
			
			<td align="center" valign="middle" class="formMain">
				<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("TransferOtherSiteBased") %><% Else %>0<% End If %>" name="frmPerformanceTransferOtherSiteBased" onchange="addUpOpenNonSchool()">
			</td>
			
			<% if y < 2004 then %>								

				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("TransferGroupMentoring") %><% Else %>0<% End If %>" name="frmPerformanceTransferGroupMentoring" onchange="addUpOpenGroup()">
				</td>		
				
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("TransferSpecialProgramsMentoring") %><% Else %>0<% End If %>" name="frmPerformanceTransferSpecialProgramsMentoring" onchange="addUpOpenSpecMent()">
				</td>					
				
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("TransferSpecialProgramsNonMentoring") %><% Else %>0<% End If %>" name="frmPerformanceTransferSpecialProgramsNonMentoring" onchange="addUpOpenSpecNonMent()">
				</td>



			<% else %>
				<td align="center" class="formMain" colspan="3"><span class="formSubHead">No Longer Required</span></td>				
					<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("TransferGroupMentoring") %><% Else %>0<% End If %>" name="frmPerformanceTransferGroupMentoring">					

					<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("TransferSpecialProgramsMentoring") %><% Else %>0<% End If %>" name="frmPerformanceTransferSpecialProgramsMentoring">					

					<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("TransferSpecialProgramsNonMentoring") %><% Else %>0<% End If %>" name="frmPerformanceTransferSpecialProgramsNonMentoring">									
			<% end if %>
			</tr>			
			
			</tr>			
			
			
			<!-- Open Matches on the Last Day of the Month -->
			<tr>
			
				<td align="center" valign="middle" class="formMain">OPEN/ACTIVE&nbsp;matches<br>on&nbsp;the&nbsp;<strong>LAST</strong>&nbsp;day&nbsp;of<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				
				<td align="center" valign="middle" bgcolor="#c0c0c0" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesCommunityBased") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesCommunityBased" onFocus="addUpOpenComm();" onchange="addUpOpenComm();" readonly><br><span class="formSubHead">calculated by system</span>
				</td>
				
				<td align="center" valign="middle" bgcolor="#c0c0c0" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesSchoolBased") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesSchoolBased" onFocus="addUpOpenSchool();" onchange="addUpOpenSchool();" readonly><br><span class="formSubHead">calculated by system</span>				
				</td>
				
				<td align="center" valign="middle" bgcolor="#c0c0c0" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesOtherSiteBased") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesOtherSiteBased" onFocus="addUpOpenNonSchool();" onchange="addUpOpenNonSchool();" readonly><br><span class="formSubHead">calculated by system</span>
				</td>
				
				<% if y < 2004 then %>								
					<td align="center" valign="middle" bgcolor="#c0c0c0" class="formMain">
						<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesGroupMentoring") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesGroupMentoring" onFocus="addUpOpenGroup();" onchange="addUpOpenGroup();" readonly><br><span class="formSubHead">calculated by system</span>
					</td>
					
					<td align="center" valign="middle" bgcolor="#c0c0c0" class="formMain">
						<input type="text"   class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesSpecialProgramsMentoring") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesSpecialProgramsMentoring" tabindex="9" onchange="addUpOpenSpecMent();" readonly><br><span class="formSubHead">calculated by system</span>
					</td>
					
					<td align="center" valign="middle" bgcolor="#c0c0c0" class="formMain">
						<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesSpecialProgramsNonMentoring") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesSpecialProgramsNonMentoring" tabindex="11" onchange="addUpOpenSpecNonMent();" readonly><br><span class="formSubHead">calculated by system</span>
					</td>					
				<% else %>
					<td align="center" class="formMain" colspan="3"><span class="formSubHead">No Longer Required</span></td>				
						<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesGroupMentoring") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesGroupMentoring">					

						<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesSpecialProgramsMentoring") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesSpecialProgramsMentoring">					

						<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesSpecialProgramsNonMentoring") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesSpecialProgramsNonMentoring">									
				<% end if %>
			</tr>
			
			
			<tr>
				<td align="center" class="formMain" colspan="7"><em><strong>Average Match Length Questions have moved to the SDM Metrics form</strong></em></td>
			</tr>


			<!-- Prepopulate AML Fields -->
					<input type="hidden"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("AverageMatchLengthCB") %><% Else %>0<% End If %>" name="frmPerformanceAverageMatchLengthCB" tabindex="8" >					
					<input type="hidden"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("AverageMatchLengthSB") %><% Else %>0<% End If %>" name="frmPerformanceAverageMatchLengthSB" tabindex="8" >					
					<input type="hidden"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("AverageMatchLengthOSB") %><% Else %>0<% End If %>" name="frmPerformanceAverageMatchLengthOSB" tabindex="8" >					
			




<% If FBIAgency = 1 and ((y > 2002 and m < 7) or (y = 2002 and m > 8) ) Then %>	

		<!-- Faith-Based / Incarcerated Questions -->

		<input type="hidden"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBIandFB") %><% Else %>0<% End If %>" name="frmPerformanceCBIandFB">				
		<input type="hidden"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBIandFB") %><% Else %>0<% End If %>" name="frmPerformanceSBIandFB">				
		<input type="hidden"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBIandFB") %><% Else %>0<% End If %>" name="frmPerformanceOSBIandFB">				
		<input type="hidden"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBInotFB") %><% Else %>0<% End If %>" name="frmPerformanceCBInotFB">				
		<input type="hidden"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBInotFB") %><% Else %>0<% End If %>" name="frmPerformanceSBInotFB">				
		<input type="hidden"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBInotFB") %><% Else %>0<% End If %>" name="frmPerformanceOSBInotFB">				
		<input type="hidden"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBFBnotI") %><% Else %>0<% End If %>" name="frmPerformanceCBFBnotI">				
		<input type="hidden"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBFBnotI") %><% Else %>0<% End If %>" name="frmPerformanceSBFBnotI">				
		<input type="hidden"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OSBFBnotI") %><% Else %>0<% End If %>" name="frmPerformanceOSBFBnotI">				

		<!-- End Faith-Based / Incarcerated Questions -->
<% else %>

		<!-- Puts in zero values for FB/I fields so that javascript validation doesn't choke on nulls -->
		
			<input type="hidden"  class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBIandFB" tabindex="14" onchange="checkForIntegerCommas(this.value);">				
			<input type="hidden"  class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceSBIandFB" tabindex="14" onchange="checkForIntegerCommas(this.value);">				
			<input type="hidden"  class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceOSBIandFB" tabindex="14" onchange="checkForIntegerCommas(this.value);">				
			<input type="hidden"  class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBInotFB" tabindex="14" onchange="checkForIntegerCommas(this.value);">				
			<input type="hidden"  class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceSBInotFB" tabindex="14" onchange="checkForIntegerCommas(this.value);">						
			<input type="hidden"  class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceOSBInotFB" tabindex="14" onchange="checkForIntegerCommas(this.value);">									
			<input type="hidden"  class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceCBFBnotI" tabindex="14" onchange="checkForIntegerCommas(this.value);">												
			<input type="hidden"  class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceSBFBnotI" tabindex="14" onchange="checkForIntegerCommas(this.value);">															
			<input type="hidden"  class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceOSBFBnotI" tabindex="14" onchange="checkForIntegerCommas(this.value);">																		
			
<% End If %>



<%  If (m=6 or m=12) And SBMAgency = 1 And Y <> 2001 Then %>

			<input type="hidden"  class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceSBMVolunteersInEnrollmentProcess">
			<input type="hidden"  class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceSBMAmountRaisedTowardsMatchPledge">					
<% End If %>


<% if (y >= 2002) then %>

		<tr>
			<td colspan="7" class="formHeaderMedium">REVENUE</td>				
		</tr>
		<tr>
			
			
			<td align="center" valign="bottom" class="formMain" colspan="1">Revenue&nbsp;<strong>booked&nbsp;for&nbsp;the&nbsp;Month&nbsp;of&nbsp;<%= MonthName(Request("m"), False) & " " & Request("y") %>&nbsp;&nbsp;</strong><a href="../helpfiles/surveyhelp.asp?HelpID=rev1" onclick="NewWindow(this.href,'name','500','275','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
			<td align="center" valign="middle" class="formMain" colspan="2">
					$<input type="text"  class="formMain" size="8" maxlength="8" value="<% If say = "edit" Then %><%= GetPerformance("Revenue") %><% Else %>0<% End If %>" name="frmPerformanceRevenue" tabindex="15" onchange="checkForIntegerCommas(this.value);">					
			</td>
			<td colspan="4">&nbsp;</td>
		</tr>			

<% End If %>	


<% if m=12 then %>

		<tr>
			<td colspan="7" class="formHeaderMedium">READY TO BE MATCHED</td>
		</tr>
		
		<tr>
			<td align="center" valign="bottom" class="formMain" colspan="1">Total Number of <strong>UNMATCHED Children</strong> as of 12/31/<%=y%>&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=rtbm1" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
			<td align="center" valign="middle" class="formMain" colspan="2"><input type="text"  class="formMain" size="8" maxlength="8" value="<% If say = "edit" Then %><%= GetPerformance("RTBM_UnmatchedChildren") %><% Else %>0<% End If %>" name="frmPerformanceRTBM_UnmatchedChildren" tabindex="15" onchange="checkForIntegerCommas(this.value);"></td>		
			<td colspan="4">&nbsp;</td>
		</tr>
		
		<tr>
			<td align="center" valign="bottom" class="formMain" colspan="1">Total Number of <strong>UNMATCHED Volunteers</strong> as of 12/31/<%=y%>&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=rtbm2" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
			<td align="center" valign="middle" class="formMain" colspan="2"><input type="text"  class="formMain" size="8" maxlength="8" value="<% If say = "edit" Then %><%= GetPerformance("RTBM_UnmatchedVolunteers") %><% Else %>0<% End If %>" name="frmPerformanceRTBM_UnmatchedVolunteers" tabindex="15" onchange="checkForIntegerCommas(this.value);"></td>		
			<td colspan="4">&nbsp;</td>
		</tr>		
		
<% else %>
		<!-- Insert Null values in form to clear validation -->
		<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceRTBM_UnmatchedChildren" onchange="checkForIntegerCommas(this.value);">		
		<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceRTBM_UnmatchedVolunteers" onchange="checkForIntegerCommas(this.value);">				

<% end if %>


<% if (SDMPilot = 1 and ( y > 2002 or (y = 2002 and m >= 7))) or y > 2003 then %>


	<!-- Prepopulate SDM Metrics -->
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Inquiries") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Inquiries">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Inquiries") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Inquiries">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Interviews") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Interviews">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Interviews") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Interviews">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Matched") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Matched">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Matched") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Matched">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_Number_Comm") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InquiryToInterview_Number_Comm">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_AveDays_Comm") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InquiryToInterview_AveDays_Comm">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_Number_School") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InquiryToInterview_Number_School">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_AveDays_School") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InquiryToInterview_AveDays_School">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_Number_Other") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InquiryToInterview_Number_Other">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InquiryToInterview_AveDays_Other") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InquiryToInterview_AveDays_Other">				
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InterviewToMatched_Number_Comm") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InterviewToMatched_Number_Comm">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InterviewToMatched_AveDays_Comm") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InterviewToMatched_AveDays_Comm">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InterviewToMatched_Number_School") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InterviewToMatched_Number_School">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InterviewToMatched_AveDays_School") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InterviewToMatched_AveDays_School">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InterviewToMatched_Number_Other") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InterviewToMatched_Number_Other">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Vol_InterviewToMatched_AveDays_Other") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Vol_InterviewToMatched_AveDays_Other">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_Number_Comm") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InquiryToInterview_Number_Comm">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_AveDays_Comm") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InquiryToInterview_AveDays_Comm">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_Number_School") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InquiryToInterview_Number_School">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_AveDays_School") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InquiryToInterview_AveDays_School">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_Number_Other") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InquiryToInterview_Number_Other">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InquiryToInterview_AveDays_Other") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InquiryToInterview_AveDays_Other">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InterviewToMatched_Number_Comm") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InterviewToMatched_Number_Comm">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InterviewToMatched_AveDays_Comm") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InterviewToMatched_AveDays_Comm">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InterviewToMatched_Number_School") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InterviewToMatched_Number_School">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InterviewToMatched_AveDays_School") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InterviewToMatched_AveDays_School">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InterviewToMatched_Number_Other") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InterviewToMatched_Number_Other">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("ProcTim_Youth_InterviewToMatched_AveDays_Other") %><% Else %>0<% End If %>" name="frmPerformanceProcTim_Youth_InterviewToMatched_AveDays_Other">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("Freq_Under3Months_Comm") %><% Else %>0<% End If %>" name="frmPerformanceFreq_Under3Months_Comm">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("Freq_Under3Months_School") %><% Else %>0<% End If %>" name="frmPerformanceFreq_Under3Months_School">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("Freq_Under3Months_Other") %><% Else %>0<% End If %>" name="frmPerformanceFreq_Under3Months_Other">				
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("Freq_3To5Months_Comm") %><% Else %>0<% End If %>" name="frmPerformanceFreq_3To5Months_Comm">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("Freq_3To5Months_School") %><% Else %>0<% End If %>" name="frmPerformanceFreq_3To5Months_School">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("Freq_3To5Months_Other") %><% Else %>0<% End If %>" name="frmPerformanceFreq_3To5Months_Other">				
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("Freq_6To8Months_Comm") %><% Else %>0<% End If %>" name="frmPerformanceFreq_6To8Months_Comm">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("Freq_6To8Months_School") %><% Else %>0<% End If %>" name="frmPerformanceFreq_6To8Months_School">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("Freq_6To8Months_Other") %><% Else %>0<% End If %>" name="frmPerformanceFreq_6To8Months_Other">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("Freq_9To11Months_Comm") %><% Else %>0<% End If %>" name="frmPerformanceFreq_9To11Months_Comm">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("Freq_9To11Months_School") %><% Else %>0<% End If %>" name="frmPerformanceFreq_9To11Months_School">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("Freq_9To11Months_Other") %><% Else %>0<% End If %>" name="frmPerformanceFreq_9To11Months_Other">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("Freq_12To23Months_Comm") %><% Else %>0<% End If %>" name="frmPerformanceFreq_12To23Months_Comm">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("Freq_12To23Months_School") %><% Else %>0<% End If %>" name="frmPerformanceFreq_12To23Months_School">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("Freq_12To23Months_Other") %><% Else %>0<% End If %>" name="frmPerformanceFreq_12To23Months_Other">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("Freq_24OrMoreMonths_Comm") %><% Else %>0<% End If %>" name="frmPerformanceFreq_24OrMoreMonths_Comm">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("Freq_24OrMoreMonths_School") %><% Else %>0<% End If %>" name="frmPerformanceFreq_24OrMoreMonths_School">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("Freq_24OrMoreMonths_Other") %><% Else %>0<% End If %>" name="frmPerformanceFreq_24OrMoreMonths_Other">			
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("Volunteers_ReMatched") %><% Else %>0<% End If %>" name="frmPerformanceVolunteers_ReMatched">
	
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Inquiries_CB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Inquiries_CB">	
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Inquiries_SB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Inquiries_SB">		
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Inquiries_OSB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Inquiries_OSB">		
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Interviews_CB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Interviews_CB">	
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Interviews_SB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Interviews_SB">		
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Interviews_OSB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Interviews_OSB">		
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Matched_CB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Matched_CB">		
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Matched_SB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Matched_SB">			
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Matched_OSB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Matched_OSB">			

	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Inquiries_CB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Inquiries_CB">	
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Inquiries_SB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Inquiries_SB">		
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Inquiries_OSB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Inquiries_OSB">		
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Interviews_CB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Interviews_CB">	
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Interviews_SB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Interviews_SB">		
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Interviews_OSB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Interviews_OSB">		
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Matched_CB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Matched_CB">		
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Matched_SB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Matched_SB">			
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Matched_OSB") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Matched_OSB">			

	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("OSBNumberClosedPrematurely") %><% Else %>0<% End If %>" name="frmPerformanceOSBNumberClosedPrematurely">			
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("OSBChildParentStatusChange") %><% Else %>0<% End If %>" name="frmPerformanceOSBChildParentStatusChange">	
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("OSBVolunteerStatusChange") %><% Else %>0<% End If %>" name="frmPerformanceOSBVolunteerStatusChange">		
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("OSBChildParentDissatisfaction") %><% Else %>0<% End If %>" name="frmPerformanceOSBChildParentDissatisfaction">			
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("OSBVolunteerDissatisfaction") %><% Else %>0<% End If %>" name="frmPerformanceOSBVolunteerDissatisfaction">				
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("SBTotalOpened6MonthsAgo") %><% Else %>0<% End If %>" name="frmPerformanceSBTotalOpened6MonthsAgo">					
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("SBNumberStillOpen") %><% Else %>0<% End If %>" name="frmPerformanceSBNumberStillOpen">						
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("OSBTotalOpened6MonthsAgo") %><% Else %>0<% End If %>" name="frmPerformanceOSBTotalOpened6MonthsAgo">							
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("OSBNumberStillOpen") %><% Else %>0<% End If %>" name="frmPerformanceOSBNumberStillOpen">								
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("OSBPOEAggregateScore") %><% Else %>0<% End If %>" name="frmPerformanceOSBPOEAggregateScore">	
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("OSBPOECount") %><% Else %>0<% End If %>" name="frmPerformanceOSBPOECount">		
				
	
	
			
<!-- Additional SDM Fields for October 2003 and beyond -->

<% if (SDMPilot = 1 and(y=2003 and m >= 10)) or (y > 2003) then %>

	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("CBNumberClosedPrematurely") %><% Else %>0<% End If %>" name="frmPerformanceCBNumberClosedPrematurely">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("SBNumberClosedPrematurely") %><% Else %>0<% End If %>" name="frmPerformanceSBNumberClosedPrematurely">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("CBChildParentStatusChange") %><% Else %>0<% End If %>" name="frmPerformanceCBChildParentStatusChange">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("SBChildParentStatusChange") %><% Else %>0<% End If %>" name="frmPerformanceSBChildParentStatusChange">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("CBVolunteerStatusChange") %><% Else %>0<% End If %>" name="frmPerformanceCBVolunteerStatusChange">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("SBVolunteerStatusChange") %><% Else %>0<% End If %>" name="frmPerformanceSBVolunteerStatusChange">				
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("CBChildParentDissatisfaction") %><% Else %>0<% End If %>" name="frmPerformanceCBChildParentDissatisfaction">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("SBChildParentDissatisfaction") %><% Else %>0<% End If %>" name="frmPerformanceSBChildParentDissatisfaction">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("CBVolunteerDissatisfaction") %><% Else %>0<% End If %>" name="frmPerformanceCBVolunteerDissatisfaction">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("SBVolunteerDissatisfaction") %><% Else %>0<% End If %>" name="frmPerformanceSBVolunteerDissatisfaction">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("CBTotalOpened6MonthsAgo") %><% Else %>0<% End If %>" name="frmPerformanceCBTotalOpened6MonthsAgo">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("CBNumberStillOpen") %><% Else %>0<% End If %>" name="frmPerformanceCBNumberStillOpen">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("EnrollmentSatAvgScore") %><% Else %>0<% End If %>" name="frmPerformanceEnrollmentSatAvgScore">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("EnrollmentSatCount") %><% Else %>0<% End If %>" name="frmPerformanceEnrollmentSatCount">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("MatchSatAvgScore") %><% Else %>0<% End If %>" name="frmPerformanceMatchSatAvgScore">		
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("MatchSatCount") %><% Else %>0<% End If %>" name="frmPerformanceMatchSatCount">								
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("CBPOEAggregateScore") %><% Else %>0<% End If %>" name="frmPerformanceCBPOEAggregateScore">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("SBPOEAggregateScore") %><% Else %>0<% End If %>" name="frmPerformanceSBPOEAggregateScore">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("CBPOECount") %><% Else %>0<% End If %>" name="frmPerformanceCBPOECount">
	<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("SBPOECount") %><% Else %>0<% End If %>" name="frmPerformanceSBPOECount">

			
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

			
<!--
			<tr>
				<TD colspan="7" class="formHeaderMedium">POE&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm9" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmark_purplesmall.gif" alt="" width="15" height="16" border="0"></a></TD>	
			</tr>				
			
			<tr>
				<td>&nbsp;</td>
				<td class="formMain" colspan="4" align="center"><strong>Community-Based</strong></td>
				<td class="formMain" colspan="4" align="center"><strong>School-Based</strong></td>				
			</tr>
			
			<tr>
				<td>&nbsp;</td>
				<td class="formMain" colspan="2" align="center">Number</td>
				<td class="formMain" colspan="2" align="center">Average</td>
				<td class="formMain" colspan="2" align="center">Number</td>
				<td class="formMain" colspan="2" align="center">Average</td>				
			</tr>
			
			<tr>
				<td class="formMain">Confidence</td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% ' If say = "edit" Then %><% '= GetPerformance("POE_Confidence_Number_Comm") %><% ' Else %>0<% ' End If %>" name="frmPerformancePOE_Confidence_Number_Comm" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% ' If say = "edit" Then %><% '= GetPerformance("POE_Confidence_Ave_Comm") %><% ' Else %>0<% ' End If %>" name="frmPerformancePOE_Confidence_Ave_Comm" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% ' If say = "edit" Then %><% '= GetPerformance("POE_Confidence_Number_School") %><% ' Else %>0<% ' End If %>" name="frmPerformancePOE_Confidence_Number_School" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% ' If say = "edit" Then %><% '= GetPerformance("POE_Confidence_Ave_School") %><% ' Else %>0<% ' End If %>" name="frmPerformancePOE_Confidence_Ave_School" onchange="checkForIntegerCommas(this.value);"></td>
			</tr>
			
			<tr>
				<td class="formMain">Competence</td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% ' If say = "edit" Then %><% '= GetPerformance("POE_Competence_Number_Comm") %><% ' Else %>0<% ' End If %>" name="frmPerformancePOE_Competence_Number_Comm" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% ' If say = "edit" Then %><% '= GetPerformance("POE_Competence_Ave_Comm") %><% ' Else %>0<% ' End If %>" name="frmPerformancePOE_Competence_Ave_Comm" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% ' If say = "edit" Then %><% '= GetPerformance("POE_Competence_Number_School") %><% ' Else %>0<% ' End If %>" name="frmPerformancePOE_Competence_Number_School" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% ' If say = "edit" Then %><% '= GetPerformance("POE_Competence_Ave_School") %><% ' Else %>0<% ' End If %>" name="frmPerformancePOE_Competence_Ave_School" onchange="checkForIntegerCommas(this.value);"></td>
			</tr>			
			
			<tr>
				<td class="formMain">Caring</td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% ' If say = "edit" Then %><% '= GetPerformance("POE_Caring_Number_Comm") %><% ' Else %>0<% ' End If %>" name="frmPerformancePOE_Caring_Number_Comm" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% ' If say = "edit" Then %><% '= GetPerformance("POE_Caring_Ave_Comm") %><% ' Else %>0<% ' End If %>" name="frmPerformancePOE_Caring_Ave_Comm" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% ' If say = "edit" Then %><% '= GetPerformance("POE_Caring_Number_School") %><% ' Else %>0<% ' End If %>" name="frmPerformancePOE_Caring_Number_School" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="2" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% ' If say = "edit" Then %><% '= GetPerformance("POE_Caring_Ave_School") %><% ' Else %>0<% ' End If %>" name="frmPerformancePOE_Caring_Ave_School" onchange="checkForIntegerCommas(this.value);"></td>
			</tr>			
			
			<tr>
				<TD colspan="7" class="formHeaderMedium">VOLUNTEER SATISFACTION&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm10" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmark_purplesmall.gif" alt="" width="15" height="16" border="0"></a></TD>	
			</tr>														
			
			<tr>
				<td>&nbsp;</td>
				<td colspan="4" class="formMain" align="center"><strong>Number</strong></td>
				<td colspan="4" class="formMain" align="center"><strong>Average</strong></td>				
			</tr>
			
			<tr>
				<td class="formMain">Post Enrollment</td>
				<td class="formMain" align="center" colspan="4"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% ' If say = "edit" Then %><% '= GetPerformance("VolSat_PostEnrollment_Number") %><% ' Else %>0<% ' End If %>" name="frmPerformanceVolSat_PostEnrollment_Number" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" align="center" colspan="4"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% ' If say = "edit" Then %><% '= GetPerformance("VolSat_PostEnrollment_Ave") %><% ' Else %>0<% ' End If %>" name="frmPerformanceVolSat_PostEnrollment_Ave" onchange="checkForIntegerCommas(this.value);"></td>
			</tr>
			
			<tr>
				<td class="formMain">Satisfaction Questionnaire</td>
				<td class="formMain" align="center" colspan="4"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% ' If say = "edit" Then %><% '= GetPerformance("VolSat_SatQuest_Number") %><% ' Else %>0<% ' End If %>" name="frmPerformanceVolSat_SatQuest_Number" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" align="center" colspan="4"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% ' If say = "edit" Then %><% '= GetPerformance("VolSat_SatQuest_Ave") %><% ' Else %>0<% ' End If %>" name="frmPerformanceVolSat_SatQuest_Ave" onchange="checkForIntegerCommas(this.value);"></td>
			</tr>			
-->			

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
			
<!-- Prepopulate Partnership Fields -->

<% ' if (y = 2003 and (m=4 or m=11)) or (y > 2003 and (m=5 or m=11)) then %>
<%  if y > 2003 and (m=5 or m=11) then %>

		<input type="hidden"  value="<% If say = "edit" Then %><%= GetPerformance("AlphaCommunityBased") %><% Else %>0<% End If %>" name="frmPerformanceAlphaCommunityBased">
		<input type="hidden"  value="<% If say = "edit" Then %><%= GetPerformance("AlphaSchoolBased") %><% Else %>0<% End If %>" name="frmperformanceAlphaSchoolBased">
		<input type="hidden"  value="<% If say = "edit" Then %><%= GetPerformance("AlphaOtherSiteBased") %><% Else %>0<% End If %>" name="frmperformanceAlphaOtherSiteBased">
		<input type="hidden" name="frmperformanceAlphaNotPartnering" value="<% If say = "edit" Then %><%=(GetPerformance("AlphaNotPartnering"))%><% Else %>0<% End If %> ">	
		<input type="hidden" name="frmperformanceAlphainterest" value="<% If say = "edit" Then%><%=(GetPerformance("AlphaInterest"))%><% Else %>0<% End If %>">
	
		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("LionsCommunityBased") %><% Else %>0<% End If %>" name="frmPerformanceLionsCommunityBased">	
		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("LionsSchoolBased") %><% Else %>0<% End If %>" name="frmperformanceLionsSchoolBased">	
		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("LionsOtherSiteBased") %><% Else %>0<% End If %>" name="frmperformanceLionsOtherSiteBased">		
		<input type="hidden" name="frmperformanceLionsNotPartnering" value="<% If say = "edit" Then %><%=(GetPerformance("LionsNotPartnering"))%><% Else %>0<% End If %> ">	
		<input type="hidden" name="frmperformanceLionsinterest" value="<% If say = "edit" Then%><%=(GetPerformance("Lionsinterest"))%><% Else %>0<% End If %>">		

		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("RotaryCommunityBased") %><% Else %>0<% End If %>" name="frmPerformanceRotaryCommunityBased">	
		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("RotarySchoolBased") %><% Else %>0<% End If %>" name="frmperformanceRotarySchoolBased">	
		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("RotaryOtherSiteBased") %><% Else %>0<% End If %>" name="frmperformanceRotaryOtherSiteBased">		
		<input type="hidden" name="frmperformanceRotaryNotPartnering" value="<% If say = "edit" Then %><% if isnull(GetPerformance("RotaryNotPartnering")) then%>0<%Else%><%=(GetPerformance("RotaryNotPartnering"))%><%End If%><% Else %>0<% End If %> ">	
		<input type="hidden" name="frmperformanceRotaryinterest" value="<% If say = "edit" Then %><% if isnull(GetPerformance("Rotaryinterest")) then%>0<%Else%><%=(GetPerformance("Rotaryinterest"))%><%End If%><% Else %>0<% End If %> ">	

		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("KiwanisCommunityBased") %><% Else %>0<% End If %>" name="frmPerformanceKiwanisCommunityBased">	
		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("KiwanisSchoolBased") %><% Else %>0<% End If %>" name="frmperformanceKiwanisSchoolBased">	
		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("KiwanisOtherSiteBased") %><% Else %>0<% End If %>" name="frmperformanceKiwanisOtherSiteBased">		
		<input type="hidden" name="frmperformanceKiwanisNotPartnering" value="<% If say = "edit" Then %><% if isnull(GetPerformance("KiwanisNotPartnering")) then%>0<%Else%><%=(GetPerformance("KiwanisNotPartnering"))%><%End If%><% Else %>0<% End If %> ">	
		<input type="hidden" name="frmperformanceKiwanisinterest" value="<% If say = "edit" Then %><% if isnull(GetPerformance("Kiwanisinterest")) then%>0<%Else%><%=(GetPerformance("Kiwanisinterest"))%><%End If%><% Else %>0<% End If %> ">	

		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("OptimistCommunityBased") %><% Else %>0<% End If %>" name="frmPerformanceOptimistCommunityBased">	
		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("OptimistSchoolBased") %><% Else %>0<% End If %>" name="frmperformanceOptimistSchoolBased">	
		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("OptimistOtherSiteBased") %><% Else %>0<% End If %>" name="frmperformanceOptimistOtherSiteBased">		
		<input type="hidden" name="frmperformanceOptimistNotPartnering" value="<% If say = "edit" Then %><% if isnull(GetPerformance("OptimistNotPartnering")) then%>0<%Else%><%=(GetPerformance("OptimistNotPartnering"))%><%End If%><% Else %>0<% End If %> ">	
		<input type="hidden" name="frmperformanceOptimistinterest" value="<% If say = "edit" Then %><% if isnull(GetPerformance("Optimistinterest")) then%>0<%Else%><%=(GetPerformance("Optimistinterest"))%><%End If%><% Else %>0<% End If %> ">	

		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("AARPCommunityBased") %><% Else %>0<% End If %>" name="frmPerformanceAARPCommunityBased">	
		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("AARPSchoolBased") %><% Else %>0<% End If %>" name="frmperformanceAARPSchoolBased">	
		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("AARPOtherSiteBased") %><% Else %>0<% End If %>" name="frmperformanceAARPOtherSiteBased">		
		<input type="hidden" name="frmperformanceAARPNotPartnering" value="<% If say = "edit" Then %><% if isnull(GetPerformance("AARPNotPartnering")) then%>0<%Else%><%=(GetPerformance("AARPNotPartnering"))%><%End If%><% Else %>0<% End If %> ">	
		<input type="hidden" name="frmperformanceAARPinterest" value="<% If say = "edit" Then %><% if isnull(GetPerformance("AARPinterest")) then%>0<%Else%><%=(GetPerformance("AARPinterest"))%><%End If%><% Else %>0<% End If %> ">	

		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("AlphaRating") %><% Else %>0<% End If %>" name="frmPerformanceAlphaRating">	
		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("LionsRating") %><% Else %>0<% End If %>" name="frmPerformanceLionsRating">			
		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("RotaryRating") %><% Else %>0<% End If %>" name="frmPerformanceRotaryRating">					
		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("KiwanisRating") %><% Else %>0<% End If %>" name="frmPerformanceKiwanisRating">							
		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("OptimistRating") %><% Else %>0<% End If %>" name="frmPerformanceOptimistRating">							
		<input type="hidden" value="<% If say = "edit" Then %><%= GetPerformance("AARPRating") %><% Else %>0<% End If %>" name="frmPerformanceAARPRating">

		<input type="hidden" name="frmperformanceAlphaFunding" value="<% If say = "edit" Then %><% if isnull(GetPerformance("AlphaFunding")) then%>0<%Else%><%=(GetPerformance("AlphaFunding"))%><%End If%><% Else %>0<% End If %> ">		
		<input type="hidden" name="frmperformanceAlphaProgramInitiative" value="<% If say = "edit" Then %><% if isnull(GetPerformance("AlphaProgramInitiative")) then%>0<%Else%><%=(GetPerformance("AlphaProgramInitiative"))%><%End If%><% Else %>0<% End If %> ">				
		<input type="hidden" name="frmperformanceAlphaLeadershipInvolvement" value="<% If say = "edit" Then %><% if isnull(GetPerformance("AlphaLeadershipInvolvement")) then%>0<%Else%><%=(GetPerformance("AlphaLeadershipInvolvement"))%><%End If%><% Else %>0<% End If %> ">						
	
		<input type="hidden" name="frmperformanceAlphaUndergradChapterName" value="<% If say = "edit" Then %><%= GetPerformance("AlphaUndergradChapterName") %><% Else %> <% End If %>">
		<input type="hidden" name="frmperformanceAlphaUndergradChapterCity" value="<% If say = "edit" Then %><%= GetPerformance("AlphaUndergradChapterCity") %><% Else %> <% End If %>">	
		<input type="hidden" name="frmperformanceAlphaUndergradChapterState" value="<% If say = "edit" Then %><%= GetPerformance("AlphaUndergradChapterState") %><% Else %> <% End If %>">			
			
		<input type="hidden" name="frmperformanceAlphaAlumniChapterName" value="<% If say = "edit" Then %><%= GetPerformance("AlphaAlumniChapterName") %><% Else %> <% End If %>">
		<input type="hidden" name="frmperformanceAlphaAlumniChapterCity" value="<% If say = "edit" Then %><%= GetPerformance("AlphaAlumniChapterCity") %><% Else %> <% End If %>">	
		<input type="hidden" name="frmperformanceAlphaAlumniChapterState" value="<% If say = "edit" Then %><%= GetPerformance("AlphaAlumniChapterState") %><% Else %> <% End If %>">			


<% else %>

	<!-- Eliminate Null Values -->	
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceAlphaCommunityBased" onchange="checkForIntegerCommas(this.value);">	
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceAlphaSchoolBased" onchange="checkForIntegerCommas(this.value);">		
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceAlphaOtherSiteBased" onchange="checkForIntegerCommas(this.value);">	
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceAlphaNotPartnering" onchange="checkForIntegerCommas(this.value);">	
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceAlphaInterest" onchange="checkForIntegerCommas(this.value);">	

	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceLionsCommunityBased" onchange="checkForIntegerCommas(this.value);">	
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceLionsSchoolBased" onchange="checkForIntegerCommas(this.value);">		
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceLionsOtherSiteBased" onchange="checkForIntegerCommas(this.value);">		
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceLionsNotPartnering" onchange="checkForIntegerCommas(this.value);">	
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceLionsInterest" onchange="checkForIntegerCommas(this.value);">	

	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceRotaryCommunityBased" onchange="checkForIntegerCommas(this.value);">	
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceRotarySchoolBased" onchange="checkForIntegerCommas(this.value);">		
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceRotaryOtherSiteBased" onchange="checkForIntegerCommas(this.value);">		
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceRotaryNotPartnering" onchange="checkForIntegerCommas(this.value);">	
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceRotaryInterest" onchange="checkForIntegerCommas(this.value);">	
	
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceKiwanisCommunityBased" onchange="checkForIntegerCommas(this.value);">	
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceKiwanisSchoolBased" onchange="checkForIntegerCommas(this.value);">		
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceKiwanisOtherSiteBased" onchange="checkForIntegerCommas(this.value);">		
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceKiwanisNotPartnering" onchange="checkForIntegerCommas(this.value);">	
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceKiwanisInterest" onchange="checkForIntegerCommas(this.value);">		
	
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceOptimistCommunityBased" onchange="checkForIntegerCommas(this.value);">	
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceOptimistSchoolBased" onchange="checkForIntegerCommas(this.value);">		
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceOptimistOtherSiteBased" onchange="checkForIntegerCommas(this.value);">		
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceOptimistNotPartnering" onchange="checkForIntegerCommas(this.value);">	
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceOptimistInterest" onchange="checkForIntegerCommas(this.value);">		

	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceAARPCommunityBased" onchange="checkForIntegerCommas(this.value);">	
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceAARPSchoolBased" onchange="checkForIntegerCommas(this.value);">		
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmPerformanceAARPOtherSiteBased" onchange="checkForIntegerCommas(this.value);">		
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceAARPNotPartnering" onchange="checkForIntegerCommas(this.value);">	
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceAARPInterest" onchange="checkForIntegerCommas(this.value);">		

	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceAlphaRating">		
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceLionsRating">			
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceRotaryRating">			
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceKiwanisRating">		
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceOptimistRating">				
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceAARPRating">		
	
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceAlphaFunding">		
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceAlphaProgramInitiative">			
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceAlphaLeadershipInvolvement">				
	
	<input type="hidden" class="formMain" size="5" maxlength="10" value=" " name="frmperformanceAlphaUndergradChapterName">		
	<input type="hidden" class="formMain" size="5" maxlength="10" value=" " name="frmperformanceAlphaUndergradChapterCity">			
	<input type="hidden" class="formMain" size="5" maxlength="10" value=" " name="frmperformanceAlphaUndergradChapterState">			
	
	<input type="hidden" class="formMain" size="5" maxlength="10" value=" " name="frmperformanceAlphaAlumniChapterName">		
	<input type="hidden" class="formMain" size="5" maxlength="10" value=" " name="frmperformanceAlphaAlumniChapterCity">			
	<input type="hidden" class="formMain" size="5" maxlength="10" value=" " name="frmperformanceAlphaAlumniChapterState">
		
<% end if %>
<!-- PARTNERSHIP QUESTIONNAIRE END -->		

<!-- ADD PREVIOUS MONTH'S MATCH FIELDS TO FORM FOR COMPARISON -->
<input type="hidden" class="formMain" size="5" maxlength="10" value="<%=Request("PrevOpenComm")%>" name="frmPerformancePrevOpenComm" onchange="checkForIntegerCommas(this.value);">	
<input type="hidden" class="formMain" size="5" maxlength="10" value="<%=Request("PrevOpenSchool")%>" name="frmPerformancePrevOpenSchool" onchange="checkForIntegerCommas(this.value);">	
<input type="hidden" class="formMain" size="5" maxlength="10" value="<%=Request("PrevOpenOther")%>" name="frmPerformancePrevOpenOther" onchange="checkForIntegerCommas(this.value);">	
<input type="hidden" class="formMain" size="5" maxlength="10" value="<%=Request("PrevOpenGroup")%>" name="frmPerformancePrevOpenGroup" onchange="checkForIntegerCommas(this.value);">	
<input type="hidden" class="formMain" size="5" maxlength="10" value="<%=Request("PrevOpenSpecMent")%>" name="frmPerformancePrevOpenSpecMent" onchange="checkForIntegerCommas(this.value);">	
<input type="hidden" class="formMain" size="5" maxlength="10" value="<%=Request("PrevOpenSpecNonMent")%>" name="frmPerformancePrevOpenSpecNonMent" onchange="checkForIntegerCommas(this.value);">	

<!-- ADD FIELD FOR VALIDATING AVERAGE MATCH LENGTH POST-2002 -->
<% if y >= 2003 then %>
	<input type="hidden" class="formMain" size="5" maxlength="10" value="1" name="frmperformanceValidateAML">		
<% else %>
	<input type="hidden" class="formMain" size="5" maxlength="10" value="0" name="frmperformanceValidateAML">		
<% end if %>
		<tr>
				<td colspan="7" class="formHeader">
				<% if SBMAgency <> 1 or (SBMAgency = 1 and m <> 6 and m <> 12 )then %>
					<input type="hidden"  value="0" name="frmPerformanceSBMVolunteersInEnrollmentProcess" tabindex="14" onchange="checkForIntegerCommas(this.value);">
					<input type="hidden"  value="0" name="frmPerformanceSBMAmountRaisedTowardsMatchPledge" tabindex="14" onchange="checkForIntegerCommas(this.value);">					
				<% end if %>

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
