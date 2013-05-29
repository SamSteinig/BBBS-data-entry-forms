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
	
	
	
	if (CurCommTotal.valueOf() < PrevOpenComm)	
		{alert( "ERROR:\n\nThe sum of your current OPEN PLUS your current CLOSED Community-Based matches ("+CurCommOpen+"+"+CurCommClosed+") must be greater than your previous month's OPEN Community-Based matches ("+PrevOpenComm+").  Please Correct and re-SAVE.");document.frmPerformance.frmPerformanceOpenMatchesCommunityBased.focus();}

	else if (CurSchoolTotal.valueOf() < PrevOpenSchool)	
		{alert( "ERROR:\n\nThe sum of your current OPEN PLUS your CLOSED School-Based matches ("+CurSchoolOpen+"+"+CurSchoolClosed+") must be greater than your previous month's OPEN School-Based matches ("+PrevOpenSchool+").  Please correct and re-SAVE.");document.frmPerformance.frmPerformanceOpenMatchesSchoolBased.focus();}

	else if (CurOtherTotal.valueOf() < PrevOpenOther)	
		{alert( "ERROR:\n\nThe sum of your current OPEN PLUS your CLOSED Other Site-Based matches ("+CurOtherOpen+"+"+CurOtherClosed+") must be greater than your previous month's OPEN Other Site-Based matches ("+PrevOpenOther+").  Please correct and re-SAVE.");document.frmPerformance.frmPerformanceOpenMatchesOtherSiteBased.focus();}

	else if (CurGroupTotal.valueOf() < PrevOpenGroup)	
		{alert( "ERROR:\n\nThe sum of your current OPEN PLUS your CLOSED Group matches ("+CurGroupOpen+"+"+CurGroupClosed+") must be greater than your previous month's OPEN Group-Based matches("+PrevOpenGroup+").  Please correct and re-SAVE.");document.frmPerformance.frmPerformanceOpenMatchesGroupMentoring.focus();}		

	else if (CurSpecMentTotal.valueOf() < PrevOpenSpecMent)	
		{alert( "ERROR:\n\nThe sum of your current OPEN PLUS your CLOSED Special Programs Mentoring Matches ("+CurSpecMentOpen+"+"+CurSpecMentClosed+") must be greater than your previous month's OPEN Special Programs Mentoring Matches ("+PrevOpenSpecMent+").  Please correct and re-SAVE.");document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsMentoring.focus();}				

	else if (CurSpecNonMentTotal.valueOf() < PrevOpenSpecNonMent)	
		{alert( "ERROR:\n\nThe sum of your current OPEN PLUS your CLOSED Special Programs Non-Mentoring Matches ("+CurSpecNonMentOpen+"+"+CurSpecNonMentClosed+") must be greater than your previous month's OPEN Special Programs Non-Mentoring Matches ("+PrevOpenSpecNonMent+").  Please correct and re-SAVE.");document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsNonMentoring.focus();}				
			
	else if(document.frmPerformance.frmperformanceValidateAML.value == 1 && document.frmPerformance.frmPerformanceAverageMatchLengthCB.value > 0 && document.frmPerformance.frmPerformanceClosedMatchesCommunityBased.value == 0)
		{alert("Average Match Length for Community-Based matches cannot be greater than 0 if you did not close any matches during the current month.");document.frmPerformance.frmPerformanceAverageMatchLengthCB.focus();}

	else if(document.frmPerformance.frmperformanceValidateAML.value == 1 && document.frmPerformance.frmPerformanceAverageMatchLengthCB.value == 0 && document.frmPerformance.frmPerformanceClosedMatchesCommunityBased.value > 0)
		{alert("Average Match Length for Community-Based matches must be greater than 0 because you closed matches during the current month.");document.frmPerformance.frmPerformanceAverageMatchLengthCB.focus();}		
		
	else if(document.frmPerformance.frmperformanceValidateAML.value == 1 && document.frmPerformance.frmPerformanceAverageMatchLengthSB.value > 0 && document.frmPerformance.frmPerformanceClosedMatchesSchoolBased.value == 0)
		{alert("Average Match Length for School-Based matches cannot be greater than 0 if you did not close any matches during the current month.");document.frmPerformance.frmPerformanceAverageMatchLengthSB.focus();}		
		
	else if(document.frmPerformance.frmperformanceValidateAML.value == 1 && document.frmPerformance.frmPerformanceAverageMatchLengthSB.value == 0 && document.frmPerformance.frmPerformanceClosedMatchesSchoolBased.value > 0)
		{alert("Average Match Length for School-Based matches must be greater than 0 because you closed matches during the current month.");document.frmPerformance.frmPerformanceAverageMatchLengthSB.focus();}				
		
	else if(document.frmPerformance.frmperformanceValidateAML.value == 1 && document.frmPerformance.frmPerformanceAverageMatchLengthOSB.value > 0 && document.frmPerformance.frmPerformanceClosedMatchesOtherSiteBased.value == 0)
		{alert("Average Match Length for Other Site-Based matches cannot be greater than 0 if you did not close any matches during the current month.");document.frmPerformance.frmPerformanceAverageMatchLengthOSB.focus();}				

	else if(document.frmPerformance.frmperformanceValidateAML.value == 1 && document.frmPerformance.frmPerformanceAverageMatchLengthOSB.value == 0 && document.frmPerformance.frmPerformanceClosedMatchesOtherSiteBased.value > 0)
		{alert("Average Match Length for Other Site-Based matches must be greater than 0 because you closed matches during the current month.");document.frmPerformance.frmPerformanceAverageMatchLengthOSB.focus();}						
		
	else if(document.frmPerformance.frmPerformanceRTBM_UnmatchedChildren.value == "")
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
		
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceOpenMatchesCommunityBased.value)))
		{alert(document.frmPerformance.frmPerformanceOpenMatchesCommunityBased.value + " is an invalid number");document.frmPerformance.frmPerformanceOpenMatchesCommunityBased.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceOpenMatchesSchoolBased.value)))
		{alert(document.frmPerformance.frmPerformanceOpenMatchesSchoolBased.value + " is an invalid number");document.frmPerformance.frmPerformanceOpenMatchesSchoolBased.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceOpenMatchesOtherSiteBased.value)))
		{alert(document.frmPerformance.frmPerformanceOpenMatchesOtherSiteBased.value + " is an invalid number");document.frmPerformance.frmPerformanceOpenMatchesOtherSiteBased.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceOpenMatchesGroupMentoring.value)))
		{alert(document.frmPerformance.frmPerformanceOpenMatchesGroupMentoring.value + " is an invalid number");document.frmPerformance.frmPerformanceOpenMatchesGroupMentoring.focus();}
	else if(!(onlyInteger.test(document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsMentoring.value)))
		{alert(document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsMentoring.value + " is an invalid number");document.frmPerformance.frmPerformanceOpenMatchesSpecialProgramsMentoring.focus();}
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
			<td colspan="7" class="formHeader">PERFORMANCE - <%= MonthName(Request("m"), False) & " " & Request("y") %></td>
		</tr>
		
		<tr>
			<td colspan="7" class="formMain"><font color="#ff0000"><div align="center"><strong>Please Note: </strong>After entering your information, you <strong>must</strong> click on the "Save" button at the bottom of the form and wait for the "Thank You" screen or your changes will be lost.</div></font></td>
		</tr>
		
			<tr>
				<td>&nbsp;</td>
				<td align="center" valign="middle" class="formMain">Community Based</td>
				<td align="center" valign="middle" class="formMain">School Based</td>
				<td align="center" valign="middle" class="formMain">Other Site Based</td>
				<td align="center" valign="middle" class="formMain">Group Mentoring</td>
				<td align="center" valign="middle" class="formMain">Special Programs: Mentoring</td>
				<td align="center" valign="middle" class="formMain">Special Programs: Non-Mentoring</td>
			</tr>
			<tr>
				<td align="center" valign="middle" class="formMain">OPEN/ACTIVE&nbsp;matches<br>on&nbsp;the&nbsp;last&nbsp;day&nbsp;of<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesCommunityBased") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesCommunityBased" tabindex="1" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesSchoolBased") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesSchoolBased" tabindex="3" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesOtherSiteBased") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesOtherSiteBased" tabindex="5" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesGroupMentoring") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesGroupMentoring" tabindex="7" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"   class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesSpecialProgramsMentoring") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesSpecialProgramsMentoring" tabindex="9" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OpenMatchesSpecialProgramsNonMentoring") %><% Else %>0<% End If %>" name="frmPerformanceOpenMatchesSpecialProgramsNonMentoring" tabindex="11" onchange="checkForIntegerCommas(this.value);">
				</td>
			</tr>
			<tr>
				<td align="center" valign="middle" class="formMain">Matches&nbsp;CLOSED&nbsp;during <b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesCommunityBased") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesCommunityBased" tabindex="2" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesSchoolBased") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesSchoolBased" tabindex="4" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesOtherSiteBased") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesOtherSiteBased" tabindex="6" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesGroupMentoring") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesGroupMentoring" tabindex="8" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesSpecialProgramsMentoring") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesSpecialProgramsMentoring" tabindex="10" onchange="checkForIntegerCommas(this.value);">
				</td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("ClosedMatchesSpecialProgramsNonMentoring") %><% Else %>0<% End If %>" name="frmPerformanceClosedMatchesSpecialProgramsNonMentoring" tabindex="12" onchange="checkForIntegerCommas(this.value);">
				</td>
			</tr>
			
			
<% if (y >= 2002) then %>

			<tr>
			<td align="center" valign="middle" class="formMain">
				AVERAGE&nbsp;LENGTH&nbsp;(In&nbsp;Months)<br>&nbsp;of&nbsp;Matches&nbsp;Closed&nbsp;during<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("AverageMatchLengthCB") %><% Else %>0<% End If %>" name="frmPerformanceAverageMatchLengthCB" tabindex="8" >					
				</td>
				
				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("AverageMatchLengthSB") %><% Else %>0<% End If %>" name="frmPerformanceAverageMatchLengthSB" tabindex="8" >					
				</td>

				<td align="center" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("AverageMatchLengthOSB") %><% Else %>0<% End If %>" name="frmPerformanceAverageMatchLengthOSB" tabindex="8" >					
				</td>		
				<td align="center" valign="middle" class="formMain">n/a</td>						
				<td align="center" valign="middle" class="formMain">n/a</td>
				<td align="center" valign="middle" class="formMain">n/a</td>
			</tr>
			
<% End If %>



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

	
			<tr>				
				<td colspan="7" class="formHeaderMedium">SCHOOL-BASED MENTORING GRANT PROGRESS REPORT</td>
			</tr>
			<tr>
				<td colspan="1" class="formMain">Number of Volunteers Currently in the Enrollment Process</td>
				<td colspan="1" align="left" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBMVolunteersInEnrollmentProcess") %><% Else %>0<% End If %>" name="frmPerformanceSBMVolunteersInEnrollmentProcess" tabindex="14" onchange="checkForIntegerCommas(this.value);">
				</td>	
				<td colspan="5">&nbsp;</td>			
			</tr>					
			<tr>
				<td colspan="1" class="formMain">Amount Raised Towards Match Pledge as of the last day of &nbsp<b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				<td colspan="1" align="left" valign="middle" class="formMain">
					<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBMAmountRaisedTowardsMatchPledge") %><% Else %>0<% End If %>" name="frmPerformanceSBMAmountRaisedTowardsMatchPledge" tabindex="14" onchange="checkForIntegerCommas(this.value);">					
				</td>
				<td colspan="5">&nbsp;</td>
			</tr>			
			
			
<% End If %>


<% if (y >= 2002) then %>

		<tr>
			<td colspan="7" class="formHeaderMedium">REVENUE</td>				
		</tr>
		<tr>
			
			
			<td align="center" valign="bottom" class="formMain" colspan="1">Revenue&nbsp;<strong>booked&nbsp;for&nbsp;the&nbsp;Month&nbsp;of&nbsp;<%= MonthName(Request("m"), False) & " " & Request("y") %>&nbsp;&nbsp;</strong><a href="../helpfiles/surveyhelp.asp?HelpID=rev1" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
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
				<td class="formMain" colspan="3" align="center"><strong>Volunteer</strong></td>
				<td class="formMain" colspan="3" align="center"><strong>Parent / Youth</strong></td>

			</tr>
			
			<tr>
				<td class="formMain">Number of Inquiries</td>
				<td class="formMain" colspan="3" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Inquiries") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Inquiries" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm1" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<td class="formMain" colspan="3" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Inquiries") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Inquiries" onchange="checkForIntegerCommas(this.value);">&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm2" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
			</tr>

			<tr>
				<td class="formMain">Number of In-Person Interviews<a href="../helpfiles/surveyhelp.asp?HelpID=sdm3" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<td class="formMain" colspan="3" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Interviews") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Interviews" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="3" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Interviews") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Interviews" onchange="checkForIntegerCommas(this.value);"></td>
			</tr>
			
			<tr>
				<td class="formMain">Number Matched&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm4" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<td class="formMain" colspan="3" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Vol_Matched") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Vol_Matched" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="3" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("YieldRate_Youth_Matched") %><% Else %>0<% End If %>" name="frmPerformanceYieldRate_Youth_Matched" onchange="checkForIntegerCommas(this.value);"></td>
			</tr>		
			
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
				<td class="formMain" colspan="2" align="center"><strong>Other Site-Based</strong></td>		
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
				<TD colspan="3" class="formMain" align="center"><b>Community-Based</b></TD>
				<TD colspan="3" class="formMain" align="center"><b>School-Based</b></TD>
	
			</TR>
			
			<tr>
				<td colspan="1" class="formMain">Number of Matches that Closed Prematurely</td>			
				<td class="formMain" colspan="3" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBNumberClosedPrematurely") %><% Else %>0<% End If %>" name="frmPerformanceCBNumberClosedPrematurely" onchange="checkForIntegerCommas(this.value);"></td>				
				<td class="formMain" colspan="3" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBNumberClosedPrematurely") %><% Else %>0<% End If %>" name="frmPerformanceSBNumberClosedPrematurely" onchange="checkForIntegerCommas(this.value);"></td>								
			</tr>
			
			<tr>
				<TD colspan="7" class="formHeaderMedium">CLOSE CODES</TD>	
			</tr>
			
			<TR>
				<TD colspan="1">&nbsp;</TD>
				<TD colspan="3" class="formMain" align="center"><b>Community-Based</b></TD>
				<TD colspan="3" class="formMain" align="center"><b>School-Based</b></TD>
	
			</TR>
			

			
			<tr>
				<td colspan="1" class="formMain">Child/Parent Status Change&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm11" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<td class="formMain" colspan="3" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBChildParentStatusChange") %><% Else %>0<% End If %>" name="frmPerformanceCBChildParentStatusChange" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="3" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBChildParentStatusChange") %><% Else %>0<% End If %>" name="frmPerformanceSBChildParentStatusChange" onchange="checkForIntegerCommas(this.value);"></td>				

			</tr>
			
			<tr>
				<td colspan="1" class="formMain">Volunteer Status Change&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm12" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>	
				<td class="formMain" colspan="3" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBVolunteerStatusChange") %><% Else %>0<% End If %>" name="frmPerformanceCBVolunteerStatusChange" onchange="checkForIntegerCommas(this.value);"></td>
				<td class="formMain" colspan="3" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBVolunteerStatusChange") %><% Else %>0<% End If %>" name="frmPerformanceSBVolunteerStatusChange" onchange="checkForIntegerCommas(this.value);"></td>				
			</tr>
			
			<tr>
				<td colspan="1" class="formMain">Child/Parent Dissatisfaction&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm13" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>	
				<td class="formMain" colspan="3" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBChildParentDissatisfaction") %><% Else %>0<% End If %>" name="frmPerformanceCBChildParentDissatisfaction" onchange="checkForIntegerCommas(this.value);"></td>				
				<td class="formMain" colspan="3" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBChildParentDissatisfaction") %><% Else %>0<% End If %>" name="frmPerformanceSBChildParentDissatisfaction" onchange="checkForIntegerCommas(this.value);"></td>								
			</tr>
			
			<tr>

				<td colspan="1" class="formMain">Volunteer Dissatisfaction&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm14" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<td class="formMain" colspan="3" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBVolunteerDissatisfaction") %><% Else %>0<% End If %>" name="frmPerformanceCBVolunteerDissatisfaction" onchange="checkForIntegerCommas(this.value);"></td>								
				<td class="formMain" colspan="3" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBVolunteerDissatisfaction") %><% Else %>0<% End If %>" name="frmPerformanceSBVolunteerDissatisfaction" onchange="checkForIntegerCommas(this.value);"></td>												
			</tr>
			
			<tr>
				<TD colspan="7" class="formHeaderMedium">6-Month Retention (Community Based Only)</TD>	
			</tr>		
			
			
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
				<td colspan="1" class="formMain">Total Opened 6 Months Ago&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm15&SixMonthsAgo=<%=SixMonthsAgo%>" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>				
				<td class="formMain" colspan="3" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBTotalOpened6MonthsAgo") %><% Else %>0<% End If %>" name="frmPerformanceCBTotalOpened6MonthsAgo" onchange="checkForIntegerCommas(this.value);"></td>								
				<td colspan="3">&nbsp;</td>
			</tr>	
			
			<tr>
				<td colspan="1" class="formMain">Number Still Open Now&nbsp;<a href="../helpfiles/surveyhelp.asp?HelpID=sdm16&SixMonthsAgo=<%=SixMonthsAgo%>&Now=<%=m%>" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a></td>
				<td class="formMain" colspan="3" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBNumberStillOpen") %><% Else %>0<% End If %>" name="frmPerformanceCBNumberStillOpen" onchange="checkForIntegerCommas(this.value);"></td>												
				<td colspan="3">&nbsp;</td>			
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
				<td colspan="3" class="formMain" align="center"><b>Community-Based</b></td>
				<td colspan="3" class="formMain" align="center"><b>School-Based</b></td>
			</tr>
			
			<tr>
				<td colspan="1" class="formMain">POE Aggregate Score</td>
				<td colspan="3" class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBPOEAggregateScore") %><% Else %>0<% End If %>" name="frmPerformanceCBPOEAggregateScore" onchange="checkForIntegerCommas(this.value);"></td>
				<td colspan="3" class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBPOEAggregateScore") %><% Else %>0<% End If %>" name="frmPerformanceSBPOEAggregateScore" onchange="checkForIntegerCommas(this.value);"></td>				
			</tr>
			
			<tr>
				<td colspan="1" class="formMain">POE Count</td>	
				<td colspan="3" class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("CBPOECount") %><% Else %>0<% End If %>" name="frmPerformanceCBPOECount" onchange="checkForIntegerCommas(this.value);"></td>
				<td colspan="3" class="formMain" align="center"><input type="text"  colspan="4"  align="center" class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("SBPOECount") %><% Else %>0<% End If %>" name="frmPerformanceSBPOECount" onchange="checkForIntegerCommas(this.value);"></td>								
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
			
<!-- PARTNERSHIP QUESTIONNAIRE -->

<% if (y = 2003 and (m=4 or m=11)) or (y > 2003 and (m=5 or m=11)) then %>

	<tr>
		<td align="center" colspan="7" class="formmain">&nbsp;</td>
	</tr>
	
	<tr>
		<td colspan="7" class="formHeader">PARTNERSHIP QUESTIONNAIRE&nbsp;&nbsp;<a href="..\helpfiles\surveyhelp.asp?HelpID=pq1" onclick="NewWindow(this.href,'name','500','250','yes');return false;"><img src="../images/qmark_purple.gif" alt="" width="21" height="22" border="0"></a></td>
	</tr>
	
	
	
	
	<!-- Active Matches -->
	<tr>
		<td align="center" colspan="7" class="formmain">If Applicable, Enter the number of <strong>Active Matches</strong> with the following organizations:</td>
	</tr>
	
	<tr>
		<td>&nbsp;</td>
		<td align="center" class="formMain">Community<br>Based</td>
		<td align="center" class="formMain">School<br>Based</td>	
		<td align="center" class="formMain">Other<br>Site Based</td>	
		<td align="center" class="formMain">Not<br>Partnering <em>(check box if not partnering)</em></td>	
		<td align="center" colspan="2" class="formMain"><em>If Not Partnering,</em> interested <br>in forming a partnership?</td>

	</tr>
	
	<!-- Alpha Phi Alpha -->
	<tr>
		<td align="left" class="formmain">Alpha Phi Alpha</td>
		
		<!-- Alpha Community Based -->
		<td align="center" class="formmain">
			<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("AlphaCommunityBased") %><% Else %>0<% End If %>" name="frmPerformanceAlphaCommunityBased" onchange="checkForIntegerCommas(this.value);">
		</td>
		
		<!-- Alpha School Based -->
		<td align="center" class="formmain">
			<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("AlphaSchoolBased") %><% Else %>0<% End If %>" name="frmperformanceAlphaSchoolBased" onchange="checkForIntegerCommas(this.value);">
		</td>	
		
		<!-- Alpha Other Site Based -->
		<td align="center" class="formmain">
			<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("AlphaOtherSiteBased") %><% Else %>0<% End If %>" name="frmperformanceAlphaOtherSiteBased" onchange="checkForIntegerCommas(this.value);">
		</td>		
		
		<!-- Alpha Not Partnering -->
		<td align="center" class="formmain">
			<input type="Checkbox" name="frmperformanceAlphaNotPartnering" value="1"<% If say = "edit" Then %><% if Trim(GetPerformance("AlphaNotPartnering"))="1" then %>checked<% end if %><% End If %> >	
		</td>	
		
		<!-- Alpha Interest -->
		<td align="center" colspan="2" class="formmain">
			<input type="radio" name="frmperformanceAlphainterest" value="1" <% If say = "edit" Then%><% if Trim(GetPerformance("AlphaInterest")) = "1" then %> checked <% End If %><% End If %> > Yes
			<input type="radio" name="frmperformanceAlphainterest" value="0" <% If say = "edit" Then%><% if Trim(GetPerformance("AlphaInterest")) = "0" then %> checked <% End If %><% End If %> > No
		</td>

		
	</tr>
	
	<!-- Lions Club -->
	<tr>
		<td align="left" class="formmain">Lions Club</td>	
		
		<!-- Lions Community Based -->
		<td align="center" class="formmain">
			<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("LionsCommunityBased") %><% Else %>0<% End If %>" name="frmPerformanceLionsCommunityBased" onchange="checkForIntegerCommas(this.value);">	
		</td>
		
		<!-- Lions School Based -->
		<td align="center" class="formmain">
			<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("LionsSchoolBased") %><% Else %>0<% End If %>" name="frmperformanceLionsSchoolBased" onchange="checkForIntegerCommas(this.value);">	
		</td>
		
		<!-- Lions Other Site Based -->
		<td align="center" class="formmain">
			<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("LionsOtherSiteBased") %><% Else %>0<% End If %>" name="frmperformanceLionsOtherSiteBased" onchange="checkForIntegerCommas(this.value);">		
		</td>	
		
		<!-- Lions Not Partnering -->
		<td align="center" class="formmain">
			<input type="Checkbox" name="frmperformanceLionsNotPartnering" value="1" <% If say = "edit" Then %><% if Trim(GetPerformance("LionsNotPartnering"))="1" then %>checked<% end if %><% End If %> >		
		</td>
	
		<!-- Lions Interest -->
		<td align="center" colspan="2" class="formmain">
			<input type="radio" name="frmperformanceLionsinterest" value="1" <% If say = "edit" Then%><% if Trim(GetPerformance("LionsInterest")) = "1" then %> checked <% End If %><% End If %> > Yes
			<input type="radio" name="frmperformanceLionsinterest" value="0" <% If say = "edit" Then%><% if Trim(GetPerformance("LionsInterest")) = "0" then %> checked <% End If %><% End If %> > No	
		</td>	

	</tr>
	
	<!-- Rotary Club -->
	<tr>
		<td align="left" class="formmain">Rotary Club</td>	
		
		<!-- Rotary Community Based -->
		<td align="center" class="formmain">
			<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("rotaryCommunityBased") %><% Else %>0<% End If %>" name="frmperformanceRotaryCommunityBased" onchange="checkForIntegerCommas(this.value);">		
		</td>
		
		<!-- Rotary School Based -->
		<td align="center" class="formmain">
			<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("RotarySchoolBased") %><% Else %>0<% End If %>" name="frmperformanceRotarySchoolBased" onchange="checkForIntegerCommas(this.value);">		
		</td>
		
		<!-- Rotary Other Site Based -->
		<td align="center" class="formmain">
			<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("RotaryOtherSiteBased") %><% Else %>0<% End If %>" name="frmperformanceRotaryOtherSiteBased"  onchange="checkForIntegerCommas(this.value);">			
		</td>	
		
		<!-- Rotary Not Partnering -->
		<td align="center" class="formmain">
			<input type="Checkbox" name="frmperformanceRotaryNotPartnering" value="1" <% If say = "edit" Then %><% if Trim(GetPerformance("RotaryNotPartnering"))="1" then %>checked<% end if %><% End If %> >			
		</td>	
		
		<!-- Rotary Interest -->
		<td align="center" colspan="2" class="formmain">
			<input type="radio" name="frmperformanceRotaryinterest" value="1" <% If say = "edit" Then%><% if Trim(GetPerformance("RotaryInterest")) = "1" then %> checked <% End If %><% End If %> > Yes	
			<input type="radio" name="frmperformanceRotaryinterest" value="0" <% If say = "edit" Then%><% if Trim(GetPerformance("RotaryInterest")) = "0" then %> checked <% End If %><% End If %> > No	
		</td>		

		
	</tr>
	
	<!-- Kiwanis Club -->
	<tr>
		<td align="left" class="formmain">Kiwanis Club</td>	
		
		<!-- Kiwanis Community Based -->
		<td align="center" class="formmain">
			<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("KiwanisCommunityBased") %><% Else %>0<% End If %>" name="frmperformanceKiwanisCommunityBased"  onchange="checkForIntegerCommas(this.value);">			
		</td>
		
		<!-- Kiwanis School Based -->
		<td align="center" class="formmain">
			<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("KiwanisSchoolBased") %><% Else %>0<% End If %>" name="frmperformanceKiwanisSchoolBased"  onchange="checkForIntegerCommas(this.value);">			
		</td>
		
		<!-- Kiwanis Other Site Based -->
		<td align="center" class="formmain">
			<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("KiwanisOtherSiteBased") %><% Else %>0<% End If %>" name="frmperformanceKiwanisOtherSiteBased"  onchange="checkForIntegerCommas(this.value);">				
		</td>	
		
		<!-- Kiwanis Not Partnering -->
		<td align="center" class="formmain">
			<input type="Checkbox" name="frmperformanceKiwanisNotPartnering" value="1" <% If say = "edit" Then %><% if Trim(GetPerformance("KiwanisNotPartnering"))="1" then %>checked<% end if %><% End If %> >				
		</td>		
		
		<!-- Kiwanis Interest -->
		<td align="center" colspan="2" class="formmain">
			<input type="radio" name="frmperformanceKiwanisinterest" value="1" <% If say = "edit" Then%><% if Trim(GetPerformance("KiwanisInterest")) = "1" then %> checked <% End If %><% End If %> > Yes	
			<input type="radio" name="frmperformanceKiwanisinterest" value="0" <% If say = "edit" Then%><% if Trim(GetPerformance("KiwanisInterest")) = "0" then %> checked <% End If %><% End If %> > No		
		</td>		

	</tr>
	
	
	<!-- Optimist Club -->
	<tr>
		<td align="left" class="formmain">Optimist Club</td>	
		
		<!-- Optimist Community Based -->
		<td align="center" class="formmain">
			<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OptimistCommunityBased") %><% Else %>0<% End If %>" name="frmperformanceOptimistCommunityBased"  onchange="checkForIntegerCommas(this.value);">				
		</td>
		
		<!-- Optimist School Based -->
		<td align="center" class="formmain">
			<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OptimistSchoolBased") %><% Else %>0<% End If %>" name="frmperformanceOptimistSchoolBased"  onchange="checkForIntegerCommas(this.value);">			
		</td>
		
		<!-- Optimist Other Site Based -->
		<td align="center" class="formmain">
			<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("OptimistOtherSiteBased") %><% Else %>0<% End If %>" name="frmperformanceOptimistOtherSiteBased"  onchange="checkForIntegerCommas(this.value);">			
		</td>	
		
		<!-- Optimist Not Partnering -->
		<td align="center" class="formmain">
			<input type="Checkbox" name="frmperformanceOptimistNotPartnering" value="1" <% If say = "edit" Then %><% if Trim(GetPerformance("OptimistNotPartnering"))="1" then %>checked<% end if %><% End If %> >					
		</td>		
		
		<!-- Optimist Interest -->
		<td align="center" colspan="2" class="formmain">
			<input type="radio" name="frmperformanceOptimistinterest" value="1" <% If say = "edit" Then%><% if Trim(GetPerformance("OptimistInterest")) = "1" then %> checked <% End If %><% End If %> > Yes	
			<input type="radio" name="frmperformanceOptimistinterest" value="0" <% If say = "edit" Then%><% if Trim(GetPerformance("OptimistInterest")) = "0" then %> checked <% End If %><% End If %> > No			
		</td>		
	
		
	</tr>
	
	
	<!-- AARP Club -->
	<tr>
		<td align="left" class="formmain">AARP</td>	
		
		<!-- AARP Community Based -->
		<td align="center" class="formmain">
			<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("AARPCommunityBased") %><% Else %>0<% End If %>" name="frmperformanceAARPCommunityBased"  onchange="checkForIntegerCommas(this.value);">					
		</td>
		
		<!-- AARP School Based -->
		<td align="center" class="formmain">
			<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("AARPSchoolBased") %><% Else %>0<% End If %>" name="frmperformanceAARPSchoolBased"  onchange="checkForIntegerCommas(this.value);">					
		</td>
		
		<!-- AARP Other Site Based -->
		<td align="center" class="formmain">
			<input type="text"  class="formMain" size="5" maxlength="10" value="<% If say = "edit" Then %><%= GetPerformance("AARPOtherSiteBased") %><% Else %>0<% End If %>" name="frmperformanceAARPOtherSiteBased"  onchange="checkForIntegerCommas(this.value);">					
		</td>	
		
		<!-- AARP Not Partnering -->
		<td align="center" class="formmain">
			<input type="Checkbox" name="frmperformanceAARPNotPartnering" value="1" <% If say = "edit" Then %><% if Trim(GetPerformance("AARPNotPartnering"))="1" then %>checked<% end if %><% End If %> >					
		</td>		
		
		<!-- AARP Interest -->
		<td align="center" colspan="2" class="formmain">
			<input type="radio" name="frmperformanceAARPinterest" value="1" <% If say = "edit" Then%><% if Trim(GetPerformance("AARPInterest")) = "1" then %> checked <% End If %><% End If %> > Yes	
			<input type="radio" name="frmperformanceAARPinterest" value="0" <% If say = "edit" Then%><% if Trim(GetPerformance("AARPInterest")) = "0" then %> checked <% End If %><% End If %> > No			
	
		</td>		

		
	</tr>
	
	<!-- Partnership Rating -->
	<tr>
		<td colspan="7" class="formHeaderMedium">PARTNERSHIP RATING</td>	
	</tr>
	
	<tr>
		<td colspan="7" align="center" class="formMain">Rate the Nature of the Partnership from 1 to 5 - based on level of interaction, with 5 being the highest -  or select 'Not Applicable'</td>	
	</tr>
	
	<!-- Alpha Rating -->
	<tr>
		<td class="formMain" colspan="2">Alpha Phi Alpha</td>
	
		<td class="formMain" colspan="7" align="left">	
		<select name="frmperformanceAlphaRating">
		<option value=0 <% If say = "edit" Then %><% if Trim(GetPerformance("AlphaRating")) = "0" then %> selected><% end if %><%end if %>Not Applicable </option>Not Applicable
		<option value=1 <% If say = "edit" Then %><% If Trim(GetPerformance("AlphaRating")) = "1" then %> selected><% end if %><%end if %>1 - Informal</option>1 - Informal
		<option value=2 <% If say = "edit" Then %><% If Trim(GetPerformance("AlphaRating")) = "2" then %> selected><% end if %><%end if %>2</option>2
		<option value=3 <% If say = "edit" Then %><% If Trim(GetPerformance("AlphaRating")) = "3" then %> selected><% end if %><%end if %>3</option>3
		<option value=4 <% If say = "edit" Then %><% If Trim(GetPerformance("AlphaRating")) = "4" then %> selected><% end if %><%end if %>4</option>4
		<option value=5 <% If say = "edit" Then %><% If Trim(GetPerformance("AlphaRating")) = "5" then %> selected><% end if %><%end if %>5 - Formal</option>5 - Formal
		</select>
		</td>
	</tr>
	
	<!-- Lions Club Rating -->
	<tr>
		<td class="formMain" colspan="2">Lions Club</td>
		<td class="formMain" colspan="7" align="left">	
		<select name="frmperformanceLionsRating">
		<option value=0 <% If say = "edit" Then %><% if Trim(GetPerformance("LionsRating")) = "0" then %> selected><% end if %><%end if %>Not Applicable </option>Not Applicable
		<option value=1 <% If say = "edit" Then %><% If Trim(GetPerformance("LionsRating")) = "1" then %> selected><% end if %><%end if %>1 - Informal</option>1 - Informal
		<option value=2 <% If say = "edit" Then %><% If Trim(GetPerformance("LionsRating")) = "2" then %> selected><% end if %><%end if %>2</option>2
		<option value=3 <% If say = "edit" Then %><% If Trim(GetPerformance("LionsRating")) = "3" then %> selected><% end if %><%end if %>3</option>3
		<option value=4 <% If say = "edit" Then %><% If Trim(GetPerformance("LionsRating")) = "4" then %> selected><% end if %><%end if %>4</option>4
		<option value=5 <% If say = "edit" Then %><% If Trim(GetPerformance("LionsRating")) = "5" then %> selected><% end if %><%end if %>5 - Formal</option>5 - Formal
		</select>
		</td>	
	</tr>
	
	<!-- Rotary Club Rating -->
	<tr>
		<td class="formMain" colspan="2">Rotary Club</td>
		<td class="formMain" colspan="7" align="left">	
		<select name="frmperformanceRotaryRating">
		<option value=0 <% If say = "edit" Then %><% if Trim(GetPerformance("RotaryRating")) = "0" then %> selected><% end if %><%end if %>Not Applicable </option>Not Applicable
		<option value=1 <% If say = "edit" Then %><% If Trim(GetPerformance("RotaryRating")) = "1" then %> selected><% end if %><%end if %>1 - Informal</option>1 - Informal
		<option value=2 <% If say = "edit" Then %><% If Trim(GetPerformance("RotaryRating")) = "2" then %> selected><% end if %><%end if %>2</option>2
		<option value=3 <% If say = "edit" Then %><% If Trim(GetPerformance("RotaryRating")) = "3" then %> selected><% end if %><%end if %>3</option>3
		<option value=4 <% If say = "edit" Then %><% If Trim(GetPerformance("RotaryRating")) = "4" then %> selected><% end if %><%end if %>4</option>4
		<option value=5 <% If say = "edit" Then %><% If Trim(GetPerformance("RotaryRating")) = "5" then %> selected><% end if %><%end if %>5 - Formal</option>5 - Formal
		</select>
		</td>	
	</tr>
	
	<!-- Kiwanis Club Rating -->
	<tr>
		<td class="formMain" colspan="2">Kiwanis Club</td>
		<td class="formMain" colspan="7" align="left">	
		<select name="frmperformanceKiwanisRating">
		<option value=0 <% If say = "edit" Then %><% if Trim(GetPerformance("KiwanisRating")) = "0" then %> selected><% end if %><%end if %>Not Applicable </option>Not Applicable
		<option value=1 <% If say = "edit" Then %><% If Trim(GetPerformance("KiwanisRating")) = "1" then %> selected><% end if %><%end if %>1 - Informal</option>1 - Informal
		<option value=2 <% If say = "edit" Then %><% If Trim(GetPerformance("KiwanisRating")) = "2" then %> selected><% end if %><%end if %>2</option>2
		<option value=3 <% If say = "edit" Then %><% If Trim(GetPerformance("KiwanisRating")) = "3" then %> selected><% end if %><%end if %>3</option>3
		<option value=4 <% If say = "edit" Then %><% If Trim(GetPerformance("KiwanisRating")) = "4" then %> selected><% end if %><%end if %>4</option>4
		<option value=5 <% If say = "edit" Then %><% If Trim(GetPerformance("KiwanisRating")) = "5" then %> selected><% end if %><%end if %>5 - Formal</option>5 - Formal
		</select>
		</td>	
	</tr>
	
	<!-- Optimist Club Rating -->
	<tr>
		<td class="formMain" colspan="2">Optimist Club</td>
		<td class="formMain" colspan="7" align="left">	
		<select name="frmperformanceOptimistRating">
		<option value=0 <% If say = "edit" Then %><% if Trim(GetPerformance("OptimistRating")) = "0" then %> selected><% end if %><%end if %>Not Applicable </option>Not Applicable
		<option value=1 <% If say = "edit" Then %><% If Trim(GetPerformance("OptimistRating")) = "1" then %> selected><% end if %><%end if %>1 - Informal</option>1 - Informal
		<option value=2 <% If say = "edit" Then %><% If Trim(GetPerformance("OptimistRating")) = "2" then %> selected><% end if %><%end if %>2</option>2
		<option value=3 <% If say = "edit" Then %><% If Trim(GetPerformance("OptimistRating")) = "3" then %> selected><% end if %><%end if %>3</option>3
		<option value=4 <% If say = "edit" Then %><% If Trim(GetPerformance("OptimistRating")) = "4" then %> selected><% end if %><%end if %>4</option>4
		<option value=5 <% If say = "edit" Then %><% If Trim(GetPerformance("OptimistRating")) = "5" then %> selected><% end if %><%end if %>5 - Formal</option>5 - Formal
		</select>
		</td>	
	</tr>
	
	<!-- AARP Rating -->
	<tr>
		<td class="formMain" colspan="2">AARP</td>
		<td class="formMain" colspan="7" align="left">	
		<select name="frmperformanceAARPRating">
		<option value=0 <% If say = "edit" Then %><% if Trim(GetPerformance("AARPRating")) = "0" then %> selected><% end if %><%end if %>Not Applicable </option>Not Applicable
		<option value=1 <% If say = "edit" Then %><% If Trim(GetPerformance("AARPRating")) = "1" then %> selected><% end if %><%end if %>1 - Informal</option>1 - Informal
		<option value=2 <% If say = "edit" Then %><% If Trim(GetPerformance("AARPRating")) = "2" then %> selected><% end if %><%end if %>2</option>2
		<option value=3 <% If say = "edit" Then %><% If Trim(GetPerformance("AARPRating")) = "3" then %> selected><% end if %><%end if %>3</option>3
		<option value=4 <% If say = "edit" Then %><% If Trim(GetPerformance("AARPRating")) = "4" then %> selected><% end if %><%end if %>4</option>4
		<option value=5 <% If say = "edit" Then %><% If Trim(GetPerformance("AARPRating")) = "5" then %> selected><% end if %><%end if %>5 - Formal</option>5 - Formal
		</select>
		</td>	
	
	</tr>
	
	<!-- Alpha Phi Alpha Partnership -->
	<tr>
		<td colspan="7" class="formHeaderMedium">ALPHA PHI ALPHA PARTNERSHIP</td>	
	</tr>
	
	<tr>
		<td colspan="7" align="center" class="formMain">I am partnering with the Alphas in the following ways (check all that apply):</td>
	</tr>
	
	<tr>
		<td colspan="7" align="left" class="formMain">
		<input type="Checkbox" name="frmperformanceAlphaFunding" value="1" <% If say = "edit" Then %><% if Trim(GetPerformance("AlphaFunding"))="1" then %>checked<% end if %><% End If %> >Funding: Alpha chapter supports BBBS funding efforts<br>
		<input type="Checkbox" name="frmperformanceAlphaProgramInitiative" value="1" <% If say = "edit" Then %><% if Trim(GetPerformance("AlphaProgramInitiative"))="1" then %>checked<% end if %><% End If %> >Program Initiative: Chapter has activities with children on waiting list<br>
		<input type="Checkbox" name="frmperformanceAlphaLeadershipInvolvement" value="1" <% If say = "edit" Then %><% if Trim(GetPerformance("AlphaLeadershipInvolvement"))="1" then %>checked<% end if %><% End If %> >Leadership Involvement: Alpha serves on board, provides agency with professional skills and resources (serves as volunteer)
		</td>

	</tr>
	
	
	<!-- Chapter Locations -->
	
	<%
	set StateChoices = Server.CreateObject("ADODB.Recordset")
	StateChoices.ActiveConnection = ConnStr
	StateChoices.Source = "SELECT DISTINCT StateSpelledOut,StateAbbreviation FROM tblAGLUST order by StateSpelledOut"
	StateChoices.CursorType = 0
	StateChoices.CursorLocation = 2
	StateChoices.Open()
	%>
	
	
	<% 
	If say = "edit" Then 
		ChosenUndergradState = GetPerformance("AlphaUndergradChapterState")
		ChosenUndergradStateSpelledOut = ""
		set UndergradStateChosen = Server.CreateObject("ADODB.Recordset")
		UndergradStateChosen.ActiveConnection = ConnStr
		UndergradStateChosen.Source = "SELECT StateSpelledOut,StateAbbreviation FROM tblAGLUST WHERE StateAbbreviation = '" & ChosenUndergradState & "'"
		UndergradStateChosen.CursorType = 0
		UndergradStateChosen.CursorLocation = 2
		UndergradStateChosen.Open()	
		If not UndergradStateChosen.EOF then
			ChosenUndergradStateSpelledOut = (UndergradStateChosen.Fields.Item("StateSpelledOut").Value)
		End if
	Else
		ChosenUndergradState = ""
		ChosenUndergradStateSpelledOut = "" 
	End If
	%>
	
	<% 
	If say = "edit" Then 
		ChosenAlumniState = GetPerformance("AlphaAlumniChapterState")
		ChosenAlumniStateSpelledOut = ""
		set AlumniStateChosen = Server.CreateObject("ADODB.Recordset")
		AlumniStateChosen.ActiveConnection = ConnStr
		AlumniStateChosen.Source = "SELECT StateSpelledOut,StateAbbreviation FROM tblAGLUST WHERE StateAbbreviation = '" & ChosenAlumniState & "'"
		AlumniStateChosen.CursorType = 0
		AlumniStateChosen.CursorLocation = 2
		AlumniStateChosen.Open()	
		If not AlumniStateChosen.EOF then
			ChosenAlumniStateSpelledOut = (AlumniStateChosen.Fields.Item("StateSpelledOut").Value)
		End if
		
	Else
		ChosenAlumniState = ""
		ChosenAlumniStateSpelledOut = "" 
	End If
	%>
	
	
	<tr>
		<td colspan="7" align="center" class="formMain">Please enter the name and location of your local Alpha Phi Alpha Chapter(s):</td>
	</tr>
	
	<tr>
		<td align="left" class="formMain">Undergraduate Chapter</td>
		<td align="left" colspan="7" class="formMain">
		
		Name:&nbsp;<input type="text" name="frmperformanceAlphaUndergradChapterName" size="50" maxlength="50" class="formMain" value="<% If say = "edit" Then %><%= GetPerformance("AlphaUndergradChapterName") %><% Else %> <% End If %>"><br>
		City:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="frmperformanceAlphaUndergradChapterCity" size="40" maxlength="40" class="formMain" value="<% If say = "edit" Then %><%= GetPerformance("AlphaUndergradChapterCity") %><% Else %> <% End If %>"><br>	
		State:&nbsp;&nbsp;
		<select NAME="frmperformanceAlphaUndergradChapterState" class="formMain">
		  <option value="<%= ChosenUndergradState %>"><%= ChosenUndergradStateSpelledOut %></option>
		  <%
		  While (NOT StateChoices.EOF)
		  %>
		  <option value="<%=(StateChoices.Fields.Item("StateAbbreviation").Value)%>"><%=(StateChoices.Fields.Item("StateSpelledOut").Value)%></option>
		  <%
		   StateChoices.MoveNext()
		  Wend
		  %>
		</select>
		</td>	
	</tr>
	
	<% StateChoices.MoveFirst() %>
	<tr>
		<td align="left" class="formMain">Alumni Chapter</td>
		<td align="left" colspan="7" class="formMain">
		Name:&nbsp;<input type="text" name="frmperformanceAlphaAlumniChapterName" size="50" maxlength="50" class="formMain" value="<% If say = "edit" Then %><%= GetPerformance("AlphaAlumniChapterName") %><% Else %> <% End If %>"><br>	
		City:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="frmperformanceAlphaAlumniChapterCity" size="40" maxlength="40" class="formMain" value="<% If say = "edit" Then %><%= GetPerformance("AlphaAlumniChapterCity") %><% Else %> <% End If %>"><br>	
		State:&nbsp;&nbsp;
		<select NAME="frmperformanceAlphaAlumniChapterState" Class="formMain">
		  <option value="<%= ChosenAlumniState %>"><%= ChosenAlumniStateSpelledOut %></option>
		  <%
		  While (NOT StateChoices.EOF)
		  %>
		  <option value="<%=(StateChoices.Fields.Item("StateAbbreviation").Value)%>"><%=(StateChoices.Fields.Item("StateSpelledOut").Value)%></option>
		  <%
		   StateChoices.MoveNext()
		  Wend
		  %>
		</select>
		</td>	
	</tr>
	
	<% If say = "edit" Then
	
		StateChoices.Close
		Set Statechoices = Nothing
		
		UndergradStateChosen.Close
		Set UndergradStateChosen = Nothing
		
		AlumniStateChosen.Close
		Set AlumniStateChosen = Nothing
		

		
	End If %>

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
