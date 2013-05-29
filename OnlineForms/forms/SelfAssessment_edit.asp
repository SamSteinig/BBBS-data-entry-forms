<% 
Section = Request("section")
If Request("status") = "addNew" Then


	' Check for duplicate records
	
	Set DupCon = Server.CreateObject("ADODB.Connection")
	DupCon.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT Count(*) As NumberOfEntries FROM tbl_frmSelfAssessment WHERE AgencyID = '" & Request("AgencyIDN") & "' and Year = " & Request("Year")	
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
		RST.Open "SELECT * FROM tbl_frmSelfAssessment", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = Request("AgencyIDN")
		RST("Year") = Request("year")
			
		RST("Std1a") = request("frmSelfAssessmentStd1a")
		RST("Std1b") = request("frmSelfAssessmentStd1b")
		RST("Form990") = request("frmSelfAssessmentForm990")
		RST("Std1c") = request("frmSelfAssessmentStd1c")
		RST("Bylaws") = request("frmSelfAssessmentBylaws")
		RST("MAA") = request("frmSelfAssessmentMAA")
		RST("Std2") = request("frmSelfAssessmentStd2")
		'RST("Std2aSO3a") = request("frmSelfAssessmentStd2aSO3a")
		'RST("BrdTrainPlan") = request("frmSelfAssessmentBrdTrainPlan")
		RST("Std2bSO3b") = request("frmSelfAssessmentStd2bSO3b")
		RST("MAA810Conf") = request("frmSelfAssessmentMAA810Conf")
		RST("Std2SO") = request("frmSelfAssessmentStd2SO")
		RST("Std3SO4m") = request("frmSelfAssessmentStd3SO4m")
		RST("Std3SO4v") = request("frmSelfAssessmentStd3SO4v")
		RST("Std4SO5") = request("frmSelfAssessmentStd4SO5")
		RST("OpPlan") = request("frmSelfAssessmentOpPlan")
		RST("Std5OpsSO6") = request("frmSelfAssessmentStd5OpsSO6")
		RST("Std5PgmSO6") = request("frmSelfAssessmentStd5PgmSO6")
		RST("Std5FilesSO6") = request("frmSelfAssessmentStd5FilesSO6")
		RST("Std6SO7b") = request("frmSelfAssessmentStd6SO7b")
		RST("Std6SO7") = request("frmSelfAssessmentStd6SO7")
		RST("Std6SO7Budget") = request("frmSelfAssessmentStd6SO7Budget")
		RST("MAA810Exp") = request("frmSelfAssessmentMAA810Exp")
		RST("MAA32") = request("frmSelfAssessmentMAA32")
		RST("Std7SO8") = request("frmSelfAssessmentStd7SO8")
		RST("MAA88") = request("frmSelfAssessmentMAA88")
		RST("MAA82") = request("frmSelfAssessmentMAA82")
		RST("StdSO8a") = request("frmSelfAssessmentStdSO8a")
		'RST("StdSO8b") = request("frmSelfAssessmentStdSO8b")
		RST("StdSO8c") = request("frmSelfAssessmentStdSO8c")
		'RST("Std8SO9Crisis") = request("frmSelfAssessmentStd8SO9Crisis")
		RST("Std8SO9Risk") = request("frmSelfAssessmentStd8SO9Risk")
		RST("MAA9") = request("frmSelfAssessmentMAA9")
		RST("Std10bSO11b") = request("frmSelfAssessmentStd10bSO11b")
		RST("Std9SO10") = request("frmSelfAssessmentStd9SO10")
		'RST("Std9a") = request("frmSelfAssessmentStd9a")		
		RST("MAA813") = request("frmSelfAssessmentMAA813")
		RST("Std9bSO10b") = request("frmSelfAssessmentStd9bSO10b")
		RST("MAA814") = request("frmSelfAssessmentMAA814")
		RST("Std10aSO11a") = request("frmSelfAssessmentStd10aSO11a")
		RST("Std10bSO11b2") = request("frmSelfAssessmentStd10bSO11b2")
		RST("Std10gSO11f") = request("frmSelfAssessmentStd10gSO11f")
		'RST("Std10hSO11g") = request("frmSelfAssessmentStd10hSO11g")
		RST("Std10hSO11g2") = request("frmSelfAssessmentStd10hSO11g2")
		RST("Std10iSO11h") = request("frmSelfAssessmentStd10iSO11h")
		RST("Std10dSO11c") = request("frmSelfAssessmentStd10dSO11c")
		RST("Std10eSO11d") = request("frmSelfAssessmentStd10eSO11d")
		RST("Std10fSO11e") = request("frmSelfAssessmentStd10fSO11e")
		RST("Std10jSO11i") = request("frmSelfAssessmentStd10jSO11i")
		RST("MAA810") = request("frmSelfAssessmentMAA810")
		RST("Std9bSO10b2") = request("frmSelfAssessmentStd9bSO10b2")
		RST("Std11SO12") = request("frmSelfAssessmentStd11SO12")
		RST("Std11SO122") = request("frmSelfAssessmentStd11SO122")
		RST("Std12aSo13a") = request("frmSelfAssessmentStd12aSo13a")
		RST("PolicyEligible") = request("frmSelfAssessmentPolicyEligible")
		'RST("ProcEligible") = request("frmSelfAssessmentProcEligible")
		RST("PolicyChildRec") = request("frmSelfAssessmentPolicyChildRec")
		'RST("ProcChildRec") = request("frmSelfAssessmentProcChildRec")
		RST("PolicyVolRec") = request("frmSelfAssessmentPolicyVolRec")
		'RST("ProcVolRec") = request("frmSelfAssessmentProcVolRec")
		RST("PolicyRef") = request("frmSelfAssessmentPolicyRef")
		'RST("ProcRef") = request("frmSelfAssessmentProcRef")
		RST("PolicyInq") = request("frmSelfAssessmentPolicyInq")
		'RST("ProcInq") = request("frmSelfAssessmentProcInq")
		RST("PolicyIntake") = request("frmSelfAssessmentPolicyIntake")
		'RST("ProcIntake") = request("frmSelfAssessmentProcIntake")
		RST("PolicyMatch") = request("frmSelfAssessmentPolicyMatch")
		'RST("ProcMatch") = request("frmSelfAssessmentProcMatch")
		RST("PolicySup") = request("frmSelfAssessmentPolicySup")
		'RST("ProcSup") = request("frmSelfAssessmentProcSup")
		RST("PolicyClosure") = request("frmSelfAssessmentPolicyClosure")
		'RST("ProcClosure") = request("frmSelfAssessmentProcClosure")
		RST("PolicyRecords") = request("frmSelfAssessmentPolicyRecords")
		'RST("ProcRecords") = request("frmSelfAssessmentProcRecords")
		RST("PolicyOvernite") = request("frmSelfAssessmentPolicyOvernite")
		RST("PolicySexAbuse") = request("frmSelfAssessmentPolicySexAbuse")
		RST("PolicyStaffAsBigs") = request("frmSelfAssessmentPolicyStaffAsBigs")
		RST("PolicyInterOthers") = request("frmSelfAssessmentPolicyInterOthers")
		RST("PolicyPriorExp") = request("frmSelfAssessmentPolicyPriorExp")
		RST("Std13SO14") = request("frmSelfAssessmentStd13SO14")
		RST("ChildConsent") = request("frmSelfAssessmentChildConsent")
		RST("ChildInterview") = request("frmSelfAssessmentChildInterview")
		RST("ChildParInterview") = request("frmSelfAssessmentChildParInterview")
		RST("ChildHomeAssess") = request("frmSelfAssessmentChildHomeAssess")
		RST("VolConsent") = request("frmSelfAssessmentVolConsent")
		RST("VolReferences") = request("frmSelfAssessmentVolReferences")
		RST("VolCriminal") = request("frmSelfAssessmentVolCriminal")
		'RST("VolInterview") = request("frmSelfAssessmentVolInterview")
		'RST("VolHomeAssess") = request("frmSelfAssessmentVolHomeAssess")
		RST("VolMatching") = request("frmSelfAssessmentVolMatching")
		RST("VolTraining") = request("frmSelfAssessmentVolTraining")
		RST("ApprovesChild") = request("frmSelfAssessmentApprovesChild")
		RST("ApprovesParent") = request("frmSelfAssessmentApprovesParent")
		RST("ApprovesVol") = request("frmSelfAssessmentApprovesVol")
		RST("InPersonMatch") = request("frmSelfAssessmentInPersonMatch")
		RST("Std17SO18") = request("frmSelfAssessmentStd17SO18")
		RST("Std18SO19") = request("frmSelfAssessmentStd18SO19")
		RST("Std19SO20") = request("frmSelfAssessmentStd19SO20")
		RST("Std20SO21") = request("frmSelfAssessmentStd20SO21")
		RST("Std21SO22") = request("frmSelfAssessmentStd21SO22")
		RST("Std22SO23") = request("frmSelfAssessmentStd22SO23")
		RST("Std1LogoAndName") = request("frmSelfAssessmentStd1LogoAndName")
		'RST("Std6FDPlan") = request("frmSelfAssessmentStd6FDPlan")
		'RST("Std9PerformanceEval") = request("frmSelfAssessmentStd9PerformanceEval")
		'RST("Std9Notification") = request("frmSelfAssessmentStd9Notification")
		RST("Std10DiscrimPolicy") = request("frmSelfAssessmentStd10DiscrimPolicy")
		RST("Std12PPHandlingDoc") = request("frmSelfAssessmentStd12PPHandlingDoc")
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'If section<>"Operational" then 'see this for trouble shooting
		RST("Std1aReason") = request("Std1aReason")
		RST("Std1bReason") = request("Std1bReason")
		RST("Form990Reason") = request("Form990Reason")
		RST("Std1cReason") = request("Std1cReason")
		RST("BylawsReason") = request("BylawsReason")
		RST("MAAReason") = request("MAAReason")
		RST("Std1LogoAndNameReason") = request("Std1LogoAndNameReason")
		RST("Std2Reason") = request("Std2Reason")
		RST("Std2bSO3bReason") = request("Std2bSO3bReason")
		RST("MAA810confReason") = request("MAA810confReason")
		RST("Std2SOReason") = request("Std2SOReason")
		RST("Std3SO4mReason") = request("Std3SO4mReason")
		RST("Std3SO4vReason") = request("Std3SO4vReason")
		RST("Std4SO5Reason") = request("Std4SO5Reason")
		RST("Std5opsSO6Reason") = request("Std5opsSO6Reason")
		RST("Std5pgmSO6Reason") = request("Std5pgmSO6Reason")
		RST("Std5filesSO6Reason") = request("Std5filesSO6Reason")
		RST("Std6SO7budgetReason") = request("Std6SO7budgetReason")
		RST("MAA32Reason") = request("MAA32Reason")
		RST("Std6SO7bReason") = request("Std6SO7bReason")
		RST("Std6SO7Reason") = request("Std6SO7Reason")
		RST("Std7SO8Reason") = request("Std7SO8Reason")
		RST("MAA88Reason") = request("MAA88Reason")
		RST("MAA82Reason") = request("MAA82Reason")
		RST("StdSO8aReason") = request("StdSO8aReason")
		RST("StdSO8cReason") = request("StdSO8cReason")
		RST("Std8SO9riskReason") = request("Std8SO9riskReason")
		RST("MAA9Reason") = request("MAA9Reason")
		RST("Std10bSO11b2Reason") = request("Std10bSO11b2Reason")
		RST("Std9SO10Reason") = request("Std9SO10Reason")
		RST("MAA813Reason") = request("MAA813Reason")
		RST("Std9bSO10bReason") = request("Std9bSO10bReason")
		RST("MAA814Reason") = request("MAA814Reason")
		RST("Std10aSO11aReason") = request("Std10aSO11aReason")
		RST("Std10bSO11bReason") = request("Std10bSO11bReason")
		RST("Std10gSO11fReason") = request("Std10gSO11fReason")
		RST("Std10hSO11g2Reason") = request("Std10hSO11g2Reason")
		RST("Std10dSO11cReason") = request("Std10dSO11cReason")
		RST("Std10eSO11dReason") = request("Std10eSO11dReason")
		RST("Std10fSO11eReason") = request("Std10fSO11eReason")
		RST("Std10jSO11iReason") = request("Std10jSO11iReason")
		RST("MAA810Reason") = request("MAA810Reason")
		RST("Std9bSO10b2Reason") = request("Std9bSO10b2Reason")
		RST("Std10iSO11hReason") = request("Std10iSO11hReason")
		RST("Std11SO12Reason") = request("Std11SO12Reason")
		RST("Std11SO122Reason") = request("Std11SO122Reason")
	'Else
		RST("Std12aSO13aReason") = request("Std12aSO13aReason")
		RST("policyeligibleReason") = request("policyeligibleReason")
		RST("policychildrecReason") = request("policychildrecReason")
		RST("policyvolrecReason") = request("policyvolrecReason")
		RST("policyrefReason") = request("policyrefReason")
		RST("policyinqReason") = request("policyinqReason")
		RST("policyintakeReason") = request("policyintakeReason")
		RST("policymatchReason") = request("policymatchReason")
		RST("policysupReason") = request("policysupReason")
		RST("policyclosureReason") = request("policyclosureReason")
		RST("policyrecordsReason") = request("policyrecordsReason")
		RST("Std12PPHandlingDocReason") = request("Std12PPHandlingDocReason")
		RST("policyoverniteReason") = request("policyoverniteReason")
		RST("policysexabuseReason") = request("policysexabuseReason")
		RST("policystaffasbigsReason") = request("policystaffasbigsReason")
		RST("policyinterothersReason") = request("policyinterothersReason")
		RST("policypriorexpReason") = request("policypriorexpReason")
		RST("Std13SO14Reason") = request("Std13SO14Reason")
		RST("childconsentReason") = request("childconsentReason")
		RST("childinterviewReason") = request("childinterviewReason")
		RST("childparinterviewReason") = request("childparinterviewReason")
		RST("childhomeassessReason") = request("childhomeassessReason")
		RST("volconsentReason") = request("volconsentReason")
		RST("volreferencesReason") = request("volreferencesReason")
		RST("volcriminalReason") = request("volcriminalReason")
		RST("volmatchingReason") = request("volmatchingReason")
		RST("voltrainingReason") = request("voltrainingReason")
		RST("ApproveschildReason") = request("ApproveschildReason")
		RST("ApprovesparentReason") = request("ApprovesparentReason")
		RST("ApprovesvolReason") = request("ApprovesvolReason")
		RST("InpersonmatchReason") = request("InpersonmatchReason")
		RST("Std17SO18Reason") = request("Std17SO18Reason")
		RST("Std18SO19Reason") = request("Std18SO19Reason")
		RST("Std19SO20Reason") = request("Std19SO20Reason")
		RST("Std20SO21Reason") = request("Std20SO21Reason")
		RST("Std21SO22Reason") = request("Std21SO22Reason")
		RST("Std22SO23Reason") = request("Std22SO23Reason")
 'End If
		''''''''''''''''''''''''''''''''''''''
		RST("CreateDate") = Now
		RST.Update
		RST.Close
		Set RST = Nothing
		form = "SelfAssessment"
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
	RST.Open "SELECT * FROM tbl_frmSelfAssessment WHERE agencyID='" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")), Con, 1, 3
	RST("AgencyID") = Request("AgencyIDN")
	RST("Year") = Request("year")

		
		RST("Std1a") = request("frmSelfAssessmentStd1a")
		RST("Std1b") = request("frmSelfAssessmentStd1b")
		RST("Form990") = request("frmSelfAssessmentForm990")
		RST("Std1c") = request("frmSelfAssessmentStd1c")
		RST("Bylaws") = request("frmSelfAssessmentBylaws")
		RST("MAA") = request("frmSelfAssessmentMAA")
		RST("Std2") = request("frmSelfAssessmentStd2")
		'RST("Std2aSO3a") = request("frmSelfAssessmentStd2aSO3a")
		'RST("BrdTrainPlan") = request("frmSelfAssessmentBrdTrainPlan")
		RST("Std2bSO3b") = request("frmSelfAssessmentStd2bSO3b")
		RST("MAA810Conf") = request("frmSelfAssessmentMAA810Conf")
		RST("Std2SO") = request("frmSelfAssessmentStd2SO")
		RST("Std3SO4m") = request("frmSelfAssessmentStd3SO4m")
		RST("Std3SO4v") = request("frmSelfAssessmentStd3SO4v")
		RST("Std4SO5") = request("frmSelfAssessmentStd4SO5")
		RST("OpPlan") = request("frmSelfAssessmentOpPlan")
		RST("Std5OpsSO6") = request("frmSelfAssessmentStd5OpsSO6")
		RST("Std5PgmSO6") = request("frmSelfAssessmentStd5PgmSO6")
		RST("Std5FilesSO6") = request("frmSelfAssessmentStd5FilesSO6")
		RST("Std6SO7b") = request("frmSelfAssessmentStd6SO7b")
		RST("Std6SO7") = request("frmSelfAssessmentStd6SO7")
		RST("Std6SO7Budget") = request("frmSelfAssessmentStd6SO7Budget")
		RST("MAA810Exp") = request("frmSelfAssessmentMAA810Exp")
		RST("MAA32") = request("frmSelfAssessmentMAA32")
		RST("Std7SO8") = request("frmSelfAssessmentStd7SO8")
		RST("MAA88") = request("frmSelfAssessmentMAA88")
		RST("MAA82") = request("frmSelfAssessmentMAA82")
		RST("StdSO8a") = request("frmSelfAssessmentStdSO8a")
		'RST("StdSO8b") = request("frmSelfAssessmentStdSO8b")
		RST("StdSO8c") = request("frmSelfAssessmentStdSO8c")
		'RST("Std8SO9Crisis") = request("frmSelfAssessmentStd8SO9Crisis")
		RST("Std8SO9Risk") = request("frmSelfAssessmentStd8SO9Risk")
		RST("MAA9") = request("frmSelfAssessmentMAA9")
		RST("Std10bSO11b") = request("frmSelfAssessmentStd10bSO11b")
		RST("Std9SO10") = request("frmSelfAssessmentStd9SO10")
		'RST("Std9a") = request("frmSelfAssessmentStd9a")			
		RST("MAA813") = request("frmSelfAssessmentMAA813")
		RST("Std9bSO10b") = request("frmSelfAssessmentStd9bSO10b")
		RST("MAA814") = request("frmSelfAssessmentMAA814")
		RST("Std10aSO11a") = request("frmSelfAssessmentStd10aSO11a")
		RST("Std10bSO11b2") = request("frmSelfAssessmentStd10bSO11b2")
		RST("Std10gSO11f") = request("frmSelfAssessmentStd10gSO11f")
		'RST("Std10hSO11g") = request("frmSelfAssessmentStd10hSO11g")
		RST("Std10hSO11g2") = request("frmSelfAssessmentStd10hSO11g2")
		RST("Std10iSO11h") = request("frmSelfAssessmentStd10iSO11h")
		RST("Std10dSO11c") = request("frmSelfAssessmentStd10dSO11c")
		RST("Std10eSO11d") = request("frmSelfAssessmentStd10eSO11d")
		RST("Std10fSO11e") = request("frmSelfAssessmentStd10fSO11e")
		RST("Std10jSO11i") = request("frmSelfAssessmentStd10jSO11i")
		RST("MAA810") = request("frmSelfAssessmentMAA810")
		RST("Std9bSO10b2") = request("frmSelfAssessmentStd9bSO10b2")
		RST("Std11SO12") = request("frmSelfAssessmentStd11SO12")
		RST("Std11SO122") = request("frmSelfAssessmentStd11SO122")
		RST("Std12aSo13a") = request("frmSelfAssessmentStd12aSo13a")
		RST("PolicyEligible") = request("frmSelfAssessmentPolicyEligible")
		'RST("ProcEligible") = request("frmSelfAssessmentProcEligible")
		RST("PolicyChildRec") = request("frmSelfAssessmentPolicyChildRec")
		'RST("ProcChildRec") = request("frmSelfAssessmentProcChildRec")
		RST("PolicyVolRec") = request("frmSelfAssessmentPolicyVolRec")
		'RST("ProcVolRec") = request("frmSelfAssessmentProcVolRec")
		RST("PolicyRef") = request("frmSelfAssessmentPolicyRef")
		'RST("ProcRef") = request("frmSelfAssessmentProcRef")
		RST("PolicyInq") = request("frmSelfAssessmentPolicyInq")
		'RST("ProcInq") = request("frmSelfAssessmentProcInq")
		RST("PolicyIntake") = request("frmSelfAssessmentPolicyIntake")
		'RST("ProcIntake") = request("frmSelfAssessmentProcIntake")
		RST("PolicyMatch") = request("frmSelfAssessmentPolicyMatch")
		'RST("ProcMatch") = request("frmSelfAssessmentProcMatch")
		RST("PolicySup") = request("frmSelfAssessmentPolicySup")
		'RST("ProcSup") = request("frmSelfAssessmentProcSup")
		RST("PolicyClosure") = request("frmSelfAssessmentPolicyClosure")
		'RST("ProcClosure") = request("frmSelfAssessmentProcClosure")
		RST("PolicyRecords") = request("frmSelfAssessmentPolicyRecords")
		'RST("ProcRecords") = request("frmSelfAssessmentProcRecords")
		RST("PolicyOvernite") = request("frmSelfAssessmentPolicyOvernite")
		RST("PolicySexAbuse") = request("frmSelfAssessmentPolicySexAbuse")
		RST("PolicyStaffAsBigs") = request("frmSelfAssessmentPolicyStaffAsBigs")
		RST("PolicyInterOthers") = request("frmSelfAssessmentPolicyInterOthers")
		RST("PolicyPriorExp") = request("frmSelfAssessmentPolicyPriorExp")
		RST("Std13SO14") = request("frmSelfAssessmentStd13SO14")
		RST("ChildConsent") = request("frmSelfAssessmentChildConsent")
		RST("ChildInterview") = request("frmSelfAssessmentChildInterview")
		RST("ChildParInterview") = request("frmSelfAssessmentChildParInterview")
		RST("ChildHomeAssess") = request("frmSelfAssessmentChildHomeAssess")
		RST("VolConsent") = request("frmSelfAssessmentVolConsent")
		RST("VolReferences") = request("frmSelfAssessmentVolReferences")
		RST("VolCriminal") = request("frmSelfAssessmentVolCriminal")
		'RST("VolInterview") = request("frmSelfAssessmentVolInterview")
		'RST("VolHomeAssess") = request("frmSelfAssessmentVolHomeAssess")
		RST("VolMatching") = request("frmSelfAssessmentVolMatching")
		RST("VolTraining") = request("frmSelfAssessmentVolTraining")
		RST("ApprovesChild") = request("frmSelfAssessmentApprovesChild")
		RST("ApprovesParent") = request("frmSelfAssessmentApprovesParent")
		RST("ApprovesVol") = request("frmSelfAssessmentApprovesVol")
		RST("InPersonMatch") = request("frmSelfAssessmentInPersonMatch")
		RST("Std17SO18") = request("frmSelfAssessmentStd17SO18")
		RST("Std18SO19") = request("frmSelfAssessmentStd18SO19")
		RST("Std19SO20") = request("frmSelfAssessmentStd19SO20")
		RST("Std20SO21") = request("frmSelfAssessmentStd20SO21")
		RST("Std21SO22") = request("frmSelfAssessmentStd21SO22")
		RST("Std22SO23") = request("frmSelfAssessmentStd22SO23")
		RST("Std1LogoAndName") = request("frmSelfAssessmentStd1LogoAndName")
		'RST("Std6FDPlan") = request("frmSelfAssessmentStd6FDPlan")
		'RST("Std9PerformanceEval") = request("frmSelfAssessmentStd9PerformanceEval")
		'RST("Std9Notification") = request("frmSelfAssessmentStd9Notification")
		RST("Std10DiscrimPolicy") = request("frmSelfAssessmentStd10DiscrimPolicy")
		RST("Std12PPHandlingDoc") = request("frmSelfAssessmentStd12PPHandlingDoc")
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'If section<>"Operational" then 'see this for trouble shooting
		If RST("Std1a") = 1 then RST("Std1aReason") = request("Std1aReason") End If
		If RST("Std1b") = 1 then RST("Std1bReason") = request("Std1bReason") End If
		If RST("Form990") = 1 then RST("Form990Reason") = request("Form990Reason") End If
		If RST("Std1c") = 1 then RST("Std1cReason") = request("Std1cReason") End If
		If RST("Bylaws") = 1 then RST("BylawsReason") = request("BylawsReason") End If
		If RST("MAA") = 1 then RST("MAAReason") = request("MAAReason") End If
		If RST("Std1LogoAndName") = 1 then RST("Std1LogoAndNameReason") = request("Std1LogoAndNameReason") End If
		If RST("Std2") = 1 then RST("Std2Reason") = request("Std2Reason") End If
		If RST("Std2bSO3b") = 1 then RST("Std2bSO3bReason") = request("Std2bSO3bReason") End If
		If RST("MAA810conf") = 1 then RST("MAA810confReason") = request("MAA810confReason") End If
		If RST("Std2SO") = 1 then RST("Std2SOReason") = request("Std2SOReason") End If
		If RST("Std3SO4m") = 1 then RST("Std3SO4mReason") = request("Std3SO4mReason") End If
		If RST("Std3SO4v") = 1 then RST("Std3SO4vReason") = request("Std3SO4vReason") End If
		If RST("Std4SO5") = 1 then RST("Std4SO5Reason") = request("Std4SO5Reason") End If
		If RST("Std5opsSO6") = 1 then RST("Std5opsSO6Reason") = request("Std5opsSO6Reason") End If
		If RST("Std5pgmSO6") = 1 then RST("Std5pgmSO6Reason") = request("Std5pgmSO6Reason") End If
		If RST("Std5filesSO6") = 1 then RST("Std5filesSO6Reason") = request("Std5filesSO6Reason") End If
		If RST("Std6SO7budget") = 1 then RST("Std6SO7budgetReason") = request("Std6SO7budgetReason") End If
		If RST("MAA32") = 1 then RST("MAA32Reason") = request("MAA32Reason") End If
		If RST("Std6SO7b") = 1 then RST("Std6SO7bReason") = request("Std6SO7bReason") End If
		If RST("Std6SO7") = 1 then RST("Std6SO7Reason") = request("Std6SO7Reason") End If
		If RST("Std7SO8") = 1 then RST("Std7SO8Reason") = request("Std7SO8Reason") End If
		If RST("MAA88") = 1 then RST("MAA88Reason") = request("MAA88Reason") End If
		If RST("MAA82") = 1 then RST("MAA82Reason") = request("MAA82Reason") End If
		If RST("StdSO8a") = 1 then RST("StdSO8aReason") = request("StdSO8aReason") End If
		If RST("StdSO8c") = 1 then RST("StdSO8cReason") = request("StdSO8cReason") End If
		If RST("Std8SO9risk") = 1 then RST("Std8SO9riskReason") = request("Std8SO9riskReason") End If
		If RST("MAA9") = 1 then RST("MAA9Reason") = request("MAA9Reason") End If
		If RST("Std10bSO11b2") = 1 then RST("Std10bSO11b2Reason") = request("Std10bSO11b2Reason") End If
		If RST("Std9SO10") = 1 then RST("Std9SO10Reason") = request("Std9SO10Reason") End If
		If RST("MAA813") = 1 then RST("MAA813Reason") = request("MAA813Reason") End If
		If RST("Std9bSO10b") = 1 then RST("Std9bSO10bReason") = request("Std9bSO10bReason") End If
		If RST("MAA814") = 1 then RST("MAA814Reason") = request("MAA814Reason") End If
		If RST("Std10aSO11a") = 1 then RST("Std10aSO11aReason") = request("Std10aSO11aReason") End If
		If RST("Std10bSO11b") = 1 then RST("Std10bSO11bReason") = request("Std10bSO11bReason") End If
		If RST("Std10gSO11f") = 1 then RST("Std10gSO11fReason") = request("Std10gSO11fReason") End If
		If RST("Std10hSO11g2") = 1 then RST("Std10hSO11g2Reason") = request("Std10hSO11g2Reason") End If
		If RST("Std10dSO11c") = 1 then RST("Std10dSO11cReason") = request("Std10dSO11cReason") End If
		If RST("Std10eSO11d") = 1 then RST("Std10eSO11dReason") = request("Std10eSO11dReason") End If
		If RST("Std10fSO11e") = 1 then RST("Std10fSO11eReason") = request("Std10fSO11eReason") End If
		If RST("Std10jSO11i") = 1 then RST("Std10jSO11iReason") = request("Std10jSO11iReason") End If
		If RST("MAA810") = 1 then RST("MAA810Reason") = request("MAA810Reason") End If
		If RST("Std9bSO10b2") = 1 then RST("Std9bSO10b2Reason") = request("Std9bSO10b2Reason") End If
		If RST("Std10iSO11h") = 1 then RST("Std10iSO11hReason") = request("Std10iSO11hReason") End If
		If RST("Std11SO12") = 1 then RST("Std11SO12Reason") = request("Std11SO12Reason") End If
		If RST("Std11SO122") = 1 then RST("Std11SO122Reason") = request("Std11SO122Reason") End If
		
		'Else
		If RST("Std12aSO13a") = 1 then RST("Std12aSO13aReason") = request("Std12aSO13aReason") End If
		If RST("policyeligible") = 1 then RST("policyeligibleReason") = request("policyeligibleReason") End If
		If RST("policychildrec") = 1 then RST("policychildrecReason") = request("policychildrecReason") End If
		If RST("policyvolrec") = 1 then RST("policyvolrecReason") = request("policyvolrecReason") End If
		If RST("policyref") = 1 then RST("policyrefReason") = request("policyrefReason") End If
		If RST("policyinq") = 1 then RST("policyinqReason") = request("policyinqReason") End If
		If RST("policyintake") = 1 then RST("policyintakeReason") = request("policyintakeReason") End If
		If RST("policymatch") = 1 then RST("policymatchReason") = request("policymatchReason") End If
		If RST("policysup") = 1 then RST("policysupReason") = request("policysupReason") End If
		If RST("policyclosure") = 1 then RST("policyclosureReason") = request("policyclosureReason") End If
		If RST("policyrecords") = 1 then RST("policyrecordsReason") = request("policyrecordsReason") End If
		If RST("Std12PPHandlingDoc") = 1 then RST("Std12PPHandlingDocReason") = request("Std12PPHandlingDocReason") End If
		If RST("policyovernite") = 1 then RST("policyoverniteReason") = request("policyoverniteReason") End If
		If RST("policysexabuse") = 1 then RST("policysexabuseReason") = request("policysexabuseReason") End If
		If RST("policystaffasbigs") = 1 then RST("policystaffasbigsReason") = request("policystaffasbigsReason") End If
		If RST("policyinterothers") = 1 then RST("policyinterothersReason") = request("policyinterothersReason") End If
		If RST("policypriorexp") = 1 then RST("policypriorexpReason") = request("policypriorexpReason") End If
		If RST("Std13SO14") = 1 then RST("Std13SO14Reason") = request("Std13SO14Reason") End If
		If RST("childconsent") = 1 then RST("childconsentReason") = request("childconsentReason") End If
		If RST("childinterview") = 1 then RST("childinterviewReason") = request("childinterviewReason") End If
		If RST("childparinterview") = 1 then RST("childparinterviewReason") = request("childparinterviewReason") End If
		If RST("childhomeassess") = 1 then RST("childhomeassessReason") = request("childhomeassessReason") End If
		If RST("volconsent") = 1 then RST("volconsentReason") = request("volconsentReason") End If
		If RST("volreferences") = 1 then RST("volreferencesReason") = request("volreferencesReason") End If
		If RST("volcriminal") = 1 then RST("volcriminalReason") = request("volcriminalReason") End If
		If RST("volmatching") = 1 then RST("volmatchingReason") = request("volmatchingReason") End If
		If RST("voltraining") = 1 then RST("voltrainingReason") = request("voltrainingReason") End If
		If RST("Approveschild") = 1 then RST("ApproveschildReason") = request("ApproveschildReason") End If
		If RST("Approvesparent") = 1 then RST("ApprovesparentReason") = request("ApprovesparentReason") End If
		If RST("Approvesvol") = 1 then RST("ApprovesvolReason") = request("ApprovesvolReason") End If
		If RST("Inpersonmatch") = 1 then RST("InpersonmatchReason") = request("InpersonmatchReason") End If
		If RST("Std17SO18") = 1 then RST("Std17SO18Reason") = request("Std17SO18Reason") End If
		If RST("Std18SO19") = 1 then RST("Std18SO19Reason") = request("Std18SO19Reason") End If
		If RST("Std19SO20") = 1 then RST("Std19SO20Reason") = request("Std19SO20Reason") End If
		If RST("Std20SO21") = 1 then RST("Std20SO21Reason") = request("Std20SO21Reason") End If
		If RST("Std21SO22") = 1 then RST("Std21SO22Reason") = request("Std21SO22Reason") End If
		If RST("Std22SO23") = 1 then RST("Std22SO23Reason") = request("Std22SO23Reason") End If
		'End If

	jMod = RST("SelfAssessmentID")
	
	RST.Update
	RST.Close
	Set RST = Nothing
	form = "SelfAssessment"
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
	<title>Agency Self Assessment Form</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
	
<SCRIPT LANGUAGE = "JavaScript">

	function formvalidation(f) {
		if ((frmSelfAssessment.frmSelfAssessmentStd1a[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd1b[2].checked) || (frmSelfAssessment.frmSelfAssessmentForm990[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else if ((frmSelfAssessment.frmSelfAssessmentStd1c[2].checked) || (frmSelfAssessment.frmSelfAssessmentBylaws[2].checked) || (frmSelfAssessment.frmSelfAssessmentMAA[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else if ((frmSelfAssessment.frmSelfAssessmentStd2[2].checked) /*|| (frmSelfAssessment.frmSelfAssessmentStd2aSO3a[2].checked) || (frmSelfAssessment.frmSelfAssessmentBrdtrainplan[2].checked)*/)
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		
		
		else if ((frmSelfAssessment.frmSelfAssessmentStd2bSO3b[2].checked) || (frmSelfAssessment.frmSelfAssessmentMAA810conf[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd2SO[3].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else if ((frmSelfAssessment.frmSelfAssessmentStd3SO4m[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd3SO4v[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd4SO5[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else if ((frmSelfAssessment.frmSelfAssessmentStd5opsSO6[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd5pgmSO6[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd5filesSO6[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		
		else if ((frmSelfAssessment.frmSelfAssessmentStd6SO7budget[2].checked) || (frmSelfAssessment.frmSelfAssessmentMAA32[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd6SO7b[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else if ((frmSelfAssessment.frmSelfAssessmentStd6SO7[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd7SO8[2].checked) || (frmSelfAssessment.frmSelfAssessmentMAA88[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else if ((frmSelfAssessment.frmSelfAssessmentMAA82[2].checked) || (frmSelfAssessment.frmSelfAssessmentStdSO8a[3].checked) /*|| (frmSelfAssessment.frmSelfAssessmentStdSO8b[3].checked)*/)
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else if ((frmSelfAssessment.frmSelfAssessmentStdSO8c[3].checked) /*|| (frmSelfAssessment.frmSelfAssessmentStd8SO9crisis[2].checked)*/ || (frmSelfAssessment.frmSelfAssessmentStd8SO9risk[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		
		else if ((frmSelfAssessment.frmSelfAssessmentMAA9[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd10bSO11b2[2].checked) /* || (frmSelfAssessment.frmSelfAssessmentStd9a[2].checked)*/)
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else if ((frmSelfAssessment.frmSelfAssessmentStd9SO10[2].checked) || (frmSelfAssessment.frmSelfAssessmentMAA813[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd9bSO10b[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else if ((frmSelfAssessment.frmSelfAssessmentMAA814[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd10aSO11a[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd10bSO11b[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		
		else if ((frmSelfAssessment.frmSelfAssessmentStd10gSO11f[3].checked) /*|| (frmSelfAssessment.frmSelfAssessmentStd10hSO11g[3].checked)*/ || (frmSelfAssessment.frmSelfAssessmentStd10hSO11g2[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else if ((frmSelfAssessment.frmSelfAssessmentStd10iSO11h[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd10dSO11c[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd10eSO11d[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else if ((frmSelfAssessment.frmSelfAssessmentStd10fSO11e[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd10jSO11i[2].checked) || (frmSelfAssessment.frmSelfAssessmentMAA810[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else if ((frmSelfAssessment.frmSelfAssessmentStd9bSO10b2[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd11SO12[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd11SO122[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else
			frmSelfAssessment.submit();
	}
	
	function formvalidationPr(f) {
		if ((frmSelfAssessment.frmSelfAssessmentStd12aSO13a[2].checked) || (frmSelfAssessment.frmSelfAssessmentpolicyeligible[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		//else if ((frmSelfAssessment.frmSelfAssessmentpolicychildrec[2].checked) || (frmSelfAssessment.frmSelfAssessmentprocchildrec[2].checked) || (frmSelfAssessment.frmSelfAssessmentpolicyvolrec[2].checked))
		else if ((frmSelfAssessment.frmSelfAssessmentpolicychildrec[2].checked) || (frmSelfAssessment.frmSelfAssessmentpolicyvolrec[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		//else if ((frmSelfAssessment.frmSelfAssessmentprocvolrec[2].checked) || (frmSelfAssessment.frmSelfAssessmentpolicyref[2].checked) || (frmSelfAssessment.frmSelfAssessmentprocref[2].checked))
		else if ((frmSelfAssessment.frmSelfAssessmentpolicyref[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		
		//else if ((frmSelfAssessment.frmSelfAssessmentpolicyinq[2].checked) || (frmSelfAssessment.frmSelfAssessmentprocinq[2].checked) || (frmSelfAssessment.frmSelfAssessmentpolicyintake[2].checked))
		else if ((frmSelfAssessment.frmSelfAssessmentpolicyinq[2].checked) || (frmSelfAssessment.frmSelfAssessmentpolicyintake[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		//else if ((frmSelfAssessment.frmSelfAssessmentprocintake[2].checked) || (frmSelfAssessment.frmSelfAssessmentpolicymatch[2].checked) || (frmSelfAssessment.frmSelfAssessmentprocmatch[2].checked))
		else if ((frmSelfAssessment.frmSelfAssessmentpolicymatch[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		//else if ((frmSelfAssessment.frmSelfAssessmentpolicysup[2].checked) || (frmSelfAssessment.frmSelfAssessmentprocsup[2].checked) || (frmSelfAssessment.frmSelfAssessmentpolicyclosure[2].checked))
		else if ((frmSelfAssessment.frmSelfAssessmentpolicysup[2].checked) || (frmSelfAssessment.frmSelfAssessmentpolicyclosure[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		
		//else if ((frmSelfAssessment.frmSelfAssessmentprocclosure[2].checked) || (frmSelfAssessment.frmSelfAssessmentprocrecords[2].checked) || (frmSelfAssessment.frmSelfAssessmentpolicyovernite[2].checked))
		else if ((frmSelfAssessment.frmSelfAssessmentpolicyovernite[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else if ((frmSelfAssessment.frmSelfAssessmentpolicysexabuse[2].checked) || (frmSelfAssessment.frmSelfAssessmentpolicystaffasbigs[2].checked) || (frmSelfAssessment.frmSelfAssessmentpolicyinterothers[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else if ((frmSelfAssessment.frmSelfAssessmentpolicypriorexp[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd13SO14[2].checked) || (frmSelfAssessment.frmSelfAssessmentchildconsent[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else if ((frmSelfAssessment.frmSelfAssessmentchildinterview[2].checked) || (frmSelfAssessment.frmSelfAssessmentchildparinterview[2].checked) || (frmSelfAssessment.frmSelfAssessmentchildhomeassess[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		
		else if ((frmSelfAssessment.frmSelfAssessmentvolconsent[2].checked) || (frmSelfAssessment.frmSelfAssessmentvolreferences[2].checked) || (frmSelfAssessment.frmSelfAssessmentvolcriminal[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else if (/*(frmSelfAssessment.frmSelfAssessmentvolinterview[2].checked) || (frmSelfAssessment.frmSelfAssessmentvolhomeassess[2].checked) || */(frmSelfAssessment.frmSelfAssessmentvolmatching[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else if ((frmSelfAssessment.frmSelfAssessmentvoltraining[2].checked) || (frmSelfAssessment.frmSelfAssessmentApproveschild[2].checked) || (frmSelfAssessment.frmSelfAssessmentApprovesparent[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		
		else if ((frmSelfAssessment.frmSelfAssessmentApprovesvol[2].checked) || (frmSelfAssessment.frmSelfAssessmentInpersonmatch[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd17SO18[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else if ((frmSelfAssessment.frmSelfAssessmentStd18SO19[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd19SO20[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd20SO21[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else if ((frmSelfAssessment.frmSelfAssessmentStd21SO22[2].checked) || (frmSelfAssessment.frmSelfAssessmentStd22SO23[2].checked))
		{
			alert("Some of the questions were not answered. Please answer all of them.");
			return false;
		}
		else
			frmSelfAssessment.submit();
	}
	
	function disableEnable(form,option_name,inout)
	{
		var area = document.getElementById(option_name); //document.getElementById('jarea');
		if (inout)
		{
		 	 switch(option_name)
			 {
			 	//first section
				case 'Std1a': area.style.display = 'block'; form.Std1aReason.focus(); break;
				case 'Std1b': area.style.display = 'block'; form.Std1bReason.focus(); break;
				case 'Form990': area.style.display = 'block'; form.Form990Reason.focus(); break;
				case 'Std1c': area.style.display = 'block'; form.Std1cReason.focus(); break;
				case 'Bylaws': area.style.display = 'block'; form.BylawsReason.focus(); break;
				case 'MAA': area.style.display = 'block'; form.MAAReason.focus(); break;
				case 'Std1LogoAndName': area.style.display = 'block'; form.Std1LogoAndNameReason.focus(); break;
				
				//Board development
				case 'Std2': area.style.display = 'block'; form.Std2Reason.focus(); break;
				case 'Std2aSO3a': area.style.display = 'block'; form.Std2aSO3aReason.focus(); break;
				case 'Brdtrainplan': area.style.display = 'block'; form.BrdtrainplanReason.focus(); break;
				case 'Std2bSO3b': area.style.display = 'block'; form.Std2bSO3bReason.focus(); break;
				case 'MAA810conf': area.style.display = 'block'; form.MAA810confReason.focus(); break;
				case 'Std2SO': area.style.display = 'block'; form.Std2SOReason.focus(); break;
				
				//Mission/Vision
				case 'Std3SO4m': area.style.display = 'block'; form.Std3SO4mReason.focus(); break;
				case 'Std3SO4v': area.style.display = 'block'; form.Std3SO4vReason.focus(); break;
				
				//Strategic Planning
				case 'Std4SO5': area.style.display = 'block'; form.Std4SO5Reason.focus(); break;
				
				//Board development
				case 'Std5opsSO6': area.style.display = 'block'; form.Std5opsSO6Reason.focus(); break;
				case 'Std5pgmSO6': area.style.display = 'block'; form.Std5pgmSO6Reason.focus(); break;
				case 'Std5filesSO6': area.style.display = 'block'; form.Std5filesSO6Reason.focus(); break;
				
				//Fund Development
				case 'Std6SO7budget': area.style.display = 'block'; form.Std6SO7budgetReason.focus(); break;
				case 'MAA32': area.style.display = 'block'; form.MAA32Reason.focus(); break;
				case 'Std6SO7b': area.style.display = 'block'; form.Std6SO7bReason.focus(); break;
				
			 	//Financial Management
				case 'Std6SO7': area.style.display = 'block'; form.Std6SO7Reason.focus(); break;
				case 'Std7SO8': area.style.display = 'block'; form.Std7SO8Reason.focus(); break;
				case 'MAA88': area.style.display = 'block'; form.MAA88Reason.focus(); break;
				case 'MAA82': area.style.display = 'block'; form.MAA82Reason.focus(); break;
				case 'StdSO8a': area.style.display = 'block'; form.StdSO8aReason.focus(); break;
				case 'StdSO8b': area.style.display = 'block'; form.StdSO8bReason.focus(); break;
				case 'StdSO8c': area.style.display = 'block'; form.StdSO8cReason.focus(); break;
				
				//Risk management
				case 'Std8SO9crisis': area.style.display = 'block'; form.Std8SO9crisisReason.focus(); break;
				case 'Std8SO9risk': area.style.display = 'block'; form.Std8SO9riskReason.focus(); break;
				case 'MAA9': area.style.display = 'block'; form.MAA9Reason.focus(); break;
				
				//Personel
				case 'Std10bSO11b2': area.style.display = 'block'; form.Std10bSO11b2Reason.focus(); break;
				case 'Std9a': area.style.display = 'block'; form.Std9aReason.focus(); break;
				case 'Std9SO10': area.style.display = 'block'; form.Std9SO10Reason.focus(); break;
				case 'MAA813': area.style.display = 'block'; form.MAA813Reason.focus(); break;
				case 'Std9bSO10b': area.style.display = 'block'; form.Std9bSO10bReason.focus(); break;
				case 'MAA814': area.style.display = 'block'; form.MAA814Reason.focus(); break;
				
			 	//Personel Part 2
				case 'Std10aSO11a': area.style.display = 'block'; form.Std10aSO11aReason.focus(); break;
				case 'Std10bSO11b': area.style.display = 'block'; form.Std10bSO11bReason.focus(); break;
				case 'Std10gSO11f': area.style.display = 'block'; form.Std10gSO11fReason.focus(); break;
				//case 'Std10hSO11g': area.style.display = 'block'; form.Std10hSO11gReason.focus(); break;
				case 'Std10hSO11g2': area.style.display = 'block'; form.Std10hSO11g2Reason.focus(); break;
				case 'Std10iSO11h': area.style.display = 'block'; form.Std10iSO11hReason.focus(); break;
				case 'Std10dSO11c': area.style.display = 'block'; form.Std10dSO11cReason.focus(); break;
				case 'Std10eSO11d': area.style.display = 'block'; form.Std10eSO11dReason.focus(); break;
				
				case 'Std10fSO11e': area.style.display = 'block'; form.Std10fSO11eReason.focus(); break;
				case 'Std10jSO11i': area.style.display = 'block'; form.Std10jSO11iReason.focus(); break;
				case 'MAA810': area.style.display = 'block'; form.MAA810Reason.focus(); break;
				case 'Std9bSO10b2': area.style.display = 'block'; form.Std9bSO10b2Reason.focus(); break;
				case 'Std11SO12': area.style.display = 'block'; form.Std11SO12Reason.focus(); break;
				case 'Std11SO122': area.style.display = 'block'; form.Std11SO122Reason.focus(); break;
				
			 	//PROGRAM STANDARDS
				case 'Std12aSO13a': area.style.display = 'block'; form.Std12aSO13aReason.focus(); break;
				case 'policyeligible': area.style.display = 'block'; form.policyeligibleReason.focus(); break;
				//case 'proceligible': area.style.display = 'block'; form.proceligibleReason.focus(); break;
				case 'policychildrec': area.style.display = 'block'; form.policychildrecReason.focus(); break;
				//case 'procchildrec': area.style.display = 'block'; form.procchildrecReason.focus(); break;
				case 'policyvolrec': area.style.display = 'block'; form.policyvolrecReason.focus(); break;
				//case 'procvolrec': area.style.display = 'block'; form.procvolrecReason.focus(); break;
				case 'policyref': area.style.display = 'block'; form.policyrefReason.focus(); break;
				
				//case 'procref': area.style.display = 'block'; form.procrefReason.focus(); break;
				case 'policyinq': area.style.display = 'block'; form.policyinqReason.focus(); break;
				//case 'procinq': area.style.display = 'block'; form.procinqReason.focus(); break;
				case 'policyintake': area.style.display = 'block'; form.policyintakeReason.focus(); break;
				//case 'procintake': area.style.display = 'block'; form.procintakeReason.focus(); break;
				case 'policymatch': area.style.display = 'block'; form.policymatchReason.focus(); break;
				
				//case 'procmatch': area.style.display = 'block'; form.procmatchReason.focus(); break;
				case 'policysup': area.style.display = 'block'; form.policysupReason.focus(); break;
				//case 'procsup': area.style.display = 'block'; form.procsupReason.focus(); break;
				case 'policyclosure': area.style.display = 'block'; form.policyclosureReason.focus(); break;
				//case 'procclosure': area.style.display = 'block'; form.procclosureReason.focus(); break;
				case 'policyrecords': area.style.display = 'block'; form.policyrecordsReason.focus(); break;
				case 'Std12PPHandlingDoc': area.style.display = 'block'; form.Std12PPHandlingDocReason.focus(); break;
				
				
			 	//Program Manual addresses risk management issues
				case 'policyovernite': area.style.display = 'block'; form.policyoverniteReason.focus(); break;
				case 'policysexabuse': area.style.display = 'block'; form.policysexabuseReason.focus(); break;
				case 'policystaffasbigs': area.style.display = 'block'; form.policystaffasbigsReason.focus(); break;
				case 'policyinterothers': area.style.display = 'block'; form.policyinterothersReason.focus(); break;
				case 'policypriorexp': area.style.display = 'block'; form.policypriorexpReason.focus(); break;
				
			 	//Financial Management
				case 'Std13SO14': area.style.display = 'block'; form.Std13SO14Reason.focus(); break;
				case 'childconsent': area.style.display = 'block'; form.childconsentReason.focus(); break;
				case 'childinterview': area.style.display = 'block'; form.childinterviewReason.focus(); break;
				case 'childparinterview': area.style.display = 'block'; form.childparinterviewReason.focus(); break;
				case 'childhomeassess': area.style.display = 'block'; form.childhomeassessReason.focus(); break;
				case 'volconsent': area.style.display = 'block'; form.volconsentReason.focus(); break;
				case 'volreferences': area.style.display = 'block'; form.volreferencesReason.focus(); break;
				case 'volcriminal': area.style.display = 'block'; form.volcriminalReason.focus(); break;
				case 'volinterview': area.style.display = 'block'; form.volinterviewReason.focus(); break;
				case 'volhomeassess': area.style.display = 'block'; form.volhomeassessReason.focus(); break;
				case 'volmatching': area.style.display = 'block'; form.volmatchingReason.focus(); break;
				case 'voltraining': area.style.display = 'block'; form.voltrainingReason.focus(); break;
				
			 	//Standard 16 -17
				case 'Approveschild': area.style.display = 'block'; form.ApproveschildReason.focus(); break;
				case 'Approvesparent': area.style.display = 'block'; form.ApprovesparentReason.focus(); break;
				case 'Approvesvol': area.style.display = 'block'; form.ApprovesvolReason.focus(); break;
				case 'Inpersonmatch': area.style.display = 'block'; form.InpersonmatchReason.focus(); break;
				case 'Std17SO18': area.style.display = 'block'; form.Std17SO18Reason.focus(); break;
				case 'Std18SO19': area.style.display = 'block'; form.Std18SO19Reason.focus(); break;
				case 'Std19SO20': area.style.display = 'block'; form.Std19SO20Reason.focus(); break;
				case 'Std20SO21': area.style.display = 'block'; form.Std20SO21Reason.focus(); break;
				case 'Std21SO22': area.style.display = 'block'; form.Std21SO22Reason.focus(); break;
				case 'Std22SO23': area.style.display = 'block'; form.Std22SO23Reason.focus(); break;
			 }
		}
		else
		{
		 	 switch(option_name)
			 {
			  //first section
				case 'Std1a': area.style.display = 'none'; break;
				case 'Std1b': area.style.display = 'none'; break;
				case 'Form990': area.style.display = 'none'; break;
				case 'Std1c': area.style.display = 'none'; break;
				case 'Bylaws': area.style.display = 'none'; break;
				case 'MAA': area.style.display = 'none'; break;
				case 'Std1LogoAndName': area.style.display = 'none'; break;
				
				//Board development
				case 'Std2': area.style.display = 'none'; break;
				case 'Std2aSO3a': area.style.display = 'none'; break;
				case 'Brdtrainplan': area.style.display = 'none'; break;
				case 'Std2bSO3b': area.style.display = 'none'; break;
				case 'MAA810conf': area.style.display = 'none'; break;
				case 'Std2SO': area.style.display = 'none'; break;
				
				//Mission/Vision
				case 'Std3SO4m': area.style.display = 'none'; break;
				case 'Std3SO4v': area.style.display = 'none'; break;
				
				//Strategic Planning
				case 'Std4SO5': area.style.display = 'none'; break;

				//Board development
				case 'Std5opsSO6': area.style.display = 'none'; break;
				case 'Std5pgmSO6': area.style.display = 'none'; break;
				case 'Std5filesSO6': area.style.display = 'none'; break;
				
				//Fund Development
				case 'Std6SO7budget': area.style.display = 'none'; break;
				case 'MAA32': area.style.display = 'none'; break;
				case 'Std6SO7b': area.style.display = 'none'; break;

			  //Financial Management
				case 'Std6SO7': area.style.display = 'none'; break;
				case 'Std7SO8': area.style.display = 'none'; break;
				case 'MAA88': area.style.display = 'none'; break;
				case 'MAA82': area.style.display = 'none'; break;
				case 'StdSO8a': area.style.display = 'none'; break;
				case 'StdSO8b': area.style.display = 'none'; break;
				case 'StdSO8c': area.style.display = 'none'; break;
				
				//Risk Management
				case 'Std8SO9crisis': area.style.display = 'none'; break;
				case 'Std8SO9risk': area.style.display = 'none'; break;
				case 'MAA9': area.style.display = 'none'; break;
				
				//Personel
				case 'Std10bSO11b2': area.style.display = 'none'; break;
				case 'Std9a': area.style.display = 'none'; break;
				case 'Std9SO10': area.style.display = 'none'; break;
				case 'MAA813': area.style.display = 'none'; break;
				case 'Std9bSO10b': area.style.display = 'none'; break;
				case 'MAA814': area.style.display = 'none'; break;
				
			  //Personel part 2
				case 'Std10aSO11a': area.style.display = 'none'; break;
				case 'Std10bSO11b': area.style.display = 'none'; break;
				case 'Std10gSO11f': area.style.display = 'none'; break;
				case 'Std10hSO11g': area.style.display = 'none'; break;
				case 'Std10hSO11g2': area.style.display = 'none'; break;
				case 'Std10iSO11h': area.style.display = 'none'; break;
				case 'Std10dSO11c': area.style.display = 'none'; break;
				case 'Std10eSO11d': area.style.display = 'none'; break;
				
				case 'Std10fSO11e': area.style.display = 'none'; break;
				case 'Std10jSO11i': area.style.display = 'none'; break;
				case 'MAA810': area.style.display = 'none'; break;
				case 'Std9bSO10b2': area.style.display = 'none'; break;
				case 'Std11SO12': area.style.display = 'none'; break;
				case 'Std11SO122': area.style.display = 'none'; break;
				
			  //PROGRAM STANDARDS
				case 'Std12aSO13a': area.style.display = 'none'; break;
				case 'policyeligible': area.style.display = 'none'; break;
				//case 'proceligible': area.style.display = 'none'; break;
				case 'policychildrec': area.style.display = 'none'; break;
				//case 'procchildrec': area.style.display = 'none'; break;
				case 'policyvolrec': area.style.display = 'none'; break;
				//case 'procvolrec': area.style.display = 'none'; break;
				case 'policyref': area.style.display = 'none'; break;
				
				//case 'procref': area.style.display = 'none'; break;
				case 'policyinq': area.style.display = 'none'; break;
				// 'procinq': area.style.display = 'none'; break;
				case 'policyintake': area.style.display = 'none'; break;
				//case 'procintake': area.style.display = 'none'; break;
				case 'policymatch': area.style.display = 'none'; break;
				
				//case 'procmatch': area.style.display = 'none'; break;
				case 'policysup': area.style.display = 'none'; break;
				//case 'procsup': area.style.display = 'none'; break;
				case 'policyclosure': area.style.display = 'none'; break;
				//case 'procclosure': area.style.display = 'none'; break;
				case 'policyrecords': area.style.display = 'none'; break;
				case 'Std12PPHandlingDoc': area.style.display = 'none'; break;
				
			  //Program Manual addresses risk management issues
				case 'policyovernite': area.style.display = 'none'; break;
				case 'policysexabuse': area.style.display = 'none'; break;
				case 'policystaffasbigs': area.style.display = 'none'; break;
				case 'policyinterothers': area.style.display = 'none'; break;
				case 'policypriorexp': area.style.display = 'none'; break;
				
			  //Financial Management
				case 'Std13SO14': area.style.display = 'none'; break;
				case 'childconsent': area.style.display = 'none'; break;
				case 'childinterview': area.style.display = 'none'; break;
				case 'childparinterview': area.style.display = 'none'; break;
				case 'childhomeassess': area.style.display = 'none'; break;
				case 'volconsent': area.style.display = 'none'; break;
				case 'volreferences': area.style.display = 'none'; break;
				case 'volcriminal': area.style.display = 'none'; break;
				case 'volinterview': area.style.display = 'none'; break;
				case 'volhomeassess': area.style.display = 'none'; break;
				case 'volmatching': area.style.display = 'none'; break;
				case 'voltraining': area.style.display = 'none'; break;
				
			  //Standard 16 - 17
				case 'Approveschild': area.style.display = 'none'; break;
				case 'Approvesparent': area.style.display = 'none'; break;
				case 'Approvesvol': area.style.display = 'none'; break;
				case 'Inpersonmatch': area.style.display = 'none'; break;
				case 'Std17SO18': area.style.display = 'none'; break;
				case 'Std18SO19': area.style.display = 'none'; break;
				case 'Std19SO20': area.style.display = 'none'; break;
				case 'Std20SO21': area.style.display = 'none'; break;
				case 'Std21SO22': area.style.display = 'none'; break;
				case 'Std22SO23': area.style.display = 'none'; break;
			 }
		}
	}
</SCRIPT>

<% '<!--#include file="../includes/top_nav_forms_yearly.inc"--><!-- include file has </head> and <body> tags --><br>      %>
<!--#include file="../includes/surveytitle.inc"-->

<table width=100% cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_fishing.jpg" alt="" width="220" height="477" border="0"></td>
<td valign="top">

<% If say = "thanks" Then %>

<font class="formMain">
<br><br>
<strong>Thank you!</strong> Your information has been saved in the BBBS database.<br><br>
<strong><font color="#FF0000">PLEASE NOTE: Make sure that you complete <em>BOTH</em> the Operational Standards and Program Standards forms.</font><br><br></strong>  
To choose another form, please select the form type from the choices above.<br>

</font>

<br>
<!--#include file="../includes/contact_info.inc"-->
<br>



<% ElseIf say <> "thanks" Then  %>
<form name="frmSelfAssessment" action="SelfAssessment_edit.asp" method="post">
<!--#include file="../includes/form_stamp.asp"-->

<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmSelfAssessment WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
	Set GetSelfAssessment = Con.Execute(query)
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
			<br>
			<table border="1" cellspacing="0" cellpadding="3"  bordercolordark="#003063" width = "650">
				<tr> 
					<td colspan="3" align="center" valign="top" class="formSubhead">BBBS - <%= y %> Agency Self-Assessment</td>
				</tr>
				<tr>
					<td colspan="3" class="formHeader">Agency Self Assessment
					<br>
					<% if section="Operational" then %>
						Business Performance - Operational Standards
					<% else %>
						Program Performance - Program Standards
					<% end if %>
					</td>
				</tr>
				
				<tr>
					<td colspan="7" class="formMain"><font color="#ff0000"><div align="center"><strong>Please Note: </strong>After entering your information, you <strong>must</strong> click on the "Save & Comeback Later" or "Save & Finish" button and wait for the "Thank You" screen or your changes will be lost.</div></font></td>
				</tr>	
				
				<tr>
				<%  if section = "Operational" then %>
					<td colspan="2" class="formHeader"><input type="submit" value="Save & Comeback Later" class="formMainBold" ID="Submit1" NAME="Submit1"></td>
					<td colspan="2" class="formHeader"><input type="button" value="Save & Finish" class="formMainBold" onclick="formvalidation(frmSelfAssessment)"></td>
				<% else %>
					<td colspan="2" class="formHeader"><input type="submit" value="Save & Comeback Later" class="formMainBold" ID="Submit2" NAME="Submit2"></td>
					<td colspan="2" class="formHeader"><input type="button" value="Save & Finish" class="formMainBold" onclick="formvalidationPr(frmSelfAssessment)"></td>
				<% end if %>
				</tr>				
				
				

		<!-- Begin Operational Section -->

		<%  if section = "Operational" then %>
		
		<!-- Prepopulate Program Fields -->
	
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std12aSO13a") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd12aSO13a">
        <input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std12aSO13aReason") %><% Else %><% End If %>" name="Std12aSO13aReason">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policyeligible") %><% Else %>0<% End If %>" name="frmSelfAssessmentpolicyeligible">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policyeligibleReason") %><% Else %><% End If %>" name="policyeligibleReason">
				<!--<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("proceligible") %><% Else %>0<% End If %>" name="frmSelfAssessmentproceligible">-->			
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policychildrec") %><% Else %>0<% End If %>" name="frmSelfAssessmentpolicychildrec">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policychildrecReason") %><% Else %><% End If %>" name="policychildrecReason">			

				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policyref") %><% Else %>0<% End If %>" name="frmSelfAssessmentpolicyref">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policyrefReason") %><% Else %><% End If %>" name="policyrefReason">	
		
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policyinq") %><% Else %>0<% End If %>" name="frmSelfAssessmentpolicyinq">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policyinqReason") %><% Else %><% End If %>" name="policyinqReason">

				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policymatch") %><% Else %>0<% End If %>" name="frmSelfAssessmentpolicymatch">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policymatchReason") %><% Else %><% End If %>" name="policymatchReason">
				<!--<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("procmatch") %><% Else %>0<% End If %>" name="frmSelfAssessmentprocmatch">-->			
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policysup") %><% Else %>0<% End If %>" name="frmSelfAssessmentpolicysup">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policysupReason") %><% Else %><% End If %>" name="policysupReason">						
				<!--<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("procsup") %><% Else %>0<% End If %>" name="frmSelfAssessmentprocsup">-->
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policyclosure") %><% Else %>0<% End If %>" name="frmSelfAssessmentpolicyclosure">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policyclosureReason") %><% Else %><% End If %>" name="policyclosureReason">			
				<!--<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("procclosure") %><% Else %>0<% End If %>" name="frmSelfAssessmentprocclosure">-->			
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policyrecords") %><% Else %>0<% End If %>" name="frmSelfAssessmentpolicyrecords">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policyrecordsReason") %><% Else %><% End If %>" name="policyrecordsReason">			
				<!--<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("procrecords") %><% Else %>0<% End If %>" name="frmSelfAssessmentprocrecords">-->
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policyovernite") %><% Else %>0<% End If %>" name="frmSelfAssessmentpolicyovernite">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policyoverniteReason") %><% Else %><% End If %>" name="policyoverniteReason">		
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policysexabuse") %><% Else %>0<% End If %>" name="frmSelfAssessmentpolicysexabuse">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policysexabuseReason") %><% Else %><% End If %>" name="policysexabuseReason">			
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policystaffasbigs") %><% Else %>0<% End If %>" name="frmSelfAssessmentpolicystaffasbigs">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policystaffasbigsReason") %><% Else %><% End If %>" name="policystaffasbigsReason">			
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policyinterothers") %><% Else %>0<% End If %>" name="frmSelfAssessmentpolicyinterothers">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policyinterothersReason") %><% Else %><% End If %>" name="policyinterothersReason">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policypriorexp") %><% Else %>0<% End If %>" name="frmSelfAssessmentpolicypriorexp">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("policypriorexpReason") %><% Else %><% End If %>" name="policypriorexpReason">			
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std13SO14") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd13SO14">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std13SO14Reason") %><% Else %><% End If %>" name="Std13SO14Reason">			
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("childconsent") %><% Else %>0<% End If %>" name="frmSelfAssessmentchildconsent">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("childconsentReason") %><% Else %><% End If %>" name="childconsentReason">						
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("childinterview") %><% Else %>0<% End If %>" name="frmSelfAssessmentchildinterview">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("childinterviewReason") %><% Else %><% End If %>" name="childinterviewReason">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("childparinterview") %><% Else %>0<% End If %>" name="frmSelfAssessmentchildparinterview">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("childparinterviewReason") %><% Else %><% End If %>" name="childparinterviewReason">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("childhomeassess") %><% Else %>0<% End If %>" name="frmSelfAssessmentchildhomeassess">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("childhomeassessReason") %><% Else %><% End If %>" name="childhomeassessReason">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("volconsent") %><% Else %>0<% End If %>" name="frmSelfAssessmentvolconsent">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("volconsentReason") %><% Else %><% End If %>" name="volconsentReason">			
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("volreferences") %><% Else %>0<% End If %>" name="frmSelfAssessmentvolreferences">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("volreferencesReason") %><% Else %><% End If %>" name="volreferencesReason">			
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("volcriminal") %><% Else %>0<% End If %>" name="frmSelfAssessmentvolcriminal">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("volcriminalReason") %><% Else %><% End If %>" name="volcriminalReason">
				<!--<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("volinterview") %><% Else %>0<% End If %>" name="frmSelfAssessmentvolinterview">-->
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("volinterviewReason") %><% Else %><% End If %>" name="volinterviewReason">			
				<!--<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("volhomeassess") %><% Else %>0<% End If %>" name="frmSelfAssessmentvolhomeassess">-->
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("volhomeassessReason") %><% Else %><% End If %>" name="volhomeassessReason">			
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("volmatching") %><% Else %>0<% End If %>" name="frmSelfAssessmentvolmatching">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("volmatchingReason") %><% Else %><% End If %>" name="volmatchingReason">			
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("voltraining") %><% Else %>0<% End If %>" name="frmSelfAssessmentvoltraining">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("voltrainingReason") %><% Else %><% End If %>" name="voltrainingReason">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Approveschild") %><% Else %>0<% End If %>" name="frmSelfAssessmentApproveschild">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("ApproveschildReason") %><% Else %><% End If %>" name="ApproveschildReason">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Approvesparent") %><% Else %>0<% End If %>" name="frmSelfAssessmentApprovesparent">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("ApprovesparentReason") %><% Else %><% End If %>" name="ApprovesparentReason">			
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Approvesvol") %><% Else %>0<% End If %>" name="frmSelfAssessmentApprovesvol">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("ApprovesvolReason") %><% Else %><% End If %>" name="ApprovesvolReason">						
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Inpersonmatch") %><% Else %>0<% End If %>" name="frmSelfAssessmentInpersonmatch">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("InpersonmatchReason") %><% Else %><% End If %>" name="InpersonmatchReason">			
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std17SO18") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd17SO18">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std17SO18Reason") %><% Else %><% End If %>" name="Std17SO18Reason">			
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std18SO19") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd18SO19">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std18SO19Reason") %><% Else %><% End If %>" name="Std18SO19Reason">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std19SO20") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd19SO20">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std19SO20Reason") %><% Else %><% End If %>" name="Std19SO20Reason">			
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std20SO21") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd20SO21">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std20SO21Reason") %><% Else %><% End If %>" name="Std20SO21Reason">			
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std21SO22") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd21SO22">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std21SO22Reason") %><% Else %><% End If %>" name="Std21SO22Reason">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std22SO23") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd22SO23">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std22SO23Reason") %><% Else %><% End If %>" name="Std22SO23Reason">			
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("PolicyVolrec") %><% Else %>0<% End If %>" name="frmSelfAssessmentPolicyVolrec">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("PolicyVolrecReason") %><% Else %><% End If %>" name="PolicyVolrecReason">							
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("PolicyIntake") %><% Else %>0<% End If %>" name="frmSelfAssessmentPolicyIntake">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("PolicyIntakeReason") %><% Else %><% End If %>" name="PolicyIntakeReason">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std12PPHandlingDoc") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd12PPHandlingDoc">
				<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std12PPHandlingDocReason") %><% Else %><% End If %>" name="Std12PPHandlingDocReason">		
				
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" width=40%>Standard 1: The affiliate operates in compliance with applicable laws</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" width=40%>Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>
				
				<!-- Articles of Incorporation -->
				<tr>
					<td align="left" valign="top" class="formMain">Articles of Incorporation</td>
					<td align="left" valign="top" class="formMain">Review Articles of Incorporation; check for approved agency name</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd1a" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std1a")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std1a',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd1a" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std1a")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std1a',true)">Out
						<br>
					
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd1a" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std1a"))) or Trim(GetSelfAssessment("Std1a")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std1a',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd1a" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std1a',false)">Not Entered						
						<% end if %>							
					</td>
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std1a" style="display:none;">
									 <label for="Std1aReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.) (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std1aReason" id="Std1aReason" colspan="3">
							</div>
					</td>
				</tr>

				<!-- Tax-exempt status documentation / IRS Letter -->
				<tr>
					<td align="left" valign="top" class="formMain">Tax-exempt Status Documentation / IRS Letter</td>
					<td align="left" valign="top" class="formMain">Review tax exempt status documents; check for current agency name</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd1b" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std1b")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std1b',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd1b" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std1b")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std1b',true)">Out
						<br>						
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd1b" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std1b"))) or Trim(GetSelfAssessment("Std1b")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std1b',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd1b" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std1b',false)">Not Entered						
						<% end if %>							
					</td>
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std1b" style="display:none;">
									 <label for="Std1bReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.) (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std1bReason" colspan="3">
							</div>
					</td>
				</tr>			
				
				<!-- 990 form -->
				<tr>
					<td align="left" valign="top" class="formMain">990 Form</td>
					<td align="left" valign="top" class="formMain"> 990 has been filed with IRS for prior fiscal year</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentForm990" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Form990")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Form990',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentForm990" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Form990")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Form990',true)">Out
						<br>						
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentForm990" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Form990"))) or Trim(GetSelfAssessment("Form990")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Form990',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentForm990" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Form990',false)">Not Entered						
						<% end if %>							
					</td>
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Form990" style="display:none;">
									 <label for="Form990Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.) (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Form990Reason" colspan="3">
							</div>
					</td>
				</tr>	
					
				
				<!-- Corporate Minutes -->
				<tr>
					<td align="left" valign="top" class="formMain">Corporate Minutes</td>
					<td align="left" valign="top" class="formMain">Board meeting minutes are on file and signed</td>

					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd1c" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std1c")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std1c',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd1c" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std1c")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std1c',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd1c" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std1c"))) or Trim(GetSelfAssessment("Std1c")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std1c',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd1c" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std1c',false)">Not Entered						
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std1c" style="display:none;">
									 <label for="Std1cReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.) (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std1cReason" colspan="3">
							</div>
					</td>
				</tr>	
				

				
				<!-- Corporate Bylaws -->
				<tr>
					<td align="left" valign="top" class="formMain">Corporate Bylaws</td>
					<td align="left" valign="top" class="formMain">Current copy of Bylaws are on file</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentBylaws" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Bylaws")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Bylaws',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentBylaws" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Bylaws")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Bylaws',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentBylaws" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Bylaws"))) or Trim(GetSelfAssessment("Bylaws")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Bylaws',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentBylaws" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Bylaws',false)">Not Entered						
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Bylaws" style="display:none;">
									 <label for="BylawsReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.) (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="BylawsReason" colspan="3">
							</div>
					</td>
				</tr>					
				
				<!-- Executed Affiliation Agreement -->
				<tr>
					<td align="left" valign="top" class="formMain">Executed Membership Affiliation Agreement (MAA)</td>
					<td align="left" valign="top" class="formMain">Signed MAA is on file and reflects current Service Community Area (SCA)</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentMAA" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("MAA")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentMAA" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("MAA")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentMAA" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("MAA"))) or Trim(GetSelfAssessment("MAA")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentMAA" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'MAA',false)">Not Entered						
						<% end if %>							
					</td>							
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="MAA" style="display:none;">
									 <label for="MAAReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.) (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="MAAReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Logo and name -->
				<tr>
					<td align="left" valign="top" class="formMain">Affiliate uses, exclusively, the logo adopted by BBBSA and operates under a name approved by BBBSA</td>
					<td align="left" valign="top" class="formMain">Signage, stationery, business cards, publications, other materials  should all reflect consistent use of BBBSA logo and approved agency names</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd1LogoAndName" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std1LogoAndName")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std1LogoAndName',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd1LogoAndName" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std1LogoAndName")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std1LogoAndName',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd1LogoAndName" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std1LogoAndName"))) or Trim(GetSelfAssessment("Std1LogoAndName")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std1LogoAndName',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd1LogoAndName" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std1LogoAndName',false)">Not Entered						
						<% end if %>							
					</td>							
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std1LogoAndName" style="display:none;">
									 <label for="Std1LogoAndNameReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.) (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std1LogoAndNameReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<tr>
					<td colspan="3" class="formHeaderMedium" align="center">Board Development</td>
				</tr>
				
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Standard 2/Standard 2,3 (sponsoring organization): The affiliate has a board recruitment and development system that focuses on providing effective and diverse representation, and provides training and leadership development to ensure that board members have the knowledge, skills, and tools necessary to effectively perform their responsibilities</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>
				
				<!-- Written Board Development Plan -->
				<tr>
					<td align="left" valign="top" class="formMain">Written Board Development Plan</td>
					<td align="left" valign="top" class="formMain">Board-approved, stand-alone document that includes: job descriptions; gap assessment and recruitment plan; board commitment specifications; orientation plan; and annual review process.</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd2" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std2")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std2',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd2" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std2")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std2',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd2" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std2"))) or Trim(GetSelfAssessment("Std2")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std2',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd2" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std2',false)">Not Entered						
						<% end if %>							
					</td>												
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std2" style="display:none;">
									 <label for="Std2Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.) (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std2Reason" colspan="3">
							</div>
					</td>
				</tr>	
								
				<!-- Board Recruitment Plan --
				<tr>
					<td align="left" valign="top" class="formMain">Board Recruitment Plan</td>
					<td align="left" valign="top" class="formMain">Review job descriptions, gap assessment, written recruitment plan, and board orientation</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd2aSO3a" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std2aSO3a")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std2aSO3a',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd2aSO3a" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std2aSO3a")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std2aSO3a',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd2aSO3a" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std2aSO3a"))) or Trim(GetSelfAssessment("Std2aSO3a")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std2aSO3a',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd2aSO3a" value="0" checked>Not Entered						
						<% end if %>							
					</td>												
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std2aSO3a" style="display:none;">
									 <label for="Std2aSO3aReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.) (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std2aSO3aReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Board Training Plan --
				<tr>
					<td align="left" valign="top" class="formMain">Board Training Plan</td>
					<td align="left" valign="top" class="formMain"><a target="_blank" href="http://agencyconnection.bbbs.org/atf/cf/{4CA344D5-890B-48AA-A80B-3EE1364E3AB7}/Self-Assessment%20Suggested%20Ideas.doc">Click Here</a> for recommendations<br>(MS Word Format)</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentBrdtrainplan" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Brdtrainplan")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Brdtrainplan',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentBrdtrainplan" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Brdtrainplan")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Brdtrainplan',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentBrdtrainplan" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Brdtrainplan"))) or Trim(GetSelfAssessment("Brdtrainplan")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Brdtrainplan',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentBrdtrainplan" value="0" checked>Not Entered						
						<% end if %>							
					</td>
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Brdtrainplan" style="display:none;">
									 <label for="BrdtrainplanReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="BrdtrainplanReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Documentation of annual review of board's performance -->
				<tr>
					<td align="left" valign="top" class="formMain">Documentation of annual review of board's performance</td>
					<td align="left" valign="top" class="formMain">Date and documentation that a review was conducted</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd2bSO3b" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std2bSO3b")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std2bSO3b',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd2bSO3b" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std2bSO3b")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std2bSO3b',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd2bSO3b" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std2bSO3b"))) or Trim(GetSelfAssessment("Std2bSO3b")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std2bSO3b',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd2bSO3b" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std2bSO3b',false)">Not Entered						
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std2bSO3b" style="display:none;">
									 <label for="Std2bSO3bReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std2bSO3bReason" colspan="3">
							</div>
					</td>
				</tr>

				<!-- Documentation that board/advisory group representative attends annual national conference -->
				<tr>
					<td align="left" valign="top" class="formMain">Documentation that board/advisory group representative (s) attend  national conference, regional conferences/meetings, workshops, and/ or trainings</td>
					<td align="left" valign="top" class="formMain">Date and documentation of attendance and type of meeting</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentMAA810conf" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("MAA810conf")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA810conf',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentMAA810conf" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("MAA810conf")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA810conf',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentMAA810conf" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("MAA810conf"))) or Trim(GetSelfAssessment("MAA810conf")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA810conf',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentMAA810conf" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'MAA810conf',false)">Not Entered						
						<% end if %>							
					</td>
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="MAA810conf" style="display:none;">
									 <label for="MAA810confReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="MAA810confReason" colspan="3">
							</div>
					</td>
				</tr>

				<!-- SO (Sponsoring Organization):  Written agreement between Corporate Board and Advisory Group re: voting representation and selection policy -->				
				<tr>
					<td align="left" valign="top" class="formMain"><strong><em>Sponsored Only:</em></strong>  Written agreement between Corporate Board and Advisory Group re: voting representation and selection policy</td>
					<td align="left" valign="top" class="formMain">Review bylaws, board minutes and governing board roster</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd2SO" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std2SO")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std2SO',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd2SO" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std2SO")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std2SO',true)">Out
						<br>
						<input type="radio" name="frmSelfAssessmentStd2SO" value="3"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std2SO")) = "3" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std2SO',false)">N/A
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd2SO" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std2SO"))) or Trim(GetSelfAssessment("Std2SO")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std2SO',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd2SO" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std2SO',false)">Not Entered						
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std2SO" style="display:none;">
									 <label for="Std2SOReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std2SOReason" colspan="3">
							</div>
					</td>
				</tr>

				<!-- Mission/Vision -->
				<tr>
					<td colspan="3" class="formHeaderMedium" align="center">Mission/Vision</td>
				</tr>
				
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Standard 3/Standard 4 (sponsored programs):  The affiliate has a clearly defined and articulated vision and mission statement that drives all agency decision making and provides focus for the assessment of the affiliate's work.</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>
				
				<!-- Written, board-approved mission statement -->
				<tr>
					<td align="left" valign="top" class="formMain">Written, board-approved mission statement</td>
					<td align="left" valign="top" class="formMain">Review mission statement to ensure that it is compatible with that of BBBSA and, as written, is used to drive decision-making</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd3SO4m" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std3SO4m")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std3SO4m',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd3SO4m" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std3SO4m")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std3SO4m',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd3SO4m" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std3SO4m"))) or Trim(GetSelfAssessment("Std3SO4m")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std3SO4m',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd3SO4m" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std3SO4m',false)">Not Entered						
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std3SO4m" style="display:none;">
									 <label for="Std3SO4mReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std3SO4mReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Written, board-approved vision statement -->
				<tr>
					<td align="left" valign="top" class="formMain">Written, board-approved vision statement</td>
					<td align="left" valign="top" class="formMain">Review vision statement to ensure that it is compatible with that of BBBSA and, as written, is used to drive decision-making</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd3SO4v" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std3SO4v")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std3SO4v',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd3SO4v" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std3SO4v")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std3SO4v',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd3SO4v" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std3SO4v"))) or Trim(GetSelfAssessment("Std3SO4v")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std3SO4v',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd3SO4v" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std3SO4v',false)">Not Entered						
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std3SO4v" style="display:none;">
									 <label for="Std3SO4vReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std3SO4vReason" colspan="3">
							</div>
					</td>
				</tr>
				
				
				<!-- Strategic Planning -->
				<tr>
					<td colspan="3" class="formHeaderMedium" align="center">Strategic Planning</td>
				</tr>
				
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Standard 4/Standard 5 (sponsored programs): The affiliate has a comprehensive strategic planning process which addresses all aspects of the affiliate's operations including, but not limited to, growth plans for One-To-One service as well as other services to children in need; marketing; technology; and facility needs</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>
				
				
				<!-- Strategic Plan in alignment with nationwide strategic plan -->
				<tr>
					<td align="left" valign="top" class="formMain">Written board-approved current Strategic Plan</td>
					<td align="left" valign="top" class="formMain">Board-approved, stand-alone document that addresses, at a minimum, services to children, marketing, technology, facilities and procedure on using the plan to drive decision-making.</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd4SO5" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std4SO5")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std4SO5',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd4SO5" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std4SO5")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std4SO5',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd4SO5" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std4SO5"))) or Trim(GetSelfAssessment("Std4SO5")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std4SO5',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd4SO5" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std4SO5',false)">Not Entered						
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std4SO5" style="display:none;">
									 <label for="Std4SO5Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std4SO5Reason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Annual Operating Plan -->
				<!-- 
				<tr>
					<td align="left" valign="top" class="formMain">Annual Operating Plan</td>
					<td align="left" valign="top" class="formMain">This is the annual plan identifying goals for the year and objectives to meeting those goals</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentOpplan" value="2"<% 'If say = "edit" Then %><% 'If Trim(GetSelfAssessment("Opplan")) = "2" Then %> checked<% 'End If %><% 'End If %>>In
						<br>
						<input type="radio" name="frmSelfAssessmentOpplan" value="1"<% 'If say = "edit" Then %><% 'If Trim(GetSelfAssessment("Opplan")) = "1" Then %> checked<% 'End If %><% 'End If %>>Out
						<br>
						<% 'if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentOpplan" value="0"<% 'If say = "edit" Then %><% 'If isnull(Trim(GetSelfAssessment("Opplan"))) or Trim(GetSelfAssessment("Opplan")) = "0" Then %> checked<% 'End If %><% 'End If %>>Not Entered
						<% 'else %>
							<input type="radio" name="frmSelfAssessmentOpplan" value="0" checked>Not Entered						
						<% 'end if %>							
					</td>	
				</tr>		
				
				<!-- Quality Assurance -->
				<tr>
					<td colspan="3" class="formHeaderMedium" align="center">Quality Assurance</td>
				</tr>								
				
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Standard 5/Standard 6 (sponsored programs):  The affiliate has a quality assurance system that ensures that all aspects of the affiliate's operations are reviewed and assessed on an annual basis, to include a review of its policies and procedures to ensure compliance with Standards of Practice for One-To-One Service related to program management for affiliates, and ensures that the affiliate is in compliance with its own program manual.</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>
				
				<!-- Documentation of annual review of all corporate policies and procedures -->
				<tr>
					<td align="left" valign="top" class="formMain">Documentation of annual review of all corporate policies and procedures</td>
					<!--<td align="left" valign="top" class="formMain"><a target="_blank" href="http://agencyconnection.bbbs.org/atf/cf/{4CA344D5-890B-48AA-A80B-3EE1364E3AB7}/Self-Assessment%20Suggested%20Ideas.doc">Click Here</a> for recommendations<br>(MS Word Format)</td>-->
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd5opsSO6" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std5opsSO6")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std5opsSO6',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd5opsSO6" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std5opsSO6")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std5opsSO6',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd5opsSO6" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std5opsSO6"))) or Trim(GetSelfAssessment("Std5opsSO6")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std5opsSO6',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd5opsSO6" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std5opsSO6',false)">Not Entered						
						<% end if %>							
					</td>
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std5opsSO6" style="display:none;">
									 <label for="Std5opsSO6Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std5opsSO6Reason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Documentation of annual review of Program Manual -->				
				<tr>
					<td align="left" valign="top" class="formMain">Documentation of annual review of Program Manual</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd5pgmSO6" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std5pgmSO6")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std5pgmSO6',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd5pgmSO6" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std5pgmSO6")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std5pgmSO6',true)">Out
						<br>

						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd5pgmSO6" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std5pgmSO6"))) or Trim(GetSelfAssessment("Std5pgmSO6")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std5pgmSO6',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd5pgmSO6" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std5pgmSO6',false)">Not Entered						
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std5pgmSO6" style="display:none;">
									 <label for="Std5pgmSO6Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std5pgmSO6Reason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Documentation of annual, random case file(s) audit -->
				<tr>
					<td align="left" valign="top" class="formMain">Documentation of annual, random case file(s) audit to ensure compliance with program standards </td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd5filesSO6" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std5filesSO6")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std5filesSO6',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd5filesSO6" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std5filesSO6")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std5filesSO6',true)">Out
						<br>

						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd5filesSO6" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std5filesSO6"))) or Trim(GetSelfAssessment("Std5filesSO6")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std5filesSO6',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd5filesSO6" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std5filesSO6',false)">Not Entered						
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std5filesSO6" style="display:none;">
									 <label for="Std5filesSO6Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std5filesSO6Reason" colspan="3">
							</div>
					</td>
				</tr>

				<!-- Fund Development -->
				<tr>
					<td colspan="3" class="formHeaderMedium" align="center">Fund Development</td>
				</tr>		
				
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Standard 6/Standard 7 (sponsoring organization):  The affiliate has a financial management and fund development plan that ensures that fund development efforts are substantial enough to address current operation needs, contingencies, and planned growth.</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>	
				
				<!-- Documentation of board-approved annual budget -->
				<tr>
					<td align="left" valign="top" class="formMain">Documentation of annual review and board-approval of annual budget</td>
					<td align="left" valign="top" class="formMain">Documentation in board minutes that annual budget has been approved. </td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd6SO7budget" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std6SO7budget")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std6SO7budget',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd6SO7budget" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std6SO7budget")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std6SO7budget',true)">Out
						<br>

						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd6SO7budget" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std6SO7budget"))) or Trim(GetSelfAssessment("Std6SO7budget")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std6SO7budget',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd6SO7budget" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std6SO7budget',false)">Not Entered						
						<% end if %>							
					</td>						
				</tr>	
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std6SO7budget" style="display:none;">
									 <label for="Std6SO7budgetReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std6SO7budgetReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Budget includes expenses for training and travel to conferences -->
				<!--
				<tr>
					<td align="left" valign="top" class="formMain">Budget includes expenses for training and travel to conferences</td>
					<td align="left" valign="top" class="formMain">Identify the line in the budget for professional development / conferences</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentMAA810exp" value="2"<% 'If say = "edit" Then %><% 'If Trim(GetSelfAssessment("MAA810exp")) = "2" Then %> checked<% 'End If %><% 'End If %>>In
						<br>
						<input type="radio" name="frmSelfAssessmentMAA810exp" value="1"<%' If say = "edit" Then %><% 'If Trim(GetSelfAssessment("MAA810exp")) = "1" Then %> checked<% 'End If %><% 'End If %>>Out
						<br>
						<% 'if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentMAA810exp" value="0"<% 'If say = "edit" Then %><% 'If isnull(Trim(GetSelfAssessment("MAA810exp"))) or Trim(GetSelfAssessment("MAA810exp")) = "0" Then %> checked<% 'End If %><% 'End If %>>Not Entered
						<% 'else %>
							<input type="radio" name="frmSelfAssessmentMAA810exp" value="0" checked>Not Entered						
						<% 'end if %>							
					</td>					
				</tr>		
				-->		
				
				<!-- Proof affiliate restricts its fund-raising activities to its own Service Community Area (SCA) or has written agreement with neighboring BBBSA affiliate -->				
				<tr>
					<td align="left" valign="top" class="formMain">Proof affiliate restricts its fund-raising activities to its own Service Community Area (SCA) or has written agreement with neighboring BBBSA affiliate</td>
					<td align="left" valign="top" class="formMain">If any fundraising activity is held in another BBBS' service community area,  your written agreement with that BBBS agency, authorizing your fundraising activity must be on file</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentMAA32" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("MAA32")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA32',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentMAA32" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("MAA32")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA32',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentMAA32" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("MAA32"))) or Trim(GetSelfAssessment("MAA32")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA32',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentMAA32" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'MAA32',false)">Not Entered						
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="MAA32" style="display:none;">
									 <label for="MAA32Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="MAA32Reason" colspan="3">
							</div>
					</td>
				</tr>					
				
				<!-- Fund Development Plan -->				
				<tr>
					<td align="left" valign="top" class="formMain">Written board-approved Fund Development Plan (NOT NEW - WAS MOVED FROM FINANCIAL MANAGEMENT)</td>
					<td align="left" valign="top" class="formMain">Review Written board-approved fundraising plan, including annual goals for diversification of funding, and planned revenue growth</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd6SO7b" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std6SO7b")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std6SO7b',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd6SO7b" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std6SO7b")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std6SO7b',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd6SO7b" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std6SO7b"))) or Trim(GetSelfAssessment("Std6SO7b")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std6SO7b',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd6SO7b" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std6SO7b',false)">Not Entered						
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std6SO7b" style="display:none;">
									 <label for="Std6SO7bReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std6SO7bReason" colspan="3">
							</div>
					</td>
				</tr>	
				
				<!-- Financial Management -->
				<tr>
					<td colspan="3" class="formHeaderMedium" align="center">Financial Management</td>
				</tr>			
											
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Standard 7/Standard 8 (sponsoring organization):  The affiliate has established financial management practices that meet generally accepted accounting practices and has an oversight structure that facilitates the early identification of potential problems</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>	
				
				<!-- Written, board-approved Fund Development Plan --
				<tr>
					<td align="left" valign="top" class="formMain">Board oversight consistent with Generally Accepted Accounting Practices (GAAP)</td>
					<td align="left" valign="top" class="formMain">Documentation in board minutes that board reviews on a regular basis the agency's financials, including balance sheet; profit and loss statement; cash flow projections and budget variance report</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd6SO7b" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std6SO7b")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std6SO7b',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd6SO7b" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std6SO7b")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std6SO7b',true)">Out
						<br>

						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd6SO7b" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std6SO7b"))) or Trim(GetSelfAssessment("Std6SO7b")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std6SO7b',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd6SO7b" value="0" checked>Not Entered						
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std6SO7b" style="display:none;">
									 <label for="Std6SO7bReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std6SO7bReason" colspan="3">
							</div>
					</td>
				</tr> -->
				
				<!-- Board oversight consistent with General Accounting Practices -->
				<tr>
					<td align="left" valign="top" class="formMain">Board oversight consistent with Generally Accepted Accounting Practices (GAAP)</td>
					<td align="left" valign="top" class="formMain">Documentation in board minutes that board reviews on a regular basis the agency's financials, including balance sheet; profit and loss statement; cash flow projections and budget variance report</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd6SO7" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std6SO7")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std6SO7',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd6SO7" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std6SO7")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std6SO7',true)">Out
						<br>

						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd6SO7" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std6SO7"))) or Trim(GetSelfAssessment("Std6SO7")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std6SO7',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd6SO7" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std6SO7',false)">Not Entered						
						<% end if %>							
					</td>
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std6SO7" style="display:none;">
									 <label for="Std6SO7Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std6SO7Reason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Written financial management practices -->
				<tr>
					<td align="left" valign="top" class="formMain">Written board-approved financial management practices</td>
					<td align="left" valign="top" class="formMain">Documentation on file of the agency's financial management practices which should include, at a minimum, managing of deposits, check writing, authorization of expenditures, managing of donations, petty cash, etc. </td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd7SO8" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std7SO8")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std7SO8',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd7SO8" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std7SO8")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std7SO8',true)">Out
						<br>

						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd7SO8" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std7SO8"))) or Trim(GetSelfAssessment("Std7SO8")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std7SO8',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd7SO8" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std7SO8',false)">Not Entered						
						<% end if %>							
					</td>
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std7SO8" style="display:none;">
									 <label for="Std7SO8Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std7SO8Reason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Annual financial audit from the last fiService Community Area (SCA)l year -->
				<tr>
					<td align="left" valign="top" class="formMain">Annual audit of its financial condition, certified by an independent, certified public accounting firm and in accordance with generally accepted accounting principles (GAAP).</td>
					<td align="left" valign="top" class="formMain">Send copy of your most recently completed financial audit to the National Office.</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentMAA88" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("MAA88")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA88',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentMAA88" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("MAA88")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA88',true)">Out
						<br>

						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentMAA88" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("MAA88"))) or Trim(GetSelfAssessment("MAA88")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA88',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentMAA88" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'MAA88',false)">Not Entered						
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="MAA88" style="display:none;">
									 <label for="MAA88Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="MAA88Reason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Affiliate is current with membership fees -->
				<tr>
					<td align="left" valign="top" class="formMain">Affiliate is current with membership fees</td>
					<td align="left" valign="top" class="formMain">Be able to track date and amount of payment of annual BBBSA fees, including any negotiated payment plans. Fee calculation forms are submitted on-line If agency affiliation fees are more than 6 months delinquent, a payment schedule must be approved and consistently followed.</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentMAA82" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("MAA82")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA82',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentMAA82" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("MAA82")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA82',true)">Out
						<br>

						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentMAA82" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("MAA82"))) or Trim(GetSelfAssessment("MAA82")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA82',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentMAA82" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'MAA82',false)">Not Entered						
						<% end if %>							
					</td>										
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="MAA82" style="display:none;">
									 <label for="MAA82Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="MAA82Reason" colspan="3">
							</div>
					</td>
				</tr>	
							
				<!-- SO (Sponsoring Organization):  Documentation that funds raised or allocated to BBBSA program used solely for BBBSA expenses -->							
				<tr>
					<td align="left" valign="top" class="formMain"><strong><em>Sponsored Only:</em></strong>  Documentation that funds raised or allocated to BBBS program are used solely for BBBS expenses and that any share of administrative costs charged to BBBS program is reasonable.</td>
					<td align="left" valign="top" class="formMain">Sponsoring Organization's annual audit must include income and expenses of the BBBS program and indicate if BBBS funds are held in a separate account or if segregated accounting is used.</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStdSO8a" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("StdSO8a")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'StdSO8a',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStdSO8a" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("StdSO8a")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'StdSO8a',true)">Out
						<br>
						<input type="radio" name="frmSelfAssessmentStdSO8a" value="3"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("StdSO8a")) = "3" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'StdSO8a',false)">N/A
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStdSO8a" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("StdSO8a"))) or Trim(GetSelfAssessment("StdSO8a")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'StdSO8a',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStdSO8a" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'StdSO8a',false)">Not Entered						
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="StdSO8a" style="display:none;">
									 <label for="StdSO8aReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="StdSO8aReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- SO (Sponsoring Organization):  Documentation that administrative costs charged to BBBSA program are reasonable, consistant and accurate --
				<tr>
					<td align="left" valign="top" class="formMain"><strong><em>Sponsored Only:</em></strong>  Documentation that administrative costs charged to BBBSA program are reasonable, consistant and accurate</td>
					<td align="left" valign="top" class="formMain">Assess the percent charged to the BBBS budget for administrative costs and the methodology used</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStdSO8b" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("StdSO8b")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'StdSO8b',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStdSO8b" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("StdSO8b")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'StdSO8b',true)">Out
						<br>
						<input type="radio" name="frmSelfAssessmentStdSO8b" value="3"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("StdSO8b")) = "3" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'StdSO8b',false)">N/A
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStdSO8b" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("StdSO8b"))) or Trim(GetSelfAssessment("StdSO8b")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'StdSO8b',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStdSO8b" value="0" checked>Not Entered						
						<% end if %>							
					</td>					
				</tr>					
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="StdSO8b" style="display:none;">
									 <label for="StdSO8bReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="StdSO8bReason" colspan="3">
							</div>
					</td>
				</tr> -->

				<!-- SO (Sponsoring Organization):  Proof that income and expense reports are provided the Advisory Group at least quarterly -->																	
				<tr>
					<td align="left" valign="top" class="formMain"><strong><em>Sponsored Only:</em></strong>  Proof that income and expense reports are provided to the Advisory Group at least quarterly</td>
					<td align="left" valign="top" class="formMain">Documentation in Advisory Group minutes that BBBS revenue and expense reports are given to and reviewed by Advisory Group.</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStdSO8c" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("StdSO8c")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'StdSO8c',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStdSO8c" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("StdSO8c")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'StdSO8c',true)">Out
						<br>
						<input type="radio" name="frmSelfAssessmentStdSO8c" value="3"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("StdSO8c")) = "3" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'StdSO8c',false)">N/A
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStdSO8c" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("StdSO8c"))) or Trim(GetSelfAssessment("StdSO8c")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'StdSO8c',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStdSO8c" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'StdSO8c',false)">Not Entered						
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="StdSO8c" style="display:none;">
									 <label for="StdSO8cReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="StdSO8cReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!--Risk Management -->
				<tr>
					<td colspan="3" class="formHeaderMedium" align="center">Risk Management</td>
				</tr>			
											
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Standard 8/Standard 9 (sponsoring organization):  The affiliate has a risk management system that ensures that agency operational risks are identified and appropriately managed through insurance, and policies and procedures</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>	
				
				<!-- Written Crisis Management Plan --		
				<tr>
					<td align="left" valign="top" class="formMain">Written Crisis Management Plan</td>
					<td align="left" valign="top" class="formMain">Review current Crisis Management Guide.  Get more information <a href="http://agencyconnection.bbbs.org/atf/cf/{4CA344D5-890B-48AA-A80B-3EE1364E3AB7}/Risk%20Mgm%20&%20Crisis%20Preparedness.doc" target="_blank">here</a>.</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd8SO9crisis" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std8SO9crisis")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std8SO9crisis',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd8SO9crisis" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std8SO9crisis")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std8SO9crisis',true)">Out
						<br>

						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd8SO9crisis" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std8SO9crisis"))) or Trim(GetSelfAssessment("Std8SO9crisis")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std8SO9crisis',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd8SO9crisis" value="0" checked>Not Entered						
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std8SO9crisis" style="display:none;">
									 <label for="Std8SO9crisisReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std8SO9crisisReason" colspan="3">
							</div>
					</td>
				</tr> -->
				
				<!-- Written Risk Management Plan -->				
				<tr>
					<td align="left" valign="top" class="formMain">Written board-approved Risk Management Plan</td>
					<td align="left" valign="top" class="formMain">Review current Risk Management Plan to ensure that it contains the following components: governance, human resources, child safety & youth protection, financial management, fundraising and public relations, facility safety & security, technology and information management, insurance, transportation, and crisis management, including BBBSA's protocol for child abuse reporting. </td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd8SO9risk" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std8SO9risk")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std8SO9risk',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd8SO9risk" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std8SO9risk")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std8SO9risk',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd8SO9risk" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std8SO9risk"))) or Trim(GetSelfAssessment("Std8SO9risk")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std8SO9risk',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd8SO9risk" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std8SO9risk',false)">Not Entered						
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std8SO9risk" style="display:none;">
									 <label for="Std8SO9riskReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std8SO9riskReason" colspan="3">
							</div>
					</td>
				</tr>	
							
				<!-- Proof of adequate insurance coverage -->
				<tr>
					<td align="left" valign="top" class="formMain">Proof of adequate insurance coverage that meets minimums established by BBBSA</td>
					<td align="left" valign="top" class="formMain">Check cover sheet of policies to assess levels of coverage for liability insurance that satisfies the risk management issues associated with the Standards of Practice. Insurance should cover, at a minimum, errors and omissions, bodily injury, property loss, sexual abuse and Director's and Officer's liability. </td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentMAA9" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("MAA9")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA9',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentMAA9" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("MAA9")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA9',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentMAA9" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("MAA9"))) or Trim(GetSelfAssessment("MAA9")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA9',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentMAA9" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'MAA9',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="MAA9" style="display:none;">
									 <label for="MAA9Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="MAA9Reason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Personnel -->
				<tr>
					<td colspan="3" class="formHeaderMedium" align="center">Personnel</td>
				</tr>			
											
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Standard 9/Standard 10 (sponsoring organization):  The affiliate employs a full time executive who is responsible to the board for the overall administration of agency operations</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>															
				
				<!-- Board approved job description for Executive (Program Director for sponsored programs) -->				
				<tr>
					<td align="left" valign="top" class="formMain">Written board-approved job description for Chief Executive (Program Director for Sponsored Programs) that specifies overall responsibility for employing, supervising, evaluating and terminating all paid and volunteer staff</td>
					<td align="left" valign="top" class="formMain">Current Job description for Chief Executive (Program Director for sponsored programs) should be kept in personnel file and referenced in personnel policies.  </td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd10bSO11b2" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10bSO11b2")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10bSO11b2',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd10bSO11b2" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10bSO11b2")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10bSO11b2',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd10bSO11b2" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std10bSO11b2"))) or Trim(GetSelfAssessment("Std10bSO11b2")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10bSO11b2',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd10bSO11b2" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std10bSO11b2',false)">Not Entered
						<% end if %>							
					</td>
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std10bSO11b2" style="display:none;">
									 <label for="Std10bSO11b2Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std10bSO11b2Reason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Documentation of annual performance evaluation of Executive (Program Director for sponsored programs) --		
				<tr>
					<td align="left" valign="top" class="formMain">Documentation of annual performance evaluation of Executive (Program Director for sponsored programs)</td>
					<td align="left" valign="top" class="formMain">Copy of annual performance evaluation is on-file in executive's personnel file</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd9a" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std9a")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std9a',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd9a" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std9a")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std9a',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd9a" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std9a"))) or Trim(GetSelfAssessment("Std9a")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std9a',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd9a" value="0" checked>Not Entered
						<% end if %>							
					</td>
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std9a" style="display:none;">
									 <label for="Std9aReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std9aReason" colspan="3">
							</div>
					</td>
				</tr> -->
				
				<!---time; for Sponsored Organizations, the affiliate employs a BBBS Program Director responsible for overall administration of BBBS Program operations -->
				<tr>
					<td align="left" valign="top" class="formMain">BBBS Chief Executive is employed full-time; and, for Sponsoring Organizations, a full-time Program Director is employed and responsible for overall administration of BBBS Program operations</td>
					<td align="left" valign="top" class="formMain">Confirmation on file that includes:  Letter of Hire and/or time sheets/payroll</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd9SO10" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std9SO10")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std9SO10',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd9SO10" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std9SO10")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std9SO10',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd9SO10" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std9SO10"))) or Trim(GetSelfAssessment("Std9SO10")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std9SO10',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd9SO10" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std9SO10',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std9SO10" style="display:none;">
									 <label for="Std9SO10Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std9SO10Reason" colspan="3">
							</div>
					</td>
				</tr>								
				
				<!-- Definition of who notifies BBBSA of a vacancy in executive position (Program Director for sponsored progrems) -->
				<tr>
					<td align="left" valign="top" class="formMain">Annual performance evaluation of Chief Executive (Program Director for sponsored programs) is conducted in accordance with agency personnel polices, approved job description and annual performance goals</td>
					<td align="left" valign="top" class="formMain">Copy of annual performance evaluation is on-file in Chief Executive's personnel file</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentMAA813" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("MAA813")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA813',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentMAA813" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("MAA813")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA813',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentMAA813" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("MAA813"))) or Trim(GetSelfAssessment("MAA813")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA813',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentMAA813" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'MAA813',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="MAA813" style="display:none;">
									 <label for="MAA813Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="MAA813Reason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Policies specify that executive (Program Director for sponsored programs) has “overall responsibility” for employing, supervising, evaluating and terminating all paid staff and volunteers -->
				<tr>
					<td align="left" valign="top" class="formMain">Notification of BBBSA of a vacancy in Chief Executive position (Program Director for sponsoring organizations)</td>
					<td align="left" valign="top" class="formMain">Identify where it is documented that BBBSA National Office must be contacted and by whom </td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd9bSO10b" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std9bSO10b")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std9bSO10b',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd9bSO10b" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std9bSO10b")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std9bSO10b',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd9bSO10b" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std9bSO10b"))) or Trim(GetSelfAssessment("Std9bSO10b")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std9bSO10b',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd9bSO10b" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std9bSO10b',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std9bSO10b" style="display:none;">
									 <label for="Std9bSO10bReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std9bSO10bReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Executive (Program Director for sponsored programs) attended new CEO training (new hires since 1/04) -->
				<tr>
					<td align="left" valign="top" class="formMain">Executive (Program Director for sponsored programs) attended new CEO training (new hires since 1/04)</td>
					<td align="left" valign="top" class="formMain">Check Personnel file for copy of transcript downloaded from the BBBS Learning Center</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentMAA814" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("MAA814")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA814',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentMAA814" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("MAA814")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA814',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentMAA814" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("MAA814"))) or Trim(GetSelfAssessment("MAA814")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA814',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentMAA814" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'MAA814',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="MAA814" style="display:none;">
									 <label for="MAA814Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="MAA814Reason" colspan="3">
							</div>
					</td>
				</tr>
				
				<tr>
				</tr>	
				
				<!-- Standard 10/Standard 11 (sponsoring organization) -->
				
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Standard 10/Standard 11 (sponsoring organization):  The affiliate, or BBBS program, has a human resource development and management system that is designed to effectively manage all paid, volunteer, and intern personnel</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>		
				
				<!-- Board-approved written personnel policies, compliant with local, state and federal labor laws -->
				<tr>
					<td align="left" valign="top" class="formMain">Written board-approved personnel policies, compliant with local, state and federal labor laws</td>
					<td align="left" valign="top" class="formMain">Ensure current personnel policies reflect all board approved changes, denoted by date and are provided to all staff, paid or volunteer.</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd10aSO11a" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10aSO11a")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10aSO11a',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd10aSO11a" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10aSO11a")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10aSO11a',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd10aSO11a" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std10aSO11a"))) or Trim(GetSelfAssessment("Std10aSO11a")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10aSO11a',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd10aSO11a" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std10aSO11a',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std10aSO11a" style="display:none;">
									 <label for="Std10aSO11aReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std10aSO11aReason" colspan="3">
							</div>
					</td>
				</tr>
									
				<!-- Written job descriptions for all paid and volunteer staff positions -->
				<tr>
					<td align="left" valign="top" class="formMain">Written job descriptions exist for all paid and volunteer staff positions</td>
					<td align="left" valign="top" class="formMain">Review job descriptions and update as necessary</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd10bSO11b" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10bSO11b")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10bSO11b',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd10bSO11b" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10bSO11b")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10bSO11b',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd10bSO11b" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std10bSO11b"))) or Trim(GetSelfAssessment("Std10bSO11b")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10bSO11b',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd10bSO11b" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std10bSO11b',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std10bSO11b" style="display:none;">
									 <label for="Std10bSO11bReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std10bSO11bReason" colspan="3">
							</div>
					</td>
				</tr>

				<!-- Volunteers functioning in staff positions meet same personnel requirements and follow same policies and procedures -->
				<tr>
					<td align="left" valign="top" class="formMain">Volunteers functioning in staff positions meet same personnel requirements and follow same policies and procedures</td>
					<td align="left" valign="top" class="formMain">Policy statement in current Personnel Manual and in job descriptions</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd10gSO11f" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10gSO11f")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10gSO11f',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd10gSO11f" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10gSO11f")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10gSO11f',true)">Out
						<br>
						<input type="radio" name="frmSelfAssessmentStd10gSO11f" value="3"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10gSO11f")) = "3" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10gSO11f',false)">N/A
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd10gSO11f" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std10gSO11f"))) or Trim(GetSelfAssessment("Std10gSO11f")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10gSO11f',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd10gSO11f" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std10gSO11f',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std10gSO11f" style="display:none;">
									 <label for="Std10gSO11fReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std10gSO11fReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Program Manual contains policies and procedures for non-degreed paraprofessionals --
				<tr>
					<td align="left" valign="top" class="formMain">Program Manual contains policies and procedures for non-degreed paraprofessionals</td>
					<td align="left" valign="top" class="formMain">Check Program Manual for policies and procedures for persons with less than a Bachelor's degree (who supervises them, trains them, and who makes service delivery decisions)</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd10hSO11g" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10hSO11g")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10hSO11g',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd10hSO11g" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10hSO11g")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10hSO11g',true)">Out
						<br>
						<input type="radio" name="frmSelfAssessmentStd10hSO11g" value="3"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10hSO11g")) = "3" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10hSO11g',false)">N/A
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd10hSO11g" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std10hSO11g"))) or Trim(GetSelfAssessment("Std10hSO11g")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10hSO11g',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd10hSO11g" value="0" checked>Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std10hSO11g" style="display:none;">
									 <label for="Std10hSO11gReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std10hSO11gReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Program Manual contains policies and procedures re: “professional/degreed staff” making all service delivery decisions -->
				<tr>
					<td align="left" valign="top" class="formMain">Agency Program Manual contains policies and procedures for the use of non-degreed paraprofessionals </td>
					<td align="left" valign="top" class="formMain">Ensure Program Manual has policies and procedures for non-degreed or paraprofessionals (persons with less than a Bachelor's degree) re: who will supervise and train them, and who will make all professional service delivery decisions.</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd10hSO11g2" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10hSO11g2")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10hSO11g2',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd10hSO11g2" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10hSO11g2")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10hSO11g2',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd10hSO11g2" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std10hSO11g2"))) or Trim(GetSelfAssessment("Std10hSO11g2")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10hSO11g2',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd10hSO11g2" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std10hSO11g2',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std10hSO11g2" style="display:none;">
									 <label for="Std10hSO11g2Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std10hSO11g2Reason" colspan="3">
							</div>
					</td>
				</tr>
								
				<!-- Non discrimination policy relative to staff --
				<tr>
					<td align="left" valign="top" class="formMain">Non discrimination policy relative to staff</td>
					<td align="left" valign="top" class="formMain">Documented in the Personnel Manual</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd10iSO11h" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10iSO11h")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10iSO11h',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd10iSO11h" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10iSO11h")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10iSO11h',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd10iSO11h" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std10iSO11h"))) or Trim(GetSelfAssessment("Std10iSO11h")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10iSO11h',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd10iSO11h" value="0" checked>Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std10iSO11h" style="display:none;">
									 <label for="Std10iSO11hReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std10iSO11hReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Board approved, competitive salary ranges -->
				<tr>
					<td align="left" valign="top" class="formMain">Board develops and approves competitive salary ranges for all paid staff</td>
					<td align="left" valign="top" class="formMain">Documentation that the Board, or committee thereof, has reviewed current salary ranges against current market and determined competitive salary ranges</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd10dSO11c" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10dSO11c")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10dSO11c',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd10dSO11c" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10dSO11c")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10dSO11c',true)">Out
						<br>

						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd10dSO11c" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std10dSO11c"))) or Trim(GetSelfAssessment("Std10dSO11c")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10dSO11c',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd10dSO11c" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std10dSO11c',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std10dSO11c" style="display:none;">
									 <label for="Std10dSO11cReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std10dSO11cReason" colspan="3">
							</div>
					</td>
				</tr>

				<!-- ED and program staff have at least a Bachelors degree -->
				<tr>
					<td align="left" valign="top" class="formMain">Chief Executive (Program Director for Sponsored affiliate) and program staff have at least a Bachelor's degree</td>
					<td align="left" valign="top" class="formMain">Review resumes, transcripts, diplomas in personnel file</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd10eSO11d" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10eSO11d")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10eSO11d',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd10eSO11d" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10eSO11d")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10eSO11d',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd10eSO11d" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std10eSO11d"))) or Trim(GetSelfAssessment("Std10eSO11d")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10eSO11d',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd10eSO11d" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std10eSO11d',false)">Not Entered
						<% end if %>							
					</td>										
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std10eSO11d" style="display:none;">
									 <label for="Std10eSO11dReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std10eSO11dReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Confidential personnel records maintained -->
				<tr>
					<td align="left" valign="top" class="formMain">Confidential personnel records on each employee, paid or volunteer, are maintained at corporate office</td>
					<td align="left" valign="top" class="formMain">Personnel files should have a cover sheet documenting content and be located in a secured location.</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd10fSO11e" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10fSO11e")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10fSO11e',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd10fSO11e" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10fSO11e")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10fSO11e',true)">Out
						<br>

						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd10fSO11e" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std10fSO11e"))) or Trim(GetSelfAssessment("Std10fSO11e")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10fSO11e',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd10fSO11e" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std10fSO11e',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std10fSO11e" style="display:none;">
									 <label for="Std10fSO11eReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std10fSO11eReason" colspan="3">
							</div>
					</td>
				</tr>

				<!-- Documentation of criminal history record check for staff / Volunteers -->
				<tr>
					<td align="left" valign="top" class="formMain">Documentation of criminal background  check for staff / volunteers</td>
					<td align="left" valign="top" class="formMain">Copy of  criminal background  check and, driver's license/ proof of insurance, if appropriate, should be located in personnel files of staff and case files of volunteers if serving in staff role</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd10jSO11i" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10jSO11i")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10jSO11i',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd10jSO11i" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10jSO11i")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10jSO11i',true)">Out
						<br>

						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd10jSO11i" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std10jSO11i"))) or Trim(GetSelfAssessment("Std10jSO11i")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10jSO11i',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd10jSO11i" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std10jSO11i',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std10jSO11i" style="display:none;">
									 <label for="Std10jSO11iReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std10jSO11iReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Documentation of attendance at BBBSA training offerings, annual meetings -->
				<tr>
					<td align="left" valign="top" class="formMain">Documentation of attendance at BBBSA training offerings, annual meetings</td>
					<td align="left" valign="top" class="formMain">Check Personnel file for copy of transcript downloaded from the BBBS Learning Center</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentMAA810" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("MAA810")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA810',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentMAA810" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("MAA810")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA810',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentMAA810" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("MAA810"))) or Trim(GetSelfAssessment("MAA810")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'MAA810',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentMAA810" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'MAA810',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="MAA810" style="display:none;">
									 <label for="MAA810Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="MAA810Reason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Documentation of annual personnel performance evaluations -->
				<tr>
					<td align="left" valign="top" class="formMain">Documentation of annual personnel performance evaluations</td>
					<td align="left" valign="top" class="formMain">Check Personnel files for copy of evaluation, signed by staff and supervisor</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd9bSO10b2" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std9bSO10b2")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std9bSO10b2',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd9bSO10b2" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std9bSO10b2")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std9bSO10b2',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd9bSO10b2" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std9bSO10b2"))) or Trim(GetSelfAssessment("Std9bSO10b2")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std9bSO10b2',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd9bSO10b2" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std9bSO10b2',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std9bSO10b2" style="display:none;">
									 <label for="Std9bSO10b2Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std9bSO10b2Reason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Non discrimination policy relative to staff -->
				<tr>
					<td align="left" valign="top" class="formMain"> Written board-approved Non-discrimination policy relative to staff and volunteers.</td>
					<td align="left" valign="top" class="formMain">Documented, at a minimum, in the Personnel Policies</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd10iSO11h" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10iSO11h")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10iSO11h',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd10iSO11h" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std10iSO11h")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10iSO11h',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd10iSO11h" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std10iSO11h"))) or Trim(GetSelfAssessment("Std10iSO11h")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std10iSO11h',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd10iSO11h" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std10iSO11h',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std10iSO11h" style="display:none;">
									 <label for="Std10iSO11hReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std10iSO11hReason" colspan="3">
							</div>
					</td>
				</tr>

				<!-- Standard 11/Standard 12 (sponsoring organization) -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0"  bgcolor="#c0c0c0">Standard 11/Standard 12 (sponsoring organization):  The affiliate provides facilities and working conditions, which are conducive  to accomplishing the operation of the affiliate including provisions to conduct private interviews, conforming to laws and regulations governing occupational health and safety</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0"  bgcolor="#c0c0c0">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0"  bgcolor="#c0c0c0"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>		
				
				<!-- Facilities meet ADA, OSHA standards -->
				<tr>
					<td align="left" valign="top" class="formMain">Facilities meet ADA, OSHA standards</td>
					<td align="left" valign="top" class="formMain">Copy of annual facilities audit on file; Inspect the environment for safety and cleanliness; Inspect the equipment used by staff to perform necessary work for safety and proper functioning</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd11SO12" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std11SO12")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std11SO12',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd11SO12" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std11SO12")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std11SO12',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd11SO12" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std11SO12"))) or Trim(GetSelfAssessment("Std11SO12")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std11SO12',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd11SO12" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std11SO12',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std11SO12" style="display:none;">
									 <label for="Std11SO12Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std11SO12Reason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Facilities allow for privacy during interviews -->
				<tr>
					<td align="left" valign="top" class="formMain">Facilities allow for privacy during interviews</td>
					<td align="left" valign="top" class="formMain">Assess that staff have private space for interviews</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd11SO122" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std11SO122")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std11SO122',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd11SO122" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std11SO122")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std11SO122',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd11SO122" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std11SO122"))) or Trim(GetSelfAssessment("Std11SO122")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std11SO122',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd11SO122" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std11SO122',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std11SO122" style="display:none;">
									 <label for="Std11SO122Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std11SO122Reason" colspan="3">
							</div>
					</td>
				</tr>
				
				
		<!-- End Operational Section -->


		<% else %>		
		<!-- Begin Program Section -->
		
			<!-- Prepopulate Operational Fields -->
		
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std1a") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd1a">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std1aReason") %><% Else %><% End If %>" name="Std1aReason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std1b") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd1b">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std1bReason") %><% Else %><% End If %>" name="Std1bReason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Form990") %><% Else %>0<% End If %>" name="frmSelfAssessmentForm990">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Form990Reason") %><% Else %><% End If %>" name="Form990Reason">			
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std1c") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd1c">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std1cReason") %><% Else %><% End If %>" name="Std1cReason">			
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Bylaws") %><% Else %>0<% End If %>" name="frmSelfAssessmentBylaws">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("BylawsReason") %><% Else %><% End If %>" name="BylawsReason">			
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("MAA") %><% Else %>0<% End If %>" name="frmSelfAssessmentMAA">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("MAAReason") %><% Else %><% End If %>" name="MAAReason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std1LogoAndName") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd1LogoAndName">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std1LogoAndNameReason") %><% Else %><% End If %>" name="Std1LogoAndNameReason">			
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std2") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd2">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std2Reason") %><% Else %><% End If %>" name="Std2Reason">						
			<!--<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std2aSO3a") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd2aSO3a">			
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Brdtrainplan") %><% Else %>0<% End If %>" name="frmSelfAssessmentBrdtrainplan">-->
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std2bSO3b") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd2bSO3b">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std2bSO3bReason") %><% Else %><% End If %>" name="Std2bSO3bReason">			
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("MAA810conf") %><% Else %>0<% End If %>" name="frmSelfAssessmentMAA810conf">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("MAA810confReason") %><% Else %><% End If %>" name="MAA810confReason">			
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std2SO") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd2SO">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std2SOReason") %><% Else %><% End If %>" name="Std2SOReason">			
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std3SO4m") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd3SO4m">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std3SO4mReason") %><% Else %><% End If %>" name="Std3SO4mReason">						
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std3SO4v") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd3SO4v">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std3SO4vReason") %><% Else %><% End If %>" name="Std3SO4vReason">			
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std4SO5") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd4SO5">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std4SO5Reason") %><% Else %><% End If %>" name="Std4SO5Reason">						
			<input type="hidden" value="0" name="frmSelfAssessmentOpplan">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std5opsSO6") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd5opsSO6">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std5opsSO6Reason") %><% Else %><% End If %>" name="Std5opsSO6Reason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std5pgmSO6") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd5pgmSO6">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std5pgmSO6Reason") %><% Else %><% End If %>" name="Std5pgmSO6Reason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std5filesSO6") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd5filesSO6">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std5filesSO6Reason") %><% Else %><% End If %>" name="Std5filesSO6Reason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std6SO7budget") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd6SO7budget">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std6SO7budgetReason") %><% Else %><% End If %>" name="Std6SO7budgetReason">	
			<input type="hidden" value="0" name="frmSelfAssessmentMAA810exp">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("MAA32") %><% Else %>0<% End If %>" name="frmSelfAssessmentMAA32">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("MAA32Reason") %><% Else %><% End If %>" name="MAA32Reason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std6SO7b") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd6SO7b">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std6SO7bReason") %><% Else %><% End If %>" name="Std6SO7bReason">			
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std6SO7") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd6SO7">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std6SO7Reason") %><% Else %><% End If %>" name="Std6SO7Reason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std7SO8") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd7SO8">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std7SO8Reason") %><% Else %><% End If %>" name="Std7SO8Reason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("MAA88") %><% Else %>0<% End If %>" name="frmSelfAssessmentMAA88">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("MAA88Reason") %><% Else %><% End If %>" name="MAA88Reason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("MAA82") %><% Else %>0<% End If %>" name="frmSelfAssessmentMAA82">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("MAA82Reason") %><% Else %><% End If %>" name="MAA82Reason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("StdSO8a") %><% Else %>0<% End If %>" name="frmSelfAssessmentStdSO8a">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("StdSO8aReason") %><% Else %><% End If %>" name="StdSO8aReason">
			<!--2009<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("StdSO8b") %><% Else %>0<% End If %>" name="frmSelfAssessmentStdSO8b">-->
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("StdSO8bReason") %><% Else %><% End If %>" name="StdSO8bReason">			
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("StdSO8c") %><% Else %>0<% End If %>" name="frmSelfAssessmentStdSO8c">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("StdSO8cReason") %><% Else %><% End If %>" name="StdSO8cReason">
			<!--<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std8SO9crisis") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd8SO9crisis">-->			
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std8SO9risk") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd8SO9risk">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std8SO9riskReason") %><% Else %><% End If %>" name="Std8SO9riskReason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("MAA9") %><% Else %>0<% End If %>" name="frmSelfAssessmentMAA9">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("MAA9Reason") %><% Else %><% End If %>" name="MAA9Reason">			
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10bSO11b2") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd10bSO11b2">
					<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10bSO11b2Reason") %><% Else %><% End If %>" name="Std10bSO11b2Reason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std9a") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd9a">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std9aReason") %><% Else %><% End If %>" name="Std9aReason">			
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std9SO10") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd9SO10">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std9SO10Reason") %><% Else %><% End If %>" name="Std9SO10Reason">						
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("MAA813") %><% Else %>0<% End If %>" name="frmSelfAssessmentMAA813">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("MAA813Reason") %><% Else %><% End If %>" name="MAA813Reason">			
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std9bSO10b") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd9bSO10b">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std9bSO10bReason") %><% Else %><% End If %>" name="Std9bSO10bReason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("MAA814") %><% Else %>0<% End If %>" name="frmSelfAssessmentMAA814">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("MAA814Reason") %><% Else %><% End If %>" name="MAA814Reason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10aSO11a") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd10aSO11a">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10aSO11aReason") %><% Else %><% End If %>" name="Std10aSO11aReason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10bSO11b") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd10bSO11b">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10bSO11bReason") %><% Else %><% End If %>" name="Std10bSO11bReason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10gSO11f") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd10gSO11f">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10gSO11fReason") %><% Else %><% End If %>" name="Std10gSO11fReason">
			<!--<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10hSO11g") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd10hSO11g">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10hSO11gReason") %><% Else %><% End If %>" name="Std10hSO11gReason">-->
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10hSO11g2") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd10hSO11g2">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10hSO11g2Reason") %><% Else %><% End If %>" name="Std10hSO11g2Reason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10iSO11h") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd10iSO11h">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10iSO11hReason") %><% Else %><% End If %>" name="Std10iSO11hReason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10dSO11c") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd10dSO11c">
  		<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10dSO11cReason") %><% Else %><% End If %>" name="Std10dSO11cReason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10eSO11d") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd10eSO11d">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10eSO11dReason") %><% Else %><% End If %>" name="Std10eSO11dReason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10fSO11e") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd10fSO11e">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10fSO11eReason") %><% Else %><% End If %>" name="Std10fSO11eReason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10jSO11i") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd10jSO11i">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std10jSO11iReason") %><% Else %><% End If %>" name="Std10jSO11iReason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("MAA810") %><% Else %>0<% End If %>" name="frmSelfAssessmentMAA810">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("MAA810Reason") %><% Else %><% End If %>" name="MAA810Reason">			
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std9bSO10b2") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd9bSO10b2">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std9bSO10b2Reason") %><% Else %><% End If %>" name="Std9bSO10b2Reason">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std11SO12") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd11SO12">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std11SO12Reason") %><% Else %><% End If %>" name="Std11SO12Reason">			
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std11SO122") %><% Else %>0<% End If %>" name="frmSelfAssessmentStd11SO122">
			<input type="hidden" value="<% If say = "edit" Then %><%= GetSelfAssessment("Std11SO122Reason") %><% Else %><% End If %>" name="Std11SO122Reason">			
			

				<!-- Standard 12/Standard 13 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Standard 12/Standard 13 (sponsoring organization):  The Program Manual contains the policies, procedures, and forms to be used for implementing all One-To-One services</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>
				
				<!-- The Program Manual contains board-approved written policies, procedures and forms compliant with Practices of One-To-One Service -->
				<tr>
					<td align="left" valign="top" class="formMain">Written board-approved Agency Program Manual contains policies, procedures and forms compliant with the Standards of Practice of One-To-One Service</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval of policies; document in Program Manual</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd12aSO13a" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std12aSO13a")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std12aSO13a',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd12aSO13a" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std12aSO13a")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std12aSO13a',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd12aSO13a" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std12aSO13a"))) or Trim(GetSelfAssessment("Std12aSO13a")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std12aSO13a',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd12aSO13a" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std12aSO13a',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std12aSO13a" style="display:none;">
									 <label for="Std12aSO13aReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std12aSO13aReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Policy on eligibility criteria for volunteers & youth -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy on eligibility criteria for volunteers & youth and procedures for determining eligibility</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentpolicyeligible" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policyeligible")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyeligible',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentpolicyeligible" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policyeligible")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyeligible',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentpolicyeligible" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("policyeligible"))) or Trim(GetSelfAssessment("policyeligible")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyeligible',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentpolicyeligible" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'policyeligible',false)">Not Entered
						<% end if %>							
					</td>
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="policyeligible" style="display:none;">
									 <label for="policyeligibleReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="policyeligibleReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Procedures --
				<tr>
					<td align="left" valign="top" class="formMain">Procedures</td>
					<td align="left" valign="top" class="formMain">Procedures for determining eligibility</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentproceligible" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("proceligible")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'proceligible',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentproceligible" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("proceligible")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'proceligible',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentproceligible" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("proceligible"))) or Trim(GetSelfAssessment("proceligible")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'proceligible',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentproceligible" value="0" checked>Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="proceligible" style="display:none;">
									 <label for="proceligibleReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="proceligibleReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Policy on youth outreach -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy and procedures on youth outreach</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentpolicychildrec" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policychildrec")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policychildrec',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentpolicychildrec" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policychildrec")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policychildrec',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentpolicychildrec" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("policychildrec"))) or Trim(GetSelfAssessment("policychildrec")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policychildrec',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentpolicychildrec" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'policychildrec',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="policychildrec" style="display:none;">
									 <label for="policychildrecReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="policychildrecReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Procedures --
				<tr>
					<td align="left" valign="top" class="formMain">Procedures</td>
					<td align="left" valign="top" class="formMain">Procedures for recruiting youth</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentprocchildrec" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("procchildrec")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procchildrec',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentprocchildrec" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("procchildrec")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procchildrec',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentprocchildrec" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("procchildrec"))) or Trim(GetSelfAssessment("procchildrec")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procchildrec',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentprocchildrec" value="0" checked>Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="procchildrec" style="display:none;">
									 <label for="procchildrecReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="procchildrecReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Policy on volunteer recruitment -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy  and procedures on volunteer recruitment</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentpolicyvolrec" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policyvolrec")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyvolrec',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentpolicyvolrec" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policyvolrec")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyvolrec',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentpolicyvolrec" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("policyvolrec"))) or Trim(GetSelfAssessment("policyvolrec")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyvolrec',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentpolicyvolrec" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'policyvolrec',false)">Not Entered
						<% end if %>							
					</td>											
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="policyvolrec" style="display:none;">
									 <label for="policyvolrecReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="policyvolrecReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Procedures --
				<tr>
					<td align="left" valign="top" class="formMain">Procedures</td>
					<td align="left" valign="top" class="formMain">Procedures for recruiting volunteers</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentprocvolrec" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("procvolrec")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procvolrec',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentprocvolrec" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("procvolrec")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procvolrec',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentprocvolrec" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("procvolrec"))) or Trim(GetSelfAssessment("procvolrec")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procvolrec',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentprocvolrec" value="0" checked>Not Entered
						<% end if %>							
					</td>																
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="procvolrec" style="display:none;">
									 <label for="procvolrecReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="procvolrecReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Policy on referrals -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy and procedures on referrals</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentpolicyref" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policyref")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyref',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentpolicyref" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policyref")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyref',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentpolicyref" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("policyref"))) or Trim(GetSelfAssessment("policyref")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyref',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentpolicyref" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'policyref',false)">Not Entered
						<% end if %>							
					</td>
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="policyref" style="display:none;">
									 <label for="policyrefReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="policyrefReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Procedures --
				<tr>
					<td align="left" valign="top" class="formMain">Procedures</td>
					<td align="left" valign="top" class="formMain">Procedures for handling referrals</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentprocref" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("procref")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procref',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentprocref" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("procref")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procref',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentprocref" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("procref"))) or Trim(GetSelfAssessment("procref")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procref',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentprocref" value="0" checked>Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="procref" style="display:none;">
									 <label for="procrefReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="procrefReason" colspan="3">
							</div>
					</td>
				</tr>

				<!-- Policy on inquiries -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy and procedures on inquiries</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentpolicyinq" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policyinq")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyinq',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentpolicyinq" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policyinq")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyinq',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentpolicyinq" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("policyinq"))) or Trim(GetSelfAssessment("policyinq")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyinq',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentpolicyinq" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'policyinq',false)">Not Entered
						<% end if %>							
					</td>										
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="policyinq" style="display:none;">
									 <label for="policyinqReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="policyinqReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Procedures --
				<tr>
					<td align="left" valign="top" class="formMain">Procedures</td>
					<td align="left" valign="top" class="formMain">Procedures for handling inquiries</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentprocinq" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("procinq")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procinq',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentprocinq" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("procinq")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procinq',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentprocinq" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("procinq"))) or Trim(GetSelfAssessment("procinq")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procinq',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentprocinq" value="0" checked>Not Entered
						<% end if %>							
					</td>
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="procinq" style="display:none;">
									 <label for="procinqReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="procinqReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Policies on intake -->
				<tr>
					<td align="left" valign="top" class="formMain">Policies and procedures on intake</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentpolicyintake" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policyintake")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyintake',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentpolicyintake" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policyintake")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyintake',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentpolicyintake" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("policyintake"))) or Trim(GetSelfAssessment("policyintake")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyintake',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentpolicyintake" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'policyintake',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="policyintake" style="display:none;">
									 <label for="policyintakeReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="policyintakeReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Procedures --
				<tr>
					<td align="left" valign="top" class="formMain">Procedures</td>
					<td align="left" valign="top" class="formMain">Procedures for handling the intake process</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentprocintake" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("procintake")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procintake',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentprocintake" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("procintake")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procintake',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentprocintake" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("procintake"))) or Trim(GetSelfAssessment("procintake")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procintake',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentprocintake" value="0" checked>Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="procintake" style="display:none;">
									 <label for="procintakeReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="procintakeReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Policies on matching -->
				<tr>
					<td align="left" valign="top" class="formMain">Policies  and procedures on matching</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentpolicymatch" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policymatch")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policymatch',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentpolicymatch" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policymatch")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policymatch',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentpolicymatch" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("policymatch"))) or Trim(GetSelfAssessment("policymatch")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policymatch',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentpolicymatch" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'policymatch',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="policymatch" style="display:none;">
									 <label for="policymatchReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="policymatchReason" colspan="3">
							</div>
					</td>
				</tr>
	
				<!-- Procedures --
				<tr>
					<td align="left" valign="top" class="formMain">Procedures</td>
					<td align="left" valign="top" class="formMain">Procedures for handling the matching process</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentprocmatch" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("procmatch")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procmatch',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentprocmatch" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("procmatch")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procmatch',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentprocmatch" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("procmatch"))) or Trim(GetSelfAssessment("procmatch")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procmatch',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentprocmatch" value="0" checked>Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="procmatch" style="display:none;">
									 <label for="procmatchReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="procmatchReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Policies on supervision -->
				<tr>
					<td align="left" valign="top" class="formMain">Policies and procedures on supervision</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentpolicysup" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policysup")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policysup',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentpolicysup" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policysup")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policysup',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentpolicysup" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("policysup"))) or Trim(GetSelfAssessment("policysup")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policysup',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentpolicysup" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'policysup',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="policysup" style="display:none;">
									 <label for="policysupReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="policysupReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Procedures --
				<tr>
					<td align="left" valign="top" class="formMain">Procedures</td>
					<td align="left" valign="top" class="formMain">Procedures for handling the match supervision process</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentprocsup" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("procsup")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procsup',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentprocsup" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("procsup")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procsup',true)">Out
						<br>

						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentprocsup" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("procsup"))) or Trim(GetSelfAssessment("procsup")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procsup',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentprocsup" value="0" checked>Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="procsup" style="display:none;">
									 <label for="procsupReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="procsupReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Policies on closure -->
				<tr>
					<td align="left" valign="top" class="formMain">Policies and procedures on closure</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentpolicyclosure" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policyclosure")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyclosure',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentpolicyclosure" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policyclosure")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyclosure',true)">Out
						<br>

						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentpolicyclosure" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("policyclosure"))) or Trim(GetSelfAssessment("policyclosure")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyclosure',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentpolicyclosure" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'policyclosure',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="policyclosure" style="display:none;">
									 <label for="policyclosureReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="policyclosureReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Procedures --
				<tr>
					<td align="left" valign="top" class="formMain">Procedures</td>
					<td align="left" valign="top" class="formMain">Procedures for handling the match closure process</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentprocclosure" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("procclosure")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procclosure',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentprocclosure" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("procclosure")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procclosure',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentprocclosure" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("procclosure"))) or Trim(GetSelfAssessment("procclosure")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procclosure',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentprocclosure" value="0" checked>Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="procclosure" style="display:none;">
									 <label for="procclosureReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="procclosureReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Policies on case record keeping -->
				<tr>
					<td align="left" valign="top" class="formMain">Policies and procedures on case record keeping</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentpolicyrecords" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policyrecords")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyrecords',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentpolicyrecords" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policyrecords")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyrecords',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentpolicyrecords" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("policyrecords"))) or Trim(GetSelfAssessment("policyrecords")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyrecords',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentpolicyrecords" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'policyrecords',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="policyrecords" style="display:none;">
									 <label for="policyrecordsReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="policyrecordsReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Procedures --
				<tr>
					<td align="left" valign="top" class="formMain">Procedures</td>
					<td align="left" valign="top" class="formMain">Procedures for handling documentation and record keeping</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentprocrecords" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("procrecords")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procrecords',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentprocrecords" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("procrecords")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procrecords',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentprocrecords" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("procrecords"))) or Trim(GetSelfAssessment("procrecords")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'procrecords',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentprocrecords" value="0" checked>Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="procrecords" style="display:none;">
									 <label for="procrecordsReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="procrecordsReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Policies on handling documentation -->
				<tr>
					<td align="left" valign="top" class="formMain">Policies and procedures for handling documentation </td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd12PPHandlingDoc" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std12PPHandlingDoc")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std12PPHandlingDoc',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd12PPHandlingDoc" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std12PPHandlingDoc")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std12PPHandlingDoc',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd12PPHandlingDoc" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std12PPHandlingDoc"))) or Trim(GetSelfAssessment("Std12PPHandlingDoc")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std12PPHandlingDoc',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd12PPHandlingDoc" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std12PPHandlingDoc',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std12PPHandlingDoc" style="display:none;">
									 <label for="Std12PPHandlingDocReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std12PPHandlingDocReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Program Manual addresses risk management issues with written Board-approved policies  -->
				
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Program Manual addresses risk management issues with written Board-approved policies </td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>		
				
				<!-- Policy on overnight visits of youth with volunteers -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy on overnight visits of youth with volunteers</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentpolicyovernite" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policyovernite")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyovernite',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentpolicyovernite" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policyovernite")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyovernite',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentpolicyovernite" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("policyovernite"))) or Trim(GetSelfAssessment("policyovernite")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyovernite',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentpolicyovernite" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'policyovernite',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="policyovernite" style="display:none;">
									 <label for="policyoverniteReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="policyoverniteReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Policy on child sexual abuse prevention orientation, education, and training -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy on child sexual abuse prevention orientation, education, and training</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentpolicysexabuse" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policysexabuse")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policysexabuse',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentpolicysexabuse" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policysexabuse")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policysexabuse',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentpolicysexabuse" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("policysexabuse"))) or Trim(GetSelfAssessment("policysexabuse")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policysexabuse',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentpolicysexabuse" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'policysexabuse',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="policysexabuse" style="display:none;">
									 <label for="policysexabuseReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="policysexabuseReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Policy on board / staff serving as Bigs -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy on board / staff serving as Bigs</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentpolicystaffasbigs" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policystaffasbigs")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policystaffasbigs',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentpolicystaffasbigs" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policystaffasbigs")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policystaffasbigs',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentpolicystaffasbigs" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("policystaffasbigs"))) or Trim(GetSelfAssessment("policystaffasbigs")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policystaffasbigs',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentpolicystaffasbigs" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'policystaffasbigs',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="policystaffasbigs" style="display:none;">
									 <label for="policystaffasbigsReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="policystaffasbigsReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Policy on interviewing other persons residing with volunteer applicant -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy on interviewing other persons residing with volunteer applicant</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentpolicyinterothers" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policyinterothers")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyinterothers',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentpolicyinterothers" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policyinterothers")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyinterothers',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentpolicyinterothers" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("policyinterothers"))) or Trim(GetSelfAssessment("policyinterothers")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policyinterothers',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentpolicyinterothers" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'policyinterothers',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="policyinterothers" style="display:none;">
									 <label for="policyinterothersReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="policyinterothersReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Procedures for obtaining information about disclosed prior BBBSA experience -->
				<tr>
					<td align="left" valign="top" class="formMain">Procedures for obtaining information about disclosed prior BBBSA experience</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentpolicypriorexp" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policypriorexp")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policypriorexp',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentpolicypriorexp" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("policypriorexp")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policypriorexp',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentpolicypriorexp" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("policypriorexp"))) or Trim(GetSelfAssessment("policypriorexp")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'policypriorexp',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentpolicypriorexp" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'policypriorexp',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="policypriorexp" style="display:none;">
									 <label for="policypriorexpReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="policypriorexpReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Recommended Best Practices for Case File Audits -->
				<tr>
					<td colspan="3" class="formMainBold" align="center">FOR RECOMMENDED BEST PRACTICE for CASE FILE AUDITS, <a href="http://agencyconnection.bbbs.org/site/c.9dJGKRNqFmG/b.1742167/k.8DA1/Child_Safety.htm" target="_blank">click here</a> to consult our Child Safety Web Page and/or contact Julie Novak, Director of Child Safety and Quality Assurance, at <a href="mailto:Julie.Novak@bbbs.org">Julie.Novak@bbbs.org</a></td>
				</tr>
				
				
				<!-- Standard 13/Standard 14 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Standard 13/Standard 14 (sponsoring organization):</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>																		
													
				<!-- Procedures for obtaining information about disclosed prior BBBSA experience -->
				<tr>
					<td align="left" valign="top" class="formMain">The children, youth, and volunteer inquiry process used by the affiliate provides the opportunity for the affiliate, parent/guardian, and volunteer to determine the appropriateness of participation and provides an orientation to all services provided by the affiliates</td>
					<td align="left" valign="top" class="formMain">Review procedures and documentation of practice for inquiry and orientation to all services. Document date of last Board review and approval</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd13SO14" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std13SO14")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std13SO14',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd13SO14" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std13SO14")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std13SO14',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd13SO14" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std13SO14"))) or Trim(GetSelfAssessment("Std13SO14")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std13SO14',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd13SO14" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std13SO14',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std13SO14" style="display:none;">
									 <label for="Std13SO14Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std13SO14Reason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Standard 14/Standard 15 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Standard 14/Standard 15 (sponsoring organization): The child intake process used by the affiliate is a consistent process to determine eligibility of children and youth for services based upon written eligibility criteria.  Children and youth are not excluded on the basis of race, religion, national origin, gender, sexual orientation, disability, or marital status of parent</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>
				
				<!-- Written consent from parent / guardian -->
				<tr>
					<td align="left" valign="top" class="formMain">Written consent from parent / guardian</td>
					<td align="left" valign="top" class="formMain">Copy of Application signed by parent / guardian is in case file</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentchildconsent" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("childconsent")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'childconsent',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentchildconsent" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("childconsent")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'childconsent',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentchildconsent" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("childconsent"))) or Trim(GetSelfAssessment("childconsent")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'childconsent',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentchildconsent" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'childconsent',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="childconsent" style="display:none;">
									 <label for="childconsentReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="childconsentReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- In-person interview with child -->
				<tr>
					<td align="left" valign="top" class="formMain">In-person interview with child</td>
					<td align="left" valign="top" class="formMain">Verify documentation of in-person interview is in case file</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentchildinterview" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("childinterview")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'childinterview',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentchildinterview" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("childinterview")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'childinterview',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentchildinterview" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("childinterview"))) or Trim(GetSelfAssessment("childinterview")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'childinterview',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentchildinterview" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'childinterview',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="childinterview" style="display:none;">
									 <label for="childinterviewReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="childinterviewReason" colspan="3">
							</div>
					</td>
				</tr>					
		
				<!-- In-person interview with parent / guardian (CBM only) -->
				<tr>
					<td align="left" valign="top" class="formMain">In-person interview with parent / guardian (CBM only)</td>
					<td align="left" valign="top" class="formMain">Verify documentation of in-person interview is in case file</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentchildparinterview" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("childparinterview")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'childparinterview',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentchildparinterview" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("childparinterview")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'childparinterview',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentchildparinterview" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("childparinterview"))) or Trim(GetSelfAssessment("childparinterview")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'childparinterview',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentchildparinterview" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'childparinterview',false)">Not Entered
						<% end if %>							
					</td>											
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="childparinterview" style="display:none;">
									 <label for="childparinterviewReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="childparinterviewReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Assessment of home environment (CBM only) -->
				<tr>
					<td align="left" valign="top" class="formMain">Assessment of home environment (CBM only)</td>
					<td align="left" valign="top" class="formMain">Review procedures for home assessment in program manual; Verify documentation of home assessment is in case file</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentchildhomeassess" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("childhomeassess")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'childhomeassess',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentchildhomeassess" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("childhomeassess")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'childhomeassess',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentchildhomeassess" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("childhomeassess"))) or Trim(GetSelfAssessment("childhomeassess")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'childhomeassess',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentchildhomeassess" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'childhomeassess',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="childhomeassess" style="display:none;">
									 <label for="childhomeassessReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="childhomeassessReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Standard 15/Standard 16 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Standard 15/Standard 16 (sponsoring organization): The professional staff conducts an in-person interview with the volunteer.  The volunteer intake process elicits necessary information enabling the professional staff to prepare recommendations based upon the volunteer's ability to help meet the needs of the child</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>

				<!-- Application -->
				<tr>
					<td align="left" valign="top" class="formMain">Application</td>
					<td align="left" valign="top" class="formMain">Document the review of written application</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentvolconsent" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("volconsent")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'volconsent',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentvolconsent" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("volconsent")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'volconsent',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentvolconsent" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("volconsent"))) or Trim(GetSelfAssessment("volconsent")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'volconsent',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentvolconsent" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'volconsent',false)">Not Entered
						<% end if %>							
					</td>
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="volconsent" style="display:none;">
									 <label for="volconsentReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="volconsentReason" colspan="3">
							</div>
					</td>
				</tr>
					
				<!-- Obtain references (CBM = 3; SBM = 1) -->
				<tr>
					<td align="left" valign="top" class="formMain">Appropriate number of references  are obtained </td>
					<td align="left" valign="top" class="formMain">Review procedures for obtaining references; Verify documentation of references is in case file</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentvolreferences" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("volreferences")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'volreferences',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentvolreferences" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("volreferences")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'volreferences',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentvolreferences" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("volreferences"))) or Trim(GetSelfAssessment("volreferences")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'volreferences',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentvolreferences" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'volreferences',false)">Not Entered
						<% end if %>							
					</td>
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="volreferences" style="display:none;">
									 <label for="volreferencesReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="volreferencesReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Obtain criminal history record -->
				<tr>
					<td align="left" valign="top" class="formMain">Obtain criminal background check(s)</td>
					<td align="left" valign="top" class="formMain">Review procedures for obtaining criminal background checks in Program Manual; Verify documentation of criminal history record is in case file</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentvolcriminal" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("volcriminal")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'volcriminal',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentvolcriminal" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("volcriminal")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'volcriminal',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentvolcriminal" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("volcriminal"))) or Trim(GetSelfAssessment("volcriminal")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'volcriminal',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentvolcriminal" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'volcriminal',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="volcriminal" style="display:none;">
									 <label for="volcriminalReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="volcriminalReason" colspan="3">
							</div>
					</td>
				</tr>
					
				<!-- In-person interview --
				<tr>
					<td align="left" valign="top" class="formMain">In-person interview</td>
					<td align="left" valign="top" class="formMain">Verify documentation of in-person interview is in case file</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentvolinterview" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("volinterview")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'volinterview',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentvolinterview" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("volinterview")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'volinterview',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentvolinterview" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("volinterview"))) or Trim(GetSelfAssessment("volinterview")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'volinterview',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentvolinterview" value="0" checked>Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="volinterview" style="display:none;">
									 <label for="volinterviewReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="volinterviewReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Assessment of home environment (CBM only) --
				<tr>
					<td align="left" valign="top" class="formMain">Assessment of home environment (CBM only)</td>
					<td align="left" valign="top" class="formMain">Review procedures for home assessment in program manual; Verify documentation of home assessment is in case file</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentvolhomeassess" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("volhomeassess")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'volhomeassess',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentvolhomeassess" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("volhomeassess")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'volhomeassess',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentvolhomeassess" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("volhomeassess"))) or Trim(GetSelfAssessment("volhomeassess")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'volhomeassess',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentvolhomeassess" value="0" checked>Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="volhomeassess" style="display:none;">
									 <label for="volhomeassessReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="volhomeassessReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Written professional matching recommendations -->
				<tr>
					<td align="left" valign="top" class="formMain">Written professional matching recommendations</td>
					<td align="left" valign="top" class="formMain">Review Program Manual for procedures; Verify documentation of written matching recommendations by professional staff is in case file</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentvolmatching" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("volmatching")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'volmatching',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentvolmatching" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("volmatching")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'volmatching',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentvolmatching" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("volmatching"))) or Trim(GetSelfAssessment("volmatching")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'volmatching',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentvolmatching" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'volmatching',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="volmatching" style="display:none;">
									 <label for="volmatchingReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="volmatchingReason" colspan="3">
							</div>
					</td>
				</tr>
		
				<!-- Provide opportunity for training -->
				<tr>
					<td align="left" valign="top" class="formMain">Provide opportunity for training</td>
					<td align="left" valign="top" class="formMain">Verify documentation that training opportunities have been offered to volunteers and parents/guardians, as needed </td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentvoltraining" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("voltraining")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'voltraining',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentvoltraining" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("voltraining")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'voltraining',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentvoltraining" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("voltraining"))) or Trim(GetSelfAssessment("voltraining")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'voltraining',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentvoltraining" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'voltraining',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="voltraining" style="display:none;">
									 <label for="voltrainingReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="voltrainingReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Standard 16/Standard 17 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Standard 16/Standard 17 (sponsoring organization): The matching process enables the professional staff to assess and take into consideration all information gathered through applications and interviews of all parties</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>								
		
				<!-- Child approves proposed match -->
				<tr>
					<td align="left" valign="top" class="formMain">Child approves proposed match</td>
					<td align="left" valign="top" class="formMain">Review program manual for procedures; Verify documentation that child approves is in case file</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentApproveschild" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Approveschild")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Approveschild',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentApproveschild" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Approveschild")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Approveschild',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentApproveschild" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Approveschild"))) or Trim(GetSelfAssessment("Approveschild")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Approveschild',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentApproveschild" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Approveschild',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Approveschild" style="display:none;">
									 <label for="ApproveschildReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="ApproveschildReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Parent / guardian approves proposed match -->
				<tr>
					<td align="left" valign="top" class="formMain">Parent / guardian approves proposed match</td>
					<td align="left" valign="top" class="formMain">Review program manual for procedures; Verify documentation that parent/guardian approves is in case file</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentApprovesparent" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Approvesparent")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Approvesparent',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentApprovesparent" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Approvesparent")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Approvesparent',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentApprovesparent" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Approvesparent"))) or Trim(GetSelfAssessment("Approvesparent")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Approvesparent',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentApprovesparent" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Approvesparent',false)">Not Entered
						<% end if %>							
					</td>						
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Approvesparent" style="display:none;">
									 <label for="ApprovesparentReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="ApprovesparentReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Volunteer approves proposed match -->
				<tr>
					<td align="left" valign="top" class="formMain">Volunteer approves proposed match</td>
					<td align="left" valign="top" class="formMain">Review program manual for procedures; Verify documentation that volunteer approves is in case file</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentApprovesvol" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Approvesvol")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Approvesvol',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentApprovesvol" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Approvesvol")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Approvesvol',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentApprovesvol" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Approvesvol"))) or Trim(GetSelfAssessment("Approvesvol")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Approvesvol',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentApprovesvol" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Approvesvol',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Approvesvol" style="display:none;">
									 <label for="ApprovesvolReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="ApprovesvolReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- In-person match introduction by BBBSA staff or designee -->
				<tr>
					<td align="left" valign="top" class="formMain">In-person match introduction by BBBSA staff or designee</td>
					<td align="left" valign="top" class="formMain">Review program manual for procedures; Verify documentation of in-person match introduction is in case file</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentInpersonmatch" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Inpersonmatch")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Inpersonmatch',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentInpersonmatch" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Inpersonmatch")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Inpersonmatch',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentInpersonmatch" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Inpersonmatch"))) or Trim(GetSelfAssessment("Inpersonmatch")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Inpersonmatch',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentInpersonmatch" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Inpersonmatch',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Inpersonmatch" style="display:none;">
									 <label for="InpersonmatchReason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="InpersonmatchReason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Standard 17/Standard 18 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Standard 17/Standard 18 (sponsoring organization)</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>		
				
				<!-- In-person match introduction by BBBSA staff or designee -->
				<tr>
					<td align="left" valign="top" class="formMain">Professional staff develops and annually updates an outcome-based plan for each match</td>
					<td align="left" valign="top" class="formMain">Review procedures in Program Manual; Verify documentation is in case file and that the annual outcome-based plan is complete, up-dated annually and on-file</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd17SO18" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std17SO18")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std17SO18',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd17SO18" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std17SO18")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std17SO18',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd17SO18" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std17SO18"))) or Trim(GetSelfAssessment("Std17SO18")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std17SO18',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd17SO18" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std17SO18',false)">Not Entered
						<% end if %>							
					</td>
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std17SO18" style="display:none;">
									 <label for="Std17SO18Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std17SO18Reason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Standard 18/Standard 19 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Standard 18/Standard 19 (sponsoring organization)</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>	
				
				<!-- Professional staff oversees regular supervisory contact -->
				<tr>
					<td align="left" valign="top" class="formMain">Professional staff oversees regular supervisory contact with volunteer, parent/guardian/ and child in accordance with the Program Manual and Standard of Practice for One-To-One Service.</td>
					<td align="left" valign="top" class="formMain">See Standard of Practice for One-To-One Service for minimum criteria; Review procedures in Program Manual; Verify documentation, indicating date and person contacted, to assure contact were made according to Standards</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd18SO19" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std18SO19")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std18SO19',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd18SO19" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std18SO19")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std18SO19',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd18SO19" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std18SO19"))) or Trim(GetSelfAssessment("Std18SO19")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std18SO19',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd18SO19" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std18SO19',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std18SO19" style="display:none;">
									 <label for="Std18SO19Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std18SO19Reason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Standard 19/Standard 20 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Standard 19/Standard 20 (sponsoring organization)</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>		
				
				<!-- Professional staff conducts closure interviews -->
				<tr>
					<td align="left" valign="top" class="formMain">Professional staff conducts closure interviews with volunteer, parent/guardian, and child in accordance with the Program Manual </td>
					<td align="left" valign="top" class="formMain">Review program Manual to ensure policies for closure are current and closures are properly documented in case files</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd19SO20" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std19SO20")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std19SO20',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd19SO20" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std19SO20")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std19SO20',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd19SO20" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std19SO20"))) or Trim(GetSelfAssessment("Std19SO20")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std19SO20',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd19SO20" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std19SO20',false)">Not Entered
						<% end if %>							
					</td>
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std19SO20" style="display:none;">
									 <label for="Std19SO20Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std19SO20Reason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Standard 20/Standard 21 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Standard 20/Standard 21 (sponsoring organization)</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>			
				
				<!-- Professional staff reassesses program participants -->
				<tr>
					<td align="left" valign="top" class="formMain">Professional staff reassesses program participants in accordance with Program Manual </td>
					<td align="left" valign="top" class="formMain">Review Program Manual to ensure polices for reassessment are current and reassessments are properly documented in case files</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd20SO21" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std20SO21")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std20SO21',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd20SO21" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std20SO21")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std20SO21',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd20SO21" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std20SO21"))) or Trim(GetSelfAssessment("Std20SO21")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std20SO21',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd20SO21" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std20SO21',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std20SO21" style="display:none;">
									 <label for="Std20SO21Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std20SO21Reason" colspan="3">
							</div>
					</td>
				</tr>

				<!-- Standard 21/Standard 22 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Standard 21/Standard 22 (sponsoring organization)</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>
				
				<!-- Policies and procedures regarding the management of confidential information -->
				<tr>
					<td align="left" valign="top" class="formMain">Written board approved policies and procedures, outlined in the Program Manual, regarding the management of confidential information</td>
					<td align="left" valign="top" class="formMain">Review the board-approved policy and procedures on confidentiality; Verify the consistent application.</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd21SO22" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std21SO22")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std21SO22',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd21SO22" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std21SO22")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std21SO22',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd21SO22" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std21SO22"))) or Trim(GetSelfAssessment("Std21SO22")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std21SO22',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd21SO22" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std21SO22',false)">Not Entered
						<% end if %>							
					</td>										
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std21SO22" style="display:none;">
									 <label for="Std21SO22Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std21SO22Reason" colspan="3">
							</div>
					</td>
				</tr>
				
				<!-- Standard 22/Standard 23 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Standard 22/Standard 23 (sponsoring organization)</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="40%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>				
				
				<!-- Non discrimination policy relative to volunteer Bigs, and Board members -->
				<tr>
					<td align="left" valign="top" class="formMain">Written board approved non discrimination policy relative to volunteer Bigs, and Board members</td>
					<td align="left" valign="top" class="formMain">Document date of Board approval  and verify where policy resides</td>
					<td align="left" valign="top" class="formMain">
						<input type="radio" name="frmSelfAssessmentStd22SO23" value="2"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std22SO23")) = "2" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std22SO23',false)">In
						<br>
						<input type="radio" name="frmSelfAssessmentStd22SO23" value="1"<% If say = "edit" Then %><% If Trim(GetSelfAssessment("Std22SO23")) = "1" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std22SO23',true)">Out
						<br>
						<% if say = "edit" then %>
							<input type="radio" name="frmSelfAssessmentStd22SO23" value="0"<% If say = "edit" Then %><% If isnull(Trim(GetSelfAssessment("Std22SO23"))) or Trim(GetSelfAssessment("Std22SO23")) = "0" Then %> checked<% End If %><% End If %> onclick="disableEnable(this.form,'Std22SO23',false)">Not Entered
						<% else %>
							<input type="radio" name="frmSelfAssessmentStd22SO23" value="0"<% If say <> "edit" Then %> checked<% End If %> onclick="disableEnable(this.form,'Std22SO23',false)">Not Entered
						<% end if %>							
					</td>					
				</tr>								
				<tr>
					<td align="left" valign="top" class="formMain" colspan="3">
							<div id="Std22SO23" style="display:none;">
									 <label for="Std22SO23Reason" style="color: #cc3300;">Please specify reason why you're out of compliance and date you plan to be in: (200 chars max.)</label><br>
									 <input type="text" class="formMain" size="120" value="" name="Std22SO23Reason" colspan="3">
							</div>
					</td>
				</tr>
											
			<!-- End Program Section -->
			<% end if %>	

				<!-- Submit The Form -->
				<tr>
				<%  if section = "Operational" then %>
					<td colspan="2" class="formHeader"><input type="submit" value="Save & Comeback Later" class="formMainBold"></td>
					<td colspan="2" class="formHeader"><input type="button" value="Save & Finish" class="formMainBold" onclick="formvalidation(frmSelfAssessment)"></td>
				<% else %>
					<td colspan="2" class="formHeader"><input type="submit" value="Save & Comeback Later" class="formMainBold"></td>
					<td colspan="2" class="formHeader"><input type="button" value="Save & Finish" class="formMainBold" onclick="formvalidationPr(frmSelfAssessment)"></td>
				<% end if %>
				</tr>

				</table>

			</form>
				
				<br>

<br>
					
<% 
If say = "edit" Then
	GetSelfAssessment.Close
	Set GetSelfAssessment = Nothing
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

