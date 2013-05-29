<% 
If Request("status") = "addNew" Then
	
	' Check for duplicate records
	
	Set DupCon = Server.CreateObject("ADODB.Connection")
	DupCon.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT Count(*) As NumberOfEntries FROM tbl_frmBenefits WHERE AgencyID = '" & Request("AgencyIDN") & "' and Year = " & Request("Year")
response.Write query
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
		RST.Open "SELECT * FROM tbl_frmBenefits", Con, 1, 3
		RST.AddNew
		RST("AgencyID") = Request("AgencyIDN")
		RST("Year") = Request("year")
		
		RST("BenMedOffered") = Request("frmBenefitsBenMedOffered")
		RST("BenMedFullEmployee") = Int(Request("frmBenefitsBenMedFullEmployee"))
		RST("BenMedFullEmployeeAmount") = Int(Request("frmBenefitsBenMedFullEmployeeAmount"))
		RST("BenMedFullFamily") = Int(Request("frmBenefitsBenMedFullFamily"))
		RST("BenMedFullFamilyAmount") = Int(Request("frmBenefitsBenMedFullFamilyAmount"))
		RST("BenMedPartEmployee") = Int(Request("frmBenefitsBenMedPartEmployee"))	
		RST("BenMedPartEmployeeAmount") = Int(Request("frmBenefitsBenMedPartEmployeeAmount"))
		RST("BenMedPartFamily") = Int(Request("frmBenefitsBenMedPartFamily"))
		RST("BenMedPartFamilyAmount") = Int(Request("frmBenefitsBenMedPartFamilyAmount"))
		
		RST("BenDentOffered") = Request("frmBenefitsBenDentOffered")
		RST("BenDentFullEmployee") = Int(Request("frmBenefitsBenDentFullEmployee"))
		RST("BenDentFullEmployeeAmount") = Int(Request("frmBenefitsBenDentFullEmployeeAmount"))
		RST("BenDentFullFamily") = Int(Request("frmBenefitsBenDentFullFamily"))
		RST("BenDentFullFamilyAmount") = Int(Request("frmBenefitsBenDentFullFamilyAmount"))
		RST("BenDentPartEmployee") = Int(Request("frmBenefitsBenDentPartEmployee"))
		RST("BenDentPartEmployeeAmount") = Int(Request("frmBenefitsBenDentPartEmployeeAmount"))
		RST("BenDentPartFamily") = Int(Request("frmBenefitsBenDentPartFamily"))
		RST("BenDentPartFamilyAmount") = Int(Request("frmBenefitsBenDentPartFamilyAmount"))
		
		RST("BenVisOffered") = Request("frmBenefitsBenVisOffered")
		RST("BenVisFullEmployee") = Int(Request("frmBenefitsBenVisFullEmployee"))
		RST("BenVisFullEmployeeAmount") = Int(Request("frmBenefitsBenVisFullEmployeeAmount"))
		RST("BenVisFullFamily") = Int(Request("frmBenefitsBenVisFullFamily"))
		RST("BenVisFullFamilyAmount") = Int(Request("frmBenefitsBenVisFullFamilyAmount"))
		RST("BenVisPartEmployee") = Int(Request("frmBenefitsBenVisPartEmployee"))
		RST("BenVisPartEmployeeAmount") = Int(Request("frmBenefitsBenVisPartEmployeeAmount"))
		RST("BenVisPartFamily") = Int(Request("frmBenefitsBenVisPartFamily"))
		RST("BenVisPartFamilyAmount") = Int(Request("frmBenefitsBenVisPartFamilyAmount"))
		
		RST("DisInsShortTermFull") = Request("frmBenefitsDisInsShortTermFull")
		RST("DisInsShortTermFullPaid") = Request("frmBenefitsDisInsShortTermFullPaid")
		RST("DisInsShortTermFullPrcnt") = Int(Request("frmBenefitsDisInsShortTermFullPrcnt"))
		RST("DisInsShortTermFullAmount") = Int(Request("frmBenefitsDisInsShortTermFullAmount"))
		RST("DisInsShortTermPart") = Request("frmBenefitsDisInsShortTermPart")
		RST("DisInsShortTermPartPaid") = Request("frmBenefitsDisInsShortTermPartPaid")
		RST("DisInsShortTermPartPrcnt") = Int(Request("frmBenefitsDisInsShortTermPartPrcnt"))
		RST("DisInsShortTermPartAmount") = Int(Request("frmBenefitsDisInsShortTermPartAmount"))
		
		RST("DisInsLongTermFull") = Request("frmBenefitsDisInsLongTermFull")
		RST("DisInsLongTermFullPaid") = Request("frmBenefitsDisInsLongTermFullPaid")
		RST("DisInsLongTermFullPrcnt") = Int(Request("frmBenefitsDisInsLongTermFullPrcnt"))
		RST("DisInsLongTermFullAmount") = Int(Request("frmBenefitsDisInsLongTermFullAmount"))
		RST("DisInsLongTermPart") = Request("frmBenefitsDisInsLongTermPart")
		RST("DisInsLongTermPartPaid") = Request("frmBenefitsDisInsLongTermPartPaid")
		RST("DisInsLongTermPartPrcnt") = Int(Request("frmBenefitsDisInsLongTermPartPrcnt"))
		RST("DisInsLongTermPartAmount") = Int(Request("frmBenefitsDisInsLongTermPartAmount"))
		
		RST("EAPFull") = Request("frmBenefitsEAPFull")
		RST("EAPFullPaid") = Request("frmBenefitsEAPFullPaid")
		RST("EAPFullPrcnt") = Int(Request("frmBenefitsEAPFullPrcnt"))
		RST("EAPFullAmount") = Int(Request("frmBenefitsEAPFullAmount"))
		RST("EAPPart") = Request("frmBenefitsEAPPart")
		RST("EAPPartPaid") = Request("frmBenefitsEAPPartPaid")
		RST("EAPPartPrcnt") = Int(Request("frmBenefitsEAPPartPrcnt"))
		RST("EAPPartAmount") = Int(Request("frmBenefitsEAPPartAmount"))
		
		'Commented due to change in requirements (not need to collect this)
		'RST("FlexFull") = Request("frmBenefitsFlexFull")
		'RST("FlexFullPaid") = Request("frmBenefitsFlexFullPaid")
		'RST("FlexFullPrcnt") = Int(Request("frmBenefitsFlexFullPrcnt"))
		'RST("FlexFullAmount") = Int(Request("frmBenefitsFlexFullAmount"))
		'RST("FlexPart") = Request("frmBenefitsFlexPart")
		'RST("FlexPartPaid") = Request("frmBenefitsFlexPartPaid")
		'RST("FlexPartPrcnt") = Int(Request("frmBenefitsFlexPartPrcnt"))
		'RST("FlexPartAmount") = Int(Request("frmBenefitsFlexPartAmount"))
		
		RST("HealthClubFull") = Request("frmBenefitsHealthClubFull")
		RST("HealthClubFullPaid") = Request("frmBenefitsHealthClubFullPaid")
		RST("HealthClubFullPrcnt") = Int(Request("frmBenefitsHealthClubFullPrcnt"))
		RST("HealthClubFullAmount") = Int(Request("frmBenefitsHealthClubFullAmount"))
		RST("HealthClubPart") = Request("frmBenefitsHealthClubPart")
		RST("HealthClubPartPaid") = Request("frmBenefitsHealthClubPartPaid")
		RST("HealthClubPartPrcnt") = Int(Request("frmBenefitsHealthClubPartPrcnt"))
		RST("HealthClubPartAmount") = Int(Request("frmBenefitsHealthClubPartAmount"))
		
		RST("LifeInsuranceFull") = Request("frmBenefitsLifeInsuranceFull")
		RST("LifeInsuranceFullPaid") = Request("frmBenefitsLifeInsuranceFullPaid")
		RST("LifeInsuranceFullPrcnt") = Int(Request("frmBenefitsLifeInsuranceFullPrcnt"))
		RST("LifeInsuranceFullAmount") = Int(Request("frmBenefitsLifeInsuranceFullAmount"))
		RST("LifeInsurancePart") = Request("frmBenefitsLifeInsurancePart")
		RST("LifeInsurancePartPaid") = Request("frmBenefitsLifeInsurancePartPaid")
		RST("LifeInsurancePartPrcnt") = Int(Request("frmBenefitsLifeInsurancePartPrcnt"))
		RST("LifeInsurancePartAmount") = Int(Request("frmBenefitsLifeInsurancePartAmount"))
		
		RST("TimeOffFull") = Request("frmBenefitsTimeOffFull")
		RST("TimeOffFullNExempt") = Request("frmBenefitsTimeOffFullNExempt")
		If LEN(Request("frmBenefitsTimeOffFullDays")) = 0 Then
			RST("TimeOffFullDays") = NULL
		Else
			RST("TimeOffFullDays") = Int(Request("frmBenefitsTimeOffFullDays"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffFullYears")) = 0 Then
			RST("TimeOffFullYears") = NULL
		Else
			RST("TimeOffFullYears") = (Request("frmBenefitsTimeOffFullYears"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffFullDaysIncreased")) = 0 Then
			RST("TimeOffFullDaysIncreased") = NULL
		Else
			RST("TimeOffFullDaysIncreased") = Int(Request("frmBenefitsTimeOffFullDaysIncreased"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffFullDaysNExempt")) = 0 Then
			RST("TimeOffFullDaysNExempt") = NULL
		Else
			RST("TimeOffFullDaysNExempt") = Int(Request("frmBenefitsTimeOffFullDaysNExempt"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffFullYearsNExempt")) = 0 Then
			RST("TimeOffFullYearsNExempt") = NULL
		Else
			RST("TimeOffFullYearsNExempt") = (Request("frmBenefitsTimeOffFullYearsNExempt"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffFullDaysIncreasedNExempt")) = 0 Then
			RST("TimeOffFullDaysIncreasedNExempt") = NULL
		Else
			RST("TimeOffFullDaysIncreasedNExempt") = Int(Request("frmBenefitsTimeOffFullDaysIncreasedNExempt"))
		End If
		
		RST("TimeOffPart") = Request("frmBenefitsTimeOffPart")
		RST("TimeOffPartNExempt") = Request("frmBenefitsTimeOffPartNExempt")
		If LEN(Request("frmBenefitsTimeOffPartDays")) = 0 Then
			RST("TimeOffPartDays") = NULL
		Else
			RST("TimeOffPartDays") = Int(Request("frmBenefitsTimeOffPartDays"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffPartYears")) = 0 Then
			RST("TimeOffPartYears") = NULL
		Else
			RST("TimeOffPartYears") = (Request("frmBenefitsTimeOffPartYears"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffPartDaysIncreased")) = 0 Then
			RST("TimeOffPartDaysIncreased") = NULL
		Else
			RST("TimeOffPartDaysIncreased") = Int(Request("frmBenefitsTimeOffPartDaysIncreased"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffPartDaysNExempt")) = 0 Then
			RST("TimeOffPartDaysNExempt") = NULL
		Else
			RST("TimeOffPartDaysNExempt") = Int(Request("frmBenefitsTimeOffPartDaysNExempt"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffPartYearsNExempt")) = 0 Then
			RST("TimeOffPartYearsNExempt") = NULL
		Else
			RST("TimeOffPartYearsNExempt") = (Request("frmBenefitsTimeOffPartYearsNExempt"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffPartDaysIncreasedNExempt")) = 0 Then
			RST("TimeOffPartDaysIncreasedNExempt") = NULL
		Else
			RST("TimeOffPartDaysIncreasedNExempt") = Int(Request("frmBenefitsTimeOffPartDaysIncreasedNExempt"))
		End If
		
		RST("TimeOffSickFull") = Request("frmBenefitsTimeOffSickFull")
		If LEN(Request("frmBenefitsTimeOffSickFullDays")) = 0 Then
			RST("TimeOffSickFullDays") = NULL
		Else
			RST("TimeOffSickFullDays") = Int(Request("frmBenefitsTimeOffSickFullDays"))
		End If
		RST("TimeOffSickPart") = Request("frmBenefitsTimeOffSickPart")
		If LEN(Request("frmBenefitsTimeOffSickPartDays")) = 0 Then
			RST("TimeOffSickPartDays") = NULL
		Else
			RST("TimeOffSickPartDays") = Int(Request("frmBenefitsTimeOffSickPartDays"))
		End If
		
		'RST("TimeOffVacFull") = Request("frmBenefitsTimeOffVacFull")
		'RST("TimeOffVacFullDays") = Int(Request("frmBenefitsTimeOffVacFullDays"))
		'RST("TimeOffVacPart") = Request("frmBenefitsTimeOffVacPart")
		'RST("TimeOffVacPartDays") = Int(Request("frmBenefitsTimeOffVacPartDays"))
		
		RST("ProfDuesFull") = Request("frmBenefitsProfDuesFull")
		RST("ProfDuesFullPaid") = Request("frmBenefitsProfDuesFullPaid")
		If LEN(Request("frmBenefitsProfDuesFullAmount")) = 0 Then
			RST("ProfDuesFullAmount") = NULL
		Else
			RST("ProfDuesFullAmount") = Int(Request("frmBenefitsProfDuesFullAmount"))
		End If
		RST("ProfDuesPart") = Request("frmBenefitsProfDuesPart")
		RST("ProfDuesPartPaid") = Request("frmBenefitsProfDuesPartPaid")
		If LEN(Request("frmBenefitsProfDuesPartAmount")) = 0 Then
			RST("ProfDuesPartAmount") = NULL
		Else
			RST("ProfDuesPartAmount") = Int(Request("frmBenefitsProfDuesPartAmount"))
		End If
						
		RST("RetirementFull") = Request("frmBenefitsRetirementFull")
		'RST("RetirementFullPaid") = Request("frmBenefitsRetirementFullPaid")
		If LEN(Request("frmBenefitsRetirementFullPrcnt")) = 0 Then
			RST("RetirementFullPrcnt") = NULL
		Else
			RST("RetirementFullPrcnt") = Int(Request("frmBenefitsRetirementFullPrcnt"))
		End If
		RST("RetirementPart") = Request("frmBenefitsRetirementPart")
		'RST("RetirementPartPaid") = Request("frmBenefitsRetirementPartPaid")
		If LEN(Request("frmBenefitsRetirementPartPrcnt")) = 0 Then
			RST("RetirementPartPrcnt") = NULL
		Else
			RST("RetirementPartPrcnt") = Int(Request("frmBenefitsRetirementPartPrcnt"))
		End If

		RST("403BFull") = Request("frmBenefits403BFull")
		RST("403BFullContrib") = Request("frmBenefits403BFullContrib")
		If LEN(Request("frmBenefits403BFullContribPrcnt")) = 0 Then 
		    RST("403BFullContribPrcnt") = NULL
		Else
			RST("403BFullContribPrcnt") = Int(Request("frmBenefits403BFullContribPrcnt"))
		End If

		RST("403BPart") = Request("frmBenefits403BPart")
		RST("403BPartContrib") = Request("frmBenefits403BPartContrib")
		If LEN(Request("frmBenefits403BPartContribPrcnt")) = 0 Then 
		    RST("403BPartContribPrcnt") = NULL
		Else
			RST("403BPartContribPrcnt") = Int(Request("frmBenefits403BPartContribPrcnt"))
		End If
								
		RST("TelecommFull") = Request("frmBenefitsTelecommFull")
		'RST("TelecommFullPaid") = Request("frmBenefitsTelecommFullPaid")
		If LEN(Request("frmBenefitsTelecommFullCount")) = 0 Then
			RST("TelecommFullCount") = NULL
		Else
			RST("TelecommFullCount") = Int(Request("frmBenefitsTelecommFullCount"))
		End If
		If LEN(Request("frmBenefitsTelecommFullPrcnt")) = 0 Then
			RST("TelecommFullPrcnt") = NULL
		Else
			RST("TelecommFullPrcnt") = Int(Request("frmBenefitsTelecommFullPrcnt"))
		End If
		RST("TelecommPart") = Request("frmBenefitsTelecommPart")
		'RST("TelecommPartPaid") = Request("frmBenefitsTelecommPartPaid")
		If LEN(Request("frmBenefitsTelecommPartCount")) = 0 Then
			RST("TelecommPartCount") = NULL
		Else
			RST("TelecommPartCount") = Int(Request("frmBenefitsTelecommPartCount"))
		End If
		If LEN(Request("frmBenefitsTelecommPartPrcnt")) = 0 Then
			RST("TelecommPartPrcnt") = NULL
		Else
			RST("TelecommPartPrcnt") = Int(Request("frmBenefitsTelecommPartPrcnt"))
		End If
		
		RST("TuitionFull") = Request("frmBenefitsTuitionFull")
		'RST("TuitionFullPaid") = Request("frmBenefitsTuitionFullPaid")
		If LEN(Request("frmBenefitsTuitionFullAmount")) = 0 Then
			RST("TuitionFullAmount") = NULL
		Else
			RST("TuitionFullAmount") = Int(Request("frmBenefitsTuitionFullAmount"))
		End If
		RST("TuitionPart") = Request("frmBenefitsTuitionPart")
		'RST("TuitionPartPaid") = Request("frmBenefitsTuitionPartPaid")
		If LEN(Request("frmBenefitsTuitionPartAmount")) = 0 Then
			RST("TuitionPartAmount") = NULL
		Else
			RST("TuitionPartAmount") = Int(Request("frmBenefitsTuitionPartAmount"))
		End If
		
		RST("CreateDate") = Now
		RST.Update
		RST.Close
		Set RST = Nothing
		form = "Benefits"
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
	RST.Open "SELECT * FROM tbl_frmBenefits WHERE AgencyID = '" & Request("AgencyIDN") & "' AND Year=" & Int(Request("year")), Con, 1, 3

		
		RST("BenMedOffered") = Request("frmBenefitsBenMedOffered")
		RST("BenMedFullEmployee") = Int(Request("frmBenefitsBenMedFullEmployee"))
		RST("BenMedFullEmployeeAmount") = Int(Request("frmBenefitsBenMedFullEmployeeAmount"))
		RST("BenMedFullFamily") = Int(Request("frmBenefitsBenMedFullFamily"))
		RST("BenMedFullFamilyAmount") = Int(Request("frmBenefitsBenMedFullFamilyAmount"))
		RST("BenMedPartEmployee") = Int(Request("frmBenefitsBenMedPartEmployee"))	
		RST("BenMedPartEmployeeAmount") = Int(Request("frmBenefitsBenMedPartEmployeeAmount"))
		RST("BenMedPartFamily") = Int(Request("frmBenefitsBenMedPartFamily"))
		RST("BenMedPartFamilyAmount") = Int(Request("frmBenefitsBenMedPartFamilyAmount"))
		
		RST("BenDentOffered") = Request("frmBenefitsBenDentOffered")
		RST("BenDentFullEmployee") = Int(Request("frmBenefitsBenDentFullEmployee"))
		RST("BenDentFullEmployeeAmount") = Int(Request("frmBenefitsBenDentFullEmployeeAmount"))
		RST("BenDentFullFamily") = Int(Request("frmBenefitsBenDentFullFamily"))
		RST("BenDentFullFamilyAmount") = Int(Request("frmBenefitsBenDentFullFamilyAmount"))
		RST("BenDentPartEmployee") = Int(Request("frmBenefitsBenDentPartEmployee"))
		RST("BenDentPartEmployeeAmount") = Int(Request("frmBenefitsBenDentPartEmployeeAmount"))
		RST("BenDentPartFamily") = Int(Request("frmBenefitsBenDentPartFamily"))
		RST("BenDentPartFamilyAmount") = Int(Request("frmBenefitsBenDentPartFamilyAmount"))
		
		RST("BenVisOffered") = Request("frmBenefitsBenVisOffered")
		RST("BenVisFullEmployee") = Int(Request("frmBenefitsBenVisFullEmployee"))
		RST("BenVisFullEmployeeAmount") = Int(Request("frmBenefitsBenVisFullEmployeeAmount"))
		RST("BenVisFullFamily") = Int(Request("frmBenefitsBenVisFullFamily"))
		RST("BenVisFullFamilyAmount") = Int(Request("frmBenefitsBenVisFullFamilyAmount"))
		RST("BenVisPartEmployee") = Int(Request("frmBenefitsBenVisPartEmployee"))
		RST("BenVisPartEmployeeAmount") = Int(Request("frmBenefitsBenVisPartEmployeeAmount"))
		RST("BenVisPartFamily") = Int(Request("frmBenefitsBenVisPartFamily"))
		RST("BenVisPartFamilyAmount") = Int(Request("frmBenefitsBenVisPartFamilyAmount"))
		
		RST("DisInsShortTermFull") = Request("frmBenefitsDisInsShortTermFull")
		RST("DisInsShortTermFullPaid") = Request("frmBenefitsDisInsShortTermFullPaid")
		RST("DisInsShortTermFullPrcnt") = Int(Request("frmBenefitsDisInsShortTermFullPrcnt"))
		RST("DisInsShortTermFullAmount") = Int(Request("frmBenefitsDisInsShortTermFullAmount"))
		RST("DisInsShortTermPart") = Request("frmBenefitsDisInsShortTermPart")
		RST("DisInsShortTermPartPaid") = Request("frmBenefitsDisInsShortTermPartPaid")
		RST("DisInsShortTermPartPrcnt") = Int(Request("frmBenefitsDisInsShortTermPartPrcnt"))
		RST("DisInsShortTermPartAmount") = Int(Request("frmBenefitsDisInsShortTermPartAmount"))
		
		RST("DisInsLongTermFull") = Request("frmBenefitsDisInsLongTermFull")
		RST("DisInsLongTermFullPaid") = Request("frmBenefitsDisInsLongTermFullPaid")
		RST("DisInsLongTermFullPrcnt") = Int(Request("frmBenefitsDisInsLongTermFullPrcnt"))
		RST("DisInsLongTermFullAmount") = Int(Request("frmBenefitsDisInsLongTermFullAmount"))
		RST("DisInsLongTermPart") = Request("frmBenefitsDisInsLongTermPart")
		RST("DisInsLongTermPartPaid") = Request("frmBenefitsDisInsLongTermPartPaid")
		RST("DisInsLongTermPartPrcnt") = Int(Request("frmBenefitsDisInsLongTermPartPrcnt"))
		RST("DisInsLongTermPartAmount") = Int(Request("frmBenefitsDisInsLongTermPartAmount"))
		
		RST("EAPFull") = Request("frmBenefitsEAPFull")
		RST("EAPFullPaid") = Request("frmBenefitsEAPFullPaid")
		RST("EAPFullPrcnt") = Int(Request("frmBenefitsEAPFullPrcnt"))
		RST("EAPFullAmount") = Int(Request("frmBenefitsEAPFullAmount"))
		RST("EAPPart") = Request("frmBenefitsEAPPart")
		RST("EAPPartPaid") = Request("frmBenefitsEAPPartPaid")
		RST("EAPPartPrcnt") = Int(Request("frmBenefitsEAPPartPrcnt"))
		RST("EAPPartAmount") = Int(Request("frmBenefitsEAPPartAmount"))
		
		'Commented due to change in requirements (not need to collect this)
		'RST("FlexFull") = Request("frmBenefitsFlexFull")
		'RST("FlexFullPaid") = Request("frmBenefitsFlexFullPaid")
		'RST("FlexFullPrcnt") = Int(Request("frmBenefitsFlexFullPrcnt"))
		'RST("FlexFullAmount") = Int(Request("frmBenefitsFlexFullAmount"))
		'RST("FlexPart") = Request("frmBenefitsFlexPart")
		'RST("FlexPartPaid") = Request("frmBenefitsFlexPartPaid")
		'RST("FlexPartPrcnt") = Int(Request("frmBenefitsFlexPartPrcnt"))
		'RST("FlexPartAmount") = Int(Request("frmBenefitsFlexPartAmount"))
		
		RST("HealthClubFull") = Request("frmBenefitsHealthClubFull")
		RST("HealthClubFullPaid") = Request("frmBenefitsHealthClubFullPaid")
		RST("HealthClubFullPrcnt") = Int(Request("frmBenefitsHealthClubFullPrcnt"))
		RST("HealthClubFullAmount") = Int(Request("frmBenefitsHealthClubFullAmount"))
		RST("HealthClubPart") = Request("frmBenefitsHealthClubPart")
		RST("HealthClubPartPaid") = Request("frmBenefitsHealthClubPartPaid")
		RST("HealthClubPartPrcnt") = Int(Request("frmBenefitsHealthClubPartPrcnt"))
		RST("HealthClubPartAmount") = Int(Request("frmBenefitsHealthClubPartAmount"))
		
		RST("LifeInsuranceFull") = Request("frmBenefitsLifeInsuranceFull")
		RST("LifeInsuranceFullPaid") = Request("frmBenefitsLifeInsuranceFullPaid")
		RST("LifeInsuranceFullPrcnt") = Int(Request("frmBenefitsLifeInsuranceFullPrcnt"))
		RST("LifeInsuranceFullAmount") = Int(Request("frmBenefitsLifeInsuranceFullAmount"))
		RST("LifeInsurancePart") = Request("frmBenefitsLifeInsurancePart")
		RST("LifeInsurancePartPaid") = Request("frmBenefitsLifeInsurancePartPaid")
		RST("LifeInsurancePartPrcnt") = Int(Request("frmBenefitsLifeInsurancePartPrcnt"))
		RST("LifeInsurancePartAmount") = Int(Request("frmBenefitsLifeInsurancePartAmount"))
		
		RST("TimeOffFull") = Request("frmBenefitsTimeOffFull")
		RST("TimeOffFullNExempt") = Request("frmBenefitsTimeOffFullNExempt")
		If LEN(Request("frmBenefitsTimeOffFullDays")) = 0 Then
			RST("TimeOffFullDays") = NULL
		Else
			RST("TimeOffFullDays") = Int(Request("frmBenefitsTimeOffFullDays"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffFullYears")) = 0 Then
			RST("TimeOffFullYears") = NULL
		Else
			RST("TimeOffFullYears") = (Request("frmBenefitsTimeOffFullYears"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffFullDaysIncreased")) = 0 Then
			RST("TimeOffFullDaysIncreased") = NULL
		Else
			RST("TimeOffFullDaysIncreased") = Int(Request("frmBenefitsTimeOffFullDaysIncreased"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffFullDaysNExempt")) = 0 Then
			RST("TimeOffFullDaysNExempt") = NULL
		Else
			RST("TimeOffFullDaysNExempt") = Int(Request("frmBenefitsTimeOffFullDaysNExempt"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffFullYearsNExempt")) = 0 Then
			RST("TimeOffFullYearsNExempt") = NULL
		Else
			RST("TimeOffFullYearsNExempt") = (Request("frmBenefitsTimeOffFullYearsNExempt"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffFullDaysIncreasedNExempt")) = 0 Then
			RST("TimeOffFullDaysIncreasedNExempt") = NULL
		Else
			RST("TimeOffFullDaysIncreasedNExempt") = Int(Request("frmBenefitsTimeOffFullDaysIncreasedNExempt"))
		End If
		
		RST("TimeOffPart") = Request("frmBenefitsTimeOffPart")
		RST("TimeOffPartNExempt") = Request("frmBenefitsTimeOffPartNExempt")
		If LEN(Request("frmBenefitsTimeOffPartDays")) = 0 Then
			RST("TimeOffPartDays") = NULL
		Else
			RST("TimeOffPartDays") = Int(Request("frmBenefitsTimeOffPartDays"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffPartYears")) = 0 Then
			RST("TimeOffPartYears") = NULL
		Else
			RST("TimeOffPartYears") = (Request("frmBenefitsTimeOffPartYears"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffPartDaysIncreased")) = 0 Then
			RST("TimeOffPartDaysIncreased") = NULL
		Else
			RST("TimeOffPartDaysIncreased") = Int(Request("frmBenefitsTimeOffPartDaysIncreased"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffPartDaysNExempt")) = 0 Then
			RST("TimeOffPartDaysNExempt") = NULL
		Else
			RST("TimeOffPartDaysNExempt") = Int(Request("frmBenefitsTimeOffPartDaysNExempt"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffPartYearsNExempt")) = 0 Then
			RST("TimeOffPartYearsNExempt") = NULL
		Else
			RST("TimeOffPartYearsNExempt") = (Request("frmBenefitsTimeOffPartYearsNExempt"))
		End If
		
		If LEN(Request("frmBenefitsTimeOffPartDaysIncreasedNExempt")) = 0 Then
			RST("TimeOffPartDaysIncreasedNExempt") = NULL
		Else
			RST("TimeOffPartDaysIncreasedNExempt") = Int(Request("frmBenefitsTimeOffPartDaysIncreasedNExempt"))
		End If
		
		RST("TimeOffSickFull") = Request("frmBenefitsTimeOffSickFull")
		RST("TimeOffSickFullDays") = Int(Request("frmBenefitsTimeOffSickFullDays"))
		RST("TimeOffSickPart") = Request("frmBenefitsTimeOffSickPart")
		RST("TimeOffSickPartDays") = Int(Request("frmBenefitsTimeOffSickPartDays"))
		
		'RST("TimeOffVacFull") = Request("frmBenefitsTimeOffVacFull")
		'RST("TimeOffVacFullDays") = Int(Request("frmBenefitsTimeOffVacFullDays"))
		'RST("TimeOffVacPart") = Request("frmBenefitsTimeOffVacPart")
		'RST("TimeOffVacPartDays") = Int(Request("frmBenefitsTimeOffVacPartDays"))
		
		RST("ProfDuesFull") = Request("frmBenefitsProfDuesFull")
		RST("ProfDuesFullPaid") = Request("frmBenefitsProfDuesFullPaid")
		RST("ProfDuesFullAmount") = Int(Request("frmBenefitsProfDuesFullAmount"))
		RST("ProfDuesPart") = Request("frmBenefitsProfDuesPart")
		RST("ProfDuesPartPaid") = Request("frmBenefitsProfDuesPartPaid")
		RST("ProfDuesPartAmount") = Int(Request("frmBenefitsProfDuesPartAmount"))
						
		RST("RetirementFull") = Request("frmBenefitsRetirementFull")
		'RST("RetirementFullPaid") = Request("frmBenefitsRetirementFullPaid")
		If LEN(Request("frmBenefitsRetirementFullPrcnt")) = 0 Then
			RST("RetirementFullPrcnt") = NULL
		Else
			RST("RetirementFullPrcnt") = Int(Request("frmBenefitsRetirementFullPrcnt"))
		End If
		RST("RetirementPart") = Request("frmBenefitsRetirementPart")
		'RST("RetirementPartPaid") = Request("frmBenefitsRetirementPartPaid")
		If LEN(Request("frmBenefitsRetirementPartPrcnt")) = 0 Then
			RST("RetirementPartPrcnt") = NULL
		Else
			RST("RetirementPartPrcnt") = Int(Request("frmBenefitsRetirementPartPrcnt"))
		End If

		RST("403BFull") = Request("frmBenefits403BFull")
		RST("403BFullContrib") = Request("frmBenefits403BFullContrib")
		If LEN(Request("frmBenefits403BFullContribPrcnt")) = 0 Then 
		    RST("403BFullContribPrcnt") = NULL
		Else
			RST("403BFullContribPrcnt") = Int(Request("frmBenefits403BFullContribPrcnt"))
		End If

		RST("403BPart") = Request("frmBenefits403BPart")
		RST("403BPartContrib") = Request("frmBenefits403BPartContrib")
		If LEN(Request("frmBenefits403BPartContribPrcnt")) = 0 Then 
		    RST("403BPartContribPrcnt") = NULL
		Else
			RST("403BPartContribPrcnt") = Int(Request("frmBenefits403BPartContribPrcnt"))
		End If
		
		RST("TelecommFull") = Request("frmBenefitsTelecommFull")
		'RST("TelecommFullPaid") = Request("frmBenefitsTelecommFullPaid")
		RST("TelecommFullCount") = Int(Request("frmBenefitsTelecommFullCount"))
		RST("TelecommFullPrcnt") = Int(Request("frmBenefitsTelecommFullPrcnt"))
		RST("TelecommPart") = Request("frmBenefitsTelecommPart")
		'RST("TelecommPartPaid") = Request("frmBenefitsTelecommPartPaid")
		RST("TelecommPartCount") = Int(Request("frmBenefitsTelecommPartCount"))
		RST("TelecommPartPrcnt") = Int(Request("frmBenefitsTelecommPartPrcnt"))
		
		RST("TuitionFull") = Request("frmBenefitsTuitionFull")
		'RST("TuitionFullPaid") = Request("frmBenefitsTuitionFullPaid")
		RST("TuitionFullAmount") = Int(Request("frmBenefitsTuitionFullAmount"))
		RST("TuitionPart") = Request("frmBenefitsTuitionPart")
		'RST("TuitionPartPaid") = Request("frmBenefitsTuitionPartPaid")
		RST("TuitionPartAmount") = Int(Request("frmBenefitsTuitionPartAmount"))
		
	jMod = RST("BenefitsID")
	RST.Update
	RST.Close
	Set RST = Nothing
	form = "Benefits"
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
	<title>Benefits</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
<script language="JavaScript">
<!--

function NewWindow(mypage, myname, w, h)
{
	var winl = (screen.width - w) / 2;
	var wint = (screen.height - h) / 2;
	winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',resizable, scrollbars'
	win = window.open(mypage, myname, winprops)
	if (parseInt(navigator.appVersion) >= 4) { win.window.focus(); }
}

function addEmUp() 
{
	//var box1 = Number(document.frmBenefits.frmBenefitsSalariesWages.value)
	//var box2 = Number(document.frmBenefits.frmBenefitsEmployeeBenefits.value)
	//var box3 = Number(document.frmBenefits.frmBenefitsInsurance.value)	
	//var box4 = Number(document.frmBenefits.frmBenefitsOther.value)
	//var boxtotal = box1 + box2 + box3 + box4
	//document.frmBenefits.frmBenefitsTotal.value = boxtotal
}


function AddUpAssets()
{
	//var box1 = Number(document.frmBenefits.frmBenefitsCashInvestments.value)
	//var box2 = Number(document.frmBenefits.frmBenefitsReceivables.value)
	//var box3 = Number(document.frmBenefits.frmBenefitsAllOtherAssets.value)	

	//var boxtotal = box1 + box2 + box3
	//document.frmBenefits.frmBenefitsTotalAssets.value = boxtotal
	
}

function AddUpLiabilities()
{
	//var box1 = Number(document.frmBenefits.frmBenefitsLiabilities.value)
	//var box2 = Number(document.frmBenefits.frmBenefitsSurplus_NetAssets.value)

	//var boxtotal = box1 + box2
	//document.frmBenefits.frmBenefitsTotalLiabNetAssets.value = boxtotal
	
}



function addUpCategory()
{
	//var box1 = Number(document.frmBenefits.frmBenefitsAdministration.value)
	//var box2 = Number(document.frmBenefits.frmBenefitsProgram.value)
	//var box3 = Number(document.frmBenefits.frmBenefitsFundRaising.value)	

	//var boxtotal = box1 + box2 + box3
	//document.frmBenefits.frmBenefitsCategoryTotal.value = boxtotal
	
}

function addUpExpenseMentNonMent()
{
	//var box1 = Number(document.frmBenefits.frmBenefitsTotalExpenseMentoring.value)
	//var box2 = Number(document.frmBenefits.frmBenefitsTotalExpenseNonMentoring.value)

	//var boxtotal = box1 + box2
	//document.frmBenefits.frmBenefitsBenefitsMentNonMentTotal.value = boxtotal
	
}

function addUpFTEProgram()
{
	//var box1 = Number(document.frmBenefits.frmBenefitsFTECommunity.value)
	//var box2 = Number(document.frmBenefits.frmBenefitsFTESchool.value)
	//var box3 = Number(document.frmBenefits.frmBenefitsFTESite.value)	

	//var boxtotal = box1 + box2 + box3
	//document.frmBenefits.frmBenefitsFTEProgramTotal.value = boxtotal
	
}

function addUpFTEFunction()
{
	//var box1 = Number(document.frmBenefits.frmBenefitsFTECustomerRelations.value)
	//var box2 = Number(document.frmBenefits.frmBenefitsFTEEnrollmentMatching.value)
	//var box3 = Number(document.frmBenefits.frmBenefitsFTEMatchSupport.value)	

	//var boxtotal = box1 + box2 + box3
	//document.frmBenefits.frmBenefitsFTEFunctionTotal.value = boxtotal
	
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
	//Medical
	if(form.frmBenefitsBenMedOffered.value == "selected")	
	{
		form.frmBenefitsBenMedOffered.focus();
		alert("Please select whether your agency offer medical benefits.");
		return false;
	}
	else if(form.frmBenefitsBenDentOffered.value == "selected")	
	{
		form.frmBenefitsBenDentOffered.focus();
		alert("Please select whether your agency offer dental benefits.");
		return false;
	}
	else if(form.frmBenefitsBenVisOffered.value == "selected")	
	{
		form.frmBenefitsBenVisOffered.focus();
		alert("Please select whether your agency offer vision benefits.");
		return false;
	}
	
	//medical benefit
	else if ((form.frmBenefitsBenMedOffered.value == "true") && (form.frmBenefitsBenMedFullEmployee.value == 0) && (form.frmBenefitsBenMedFullEmployeeAmount.value == 0) && (form.frmBenefitsBenMedFullFamily.value == 0) && (form.frmBenefitsBenMedFullFamilyAmount.value == 0) && (form.frmBenefitsBenMedPartEmployee.value == 0) && (form.frmBenefitsBenMedPartEmployeeAmount.value == 0) && (form.frmBenefitsBenMedPartFamily.value == 0) && (form.frmBenefitsBenMedPartFamilyAmount.value == 0))
	{
		alert("You selected that you offer Medical Benefits, but did not enter Percentages and Amounts. Please enter required data or select that you do NOT offer Medical Benefits");
		form.frmBenefitsBenMedFullEmployee.focus();
		return false;
	}

// Medical Full Time Employee Validation(1)
	else if ((form.frmBenefitsBenMedFullEmployee.value != 0) && (form.frmBenefitsBenMedFullEmployeeAmount.value == 0))
	{
		alert("Please enter the total monthly premium or set the percentage paid by agency to 0.");
		form.frmBenefitsBenMedFullEmployee.focus();
		return false;
	}

//Medical Full Time Family Validation(2)
   else if ((form.frmBenefitsBenMedFullFamily.value != 0) && (form.frmBenefitsBenMedFullFamilyAmount.value == 0))
	{
		alert("Please enter the total monthly premium or set the percentage paid by agency to 0.");
		form.frmBenefitsBenMedFullFamily.focus();
		return false;
	}

//Medical Part Time Employee Validation(3)
   else if ((form.frmBenefitsBenMedPartEmployee.value != 0) && (form.frmBenefitsBenMedPartEmployeeAmount.value == 0))
	{
		alert("Please enter the total monthly premium or set the percentage paid by agency to 0.");
		form.frmBenefitsBenMedPartEmployee.focus();
		return false;
	}

//Medical Part Time Family Validation(4)

    else if ((form.frmBenefitsBenMedPartFamily.value != 0) && (form.frmBenefitsBenMedPartFamilyAmount.value == 0))
	{
		alert("Please enter the total monthly premium or set the percentage paid by agency to 0.");
		form.frmBenefitsBenMedPartFamily.focus();
		return false;
	}

	else if((!(myRegularExpression3.test(form.frmBenefitsBenMedFullEmployee.value))) || (form.frmBenefitsBenMedFullEmployee.value > 100) || (form.frmBenefitsBenMedFullEmployee.value < 0))
	{
		
		alert((form.frmBenefitsBenMedFullEmployee.value) + " is invalid. Please enter a whole number between 0 and 100.");
		form.frmBenefitsBenMedFullEmployee.focus();
		return false;
	}
	else if(!(myRegularExpression3.test(form.frmBenefitsBenMedFullEmployeeAmount.value)))	
	{
		form.frmBenefitsBenMedFullEmployeeAmount.focus();
		alert((form.frmBenefitsBenMedFullEmployeeAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}						
	
	else if((!(myRegularExpression3.test(form.frmBenefitsBenMedFullFamily.value))) || (form.frmBenefitsBenMedFullFamily.value > 100) || (form.frmBenefitsBenMedFullFamily.value < 0))
	{
		form.frmBenefitsBenMedFullFamily.focus();
		alert((form.frmBenefitsBenMedFullFamily.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression3.test(form.frmBenefitsBenMedFullFamilyAmount.value)))	
	{
		form.frmBenefitsBenMedFullFamilyAmount.focus();
		alert((form.frmBenefitsBenMedFullFamilyAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}						
	
	else if((!(myRegularExpression3.test(form.frmBenefitsBenMedPartEmployee.value))) || (form.frmBenefitsBenMedPartEmployee.value > 100) || (form.frmBenefitsBenMedPartEmployee.value < 0))
	{
		form.frmBenefitsBenMedPartEmployee.focus();
		alert((form.frmBenefitsBenMedPartEmployee.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression3.test(form.frmBenefitsBenMedPartEmployeeAmount.value)))	
	{
		form.frmBenefitsBenMedPartEmployeeAmount.focus();
		alert((form.frmBenefitsBenMedPartEmployeeAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}						
	
	else if((!(myRegularExpression3.test(form.frmBenefitsBenMedPartFamily.value))) || (form.frmBenefitsBenMedPartFamily.value > 100) || (form.frmBenefitsBenMedPartFamily.value < 0))
	{
		form.frmBenefitsBenMedPartFamily.focus();
		alert((form.frmBenefitsBenMedPartFamily.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression3.test(form.frmBenefitsBenMedPartFamilyAmount.value)))	
	{
		form.frmBenefitsBenMedPartFamilyAmount.focus();
		alert((form.frmBenefitsBenMedPartFamilyAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
//dental benefits
/* Commented on 10/19/2007 due to request for allowing unknown amount for vision if it is part of medical plan.
	else if ((form.frmBenefitsBenDentOffered.value == "true") && (form.frmBenefitsBenDentFullEmployee.value == 0) && (form.frmBenefitsBenDentFullEmployeeAmount.value == 0) && (form.frmBenefitsBenDentFullFamily.value == 0) && (form.frmBenefitsBenDentFullFamilyAmount.value == 0) && (form.frmBenefitsBenDentPartEmployee.value == 0) && (form.frmBenefitsBenDentPartEmployeeAmount.value == 0) && (form.frmBenefitsBenDentPartFamily.value == 0) && (form.frmBenefitsBenDentPartFamilyAmount.value == 0))
	{
		alert("You selected that you offer Dental Benefits, but did not entered Percentages and Amounts. Please enter required data or select that you do NOT offer Dental Benefits");
		form.frmBenefitsBenDentFullEmployee.focus();
		return false;
	}
*/	
	
// Dental Full Time Employee Validation(1)
	else if ((form.frmBenefitsBenDentFullEmployee.value != 0) && (form.frmBenefitsBenDentFullEmployeeAmount.value == 0))
	{
		alert("Please enter the total monthly premium or set the percentage paid by agency to 0.");
		form.frmBenefitsBenDentFullEmployee.focus();
		return false;
	}

//Dental Full Time Employee Family Validation(2)
   else if ((form.frmBenefitsBenDentFullFamily.value != 0) && (form.frmBenefitsBenDentFullFamilyAmount.value == 0))
	{
		alert("Please enter the total monthly premium or set the percentage paid by agency to 0.");
		form.frmBenefitsBenDentFullFamily.focus();
		return false;
	}

//Dental Part Time Employee Validation(3)
   else if ((form.frmBenefitsBenDentPartEmployee.value != 0) && (form.frmBenefitsBenDentPartEmployeeAmount.value == 0))
	{
		alert("Please enter the total monthly premium or set the percentage paid by agency to 0.");
		form.frmBenefitsBenDentPartEmployee.focus();
		return false;
	}

//Dental Part Time Family Validation(4)

    else if ((form.frmBenefitsBenDentPartFamily.value != 0) && (form.frmBenefitsBenDentPartFamilyAmount.value == 0))
	{
		alert("Please enter the total monthly premium or set the percentage paid by agency to 0.");
		form.frmBenefitsBenDentPartFamily.focus();
		return false;
	}
	
	
	else if((!(myRegularExpression3.test(form.frmBenefitsBenDentFullEmployee.value))) || (form.frmBenefitsBenDentFullEmployee.value > 100) || (form.frmBenefitsBenDentFullEmployee.value < 0))
	{
		form.frmBenefitsBenDentFullEmployee.focus();
		alert((form.frmBenefitsBenDentFullEmployee.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression3.test(form.frmBenefitsBenDentFullEmployeeAmount.value)))	
	{
		form.frmBenefitsBenDentFullEmployeeAmount.focus();
		alert((form.frmBenefitsBenDentFullEmployeeAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
	else if((!(myRegularExpression3.test(form.frmBenefitsBenDentFullFamily.value))) || (form.frmBenefitsBenDentFullFamily.value > 100) || (form.frmBenefitsBenDentFullFamily.value < 0))
	{
		form.frmBenefitsBenDentFullFamily.focus();
		alert((form.frmBenefitsBenDentFullFamily.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression3.test(form.frmBenefitsBenDentFullFamilyAmount.value)))	
	{
		form.frmBenefitsBenDentFullFamilyAmount.focus();
		alert((form.frmBenefitsBenDentFullFamilyAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
	else if((!(myRegularExpression3.test(form.frmBenefitsBenDentPartEmployee.value))) || (form.frmBenefitsBenDentPartEmployee.value > 100) || (form.frmBenefitsBenDentPartEmployee.value < 0))
	{
		form.frmBenefitsBenDentPartEmployee.focus();
		alert((form.frmBenefitsBenDentPartEmployee.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression3.test(form.frmBenefitsBenDentPartEmployeeAmount.value)))	
	{
		form.frmBenefitsBenDentPartEmployeeAmount.focus();
		alert((form.frmBenefitsBenDentPartEmployeeAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
	else if((!(myRegularExpression3.test(form.frmBenefitsBenDentPartFamily.value))) || (form.frmBenefitsBenDentPartFamily.value > 100) || (form.frmBenefitsBenDentPartFamily.value < 0))
	{
		form.frmBenefitsBenDentPartFamily.focus();
		alert((form.frmBenefitsBenDentPartFamily.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression3.test(form.frmBenefitsBenDentPartFamilyAmount.value)))	
	{
		form.frmBenefitsBenDentPartFamilyAmount.focus();
		alert((form.frmBenefitsBenDentPartFamilyAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
// vision benefits
/* Commented on 10/17/2007 due to request for allowing unknown amounts for vision if it is part of medical plan.
	else if ((form.frmBenefitsBenVisOffered.value == "true") && (form.frmBenefitsBenVisFullEmployee.value == 0) && (form.frmBenefitsBenVisFullEmployeeAmount.value == 0) && (form.frmBenefitsBenVisFullFamily.value == 0) && (form.frmBenefitsBenVisFullFamilyAmount.value == 0) && (form.frmBenefitsBenVisPartEmployee.value == 0) && (form.frmBenefitsBenVisPartEmployeeAmount.value == 0) && (form.frmBenefitsBenVisPartFamily.value == 0) && (form.frmBenefitsBenVisPartFamilyAmount.value == 0))
	{
		alert("You selected that you offer Vision Benefits, but did not entered Percentages and Amounts. Please enter required data or select that you do NOT offer Vision Benefits");
		form.frmBenefitsBenVisFullEmployee.focus();
		return false;
	}
*/
	
	
// Vision Full Time Employee Validation(1)
	else if ((form.frmBenefitsBenVisFullEmployee.value != 0) && (form.frmBenefitsBenVisFullEmployeeAmount.value == 0)) 
	{
		alert("Please enter the total monthly premium or set the percentage paid by agency to 0.");
		form.frmBenefitsBenVisFullEmployee.focus();
		return false;
	}


//Vision Full Time Employee Family Validation(2)
   else if ((form.frmBenefitsBenVisFullFamily.value != 0) && (form.frmBenefitsBenVisFullFamilyAmount.value == 0))
	{
		alert("Please enter the total monthly premium or set the percentage paid by agency to 0.");
		form.frmBenefitsBenVisFullFamily.focus();
		return false;
	}

//Vision Part Time Employee Validation(3)
   else if ((form.frmBenefitsBenVisPartEmployee.value != 0) && (form.frmBenefitsBenVisPartEmployeeAmount.value == 0))
	{
		alert("Please enter the total monthly premium or set the percentage paid by agency to 0.");
		form.frmBenefitsBenVisPartEmployee.focus();
		return false;
	}

//Vision Part Time Family Validation(4)

    else if ((form.frmBenefitsBenVisPartFamily.value != 0) && (form.frmBenefitsBenVisPartFamilyAmount.value == 0))
	{
		alert("Please enter the total monthly premium or set the percentage paid by agency to 0.");
		form.frmBenefitsBenVisPartFamily.focus();
		return false;
	}
		
	else if((!(myRegularExpression3.test(form.frmBenefitsBenVisFullEmployee.value))) || (form.frmBenefitsBenVisFullEmployee.value > 100) || (form.frmBenefitsBenVisFullEmployee.value < 0))
	{
		form.frmBenefitsBenVisFullEmployee.focus();
		alert((form.frmBenefitsBenVisFullEmployee.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression3.test(form.frmBenefitsBenVisFullEmployeeAmount.value)))	
	{
		form.frmBenefitsBenVisFullEmployeeAmount.focus();
		alert((form.frmBenefitsBenVisFullEmployeeAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
	else if((!(myRegularExpression3.test(form.frmBenefitsBenVisFullFamily.value))) || (form.frmBenefitsBenVisFullFamily.value > 100) || (form.frmBenefitsBenVisFullFamily.value < 0))
	{
		form.frmBenefitsBenVisFullFamily.focus();
		alert((form.frmBenefitsBenVisFullFamily.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression3.test(form.frmBenefitsBenVisFullFamilyAmount.value)))	
	{
		form.frmBenefitsBenVisFullFamilyAmount.focus();
		alert((form.frmBenefitsBenVisFullFamilyAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
	else if((!(myRegularExpression3.test(form.frmBenefitsBenVisPartEmployee.value))) || (form.frmBenefitsBenVisPartEmployee.value > 100) || (form.frmBenefitsBenVisPartEmployee.value < 0))
	{
		form.frmBenefitsBenVisPartEmployee.focus();
		alert((form.frmBenefitsBenVisPartEmployee.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression3.test(form.frmBenefitsBenVisPartEmployeeAmount.value)))	
	{
		form.frmBenefitsBenVisPartEmployeeAmount.focus();
		alert((form.frmBenefitsBenVisPartEmployeeAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
	else if((!(myRegularExpression3.test(form.frmBenefitsBenVisPartFamily.value))) || (form.frmBenefitsBenVisPartFamily.value > 100) || (form.frmBenefitsBenVisPartFamily.value < 0))
	{
		form.frmBenefitsBenVisPartFamily.focus();
		alert((form.frmBenefitsBenVisPartFamily.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression3.test(form.frmBenefitsBenVisPartFamilyAmount.value)))	
	{
		form.frmBenefitsBenVisPartFamilyAmount.focus();
		alert((form.frmBenefitsBenVisPartFamilyAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
	//SHORT TERM
	else if ((form.frmBenefitsDisInsShortTermFullPaid.checked == true) && (form.frmBenefitsDisInsShortTermFullPrcnt.value == 0) && (form.frmBenefitsDisInsShortTermFullAmount.value == 0))
	{
		alert("You selected that you pay some portion of Short Term Disability, but did not enter Percentages and Amounts. Please enter required data or select that you do NOT offer Short Term Disability");
		form.frmBenefitsDisInsShortTermFullPrcnt.focus();
		return false;
	}
	else if ((form.frmBenefitsDisInsShortTermPartPaid.checked == true) && (form.frmBenefitsDisInsShortTermPartPrcnt.value == 0) && (form.frmBenefitsDisInsShortTermPartAmount.value == 0))
	{
	    alert("You selected that you pay some portion of Short Term Disability, but did not enter Percentage and Amount. Please enter required data or select that you do NOT Pay for Short Term Disability");
		form.frmBenefitsDisInsShortTermPartPrcnt.focus();
		return false;
	}
    else if (((form.frmBenefitsDisInsShortTermFullPrcnt.value == 0) && (form.frmBenefitsDisInsShortTermFullAmount.value != 0)) || ((form.frmBenefitsDisInsShortTermFullPrcnt.value != 0) && (form.frmBenefitsDisInsShortTermFullAmount.value == 0)))
	{
		alert("Both Percent and Total Premium must be greater than 0 or equal 0. Please enter missing information.");
		form.frmBenefitsDisInsShortTermFullPrcnt.focus();
		return false;
	}
    else if (((form.frmBenefitsDisInsShortTermPartPrcnt.value == 0) && (form.frmBenefitsDisInsShortTermPartAmount.value != 0)) || ((form.frmBenefitsDisInsShortTermPartPrcnt.value != 0) && (form.frmBenefitsDisInsShortTermPartAmount.value == 0)))
	{
		alert("Both Percent and Total Premium must be greater than 0 or equal 0. Please enter missing information.");
		form.frmBenefitsDisInsShortTermPartPrcnt.focus();
		return false;
	}
	else if((!(myRegularExpression1.test(form.frmBenefitsDisInsShortTermFullPrcnt.value))) || (form.frmBenefitsDisInsShortTermFullPrcnt.value > 100) || (form.frmBenefitsDisInsShortTermFullPrcnt.value < 0))
	{
		form.frmBenefitsDisInsShortTermFullPrcnt.focus();
		alert((form.frmBenefitsDisInsShortTermFullPrcnt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBenefitsDisInsShortTermFullAmount.value)))	
	{
		form.frmBenefitsDisInsShortTermFullAmount.focus();
		alert((form.frmBenefitsDisInsShortTermFullAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
	else if((!(myRegularExpression1.test(form.frmBenefitsDisInsShortTermPartPrcnt.value))) || (form.frmBenefitsDisInsShortTermPartPrcnt.value > 100) || (form.frmBenefitsDisInsShortTermPartPrcnt.value < 0))
	{
		form.frmBenefitsDisInsShortTermPartPrcnt.focus();
		alert((form.frmBenefitsDisInsShortTermPartPrcnt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBenefitsDisInsShortTermPartAmount.value)))	
	{
		form.frmBenefitsDisInsShortTermPartAmount.focus();
		alert((form.frmBenefitsDisInsShortTermPartAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
	//LONG TERM
	else if ((form.frmBenefitsDisInsLongTermFullPaid.checked == true) && (form.frmBenefitsDisInsLongTermFullPrcnt.value == 0) && (form.frmBenefitsDisInsLongTermFullAmount.value == 0))
	{
		alert("You selected that you pay some portion of LONG TERM Disability, but did not enter Percentage and Amount. Please enter required data or select that you do NOT Pay for LONG TERM Disability");
		form.frmBenefitsDisInsLongTermFullPrcnt.focus();
		return false;
	}
	else if ((form.frmBenefitsDisInsLongTermPartPaid.checked == true) && (form.frmBenefitsDisInsLongTermPartPrcnt.value == 0) && (form.frmBenefitsDisInsLongTermPartAmount.value == 0))
	{
		alert("You selected that you pay some portion of LONG TERM Disability, but did not enter Percentage and Amount. Please enter required data or select that you do NOT Pay for LONG TERM Disability");
		form.frmBenefitsDisInsLongTermPartPrcnt.focus();
		return false;
	}
    else if (((form.frmBenefitsDisInsLongTermFullPrcnt.value == 0) && (form.frmBenefitsDisInsLongTermFullAmount.value != 0)) || ((form.frmBenefitsDisInsLongTermFullPrcnt.value != 0) && (form.frmBenefitsDisInsLongTermFullAmount.value == 0)))
	{
		alert("Both Percent and Total Premium must be greater than 0 or equal 0. Please enter missing information.");
		form.frmBenefitsDisInsLongTermFullPrcnt.focus();
		return false;
	}
    else if (((form.frmBenefitsDisInsLongTermPartPrcnt.value == 0) && (form.frmBenefitsDisInsLongTermPartAmount.value != 0)) || ((form.frmBenefitsDisInsLongTermPartPrcnt.value != 0) && (form.frmBenefitsDisInsLongTermPartAmount.value == 0)))
	{
		alert("Both Percent and Total Premium must be greater than 0 or equal 0. Please enter missing information.");
		form.frmBenefitsDisInsLongTermPartPrcnt.focus();
		return false;
	}
	else if((!(myRegularExpression1.test(form.frmBenefitsDisInsLongTermFullPrcnt.value))) || (form.frmBenefitsDisInsLongTermFullPrcnt.value > 100) || (form.frmBenefitsDisInsLongTermFullPrcnt.value < 0))
	{
		form.frmBenefitsDisInsLongTermFullPrcnt.focus();
		alert((form.frmBenefitsDisInsLongTermFullPrcnt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBenefitsDisInsLongTermFullAmount.value)))	
	{
		form.frmBenefitsDisInsLongTermFullAmount.focus();
		alert((form.frmBenefitsDisInsLongTermFullAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
	else if((!(myRegularExpression1.test(form.frmBenefitsDisInsLongTermPartPrcnt.value))) || (form.frmBenefitsDisInsLongTermPartPrcnt.value > 100) || (form.frmBenefitsDisInsLongTermPartPrcnt.value < 0))
	{
		form.frmBenefitsDisInsLongTermPartPrcnt.focus();
		alert((form.frmBenefitsDisInsLongTermPartPrcnt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBenefitsDisInsLongTermPartAmount.value)))	
	{
		form.frmBenefitsDisInsLongTermPartAmount.focus();
		alert((form.frmBenefitsDisInsLongTermPartAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
	//EAP
/*	else if ((form.frmBenefitsEAPFullPaid.checked == true) && (form.frmBenefitsEAPFullPrcnt.value == 0) && (form.frmBenefitsEAPFullAmount.value == 0))
	{
		alert("You selected that you pay some portion of Employee Assistance Program, but did not enter Percentage and Amount. Please enter required data or select that you do NOT Pay for Employee Assistance Program");
		form.frmBenefitsEAPFullPrcnt.focus();
		return false;
	}
	else if ((form.frmBenefitsEAPPartPaid.checked == true) && (form.frmBenefitsEAPPartPrcnt.value == 0) && (form.frmBenefitsEAPPartAmount.value == 0))
	{
		alert("You selected that you pay some portion of Employee Assistance Program, but did not enter Percentage and Amount. Please enter required data or select that you do NOT Pay for Employee Assistance Program");
		form.frmBenefitsEAPPartPrcnt.focus();
		return false;
	}
*/
    else if (((form.frmBenefitsEAPFullPrcnt.value == 0) && (form.frmBenefitsEAPFullAmount.value != 0)) || ((form.frmBenefitsEAPFullPrcnt.value != 0) && (form.frmBenefitsEAPFullAmount.value == 0)))
	{
		alert("Both Percent and Total Premium must be greater than 0 or equal 0. Please enter missing information.");
		form.frmBenefitsEAPFullPrcnt.focus();
		return false;
	}
    else if (((form.frmBenefitsEAPPartPrcnt.value == 0) && (form.frmBenefitsEAPPartAmount.value != 0)) || ((form.frmBenefitsEAPPartPrcnt.value != 0) && (form.frmBenefitsEAPPartAmount.value == 0)))
	{
		alert("Both Percent and Total Premium must be greater than 0 or equal 0. Please enter missing information.");
		form.frmBenefitsEAPPartPrcnt.focus();
		return false;
	}
	else if((!(myRegularExpression1.test(form.frmBenefitsEAPFullPrcnt.value))) || (form.frmBenefitsEAPFullPrcnt.value > 100) || (form.frmBenefitsEAPFullPrcnt.value < 0))
	{
		form.frmBenefitsEAPFullPrcnt.focus();
		alert((form.frmBenefitsEAPFullPrcnt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBenefitsEAPFullAmount.value)))	
	{
		form.frmBenefitsEAPFullAmount.focus();
		alert((form.frmBenefitsEAPFullAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
	else if((!(myRegularExpression1.test(form.frmBenefitsEAPPartPrcnt.value))) || (form.frmBenefitsEAPPartPrcnt.value > 100) || (form.frmBenefitsEAPPartPrcnt.value < 0))
	{
		form.frmBenefitsEAPPartPrcnt.focus();
		alert((form.frmBenefitsEAPPartPrcnt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBenefitsEAPPartAmount.value)))	
	{
		form.frmBenefitsEAPPartAmount.focus();
		alert((form.frmBenefitsEAPPartAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
	//FLEX PRE TAX
	//Commented due to change in requirements (not need to collect this)
	/*else if((!(myRegularExpression3.test(form.frmBenefitsFlexFullPrcnt.value))) || (form.frmBenefitsFlexFullPrcnt.value > 100) || (form.frmBenefitsFlexFullPrcnt.value < 0))
	{
		form.frmBenefitsFlexFullPrcnt.focus();
		alert((form.frmBenefitsFlexFullPrcnt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression3.test(form.frmBenefitsFlexFullAmount.value)))	
	{
		form.frmBenefitsFlexFullAmount.focus();
		alert((form.frmBenefitsFlexFullAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
	else if((!(myRegularExpression3.test(form.frmBenefitsFlexPartPrcnt.value))) || (form.frmBenefitsFlexPartPrcnt.value > 100) || (form.frmBenefitsFlexPartPrcnt.value < 0))
	{
		form.frmBenefitsFlexPartPrcnt.focus();
		alert((form.frmBenefitsFlexPartPrcnt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression3.test(form.frmBenefitsFlexPartAmount.value)))	
	{
		form.frmBenefitsFlexPartAmount.focus();
		alert((form.frmBenefitsFlexPartAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}*/
	
	//HEALTH CLUB
/*	else if ((form.frmBenefitsHealthClubFullPaid.checked == true) && (form.frmBenefitsHealthClubFullPrcnt.value == 0) && (form.frmBenefitsHealthClubFullAmount.value == 0))
	{
		alert("You selected that you pay some portion for Health Club, but did not enter Percentage and Amount. Please enter required data or select that you do NOT Pay for Health Club");
		form.frmBenefitsHealthClubFullPrcnt.focus();
		return false;
	}
	else if ((form.frmBenefitsHealthClubPartPaid.checked == true) && (form.frmBenefitsHealthClubPartPrcnt.value == 0) && (form.frmBenefitsHealthClubPartAmount.value == 0))
	{
		alert("You selected that you pay some portion for Health Club, but did not enter Percentage and Amount. Please enter required data or select that you do NOT Pay for Health Club");
		form.frmBenefitsHealthClubPartPrcnt.focus();
		return false;
	}
*/
    else if (((form.frmBenefitsHealthClubFullPrcnt.value == 0) && (form.frmBenefitsHealthClubFullAmount.value != 0)) || ((form.frmBenefitsHealthClubFullPrcnt.value != 0) && (form.frmBenefitsHealthClubFullAmount.value == 0)))
	{
		alert("Both Percent and Total Premium must be greater than 0 or equal 0. Please enter missing information.");
		form.frmBenefitsHealthClubFullPrcnt.focus();
		return false;
	}
    else if (((form.frmBenefitsHealthClubPartPrcnt.value == 0) && (form.frmBenefitsHealthClubPartAmount.value != 0)) || ((form.frmBenefitsHealthClubPartPrcnt.value != 0) && (form.frmBenefitsHealthClubPartAmount.value == 0)))
	{
		alert("Both Percent and Total Premium must be greater than 0 or equal 0. Please enter missing information.");
		form.frmBenefitsHealthClubPartPrcnt.focus();
		return false;
	}
	else if((!(myRegularExpression1.test(form.frmBenefitsHealthClubFullPrcnt.value))) || (form.frmBenefitsHealthClubFullPrcnt.value > 100) || (form.frmBenefitsHealthClubFullPrcnt.value < 0))
	{
		form.frmBenefitsHealthClubFullPrcnt.focus();
		alert((form.frmBenefitsHealthClubFullPrcnt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBenefitsHealthClubFullAmount.value)))	
	{
		form.frmBenefitsHealthClubFullAmount.focus();
		alert((form.frmBenefitsHealthClubFullAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
	else if((!(myRegularExpression1.test(form.frmBenefitsHealthClubPartPrcnt.value))) || (form.frmBenefitsHealthClubPartPrcnt.value > 100) || (form.frmBenefitsHealthClubPartPrcnt.value < 0))
	{
		form.frmBenefitsHealthClubPartPrcnt.focus();
		alert((form.frmBenefitsHealthClubPartPrcnt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBenefitsHealthClubPartAmount.value)))	
	{
		form.frmBenefitsHealthClubPartAmount.focus();
		alert((form.frmBenefitsHealthClubPartAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
	//LIFE INSURANCE
	else if ((form.frmBenefitsLifeInsuranceFullPaid.checked == true) && (form.frmBenefitsLifeInsuranceFullPrcnt.value == 0) && (form.frmBenefitsLifeInsuranceFullAmount.value == 0))
	{
		alert("You selected that you pay some portion for Life Insurance, but did not enter Percentage and Amount. Please enter required data or select that you do NOT Pay for Life Insurance");
		form.frmBenefitsLifeInsuranceFullPrcnt.focus();
		return false;
	}
	else if ((form.frmBenefitsLifeInsurancePartPaid.checked == true) && (form.frmBenefitsLifeInsurancePartPrcnt.value == 0) && (form.frmBenefitsLifeInsurancePartAmount.value == 0))
	{
		alert("You selected that you pay some portion for Life Insurance, but did not enter Percentage and Amount. Please enter required data or select that you do NOT Pay for Life Insurance");
		form.frmBenefitsLifeInsurancePartPrcnt.focus();
		return false;
	}
    else if (((form.frmBenefitsLifeInsuranceFullPrcnt.value == 0) && (form.frmBenefitsLifeInsuranceFullAmount.value != 0)) || ((form.frmBenefitsLifeInsuranceFullPrcnt.value != 0) && (form.frmBenefitsLifeInsuranceFullAmount.value == 0)))
	{
		alert("Both Percent and Total Premium must be greater than 0 or equal 0. Please enter missing information.");
		form.frmBenefitsLifeInsuranceFullPrcnt.focus();
		return false;
	}
    else if (((form.frmBenefitsLifeInsurancePartPrcnt.value == 0) && (form.frmBenefitsLifeInsurancePartAmount.value != 0)) || ((form.frmBenefitsLifeInsurancePartPrcnt.value != 0) && (form.frmBenefitsLifeInsurancePartAmount.value == 0)))
	{
		alert("Both Percent and Total Premium must be greater than 0 or equal 0. Please enter missing information.");
		form.frmBenefitsLifeInsurancePartPrcnt.focus();
		return false;
	}
	else if((!(myRegularExpression1.test(form.frmBenefitsLifeInsuranceFullPrcnt.value))) || (form.frmBenefitsLifeInsuranceFullPrcnt.value > 100) || (form.frmBenefitsLifeInsuranceFullPrcnt.value < 0))
	{
		form.frmBenefitsLifeInsuranceFullPrcnt.focus();
		alert((form.frmBenefitsLifeInsuranceFullPrcnt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBenefitsLifeInsuranceFullAmount.value)))	
	{
		form.frmBenefitsLifeInsuranceFullAmount.focus();
		alert((form.frmBenefitsLifeInsuranceFullAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
	else if((!(myRegularExpression1.test(form.frmBenefitsLifeInsurancePartPrcnt.value))) || (form.frmBenefitsLifeInsurancePartPrcnt.value > 100) || (form.frmBenefitsLifeInsurancePartPrcnt.value < 0))
	{
		form.frmBenefitsLifeInsurancePartPrcnt.focus();
		alert((form.frmBenefitsLifeInsurancePartPrcnt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression1.test(form.frmBenefitsLifeInsurancePartAmount.value)))	
	{
		form.frmBenefitsLifeInsurancePartAmount.focus();
		alert((form.frmBenefitsLifeInsurancePartAmount.value) + " is invalid. Please enter a whole number between 0 and 10000.");
		return false;
	}
	
	//TIME OFF
	else if ((form.frmBenefitsTimeOffFull.checked == true) && ((form.frmBenefitsTimeOffFullDays.value == 0)&&(form.frmBenefitsTimeOffFullYears.value == 0)&&(form.frmBenefitsTimeOffFullDaysIncreased.value == 0)))
	{
		alert("You selected that you offer Paid Time Off, but did not enter the number of paid days or years or days after increase. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffFullDays.focus();
		return false;
	}
	else if ((form.frmBenefitsTimeOffFull.checked == true) && ((form.frmBenefitsTimeOffFullDays.value == 0)&&(form.frmBenefitsTimeOffFullYears.value == 0)&&(form.frmBenefitsTimeOffFullDaysIncreased.value > 0)))
	{
		alert("You selected that you offer Paid Time Off and entered number of days after increase, but did not enter the number of paid days for new employee and years before increase. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffFullDays.focus();
		return false;
	}
	else if ((form.frmBenefitsTimeOffFull.checked == true) && ((form.frmBenefitsTimeOffFullDays.value == 0)&&(form.frmBenefitsTimeOffFullYears.value > 0)&&(form.frmBenefitsTimeOffFullDaysIncreased.value == 0)))
	{
		alert("You selected that you offer Paid Time Off and entered number of years before increase, but did not enter the number of paid days for new employee and number of days after increase. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffFullDays.focus();
		return false;
	}
	else if ((form.frmBenefitsTimeOffFull.checked == true) && ((form.frmBenefitsTimeOffFullDays.value > 0)&&(form.frmBenefitsTimeOffFullYears.value == 0)&&(form.frmBenefitsTimeOffFullDaysIncreased.value == 0)))
	{
		alert("You selected that you offer Paid Time Off and entered number of days for new employee, but did not enter the number of years before increase and number of days after increase. If you do not offer increase, make the number of days after increase the same as number of days for new employee. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffFullDays.focus();
		return false;
	}
	else if ((form.frmBenefitsTimeOffFull.checked == true) && ((form.frmBenefitsTimeOffFullDays.value > 0)&&(form.frmBenefitsTimeOffFullYears.value > 0)&&(form.frmBenefitsTimeOffFullDaysIncreased.value == 0)))
	{
		alert("You selected that you offer Paid Time Off and entered number of days for new employee and years before increase, but did not enter the number of days after increase. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffFullDays.focus();
		return false;
	}
	//
	else if ((form.frmBenefitsTimeOffFullNExempt.checked == true) && ((form.frmBenefitsTimeOffFullDaysNExempt.value == 0)&&(form.frmBenefitsTimeOffFullYearsNExempt.value == 0)&&(form.frmBenefitsTimeOffFullDaysIncreasedNExempt.value == 0)))
	{
		alert("You selected that you offer Paid Time Off, but did not enter the number of paid days or years or days after increase. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffFullDaysNExempt.focus();
		return false;
	}
	else if ((form.frmBenefitsTimeOffFullNExempt.checked == true) && ((form.frmBenefitsTimeOffFullDaysNExempt.value == 0)&&(form.frmBenefitsTimeOffFullYearsNExempt.value == 0)&&(form.frmBenefitsTimeOffFullDaysIncreasedNExempt.value > 0)))
	{
		alert("You selected that you offer Paid Time Off and entered number of days after increase, but did not enter the number of paid days for new employee and years before increase. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffFullDaysNExempt.focus();
		return false;
	}
	else if ((form.frmBenefitsTimeOffFullNExempt.checked == true) && ((form.frmBenefitsTimeOffFullDaysNExempt.value == 0)&&(form.frmBenefitsTimeOffFullYearsNExempt.value > 0)&&(form.frmBenefitsTimeOffFullDaysIncreasedNExempt.value == 0)))
	{
		alert("You selected that you offer Paid Time Off and entered number of years before increase, but did not enter the number of paid days for new employee and number of days after increase. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffFullDaysNExempt.focus();
		return false;
	}
	else if ((form.frmBenefitsTimeOffFullNExempt.checked == true) && ((form.frmBenefitsTimeOffFullDaysNExempt.value > 0)&&(form.frmBenefitsTimeOffFullYearsNExempt.value == 0)&&(form.frmBenefitsTimeOffFullDaysIncreasedNExempt.value == 0)))
	{
		alert("You selected that you offer Paid Time Off and entered number of days for new employee, but did not enter the number of years before increase and number of days after increase. If you do not offer increase, make the number of days after increase the same as number of days for new employee. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffFullDaysNExempt.focus();
		return false;
	}
	else if ((form.frmBenefitsTimeOffFullNExempt.checked == true) && ((form.frmBenefitsTimeOffFullDaysNExempt.value > 0)&&(form.frmBenefitsTimeOffFullYearsNExempt.value > 0)&&(form.frmBenefitsTimeOffFullDaysIncreasedNExempt.value == 0)))
	{
		alert("You selected that you offer Paid Time Off and entered number of days for new employee and years before increase, but did not enter the number of days after increase. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffFullDaysNExempt.focus();
		return false;
	}
	//
	else if ((form.frmBenefitsTimeOffPart.checked == true) && ((form.frmBenefitsTimeOffPartDays.value == 0)&&(form.frmBenefitsTimeOffPartYears.value == 0)&&(form.frmBenefitsTimeOffPartDaysIncreased.value == 0)))
	{
		alert("You selected that you offer Paid Time Off, but did not enter the number of paid days. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffPartDays.focus();
		return false;
	}
	else if ((form.frmBenefitsTimeOffPart.checked == true) && ((form.frmBenefitsTimeOffPartDays.value == 0)&&(form.frmBenefitsTimeOffPartYears.value == 0)&&(form.frmBenefitsTimeOffPartDaysIncreased.value > 0)))
	{
		alert("You selected that you offer Paid Time Off and entered number of days after increase, but did not enter the number of paid days for new employee and years before increase. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffPartDays.focus();
		return false;
	}
	else if ((form.frmBenefitsTimeOffPart.checked == true) && ((form.frmBenefitsTimeOffPartDays.value == 0)&&(form.frmBenefitsTimeOffPartYears.value > 0)&&(form.frmBenefitsTimeOffPartDaysIncreased.value == 0)))
	{
		alert("You selected that you offer Paid Time Off and entered number of years before increase, but did not enter the number of paid days for new employee and number of days after increase. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffPartDays.focus();
		return false;
	}
	else if ((form.frmBenefitsTimeOffPart.checked == true) && ((form.frmBenefitsTimeOffPartDays.value > 0)&&(form.frmBenefitsTimeOffPartYears.value == 0)&&(form.frmBenefitsTimeOffPartDaysIncreased.value == 0)))
	{
		alert("You selected that you offer Paid Time Off and entered number of days for new employee, but did not enter the number of years before increase and number of days after increase. If you do not offer increase, make the number of days after increase the same as number of days for new employee. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffPartDays.focus();
		return false;
	}
	else if ((form.frmBenefitsTimeOffPart.checked == true) && ((form.frmBenefitsTimeOffPartDays.value > 0)&&(form.frmBenefitsTimeOffPartYears.value > 0)&&(form.frmBenefitsTimeOffPartDaysIncreased.value == 0)))
	{
		alert("You selected that you offer Paid Time Off and entered number of days for new employee and years before increase, but did not enter the number of days after increase. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffPartDays.focus();
		return false;
	}
	//
	else if ((form.frmBenefitsTimeOffPartNExempt.checked == true) && ((form.frmBenefitsTimeOffPartDaysNExempt.value == 0)&&(form.frmBenefitsTimeOffPartYearsNExempt.value == 0)&&(form.frmBenefitsTimeOffPartDaysIncreasedNExempt.value == 0)))
	{
		alert("You selected that you offer Paid Time Off, but did not enter the number of paid days or years or days after increase. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffPartDaysNExempt.focus();
		return false;
	}
	else if ((form.frmBenefitsTimeOffPartNExempt.checked == true) && ((form.frmBenefitsTimeOffPartDaysNExempt.value == 0)&&(form.frmBenefitsTimeOffPartYearsNExempt.value == 0)&&(form.frmBenefitsTimeOffPartDaysIncreasedNExempt.value > 0)))
	{
		alert("You selected that you offer Paid Time Off and entered number of days after increase, but did not enter the number of paid days for new employee and years before increase. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffPartDaysNExempt.focus();
		return false;
	}
	else if ((form.frmBenefitsTimeOffPartNExempt.checked == true) && ((form.frmBenefitsTimeOffPartDaysNExempt.value == 0)&&(form.frmBenefitsTimeOffPartYearsNExempt.value > 0)&&(form.frmBenefitsTimeOffPartDaysIncreasedNExempt.value == 0)))
	{
		alert("You selected that you offer Paid Time Off and entered number of years before increase, but did not enter the number of paid days for new employee and number of days after increase. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffPartDaysNExempt.focus();
		return false;
	}
	else if ((form.frmBenefitsTimeOffPartNExempt.checked == true) && ((form.frmBenefitsTimeOffPartDaysNExempt.value > 0)&&(form.frmBenefitsTimeOffPartYearsNExempt.value == 0)&&(form.frmBenefitsTimeOffPartDaysIncreasedNExempt.value == 0)))
	{
		alert("You selected that you offer Paid Time Off and entered number of days for new employee, but did not enter the number of years before increase and number of days after increase. If you do not offer increase, make the number of days after increase the same as number of days for new employee. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffPartDaysNExempt.focus();
		return false;
	}
	else if ((form.frmBenefitsTimeOffPartNExempt.checked == true) && ((form.frmBenefitsTimeOffPartDaysNExempt.value > 0)&&(form.frmBenefitsTimeOffPartYearsNExempt.value > 0)&&(form.frmBenefitsTimeOffPartDaysIncreasedNExempt.value == 0)))
	{
		alert("You selected that you offer Paid Time Off and entered number of days for new employee and years before increase, but did not enter the number of days after increase. Please enter required data or select that you do NOT offer Paid Time Off");
		form.frmBenefitsTimeOffPartDaysNExempt.focus();
		return false;
	}
	//
	
	else if((!(myRegularExpression3.test(form.frmBenefitsTimeOffFullDays.value))) || (form.frmBenefitsTimeOffFullDays.value > 100) || (form.frmBenefitsTimeOffFullDays.value < 0))
	{
		form.frmBenefitsTimeOffFullDays.focus();
		alert((form.frmBenefitsTimeOffFullDays.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	/*else if((!(myRegularExpression3.test(form.frmBenefitsTimeOffFullYears.value))) || (form.frmBenefitsTimeOffFullYears.value > 100) || (form.frmBenefitsTimeOffFullYears.value < 0))
	{
		form.frmBenefitsTimeOffFullYears.focus();
		alert((form.frmBenefitsTimeOffFullYears.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}*/
	else if((!(myRegularExpression3.test(form.frmBenefitsTimeOffFullDaysIncreased.value))) || (form.frmBenefitsTimeOffFullDaysIncreased.value > 100) || (form.frmBenefitsTimeOffFullDaysIncreased.value < 0))
	{
		form.frmBenefitsTimeOffFullDaysIncreased.focus();
		alert((form.frmBenefitsTimeOffFullDays.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}

	else if((!(myRegularExpression3.test(form.frmBenefitsTimeOffFullDaysNExempt.value))) || (form.frmBenefitsTimeOffFullDaysNExempt.value > 100) || (form.frmBenefitsTimeOffFullDaysNExempt.value < 0))
	{
		form.frmBenefitsTimeOffFullDaysNExempt.focus();
		alert((form.frmBenefitsTimeOffFullDaysNExempt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	/*else if((!(myRegularExpression3.test(form.frmBenefitsTimeOffFullYearsNExempt.value))) || (form.frmBenefitsTimeOffFullYearsNExempt.value > 100) || (form.frmBenefitsTimeOffFullYearsNExempt.value < 0))
	{
		form.frmBenefitsTimeOffFullYearsNExempt.focus();
		alert((form.frmBenefitsTimeOffFullYearsNExempt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}*/
	else if((!(myRegularExpression3.test(form.frmBenefitsTimeOffFullDaysIncreasedNExempt.value))) || (form.frmBenefitsTimeOffFullDaysIncreasedNExempt.value > 100) || (form.frmBenefitsTimeOffFullDaysIncreasedNExempt.value < 0))
	{
		form.frmBenefitsTimeOffFullDaysIncreasedNExempt.focus();
		alert((form.frmBenefitsTimeOffFullDaysNExempt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}

	else if((!(myRegularExpression3.test(form.frmBenefitsTimeOffPartDays.value))) || (form.frmBenefitsTimeOffPartDays.value > 100) || (form.frmBenefitsTimeOffPartDays.value < 0))
	{
		form.frmBenefitsTimeOffPartDays.focus();
		alert((form.frmBenefitsTimeOffPartDays.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	/*else if((!(myRegularExpression3.test(form.frmBenefitsTimeOffPartYears.value))) || (form.frmBenefitsTimeOffPartYears.value > 100) || (form.frmBenefitsTimeOffPartYears.value < 0))
	{
		form.frmBenefitsTimeOffPartYears.focus();
		alert((form.frmBenefitsTimeOffPartYears.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}*/
	else if((!(myRegularExpression3.test(form.frmBenefitsTimeOffPartDaysIncreased.value))) || (form.frmBenefitsTimeOffPartDaysIncreased.value > 100) || (form.frmBenefitsTimeOffPartDaysIncreased.value < 0))
	{
		form.frmBenefitsTimeOffPartDaysIncreased.focus();
		alert((form.frmBenefitsTimeOffPartDays.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}

	else if((!(myRegularExpression3.test(form.frmBenefitsTimeOffPartDaysNExempt.value))) || (form.frmBenefitsTimeOffPartDaysNExempt.value > 100) || (form.frmBenefitsTimeOffPartDaysNExempt.value < 0))
	{
		form.frmBenefitsTimeOffPartDaysNExempt.focus();
		alert((form.frmBenefitsTimeOffPartDaysNExempt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	/*else if((!(myRegularExpression3.test(form.frmBenefitsTimeOffPartYearsNExempt.value))) || (form.frmBenefitsTimeOffPartYearsNExempt.value > 100) || (form.frmBenefitsTimeOffPartYearsNExempt.value < 0))
	{
		form.frmBenefitsTimeOffPartYearsNExempt.focus();
		alert((form.frmBenefitsTimeOffPartYearsNExempt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}*/
	else if((!(myRegularExpression3.test(form.frmBenefitsTimeOffPartDaysIncreasedNExempt.value))) || (form.frmBenefitsTimeOffPartDaysIncreasedNExempt.value > 100) || (form.frmBenefitsTimeOffPartDaysIncreasedNExempt.value < 0))
	{
		form.frmBenefitsTimeOffPartDaysIncreasedNExempt.focus();
		alert((form.frmBenefitsTimeOffPartDaysNExempt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	
	//TIME SICK
	else if ((form.frmBenefitsTimeOffSickFull.checked == true) && (form.frmBenefitsTimeOffSickFullDays.value == 0))
	{
		alert("You selected that you offer Paid Sick Time, but did not enter the number of paid sick days. Please enter required data or select that you do NOT offer Paid Sick Time");
		form.frmBenefitsTimeOffSickFullDays.focus();
		return false;
	}
	else if ((form.frmBenefitsTimeOffSickPart.checked == true) && (form.frmBenefitsTimeOffSickPartDays.value == 0))
	{
		alert("You selected that you offer Paid Sick Time, but did not enter the number of paid sick days. Please enter required data or select that you do NOT offer Paid Sick Time");
		form.frmBenefitsTimeOffSickPartDays.focus();
		return false;
	}
	else if((!(myRegularExpression3.test(form.frmBenefitsTimeOffSickFullDays.value))) || (form.frmBenefitsTimeOffSickFullDays.value > 100) || (form.frmBenefitsTimeOffSickFullDays.value < 0))
	{
		form.frmBenefitsTimeOffSickFullDays.focus();
		alert((form.frmBenefitsTimeOffSickFullDays.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if((!(myRegularExpression3.test(form.frmBenefitsTimeOffSickPartDays.value))) || (form.frmBenefitsTimeOffSickPartDays.value > 100) || (form.frmBenefitsTimeOffSickPartDays.value < 0))
	{
		form.frmBenefitsTimeOffSickPartDays.focus();
		alert((form.frmBenefitsTimeOffSickPartDays.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	
	//TIME VACATION
	/*else if((!(myRegularExpression3.test(form.frmBenefitsTimeOffVacFullDays.value))) || (form.frmBenefitsTimeOffVacFullDays.value > 100) || (form.frmBenefitsTimeOffVacFullDays.value < 0))
	{
		form.frmBenefitsTimeOffVacFullDays.focus();
		alert((form.frmBenefitsTimeOffVacFullDays.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if((!(myRegularExpression3.test(form.frmBenefitsTimeOffVacPartDays.value))) || (form.frmBenefitsTimeOffVacPartDays.value > 100) || (form.frmBenefitsTimeOffVacPartDays.value < 0))
	{
		form.frmBenefitsTimeOffVacPartDays.focus();
		alert((form.frmBenefitsTimeOffVacPartDays.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}*/
	
	//Prof DUES
	else if ((form.frmBenefitsProfDuesFullPaid.checked == true) && (form.frmBenefitsProfDuesFullAmount.value == 0))
	{
		alert("You selected that you pay some portion of Professional Dues, but did not enter the Amount. Please enter required data or select that you do NOT pay Professional Dues");
		form.frmBenefitsProfDuesFullAmount.focus();
		return false;
	}
	else if ((form.frmBenefitsProfDuesPartPaid.checked == true) && (form.frmBenefitsProfDuesPartAmount.value == 0))
	{
		alert("You selected that you pay some portion of Professional Dues, but did not enter the Amount. Please enter required data or select that you do NOT pay Professional Dues");
		form.frmBenefitsProfDuesPartAmount.focus();
		return false;
	}
	else if(!(myRegularExpression3.test(form.frmBenefitsProfDuesFullAmount.value))) // || (form.frmBenefitsProfDuesFullAmount.value > 100) || (form.frmBenefitsProfDuesFullAmount.value < 0))
	{
		form.frmBenefitsProfDuesFullAmount.focus();
		alert((form.frmBenefitsProfDuesFullAmount.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression3.test(form.frmBenefitsProfDuesPartAmount.value))) // || (form.frmBenefitsProfDuesPartAmount.value > 100) || (form.frmBenefitsProfDuesPartAmount.value < 0))
	{
		form.frmBenefitsProfDuesPartAmount.focus();
		alert((form.frmBenefitsProfDuesPartAmount.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	
	//PENSION PLAN
	else if ((form.frmBenefitsRetirementFull.checked == true) && (form.frmBenefitsRetirementFullPrcnt.value == 0))
	{
		alert("You selected that you offer Pension Plan, but did not enter the Pecent of Agency Contribution. Please enter required data or select that you do NOT offer Pension Plan");
		form.frmBenefitsRetirementFullPrcnt.focus();
		return false;
	}
	else if ((form.frmBenefitsRetirementPart.checked == true) && (form.frmBenefitsRetirementPartPrcnt.value == 0))
	{
		alert("You selected that you offer Pension Plan, but did not enter the Pecent of Agency Contribution. Please enter required data or select that you do NOT offer Pension Plan");
		form.frmBenefitsRetirementPartPrcnt.focus();
		return false;
	}
	else if((!(myRegularExpression3.test(form.frmBenefitsRetirementFullPrcnt.value))) || (form.frmBenefitsRetirementFullPrcnt.value > 100) || (form.frmBenefitsRetirementFullPrcnt.value < 0))
	{
		form.frmBenefitsRetirementFullPrcnt.focus();
		alert((form.frmBenefitsRetirementFullPrcnt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if((!(myRegularExpression3.test(form.frmBenefitsRetirementPartPrcnt.value))) || (form.frmBenefitsRetirementPartPrcnt.value > 100) || (form.frmBenefitsRetirementPartPrcnt.value < 0))
	{
		form.frmBenefitsRetirementPartPrcnt.focus();
		alert((form.frmBenefitsRetirementPartPrcnt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	// 403B
	else if(form.frmBenefits403BFullContrib.value == "selected")	
	{
		form.frmBenefits403BFullContrib.focus();
		alert("Please select whether your agency matches employee contribution for 403B.");
		return false;
	}
	else if ((form.frmBenefits403BFullContrib.value == "true") && (form.frmBenefits403BFullContribPrcnt.value == 0))
	{
		alert("You selected that your agency contributes to EE 403B, but did not enter the Pecent of Matching Contribution. Please enter required data or select that you do NOT match EE contribution");
		form.frmBenefits403BFullContribPrcnt.focus();
		return false;
	}
	else if ((form.frmBenefits403BPartContrib.value == "true") && (form.frmBenefits403BPartContribPrcnt.value == 0))
	{
		alert("You selected that your agency contributes to EE 403B, but did not enter the Pecent of Matching Contribution. Please enter required data or select that you do NOT match EE contribution");
		form.frmBenefits403BPartContribPrcnt.focus();
		return false;
	}
	else if((form.frmBenefits403BFullContrib.value == "true")&&((form.frmBenefits403BFullContribPrcnt.value.length < 1)||(form.frmBenefits403BFullContribPrcnt.value == 0)))	
	{
		form.frmBenefits403BFullContribPrcnt.focus();
		alert("Please enter pecent of matching contribution for 403B.");
		return false;
	}
	else if(form.frmBenefits403BPartContrib.value == "selected")	
	{
		form.frmBenefits403BPartContrib.focus();
		alert("Please select whether your agency matches employee contribution for 403B.");
		return false;
	}
	else if((form.frmBenefits403BPartContrib.value == "true")&&((form.frmBenefits403BPartContribPrcnt.value.length < 1)||(form.frmBenefits403BPartContribPrcnt.value == 0)))
	{
		form.frmBenefits403BPartContribPrcnt.focus();
		alert("Please enter pecent of matching contribution for 403B.");
		return false;
	}
	
	//TELECOMMUTING
	else if (((form.frmBenefitsTelecommFullCount.value == 0) && (form.frmBenefitsTelecommFullPrcnt.value != 0)) || ((form.frmBenefitsTelecommFullCount.value != 0) && (form.frmBenefitsTelecommFullPrcnt.value == 0)))
	{
		alert("Both Number of EEs and Percent of EE Population must be greater than 0 or equal 0. Please enter missing information.");
		form.frmBenefitsTelecommFullCount.focus();
		return false;
	}
	else if (((form.frmBenefitsTelecommPartCount.value == 0) && (form.frmBenefitsTelecommPartPrcnt.value != 0)) || ((form.frmBenefitsTelecommPartCount.value != 0) && (form.frmBenefitsTelecommPartPrcnt.value == 0)))
	{
		alert("Both Number of EEs and Percent of EE Population must be greater than 0 or equal 0. Please enter missing information.");
		form.frmBenefitsTelecommPartCount.focus();
		return false;
	}
	else if((!(myRegularExpression3.test(form.frmBenefitsTelecommFullCount.value))) || (form.frmBenefitsTelecommFullCount.value > 100) || (form.frmBenefitsTelecommFullCount.value < 0))
	{
		form.frmBenefitsTelecommFullCount.focus();
		alert((form.frmBenefitsTelecommFullCount.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if((!(myRegularExpression3.test(form.frmBenefitsTelecommFullPrcnt.value))) || (form.frmBenefitsTelecommFullPrcnt.value > 100) || (form.frmBenefitsTelecommFullPrcnt.value < 0))
	{
		form.frmBenefitsTelecommFullPrcnt.focus();
		alert((form.frmBenefitsTelecommFullPrcnt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	
	else if((!(myRegularExpression3.test(form.frmBenefitsTelecommPartCount.value))) || (form.frmBenefitsTelecommPartCount.value > 100) || (form.frmBenefitsTelecommPartCount.value < 0))
	{
		form.frmBenefitsTelecommPartCount.focus();
		alert((form.frmBenefitsTelecommPartCount.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if((!(myRegularExpression3.test(form.frmBenefitsTelecommPartPrcnt.value))) || (form.frmBenefitsTelecommPartPrcnt.value > 100) || (form.frmBenefitsTelecommPartPrcnt.value < 0))
	{
		form.frmBenefitsTelecommPartPrcnt.focus();
		alert((form.frmBenefitsTelecommPartPrcnt.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	
	//TUITION
	else if ((form.frmBenefitsTuitionFull.checked == true) && (form.frmBenefitsTuitionFullAmount.value == 0))
	{
		alert("You selected that you offer Tuition Reimbursement, but did not enter the maximum Amount. Please enter required data or select that you do NOT offer Tuition Reimbursement");
		form.frmBenefitsTuitionFullAmount.focus();
		return false;
	}
	else if ((form.frmBenefitsTuitionPart.checked == true) && (form.frmBenefitsTuitionPartAmount.value == 0))
	{
		alert("You selected that you offer Tuition Reimbursement, but did not enter the maximum Amount. Please enter required data or select that you do NOT offer Tuition Reimbursement");
		form.frmBenefitsTuitionPartAmount.focus();
		return false;
	}
	else if(!(myRegularExpression3.test(form.frmBenefitsTuitionFullAmount.value))) //|| (form.frmBenefitsTuitionFullAmount.value > 100) || (form.frmBenefitsTuitionFullAmount.value < 0))
	{
		form.frmBenefitsTuitionFullAmount.focus();
		alert((form.frmBenefitsTuitionFullAmount.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	else if(!(myRegularExpression3.test(form.frmBenefitsTuitionPartAmount.value))) //|| (form.frmBenefitsTuitionPartAmount.value > 100) || (form.frmBenefitsTuitionPartAmount.value < 0))
	{
		form.frmBenefitsTuitionPartAmount.focus();
		alert((form.frmBenefitsTuitionPartAmount.value) + " is invalid. Please enter a whole number between 0 and 100.");
		return false;
	}
	
	else
	{
		//medical
		if((form.frmBenefitsBenMedOffered.value == "true") && (Number(form.frmBenefitsBenMedFullEmployeeAmount.value)>300))
		{
		form.frmBenefitsBenMedFullEmployeeAmount.focus();
		return confirm("The medical monthly premium ( $" + form.frmBenefitsBenMedFullEmployeeAmount.value + " ) you entered is outside of normal range. Click OK to continue or CANCEL to re enter the premium amount.");
		}
		else if((form.frmBenefitsBenMedOffered.value == "true") && (Number(form.frmBenefitsBenMedFullFamilyAmount.value)>800))
		{
		form.frmBenefitsBenMedFullFamilyAmount.focus();
		return confirm("The medical monthly premium ( $" + form.frmBenefitsBenMedFullFamilyAmount.value + " ) you entered is outside of normal range. Click OK to continue or CANCEL to re enter the premium amount.");
		}
		//dental
		else if((form.frmBenefitsBenDentOffered.value == "true") && (Number(form.frmBenefitsBenDentFullEmployeeAmount.value)>50))
		{
		form.frmBenefitsBenDentFullEmployeeAmount.focus();
		return confirm("The dental monthly premium ( $" + form.frmBenefitsBenDentFullEmployeeAmount.value + " ) you entered is outside of normal range. Click OK to continue or CANCEL to re enter the premium amount.");
		}
		else if((form.frmBenefitsBenDentOffered.value == "true") && (Number(form.frmBenefitsBenDentFullFamilyAmount.value)>80))
		{
		form.frmBenefitsBenDentFullFamilyAmount.focus();
		return confirm("The dental monthly premium ( $" + form.frmBenefitsBenDentFullFamilyAmount.value + " ) you entered is outside of normal range. Click OK to continue or CANCEL to re enter the premium amount.");
		}
		//vision
		else if((form.frmBenefitsBenVisOffered.value == "true") && (Number(form.frmBenefitsBenVisFullEmployeeAmount.value)>50))
		{
		form.frmBenefitsBenVisFullEmployeeAmount.focus();
		return confirm("The vision monthly premium ( $" + form.frmBenefitsBenVisFullEmployeeAmount.value + " ) you entered is outside of normal range. Click OK to continue or CANCEL to re enter the premium amount.");
		}
		else if((form.frmBenefitsBenVisOffered.value == "true") && (Number(form.frmBenefitsBenVisFullFamilyAmount.value)>75))
		{
		form.frmBenefitsBenVisFullFamilyAmount.focus();
		return confirm("The vision monthly premium ( $" + form.frmBenefitsBenVisFullFamilyAmount.value + " ) you entered is outside of normal range. Click OK to continue or CANCEL to re enter the premium amount.");
		}
		//short term
		else if((form.frmBenefitsDisInsShortTermFullPaid.checked == true) && (Number(form.frmBenefitsDisInsShortTermFullAmount.value)>10))
		{
		form.frmBenefitsDisInsShortTermFullAmount.focus();
		return confirm("The short term disability monthly premium ( $" + form.frmBenefitsDisInsShortTermFullAmount.value + " ) you entered is outside of normal range. Click OK to continue or CANCEL to re enter the premium amount.");
		}
		//long term
		else if((form.frmBenefitsDisInsLongTermFullPaid.checked == true) && (Number(form.frmBenefitsDisInsLongTermFullAmount.value)>10))
		{
		form.frmBenefitsDisInsLongTermFullAmount.focus();
		return confirm("The long term disability monthly premium ( $" + form.frmBenefitsDisInsLongTermFullAmount.value + " ) you entered is outside of normal range. Click OK to continue or CANCEL to re enter the premium amount.");
		}
		else
		{
			return true;
		}
	}
}
function ZeroIfNotChecked(form)
{
	//Short Term
	if (form.frmBenefitsDisInsShortTermFullPaid.checked == true)
	{
		form.frmBenefitsDisInsShortTermFullPrcnt.disabled = false;
		form.frmBenefitsDisInsShortTermFullAmount.disabled = false;
	}
	else
	{
		form.frmBenefitsDisInsShortTermFullPrcnt.value = 0;
		form.frmBenefitsDisInsShortTermFullAmount.value = 0;
		form.frmBenefitsDisInsShortTermFullPrcnt.disabled = true;
		form.frmBenefitsDisInsShortTermFullAmount.disabled = true;
	}
	if (form.frmBenefitsDisInsShortTermPartPaid.checked == true)
	{
		form.frmBenefitsDisInsShortTermPartPrcnt.disabled = false;
		form.frmBenefitsDisInsShortTermPartAmount.disabled = false;
	}
	else
	{
		form.frmBenefitsDisInsShortTermPartPrcnt.value = 0;
		form.frmBenefitsDisInsShortTermPartAmount.value = 0;
		form.frmBenefitsDisInsShortTermPartPrcnt.disabled = true;
		form.frmBenefitsDisInsShortTermPartAmount.disabled = true;
	}
	//Long Term
	if (form.frmBenefitsDisInsLongTermFullPaid.checked == true)
	{
		form.frmBenefitsDisInsLongTermFullPrcnt.disabled = false;
		form.frmBenefitsDisInsLongTermFullAmount.disabled = false;
	}
	else
	{
		form.frmBenefitsDisInsLongTermFullPrcnt.value = 0;
		form.frmBenefitsDisInsLongTermFullAmount.value = 0;
		form.frmBenefitsDisInsLongTermFullPrcnt.disabled = true;
		form.frmBenefitsDisInsLongTermFullAmount.disabled = true;
	}
	if (form.frmBenefitsDisInsLongTermPartPaid.checked == true)
	{
		form.frmBenefitsDisInsLongTermPartPrcnt.disabled = false;
		form.frmBenefitsDisInsLongTermPartAmount.disabled = false;
	}
	else
	{
		form.frmBenefitsDisInsLongTermPartPrcnt.value = 0;
		form.frmBenefitsDisInsLongTermPartAmount.value = 0;
		form.frmBenefitsDisInsLongTermPartPrcnt.disabled = true;
		form.frmBenefitsDisInsLongTermPartAmount.disabled = true;
	}
	//EAP
	if (form.frmBenefitsEAPFullPaid.checked == true)
	{
		form.frmBenefitsEAPFullPrcnt.disabled = false;
		form.frmBenefitsEAPFullAmount.disabled = false;
	}
	else
	{
		form.frmBenefitsEAPFullPrcnt.value = 0;
		form.frmBenefitsEAPFullAmount.value = 0;
		form.frmBenefitsEAPFullPrcnt.disabled = true;
		form.frmBenefitsEAPFullAmount.disabled = true;
	}
	if (form.frmBenefitsEAPPartPaid.checked == true)
	{
		form.frmBenefitsEAPPartPrcnt.disabled = false;
		form.frmBenefitsEAPPartAmount.disabled = false;
	}
	else
	{
		form.frmBenefitsEAPPartPrcnt.value = 0;
		form.frmBenefitsEAPPartAmount.value = 0;
		form.frmBenefitsEAPPartPrcnt.disabled = true;
		form.frmBenefitsEAPPartAmount.disabled = true;
	}
	//HealthClub
	if (form.frmBenefitsHealthClubFullPaid.checked == true)
	{
		form.frmBenefitsHealthClubFullPrcnt.disabled = false;
		form.frmBenefitsHealthClubFullAmount.disabled = false;
	}
	else
	{
		form.frmBenefitsHealthClubFullPrcnt.value = 0;
		form.frmBenefitsHealthClubFullAmount.value = 0;
		form.frmBenefitsHealthClubFullPrcnt.disabled = true;
		form.frmBenefitsHealthClubFullAmount.disabled = true;
	}
	if (form.frmBenefitsHealthClubPartPaid.checked == true)
	{
		form.frmBenefitsHealthClubPartPrcnt.disabled = false;
		form.frmBenefitsHealthClubPartAmount.disabled = false;
	}
	else
	{
		form.frmBenefitsHealthClubPartPrcnt.value = 0;
		form.frmBenefitsHealthClubPartAmount.value = 0;
		form.frmBenefitsHealthClubPartPrcnt.disabled = true;
		form.frmBenefitsHealthClubPartAmount.disabled = true;
	}
	//Life Insurance
	if (form.frmBenefitsLifeInsuranceFullPaid.checked == true)
	{
		form.frmBenefitsLifeInsuranceFullPrcnt.disabled = false;
		form.frmBenefitsLifeInsuranceFullAmount.disabled = false;
	}
	else
	{
		form.frmBenefitsLifeInsuranceFullPrcnt.value = 0;
		form.frmBenefitsLifeInsuranceFullAmount.value = 0;
		form.frmBenefitsLifeInsuranceFullPrcnt.disabled = true;
		form.frmBenefitsLifeInsuranceFullAmount.disabled = true;
	}
	if (form.frmBenefitsLifeInsurancePartPaid.checked == true)
	{
		form.frmBenefitsLifeInsurancePartPrcnt.disabled = false;
		form.frmBenefitsLifeInsurancePartAmount.disabled = false;
	}
	else
	{
		form.frmBenefitsLifeInsurancePartPrcnt.value = 0;
		form.frmBenefitsLifeInsurancePartAmount.value = 0;
		form.frmBenefitsLifeInsurancePartPrcnt.disabled = true;
		form.frmBenefitsLifeInsurancePartAmount.disabled = true;
	}
	//Prof Dues
	if (form.frmBenefitsProfDuesFullPaid.checked == true)
	{
		form.frmBenefitsProfDuesFullAmount.disabled = false;
	}
	else
	{
		form.frmBenefitsProfDuesFullAmount.value = 0;
		form.frmBenefitsProfDuesFullAmount.disabled = true;
	}
	if (form.frmBenefitsProfDuesPartPaid.checked == true)
	{
		form.frmBenefitsProfDuesPartAmount.disabled = false;
	}
	else
	{
		form.frmBenefitsProfDuesPartAmount.value = 0;
		form.frmBenefitsProfDuesPartAmount.disabled = true;
	}
	//403B
	if (form.frmBenefits403BFullContrib.value == 'true')
	{
		form.frmBenefits403BFullContribPrcnt.disabled = false;
	}
	else
	{
		form.frmBenefits403BFullContribPrcnt.value = 0;
		form.frmBenefits403BFullContribPrcnt.disabled = true;
	}
	if (form.frmBenefits403BPartContrib.value == 'true')
	{
		form.frmBenefits403BPartContribPrcnt.disabled = false;
	}
	else
	{
		form.frmBenefits403BPartContribPrcnt.value = 0;
		form.frmBenefits403BPartContribPrcnt.disabled = true;
	}
	
}
function CheckForOffer(form)
{
	//Madical
	if (form.frmBenefitsBenMedOffered.value == 'true')
	{
		form.frmBenefitsBenMedFullEmployee.disabled = false;
		form.frmBenefitsBenMedFullEmployeeAmount.disabled = false;
		form.frmBenefitsBenMedFullFamily.disabled = false;
		form.frmBenefitsBenMedFullFamilyAmount.disabled = false;
		form.frmBenefitsBenMedPartEmployee.disabled = false;
		form.frmBenefitsBenMedPartEmployeeAmount.disabled = false;
		form.frmBenefitsBenMedPartFamily.disabled = false;
		form.frmBenefitsBenMedPartFamilyAmount.disabled = false;
	}
	else
	{
		form.frmBenefitsBenMedFullEmployee.value = 0;
		form.frmBenefitsBenMedFullEmployeeAmount.value = 0;
		form.frmBenefitsBenMedFullFamily.value = 0;
		form.frmBenefitsBenMedFullFamilyAmount.value = 0;
		form.frmBenefitsBenMedPartEmployee.value = 0;
		form.frmBenefitsBenMedPartEmployeeAmount.value = 0;
		form.frmBenefitsBenMedPartFamily.value = 0;
		form.frmBenefitsBenMedPartFamilyAmount.value = 0;
		
		form.frmBenefitsBenMedFullEmployee.disabled = true;
		form.frmBenefitsBenMedFullEmployeeAmount.disabled = true;
		form.frmBenefitsBenMedFullFamily.disabled = true;
		form.frmBenefitsBenMedFullFamilyAmount.disabled = true;
		form.frmBenefitsBenMedPartEmployee.disabled = true;
		form.frmBenefitsBenMedPartEmployeeAmount.disabled = true;
		form.frmBenefitsBenMedPartFamily.disabled = true;
		form.frmBenefitsBenMedPartFamilyAmount.disabled = true;
	}
	//Dental
	if (form.frmBenefitsBenDentOffered.value == 'true')
	{
		form.frmBenefitsBenDentFullEmployee.disabled = false;
		form.frmBenefitsBenDentFullEmployeeAmount.disabled = false;
		form.frmBenefitsBenDentFullFamily.disabled = false;
		form.frmBenefitsBenDentFullFamilyAmount.disabled = false;
		form.frmBenefitsBenDentPartEmployee.disabled = false;
		form.frmBenefitsBenDentPartEmployeeAmount.disabled = false;
		form.frmBenefitsBenDentPartFamily.disabled = false;
		form.frmBenefitsBenDentPartFamilyAmount.disabled = false;
	}
	else
	{
		form.frmBenefitsBenDentFullEmployee.value = 0;
		form.frmBenefitsBenDentFullEmployeeAmount.value = 0;
		form.frmBenefitsBenDentFullFamily.value = 0;
		form.frmBenefitsBenDentFullFamilyAmount.value = 0;
		form.frmBenefitsBenDentPartEmployee.value = 0;
		form.frmBenefitsBenDentPartEmployeeAmount.value = 0;
		form.frmBenefitsBenDentPartFamily.value = 0;
		form.frmBenefitsBenDentPartFamilyAmount.value = 0;
		
		form.frmBenefitsBenDentFullEmployee.disabled = true;
		form.frmBenefitsBenDentFullEmployeeAmount.disabled = true;
		form.frmBenefitsBenDentFullFamily.disabled = true;
		form.frmBenefitsBenDentFullFamilyAmount.disabled = true;
		form.frmBenefitsBenDentPartEmployee.disabled = true;
		form.frmBenefitsBenDentPartEmployeeAmount.disabled = true;
		form.frmBenefitsBenDentPartFamily.disabled = true;
		form.frmBenefitsBenDentPartFamilyAmount.disabled = true;
	}
	//Vision
	if (form.frmBenefitsBenVisOffered.value == 'true')
	{
		form.frmBenefitsBenVisFullEmployee.disabled = false;
		form.frmBenefitsBenVisFullEmployeeAmount.disabled = false;
		form.frmBenefitsBenVisFullFamily.disabled = false;
		form.frmBenefitsBenVisFullFamilyAmount.disabled = false;
		form.frmBenefitsBenVisPartEmployee.disabled = false;
		form.frmBenefitsBenVisPartEmployeeAmount.disabled = false;
		form.frmBenefitsBenVisPartFamily.disabled = false;
		form.frmBenefitsBenVisPartFamilyAmount.disabled = false;
	}
	else
	{
		form.frmBenefitsBenVisFullEmployee.value = 0;
		form.frmBenefitsBenVisFullEmployeeAmount.value = 0;
		form.frmBenefitsBenVisFullFamily.value = 0;
		form.frmBenefitsBenVisFullFamilyAmount.value = 0;
		form.frmBenefitsBenVisPartEmployee.value = 0;
		form.frmBenefitsBenVisPartEmployeeAmount.value = 0;
		form.frmBenefitsBenVisPartFamily.value = 0;
		form.frmBenefitsBenVisPartFamilyAmount.value = 0;
		
		form.frmBenefitsBenVisFullEmployee.disabled = true;
		form.frmBenefitsBenVisFullEmployeeAmount.disabled = true;
		form.frmBenefitsBenVisFullFamily.disabled = true;
		form.frmBenefitsBenVisFullFamilyAmount.disabled = true;
		form.frmBenefitsBenVisPartEmployee.disabled = true;
		form.frmBenefitsBenVisPartEmployeeAmount.disabled = true;
		form.frmBenefitsBenVisPartFamily.disabled = true;
		form.frmBenefitsBenVisPartFamilyAmount.disabled = true;
	}
	//Short Term
	if (form.frmBenefitsDisInsShortTermFull.checked == true)
	{
		form.frmBenefitsDisInsShortTermFullPaid.disabled = false;
		if (form.frmBenefitsDisInsShortTermFullPaid.checked == false)
		{
			form.frmBenefitsDisInsShortTermFullPrcnt.disabled = true;
			form.frmBenefitsDisInsShortTermFullAmount.disabled = true;
		}
	}
	else
	{
		form.frmBenefitsDisInsShortTermFullPaid.checked = false;
		form.frmBenefitsDisInsShortTermFullPrcnt.value = 0;
		form.frmBenefitsDisInsShortTermFullAmount.value = 0;
		form.frmBenefitsDisInsShortTermFullPaid.disabled = true;
		form.frmBenefitsDisInsShortTermFullPrcnt.disabled = true;
		form.frmBenefitsDisInsShortTermFullAmount.disabled = true;
	}
	if (form.frmBenefitsDisInsShortTermPart.checked == true)
	{
		form.frmBenefitsDisInsShortTermPartPaid.disabled = false;
		if (form.frmBenefitsDisInsShortTermPartPaid.checked == false)
		{
			form.frmBenefitsDisInsShortTermPartPrcnt.disabled = true;
			form.frmBenefitsDisInsShortTermPartAmount.disabled = true;
		}
	}
	else
	{
		form.frmBenefitsDisInsShortTermPartPaid.checked = false;
		form.frmBenefitsDisInsShortTermPartPrcnt.value = 0;
		form.frmBenefitsDisInsShortTermPartAmount.value = 0;
		form.frmBenefitsDisInsShortTermPartPaid.disabled = true;
		form.frmBenefitsDisInsShortTermPartPrcnt.disabled = true;
		form.frmBenefitsDisInsShortTermPartAmount.disabled = true;
	}
	//Long Term
	if (form.frmBenefitsDisInsLongTermFull.checked == true)
	{
		form.frmBenefitsDisInsLongTermFullPaid.disabled = false;
		if (form.frmBenefitsDisInsLongTermFullPaid.checked == false)
		{
			form.frmBenefitsDisInsLongTermFullPrcnt.disabled = true;
			form.frmBenefitsDisInsLongTermFullAmount.disabled = true;
		}
	}
	else
	{
		form.frmBenefitsDisInsLongTermFullPaid.checked = false;
		form.frmBenefitsDisInsLongTermFullPrcnt.value = 0;
		form.frmBenefitsDisInsLongTermFullAmount.value = 0;
		form.frmBenefitsDisInsLongTermFullPaid.disabled = true;
		form.frmBenefitsDisInsLongTermFullPrcnt.disabled = true;
		form.frmBenefitsDisInsLongTermFullAmount.disabled = true;
	}
	if (form.frmBenefitsDisInsLongTermPart.checked == true)
	{
		form.frmBenefitsDisInsLongTermPartPaid.disabled = false;
		if (form.frmBenefitsDisInsLongTermPartPaid.checked == false)
		{
			form.frmBenefitsDisInsLongTermPartPrcnt.disabled = true;
			form.frmBenefitsDisInsLongTermPartAmount.disabled = true;
		}
	}
	else
	{
		form.frmBenefitsDisInsLongTermPartPaid.checked = false;
		form.frmBenefitsDisInsLongTermPartPrcnt.value = 0;
		form.frmBenefitsDisInsLongTermPartAmount.value = 0;
		form.frmBenefitsDisInsLongTermPartPaid.disabled = true;
		form.frmBenefitsDisInsLongTermPartPrcnt.disabled = true;
		form.frmBenefitsDisInsLongTermPartAmount.disabled = true;
	}
	//EAP
	if (form.frmBenefitsEAPFull.checked == true)
	{
		form.frmBenefitsEAPFullPaid.disabled = false;
		if (form.frmBenefitsEAPFullPaid.checked == false)
		{
			form.frmBenefitsEAPFullPrcnt.disabled = true;
			form.frmBenefitsEAPFullAmount.disabled = true;
		}
	}
	else
	{
		form.frmBenefitsEAPFullPaid.checked = false;
		form.frmBenefitsEAPFullPrcnt.value = 0;
		form.frmBenefitsEAPFullAmount.value = 0;
		form.frmBenefitsEAPFullPaid.disabled = true;
		form.frmBenefitsEAPFullPrcnt.disabled = true;
		form.frmBenefitsEAPFullAmount.disabled = true;
	}
	if (form.frmBenefitsEAPPart.checked == true)
	{
		form.frmBenefitsEAPPartPaid.disabled = false;
		if (form.frmBenefitsEAPPartPaid.checked == false)
		{
			form.frmBenefitsEAPPartPrcnt.disabled = true;
			form.frmBenefitsEAPPartAmount.disabled = true;
		}
	}
	else
	{
		form.frmBenefitsEAPPartPaid.checked = false;
		form.frmBenefitsEAPPartPrcnt.value = 0;
		form.frmBenefitsEAPPartAmount.value = 0;
		form.frmBenefitsEAPPartPaid.disabled = true;
		form.frmBenefitsEAPPartPrcnt.disabled = true;
		form.frmBenefitsEAPPartAmount.disabled = true;
	}
	//Health Club
	if (form.frmBenefitsHealthClubFull.checked == true)
	{
		form.frmBenefitsHealthClubFullPaid.disabled = false;
		if (form.frmBenefitsHealthClubFullPaid.checked == false)
		{
			form.frmBenefitsHealthClubFullPrcnt.disabled = true;
			form.frmBenefitsHealthClubFullAmount.disabled = true;
		}
	}
	else
	{
		form.frmBenefitsHealthClubFullPaid.checked = false;
		form.frmBenefitsHealthClubFullPrcnt.value = 0;
		form.frmBenefitsHealthClubFullAmount.value = 0;
		form.frmBenefitsHealthClubFullPaid.disabled = true;
		form.frmBenefitsHealthClubFullPrcnt.disabled = true;
		form.frmBenefitsHealthClubFullAmount.disabled = true;
	}
	if (form.frmBenefitsHealthClubPart.checked == true)
	{
		form.frmBenefitsHealthClubPartPaid.disabled = false;
		if (form.frmBenefitsHealthClubPartPaid.checked == false)
		{
			form.frmBenefitsHealthClubPartPrcnt.disabled = true;
			form.frmBenefitsHealthClubPartAmount.disabled = true;
		}
	}
	else
	{
		form.frmBenefitsHealthClubPartPaid.checked = false;
		form.frmBenefitsHealthClubPartPrcnt.value = 0;
		form.frmBenefitsHealthClubPartAmount.value = 0;
		form.frmBenefitsHealthClubPartPaid.disabled = true;
		form.frmBenefitsHealthClubPartPrcnt.disabled = true;
		form.frmBenefitsHealthClubPartAmount.disabled = true;
	}
	//Life Insurance
	if (form.frmBenefitsLifeInsuranceFull.checked == true)
	{
		form.frmBenefitsLifeInsuranceFullPaid.disabled = false;
		if (form.frmBenefitsLifeInsuranceFullPaid.checked == false)
		{
			form.frmBenefitsLifeInsuranceFullPrcnt.disabled = true;
			form.frmBenefitsLifeInsuranceFullAmount.disabled = true;
		}
	}
	else
	{
		form.frmBenefitsLifeInsuranceFullPaid.checked = false;
		form.frmBenefitsLifeInsuranceFullPrcnt.value = 0;
		form.frmBenefitsLifeInsuranceFullAmount.value = 0;
		form.frmBenefitsLifeInsuranceFullPaid.disabled = true;
		form.frmBenefitsLifeInsuranceFullPrcnt.disabled = true;
		form.frmBenefitsLifeInsuranceFullAmount.disabled = true;
	}
	if (form.frmBenefitsLifeInsurancePart.checked == true)
	{
		form.frmBenefitsLifeInsurancePartPaid.disabled = false;
		if (form.frmBenefitsLifeInsurancePartPaid.checked == false)
		{
			form.frmBenefitsLifeInsurancePartPrcnt.disabled = true;
			form.frmBenefitsLifeInsurancePartAmount.disabled = true;
		}
	}
	else
	{
		form.frmBenefitsLifeInsurancePartPaid.checked = false;
		form.frmBenefitsLifeInsurancePartPrcnt.value = 0;
		form.frmBenefitsLifeInsurancePartAmount.value = 0;
		form.frmBenefitsLifeInsurancePartPaid.disabled = true;
		form.frmBenefitsLifeInsurancePartPrcnt.disabled = true;
		form.frmBenefitsLifeInsurancePartAmount.disabled = true;
	}
	//Vacation days
	if (form.frmBenefitsTimeOffFull.checked == true)
	{
			form.frmBenefitsTimeOffFullDays.disabled = false;
			form.frmBenefitsTimeOffFullYears.disabled = false;
			form.frmBenefitsTimeOffFullDaysIncreased.disabled = false;
	}
	else
	{
		form.frmBenefitsTimeOffFullDays.value = 0;
		form.frmBenefitsTimeOffFullYears.value = 0;
		form.frmBenefitsTimeOffFullDaysIncreased.value = 0;
		form.frmBenefitsTimeOffFullDays.disabled = true;
		form.frmBenefitsTimeOffFullYears.disabled = true;
		form.frmBenefitsTimeOffFullDaysIncreased.disabled = true;
	}
	if (form.frmBenefitsTimeOffFullNExempt.checked == true)
	{
			form.frmBenefitsTimeOffFullDaysNExempt.disabled = false;
			form.frmBenefitsTimeOffFullYearsNExempt.disabled = false;
			form.frmBenefitsTimeOffFullDaysIncreasedNExempt.disabled = false;
	}
	else
	{
		form.frmBenefitsTimeOffFullDaysNExempt.value = 0;
		form.frmBenefitsTimeOffFullYearsNExempt.value = 0;
		form.frmBenefitsTimeOffFullDaysIncreasedNExempt.value = 0;
		form.frmBenefitsTimeOffFullDaysNExempt.disabled = true;
		form.frmBenefitsTimeOffFullYearsNExempt.disabled = true;
		form.frmBenefitsTimeOffFullDaysIncreasedNExempt.disabled = true;
	}
	if (form.frmBenefitsTimeOffPart.checked == true)
	{
			form.frmBenefitsTimeOffPartDays.disabled = false;
			form.frmBenefitsTimeOffPartYears.disabled = false;
			form.frmBenefitsTimeOffPartDaysIncreased.disabled = false;
	}
	else
	{
		form.frmBenefitsTimeOffPartDays.value = 0;
		form.frmBenefitsTimeOffPartYears.value = 0;
		form.frmBenefitsTimeOffPartDaysIncreased.value = 0;
		form.frmBenefitsTimeOffPartDays.disabled = true;
		form.frmBenefitsTimeOffPartYears.disabled = true;
		form.frmBenefitsTimeOffPartDaysIncreased.disabled = true;
	}
	if (form.frmBenefitsTimeOffPartNExempt.checked == true)
	{
			form.frmBenefitsTimeOffPartDaysNExempt.disabled = false;
			form.frmBenefitsTimeOffPartYearsNExempt.disabled = false;
			form.frmBenefitsTimeOffPartDaysIncreasedNExempt.disabled = false;
	}
	else
	{
		form.frmBenefitsTimeOffPartDaysNExempt.value = 0;
		form.frmBenefitsTimeOffPartYearsNExempt.value = 0;
		form.frmBenefitsTimeOffPartDaysIncreasedNExempt.value = 0;
		form.frmBenefitsTimeOffPartDaysNExempt.disabled = true;
		form.frmBenefitsTimeOffPartYearsNExempt.disabled = true;
		form.frmBenefitsTimeOffPartDaysIncreasedNExempt.disabled = true;
	}
	//Sick days
	if (form.frmBenefitsTimeOffSickFull.checked == true)
	{
			form.frmBenefitsTimeOffSickFullDays.disabled = false;
	}
	else
	{
		form.frmBenefitsTimeOffSickFullDays.value = 0;
		form.frmBenefitsTimeOffSickFullDays.disabled = true;
	}
	if (form.frmBenefitsTimeOffSickPart.checked == true)
	{
			form.frmBenefitsTimeOffSickPartDays.disabled = false;
	}
	else
	{
		form.frmBenefitsTimeOffSickPartDays.value = 0;
		form.frmBenefitsTimeOffSickPartDays.disabled = true;
	}
	//Prof Dues
	if (form.frmBenefitsProfDuesFull.checked == true)
	{
		form.frmBenefitsProfDuesFullPaid.disabled = false;
		if (form.frmBenefitsProfDuesFullPaid.checked == false)
		{
			form.frmBenefitsProfDuesFullAmount.disabled = true;
		}
	}
	else
	{
		form.frmBenefitsProfDuesFullPaid.checked = false;
		form.frmBenefitsProfDuesFullAmount.value = 0;
		form.frmBenefitsProfDuesFullPaid.disabled = true;
		form.frmBenefitsProfDuesFullAmount.disabled = true;
	}
	if (form.frmBenefitsProfDuesPart.checked == true)
	{
		form.frmBenefitsProfDuesPartPaid.disabled = false;
		if (form.frmBenefitsProfDuesPartPaid.checked == false)
		{
			form.frmBenefitsProfDuesPartAmount.disabled = true;
		}
	}
	else
	{
		form.frmBenefitsProfDuesPartPaid.checked = false;
		form.frmBenefitsProfDuesPartAmount.value = 0;
		form.frmBenefitsProfDuesPartPaid.disabled = true;
		form.frmBenefitsProfDuesPartAmount.disabled = true;
	}
	//Retirement
	if (form.frmBenefitsRetirementFull.checked == true)
	{
			form.frmBenefitsRetirementFullPrcnt.disabled = false;
	}
	else
	{
		form.frmBenefitsRetirementFullPrcnt.value = 0;
		form.frmBenefitsRetirementFullPrcnt.disabled = true;
	}
	if (form.frmBenefitsRetirementPart.checked == true)
	{
			form.frmBenefitsRetirementPartPrcnt.disabled = false;
	}
	else
	{
		form.frmBenefitsRetirementPartPrcnt.value = 0;
		form.frmBenefitsRetirementPartPrcnt.disabled = true;
	}
	//403B
	if (form.frmBenefits403BFull.checked == true)
	{
		form.frmBenefits403BFullContrib.disabled = false;
		if (form.frmBenefits403BFullContrib.value == 'false')
		{
			form.frmBenefits403BFullContribPrcnt.disabled = true;
		}
	}
	else
	{
		form.frmBenefits403BFullContrib.value = 'false';
		form.frmBenefits403BFullContribPrcnt.value = 0;
		form.frmBenefits403BFullContrib.disabled = true;
		form.frmBenefits403BFullContribPrcnt.disabled = true;
	}
	if (form.frmBenefits403BPart.checked == true)
	{
		form.frmBenefits403BPartContrib.disabled = false;
		if (form.frmBenefits403BPartContrib.value == 'false')
		{
			form.frmBenefits403BPartContribPrcnt.disabled = true;
		}
	}
	else
	{
		form.frmBenefits403BPartContrib.value = 'false';
		form.frmBenefits403BPartContribPrcnt.value = 0;
		form.frmBenefits403BPartContrib.disabled = true;
		form.frmBenefits403BPartContribPrcnt.disabled = true;
	}
	//Telecomm
	if (form.frmBenefitsTelecommFull.checked == true)
	{
			form.frmBenefitsTelecommFullCount.disabled = false;
			form.frmBenefitsTelecommFullPrcnt.disabled = false;
	}
	else
	{
		form.frmBenefitsTelecommFullCount.value = 0;
		form.frmBenefitsTelecommFullCount.disabled = true;
		form.frmBenefitsTelecommFullPrcnt.value = 0;
		form.frmBenefitsTelecommFullPrcnt.disabled = true;
	}
	if (form.frmBenefitsTelecommPart.checked == true)
	{
			form.frmBenefitsTelecommPartCount.disabled = false;
			form.frmBenefitsTelecommPartPrcnt.disabled = false;
	}
	else
	{
		form.frmBenefitsTelecommPartCount.value = 0;
		form.frmBenefitsTelecommPartCount.disabled = true;
		form.frmBenefitsTelecommPartPrcnt.value = 0;
		form.frmBenefitsTelecommPartPrcnt.disabled = true;
	}
	//Tuition
	if (form.frmBenefitsTuitionFull.checked == true)
	{
			form.frmBenefitsTuitionFullAmount.disabled = false;
	}
	else
	{
		form.frmBenefitsTuitionFullAmount.value = 0;
		form.frmBenefitsTuitionFullAmount.disabled = true;
	}
	if (form.frmBenefitsTuitionPart.checked == true)
	{
			form.frmBenefitsTuitionPartAmount.disabled = false;
	}
	else
	{
		form.frmBenefitsTuitionPartAmount.value = 0;
		form.frmBenefitsTuitionPartAmount.disabled = true;
	}
}

addEvent(window, 'load', function()
{
 //Medical
 if (document.getElementById('frmBenefitsBenMedOffered').value == 'false'){
	document.getElementById('frmBenefitsBenMedFullEmployee').disabled = true;
	document.getElementById('frmBenefitsBenMedFullEmployeeAmount').disabled = true;
	document.getElementById('frmBenefitsBenMedFullFamily').disabled = true;
	document.getElementById('frmBenefitsBenMedFullFamilyAmount').disabled = true;
	document.getElementById('frmBenefitsBenMedPartEmployee').disabled = true;
	document.getElementById('frmBenefitsBenMedPartEmployeeAmount').disabled = true;
	document.getElementById('frmBenefitsBenMedPartFamily').disabled = true;
	document.getElementById('frmBenefitsBenMedPartFamilyAmount').disabled = true;
 }
 //Dental
 if (document.getElementById('frmBenefitsBenDentOffered').value == 'false'){
	document.getElementById('frmBenefitsBenDentFullEmployee').disabled = true;
	document.getElementById('frmBenefitsBenDentFullEmployeeAmount').disabled = true;
	document.getElementById('frmBenefitsBenDentFullFamily').disabled = true;
	document.getElementById('frmBenefitsBenDentFullFamilyAmount').disabled = true;
	document.getElementById('frmBenefitsBenDentPartEmployee').disabled = true;
	document.getElementById('frmBenefitsBenDentPartEmployeeAmount').disabled = true;
	document.getElementById('frmBenefitsBenDentPartFamily').disabled = true;
	document.getElementById('frmBenefitsBenDentPartFamilyAmount').disabled = true;
 }
 //Vision
 if (document.getElementById('frmBenefitsBenVisOffered').value == 'false'){
	document.getElementById('frmBenefitsBenVisFullEmployee').disabled = true;
	document.getElementById('frmBenefitsBenVisFullEmployeeAmount').disabled = true;
	document.getElementById('frmBenefitsBenVisFullFamily').disabled = true;
	document.getElementById('frmBenefitsBenVisFullFamilyAmount').disabled = true;
	document.getElementById('frmBenefitsBenVisPartEmployee').disabled = true;
	document.getElementById('frmBenefitsBenVisPartEmployeeAmount').disabled = true;
	document.getElementById('frmBenefitsBenVisPartFamily').disabled = true;
	document.getElementById('frmBenefitsBenVisPartFamilyAmount').disabled = true;
 }
 //Short Term
 if (document.getElementById('frmBenefitsDisInsShortTermFullPaid').checked == false){
	document.getElementById('frmBenefitsDisInsShortTermFullPrcnt').disabled = true;
	document.getElementById('frmBenefitsDisInsShortTermFullAmount').disabled = true;
 }
 if(document.getElementById('frmBenefitsDisInsShortTermFull').checked == false){
	document.getElementById('frmBenefitsDisInsShortTermFullPaid').disabled = true;
	document.getElementById('frmBenefitsDisInsShortTermFullPrcnt').disabled = true;
	document.getElementById('frmBenefitsDisInsShortTermFullAmount').disabled = true;
 }
 if (document.getElementById('frmBenefitsDisInsShortTermPartPaid').checked == false){
	document.getElementById('frmBenefitsDisInsShortTermPartPrcnt').disabled = true;
	document.getElementById('frmBenefitsDisInsShortTermPartAmount').disabled = true;
 }
 if(document.getElementById('frmBenefitsDisInsShortTermPart').checked == false){
	document.getElementById('frmBenefitsDisInsShortTermPartPaid').disabled = true;
	document.getElementById('frmBenefitsDisInsShortTermPartPrcnt').disabled = true;
	document.getElementById('frmBenefitsDisInsShortTermPartAmount').disabled = true;
 }
//Long Term
 if (document.getElementById('frmBenefitsDisInsLongTermFullPaid').checked == false){
	document.getElementById('frmBenefitsDisInsLongTermFullPrcnt').disabled = true;
	document.getElementById('frmBenefitsDisInsLongTermFullAmount').disabled = true;
 }
 if(document.getElementById('frmBenefitsDisInsLongTermFull').checked == false){
	document.getElementById('frmBenefitsDisInsLongTermFullPaid').disabled = true;
	document.getElementById('frmBenefitsDisInsLongTermFullPrcnt').disabled = true;
	document.getElementById('frmBenefitsDisInsLongTermFullAmount').disabled = true;
 }
 if (document.getElementById('frmBenefitsDisInsLongTermPartPaid').checked == false){
	document.getElementById('frmBenefitsDisInsLongTermPartPrcnt').disabled = true;
	document.getElementById('frmBenefitsDisInsLongTermPartAmount').disabled = true;
 }
 if(document.getElementById('frmBenefitsDisInsLongTermPart').checked == false){
	document.getElementById('frmBenefitsDisInsLongTermPartPaid').disabled = true;
	document.getElementById('frmBenefitsDisInsLongTermPartPrcnt').disabled = true;
	document.getElementById('frmBenefitsDisInsLongTermPartAmount').disabled = true;
 }
 //EAP
 if (document.getElementById('frmBenefitsEAPFullPaid').checked == false){
	document.getElementById('frmBenefitsEAPFullPrcnt').disabled = true;
	document.getElementById('frmBenefitsEAPFullAmount').disabled = true;
 }
 if(document.getElementById('frmBenefitsEAPFull').checked == false){
	document.getElementById('frmBenefitsEAPFullPaid').disabled = true;
	document.getElementById('frmBenefitsEAPFullPrcnt').disabled = true;
	document.getElementById('frmBenefitsEAPFullAmount').disabled = true;
 }
 if (document.getElementById('frmBenefitsEAPPartPaid').checked == false){
	document.getElementById('frmBenefitsEAPPartPrcnt').disabled = true;
	document.getElementById('frmBenefitsEAPPartAmount').disabled = true;
 }
 if(document.getElementById('frmBenefitsEAPPart').checked == false){
	document.getElementById('frmBenefitsEAPPartPaid').disabled = true;
	document.getElementById('frmBenefitsEAPPartPrcnt').disabled = true;
	document.getElementById('frmBenefitsEAPPartAmount').disabled = true;
 }
 //HealthClub
 if (document.getElementById('frmBenefitsHealthClubFullPaid').checked == false){
	document.getElementById('frmBenefitsHealthClubFullPrcnt').disabled = true;
	document.getElementById('frmBenefitsHealthClubFullAmount').disabled = true;
 }
 if(document.getElementById('frmBenefitsHealthClubFull').checked == false){
	document.getElementById('frmBenefitsHealthClubFullPaid').disabled = true;
	document.getElementById('frmBenefitsHealthClubFullPrcnt').disabled = true;
	document.getElementById('frmBenefitsHealthClubFullAmount').disabled = true;
 }
 if (document.getElementById('frmBenefitsHealthClubPartPaid').checked == false){
	document.getElementById('frmBenefitsHealthClubPartPrcnt').disabled = true;
	document.getElementById('frmBenefitsHealthClubPartAmount').disabled = true;
 }
 if(document.getElementById('frmBenefitsHealthClubPart').checked == false){
	document.getElementById('frmBenefitsHealthClubPartPaid').disabled = true;
	document.getElementById('frmBenefitsHealthClubPartPrcnt').disabled = true;
	document.getElementById('frmBenefitsHealthClubPartAmount').disabled = true;
 }
 //LifeInsurance
 if (document.getElementById('frmBenefitsLifeInsuranceFullPaid').checked == false){
	document.getElementById('frmBenefitsLifeInsuranceFullPrcnt').disabled = true;
	document.getElementById('frmBenefitsLifeInsuranceFullAmount').disabled = true;
 }
 if(document.getElementById('frmBenefitsLifeInsuranceFull').checked == false){
	document.getElementById('frmBenefitsLifeInsuranceFullPaid').disabled = true;
	document.getElementById('frmBenefitsLifeInsuranceFullPrcnt').disabled = true;
	document.getElementById('frmBenefitsLifeInsuranceFullAmount').disabled = true;
 }
 if (document.getElementById('frmBenefitsLifeInsurancePartPaid').checked == false){
	document.getElementById('frmBenefitsLifeInsurancePartPrcnt').disabled = true;
	document.getElementById('frmBenefitsLifeInsurancePartAmount').disabled = true;
 }
 if(document.getElementById('frmBenefitsLifeInsurancePart').checked == false){
	document.getElementById('frmBenefitsLifeInsurancePartPaid').disabled = true;
	document.getElementById('frmBenefitsLifeInsurancePartPrcnt').disabled = true;
	document.getElementById('frmBenefitsLifeInsurancePartAmount').disabled = true;
 }
 //Vacation
 if (document.getElementById('frmBenefitsTimeOffFull').checked == false){
	document.getElementById('frmBenefitsTimeOffFullDays').disabled = true;
	document.getElementById('frmBenefitsTimeOffFullYears').disabled = true;
	document.getElementById('frmBenefitsTimeOffFullDaysIncreased').disabled = true;
 }
 if (document.getElementById('frmBenefitsTimeOffFullNExempt').checked == false){
	document.getElementById('frmBenefitsTimeOffFullDaysNExempt').disabled = true;
	document.getElementById('frmBenefitsTimeOffFullYearsNExempt').disabled = true;
	document.getElementById('frmBenefitsTimeOffFullDaysIncreasedNExempt').disabled = true;
 }
 if (document.getElementById('frmBenefitsTimeOffPart').checked == false){
	document.getElementById('frmBenefitsTimeOffPartDays').disabled = true;
	document.getElementById('frmBenefitsTimeOffPartYears').disabled = true;
	document.getElementById('frmBenefitsTimeOffPartDaysIncreased').disabled = true;
 }
 if (document.getElementById('frmBenefitsTimeOffPartNExempt').checked == false){
	document.getElementById('frmBenefitsTimeOffPartDaysNExempt').disabled = true;
	document.getElementById('frmBenefitsTimeOffPartYearsNExempt').disabled = true;
	document.getElementById('frmBenefitsTimeOffPartDaysIncreasedNExempt').disabled = true;
 }
 //Sick Time
 if (document.getElementById('frmBenefitsTimeOffSickFull').checked == false){
	document.getElementById('frmBenefitsTimeOffSickFullDays').disabled = true;
 }
 if (document.getElementById('frmBenefitsTimeOffSickPart').checked == false){
	document.getElementById('frmBenefitsTimeOffSickPartDays').disabled = true;
 }
 //Prof Dies
 if (document.getElementById('frmBenefitsProfDuesFullPaid').checked == false){
	document.getElementById('frmBenefitsProfDuesFullAmount').disabled = true;
 }
 if(document.getElementById('frmBenefitsProfDuesFull').checked == false){
	document.getElementById('frmBenefitsProfDuesFullPaid').disabled = true;
	document.getElementById('frmBenefitsProfDuesFullAmount').disabled = true;
 }
 if (document.getElementById('frmBenefitsProfDuesPartPaid').checked == false){
	document.getElementById('frmBenefitsProfDuesPartAmount').disabled = true;
 }
 if(document.getElementById('frmBenefitsProfDuesPart').checked == false){
	document.getElementById('frmBenefitsProfDuesPartPaid').disabled = true;
	document.getElementById('frmBenefitsProfDuesPartAmount').disabled = true;
 }
 //Pension Plan
 if (document.getElementById('frmBenefitsRetirementFull').checked == false){
	document.getElementById('frmBenefitsRetirementFullPrcnt').disabled = true;
 }
 if (document.getElementById('frmBenefitsRetirementPart').checked == false){
	document.getElementById('frmBenefitsRetirementPartPrcnt').disabled = true;
 }
 //403B
 if (document.getElementById('frmBenefits403BFullContrib').value == 'false'){
	document.getElementById('frmBenefits403BFullContribPrcnt').disabled = true;
 }
 if(document.getElementById('frmBenefits403BFull').checked == false){
	document.getElementById('frmBenefits403BFullContrib').disabled = true;
	document.getElementById('frmBenefits403BFullContribPrcnt').disabled = true;
 }
 if (document.getElementById('frmBenefits403BPartContrib').value == 'false'){
	document.getElementById('frmBenefits403BPartContribPrcnt').disabled = true;
 }
 if(document.getElementById('frmBenefits403BPart').checked == false){
	document.getElementById('frmBenefits403BPartContrib').disabled = true;
	document.getElementById('frmBenefits403BPartContribPrcnt').disabled = true;
 }
 //Telecomm
 if (document.getElementById('frmBenefitsTelecommFull').checked == false){
	document.getElementById('frmBenefitsTelecommFullCount').disabled = true;
	document.getElementById('frmBenefitsTelecommFullPrcnt').disabled = true;
 }
 if (document.getElementById('frmBenefitsTelecommPart').checked == false){
	document.getElementById('frmBenefitsTelecommPartCount').disabled = true;
	document.getElementById('frmBenefitsTelecommPartPrcnt').disabled = true;
 }
 //Tuition
 if (document.getElementById('frmBenefitsTuitionFull').checked == false){
	document.getElementById('frmBenefitsTuitionFullAmount').disabled = true;
 }
 if (document.getElementById('frmBenefitsTuitionPart').checked == false){
	document.getElementById('frmBenefitsTuitionPartAmount').disabled = true;
 }
}); 

function addEvent(obj, evType, fn)
{ 
 if (obj.addEventListener){ 
    obj.addEventListener(evType, fn, true); 
    return true; 
 }
 else if (obj.attachEvent){ 
    var r = obj.attachEvent("on"+evType, fn); 
    return r; 
 }
 else { 
    return false; 
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
<table border="1" cellspacing="0" cellpadding="1" width = "650" bordercolordark="#003063">

<form name="frmBenefits" action="benefits_edit.asp" method="post" onsubmit="return submitFormValidate(this)">
<!--<form name="frmBenefits" action="benefits_edit.asp" method="post">-->
<!--#include file="../includes/form_stamp.asp"-->
<script type="text/javascript">


</script>
<% 
If say = "edit" Then
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"
	query = "SELECT * FROM tbl_frmBenefits WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
	Set GetBenefits = Con.Execute(query)
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
		<td colspan="6" class="formHeader">BENEFITS</td>
	</tr>
	<tr>
		<td colspan="6" class="formMain"><font color="#ff0000"><div align="center"><strong>Please Note: </strong>After entering your information, you <strong>must</strong> click on the "Save Form" button at the bottom of the form and wait for the "Thank You" screen or your changes will be lost.</div></font></td>
	</tr>	
	<tr>
		<td colspan="6" class="formSubhead" align="center">If you need help with understanding the topic, please click on question-mark next to the topic of your interest</td>
	</tr>
	
	<tr>
	
		<td colspan="6">
			<table width="100%" border="1" bordercolordark="#003063" cellspacing="0" cellpadding="2">		
			
			<tr>
				<td colspan="6" class="formHeaderSmall">BENEFITS - MEDICAL<BR>% paid by BBBS & total benefit cost for Employee and Employee's Family</td>
			</tr>
			
			<tr>
				<td class="formMain">&nbsp;</td>
				<td class="formMain" align="center" colspan="2">Full Time</td>
				<td class="formMain" align="center" colspan="2">Part Time</td>				
			</tr>
			
			<tr>
				<td class="formMain" align="center" cellpadding="1">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" ID="Table9">
					<tr>
						<td class="formMain" valign="top" align="center">Medical&nbsp;
							<a href="../helpfiles/surveyhelp.asp?HelpID=BenMedOffered" onclick="NewWindow(this.href,'name','550','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="Center">
						<select name="frmBenefitsBenMedOffered" size=1 class="formMain" ID="Select1" onchange="CheckForOffer(this.form)">
							<%If say = "edit" Then%>
									<%If GetBenefits("BenMedOffered") then%>
										<option value="true">Yes</option>
										<option value="false">No</option>
									<%else%>
										<option value="false">No</option>
										<option value="true">Yes</option>
									<% End If %>
								<%Else%>
									<option value="selected">Offered?</option>
									<option value="true">Yes</option>
									<option value="false">No</option>
								<%End If%>
						</select>&nbsp;
						</td>
					</tr>
					</table>
				</td>
				<td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table1">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2">Employee</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("BenMedFullEmployee") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenMedFullEmployee" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("BenMedFullEmployeeAmount") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenMedFullEmployeeAmount" onchange="checkForWholeNumber(this.value)"><br></td>
						<td class="formMain" valign="top" align="left">Total monthly premium<br>(agency + EE contribution)</td>
					</tr>
					</table>
				</td>
				
				<td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table2">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2">Family</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("BenMedFullFamily") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenMedFullFamily" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("BenMedFullFamilyAmount") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenMedFullFamilyAmount" onchange="checkForWholeNumber(this.value)" ID="Text1"><br></td>
						<td class="formMain" valign="top" align="left">Total monthly premium<br>(agency + EE contribution)</td>
					</tr>
					</table>
				</td>
				
				<td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table3">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2">Employee</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("BenMedPartEmployee") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenMedPartEmployee" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("BenMedPartEmployeeAmount") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenMedPartEmployeeAmount" onchange="checkForWholeNumber(this.value)" ID="Text4"><br></td>
						<td class="formMain" valign="top" align="left">Total monthly premium<br>(agency + EE contribution)</td>
					</tr>
					</table>
				</td>		
				
				<td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table4">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2">Family</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("BenMedPartFamily") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenMedPartFamily" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);" ></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("BenMedPartFamilyAmount") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenMedPartFamilyAmount" onchange="checkForWholeNumber(this.value)" ID="Text5"><br></td>
						<td class="formMain" valign="top" align="left">Total monthly premium<br>(agency + EE contribution)</td>
					</tr>
					</table>
				</td>						
			
			</tr>
			
			<tr>
				<td class="formMain" align="center" cellpadding="1">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" ID="Table10">
					<tr>
						<td class="formMain" valign="top" align="center">Dental&nbsp;
							<a href="../helpfiles/surveyhelp.asp?HelpID=BenDentOffered" onclick="NewWindow(this.href,'name','550','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="Center">
						<select name="frmBenefitsBenDentOffered" size=1 class="formMain" ID="Select2" onchange="CheckForOffer(this.form)">
							<%If say = "edit" Then%>
									<%If GetBenefits("BenDentOffered") then%>
										<option value="true">Yes</option>
										<option value="false">No</option>
									<%else%>
										<option value="false">No</option>
										<option value="true">Yes</option>
									<% End If %>
								<%Else%>
									<option value="selected">Offered?</option>
									<option value="true">Yes</option>
									<option value="false">No</option>
								<%End If%>
						</select>&nbsp;
						</td>
					</tr>
					</table>
				</td>

				<td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table5">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2">For Employee</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("BenDentFullEmployee") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenDentFullEmployee" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);" ID="Text6"></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("BenDentFullEmployeeAmount") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenDentFullEmployeeAmount" onchange="checkForWholeNumber(this.value)" ID="Text7"><br></td>
						<td class="formMain" valign="top" align="left">Total monthly premium<br>(agency + EE contribution)</td>
					</tr>
					</table>
				</td>
				
				<td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table6">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2">For Family</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("BenDentFullFamily") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenDentFullFamily" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);" ID="Text8"></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("BenDentFullFamilyAmount") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenDentFullFamilyAmount" onchange="checkForWholeNumber(this.value)" ID="Text9"><br></td>
						<td class="formMain" valign="top" align="left">Total monthly premium<br>(agency + EE contribution)</td>
					</tr>
					</table>
				</td>
				
				<td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table7">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2">For Employee</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("BenDentPartEmployee") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenDentPartEmployee" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);" ID="Text10"></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("BenDentPartEmployeeAmount") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenDentPartEmployeeAmount" onchange="checkForWholeNumber(this.value)" ID="Text11"><br></td>
						<td class="formMain" valign="top" align="left">Total monthly premium<br>(agency + EE contribution)</td>
					</tr>
					</table>
				</td>		
				
				<td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table8">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2">For Family</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("BenDentPartFamily") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenDentPartFamily" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);" ID="Text12"></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("BenDentPartFamilyAmount") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenDentPartFamilyAmount" onchange="checkForWholeNumber(this.value)" ID="Text13"><br></td>
						<td class="formMain" valign="top" align="left">Total monthly premium<br>(agency + EE contribution)</td>
					</tr>
					</table>
				</td>						
			
			</tr>
			
			<tr>
				<td class="formMain" align="center" cellpadding="1">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" ID="Table11">
					<tr>
						<td class="formMain" valign="top" align="center">Vision&nbsp;
							<a href="../helpfiles/surveyhelp.asp?HelpID=BenVisOffered" onclick="NewWindow(this.href,'name','550','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="Center">
						<select name="frmBenefitsBenVisOffered" size=1 class="formMain" ID="Select3" onchange="CheckForOffer(this.form)">
							<%If say = "edit" Then%>
									<%If GetBenefits("BenVisOffered") then%>
										<option value="true">Yes</option>
										<option value="false">No</option>
									<%else%>
										<option value="false">No</option>
										<option value="true">Yes</option>
									<% End If %>
								<%Else%>
									<option value="selected">Offered?</option>
									<option value="true">Yes</option>
									<option value="false">No</option>
								<%End If%>
						</select>&nbsp;
						</td>
					</tr>
					</table>
				</td>

				<td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table12">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2">For Employee</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("BenVisFullEmployee") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenVisFullEmployee" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);" ID="Text14"></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("BenVisFullEmployeeAmount") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenVisFullEmployeeAmount" onchange="checkForWholeNumber(this.value)" ID="Text15"><br></td>
						<td class="formMain" valign="top" align="left">Total monthly premium<br>(agency + EE contribution)</td>
					</tr>
					</table>
				</td>
				
				<td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table13">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2">For Family</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("BenVisFullFamily") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenVisFullFamily" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);" ID="Text16"></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("BenVisFullFamilyAmount") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenVisFullFamilyAmount" onchange="checkForWholeNumber(this.value)" ID="Text17"><br></td>
						<td class="formMain" valign="top" align="left">Total monthly premium<br>(agency + EE contribution)</td>
					</tr>
					</table>
				</td>
				
				<td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table14">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2">For Employee</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("BenVisPartEmployee") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenVisPartEmployee" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);" ID="Text18"></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("BenVisPartEmployeeAmount") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenVisPartEmployeeAmount" onchange="checkForWholeNumber(this.value)" ID="Text19"><br></td>
						<td class="formMain" valign="top" align="left">Total monthly premium<br>(agency + EE contribution)</td>
					</tr>
					</table>
				</td>		
				
				<td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table15">
					<tr>
						<td class="formMain" valign="top" align="center" colspan="2">For Family</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("BenVisPartFamily") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenVisPartFamily" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);" ID="Text20"></td>
						<td class="formMain" valign="top" align="left">% of premium <br> paid by agency</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="top" align="left"><input type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("BenVisPartFamilyAmount") %><% Else %>0<% End If %>" class="formMain" name="frmBenefitsBenVisPartFamilyAmount" onchange="checkForWholeNumber(this.value)" ID="Text21"><br></td>
						<td class="formMain" valign="top" align="left">Total monthly premium<br>(agency + EE contribution)</td>
					</tr>
					</table>
				</td>						
			
			</tr>
			</table>
			<table width="100%" border="1" bordercolordark="#003063" cellspacing="0" cellpadding="1" ID="Table17">
			<tr>
				<td colspan="6" class="formHeaderSmall">BENEFITS - NON-MEDICAL<br>(check all that apply)</td>
			</tr>	
			
			<tr>
				<!--<td class="formMain" colspan="3">&nbsp;</td>-->
				<td class="formMain" width="40%">&nbsp;</td>
				<td class="formMain" width="30%">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table16">
					<tr>
						
						<td class="formMain" valign="top" align="center" colspan="4">Full Time</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="bottom" align="center">Offered</td>
						<td class="formMain" valign="bottom" align="left">Paid</td>
						<td class="formMain" valign="bottom" align="left">%<br>of Prem<br>Pd by Ag</td>
						<td class="formMain" valign="bottom" align="left">Total<br>Mthly<br>Prem</td>
					</tr>
					</table>
				</td>
				
				<!--<td class="formMain" colspan="3">&nbsp;</td>
				<td class="formMain" align="center">Full Time</td>-->
				<td class="formMain" align="center" width="30%">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table18">
					<tr>
						
						<td class="formMain" valign="top" align="center" colspan="4">Part Time</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td class="formMain" valign="bottom" align="center">Offered</td>
						<td class="formMain" valign="bottom" align="left">Paid</td>
						<td class="formMain" valign="bottom" align="left">%<br>of Prem<br>Pd by Ag</td>
						<td class="formMain" valign="bottom" align="left">Total<br>Mthly<br>Prem</td>
					</tr>
					</table>
				</td>

			</tr>			
			
			<tr>
				<td class="formMain" width="40%">
					<a href="../helpfiles/surveyhelp.asp?HelpID=DisInsShortTermFull" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Disability Insurance SHORT Term</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table20">
					<tr>
					

						<td class=formMain width="25%" align=center><input id="frmBenefitsDisInsShortTermFull" type="checkbox" class="formMain" name="frmBenefitsDisInsShortTermFull" value="1"  <% If say = "edit" Then %><% if GetBenefits("DisInsShortTermFull")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain width="25%" align=center><input id="frmBenefitsDisInsShortTermFullPaid" type="checkbox" class="formMain" name="frmBenefitsDisInsShortTermFullPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("DisInsShortTermFullPaid")=true then%>checked="true"<% end if %><% end if %> onclick="ZeroIfNotChecked(this.form)"</td>
						<td class=formMain width="25%" align=center><input id="frmBenefitsDisInsShortTermFullPrcnt" type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("DisInsShortTermFullPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsDisInsShortTermFullPrcnt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">%</td>
						<td class=formMain width="25%" align=right>$<input id="frmBenefitsDisInsShortTermFullAmount" type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("DisInsShortTermFullAmount")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsDisInsShortTermFullAmount" onchange="checkForWholeNumber(this.value)"></td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table21">
					<tr>
						<td class=formMain width="25%" align="center"><input id="frmBenefitsDisInsShortTermPart" type="checkbox" class="formMain" name="frmBenefitsDisInsShortTermPart" value="1"  <% If say = "edit" Then %><% if GetBenefits("DisInsShortTermPart")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain width="25%" align="center"><input id="frmBenefitsDisInsShortTermPartPaid" type="checkbox" class="formMain" name="frmBenefitsDisInsShortTermPartPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("DisInsShortTermPartPaid")=true then%>checked="true"<% end if %><% end if %> onclick="ZeroIfNotChecked(this.form)"</td>
						<td class=formMain width="25%" align="center"><input id="frmBenefitsDisInsShortTermPartPrcnt" type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("DisInsShortTermPartPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsDisInsShortTermPartPrcnt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">%</td>
						<td class=formMain width="25%" align="center">$<input id="frmBenefitsDisInsShortTermPartAmount" type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("DisInsShortTermPartAmount")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsDisInsShortTermPartAmount" onchange="checkForWholeNumber(this.value)"></td>
					</tr>
					</table>
				</td>
			</tr>
			
			<tr>
				<td class="formMain" width="40%">
					<a href="../helpfiles/surveyhelp.asp?HelpID=DisInsLongTermFull" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Disability Insurance LONG Term</td>				
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table22">
					<tr>
						<td class=formMain width="25%" align=center><input id="frmBenefitsDisInsLongTermFull" type="checkbox" class="formMain" name="frmBenefitsDisInsLongTermFull" value="1"  <% If say = "edit" Then %><% if GetBenefits("DisInsLongTermFull")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain width="25%" align=center><input id="frmBenefitsDisInsLongTermFullPaid" type="checkbox" class="formMain" name="frmBenefitsDisInsLongTermFullPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("DisInsLongTermFullPaid")=true then%>checked="true"<% end if %><% end if %> onclick="ZeroIfNotChecked(this.form)"</td>
						<td class=formMain width="25%" align=center><input id="frmBenefitsDisInsLongTermFullPrcnt" type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("DisInsLongTermFullPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsDisInsLongTermFullPrcnt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">%</td>
						<td class=formMain width="25%" align=center>$<input id="frmBenefitsDisInsLongTermFullAmount" type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("DisInsLongTermFullAmount")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsDisInsLongTermFullAmount" onchange="checkForWholeNumber(this.value)"></td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table23">
					<tr>
						<td class=formMain width="25%" align=center><input id="frmBenefitsDisInsLongTermPart" type="checkbox" class="formMain" name="frmBenefitsDisInsLongTermPart" value="1"  <% If say = "edit" Then %><% if GetBenefits("DisInsLongTermPart")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain width="25%" align=center><input id="frmBenefitsDisInsLongTermPartPaid" type="checkbox" class="formMain" name="frmBenefitsDisInsLongTermPartPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("DisInsLongTermPartPaid")=true then%>checked="true"<% end if %><% end if %> onclick="ZeroIfNotChecked(this.form)"</td>
						<td class=formMain width="25%" align=center><input id="frmBenefitsDisInsLongTermPartPrcnt" type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("DisInsLongTermPartPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsDisInsLongTermPartPrcnt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">%</td>
						<td class=formMain width="25%" align=center>$<input id="frmBenefitsDisInsLongTermPartAmount" type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("DisInsLongTermPartAmount")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsDisInsLongTermPartAmount" onchange="checkForWholeNumber(this.value)"></td>
					</tr>
					</table>
				</td>
			</tr>
			
			<tr>
				<td class="formMain" width="40%">
					<a href="../helpfiles/surveyhelp.asp?HelpID=EAPFull" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;EAP: Employee Assistance Programs</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table24">
					<tr>
						<td class=formMain width="25%" align=center><input id="frmBenefitsEAPFull" type="checkbox" class="formMain" name="frmBenefitsEAPFull" value="1"  <% If say = "edit" Then %><% if GetBenefits("EAPFull")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>				
						<td class=formMain width="25%" align=center><input id="frmBenefitsEAPFullPaid" type="checkbox" class="formMain" name="frmBenefitsEAPFullPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("EAPFullPaid")=true then%>checked="true"<% end if %><% end if %> onclick="ZeroIfNotChecked(this.form)"</td>
						<td class=formMain width="25%" align=center><input id="frmBenefitsEAPFullPrcnt" type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("EAPFullPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsEAPFullPrcnt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">%</td>
						<td class=formMain width="25%" align=center>$<input id="frmBenefitsEAPFullAmount" type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("EAPFullAmount")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsEAPFullAmount" onchange="checkForWholeNumber(this.value)"></td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table25">
					<tr>
						<td class=formMain width="25%" align=center><input id="frmBenefitsEAPPart" type="checkbox" class="formMain" name="frmBenefitsEAPPart" value="1"  <% If say = "edit" Then %><% if GetBenefits("EAPPart")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>				
						<td class=formMain width="25%" align=center><input id="frmBenefitsEAPPartPaid" type="checkbox" class="formMain" name="frmBenefitsEAPPartPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("EAPPart")=true then%>checked="true"<% end if %><% end if %> onclick="ZeroIfNotChecked(this.form)"</td>
						<td class=formMain width="25%" align=center><input id="frmBenefitsEAPPartPrcnt" type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("EAPPartPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsEAPPartPrcnt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">%</td>
						<td class=formMain width="25%" align=center>$<input id="frmBenefitsEAPPartAmount" type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("EAPPartAmount")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsEAPPartAmount" onchange="checkForWholeNumber(this.value)"></td>
					</tr>
					</table>
				</td>
			</tr>
			
			<!--Commented due to change in requirements (not need to collect this)
			<tr>
				<td class="formMain" width="40%">
					<a href="../helpfiles/surveyhelp.asp?HelpID=FlexFull" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;"Flex" Pre-Tax Plan (medical, dependent)</td>			
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table26">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsFlexFull" value="1"  <% If say = "edit" Then %><% if GetBenefits("FlexFull")=true then%>checked="true"<% end if %><% end if %></td>				
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsFlexFullPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("FlexFullPaid")=true then%>checked="true"<% end if %><% end if %></td>
						<td class=formMain width="25%" align=center><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("FlexFullPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsFlexFullPrcnt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">%</td>
						<td class=formMain width="25%" align=center>$<input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("FlexFullAmount")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsFlexFullAmount" onchange="checkForWholeNumber(this.value)"></td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table27">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsFlexPart" value="1"  <% If say = "edit" Then %><% if GetBenefits("FlexPart")=true then%>checked="true"<% end if %><% end if %></td>				
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsFlexPartPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("FlexPartPaid")=true then%>checked="true"<% end if %><% end if %></td>
						<td class=formMain width="25%" align=center><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("FlexPartPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsFlexPartPrcnt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">%</td>
						<td class=formMain width="25%" align=center>$<input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("FlexPartAmount")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsFlexPartAmount" onchange="checkForWholeNumber(this.value)"></td>
					</tr>
					</table>
				</td>
			</tr>-->
			
			<tr>
				<td class="formMain" width="40%">
					<a href="../helpfiles/surveyhelp.asp?HelpID=HealthClubFull" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Health Club</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table28">
					<tr>
						<td class=formMain width="25%" align=center><input id="frmBenefitsHealthClubFull" type="checkbox" class="formMain" name="frmBenefitsHealthClubFull" value="1"  <% If say = "edit" Then %><% if GetBenefits("HealthClubFull")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain width="25%" align=center><input id="frmBenefitsHealthClubFullPaid" type="checkbox" class="formMain" name="frmBenefitsHealthClubFullPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("HealthClubFullPaid")=true then%>checked="true"<% end if %><% end if %> onclick="ZeroIfNotChecked(this.form)"</td>
						<td class=formMain width="25%" align=center><input id="frmBenefitsHealthClubFullPrcnt" type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("HealthClubFullPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsHealthClubFullPrcnt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">%</td>
						<td class=formMain width="25%" align=center>$<input id="frmBenefitsHealthClubFullAmount" type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("HealthClubFullAmount")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsHealthClubFullAmount" onchange="checkForWholeNumber(this.value)"></td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table29">
					<tr>
						<td class=formMain width="25%" align=center><input id="frmBenefitsHealthClubPart" type="checkbox" class="formMain" name="frmBenefitsHealthClubPart" value="1"  <% If say = "edit" Then %><% if GetBenefits("HealthClubPart")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain width="25%" align=center><input id="frmBenefitsHealthClubPartPaid" type="checkbox" class="formMain" name="frmBenefitsHealthClubPartPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("HealthClubPartPaid")=true then%>checked="true"<% end if %><% end if %> onclick="ZeroIfNotChecked(this.form)"</td>
						<td class=formMain width="25%" align=center><input id="frmBenefitsHealthClubPartPrcnt" type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("HealthClubPartPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsHealthClubPartPrcnt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">%</td>
						<td class=formMain width="25%" align=center>$<input id="frmBenefitsHealthClubPartAmount" type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("HealthClubPartAmount")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsHealthClubPartAmount" onchange="checkForWholeNumber(this.value)"></td>
					</tr>
					</table>
				</td>
			</tr>		
			
			<tr>
				<td class="formMain" width="40%">
					<a href="../helpfiles/surveyhelp.asp?HelpID=LifeInsuranceFull" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Life Insurance</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table30">
					<tr>
						<td class=formMain width="25%" align=center><input id="frmBenefitsLifeInsuranceFull" type="checkbox" class="formMain" name="frmBenefitsLifeInsuranceFull" value="1"  <% If say = "edit" Then %><% if GetBenefits("LifeInsuranceFull")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain width="25%" align=center><input id="frmBenefitsLifeInsuranceFullPaid" type="checkbox" class="formMain" name="frmBenefitsLifeInsuranceFullPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("LifeInsuranceFullPaid")=true then%>checked="true"<% end if %><% end if %> onclick="ZeroIfNotChecked(this.form)"</td>
						<td class=formMain width="25%" align=center><input id="frmBenefitsLifeInsuranceFullPrcnt" type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("LifeInsuranceFullPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsLifeInsuranceFullPrcnt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">%</td>
						<td class=formMain width="25%" align=center>$<input id="frmBenefitsLifeInsuranceFullAmount" type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("LifeInsuranceFullAmount")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsLifeInsuranceFullAmount" onchange="checkForWholeNumber(this.value)"></td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table31">
					<tr>
						<td class=formMain width="25%" align=center><input id="frmBenefitsLifeInsurancePart" type="checkbox" class="formMain" name="frmBenefitsLifeInsurancePart" value="1"  <% If say = "edit" Then %><% if GetBenefits("LifeInsurancePart")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain width="25%" align=center><input id="frmBenefitsLifeInsurancePartPaid" type="checkbox" class="formMain" name="frmBenefitsLifeInsurancePartPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("LifeInsurancePartPaid")=true then%>checked="true"<% end if %><% end if %> onclick="ZeroIfNotChecked(this.form)"</td>
						<td class=formMain width="25%" align=center><input id="frmBenefitsLifeInsurancePartPrcnt" type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("LifeInsurancePartPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsLifeInsurancePartPrcnt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);">%</td>
						<td class=formMain width="25%" align=center>$<input id="frmBenefitsLifeInsurancePartAmount" type="text" size="1" maxlength="4" value="<% If say = "edit" Then %><%= GetBenefits("LifeInsurancePartAmount")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsLifeInsurancePartAmount" onchange="checkForWholeNumber(this.value)"></td>
					</tr>
					</table>
				</td>
			</tr>								
			</table>
			
			
			
			<table width="100%" border="1" bordercolordark="#003063" cellspacing="0" cellpadding="1" ID="Table19">
			<tr>
				<td class="formMain" width="40%">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table26">
					<tr>
						<td class=formMain width="50%" rowspan=2"><a href="../helpfiles/surveyhelp.asp?HelpID=TimeOffFull" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Paid Time Off (Vacation, Floating Holidays, Personal Days)</td>
						<td class=formMain width="50%" align=left valign="bottom" height="80">&nbsp;&nbsp;Exempt (Salaried)</td>
					</tr>
					<tr>
						<td class=formMain width="50%" align=left valign="bottom" height="20">&nbsp;&nbsp;Non-Exempt (Hourly)</td>
					</tr>
					</table>
				<td class="formMain" align="center" width="30%">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table32">
					<tr>
						<td class=formMain width="25%" align=center valign=bottom><input type="checkbox" class="formMain" name="frmBenefitsTimeOffFull" value="1"  <% If say = "edit" Then %><% if GetBenefits("TimeOffFull")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain width="25%" align=center>&nbsp;&nbsp;# of days for new employee<br><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TimeOffFullDays")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffFullDays" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"></td>
						<td class=formMain width="25%" align=center># of years before increase<br><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TimeOffFullYears")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffFullYears" onchange="equalsLessThan101(this.value)"></td>
						<td class=formMain width="25%" align=center># of days after increase<br><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TimeOffFullDaysIncreased")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffFullDaysIncreased" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"></td>
					</tr>
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsTimeOffFullNExempt" value="1"  <% If say = "edit" Then %><% if GetBenefits("TimeOffFullNExempt")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain width="25%" align=center><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TimeOffFullDaysNExempt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffFullDaysNExempt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"></td>
						<td class=formMain width="25%" align=center><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TimeOffFullYearsNExempt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffFullYearsNExempt" onchange="equalsLessThan101(this.value)"></td>
						<td class=formMain width="25%" align=center><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TimeOffFullDaysIncreasedNExempt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffFullDaysIncreasedNExempt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"></td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center" width="30%">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table27">
					<tr>
						<td class=formMain width="25%" align=center valign=bottom><input type="checkbox" class="formMain" name="frmBenefitsTimeOffPart" value="1"  <% If say = "edit" Then %><% if GetBenefits("TimeOffPart")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain width="25%" align=center>&nbsp;&nbsp;# of days for new employee<br><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TimeOffPartDays")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffPartDays" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"></td>
						<td class=formMain width="25%" align=center># of years before increase<br><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TimeOffPartYears")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffPartYears" onchange="equalsLessThan101(this.value)"></td>
						<td class=formMain width="25%" align=center># of days after increase<br><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TimeOffPartDaysIncreased")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffPartDaysIncreased" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"></td>
					</tr>
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsTimeOffPartNExempt" value="1"  <% If say = "edit" Then %><% if GetBenefits("TimeOffPartNExempt")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain width="25%" align=center><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TimeOffPartDaysNExempt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffPartDaysNExempt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"></td>
						<td class=formMain width="25%" align=center><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TimeOffPartYearsNExempt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffPartYearsNExempt" onchange="equalsLessThan101(this.value)"></td>
						<td class=formMain width="25%" align=center><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TimeOffPartDaysIncreasedNExempt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffPartDaysIncreasedNExempt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"></td>
					</tr>
					</table>
				</td>		
				<!--<td class="formMain" align="center" width="30%">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table33">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsTimeOffPart" value="1"  <% If say = "edit" Then %><% if GetBenefits("TimeOffPart")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain width="75%" align=center colspan="3"><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TimeOffPartDays")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffPartDays" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"> days</td>
					</tr>
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsTimeOffPartNExempt" value="1"  <% If say = "edit" Then %><% if GetBenefits("TimeOffPartNExempt")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain width="75%" align=center colspan="3"><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TimeOffPartDaysNExempt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffPartDaysNExempt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"> days</td>
					</tr>
					</table>
				</td>-->
			</tr>

			
			<tr>
				<td class="formMain" width="40%">
					<a href="../helpfiles/surveyhelp.asp?HelpID=TimeOffSickFull" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Paid Time Off (Sick Time)</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table34">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsTimeOffSickFull" value="1"  <% If say = "edit" Then %><% if GetBenefits("TimeOffSickFull")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain width="75%" align=center colspan="3"><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TimeOffSickFullDays")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffSickFullDays" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"> days</td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table35">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsTimeOffSickPart" value="1"  <% If say = "edit" Then %><% if GetBenefits("TimeOffSickPart")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain width="75%" align=center colspan="3"><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TimeOffSickPartDays")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffSickPartDays" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"> days</td>
					</tr>
					</table>
				</td>
			</tr>													

<!-- Commented due changes from edit by Cindy and Jeff on Sept 6
			<tr>
				<td class="formMain" width="40%">
					<a href="../helpfiles/surveyhelp.asp?HelpID=TimeOffVacFull" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Paid Time Off (Vacation)</td>			
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table36">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsTimeOffVacFull" value="1"  <% If say = "edit" Then %><% if GetBenefits("TimeOffVacFull")=true then%>checked="true"<% end if %><% end if %></td>
						<td class=formMain width="75%" align=center colspan="3"><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TimeOffVacFullDays")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffVacFullDays" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"> days</td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table37">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsTimeOffVacPart" value="1"  <% If say = "edit" Then %><% if GetBenefits("TimeOffVacPart")=true then%>checked="true"<% end if %><% end if %></td>
						<td class=formMain width="75%" align=center colspan="3"><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TimeOffVacPartDays")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTimeOffVacPartDays" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"> days</td>
					</tr>
					</table>
				</td>
			</tr>	
-->

			<tr>
				<td class="formMain" width="40%">
					<a href="../helpfiles/surveyhelp.asp?HelpID=ProfDuesFull" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Professional Dues, Conferences, etc.</td>			
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table38">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsProfDuesFull" value="1"  <% If say = "edit" Then %><% if GetBenefits("ProfDuesFull")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsProfDuesFullPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("ProfDuesFullPaid")=true then%>checked="true"<% end if %><% end if %> onclick="ZeroIfNotChecked(this.form)"</td>
						<td class=formMain width="50%" align=center colspan="2">avg $ amount per EE<br><input type="text" size="2" maxlength="5" value="<% If say = "edit" Then %><%= GetBenefits("ProfDuesFullAmount")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsProfDuesFullAmount" onchange="checkForWholeNumber(this.value)"></td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table39">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsProfDuesPart" value="1"  <% If say = "edit" Then %><% if GetBenefits("ProfDuesPart")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsProfDuesPartPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("ProfDuesPartPaid")=true then%>checked="true"<% end if %><% end if %> onclick="ZeroIfNotChecked(this.form)"</td>
						<td class=formMain width="50%" align=center colspan="2">avg $ amount per EE<br><input type="text" size="2" maxlength="5" value="<% If say = "edit" Then %><%= GetBenefits("ProfDuesPartAmount")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsProfDuesPartAmount" onchange="checkForWholeNumber(this.value)"></td>
					</tr>
					</table>
				</td>
			</tr>			
			
			<tr>
				<td class="formMain" width="40%">
					<a href="../helpfiles/surveyhelp.asp?HelpID=RetirementFull" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Pension Plan (Defined benefit plan)</td>			
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table40">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsRetirementFull" value="1"  <% If say = "edit" Then %><% if GetBenefits("RetirementFull")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsRetirementFullPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("RetirementFullPaid")=true then%>checked="true"<% end if %><% end if %></td>-->
						<td class=formMain width="75%" align=center colspan="3">% of contribution by agency<input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("RetirementFullPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsRetirementFullPrcnt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"></td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table41">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsRetirementPart" value="1"  <% If say = "edit" Then %><% if GetBenefits("RetirementPart")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsRetirementPartPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("RetirementpartPaid")=true then%>checked="true"<% end if %><% end if %></td>-->
						<td class=formMain width="75%" align=center colspan="3">% of contribution by agency<input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("RetirementPartPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsRetirementPartPrcnt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"></td>
					</tr>
					</table>
				</td>
			</tr>
			
			<tr>
				<td class="formMain" width="40%">
					<a href="../helpfiles/surveyhelp.asp?HelpID=403BFull" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;403 B</td>			
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table42">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefits403BFull" value="1"  <% If say = "edit" Then %><% if GetBenefits("403BFull")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain align=center>Employer Contribution<!--<input type="text" size="1" maxlength="3" value=" <% If say = "edit" Then %><%= GetBenefits("403BFullContrib")%><% Else %>0<% End If %>" class="formMain" name="frmBenefits403BFullContrib">-->
							<select name="frmBenefits403BFullContrib" size=1 class="formMain" ID="Select4" onchange="ZeroIfNotChecked(this.form)">
								<!--<option value="selected">Y/N</option>-->
								<%If say = "edit" Then%>
									<%If GetBenefits("403BFullContrib") then%>
										<option value="true">Yes</option>
										<option value="false">No</option>
									<%else%>
										<option value="false">No</option>
										<option value="true">Yes</option>
									<% End If %>
								<%Else%>
									<option value="selected">Y/N</option>
									<option value="true">Yes</option>
									<option value="false">No</option>
								<%End If%>
								
							</select>&nbsp;
						</td>
						<td class=formMain align=center>% of matching<br>contribution<br><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("403BFullContribPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefits403BFullContribPrcnt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"></td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table43">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefits403BPart" value="1"  <% If say = "edit" Then %><% if GetBenefits("403BPart")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<td class=formMain align=center>Employer<br>Contribution<!--<input type="text" size="1" maxlength="3" value=" <% If say = "edit" Then %><%= GetBenefits("403BPartContrib")%><% Else %>0<% End If %>" class="formMain" name="frmBenefits403BPartContrib">-->
							<select name="frmBenefits403BPartContrib" size=1 class="formMain" ID="Select5" onchange="ZeroIfNotChecked(this.form)">
								<%If say = "edit" Then%>
									<%If GetBenefits("403BPartContrib") then%>
										<option value="true">Yes</option>
										<option value="false">No</option>
									<%else%>
										<option value="false">No</option>
										<option value="true">Yes</option>
									<% End If %>
								<%Else%>
									<option value="selected">Y/N</option>
									<option value="true">Yes</option>
									<option value="false">No</option>
								<%End If%>
							</select>&nbsp;
						</td>
						<td class=formMain align=center>% of matching<br>contribution<br><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("403BPartContribPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefits403BPartContribPrcnt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"></td>
					</tr>
					</table>
				</td>
			</tr>						
			
			<tr>
				<td class="formMain" width="40%">
					<a href="../helpfiles/surveyhelp.asp?HelpID=TelecommFull" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Telecommuting</td>			
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table44">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsTelecommFull" value="1"  <% If say = "edit" Then %><% if GetBenefits("TelecommFull")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<!--<td class=formMain width="20%" align=center><input type="checkbox" class="formMain" name="frmBenefitsTelecommFullPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("TelecommFullPaid")=true then%>checked="true"<% end if %><% end if %></td>-->
						<td class=formMain align=center># of EEs<br><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TelecommFullCount")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTelecommFullCount" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"></td>
						<td class=formMain align=center>% of EE<br>population<br><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TelecommFullPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTelecommFullPrcnt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"></td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table45">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsTelecommPart" value="1"  <% If say = "edit" Then %><% if GetBenefits("TelecommPart")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<!--<td class=formMain width="20%" align=center><input type="checkbox" class="formMain" name="frmBenefitsTelecommPartPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("TelecommPartPaid")=true then%>checked="true"<% end if %><% end if %></td>-->
						<td class=formMain align=center># of EEs<br><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TelecommPartCount")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTelecommPartCount" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"></td>
						<td class=formMain align=center>% of EE<br>population<br><input type="text" size="1" maxlength="3" value="<% If say = "edit" Then %><%= GetBenefits("TelecommPartPrcnt")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTelecommPartPrcnt" onchange="checkForWholeNumber(this.value); equalsLessThan101(this.value);"></td>
					</tr>
					</table>
				</td>
			</tr>
			
			<tr>
				<td class="formMain" width="40%">
					<a href="../helpfiles/surveyhelp.asp?HelpID=TuitionFull" onclick="NewWindow(this.href,'name','600','450','yes');return false;"><img src="../images/qmarksmall.gif" alt="" width="15" height="16" border="0"></a>&nbsp;Tuition Reimbursement</td>			
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table46">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsTuitionFull" value="1"  <% If say = "edit" Then %><% if GetBenefits("TuitionFull")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsTuitionFullPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("TuitionFullPaid")=true then%>checked="true"<% end if %><% end if %></td>-->
						<td class=formMain width="75%" align=center colspan="3">Maximum $ paid<br><input type="text" size="2" maxlength="5" value="<% If say = "edit" Then %><%= GetBenefits("TuitionFullAmount")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTuitionFullAmount" onchange="checkForWholeNumber(this.value)"></td>
					</tr>
					</table>
				</td>
				<td class="formMain" align="center">
					<table width="100%" border="0" cellspacing="0" cellpadding="1" ID="Table47">
					<tr>
						<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsTuitionPart" value="1"  <% If say = "edit" Then %><% if GetBenefits("TuitionPart")=true then%>checked="true"<% end if %><% end if %> onclick="CheckForOffer(this.form)"</td>
						<!--<td class=formMain width="25%" align=center><input type="checkbox" class="formMain" name="frmBenefitsTuitionPartPaid" value="1"  <% If say = "edit" Then %><% if GetBenefits("TuitionPartPaid")=true then%>checked="true"<% end if %><% end if %></td>-->
						<td class=formMain width="75%" align=center colspan="3">Maximum $ paid<br><input type="text" size="2" maxlength="5" value="<% If say = "edit" Then %><%= GetBenefits("TuitionPartAmount")%><% Else %>0<% End If %>" class="formMain" name="frmBenefitsTuitionPartAmount" onchange="checkForWholeNumber(this.value)"></td>
					</tr>
					</table>
				</td>
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
	GetBenefits.Close
	Set GetBenefits = Nothing
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
