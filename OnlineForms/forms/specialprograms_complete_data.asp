	<table border="1" cellspacing="0" cellpadding="3" width="625" bordercolordark="#003063">
	<form name="frmSpecialPrograms" action="SpecialPrograms_edit.asp?y=<%= Request("y") %>" method="post">
	<!--#include file="../includes/form_stamp.asp"-->
	<input type="hidden" name="status" value="editOld">
	<tr>
			<td colspan="4" align="center" valign="top" class="formSubhead">BBBS - <%= y %> Annual Agency Information (AAI)</td>
		</tr>
		
	<% Dim ReadOnlyLevel
	If Session("ReadOnly") then
		ReadOnlyLevel=1
	Else
		ReadOnlyLevel=0
	End If
	%>		
		
	<% if printform = "No" then %>
	
		<tr>
			<td colspan="4" class="formHeader">SPECIAL PROGRAMS</td>
		</tr>
		
	<% else %>
	
		<tr>
			<td colspan="4" class="formIndex">SPECIAL PROGRAMS</td>
		</tr>	
		
	<% end if %>
	
	
		<tr>
			<td colspan="4" class="formMainBold">Created: <%= GetSpecialPrograms("CreateDate") %><br>
		<% form = "SpecialPrograms" %> 
		<% gid = GetSpecialPrograms("SpecialProgramsID") %>
		<!--#include file="../includes/lastmodified_stamp.asp"-->
			 </td>
		</tr>
		<tr>
		<tr>
			<td align="left" valign="top" class="formMain">55 or Older as Bigs (IntergenerationalProgram):</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("FiftyFiveOrOlderBigs") %></td>
			<td align="left" valign="top" class="formMain">Group Activities:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("GroupActivities") %></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">Academic Achievement/Enrichment (Literacy/Tutoring):</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("AcademicAchievement") %></td>
			<td align="left" valign="top" class="formMain">High School Students as Bigs:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("HighSchoolStudentsAsBigs") %></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">After School Mentoring:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("AfterSchoolMentoring") %></td>
			<td align="left" valign="top" class="formMain">Life Skills/Life Choices:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("LifeSkillsLifeChoices") %></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">AIDS Prevention/Intervention:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("AIDSPreventionIntervention") %></td>
			<td align="left" valign="top" class="formMain">Parent Support Groups:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("ParentSupportGroups") %></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">Alcohol Abuse Prevention/Intervention:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("AlcoholAbusePreventionIntervention") %></td>
			<td align="left" valign="top" class="formMain">Partnerships: Civic Organizations:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("PartnershipsCivicOrganizations") %></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">Camping:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("Camping") %></td>
			<td align="left" valign="top" class="formMain">Partnerships: Colleges/Universities:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("PartnershipsCollegesUniversities") %></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">Character Counts:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("CharacterCounts") %></td>
			<td align="left" valign="top" class="formMain">Partnerships: Corporations/Businesses:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("PartnershipsCorporationsBusinesses") %></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">Children with Disabilities:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("ChildrenWithDisabilities") %></td>
			<td align="left" valign="top" class="formMain">Partnerships:Other Youth Serving Organizations:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("PartnershipsOtherYouthServingOrganizations") %></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">College/University Students as Bigs:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("CollegeStudentsAsBigs") %></td>
			<td align="left" valign="top" class="formMain">Partnerships: Religious Organizations:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("PartnershipsReligiousOrganizations") %></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">Community Service Projects:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("CommunityServiceProjects") %></td>
			<td align="left" valign="top" class="formMain">Pregnancy Prevention:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("PregnancyPrevention") %></td>
		</tr>
		<tr>
			<td align="left" valign="top" class="formMain">Drop Out Prevention:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("DropOutPrevention") %></td>
			<td align="left" valign="top" class="formMain">Scholarships:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("Scholarships") %></td>
		</tr>
		<tr>
			<td rowspan="2" class="formMain" align="left" valign="top">Drug Abuse Prevention/Intervention:</td>
			<td rowspan="2" class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("DrugAbusePreventionIntervention") %></td>
			<td rowspan="2" class="formMain" align="left" valign="top">Sexual Abuse Prevention/Intervention</td>
			<td class="formMainRightJ">EMPOWER: <%= GetSpecialPrograms("SexualAbusePreventionInterventionEmpower") %></td>
		</tr>
		<tr>
			<td class="formMainRightJ" align="left" valign="top"><b>NOT</b> EMPOWER: <%= GetSpecialPrograms("SexualAbusePreventionInterventionNOTEmpower") %></td>
		</tr>
		<tr>
			<td rowspan="2" class="formMain" align="left" valign="top">Emergency Financial Assistance:</td>
			<td rowspan="2" class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("EmergencyFinancialAssistance") %></td>
			<td rowspan="2" class="formMain" align="left" valign="top">Site-Based Mentoring<br></td>
			<td class="formMainRightJ" align="left" valign="top">School Based: <%= GetSpecialPrograms("SiteBasedMentoringSchoolBased") %></td>
		</tr>
		<tr>	
			<td class="formMainRightJ" align="left" valign="top"><b>NOT</b> School Based: <%= GetSpecialPrograms("SiteBasedMentoringNOTSchoolBased") %></td>
		</tr>
		<tr>
			<td class="formMain" align="left" valign="top">Employability/Job Readiness (School To Work):</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("EmployabilityJobReadiness") %></td>
			<td class="formMain" align="left" valign="top">Teen Parenting:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("TeenParenting") %></td>
		</tr>
		<tr>
			<td class="formMain" align="left" valign="top">Family Counseling:</td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("FamilyCounseling") %></td>
			<td class="formMain" align="left" valign="top">Other: <%= GetSpecialPrograms("OtherText") %></td>
			<td class="formMainRightJ" align="left" valign="top"><%= GetSpecialPrograms("Other") %></td>
		</tr>
		
		<% if printform = "No" then %>
		
			<% if ReadOnlyLevel = 0 then %>
				<tr>
					<td colspan="4" class="formHeader" align="left" valign="top"><input type="submit" value="Edit Form" class="formMainBold"></td>
				</tr>
			<% end if %>
		<tr>
			<td colspan="4"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
		</tr>
		
		<% end if %>
		
	</table>
