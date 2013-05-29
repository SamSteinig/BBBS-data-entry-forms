<table border=0>
<tr><td align="center"><form name="frmGeneralInformation" action="GeneralInformation_edit.asp?y=<%= Request("y") %>" method="post">
<!--#include file="../includes/form_stamp.asp"-->
<input type="hidden" name="status" value="editOld">
	<P>
			<table width="600" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063">
				<tr> 
					<td colspan="3" align="center" valign="top" class="formSubhead">BBBS - <%= y %> Annual Agency Information (AAI)</td>
			</tr>
				
<% if printform = "No" then %>				
				
				<tr>
					<td colspan="3" class="formHeader">GENERAL INFORMATION</td>
				</tr>
				
<% else %>

				<tr>
					<td colspan="3" class="formIndex">GENERAL INFORMATION</td>
				</tr>
				
<% end if %>

<% Dim ReadOnlyLevel
If Session("ReadOnly") then
	ReadOnlyLevel=1
Else
	ReadOnlyLevel=0
End If
%>
				
				
				
				<tr>
					<td colspan="3" class="formMainBold">Created: <%= GetGeneralInformation("CreateDate") %><br>
		<% form = "GeneralInformation" %> 
		<% gid = GetGeneralInformation("GeneralInformationID") %>
		<!--#include file="../includes/lastmodified_stamp.asp"-->
					</td>
				</tr>
				
<!-- Question Number 1 -->
				<tr>
					<td align="right" valign="top" class="formMain">1.</td>
					<td align="left" valign="top" class="formMain">Population of your Service Community Area (SCA):</td>
					<td align="right" valign="top" class="formMainRightJ"><%= GetGeneralInformation("PopulationSCA") %></td>
				</tr>
				
<!-- Question Number 2 -->
				<tr>
					<td align="right" valign="top" class="formMain">2.</td>
					<td align="left" valign="top" class="formMain">Number of school age children (K-12) in SCA:</td>
					<td align="right" valign="top" class="formMainRightJ"><%= GetGeneralInformation("SchoolAgeSCA") %></td>
				</tr>
<!-- Question Number 3 -->
				<tr> 
					<td align="right" valign="top" class="formMain">3.</td>
					<td align="left" valign="top" class="formMain">What is the total number of volunteer inquiries you received?<br><em>A volunteer is considered to have inquired when he/she contacts the agency, expresses an interest in being a Big and provides basic contact information.  Contact includes web-based inquiries.</em></td>
					<td align="right" valign="top" class="formMainRightJ"><%= GetGeneralInformation("VolunteerInquiries") %></td>
				</tr>
<!-- Question Number 4 -->
				<tr>
					<td align="right" valign="top" class="formMain">4.</td>
					<td align="left" valign="top" class="formMain">What is the total number of volunteer in person interviews?</td>
					<td align="right" valign="top" class="formMainRightJ"><%=GetGeneralInformation("VolunteerInPersonInterviews")%></td>
				</tr>

<!-- Question Number 5 -->
				<tr>
					<td align="right" valign="top" class="formMain">5.</td>
					<td align="left" valign="top" class="formMain">What is the total number of volunteers that were matched?</td>
					<td align="right" valign="top" class="formMainRightJ"><%= GetGeneralInformation("TotalVolunteersMatched") %></td>
				</tr>				

				
<!-- Question Number 6 NOT USED
				<tr>
					<td align="right" valign="top" class="formMain">6.</td>
					<td align="left" valign="top" class="formMain">What is the total number of volunteers that were accepted?</td>
					<td align="right" valign="top" class="formMainRightJ"><%' = GetGeneralInformation("VolunteersAccepted") %></td>
				</tr>
-->

<!-- Question Number 6 -->
				<tr>
					<td align="right" valign="top" class="formMain">6.</td>
					<td align="left" valign="top" class="formMain">Do you have a Strategic Growth Plan in place?</td>
					<td align="right" valign="top" class="formMain"><% IF GetGeneralInformation("StrategicGrowthPlan") = True Then %>Yes<% Else %>No<% End If %></td>
				</tr>
<!-- Question Number 7 -->
				<tr>
					<td align="right" valign="top" class="formMain">7.</td>
					<td align="left" valign="top" class="formMain">According to the Strategic Growth Plan, how many children do you plan to serve by 2004?</td>
					<td align="right" valign="top" class="formMainRightJ"><%= GetGeneralInformation("ChildrenBy2004") %></td>
				</tr>
<!-- Question Number 8 -->
				<tr>
					<td align="right" valign="top" class="formMain">8.</td>
					<td align="left" valign="top" class="formMain">Do you use EMPOWER or similar sexual-prevention curriculum?</td>
					<td align="right" valign="top" class="formMain"><% IF GetGeneralInformation("SexualPreventionCurriculum") = True Then %>Yes<% Else %>No<% End If %></td>
				</tr>
<!-- Question Number 9 --> 
				<tr>
					<td align="right" valign="top" class="formMain">9.</td>
					<td align="left" valign="top" class="formMain">Do you provide training for other mentoring organizations?</td>
					<td align="right" valign="top" class="formMain"><% IF GetGeneralInformation("TrainingMentoringOrganizations") = True Then %>Yes<% Else %>No<% End If %></td>
				</tr>
<!-- Question Number 10 -->
				<tr>
					<td align="right" valign="top" class="formMain">10.</td>
					<td align="left" valign="top" class="formMain">Do you provide post-match training for your volunteers?</td>
					<td align="right" valign="top" class="formMain"><% IF GetGeneralInformation("TrainingPostMatch") = True Then %>Yes<% Else %>No<% End If %></td>
				</tr>
				
<!-- Question Number 18 NOT USED
				<tr>
					<td rowspan="2" align="right" valign="top" class="formMain">17.</td>
					<td align="left" valign="top" class="formMain">Do you have an After School Mentoring Program?</td>
					<td align="right" valign="top" class="formMain"><% ' IF GetGeneralInformation("AfterSchoolMentoringProgram") = True Then %>Yes<% ' Else %>No<% ' End If %></td>
				</tr>
				<tr>
					<td align="left" valign="top" class="formMain">If yes, how many children do you serve?</td>
					<td align="right" valign="top" class="formMainRightJ"><%' =  GetGeneralInformation("ASMPHowManyChildren") %></td>
				</tr> -->
				
				<tr>
				<td colspan="3" align="center" class="formMain"><em>Below please list all <strong>RTBM</strong> Clients and Volunteers, Open and Closed.<br>Categories are <strong>mutually exclusive</strong></em></td>
				</tr>
				
<!-- Question Number 11 -->
				<tr>
					<td align="right" valign="top" class="formMain">11.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Unmatched Clients (RTBM)<br>
					OPEN as of 12/31/<%= y %></td>
					<td align="right" valign="top" class="formMainRightJ"><%= GetGeneralInformation("UnmatchedClientsOpen") %></td>
				</tr>
<!-- Question Number 12 -->
				<tr>
					<td align="right" valign="top" class="formMain">12.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Unmatched Clients (RTBM)<br>
					CLOSED between 1/1/<%= y %>-12/31/<%= y %></td>
					<td align="right" valign="top" class="formMainRightJ"><%= GetGeneralInformation("UnmatchedClientsForTheYear") %></td>
				</tr>
<!-- Question Number 13 -->
				<tr>
					<td align="right" valign="top" class="formMain">13.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Unmatched Volunteers<br>
					OPEN as of 12/31/<%= y %></td>
					<td align="right" valign="top" class="formMainRightJ"><%= GetGeneralInformation("UnmatchedVolunteersOpen") %></td>
				</tr>
<!-- Question Number 14 -->
				<tr>
					<td align="right" valign="top" class="formMain">14.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Unmatched Volunteers<br>
					CLOSED between 1/1/<%= y %>-12/31/<%= y %></td>
					<td align="right" valign="top" class="formMainRightJ"><%= GetGeneralInformation("UnmatchedVolunteersForTheYear") %></td>
				</tr>
<!-- Question Number 15 -->
				<tr>
					<td align="right" valign="top" class="formMain">15.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Group Volunteers<br>
					OPEN as of 12/31/<%= y %></td>
					<td align="right" valign="top" class="formMainRightJ"><%= GetGeneralInformation("GroupVolunteersOpen") %></td>
				</tr>
<!-- Question Number 16 -->
				<tr>
					<td align="right" valign="top" class="formMain">16.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Group Volunteers<br>
					CLOSED between 1/1/<%= y %>-12/31/<%= y %></td>
					<td align="right" valign="top" class="formMainRightJ"><%= GetGeneralInformation("GroupVolunteersForTheYear") %></td>
				</tr>

<% if printform = "No" then %>	

	<% if ReadOnlyLevel = 0 then %>					
		<tr>
			<td colspan="3" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold"></td>
		</tr>
		<tr>
	<% end if %>
	
		<td colspan="3" align="center">
		<!--#include file="../includes/contact_info.inc"-->
		</td>
		</tr>
		
<% end if %>
				
				
				</table>  
							
</form>

</td>
</tr>
</table>