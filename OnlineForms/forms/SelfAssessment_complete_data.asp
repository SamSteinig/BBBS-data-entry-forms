<% If (GetSelfAssessment.EOF) and (GetSelfAssessment.BOF) Then%>
	<span class="formMainBold">No Self Assessment Data Entered for this Agency</span>
<% else %>	

<%  section = Request("section") %>
<table border=0>
<tr><td align="center"><form name="frmSelfAssessment" action="SelfAssessment_edit.asp?y=<%= Request("y") %>&section=<%=Request("section")%>" method="post">
<!--#include file="../includes/form_stamp.asp"-->
<br>

<input type="hidden" name="status" value="editOld">
	<P>
			<table width="650" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063">
				<tr> 
					<td colspan="3" align="center" valign="top" class="formSubhead">BBBS - <%= y %> Agency Self Assessment Form</td>
			</tr>

<% if printform = "No" then %>				
				
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
				
<% else %>

				<tr>
					<td colspan="3" class="formIndex">Agency Self Assessment - Agency ID: <%=AgencyID%>
					<br>
					<% if section="Operational" then %>
						Business Peformance - Operational Standards
					<% else %>
						Program Performance - Program Standards
					<% end if %>
					</td>
				</tr>
				
<% end if %>


<% Dim ReadOnlyLevel
If Session("ReadOnly") then
	ReadOnlyLevel=1
Else
	ReadOnlyLevel=0
End If
%>
				
<!-- Table Header -->				
				
				<tr>
					<td colspan="3" class="formMain">Created: <%= GetSelfAssessment("CreateDate") %><br>
					<% form = "SelfAssessment" %> 
					<% gid = GetSelfAssessment("SelfAssessmentID") %>
					<!--#include file="../includes/lastmodified_stamp.asp"-->
					
					</td>
				</tr>
				
		<!-- Begin Operational Section -->
		
		
		<% if printform = "No" then %>	
		

		
			<% if ReadOnlyLevel = 0 then %>					
				<tr>
					<td colspan="3" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold"  bgcolor="#c0c0c0"></td>
				</tr>
				<tr>
				
			<% else %>
				<tr>
					<td colspan="9" class="formMainCentered">Where did the <strong>Edit Button</strong> go?  <a href="..\helpfiles\surveyhelp.asp?HelpID=password1" onclick="NewWindow(this.href,'name','500','250','yes');return false;">Click Here</a> for an explanation.</td>
				</tr>	
				
			<% end if %>
			

		<% end if %>		

		<% if section = "Operational" then %>


				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" width=45%>Standard 1: The affiliate operates in compliance with applicable laws</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" width=45%>Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>
				
				<!-- Articles of Incorporation -->
				<tr>
					<td align="left" valign="top" class="formMain">Articles of Incorporation</td>
					<td align="left" valign="top" class="formMain">Review Articles of Incorporation; check for approved agency name</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std1a") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std1a") = 2 then%>In<%else if GetSelfAssessment("Std1a") = 1 then%>Out<%else%><font color="#FF0000"><font color="#FF0000"><font color="#FF0000">Not Entered</font></font></font><%end if%><%end if%><%end if%></td>
				</tr>
				<% if GetSelfAssessment("Std1aReason") <> null or GetSelfAssessment("Std1aReason") <> "" then%>
					<tr>
						
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std1aReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Tax-exempt status documentation / IRS Letter -->
				<tr>
					<td align="left" valign="top" class="formMain">Tax-exempt Status Documentation / IRS Letter</td>
					<td align="left" valign="top" class="formMain">Review tax exempt status documents; check for approved agency name</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std1b") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std1b") = 2 then%>In<%else if GetSelfAssessment("Std1b") = 1 then%>Out<%else%><font color="#FF0000"><font color="#FF0000">Not Entered</font></font><%end if%><%end if%><%end if%></td>					
				</tr>	
				<% if GetSelfAssessment("Std1bReason") <> null or GetSelfAssessment("Std1bReason") <> "" then%>
					<tr>
						
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std1bReason")%></td>
					</tr>
				<%end if%>			
				
				<!-- 990 form -->
				<tr>
					<td align="left" valign="top" class="formMain">990 Form</td>
					<td align="left" valign="top" class="formMain">990 has been filed with IRS for prior fiscal year</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Form990") = 3 then%>N/A<%else%><% if GetSelfAssessment("Form990") = 2 then%>In<%else if GetSelfAssessment("Form990") = 1 then%>Out<%else%><font color="#FF0000"><font color="#FF0000">Not Entered</font></font><%end if%><%end if%><%end if%></td>										
				</tr>
				<% if GetSelfAssessment("Form990Reason") <> null or GetSelfAssessment("Form990Reason") <> "" then%>
					<tr>
						
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Form990Reason")%></td>
					</tr>
				<%end if%>				
				
				<!-- Corporate Minutes -->
				<tr>
					<td align="left" valign="top" class="formMain">Corporate Minutes</td>
					<td align="left" valign="top" class="formMain">Board meeting minutes are on file and signed</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std1c") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std1c") = 2 then%>In<%else if GetSelfAssessment("Std1c") = 1 then%>Out<%else%><font color="#FF0000"><font color="#FF0000">Not Entered</font></font><%end if%><%end if%><%end if%></td>															
				</tr>
				<% if GetSelfAssessment("Std1cReason") <> null or GetSelfAssessment("Std1cReason") <> "" then%>
					<tr>
						
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std1cReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Corporate Bylaws -->
				<tr>
					<td align="left" valign="top" class="formMain">Corporate Bylaws</td>
					<td align="left" valign="top" class="formMain">Current copy of Bylaws are on file</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Bylaws") = 3 then%>N/A<%else%><% if GetSelfAssessment("Bylaws") = 2 then%>In<%else if GetSelfAssessment("Bylaws") = 1 then%>Out<%else%><font color="#FF0000"><font color="#FF0000">Not Entered</font></font><%end if%><%end if%><%end if%></td>															
				</tr>
				<% if GetSelfAssessment("BylawsReason") <> null or GetSelfAssessment("BylawsReason") <> "" then%>
					<tr>
						
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("BylawsReason")%></td>
					</tr>
				<%end if%>								
				
				<!-- Executed Affiliation Agreement -->
				<tr>
					<td align="left" valign="top" class="formMain">Executed Membership Affiliation Agreement (MAA)</td>
					<td align="left" valign="top" class="formMain">Signed MAA is on file and reflects current Service Community Area (SCA)</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("MAA") = 3 then%>N/A<%else%><% if GetSelfAssessment("MAA") = 2 then%>In<%else if GetSelfAssessment("MAA") = 1 then%>Out<%else%><font color="#FF0000"><font color="#FF0000">Not Entered</font></font><%end if%><%end if%><%end if%></td>															
				</tr>
				<% if GetSelfAssessment("MAAReason") <> null or GetSelfAssessment("MAAReason") <> "" then%>
					<tr>
						
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("MAAReason")%></td>
					</tr>
				<%end if%>		
				
				<!-- Logo and name -->
				<tr>
					<td align="left" valign="top" class="formMain">Affiliate uses, exclusively, the logo adopted by BBBSA and operates under a name approved by BBBSA</td>
					<td align="left" valign="top" class="formMain">Signage, stationery, business cards, publications, other materials  should all reflect consistent use of BBBSA logo and approved agency names</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std1LogoAndName") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std1LogoAndName") = 2 then%>In<%else if GetSelfAssessment("Std1LogoAndName") = 1 then%>Out<%else%><font color="#FF0000"><font color="#FF0000">Not Entered</font></font><%end if%><%end if%><%end if%></td>															
				</tr>
				<% if GetSelfAssessment("Std1LogoAndNameReason") <> null or GetSelfAssessment("Std1LogoAndNameReason") <> "" then%>
					<tr>
						
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std1LogoAndNameReason")%></td>
					</tr>
				<%end if%>
				
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
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std2") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std2") = 2 then%>In<%else if GetSelfAssessment("Std2") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>															
				</tr>
				<% if GetSelfAssessment("Std2Reason") <> null or GetSelfAssessment("Std2Reason") <> "" then%>
					<tr>
						
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std2Reason")%></td>
					</tr>
				<%end if%>		
				
				<!-- Board Recruitment Plan --
				<tr>
					<td align="left" valign="top" class="formMain">Board Recruitment Plan</td>
					<td align="left" valign="top" class="formMain">Review job descriptions, gap assessment, written recruitment plan, and board orientation</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std2aSO3a") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std2aSO3a") = 2 then%>In<%else if GetSelfAssessment("Std2aSO3a") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>															
				</tr>	
				
				<!-- Board Training Plan --
				<tr>
					<td align="left" valign="top" class="formMain">Board Training Plan</td>
					<td align="left" valign="top" class="formMain"><a target="_blank" href="http://agencyconnection.bbbs.org/atf/cf/{4CA344D5-890B-48AA-A80B-3EE1364E3AB7}/Self-Assessment%20Suggested%20Ideas.doc">Click Here</a> for recommendations<br>(MS Word Format)</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Brdtrainplan") = 3 then%>N/A<%else%><% if GetSelfAssessment("Brdtrainplan") = 2 then%>In<%else if GetSelfAssessment("Brdtrainplan") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>																	

				<!-- Documentation of annual review of board's performance -->
				<tr>
					<td align="left" valign="top" class="formMain">Documentation of annual review of board's performance</td>
					<td align="left" valign="top" class="formMain">Date and documentation that a review was conducted</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std2bSO3b") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std2bSO3b") = 2 then%>In<%else if GetSelfAssessment("Std2bSO3b") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std2bSO3bReason") <> null or GetSelfAssessment("Std2bSO3bReason") <> "" then%>
					<tr>
						
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std2bSO3bReason")%></td>
					</tr>
				<%end if%>	
				
				<!-- Documentation that board/advisory group representative attends annual national conference -->
				<tr>
					<td align="left" valign="top" class="formMain">Documentation that board/advisory group representative (s) attend national conference, regional conferences/meetings, workshops, and/ or trainings</td>
					<td align="left" valign="top" class="formMain">&nbsp;</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("MAA810conf") = 3 then%>N/A<%else%><% if GetSelfAssessment("MAA810conf") = 2 then%>In<%else if GetSelfAssessment("MAA810conf") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("MAA810confReason") <> null or GetSelfAssessment("MAA810confReason") <> "" then%>
					<tr>
						
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("MAA810confReason")%></td>
					</tr>
				<%end if%>		
				
				<!-- SO (Sponsoring Organization):  Written agreement between Corporate Board and Advisory Group re: voting representation and selection policy -->				
				<tr>
					<td align="left" valign="top" class="formMain"><em>Sponsored Only:</em>  Written agreement between Corporate Board and Advisory Group re: voting representation and selection policy</td>
					<td align="left" valign="top" class="formMain">Review bylaws, board minutes and governing board roster</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std2SO") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std2SO") = 2 then%>In<%else if GetSelfAssessment("Std2SO") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>										
				</tr>
				<% if GetSelfAssessment("Std2SOReason") <> null or GetSelfAssessment("Std2SOReason") <> "" then%>
					<tr>
						
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std2SOReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Mission/Vision -->
				<tr>
					<td colspan="3" class="formHeaderMedium" align="center">Mission/Vision</td>
				</tr>
				
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Standard 3/Standard 4 (sponsored programs): The affiliate has a clearly defined and articulated vision and mission statement that drives all agency decision making and provides focus for the assessment of the affiliate's work.</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>
				
				<!-- Written, board-approved mission statement -->
				<tr>
					<td align="left" valign="top" class="formMain">Written, board-approved mission statement</td>
					<td align="left" valign="top" class="formMain">Review mission statement to ensure that it is compatible with that of BBBSA and, as written, is used to drive decision-making</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std3SO4m") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std3SO4m") = 2 then%>In<%else if GetSelfAssessment("Std3SO4m") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std3SO4mReason") <> null or GetSelfAssessment("Std3SO4mReason") <> "" then%>
					<tr>
						
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std3SO4mReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Written, board-approved vision statement -->
				<tr>
					<td align="left" valign="top" class="formMain">Written, board-approved vision statement</td>
					<td align="left" valign="top" class="formMain">Review vision statement to ensure that it is compatible with that of BBBSA and, as written, is used to drive decision-making</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std3SO4v") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std3SO4v") = 2 then%>In<%else if GetSelfAssessment("Std3SO4v") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>										
				</tr>
				<% if GetSelfAssessment("Std3SO4vReason") <> null or GetSelfAssessment("Std3SO4vReason") <> "" then%>
					<tr>
						
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std3SO4vReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Strategic Planning -->
				<tr>
					<td colspan="3" class="formHeaderMedium" align="center">Strategic Planning </td>
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
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std4SO5") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std4SO5") = 2 then%>In<%else if GetSelfAssessment("Std4SO5") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std4SO5Reason") <> null or GetSelfAssessment("Std4SO5Reason") <> "" then%>
					<tr>
						
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std4SO5Reason")%></td>
					</tr>
				<%end if%>

				
				<!-- Quality Assurance -->
				<tr>
					<td colspan="3" class="formHeaderMedium" align="center">Quality Assurance</td>
				</tr>								
				
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Standard 5/Standard 6 (sponsored programs): The affiliate has a quality assurance system that ensures that all aspects of the affiliate's operations are reviewed and assessed on an annual basis, to include a review of its policies and procedures to ensure compliance with Standards of Practice for One-To-One Service related to program management for affiliates, and ensures that the affiliate is in compliance with its own program manual.</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>
				
				<!-- Documentation of annual review of all corporate policies and procedures -->
				<tr>
					<td align="left" valign="top" class="formMain">Documentation of annual review of all corporate policies and procedures</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std5opsSO6") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std5opsSO6") = 2 then%>In<%else if GetSelfAssessment("Std5opsSO6") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>
				</tr>
				<% if GetSelfAssessment("Std5opsSO6Reason") <> null or GetSelfAssessment("Std5opsSO6Reason") <> "" then%>
					<tr>
						
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std5opsSO6Reason")%></td>
					</tr>
				<%end if%>
				
				<!-- Documentation of annual review of Program Manual -->				
				<tr>
					<td align="left" valign="top" class="formMain">Documentation of annual review of Program Manual</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std5pgmSO6") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std5pgmSO6") = 2 then%>In<%else if GetSelfAssessment("Std5pgmSO6") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>
				</tr>
				<% if GetSelfAssessment("Std5pgmSO6Reason") <> null or GetSelfAssessment("Std5pgmSO6Reason") <> "" then%>
					<tr>
						
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std5pgmSO6Reason")%></td>
					</tr>
				<%end if%>
				
				<!-- Documentation of annual, random case file(s) audit -->
				<tr>
					<td align="left" valign="top" class="formMain">Documentation of annual, random case file(s) audit to ensure compliance with program standards </td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std5filesSO6") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std5filesSO6") = 2 then%>In<%else if GetSelfAssessment("Std5filesSO6") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std5filesSO6Reason") <> null or GetSelfAssessment("Std5filesSO6Reason") <> "" then%>
					<tr>
						
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std5filesSO6Reason")%></td>
					</tr>
				<%end if%>
				
				<!-- Fund Development -->
				<tr>
					<td colspan="3" class="formHeaderMedium" align="center">Fund Development</td>
				</tr>		
				
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Standard 6/Standard 7 (sponsoring organization): The affiliate has a financial management and fund development plan that ensures that fund development efforts are substantial enough to address current operation needs, contingencies, and planned growth.</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>			
				
				<!-- Documentation of board-approved annual budget -->
				<tr>
					<td align="left" valign="top" class="formMain">Documentation of annual review and board-approval of annual budget</td>
					<td align="left" valign="top" class="formMain">Documentation in board minutes that annual budget has been approved.</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std6SO7budget") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std6SO7budget") = 2 then%>In<%else if GetSelfAssessment("Std6SO7budget") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std6SO7budgetReason") <> null or GetSelfAssessment("Std6SO7budgetReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std6SO7budgetReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Budget includes expenses for training and travel to conferences -->
				<!--
				<tr>
					<td align="left" valign="top" class="formMain">Budget includes expenses for training and travel to conferences</td>
					<td align="left" valign="top" class="formMain">Identify the line in the budget for professional development / conferences</td>
					<td align="center" valign="top" class="formMain"><% ' if GetSelfAssessment("MAA810exp") = 2 then%>In<% 'else if GetSelfAssessment("MAA810exp") = 1 then%>Out<% 'else%><font color="#FF0000">Not Entered</font><% 'end if%><% 'end if%></td>					
				</tr>				
				-->
				 				
				<!-- Proof affiliate restricts its fund-raising activities to its own Service Community Area (SCA) or has written agreement with neighboring BBBSA affiliate -->				
				<tr>
					<td align="left" valign="top" class="formMain">Proof affiliate restricts its fund-raising activities to its own Service Community Area (SCA) or has written agreement with neighboring BBBSA affiliate</td>
					<td align="left" valign="top" class="formMain">If any fundraising activity is held in another BBBS' service community area, your written agreement with that BBBS agency, authorizing your fundraising activity must be on file</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("MAA32") = 3 then%>N/A<%else%><% if GetSelfAssessment("MAA32") = 2 then%>In<%else if GetSelfAssessment("MAA32") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("MAA32Reason") <> null or GetSelfAssessment("MAA32Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("MAA32Reason")%></td>
					</tr>
				<%end if%>
				
				
				<!-- Written, board-approved Fund Development Plan -->
				<tr>
					<td align="left" valign="top" class="formMain">Written, board-approved Fund Development Plan</td>
					<td align="left" valign="top" class="formMain">Review written, board-approved fundraising plan, including goals, diversification of funding, and planned revenue growth</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std6SO7b") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std6SO7b") = 2 then%>In<%else if GetSelfAssessment("Std6SO7b") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std6SO7bReason") <> null or GetSelfAssessment("Std6SO7bReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std6SO7bReason")%></td>
					</tr>
				<%end if%>	
				
				<!-- Financial Management -->
				<tr>
					<td colspan="3" class="formHeaderMedium" align="center">Financial Management</td>
				</tr>			
											
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Standard 7/Standard 8 (sponsoring organization): The affiliate has established financial management practices that meet generally accepted accounting practices and has an oversight structure that facilitates the early identification of potential problems</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>	
		
				
				<!-- Board oversight consistent with General Accounting Practices -->
				<tr>
					<td align="left" valign="top" class="formMain">Board oversight consistent with Generally Accepted Accounting Practices (GAAP)</td>
					<td align="left" valign="top" class="formMain">Documentation in board minutes that board reviews on a regular basis the agency's financials, including balance sheet; profit and loss statement; cash flow projections and budget variance report</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std6SO7") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std6SO7") = 2 then%>In<%else if GetSelfAssessment("Std6SO7") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std6SO7Reason") <> null or GetSelfAssessment("Std6SO7Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std6SO7Reason")%></td>
					</tr>
				<%end if%>
				
				<!-- Written financial management practices -->
				<tr>
					<td align="left" valign="top" class="formMain">Written board-approved financial management practices</td>
					<td align="left" valign="top" class="formMain">Documentation on file of the agency's financial management practices which should include, at a minimum, managing of deposits, check writing, authorization of expenditures, managing of donations, petty cash, etc. </td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std7SO8") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std7SO8") = 2 then%>In<%else if GetSelfAssessment("Std7SO8") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std7SO8Reason") <> null or GetSelfAssessment("Std7SO8Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std7SO8Reason")%></td>
					</tr>
				<%end if%>
				
				<!-- Annual financial audit from the last fiscal year -->
				<tr>
					<td align="left" valign="top" class="formMain">Annual audit of its financial condition, certified by an independent, certified public accounting firm and in accordance with generally accepted accounting principles (GAAP).</td>
					<td align="left" valign="top" class="formMain">Send copy of your most recently completed financial audit to the National Office.</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("MAA88") = 3 then%>N/A<%else%><% if GetSelfAssessment("MAA88") = 2 then%>In<%else if GetSelfAssessment("MAA88") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("MAA88Reason") <> null or GetSelfAssessment("MAA88Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("MAA88Reason")%></td>
					</tr>
				<%end if%>
				
				<!-- Affiliate is current with membership fees -->
				<tr>
					<td align="left" valign="top" class="formMain">Affiliate is current with membership fees</td>
					<td align="left" valign="top" class="formMain">Be able to track date and amount of payment of annual BBBSA fees, including any negotiated payment plans. Fee calculation forms are submitted on-line If agency affiliation fees are more than 6 months delinquent, a payment schedule must be approved and consistently followed.</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("MAA82") = 3 then%>N/A<%else%><% if GetSelfAssessment("MAA82") = 2 then%>In<%else if GetSelfAssessment("MAA82") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("MAA82Reason") <> null or GetSelfAssessment("MAA82Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("MAA82Reason")%></td>
					</tr>
				<%end if%>
							
				<!-- SO (Sponsoring Organization):  Documentation that funds raised or allocated to BBBSA program used solely for BBBSA expenses -->							
				<tr>
					<td align="left" valign="top" class="formMain"><strong><em>Sponsored Only:</em></strong>  Documentation that funds raised or allocated to BBBS program are used solely for BBBS expenses and that any share of administrative costs charged to BBBS program is reasonable.</td>
					<td align="left" valign="top" class="formMain">Sponsoring Organization's annual audit must include income and expenses of the BBBS program and indicate if BBBS funds are held in a separate account or if segregated accounting is used.</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("StdSO8a") = 3 then%>N/A<%else%><% if GetSelfAssessment("StdSO8a") = 2 then%>In<%else if GetSelfAssessment("StdSO8a") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("StdSO8aReason") <> null or GetSelfAssessment("StdSO8aReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("StdSO8aReason")%></td>
					</tr>
				<%end if%>
				
				<!-- SO (Sponsoring Organization):  Documentation that administrative costs charged to BBBSA program are reasonable, consistant and accurate --
				<tr>
					<td align="left" valign="top" class="formMain"><strong><em>Sponsored Only:</em></strong>  Documentation that administrative costs charged to BBBSA program are reasonable, consistant and accurate</td>
					<td align="left" valign="top" class="formMain">Assess the percent charged to the BBBS budget for administrative costs and the methodology used</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("StdSO8b") = 3 then%>N/A<%else%><% if GetSelfAssessment("StdSO8b") = 2 then%>In<%else if GetSelfAssessment("StdSO8b") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>		
				
				<!-- SO (Sponsoring Organization):  Proof that income and expense reports are provided the Advisory Group at least quarterly -->																	
				<tr>
					<td align="left" valign="top" class="formMain"><strong><em>Sponsored Only:</em></strong> Proof that income and expense reports are provided to the Advisory Group at least quarterly</td>
					<td align="left" valign="top" class="formMain">Documentation in Advisory Group minutes that BBBS revenue and expense reports are given to and reviewed by Advisory Group.</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("StdSO8c") = 3 then%>N/A<%else%><% if GetSelfAssessment("StdSO8c") = 2 then%>In<%else if GetSelfAssessment("StdSO8c") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("StdSO8cReason") <> null or GetSelfAssessment("StdSO8cReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("StdSO8cReason")%></td>
					</tr>
				<%end if%>
				
				
				<!--Risk Management -->
				<tr>
					<td colspan="3" class="formHeaderMedium" align="center">Risk Management</td>
				</tr>			
											
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Standard 8/Standard 9 (sponsoring organization): The affiliate has a risk management system that ensures that agency operational risks are identified and appropriately managed through insurance, and policies and procedures</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>	
				
				<!-- Written Crisis Management Plan --			
				<tr>
					<td align="left" valign="top" class="formMain">Written Crisis Management Plan</td>
					<td align="left" valign="top" class="formMain">Review currentCrisis Management Guide.  Get more information <a href="http://agencyconnection.bbbs.org/atf/cf/{4CA344D5-890B-48AA-A80B-3EE1364E3AB7}/Risk%20Mgm%20&%20Crisis%20Preparedness.doc" target="_blank">here</a>.</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std8SO9crisis") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std8SO9crisis") = 2 then%>In<%else if GetSelfAssessment("Std8SO9crisis") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				
				<!-- Written Risk Management Plan -->				
				<tr>
					<td align="left" valign="top" class="formMain">Written board-approved Risk Management Plan</td>
					<td align="left" valign="top" class="formMain">Review current Risk Management Plan to ensure that it contains the following components: governance, human resources, child safety & youth protection, financial management, fundraising and public relations, facility safety & security, technology and information management, insurance, transportation, and crisis management, including BBBSA's protocol for child abuse reporting. </td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std8SO9risk") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std8SO9risk") = 2 then%>In<%else if GetSelfAssessment("Std8SO9risk") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>
				</tr>
				<% if GetSelfAssessment("Std8SO9riskReason") <> null or GetSelfAssessment("Std8SO9riskReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std8SO9riskReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Proof of adequate insurance coverage -->
				<tr>
					<td align="left" valign="top" class="formMain">Proof of adequate insurance coverage that meets minimums established by BBBSA</td>
					<td align="left" valign="top" class="formMain">Check cover sheet of policies to assess levels of coverage for liability insurance that satisfies the risk management issues associated with the Standards of Practice. Insurance should cover, at a minimum, errors and omissions, bodily injury, property loss, sexual abuse and Director's and Officer's liability. </td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("MAA9") = 3 then%>N/A<%else%><% if GetSelfAssessment("MAA9") = 2 then%>In<%else if GetSelfAssessment("MAA9") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("MAA9Reason") <> null or GetSelfAssessment("MAA9Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("MAA9Reason")%></td>
					</tr>
				<%end if%>
				
				<!-- Personnel -->
				<tr>
					<td colspan="3" class="formHeaderMedium" align="center">Personnel</td>
				</tr>			
											
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Standard 9/Standard 10 (sponsoring organization): The affiliate employs a full time executive who is responsible to the board for the overall administration of agency operations</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>															
				
				<!-- Board approved job description for Executive (Program Director for sponsored programs) -->				
				<tr>
					<td align="left" valign="top" class="formMain">Written board-approved job description for Chief Executive (Program Director for Sponsored Programs) that specifies overall responsibility for employing, supervising, evaluating and terminating all paid and volunteer staff</td>
					<td align="left" valign="top" class="formMain">Current Job description for Chief Executive (Program Director for sponsored programs) should be kept in personnel file and referenced in personnel policies. </td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std10bSO11b2") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std10bSO11b2") = 2 then%>In<%else if GetSelfAssessment("Std10bSO11b2") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std10bSO11b2Reason") <> null or GetSelfAssessment("Std10bSO11b2Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std10bSO11b2Reason")%></td>
					</tr>
				<%end if%>		
				
				<!-- BBBS executive is employed full-time; for Sponsored Organizations, the affiliate employs a BBBS Program Director responsible for overall administration of BBBS Program operations -->
				<tr>
					<td align="left" valign="top" class="formMain">BBBS Chief Executive is employed full-time; and, for Sponsoring Organizations, a full-time Program Director is employed and responsible for overall administration of BBBS Program operations</td>
					<td align="left" valign="top" class="formMain">Confirmation on file that includes: Letter of Hire and/or time sheets/payroll</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std9SO10") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std9SO10") = 2 then%>In<%else if GetSelfAssessment("Std9SO10") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std9SO10Reason") <> null or GetSelfAssessment("Std9SO10Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std9SO10Reason")%></td>
					</tr>
				<%end if%>
				
				<!-- Documentation of annual performance evaluation of Executive (Program Director for sponsored programs) -->				
				<tr>
					<td align="left" valign="top" class="formMain">Annual performance evaluation of Chief Executive (Program Director for sponsored programs) is conducted in accordance with agency personnel polices, approved job description and annual performance goals</td>
					<td align="left" valign="top" class="formMain">Copy of annual performance evaluation is on-file in Chief Executive's personnel file</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("MAA813") = 3 then%>N/A<%else%><% if GetSelfAssessment("MAA813") = 2 then%>In<%else if GetSelfAssessment("MAA813") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>
				</tr>
				<% if GetSelfAssessment("MAA813Reason") <> null or GetSelfAssessment("MAA813Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("MAA813Reason")%></td>
					</tr>
				<%end if%>
				
				<!-- Definition of who notifies BBBSA of a vacancy in executive position (Program Director for sponsored progrems) -->
				<tr>
					<td align="left" valign="top" class="formMain">Notification of BBBSA of a vacancy in Chief Executive position (Program Director for sponsoring organizations)</td>
					<td align="left" valign="top" class="formMain">Identify where it is documented that BBBSA National Office must be contacted and by whom </td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std9bSO10b") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std9bSO10b") = 2 then%>In<%else if GetSelfAssessment("Std9bSO10b") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std9bSO10bReason") <> null or GetSelfAssessment("Std9bSO10bReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std9bSO10bReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Policies specify that executive (Program Director for sponsored programs) has “overall responsibility” for employing, supervising, evaluating and terminating all paid staff and volunteers --
				<tr>
					<td align="left" valign="top" class="formMain">Policies specify that executive (Program Director for sponsored programs) has “overall responsibility” for employing, supervising, evaluating and terminating all paid staff and volunteers</td>
					<td align="left" valign="top" class="formMain">Documented in the Personnel Manual and job description</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std9bSO10b") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std9bSO10b") = 2 then%>In<%else if GetSelfAssessment("Std9bSO10b") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>				
				
				<!-- Executive (Program Director for sponsored programs) attended new CEO training (new hires since 1/04) -->
				<tr>
					<td align="left" valign="top" class="formMain">Executive (Program Director for sponsored programs) attended new CEO training (new hires since 1/04)</td>
					<td align="left" valign="top" class="formMain">Check Personnel file for copy of transcript downloaded from the BBBS Learning Center</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("MAA814") = 3 then%>N/A<%else%><% if GetSelfAssessment("MAA814") = 2 then%>In<%else if GetSelfAssessment("MAA814") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("MAA814Reason") <> null or GetSelfAssessment("MAA814Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("MAA814Reason")%></td>
					</tr>
				<%end if%>
				
				<tr>
				
				</tr>	
				
				<!-- Standard 10/Standard 11 (sponsoring organization) -->
				
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Standard 10/Standard 11 (sponsoring organization): The affiliate, or BBBS program, has a human resource development and management system that is designed to effectively manage all paid, volunteer, and intern personnel</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>		
				
				<!-- Board-approved written personnel policies, compliant with local, state and federal labor laws -->
				<tr>
					<td align="left" valign="top" class="formMain">Written board-approved personnel policies, compliant with local, state and federal labor laws</td>
					<td align="left" valign="top" class="formMain">Ensure current personnel policies reflect all board approved changes, denoted by date and are provided to all staff, paid or volunteer.</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std10aSO11a") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std10aSO11a") = 2 then%>In<%else if GetSelfAssessment("Std10aSO11a") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>
				</tr>
				<% if GetSelfAssessment("Std10aSO11aReason") <> null or GetSelfAssessment("Std10aSO11aReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std10aSO11aReason")%></td>
					</tr>
				<%end if%>
									
				<!-- Written job descriptions for all paid and volunteer staff positions -->
				<tr>
					<td align="left" valign="top" class="formMain">Written job descriptions exist for all paid and volunteer staff positions</td>
					<td align="left" valign="top" class="formMain">Review job descriptions and update as necessary</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std10bSO11b") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std10bSO11b") = 2 then%>In<%else if GetSelfAssessment("Std10bSO11b") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std10bSO11bReason") <> null or GetSelfAssessment("Std10bSO11bReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std10bSO11bReason")%></td>
					</tr>
				<%end if%>

				<!-- Volunteers functioning in staff positions meet same personnel requirements and follow same policies and procedures -->
				<tr>
					<td align="left" valign="top" class="formMain">Volunteers functioning in staff positions meet same personnel requirements and follow same policies and procedures</td>
					<td align="left" valign="top" class="formMain">Policy statement in current Personnel Manual and in job descriptions</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std10gSO11f") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std10gSO11f") = 2 then%>In<%else if GetSelfAssessment("Std10gSO11f") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std10gSO11fReason") <> null or GetSelfAssessment("Std10gSO11fReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std10gSO11fReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Program Manual contains policies and procedures for non-degreed paraprofessionals -->
				<tr>
					<td align="left" valign="top" class="formMain">Agency Program Manual contains policies and procedures for the use of non-degreed paraprofessionals </td>
					<td align="left" valign="top" class="formMain">Ensure Program Manual has policies and procedures for non-degreed or paraprofessionals (persons with less than a Bachelor's degree) re: who will supervise and train them, and who will make all professional service delivery decisions.</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std10hSO11g2") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std10hSO11g2") = 2 then%>In<%else if GetSelfAssessment("Std10hSO11g2") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std10hSO11g2Reason") <> null or GetSelfAssessment("Std10hSO11g2Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std10hSO11g2Reason")%></td>
					</tr>
				<%end if%>
				
				<!-- Program Manual contains policies and procedures re: “professional/degreed staff” making all service delivery decisions --
				<tr>
					<td align="left" valign="top" class="formMain">Program Manual contains policies and procedures re: “professional/degreed staff” making all service delivery decisions</td>
					<td align="left" valign="top" class="formMain">Check Program Manual for clearly stated policies and procedures; determine which staff conduct service delivery and their degree (minimum is a Bachelors)</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std10hSO11g2") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std10hSO11g2") = 2 then%>In<%else if GetSelfAssessment("Std10hSO11g2") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>													
				
				<!-- Board approved, competitive salary ranges -->
				<tr>
					<td align="left" valign="top" class="formMain">Board develops and approves competitive salary ranges for all paid staff</td>
					<td align="left" valign="top" class="formMain">Documentation that the Board, or committee thereof, has reviewed current salary ranges against current market and determined competitive salary ranges</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std10dSO11c") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std10dSO11c") = 2 then%>In<%else if GetSelfAssessment("Std10dSO11c") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std10dSO11cReason") <> null or GetSelfAssessment("Std10dSO11cReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std10dSO11cReason")%></td>
					</tr>
				<%end if%>

				<!-- ED and program staff have at least a Bachelors degree -->
				<tr>
					<td align="left" valign="top" class="formMain">Chief Executive (Program Director for Sponsored affiliate) and program staff have at least a Bachelor's degree</td>
					<td align="left" valign="top" class="formMain">Review resumes, transcripts, diplomas in personnel file</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std10eSO11d") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std10eSO11d") = 2 then%>In<%else if GetSelfAssessment("Std10eSO11d") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std10eSO11dReason") <> null or GetSelfAssessment("Std10eSO11dReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std10eSO11dReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Confidential personnel records maintained -->
				<tr>
					<td align="left" valign="top" class="formMain">Confidential personnel records on each employee, paid or volunteer, are maintained at corporate office</td>
					<td align="left" valign="top" class="formMain">Personnel files should have a cover sheet documenting content and be located in a secured location.</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std10fSO11e") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std10fSO11e") = 2 then%>In<%else if GetSelfAssessment("Std10fSO11e") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std10fSO11eReason") <> null or GetSelfAssessment("Std10fSO11eReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std10fSO11eReason")%></td>
					</tr>
				<%end if%>

				<!-- Documentation of criminal history record check for staff / Volunteers -->
				<tr>
					<td align="left" valign="top" class="formMain">Documentation of criminal background check for staff / volunteers</td>
					<td align="left" valign="top" class="formMain">Copy of criminal background check and, driver's license/ proof of insurance, if appropriate, should be located in personnel files of staff and case files of volunteers if serving in staff role</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std10jSO11i") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std10jSO11i") = 2 then%>In<%else if GetSelfAssessment("Std10jSO11i") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std10jSO11iReason") <> null or GetSelfAssessment("Std10jSO11iReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std10jSO11iReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Documentation of attendance at BBBSA training offerings, annual meetings -->
				<tr>
					<td align="left" valign="top" class="formMain">Documentation of attendance at BBBSA training offerings, annual meetings</td>
					<td align="left" valign="top" class="formMain">Check Personnel file for copy of transcript downloaded from the BBBS Learning Center</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("MAA810") = 3 then%>N/A<%else%><% if GetSelfAssessment("MAA810") = 2 then%>In<%else if GetSelfAssessment("MAA810") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("MAA810Reason") <> null or GetSelfAssessment("MAA810Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("MAA810Reason")%></td>
					</tr>
				<%end if%>
				
				<!-- Documentation of annual personnel performance evaluations -->
				<tr>
					<td align="left" valign="top" class="formMain">Documentation of annual personnel performance evaluations</td>
					<td align="left" valign="top" class="formMain">Check Personnel files for copy of evaluation, signed by staff and supervisor</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std9bSO10b2") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std9bSO10b2") = 2 then%>In<%else if GetSelfAssessment("Std9bSO10b2") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std9bSO10b2Reason") <> null or GetSelfAssessment("Std9bSO10b2Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std9bSO10b2Reason")%></td>
					</tr>
				<%end if%>
				
				<!-- Non discrimination policy relative to staff -->
				<tr>
					<td align="left" valign="top" class="formMain">Written board-approved Non-discrimination policy relative to staff and volunteers.</td>
					<td align="left" valign="top" class="formMain">Documented, at a minimum, in the Personnel Policies</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std10iSO11h") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std10iSO11h") = 2 then%>In<%else if GetSelfAssessment("Std10iSO11h") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std10iSO11hReason") <> null or GetSelfAssessment("Std10iSO11hReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std10iSO11hReason")%></td>
					</tr>
				<%end if%>

				<!-- Standard 11/Standard 12 (sponsoring organization) -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0"  bgcolor="#c0c0c0">Standard 11/Standard 12 (sponsoring organization): The affiliate provides facilities and working conditions, which are conducive to accomplishing the operation of the affiliate including provisions to conduct private interviews, conforming to laws and regulations governing occupational health and safety</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0"  bgcolor="#c0c0c0">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0"  bgcolor="#c0c0c0"  bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>
				
				<!-- Facilities meet ADA, OSHA standards -->
				<tr>
					<td align="left" valign="top" class="formMain">Facilities meet ADA, OSHA standards</td>
					<td align="left" valign="top" class="formMain">Copy of annual facilities audit on file; Inspect the environment for safety and cleanliness; Inspect the equipment used by staff to perform necessary work for safety and proper functioning</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std11SO12") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std11SO12") = 2 then%>In<%else if GetSelfAssessment("Std11SO12") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std11SO12Reason") <> null or GetSelfAssessment("Std11SO12Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std11SO12Reason")%></td>
					</tr>
				<%end if%>
				
				<!-- Facilities allow for privacy during interviews -->
				<tr>
					<td align="left" valign="top" class="formMain">Facilities allow for privacy during interviews</td>
					<td align="left" valign="top" class="formMain">Assess that staff have private space for interviews</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std11SO122") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std11SO122") = 2 then%>In<%else if GetSelfAssessment("Std11SO122") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std11SO122Reason") <> null or GetSelfAssessment("Std11SO122Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std11SO122Reason")%></td>
					</tr>
				<%end if%>
				


			<!-- End Operational Section -->

		<% else %>

			<!-- Begin Program Section -->

				<!-- Standard 12/Standard 13 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Standard 12/Standard 13 (sponsoring organization): The Program Manual contains the policies, procedures, and forms to be used for implementing all One-To-One services</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>
				
				<!-- The Program Manual contains board-approved written policies, procedures and forms compliant with Practices of One-To-One Service -->
				<tr>
					<td align="left" valign="top" class="formMain">Written board-approved Agency Program Manual contains policies, procedures and forms compliant with the Standards of Practice of One-To-One Service</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval of policies; document in Program Manual</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std12aSO13a") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std12aSO13a") = 2 then%>In<%else if GetSelfAssessment("Std12aSO13a") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std12aSO13aReason") <> null or GetSelfAssessment("Std12aSO13aReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std12aSO13aReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Policy on eligibility criteria for volunteers & youth -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy and procedures on eligibility criteria for volunteers & youth and procedures for determining eligibility</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("policyeligible") = 3 then%>N/A<%else%><% if GetSelfAssessment("policyeligible") = 2 then%>In<%else if GetSelfAssessment("policyeligible") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("policyeligibleReason") <> null or GetSelfAssessment("policyeligibleReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("policyeligibleReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Procedures --
				<tr>
					<td align="left" valign="top" class="formMain">Procedures</td>
					<td align="left" valign="top" class="formMain">Procedures for determining eligibility</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("proceligible") = 3 then%>N/A<%else%><% if GetSelfAssessment("proceligible") = 2 then%>In<%else if GetSelfAssessment("proceligible") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>		
				
				<!-- Policy on youth outreach -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy and procedures on youth outreach</td>
					<td align="left" valign="top" class="formMain">Determine date of last Board review and approval</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("policychildrec") = 3 then%>N/A<%else%><% if GetSelfAssessment("policychildrec") = 2 then%>In<%else if GetSelfAssessment("policychildrec") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("policychildrecReason") <> null or GetSelfAssessment("policychildrecReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("policychildrecReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Procedures --
				<tr>
					<td align="left" valign="top" class="formMain">Procedures</td>
					<td align="left" valign="top" class="formMain">Procedures for recruiting youth</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("procchildrec") = 3 then%>N/A<%else%><% if GetSelfAssessment("procchildrec") = 2 then%>In<%else if GetSelfAssessment("procchildrec") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				
				<!-- Policy on volunteer recruitment -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy and procedures on volunteer recruitment</td>
					<td align="left" valign="top" class="formMain">Determine date of last Board review and approval</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("policyvolrec") = 3 then%>N/A<%else%><% if GetSelfAssessment("policyvolrec") = 2 then%>In<%else if GetSelfAssessment("policyvolrec") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("policyvolrecReason") <> null or GetSelfAssessment("policyvolrecReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("policyvolrecReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Procedures --
				<tr>
					<td align="left" valign="top" class="formMain">Procedures</td>
					<td align="left" valign="top" class="formMain">Procedures for recruiting volunteers</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("procvolrec") = 3 then%>N/A<%else%><% if GetSelfAssessment("procvolrec") = 2 then%>In<%else if GetSelfAssessment("procvolrec") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>		
				
				<!-- Policy on referrals -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy and procedures on referrals</td>
					<td align="left" valign="top" class="formMain">Determine date of last Board review and approval</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("policyref") = 3 then%>N/A<%else%><% if GetSelfAssessment("policyref") = 2 then%>In<%else if GetSelfAssessment("policyref") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("policyrefReason") <> null or GetSelfAssessment("policyrefReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("policyrefReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Procedures --
				<tr>
					<td align="left" valign="top" class="formMain">Procedures</td>
					<td align="left" valign="top" class="formMain">Procedures for handling referrals</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("procref") = 3 then%>N/A<%else%><% if GetSelfAssessment("procref") = 2 then%>In<%else if GetSelfAssessment("procref") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>									

				<!-- Policy on inquiries -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy and procedures on inquiries</td>
					<td align="left" valign="top" class="formMain">Determine date of last Board review and approval</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("policyinq") = 3 then%>N/A<%else%><% if GetSelfAssessment("policyinq") = 2 then%>In<%else if GetSelfAssessment("policyinq") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("policyinqReason") <> null or GetSelfAssessment("policyinqReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("policyinqReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Procedures --
				<tr>
					<td align="left" valign="top" class="formMain">Procedures</td>
					<td align="left" valign="top" class="formMain">Procedures for handling inquiries</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("procinq") = 3 then%>N/A<%else%><% if GetSelfAssessment("procinq") = 2 then%>In<%else if GetSelfAssessment("procinq") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>		
				
				<!-- Policies on intake -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy and procedures on intake</td>
					<td align="left" valign="top" class="formMain">Determine date of last Board review and approval</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("policyintake") = 3 then%>N/A<%else%><% if GetSelfAssessment("policyintake") = 2 then%>In<%else if GetSelfAssessment("policyintake") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("policyintakeReason") <> null or GetSelfAssessment("policyintakeReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("policyintakeReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Procedures --
				<tr>
					<td align="left" valign="top" class="formMain">Procedures</td>
					<td align="left" valign="top" class="formMain">Procedures for handling the intake process</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("procintake") = 3 then%>N/A<%else%><% if GetSelfAssessment("procintake") = 2 then%>In<%else if GetSelfAssessment("procintake") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>	
				
				<!-- Policies on matching -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy and procedures on matching</td>
					<td align="left" valign="top" class="formMain">Determine date of last Board review and approval</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("policymatch") = 3 then%>N/A<%else%><% if GetSelfAssessment("policymatch") = 2 then%>In<%else if GetSelfAssessment("policymatch") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("policymatchReason") <> null or GetSelfAssessment("policymatchReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("policymatchReason")%></td>
					</tr>
				<%end if%>
	
				<!-- Procedures --
				<tr>
					<td align="left" valign="top" class="formMain">Procedures</td>
					<td align="left" valign="top" class="formMain">Procedures for handling the matching process</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("procmatch") = 3 then%>N/A<%else%><% if GetSelfAssessment("procmatch") = 2 then%>In<%else if GetSelfAssessment("procmatch") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>			
				
				<!-- Policies on supervision -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy and procedures on supervision</td>
					<td align="left" valign="top" class="formMain">Determine date of last Board review and approval</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("policysup") = 3 then%>N/A<%else%><% if GetSelfAssessment("policysup") = 2 then%>In<%else if GetSelfAssessment("policysup") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("policysupReason") <> null or GetSelfAssessment("policysupReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("policysupReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Procedures --
				<tr>
					<td align="left" valign="top" class="formMain">Procedures</td>
					<td align="left" valign="top" class="formMain">Procedures for handling the match supervision process</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("procsup") = 3 then%>N/A<%else%><% if GetSelfAssessment("procsup") = 2 then%>In<%else if GetSelfAssessment("procsup") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>	
				
				<!-- Policies on closure -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy and procedures on closure</td>
					<td align="left" valign="top" class="formMain">Determine date of last Board review and approval</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("policyclosure") = 3 then%>N/A<%else%><% if GetSelfAssessment("policyclosure") = 2 then%>In<%else if GetSelfAssessment("policyclosure") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("policyclosureReason") <> null or GetSelfAssessment("policyclosureReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("policyclosureReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Procedures --
				<tr>
					<td align="left" valign="top" class="formMain">Procedures</td>
					<td align="left" valign="top" class="formMain">Procedures for handling the match closure process</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("procclosure") = 3 then%>N/A<%else%><% if GetSelfAssessment("procclosure") = 2 then%>In<%else if GetSelfAssessment("procclosure") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>	
				
				<!-- Policies on case record keeping -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy and procedures on case record keeping</td>
					<td align="left" valign="top" class="formMain">Determine date of last Board review and approval</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("policyrecords") = 3 then%>N/A<%else%><% if GetSelfAssessment("policyrecords") = 2 then%>In<%else if GetSelfAssessment("policyrecords") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%END IF%></td>					
				</tr>
				<% if GetSelfAssessment("policyrecordsReason") <> null or GetSelfAssessment("policyrecordsReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("policyrecordsReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Policies on handling documentation -->
				<tr>
					<td align="left" valign="top" class="formMain">Policies and procedures for handling documentation</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std12PPHandlingDoc") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std12PPHandlingDoc") = 2 then%>In<%else if GetSelfAssessment("Std12PPHandlingDoc") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std12PPHandlingDocReason") <> null or GetSelfAssessment("Std12PPHandlingDocReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std12PPHandlingDocReason")%></td>
					</tr>
				<%end if%>
				
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
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("policyovernite") = 3 then%>N/A<%else%><% if GetSelfAssessment("policyovernite") = 2 then%>In<%else if GetSelfAssessment("policyovernite") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("policyoverniteReason") <> null or GetSelfAssessment("policyoverniteReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("policyoverniteReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Policy on child sexual abuse prevention orientation, education, and training -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy on child sexual abuse prevention orientation, education, and training</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("policysexabuse") = 3 then%>N/A<%else%><% if GetSelfAssessment("policysexabuse") = 2 then%>In<%else if GetSelfAssessment("policysexabuse") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("policysexabuseReason") <> null or GetSelfAssessment("policysexabuseReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("policysexabuseReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Policy on board / staff serving as Bigs -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy on board / staff serving as Bigs</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("policystaffasbigs") = 3 then%>N/A<%else%><% if GetSelfAssessment("policystaffasbigs") = 2 then%>In<%else if GetSelfAssessment("policystaffasbigs") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("policystaffasbigsReason") <> null or GetSelfAssessment("policystaffasbigsReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("policystaffasbigsReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Policy on interviewing other persons residing with volunteer applicant -->
				<tr>
					<td align="left" valign="top" class="formMain">Policy on interviewing other persons residing with volunteer applicant</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("policyinterothers") = 3 then%>N/A<%else%><% if GetSelfAssessment("policyinterothers") = 2 then%>In<%else if GetSelfAssessment("policyinterothers") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("policyinterothersReason") <> null or GetSelfAssessment("policyinterothersReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("policyinterothersReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Procedures for obtaining information about disclosed prior BBBSA experience -->
				<tr>
					<td align="left" valign="top" class="formMain">Procedures for obtaining information about disclosed prior BBBSA experience</td>
					<td align="left" valign="top" class="formMain">Document date of last Board review and approval</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("policypriorexp") = 3 then%>N/A<%else%><% if GetSelfAssessment("policypriorexp") = 2 then%>In<%else if GetSelfAssessment("policypriorexp") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("policypriorexpReason") <> null or GetSelfAssessment("policypriorexpReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("policypriorexpReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Recommended Best Practice for Case File Audits -->
				<tr>
					<td colspan="3" class="formMainBold" align="center">FOR RECOMMENDED BEST PRACTICE for CASE FILE AUDITS, <a href="http://agencyconnection.bbbs.org/site/c.9dJGKRNqFmG/b.1742167/k.8DA1/Child_Safety.htm" target="_blank">click here</a> to consult our Child Safety Web Page and/or contact Julie Novak, Director of Child Safety and Quality Assurance, at <a href="mailto:Julie.Novak@bbbs.org">Julie.Novak@bbbs.org</a></td>
				</tr>
				
				
				<!-- Standard 13/Standard 14 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Standard 13/Standard 14 (sponsoring organization):</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>																		
													
				<!-- Procedures for obtaining information about disclosed prior BBBSA experience -->
				<tr>
					<td align="left" valign="top" class="formMain">The children, youth, and volunteer inquiry process used by the affiliate provides the opportunity for the affiliate, parent/guardian, and volunteer to determine the appropriateness of participation and provides an orientation to all services provided by the affiliates</td>
					<td align="left" valign="top" class="formMain">Review procedures and documentation of practice for inquiry and orientation to all services. Document date of last Board review and approval</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std13SO14") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std13SO14") = 2 then%>In<%else if GetSelfAssessment("Std13SO14") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std13SO14Reason") <> null or GetSelfAssessment("Std13SO14Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std13SO14Reason")%></td>
					</tr>
				<%end if%>
				
				<!-- Standard 14/Standard 15 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Standard 14/Standard 15 (sponsoring organization): The child intake process used by the affiliate is a consistent process to determine eligibility of children and youth for services based upon written eligibility criteria. Children and youth are not excluded on the basis of race, religion, national origin, gender, sexual orientation, disability, or marital status of parent</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>
				
				<!-- Written consent from parent / guardian -->
				<tr>
					<td align="left" valign="top" class="formMain">Written consent from parent / guardian</td>
					<td align="left" valign="top" class="formMain">Copy of Application signed by parent / guardian is in case file</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("childconsent") = 3 then%>N/A<%else%><% if GetSelfAssessment("childconsent") = 2 then%>In<%else if GetSelfAssessment("childconsent") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("childconsentReason") <> null or GetSelfAssessment("childconsentReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("childconsentReason")%></td>
					</tr>
				<%end if%>
				
				<!-- In-person interview with child -->
				<tr>
					<td align="left" valign="top" class="formMain">In-person interview with child</td>
					<td align="left" valign="top" class="formMain">Verify documentation of in-person interview is in case file</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("childinterview") = 3 then%>N/A<%else%><% if GetSelfAssessment("childinterview") = 2 then%>In<%else if GetSelfAssessment("childinterview") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("childinterviewReason") <> null or GetSelfAssessment("childinterviewReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("childinterviewReason")%></td>
					</tr>
				<%end if%>					
		
				<!-- In-person interview with parent / guardian (CBM only) -->
				<tr>
					<td align="left" valign="top" class="formMain">In-person interview with parent / guardian (CBM only)</td>
					<td align="left" valign="top" class="formMain">Verify documentation of in-person interview is in case file</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("childparinterview") = 3 then%>N/A<%else%><% if GetSelfAssessment("childparinterview") = 2 then%>In<%else if GetSelfAssessment("childparinterview") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("childparinterviewReason") <> null or GetSelfAssessment("childparinterviewReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("childparinterviewReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Assessment of home environment (CBM only) -->
				<tr>
					<td align="left" valign="top" class="formMain">Assessment of home environment (CBM only)</td>
					<td align="left" valign="top" class="formMain">Review procedures for home assessment in program manual; Verify documentation of home assessment is in case file</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("childhomeassess") = 3 then%>N/A<%else%><% if GetSelfAssessment("childhomeassess") = 2 then%>In<%else if GetSelfAssessment("childhomeassess") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("childhomeassessReason") <> null or GetSelfAssessment("childhomeassessReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("childhomeassessReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Standard 15/Standard 16 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Standard 15/Standard 16 (sponsoring organization): The professional staff conducts an in-person interview with the volunteer. The volunteer intake process elicits necessary information enabling the professional staff to prepare recommendations based upon the volunteer's ability to help meet the needs of the child</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>

				<!-- Application -->
				<tr>
					<td align="left" valign="top" class="formMain">Application</td>
					<td align="left" valign="top" class="formMain">Document the review of written application</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("volconsent") = 3 then%>N/A<%else%><% if GetSelfAssessment("volconsent") = 2 then%>In<%else if GetSelfAssessment("volconsent") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("volconsentReason") <> null or GetSelfAssessment("volconsentReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("volconsentReason")%></td>
					</tr>
				<%end if%>
					
				<!-- Obtain references (CBM = 3; SBM = 1) -->
				<tr>
					<td align="left" valign="top" class="formMain">Appropriate number of references are obtained </td>
					<td align="left" valign="top" class="formMain">Review procedures for obtaining references; Verify documentation of references is in case file</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("volreferences") = 3 then%>N/A<%else%><% if GetSelfAssessment("volreferences") = 2 then%>In<%else if GetSelfAssessment("volreferences") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("volreferencesReason") <> null or GetSelfAssessment("volreferencesReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("volreferencesReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Obtain criminal history record -->
				<tr>
					<td align="left" valign="top" class="formMain">Obtain criminal background check(s)</td>
					<td align="left" valign="top" class="formMain">Review procedures for obtaining criminal background checks in Program Manual; Verify documentation of criminal history record is in case file</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("volcriminal") = 3 then%>N/A<%else%><% if GetSelfAssessment("volcriminal") = 2 then%>In<%else if GetSelfAssessment("volcriminal") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("volcriminalReason") <> null or GetSelfAssessment("volcriminalReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("volcriminalReason")%></td>
					</tr>
				<%end if%>
					
				<!-- In-person interview --
				<tr>
					<td align="left" valign="top" class="formMain">In-person interview</td>
					<td align="left" valign="top" class="formMain">Verify documentation of in-person interview is in case file</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("volinterview") = 3 then%>N/A<%else%><% if GetSelfAssessment("volinterview") = 2 then%>In<%else if GetSelfAssessment("volinterview") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>			
				
				<!-- Assessment of home environment (CBM only) --
				<tr>
					<td align="left" valign="top" class="formMain">Assessment of home environment (CBM only)</td>
					<td align="left" valign="top" class="formMain">Review procedures for home assessment in program manual; Verify documentation of home assessment is in case file</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("volhomeassess") = 3 then%>N/A<%else%><% if GetSelfAssessment("volhomeassess") = 2 then%>In<%else if GetSelfAssessment("volhomeassess") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				
				<!-- Written professional matching recommendations -->
				<tr>
					<td align="left" valign="top" class="formMain">Written professional matching recommendations</td>
					<td align="left" valign="top" class="formMain">Review Program Manual for procedures; Verify documentation of written matching recommendations by professional staff is in case file</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("volmatching") = 3 then%>N/A<%else%><% if GetSelfAssessment("volmatching") = 2 then%>In<%else if GetSelfAssessment("volmatching") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("volmatchingReason") <> null or GetSelfAssessment("volmatchingReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("volmatchingReason")%></td>
					</tr>
				<%end if%>
		
				<!-- Provide opportunity for training -->
				<tr>
					<td align="left" valign="top" class="formMain">Provide opportunity for training</td>
					<td align="left" valign="top" class="formMain">Verify documentation that training opportunities have been offered to volunteers and parents/guardians, as needed </td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("voltraining") = 3 then%>N/A<%else%><% if GetSelfAssessment("voltraining") = 2 then%>In<%else if GetSelfAssessment("voltraining") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("voltrainingReason") <> null or GetSelfAssessment("voltrainingReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("voltrainingReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Standard 16/Standard 17 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Standard 16/Standard 17 (sponsoring organization): The matching process enables the professional staff to assess and take into consideration all information gathered through applications and interviews of all parties</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>								
		
				<!-- Child approves proposed match -->
				<tr>
					<td align="left" valign="top" class="formMain">Child approves proposed match</td>
					<td align="left" valign="top" class="formMain">Review program manual for procedures; Verify documentation that child approves is in case file</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Approveschild") = 3 then%>N/A<%else%><% if GetSelfAssessment("Approveschild") = 2 then%>In<%else if GetSelfAssessment("Approveschild") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("ApproveschildReason") <> null or GetSelfAssessment("ApproveschildReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("ApproveschildReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Parent / guardian approves proposed match -->
				<tr>
					<td align="left" valign="top" class="formMain">Parent / guardian approves proposed match</td>
					<td align="left" valign="top" class="formMain">Review program manual for procedures; Verify documentation that parent/guardian approves is in case filee</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Approvesparent") = 3 then%>N/A<%else%><% if GetSelfAssessment("Approvesparent") = 2 then%>In<%else if GetSelfAssessment("Approvesparent") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("ApprovesparentReason") <> null or GetSelfAssessment("ApprovesparentReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("ApprovesparentReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Volunteer approves proposed match -->
				<tr>
					<td align="left" valign="top" class="formMain">Volunteer approves proposed match</td>
					<td align="left" valign="top" class="formMain">Review program manual for procedures; Verify documentation that volunteer approves is in case file</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Approvesvol") = 3 then%>N/A<%else%><% if GetSelfAssessment("Approvesvol") = 2 then%>In<%else if GetSelfAssessment("Approvesvol") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("ApprovesvolReason") <> null or GetSelfAssessment("ApprovesvolReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("ApprovesvolReason")%></td>
					</tr>
				<%end if%>
				
				<!-- In-person match introduction by BBBSA staff or designee -->
				<tr>
					<td align="left" valign="top" class="formMain">In-person match introduction by BBBSA staff or designee</td>
					<td align="left" valign="top" class="formMain">Review program manual for procedures; Verify documentation of in-person match introduction is in case file</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Inpersonmatch") = 3 then%>N/A<%else%><% if GetSelfAssessment("Inpersonmatch") = 2 then%>In<%else if GetSelfAssessment("Inpersonmatch") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("InpersonmatchReason") <> null or GetSelfAssessment("InpersonmatchReason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("InpersonmatchReason")%></td>
					</tr>
				<%end if%>
				
				<!-- Standard 17/Standard 18 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Standard 17/Standard 18 (sponsoring organization)</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>		
				
				<!-- In-person match introduction by BBBSA staff or designee -->
				<tr>
					<td align="left" valign="top" class="formMain">Professional staff develops and annually updates an outcome-based plan for each match</td>
					<td align="left" valign="top" class="formMain">Review procedures in Program Manual; Verify documentation is in case file and that the annual outcome-based plan is complete, up-dated annually and on-file</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std17SO18") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std17SO18") = 2 then%>In<%else if GetSelfAssessment("Std17SO18") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std17SO18Reason") <> null or GetSelfAssessment("Std17SO18Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std17SO18Reason")%></td>
					</tr>
				<%end if%>
				
				<!-- Standard 18/Standard 19 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Standard 18/Standard 19 (sponsoring organization)</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>	
				
				<!-- Professional staff oversees regular supervisory contact -->
				<tr>
					<td align="left" valign="top" class="formMain">Professional staff oversees regular supervisory contact with volunteer, parent/guardian/ and child in accordance with the Program Manual and Standard of Practice for One-To-One Service.</td>
					<td align="left" valign="top" class="formMain">See Standard of Practice for One-To-One Service for minimum criteria; Review procedures in Program Manual; Verify documentation, indicating date and person contacted, to assure contact were made according to Standards</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std18SO19") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std18SO19") = 2 then%>In<%else if GetSelfAssessment("Std18SO19") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std18SO19Reason") <> null or GetSelfAssessment("Std18SO19Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std18SO19Reason")%></td>
					</tr>
				<%end if%>
				
				<!-- Standard 19/Standard 20 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Standard 19/Standard 20 (sponsoring organization)</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>		
				
				<!-- Professional staff conducts closure interviews -->
				<tr>
					<td align="left" valign="top" class="formMain">Professional staff conducts closure interviews with volunteer, parent/guardian, and child in accordance with the Program Manual </td>
					<td align="left" valign="top" class="formMain">Review program Manual to ensure policies for closure are current and closures are properly documented in case files</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std19SO20") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std19SO20") = 2 then%>In<%else if GetSelfAssessment("Std19SO20") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std19SO20Reason") <> null or GetSelfAssessment("Std19SO20Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std19SO20Reason")%></td>
					</tr>
				<%end if%>
				
				<!-- Standard 20/Standard 21 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Standard 20/Standard 21 (sponsoring organization)</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>			
				
				<!-- Professional staff reassesses program participants -->
				<tr>
					<td align="left" valign="top" class="formMain">Professional staff reassesses program participants in accordance with Program Manual </td>
					<td align="left" valign="top" class="formMain">Review Program Manual to ensure polices for reassessment are current and reassessments are properly documented in case files</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std20SO21") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std20SO21") = 2 then%>In<%else if GetSelfAssessment("Std20SO21") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std20SO21Reason") <> null or GetSelfAssessment("Std20SO21Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std20SO21Reason")%></td>
					</tr>
				<%end if%>

				<!-- Standard 21/Standard 22 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Standard 21/Standard 22 (sponsoring organization)</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>
				
				<!-- Policies and procedures regarding the management of confidential information -->
				<tr>
					<td align="left" valign="top" class="formMain">Written board approved policies and procedures, outlined in the Program Manual, regarding the management of confidential information</td>
					<td align="left" valign="top" class="formMain">Review the board-approved policy and procedures on confidentiality; Verify the consistent application.</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std21SO22") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std21SO22") = 2 then%>In<%else if GetSelfAssessment("Std21SO22") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std21SO22Reason") <> null or GetSelfAssessment("Std21SO22Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std21SO22Reason")%></td>
					</tr>
				<%end if%>
				
				<!-- Standard 22/Standard 23 (sponsoring organization): -->
				<tr>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Standard 22/Standard 23 (sponsoring organization)</td>
					<td align="left" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0" width="45%">Criteria; Support Materials</td>
					<td align="center" valign="top" class="formMainBold"  bgcolor="#c0c0c0" bgcolor="#c0c0c0">Compliance Level<br>(In/Out)</td>
				</tr>				
				
				<!-- Non discrimination policy relative to volunteer Bigs, and Board members -->
				<tr>
					<td align="left" valign="top" class="formMain">Written board approved non discrimination policy relative to volunteer Bigs, and Board members</td>
					<td align="left" valign="top" class="formMain">Document date of Board approval and verify where policy resides</td>
					<td align="center" valign="top" class="formMain"><% if GetSelfAssessment("Std22SO23") = 3 then%>N/A<%else%><% if GetSelfAssessment("Std22SO23") = 2 then%>In<%else if GetSelfAssessment("Std22SO23") = 1 then%>Out<%else%><font color="#FF0000">Not Entered</font><%end if%><%end if%><%end if%></td>					
				</tr>
				<% if GetSelfAssessment("Std22SO23Reason") <> null or GetSelfAssessment("Std22SO23Reason") <> "" then%>
					<tr>
						<td align="left" valign="top" class="formMain" colspan="3"><label style="color: #cc3300;">Reason for being out of compliance</label><br><%=GetSelfAssessment("Std22SO23Reason")%></td>
					</tr>
				<%end if%>
																													
						
		
			<!-- End Program Section -->
		
		<% end if %>



					
								
<% if printform = "No" then %>	
	

	
	<%  if ReadOnlyLevel = 0 then %>			
	
			
		<tr>
			<td colspan="3" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold"  bgcolor="#c0c0c0"></td>
		</tr>
		<tr>
		
	<%  else %>
		<tr>
			<td colspan="9" class="formMainCentered">Where did the <strong>Edit Button</strong> go?  <a href="..\helpfiles\surveyhelp.asp?HelpID=password1" onclick="NewWindow(this.href,'name','500','250','yes');return false;">Click Here</a> for an explanation.</td>
		</tr>	
		
	<%  end if %>
	
	
		<td colspan="3" align="center" class="formMain"><br>
		For Self-Assessment Questions, contact <a href="mailto:affiliatereview@bbbsa.org">affiliatereview@bbbsa.org</a>
		<!--#include file="../includes/contact_info.inc"-->
		</td>
		</tr>
		
<% end if %>
				
				
				</table>  
							
</form>

</td>
</tr>
</table>

<% end if %>