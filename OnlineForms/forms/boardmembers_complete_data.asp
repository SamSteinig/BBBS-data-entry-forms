<% Dim ReadOnlyLevel
	If Session("ReadOnly") then
		ReadOnlyLevel=1
	Else
		ReadOnlyLevel=0
	End If
	%>

			<table width="550" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063">
			<form name="frmBoardMembers" action="BoardMembers_edit.asp?y=<%= Request("y") %>" method="post">
			<!--#include file="../includes/form_stamp.asp"-->
			<input type="hidden" name="status" value="editOld">
				<tr> 
					<td colspan="3" align="center" valign="top" class="formSubhead">BBBS - <%= y %> Annual Agency Information (AAI)</td>
				</tr>
				<%if printform="No" then %>
				<tr>
					<td colspan="3" class="formHeader">BOARD MEMBERS</td>
				</tr>
				<% else %>
				<tr>
					<td colspan="3" class="formIndex">BOARD MEMBERS</td>
				</tr>				
				<%end if%>
				
					<tr>
					<td colspan="3" class="formMainBold">Created: <%= GetBoardMembers("CreateDate") %>
		<% form = "BoardMembers" %> 
		<% gid = GetBoardMembers("BoardMembersID") %>
		<!--#include file="../includes/lastmodified_stamp.asp"-->
					</td>
					</tr>

<!-- Question Number 1 -->
				<tr>
					<td align="left" valign="top" class="formMain">1.</td>
					<td align="left" valign="top" class="formMain" colspan="2">Number of Board Members as of 06/30:&nbsp;<strong><%= GetBoardMembers("NumberBoardMembers") %></strong></td>

				</tr>

<!-- Question Number 2 -->
				<tr>
					<td align="left" valign="top" class="formMain">2.</td>
					<td colspan="2" align="left" valign="top" class="formMain">Number of <strong>FEMALE</strong>board members by ethnicity:<br>
<!-- nested table -->
						<table width="550" border="0" cellspacing="3" cellpadding="3" align="center">
							<tr>
								<td align="left" valign="top" class="formMain">White (Not Hispanic or Latino)<br><strong><%= GetBoardMembers("FemaleWhite") %></strong></td>
								<td align="left" valign="top" class="formMain">Black or African American (Not Hispanic or Latino)<br><strong><%= GetBoardMembers("FemaleBlack") %></strong></td>
								<td align="left" valign="top" class="formMain">Hispanic or Latino<br><strong><%= GetBoardMembers("FemaleHispanic") %></strong></td>
								<td align="left" valign="top" class="formMain">Asian (Not Hispanic or Latino)<br><strong><%= GetBoardMembers("FemaleAsian") %></strong></td>
								
							</tr>
							
            
							<tr>
								<td align="left" valign="top" class="formMain">Native Hawaiian or Other Pacific Islander(Not Hispanic or Latino)<br><strong><%= GetBoardMembers("FemaleIslander") %></strong></td>
								<td align="left" valign="top" class="formMain">American Indian or Alaska Native (Not Hispanic or Latino)<br><strong><%= GetBoardMembers("FemaleNative") %></strong></td>
								<td align="left" valign="top" class="formMain">Two or More Races (Not Hispanic or Latino)<br><strong><%= GetBoardMembers("FemaleMulti") %></strong></td>
								<td align="left" valign="top" class="formMain">Race missing or Unknown<br><strong><%= GetBoardMembers("FemaleUnknown") %></strong></td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td align="left" valign="top" class="formMain"><strong>TOTAL</strong><br><strong><%=GetBoardMembers("FemaleTotal")%></strong></td>																
							</tr>							
						</table> 
					</td>		
				</tr>
<!-- Question Number 3 -->
				<tr>
					<td align="left" valign="top" class="formMain">3.</td>
					<td colspan="2" align="left" valign="top" class="formMain">Number of <strong>MALE</strong> board members by ethnicity:<br>
<!-- nested table -->
						<table width="550" border="0" cellspacing="3" cellpadding="3" align="center">
							<tr>
								<td align="left" valign="top" class="formMain">White (Not Hispanic or Latino)<br><strong><right><%= GetBoardMembers("MaleWhite") %></strong></td>
								<td align="left" valign="top" class="formMain">Black or African American (Not Hispanic or Latino)<br><strong><%= GetBoardMembers("MaleBlack") %></strong></td>
								<td align="left" valign="top" class="formMain">Hispanic or Latino<br><strong><%= GetBoardMembers("MaleHispanic") %></strong></td>
								<td align="left" valign="top" class="formMain">Asian (Not Hispanic or Latino)<br><strong><%= GetBoardMembers("MaleAsian") %></strong></td>
							</tr>
							<tr>
								<td align="left" valign="top" class="formMain">Native Hawaiian or Other Pacific Islander(Not Hispanic or Latino)<br><strong><%= GetBoardMembers("MaleIslander") %></strong></td>
								<td align="left" valign="top" class="formMain">American Indian or Alaska Native (Not Hispanic or Latino)<br><strong><%= GetBoardMembers("MaleNative") %></strong></td>
								<td align="left" valign="top" class="formMain">Two or More Races (Not Hispanic or Latino)<br><strong><%= GetBoardMembers("MaleMulti") %></strong></td>
								<td align="left" valign="top" class="formMain">Race missing or Unknown<br><strong><%= GetBoardMembers("MaleUnknown") %></strong></td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td align="left" valign="top" class="formMain"><strong>TOTAL</strong><br><strong><right><%=GetBoardMembers("MaleTotal")%></strong></right></td>																
							</tr>
						</table>
					</td>		
				</tr>

<!-- Question Number 4 --> 
<!--				<tr> 
					<td align="left" valign="top" class="formMain">4.</td>
					<td colspan="3" align="left" valign="top" class="formMain">Do you have a policy of 100% Board donating?
						<% If (GetBoardMembers("BoardDonatingPolicy") = True) Then %> <strong>Yes</strong><%else%> <strong>No</strong><% End If %><br>
						<% If (GetBoardMembers("BoardDonatingPolicy") = True) Then %> 
							Minimum Donation Amount: <strong><%= formatcurrency(GetBoardMembers("MinimumBoardDonation"))%></strong>
						<% end if %>
					</td>
				</tr>
				
<!-- Question Number 4 --> 
				<tr> 
					<td align="left" valign="middle" class="formMain">4.</td>
					<td colspan="3" align="left" valign="top" class="formMain">Percentage of Board Members Donating to the Agency:&nbsp;<strong><%=GetBoardMembers("BoardDonationPrcnt")%>%</td></strong></td>
				</tr>								
				
				
<!-- Question Number 5 --> 
				<tr> 
					<td align="left" valign="middle" class="formMain">5</td>
					<td colspan="3" align="left" valign="top" class="formMain">Percentage of Board Members Connected the agency to potential Corporate and Individual Donors:&nbsp;<strong><%=GetBoardMembers("BoardConnectedPrcnt")%>%</td></strong></td>
				</tr>								
				
				
				
				
<!-- Question Number 6 --> 
				<tr> 
					<td align="left" valign="middle" class="formMain">6</td>
					<td colspan="3" align="left" valign="top" class="formMain">Average Donation by Board Members: <strong>$<strong><%=GetBoardMembers("AvgDonationBoardMember")%></strong></td>
				</tr>								


<!-- Question Number 7 --> 
      	<tr> 
			<td align="left" valign="top" class="formMain">7.</td>
			<td colspan="3" align="left" valign="top" class="formMain">Does your Agency has a Board Development Plan written in Last Year?
						<% If (GetBoardMembers("BoardDevelopmentPlan") = True) Then %> <strong>Yes</strong><%else%> <strong>No</strong><% End If %><br>
						<% If (GetBoardMembers("BoardDevelopmentPlan") = True) Then %> 
						<% end if %>
					</td>
				</tr>


<!-- Question Number 8 --> 
      	<tr> 
			<td align="left" valign="top" class="formMain">8.</td>
			<td colspan="3" align="left" valign="top" class="formMain">Has your board done an assessment in the last year?
						<% If (GetBoardMembers("AssessmentDone") = True) Then %> <strong>Yes</strong><%else%> <strong>No</strong><% End If %><br>
						<% If (GetBoardMembers("AssessmentDone") = True) Then %> 
						<% end if %>
					</td>
				</tr>





<!-- Question Number 9 --> 
				<tr> 
					<td align="left" valign="middle" class="formMain">9.</td>
					<td colspan="3" align="left" valign="top" class="formMain">Number of Board Members as of June 30th of current AAI year who are or have been a BIG:&nbsp;<strong><%=GetBoardMembers("BoardBigs")%></strong></td>
				</tr>				
				
				
				
				<% if printform="No" then %>
				
					<% if ReadOnlyLevel=0 then %>
						<tr>
							<td colspan="3" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold"></td>
						</tr>
						
					<% else %>
						<tr>
							<td colspan="9" class="formMainCentered">Where did the <strong>Edit Button</strong> go?  <a href="..\helpfiles\surveyhelp.asp?HelpID=password1" onclick="NewWindow(this.href,'name','500','250','yes');return false;">Click Here</a> for an explanation.</td>
						</tr>			
					<% end if %>						

				<tr>
					<td colspan="3"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
				</tr>
				<%end if%>
				

			</table>	
			</form>