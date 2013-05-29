<% 

		<table width="550" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063">
			<tr>
				<td colspan="7" class="formHeader">PERFORMANCE</td>
			</tr>
			<tr>
				<td colspan="7" class="formMainBold">Created: <%= GetPerformance("CreateDate") %><br>
		<% form = "Performance" %> 
		<% gid = GetPerformance("PerformanceID") %>
		<%= GetPerformance("PerformanceID") %>
		<!--#include file="../includes/lastmodified_stamp.asp"-->
				</td>
			</tr>
			<tr>
				<td align="center" valign="middle" class="formMain">&nbsp;</td>
				<td align="center" valign="middle" class="formMain">Community Based</td>
				<td align="center" valign="middle" class="formMain">School Based</td>
				<td align="center" valign="middle" class="formMain">Other Site Based</td>
				<td align="center" valign="middle" class="formMain">Group Mentoring</td>
				<td align="center" valign="middle" class="formMain">Special Programs: Mentoring</td>
				<td align="center" valign="middle" class="formMain">Special Programs: Non-Mentoring</td>
			</tr>
			<tr>
				<td align="center" valign="middle" class="formMain">
				OPEN/ACTIVE&nbsp;matches<br>on&nbsp;the&nbsp;last&nbsp;day&nbsp;of<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("OpenMatchesCommunityBased") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("OpenMatchesSchoolBased") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("OpenMatchesOtherSiteBased") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("OpenMatchesGroupMentoring") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("OpenMatchesSpecialProgramsMentoring") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("OpenMatchesSpecialProgramsNonMentoring") %></td>
			</tr>
			<tr>
				<td align="center" valign="middle" class="formMain">
				Matches&nbsp;CLOSED&nbsp;during<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("ClosedMatchesCommunityBased") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("ClosedMatchesSchoolBased") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("ClosedMatchesOtherSiteBased") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("ClosedMatchesGroupMentoring") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("ClosedMatchesSpecialProgramsMentoring") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("ClosedMatchesSpecialProgramsNonMentoring") %></td>
			</tr>

<% if (y >= 2003) Then %>		
			<tr>
				<td align="center" valign="middle" class="formMain">
				AVERAGE&nbsp;LENGTH&nbsp;(In&nbsp;Months)<br>&nbsp;of&nbsp;Matches&nbsp;Closed&nbsp;during<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("AverageMatchLengthCB") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("AverageMatchLengthSB") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("AverageMatchLengthOSB") %></td>		
				<td align="right" valign="middle" class="formMain">n/a</td>	
				<td align="right" valign="middle" class="formMain">n/a</td>	
				<td align="right" valign="middle" class="formMain">n/a</td>					
			</tr>
			

			
<% End If %>			
			
			
<% If FBIAgency = 1 and (y > 2002 or (y = 2002 and m > 8) ) Then %>			
			
<!-- Fields added for Faith-Based / Incarcerated Questions -->

			<tr>
				<td colspan="7" class="formHeaderMedium">FAITH-BASED / INCARCERATED PARENTS REPORTING</td>				
			</tr>
			
			<tr>
				<td align="left">&nbsp;</td>				
				<td class="formMain" align="center">Community Based</td>
				<td class="formMain" align="center">School Based</td>
				<td class="formMain" align="center">Other Site Based</td>				
				<td align="left" colspan="3">&nbsp;</td>
			</tr>
			
			<tr>
				<td align="center" class="formMain">Faith-Based Partnerships <strong>And</strong> Children with Incarcerated Parents</td>
				<td align="right" class="formMain"><%= GetPerformance("CBIandFB") %></td>
				<td align="right" class="formMain"><%= GetPerformance("SBIandFB") %></td>
				<td align="right" class="formMain"><%= GetPerformance("OSBIandFB") %></td>
				<td align="left" colspan="3">&nbsp;</td>				

			</tr>
			
			<tr>
				<td align="center" class="formMain">Children with Incarcerated Parents <strong>Only</strong></td>			
				<td align="right" class="formMain"><%= GetPerformance("CBInotFB") %></td>
				<td align="right" class="formMain"><%= GetPerformance("SBInotFB") %></td>
				<td align="right" class="formMain"><%= GetPerformance("OSBInotFB") %></td>
				<td align="right" class="formMain" colspan="3">&nbsp;</td>
								
			</tr>
			
			<tr>
				<td align="center" class="formMain">Faith-Based Partnerships <strong>Only</strong></td>
				<td align="right" class="formMain"><%= GetPerformance("CBFBnotI") %></td>
				<td align="right" class="formMain"><%= GetPerformance("SBFBnotI") %></td>
				<td align="right" class="formMain"><%= GetPerformance("OSBFBnotI") %></td>
				<td align="right" class="formMain" colspan="3">&nbsp;</td>				
				
			</tr>			
			
<% End If %>			

<!-- Fields added for SBM Questions for June and December only -->
			
<% ' if (m=6 or m=12) and SBMAgency = 1 and y <> 2001 then 
	if m=12 and SBMAgency = 1 and y <> 2001 then	%>	
			<tr>
				<td colspan="7" class="formHeaderMedium">SCHOOL-BASED MENTORING GRANT PROGRESS REPORT</td>
			</tr>
			
			<tr>
				<td colspan="6" class="formMain">Number of Volunteers Currently in the Enrollment Process</td>
				<td colspan="1" align="right" valign="middle" class="formMain"><%= GetPerformance("SBMVolunteersInEnrollmentProcess") %></td>
			</tr>					
					
			<tr>
				<td colspan="6" class="formMain">Amount Raised Towards Match Pledge as of the last day of &nbsp<b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				<td colspan="1" align="right" valign="middle" class="formMain"><%= GetPerformance("SBMAmountRaisedTowardsMatchPledge") %></td>				
			</tr>				
<% end if %>


<!-- Revenue -->

<% if (y >= 2003) then %>

			<tr>
				<td colspan="7" class="formHeaderMedium">REVENUE</td>				
			</tr>
			<tr>
				<td align="center" valign="middle" class="formMain" colspan="4">Revenue&nbsp;<strong>for&nbsp;the&nbsp;Month&nbsp;of&nbsp;<%= MonthName(Request("m"), False) & " " & Request("y") %></strong></td>
				<td align="center" valign="middle" class="formMain" colspan="3">$<%= GetPerformance("Revenue")%></td>
			</tr>			

<% End If %>	

				<tr>
					<td colspan="7" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold"></td>
				</tr>
				
				<tr>
					<td colspan="7"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
				</tr>
				
				
			
				
				
		</table>


</body>
</html>
