<table width="650" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063">

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

<%

dim DisplayMetrics
DisplayMetrics = 1

%>

<% Dim ReadOnlyLevel
If Session("ReadOnly") then
	ReadOnlyLevel=1
Else
	ReadOnlyLevel=0
End If
%>


<br>

<% if printform = "No" then %>

	<tr>
		<td colspan="10" class="formHeader">SDM METRIC COMPONENTS - <%= MonthName(Request("m"), False) & " " & Request("y") %></td>
	</tr>
			
<% else %>			

	<tr>
		<td colspan="10" class="formIndex">SDM METRIC COMPONENTS - <%= MonthName(Request("m"), False) & " " & Request("y") %></td>
	</tr>	
			
<% end if %>

		<tr>
		<!-- Date / Time Info -->
		
			<td colspan="10" class="formMainBold">Created: <%= GetPerformance("CreateDate") %><br>		
			
			<% form = "SDMPerformance" %> 
			<% gid = GetPerformance("SDMPerformanceID") %>
			<!--#include file="../includes/lastmodified_stamp.asp"-->
			</td>
		</tr>
		
			
			

<% if DisplayMetrics = 1 then %>
		
	<% if y < 2006 then %>
		

		<% if Printform = "No" then %>
			
			<TR>
				<TD colspan="7" class="formHeaderMedium">AVERAGE MATCH LENGTH</TD>
			</TR>
		<% else %>
			<TR>
				<TD colspan="7" class="formMain"><div align="center"><strong>AVERAGE MATCH LENGTH</strong></div></TD>
			</TR>		
		<% end if %>	
		
			<tr>
				<td>&nbsp;</td>
				<td class="formMain" colspan="2" align="center">Community-Based</td>
				<td class="formMain" colspan="2" align="center">School-Based</td>				
				<td class="formMain" colspan="2" align="center">Non-School<br>Site-Based</td>				
			</tr>	
		
			<tr>
				<td align="center" valign="middle" class="formMain">
					AVERAGE&nbsp;LENGTH&nbsp;(In&nbsp;Months)<br>&nbsp;of&nbsp;Matches&nbsp;Closed&nbsp;during<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
					<td align="center" colspan = "2" valign="middle" class="formMain"><%= GetPerformance("AverageMatchLengthCB") %></td>
					<td align="center" colspan = "2" valign="middle" class="formMain"><%= GetPerformance("AverageMatchLengthSB") %></td>
					<td align="center" colspan = "2" valign="middle" class="formMain"><%= GetPerformance("AverageMatchLengthOSB") %></td>
			</tr>		
		
	<% end if %>
	
		<% if y < 2006 then %>	
		<!-- Only Display Yield Heading for older components (prior to 2006) -->	
			<TR>
				<TD colspan="7" class="formHeaderMedium">YIELD AND PROCESSING TIME</TD>
			</TR>	
				
		<% end if %>
			
			<% if PrintForm="No" then %>
			<TR>
				<TD colspan="10" <%if y < 2006 then %>class="formMain"<%else%>class="formHeaderMedium"<%end if%> align="center"><strong>Volunteer</strong></TD>
			</TR>		
			<% else %>
			<TR>
				<TD colspan="10" class="formMain"><div align="center"><strong>Volunteer</strong></div></TD>
			</TR>					
			<% end if %>
			
			
			<tr>
				<td>&nbsp;</td>
				<td colspan="2" class="formMain" align="center" width="150">Community-Based</td>
				<td colspan="2" class="formMain" align="center" width="150">School-Based</td>
				<td colspan="2" class="formMain" align="center" width="150"> Non-School<br>Site-Based</td>			
				<% if y >= 2006 then %>
					<td colspan="2" class="formMain" align="center" width="150">Totals</td>			
				<% end if %>
						
			</tr>
			
			<tr>
				<td>&nbsp;</td>
				<% if y < 2006 then %>
					<td class="formMain">Number of Individuals</td>
					<td class="formMain">Average Days</td>
				<% else %>
					<td class="formMain" colspan="2" align="center" width="150">Number of Individuals</td>				
				<% end if %>
				<% if y < 2006 then %>
					<td class="formMain">Number of Individuals</td>
					<td class="formMain">Average Days</td>	
				<% else %>
					<td class="formMain" colspan="2" align="center" width="150">Number of Individuals</td>				
				<% end if %>
				<% if y < 2006 then %>
					<td class="formMain">Number of Individuals</td>
					<td class="formMain">Average Days</td>
				<% else %>
					<td class="formMain" colspan="2" align="center" width="150">Number of Individuals</td>
					<td>&nbsp;</td>		
				<% end if %>
				
			</tr>
			
			<tr>
				<td class="formMain" align="left" width="225">Volunteer Inquiries</td>
				<% if y < 2006 then %>
					<td class="formMain"  align="center"><%= GetPerformance("YieldRate_Vol_Inquiries_CB") %></td>				
					<td class="formMain" bgcolor="#c0c0c0">&nbsp;</td>
				<% else %>
					<td class="formMain"  align="center" colspan="2"><%= GetPerformance("YieldRate_Vol_Inquiries_CB") %></td>								
				<% end if %>
				<% if y < 2006 then %>
					<td class="formMain" align="center"><%= GetPerformance("YieldRate_Vol_Inquiries_SB") %></td>								
					<td class="formMain" bgcolor="#c0c0c0">&nbsp;</td>		
				<% else %>
					<td class="formMain" colspan = "2" align="center"><%= GetPerformance("YieldRate_Vol_Inquiries_SB") %></td>												
				<% end if %>		
				<% if y < 2006 then %>
					<td class="formMain" align="center"><%= GetPerformance("YieldRate_Vol_Inquiries_OSB") %></td>												
					<td class="formMain" bgcolor="#c0c0c0">&nbsp;</td>		
				<% else %>
					<td class="formMain" align="center" colspan="2"><%= GetPerformance("YieldRate_Vol_Inquiries_OSB") %></td>	
					<% dim VolInqTotal
					VolInqTotal = GetPerformance("YieldRate_Vol_Inquiries_CB") + GetPerformance("YieldRate_Vol_Inquiries_SB") + GetPerformance("YieldRate_Vol_Inquiries_OSB")
					%>
					<td class="formMain" align="center" colspan="2"><%=VolInqTotal%></td>						
																				
				<% end if %>						
			</tr>			
			
			
			
			<tr>
				<td class="formMain">Volunteer Interviews</strong></td>
				<% if y < 2006 then %>
					<td class="formMain" align="center"><%= GetPerformance("ProcTim_Vol_InquiryToInterview_Number_Comm") %></td>
					<td class="formMain" align="center"><%= GetPerformance("ProcTim_Vol_InquiryToInterview_AveDays_Comm") %></td>
				<% else %>
					<td class="formMain" align="center" colspan = "2"><%= GetPerformance("ProcTim_Vol_InquiryToInterview_Number_Comm") %></td>				
				<% end if %>
				<% if y < 2006 then %>
					<td class="formMain" align="center"><%= GetPerformance("ProcTim_Vol_InquiryToInterview_Number_School") %></td>
					<td class="formMain" align="center"><%= GetPerformance("ProcTim_Vol_InquiryToInterview_AveDays_School") %></td>
				<% else %>
					<td class="formMain" align="center" colspan = "2"><%= GetPerformance("ProcTim_Vol_InquiryToInterview_Number_School") %></td>				
				<% end if %>
				<% if y < 2006 then %>
					<td class="formMain" align="center"><%= GetPerformance("ProcTim_Vol_InquiryToInterview_Number_Other") %></td>
					<td class="formMain" align="center"><%= GetPerformance("ProcTim_Vol_InquiryToInterview_AveDays_Other") %></td>				
				<% else %>
					<td class="formMain" align="center" colspan="2"><%= GetPerformance("ProcTim_Vol_InquiryToInterview_Number_Other") %></td>				
					<% dim VolIntTotal
					VolIntTotal = GetPerformance("ProcTim_Vol_InquiryToInterview_Number_Comm") + GetPerformance("ProcTim_Vol_InquiryToInterview_Number_School") + GetPerformance("ProcTim_Vol_InquiryToInterview_Number_Other")
					%>
					<td class="formMain" align="center" colspan="2"><%=VolIntTotal%></td>						
				<% end if %>
			</tr>

	<% if y < 2006 then %>			
			<tr>
				<td class="formMain">Volunteer Interview <strong>to Matched</strong></td>
				<td class="formMain" align="center"><%= GetPerformance("ProcTim_Vol_InterviewToMatched_Number_Comm") %></td>
				<td class="formMain" align="center"><%= GetPerformance("ProcTim_Vol_InterviewToMatched_AveDays_Comm") %></td>
				<td class="formMain" align="center"><%= GetPerformance("ProcTim_Vol_InterviewToMatched_Number_School") %></td>
				<td class="formMain" align="center"><%= GetPerformance("ProcTim_Vol_InterviewToMatched_AveDays_School") %></td>
				<td class="formMain" align="center"><%= GetPerformance("ProcTim_Vol_InterviewToMatched_Number_Other") %></td>
				<td class="formMain" align="center"><%= GetPerformance("ProcTim_Vol_InterviewToMatched_AveDays_Other") %></td>				
			</tr>
			
	<% end if %>
			
			
			<% if PrintForm="No" then %>
			<TR>
				<TD colspan="10"  <%if y < 2006 then %>class="formMain"<%else%>class="formHeaderMedium"<%end if%> align="center"><strong>Child</strong></TD>
			</TR>		
			<% else %>
			<TR>
				<TD colspan="10" class="formMain"><div align="center"><strong>Child</strong></div></TD>
			</TR>					
			<% end if %>
				
			
			<tr>
				<td>&nbsp;</td>
				<td colspan="2" class="formMain" align="center">Community-Based</td>
				<td colspan="2" class="formMain" align="center">School-Based</td>
				<td colspan="2" class="formMain" align="center">Non-School<br>Site-Based</td>				
				<% if y >= 2006 then %>
					<td colspan="2" class="formMain" align="center">Totals</td>													
				<% end if %>
			</tr>
			
			<tr>
				<td>&nbsp;</td>
				<% if y < 2006 then %>
					<td class="formMain">Number of Individuals</td>
					<td class="formMain">Average Days</td>
				<% else %>
					<td class="formMain" colspan="2" align="center">Number of Individuals</td>				
				<% end if %>
				<% if y < 2006 then %>
					<td class="formMain">Number of Individuals</td>
					<td class="formMain">Average Days</td>	
				<% else %>
					<td class="formMain" colspan = "2" align="center">Number of Individuals</td>				
				<% end if %>
				<% if y < 2006 then %>
					<td class="formMain">Number of Individuals</td>
					<td class="formMain">Average Days</td>
				<% else %>
					<td class="formMain" colspan="2" align="center">Number of Individuals</td>		
					<td>&nbsp;</td>		
				<% end if %>
			</tr>
			
			<tr>
				<td class="formMain" align="left">Child Inquiries</td>
				<% if y < 2006 then %>
					<td class="formMain" align="center"><%= GetPerformance("YieldRate_Youth_Inquiries_CB") %></td>		
					<td class="formMain" bgcolor="#c0c0c0">&nbsp;</td>		
				<% else %>
					<td class="formMain" align="center" colspan = "2"><%= GetPerformance("YieldRate_Youth_Inquiries_CB") %></td>						
				<% end if %>	
				
				<% if y < 2006 then %>							
					<td class="formMain" align="center"><%= GetPerformance("YieldRate_Youth_Inquiries_SB") %></td>								
					<td class="formMain" bgcolor="#c0c0c0">&nbsp;</td>				
				<% else %>
					<td class="formMain" align="center" colspan = "2"><%= GetPerformance("YieldRate_Youth_Inquiries_SB") %></td>								
				<% end if %>	
				
				<% if y < 2006 then %>
					<td class="formMain" align="center"><%= GetPerformance("YieldRate_Youth_Inquiries_OSB") %></td>												
					<td class="formMain" bgcolor="#c0c0c0">&nbsp;</td>		
				<% else %>
					<td class="formMain" align="center" colspan="2"><%= GetPerformance("YieldRate_Youth_Inquiries_OSB") %></td>																		
					<%
					dim ChildInqTotal
					ChildInqTotal = GetPerformance("YieldRate_Youth_Inquiries_CB") + GetPerformance("YieldRate_Youth_Inquiries_SB") + GetPerformance("YieldRate_Youth_Inquiries_OSB")
					%>
					<td class="formMain" align="center" colspan="2"><%= ChildInqTotal %></td>																							
				<% end if %>
			</tr>			
			
			
			
			<tr>
				<td class="formMain">Child Interviews</strong></td>
				<% if y < 2006 then %>
					<td class="formMain" align="center"><%= GetPerformance("ProcTim_Youth_InquiryToInterview_Number_Comm") %></td>
					<td class="formMain" align="center"><%= GetPerformance("ProcTim_Youth_InquiryToInterview_AveDays_Comm") %></td>
				<% else %>
					<td class="formMain" align="center" colspan="2"><%= GetPerformance("ProcTim_Youth_InquiryToInterview_Number_Comm") %></td>				
				<% end if %>
				<% if y < 2006 then %>
					<td class="formMain" align="center"><%= GetPerformance("ProcTim_Youth_InquiryToInterview_Number_School") %></td>
					<td class="formMain" align="center"><%= GetPerformance("ProcTim_Youth_InquiryToInterview_AveDays_School") %></td>
				<% else %>
					<td class="formMain" align="center" colspan="2"><%= GetPerformance("ProcTim_Youth_InquiryToInterview_Number_School") %></td>				
				<% end if %>
				<% if y < 2006 then %>
					<td class="formMain" align="center"><%= GetPerformance("ProcTim_Youth_InquiryToInterview_Number_Other") %></td>
					<td class="formMain" align="center"><%= GetPerformance("ProcTim_Youth_InquiryToInterview_AveDays_Other") %></td>				
				<% else %>
					<td class="formMain" align="center" colspan = "2"><%= GetPerformance("ProcTim_Youth_InquiryToInterview_Number_Other") %></td>				
					<%
					dim ChildIntTotal
					ChildIntTotal = GetPerformance("ProcTim_Youth_InquiryToInterview_Number_Comm") + GetPerformance("ProcTim_Youth_InquiryToInterview_Number_School") + GetPerformance("ProcTim_Youth_InquiryToInterview_Number_Other")
					%>
					<td class="formMain" align="center" colspan = "2"><%= ChildIntTotal %></td>									
				<% end if %>
			</tr>

	<% if y < 2006 then %>			
			
			<tr>
				<td class="formMain">Child Interview <strong>to Matched</strong></td>
				<td class="formMain" align="center"><%= GetPerformance("ProcTim_Youth_InterviewToMatched_Number_Comm") %></td>
				<td class="formMain" align="center"><%= GetPerformance("ProcTim_Youth_InterviewToMatched_AveDays_Comm") %></td>
				<td class="formMain" align="center"><%= GetPerformance("ProcTim_Youth_InterviewToMatched_Number_School") %></td>
				<td class="formMain" align="center"><%= GetPerformance("ProcTim_Youth_InterviewToMatched_AveDays_School") %></td>
				<td class="formMain" align="center"><%= GetPerformance("ProcTim_Youth_InterviewToMatched_Number_Other") %></td>
				<td class="formMain" align="center"><%= GetPerformance("ProcTim_Youth_InterviewToMatched_AveDays_Other") %></td>				
			</tr>
	<% end if %>
			
	<% if y < 2006 then %>
	
		<% if PrintForm="No" then %>
			<tr>
				<TD colspan="7" class="formHeaderMedium">NUMBER OF MATCH CLOSURES</TD>	
			</tr>
		<% else %>
			<tr>
				<TD colspan="7" class="formMain"><strong><div align="center">NUMBER OF MATCH CLOSURES</div></strong></TD>	
			</tr>		
		<% end if %>
			
			<tr>		
				<td>&nbsp;</td>
				<td class="formMain" colspan="2" align="center">Community-Based</td>
				<td class="formMain" colspan="2" align="center">School-Based</td>	
				<td class="formMain" colspan="2" align="center">Non-School<br>Site-Based</td>							
			</tr>
			
			<tr>
				<td class="formMain">Less Than 3 Months</td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Freq_Under3Months_Comm") %></td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Freq_Under3Months_School") %></td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Freq_Under3Months_Other") %></td>				
			</tr>
			
			<tr>
				<td class="formMain">3-6 Months</td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Freq_3To6Months_Comm") %></td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Freq_3To6Months_School") %></td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Freq_3To6Months_Other") %></td>				
			</tr>		
			
			<tr>
				<td class="formMain">7-9 Months</td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Freq_7To9Months_Comm") %></td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Freq_7To9Months_School") %></td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Freq_7To9Months_Other") %></td>				
			</tr>		
			
			<tr>
				<td class="formMain">10-12 Months</td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Freq_10To12Months_Comm") %></td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Freq_10To12Months_School") %></td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Freq_10To12Months_Other") %></td>				
			</tr>
			
			<tr>
				<td class="formMain">13-23 Months</td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Freq_13To23Months_Comm") %></td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Freq_13To23Months_School") %></td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Freq_13To23Months_Other") %></td>				
			</tr>
			
			<tr>
				<td class="formMain">24 or More Months</td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Freq_24OrMoreMonths_Comm") %></td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Freq_24OrMoreMonths_School") %></td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Freq_24OrMoreMonths_Other") %></td>				
			</tr>
			
			<% if PrintForm="No" then %>
				<tr>
					<TD colspan="7" class="formHeaderMedium">VOLUNTEERS RE-MATCHED</TD>	
				</tr>		
			<% else %>
				<tr>
					<TD colspan="7" class="formMain"><div align="center"><strong>VOLUNTEERS RE-MATCHED</strong></div></TD>	
				</tr>				
			<% end if %>	
			
			<tr>
				<td>&nbsp;</td>
				<td class="formMain" colspan="2" align="center">Community-Based</td>
				<td class="formMain" colspan="2" align="center">School-Based</td>				
				<td class="formMain" colspan="2" align="center">Non-School<br>Site-Based</td>				
			</tr>											
			
			<tr>
				<td class="formMain">Volunteers Re-Matched</td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Volunteers_ReMatchedCB") %></td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Volunteers_ReMatchedSB") %></td>				
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("Volunteers_ReMatchedOSB") %></td>				
				
				
			</tr>
			



			<% if PrintForm="No" then %>
				<tr>
					<TD colspan="7" class="formHeaderMedium">PREMATURE CLOSURE</TD>
				</tr>
			<% else %>
				<tr>
					<TD colspan="7" class="formMain"><div align="center"><strong>PREMATURE CLOSURE</strong></div></TD>
				</tr>		
			<% end if %>
			
			<TR>
				<TD colspan="1">&nbsp;</TD>
				<TD colspan="2" class="formMain" align="center">Community-Based</TD>
				<TD colspan="2" class="formMain" align="center">School-Based</TD>
				<TD colspan="2" class="formMain" align="center">Non-School<br>Site-Based</TD>				
	
			</TR>
			
			<tr>
				<td colspan="1" class="formMain">Number of Matches that Closed Prematurely</td>			
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("CBNumberClosedPrematurely") %></td>				
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("SBNumberClosedPrematurely") %></td>								
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("OSBNumberClosedPrematurely") %></td>												
			</tr>
			
			<% if PrintForm="No" then %>
				<tr>
					<TD colspan="7" class="formHeaderMedium">CLOSE CODES</TD>	
				</tr>
			<% else %>
				<tr>
					<TD colspan="7" class="formMain"><strong><div align="center">CLOSE CODES</div></strong></TD>	
				</tr>			
			<% end if %>
			
			<TR>
				<TD colspan="1">&nbsp;</TD>
				<TD colspan="2" class="formMain" align="center">Community-Based</TD>
				<TD colspan="2" class="formMain" align="center">School-Based</TD>
				<TD colspan="2" class="formMain" align="center">Non-School<br>Site-Based</TD>				
	
			</TR>
			
			<tr>
				<td colspan="1" class="formMain">Child/Parent Status Change</td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("CBChildParentStatusChange") %></td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("SBChildParentStatusChange") %></td>				
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("OSBChildParentStatusChange") %></td>
			</tr>

			<tr>
				<td colspan="1" class="formMain">Volunteer Status Change</td>	
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("CBVolunteerStatusChange") %></td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("SBVolunteerStatusChange") %></td>				
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("OSBVolunteerStatusChange") %></td>								
			</tr>
			
			<tr>
				<td colspan="1" class="formMain">Child/Parent Dissatisfaction</td>	
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("CBChildParentDissatisfaction") %></td>				
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("SBChildParentDissatisfaction") %></td>								
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("OSBChildParentDissatisfaction") %></td>												
			</tr>
			
			<tr>

				<td colspan="1" class="formMain">Volunteer Dissatisfaction</td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("CBVolunteerDissatisfaction") %></td>								
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("SBVolunteerDissatisfaction") %></td>												
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("OSBVolunteerDissatisfaction") %></td>																
			</tr>
			
			<tr>

				<td colspan="1" class="formMain">Successful Matches</td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("CBSuccessfulMatches") %></td>								
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("SBSuccessfulMatches") %></td>												
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("OSBSuccessfulMatches") %></td>																
			</tr>			
			
			<!-- 6-Month Retention -->
			
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
								else
									if SixMonthsAgo = 0 then
										SixMonthsAgo = 12
									end if
								end if
							end if
						end if 
					end if
				end if %>
			
			<% if PrintForm = "No" then %>			
				<tr>
					<TD colspan="7" class="formHeaderMedium">6-Month Retention</TD>	
				</tr>		
			<% else %>
				<tr>
					<TD colspan="7" class="formMain"><strong><div align="center">6-Month Retention</div></strong></TD>	
				</tr>	
			<% end if %>				
			
			<TR>
				<TD colspan="1">&nbsp;</TD>
				<TD colspan="2" class="formMain" align="center">Community-Based</TD>
				<TD colspan="2" class="formMain" align="center">School-Based</TD>
				<TD colspan="2" class="formMain" align="center">Non-School<br>Site-Based</TD>				
	
			</TR>			
			
			<tr>
				<td colspan="1" class="formMain">Number of <strong>New</strong> Matches Made in <strong><%=monthname(SixMonthsAgo)%></strong></td>				
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("CBTotalOpened6MonthsAgo") %></td>								
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("SBTotalOpened6MonthsAgo") %></td>	
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("OSBTotalOpened6MonthsAgo") %></td>							
			</tr>	
			
			<tr>
				<td colspan="1" class="formMain">Number of These Matches that CLOSED before the end of <strong><%=monthname(m)%>&nbsp;<%=y%></strong></td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("CBNumberStillOpen") %></td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("SBNumberStillOpen") %></td>
				<td class="formMain" colspan="2" align="center"><%= GetPerformance("OSBNumberStillOpen") %></td>				
			</tr>
			
		<% end if %>
			
			<% if (m = 3 or m = 6 or m = 9 or m = 12) and y < 2005 then %>
			
				<% if PrintForm="No" then %>
					<tr>
						<TD colspan="7" class="formHeaderMedium">Customer Satisfaction</TD>	
					</tr>	
				<% else %>
					<tr>
						<TD colspan="7" class="formMain"><strong><div align="center">Customer Satisfaction</div></strong></TD>	
					</tr>	
				<% end if%>			
				
				
				<tr>
					<td colspan="1" class="formMain">Enrollment Satisfaction Average Score</td>
					<td class="formMain" colspan="6" align="center"><%= GetPerformance("EnrollmentSatAvgScore") %></td>				
				</tr>				
				
				<tr>
					<td colspan="1" class="formMain">Enrollment Satisfaction Count</td>
					<td class="formMain" colspan="6" align="center"><%= GetPerformance("EnrollmentSatCount") %></td>								
				</tr>
	
				<tr>
					<td colspan="1" class="formMain">Match Satisfaction Average Score</td>
					<td class="formMain" colspan="6" align="center"><%= GetPerformance("MatchSatAvgScore") %></td>								
				</tr>			
				
				<tr>
					<td colspan="1" class="formMain">Match Satisfaction Count</td>
					<td class="formMain" colspan="6" align="center"><%= GetPerformance("MatchSatCount") %></td>								
				</tr>	
				
				<% if PrintForm="No" then %>
					<tr>
						<TD colspan="7" class="formHeaderMedium">POE</TD>	
					</tr>								
				<% else %>
					<tr>
						<TD colspan="7" class="formMain"><strong><div align="center">POE</div></strong></TD>	
					</tr>		
				<% end if %>									
				
				
				<tr>
					<td colspan="1">&nbsp;</td>
					<td colspan="2" class="formMain" align="center"><b>Community-Based</b></td>
					<td colspan="2" class="formMain" align="center"><b>School-Based</b></td>
					<td colspan="2" class="formMain" align="center"><b>Non-School<br>Site-Based</b></td>				
				</tr>
				
				<tr>
					<td colspan="1" class="formMain">POE Aggregate Score</td>
					<td colspan="2" class="formMain" align="center"><%= GetPerformance("CBPOEAggregateScore") %></td>
					<td colspan="2" class="formMain" align="center"><%= GetPerformance("SBPOEAggregateScore") %></td>				
					<td colspan="2" class="formMain" align="center"><%= GetPerformance("OSBPOEAggregateScore") %></td>								
				</tr>
				
				<tr>
					<td colspan="1" class="formMain">POE Count</td>	
					<td colspan="2" class="formMain" align="center"><%= GetPerformance("CBPOECount") %></td>
					<td colspan="2" class="formMain" align="center"><%= GetPerformance("SBPOECount") %></td>								
					<td colspan="2" class="formMain" align="center"><%= GetPerformance("OSBPOECount") %></td>												
				</tr>
				
			<% else %>
			
				<% if y < 2005 then %>
					<tr>
						<td colspan="10" class="formMain" align="center"><em><strong>Customer Satisfaction and POE Questions are answered Quarterly<br>(March, June, September, and December)</strong></em></td>
					</tr>
				<% else %>
					<tr>
						<td colspan="10" class="formMain" align="center"><em><strong>Starting in 2005, POE and Customer Satisfaction questions are no longer answered using this form.  Use the online POE and Customer Satisfaction Forms found <a href="http://agencies.bbbsa.org/myagency/POESat.asp">here</a></strong></em></td>
					</tr>				
				<% end if %>
			
			<% end if %>

<% end if %>				
			


				
<!-- END SDM METRICS -->			







<% If printform = "No" Then %>

		<% if DisplayMetrics = 1 then %>
			<% if ReadOnlyLevel=0 then %>
				<tr>
					<td colspan="9" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold"></td>
				</tr>
			<% else %>
				<tr>
					<td colspan="9" class="formMain">Where did the <strong>Edit Button</strong> go?  <a href="..\helpfiles\surveyhelp.asp?HelpID=password1" onclick="NewWindow(this.href,'name','500','250','yes');return false;">Click Here</a> for an explanation.</td>
				</tr>			
			<% end if %>
				
				<tr>
					<td colspan="9"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
				</tr>				
		<% else %>
		
			<% if DisplayMetrics=1 then %>
				<% if ReadOnlyLevel=0 then %>
					<tr>
						<td colspan="7" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold"></td>
					</tr>
				<% else %>
					<tr>
						<td colspan="7" class="formMain" align="center"><strong>HEY!!! Where did the <em>Edit</em> Button</strong> go?  <a href="..\helpfiles\surveyhelp.asp?HelpID=password1" onclick="NewWindow(this.href,'name','500','250','yes');return false;">Click Here</a> for an explanation.</td>
					</tr>			
				<% end if %>	
			<% end if %>			
				<tr>
					<td colspan="7"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
				</tr>
		<% end if %>
				
				
<% End If %>			
				
				
		</table>
		
		

