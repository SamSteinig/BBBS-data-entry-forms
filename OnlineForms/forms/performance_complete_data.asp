<table width="60%" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063">

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
if (SDMPilot = 1 and ( y > 2002 or (y = 2002 and m >= 7))) or y > 2003 then 
	DisplayMetrics = 1
else
	DisplayMetrics = 0
end if

%>

<% Dim ReadOnlyLevel
If Session("ReadOnly") then
	ReadOnlyLevel=1
Else
	ReadOnlyLevel=0
End If
%>




<% if printform = "No" then %>

	<tr>
		<td colspan="7" class="formHeader">PERFORMANCE - CORE BUSINESS <br><%= MonthName(Request("m"), False) & " " & Request("y") %></td>
	</tr>
			
<% else %>			

	<tr>
		<td colspan="7" class="formIndex">PERFORMANCE - CORE BUSINESS <br><%= MonthName(Request("m"), False) & " " & Request("y") %></td>
	</tr>	
			
<% end if %>

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
			<td align="center" valign="middle" class="formMain">Non-School Site Based</td>
		<% if y < 2004 then %>
			<td align="center" valign="middle" class="formMain">Group Mentoring</td>
			<td align="center" valign="middle" class="formMain">Special Programs: Mentoring</td>
			<td align="center" valign="middle" class="formMain">Special Programs: Non-Mentoring</td>
		<% end if %>
		</tr>
		
		<!-- Matches Open/Active in the Beginning of the Month -->
		<tr>
			<td align="center" valign="middle" class="formMain">
			OPEN/ACTIVE&nbsp;matches<br>on&nbsp;the&nbsp;<strong>FIRST</strong>&nbsp;day&nbsp;of<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
			<td align="right" valign="middle" class="formMain"><%= PrevOpenComm %></td>
			<td align="right" valign="middle" class="formMain"><%= PrevOpenSchool %></td>
			<td align="right" valign="middle" class="formMain"><%= PrevOpenOther %></td>
			<% if y < 2004 then %>
				<td align="right" valign="middle" class="formMain"><%= PrevOpenGroup %></td>
				<td align="right" valign="middle" class="formMain"><%= PrevOpenSpecMent %></td>
				<td align="right" valign="middle" class="formMain"><%= PrevOpenSpecNonMent %></td>
			<% end if %>				
		</tr>		

		<!-- Matches Closed During the Month -->
		<tr>
			<td align="center" valign="middle" class="formMain">
			Matches&nbsp;CLOSED&nbsp;during<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
			<td align="right" valign="middle" class="formMain"><%= GetPerformance("ClosedMatchesCommunityBased") %></td>
			<td align="right" valign="middle" class="formMain"><%= GetPerformance("ClosedMatchesSchoolBased") %></td>
			<td align="right" valign="middle" class="formMain"><%= GetPerformance("ClosedMatchesOtherSiteBased") %></td>

			<% if y < 2004 then %>			
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("ClosedMatchesGroupMentoring") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("ClosedMatchesSpecialProgramsMentoring") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("ClosedMatchesSpecialProgramsNonMentoring") %></td>
			<% end if %>			
			
			
		</tr>		

		
		<!-- New Matches Opened During the Month -->
		<tr>	
			<td align="center" valign="middle" class="formMain">NEW&nbsp;matches opened<br>during&nbsp;<b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
			<td align="right" valign="middle" class="formMain"><%= GetPerformance("NewMatchesCommunityBased") %></td>
			<td align="right" valign="middle" class="formMain"><%= GetPerformance("NewMatchesSchoolBased") %></td>
			<td align="right" valign="middle" class="formMain"><%= GetPerformance("NewMatchesSiteBasedNonSchool") %></td>

			<% if y < 2004 then %>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("NewMatchesGroupMentoring") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("NewMatchesSpecialProgramsMentoring") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("NewMatchesSpecialProgramsNonMentoring") %></td>
			<% end if %>		
		</tr>	
		
		
		<!-- Match Transfers -->
		
		<tr>
			<td align="center" valign="middle" class="formMain">Transfer Matches</td>			
		
		<td align="right" valign="middle" class="formMain">
			<%= GetPerformance("TransferCommunityBased") %>
		</td>		
		
		<td align="right" valign="middle" class="formMain">
			<%= GetPerformance("TransferSchoolBased") %>
		</td>					
		
		<td align="right" valign="middle" class="formMain">
			<%= GetPerformance("TransferOtherSiteBased") %>
		</td>
		
		<% if y < 2004 then %>								

			<td align="right" valign="middle" class="formMain">
				<%= GetPerformance("TransferGroupMentoring") %>
			</td>		
			
			<td align="right" valign="middle" class="formMain">
				<%= GetPerformance("TransferSpecialProgramsMentoring") %>
			</td>					
			
			<td align="right" valign="middle" class="formMain">
				<%= GetPerformance("TransferSpecialProgramsNonMentoring") %>
			</td>
		<% end if %>

		</tr>					
	
		
		<tr>
			<td align="center" valign="middle" class="formMain">
			OPEN/ACTIVE&nbsp;matches<br>on&nbsp;the&nbsp;last&nbsp;day&nbsp;of<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
			<td align="right" valign="middle" class="formMain"><%= GetPerformance("OpenMatchesCommunityBased") %></td>
			<td align="right" valign="middle" class="formMain"><%= GetPerformance("OpenMatchesSchoolBased") %></td>
			<td align="right" valign="middle" class="formMain"><%= GetPerformance("OpenMatchesOtherSiteBased") %></td>
			
			<% if y < 2004 then %>			
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("OpenMatchesGroupMentoring") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("OpenMatchesSpecialProgramsMentoring") %></td>
				<td align="right" valign="middle" class="formMain"><%= GetPerformance("OpenMatchesSpecialProgramsNonMentoring") %></td>
			<% end if %>			
		</tr>		
		
		<!-- YTD - Total Matches  -->
		<tr>
			<td align="center" valign="middle" class="formMain">
			<strong>YEAR TO DATE</strong> - Total number of children served as of the end of <strong><%= MonthName(Request("m"), False) & " " & Request("y") %></strong>
			</td>
			<td align="right" valign="middle" class="formMain"><%=CommunityYTD%></td>
			<td align="right" valign="middle" class="formMain"><%=SchoolYTD%></td>			
			<td align="right" valign="middle" class="formMain"><%=OtherYTD%></td>	
			<% if y < 2004 then %>					
				<td colspan="3">&nbsp;</td>
			<% end if %>				
		
		
		</tr>
		


			<tr>
				<td align="center" valign="middle" class="formHeaderMedium" colspan="7"><strong>REVENUE</strong></td>			
			</tr>
			<tr>
				<td align="center" valign="middle" class="formMain" colspan="7"><font color="red"><strong>NOTE: </font>Revenue Questions Have Been Moved to the <BR>New Monthly Revenue / Expense Form</strong></td>
			</tr>
		

			

			

<!-- Revenue 

<% 'if (y >= 2002) and (y < 2005) then %>


<% 'if printform = "No" Then %>
			<tr>
				<td colspan="7" class="formHeaderMedium">REVENUE</td>				
			</tr>
			
<% 'else %>
			<tr>
				<td colspan="7" class="formmain"><strong><div align="center">REVENUE</div></strong></td>
			</tr>
<% 'end if %>

			<tr>
				<td valign="middle" class="formMain">Revenue&nbsp;<strong>booked&nbsp;for&nbsp;<% '= MonthName(Request("m"), False) & " " & Request("y") %></strong></td>				
				<td valign="middle" class="formMain" colspan="3"><% '= formatcurrency(GetPerformance("Revenue"))%></td>		
				<% 'if y < 2004 then %>
					<td colspan="3">&nbsp;</td>	
				<% 'end if %>
			</tr>	
			
			<tr>
				<td valign="middle" class="formMain"><strong>YEAR TO DATE</strong> Revenue as of the end of <strong><% '= MonthName(Request("m"), False) & " " & Request("y") %></strong></td>				
				<td valign="middle" class="formMain" colspan="3"><% '=formatcurrency(RevenueYTD)%></td>
				<% 'if y < 2004 then %>
					<td colspan="3">&nbsp;</td>	
				<% 'end if %>
			</tr>				
<% 'End If %>	

<!-- RTBM -->

<% if m=12 then %>

<% if printform="No" then %>
			<tr>
				<td colspan="7" class="formHeaderMedium">READY TO BE MATCHED</td>
			</tr>
<% else %>
			<tr>
				<td colspan="7" class="formMain"><strong><div align="center">READY TO BE MATCHED</div></strong></td>
			</tr>
<% end if %>
			<tr>
				<td valign="middle" class="formMain">Total number of <strong>UNMATCHED Children </strong>as of <b>12/31/<%=y%></b></td>
				<td valign="middle" class="formMain" colspan="2"><%=GetPerformance("RTBM_UnmatchedChildren")%></td>		
				<td colspan="4">&nbsp;</td>	
			</tr>

			<tr>
				<td valign="middle" class="formMain">Total number of <strong>UNMATCHED Volunteers</strong> as of <b>12/31/<%=y%></b></td>
				<td valign="middle" class="formMain" colspan="2"><%=GetPerformance("RTBM_UnmatchedVolunteers")%></td>			
				<td colspan="4">&nbsp;</td>
			</tr>			
			


<% end if%>





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
		
			<% if ReadOnlyLevel=0 then %>
				<tr>
					<td colspan="7" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold"></td>
				</tr>
			<% else %>
				<tr>
					<td colspan="7" class="formMain" align="center"><strong>HEY!!! Where did the <em>Edit</em> Button</strong> go?  <a href="..\helpfiles\surveyhelp.asp?HelpID=password1" onclick="NewWindow(this.href,'name','500','250','yes');return false;">Click Here</a> for an explanation.</td>
				</tr>			
			<% end if %>				
				<tr>
					<td colspan="7"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
				</tr>
		<% end if %>
				
				
<% End If %>			
				
				
		</table>
		
		

