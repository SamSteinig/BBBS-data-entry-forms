<table width="50%" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063">

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

<% Dim ReadOnlyLevel
If Session("ReadOnly") then
	ReadOnlyLevel=1
Else
	ReadOnlyLevel=0
End If
%>




<% if printform = "No" then %>

	<tr>
		<td colspan="7" class="formHeader">OTHER PERFORMANCE - <%= MonthName(Request("m"), False) & " " & Request("y") %></td>
	</tr>
			
<% else %>			

	<tr>
		<td colspan="7" class="formIndex">OTHER PERFORMANCE - <%= MonthName(Request("m"), False) & " " & Request("y") %></td>
	</tr>	
			
<% end if %>

		<tr>
		<!-- Date / Time Info -->
		
			<td colspan="7" class="formMainBold">Created: <%= GetPerformance("CreateDate") %><br>		
			
			<% form = "Performance" %> 
			<% gid = GetPerformance("PerformanceID") %>
			<%= GetPerformance("PerformanceID") %>
			<!--#include file="../includes/lastmodified_stamp.asp"-->
			</td>
		</tr>
			

			

<%
dim FBIEdit
FBIEdit = 0			
			
If FBIAgency = 1 and ((y > 2002 and m < 7) or (y = 2002 and m > 8) ) Then 
FBIEdit = 1
%>			
			
<!-- Fields added for Faith-Based / Incarcerated Questions -->

<% if printform = "No" then %>

	<tr>
		<td colspan="7" class="formHeaderMedium">FAITH-BASED / INCARCERATED PARENTS REPORTING</td>				
	</tr>

			
<% else %>

	<tr>
		<td colspan="7" class="formMain"><strong><div align="center">FAITH-BASED / INCARCERATED PARENTS REPORTING</div></strong></td>				
	</tr>

			
<% end if %>

			
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
				<td align="left" colspan="3">&nbsp;</td>
			</tr>
			
			<tr>
				<td align="center" class="formMain">Faith-Based Partnerships <strong>Only</strong></td>
				<td align="right" class="formMain"><%= GetPerformance("CBFBnotI") %></td>
				<td align="right" class="formMain"><%= GetPerformance("SBFBnotI") %></td>
				<td align="right" class="formMain"><%= GetPerformance("OSBFBnotI") %></td>
				<td align="left" colspan="3">&nbsp;</td>
			</tr>			
			
<% End If %>			

<!-- Fields added for SBM Questions for June and December only -->
			
<% 	dim SBMEdit
	SBMEdit = 0
if (m=6 or m=12) and SBMAgency = 1 and y <> 2001 then	
	SBMEdit = 1 	
%>	
	
	<% if printform = "No" then %>
		<tr>
			<td colspan="7" class="formHeaderMedium">SCHOOL-BASED MENTORING GRANT PROGRESS REPORT</td>
		</tr>			
	<% else %>
		<tr>
			<td colspan="7" class="formmain"><div align="center"><strong>SCHOOL-BASED MENTORING GRANT PROGRESS REPORT</strong></div></td>
		</tr>
	<% end if %>

			
			<tr>
				<td colspan="6" class="formMain">Number of Volunteers Currently in the Enrollment Process</td>
				<td colspan="1" align="right" valign="middle" class="formMain"><%= GetPerformance("SBMVolunteersInEnrollmentProcess") %></td>
			</tr>					
					
			<tr>
				<td colspan="6" class="formMain">Amount Raised Towards Match Pledge as of the last day of &nbsp<b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				<td colspan="1" align="right" valign="middle" class="formMain"><%= GetPerformance("SBMAmountRaisedTowardsMatchPledge") %></td>				
			</tr>				
	<% end if %>



<!-- Partnership Questionnaire -->

<%
dim PartnershipEdit
PartnershipEdit = 0

if (y = 2003 and (m=4 or m=11)) or (y > 2003 and (m=5 or m=11)) then 
PartnershipEdit = 1
%>

<tr>

	<td align="center" colspan="7" class="formmain">&nbsp;</td>

</tr>

<tr>
	<td colspan="7" class="formHeader">PARTNERSHIP QUESTIONNAIRE</td>
</tr>

<!-- Active Matches -->
<tr>
	<td align="center" colspan="7" class="formmain">The number of <strong>Active Matches</strong> with the following organizations:</td>
</tr>

<tr>
	<td>&nbsp;</td>
	<td align="center" class="formMain">Community<br>Based</td>
	<td align="center" class="formMain">School<br>Based</td>	
	<td align="center" class="formMain">Other<br>Site Based</td>	
	<td align="center" class="formMain">Not<br>Partnering <em>(indicated with an 'x')</em></td>	
	<td align="center" colspan="2" class="formMain"><em>If Not Partnering,</em> interested <br>in forming a partnership?</td>

</tr>

<!-- Alpha Phi Alpha -->
<tr>
	<td align="left" class="formmain">Alpha Phi Alpha</td>
	
	<!-- Alpha Community Based -->
	<td align="center" class="formmain">
		<%= GetPerformance("AlphaCommunityBased")%>
	</td>
	
	<!-- Alpha School Based -->
	<td align="center" class="formmain">
		<%= GetPerformance("AlphaSchoolBased") %>
	</td>	
	
	<!-- Alpha Other Site Based -->
	<td align="center" class="formmain">
		<%= GetPerformance("AlphaOtherSiteBased") %>
	</td>		
	
	<!-- Alpha Not Partnering -->
	<td align="center" class="formmain">
		<% if Trim(GetPerformance("AlphaNotPartnering"))="1" then %>x<%else%>&nbsp;<% End If %>
	</td>	
	
	<!-- Alpha Interest -->
	<td align="center" colspan="2" class="formmain">
		<% if Trim(GetPerformance("AlphaNotPartnering"))="1" then %>
		
			<% if Trim(GetPerformance("AlphaInterest")) = "1" then %>Yes<%else%>No<% End If %>
		
		<%else%>n/a<%end if%>
	</td>
	

	
</tr>

<!-- Lions Club -->
<tr>
	<td align="left" class="formmain">Lions Club</td>	
	
	<!-- Lions Community Based -->
	<td align="center" class="formmain">
		<%= GetPerformance("LionsCommunityBased") %>
	</td>
	
	<!-- Lions School Based -->
	<td align="center" class="formmain">
		<%= GetPerformance("LionsSchoolBased") %>
	</td>
	
	<!-- Lions Other Site Based -->
	<td align="center" class="formmain">
		<%= GetPerformance("LionsOtherSiteBased") %>
	</td>	
	
	<!-- Lions Not Partnering -->
	<td align="center" class="formmain">
		<% if Trim(GetPerformance("LionsNotPartnering"))="1" then %>x<%else%>&nbsp;<% End If %>
	</td>

	<!-- Lions Interest -->
	<td align="center" colspan="2" class="formmain">
		<% if Trim(GetPerformance("LionsNotPartnering"))="1" then %>
		
			<% if Trim(GetPerformance("LionsInterest")) = "1" then %>Yes<%else%>No<% End If %>
		
		<%else%>n/a<%end if%>
	</td>	
			
</tr>

<!-- Rotary Club -->
<tr>
	<td align="left" class="formmain">Rotary Club</td>	
	
	<!-- Rotary Community Based -->
	<td align="center" class="formmain">
		<%= GetPerformance("rotaryCommunityBased") %>
	</td>
	
	<!-- Rotary School Based -->
	<td align="center" class="formmain">
		<%= GetPerformance("RotarySchoolBased") %>
	</td>
	
	<!-- Rotary Other Site Based -->
	<td align="center" class="formmain">
		<%= GetPerformance("RotaryOtherSiteBased") %>
	</td>	
	
	<!-- Rotary Not Partnering -->
	<td align="center" class="formmain">
		<% if Trim(GetPerformance("RotaryNotPartnering"))="1" then %>x<%else%>&nbsp;<% End If %>
	</td>	
	
	<!-- Rotary Interest -->
	<td align="center" colspan="2" class="formmain">
		<% if Trim(GetPerformance("RotaryNotPartnering"))="1" then %>
		
			<% if Trim(GetPerformance("RotaryInterest")) = "1" then %>Yes<%else%>No<% End If %>
		
		<%else%>n/a<%end if%>
	</td>		
	
</tr>

<!-- Kiwanis Club -->
<tr>
	<td align="left" class="formmain">Kiwanis Club</td>	
	
	<!-- Kiwanis Community Based -->
	<td align="center" class="formmain">
		<%= GetPerformance("KiwanisCommunityBased") %>
	</td>
	
	<!-- Kiwanis School Based -->
	<td align="center" class="formmain">
		<%= GetPerformance("KiwanisSchoolBased") %>
	</td>
	
	<!-- Kiwanis Other Site Based -->
	<td align="center" class="formmain">
		<%= GetPerformance("KiwanisOtherSiteBased") %>
	</td>	
	
	<!-- Kiwanis Not Partnering -->
	<td align="center" class="formmain">
		<% if Trim(GetPerformance("KiwanisNotPartnering"))="1" then %>x<%else%>&nbsp;<% End If %>
	</td>		
	
	<!-- Kiwanis Interest -->
	<td align="center" colspan="2" class="formmain">
		<% if Trim(GetPerformance("KiwanisNotPartnering"))="1" then %>
		
			<% if Trim(GetPerformance("KiwanisInterest")) = "1" then %>Yes<%else%>No<% End If %>
		
		<%else%>n/a<%end if%>
	</td>		
	
</tr>


<!-- Optimist Club -->
<tr>
	<td align="left" class="formmain">Optimist Club</td>	
	
	<!-- Optimist Community Based -->
	<td align="center" class="formmain">
		<%= GetPerformance("OptimistCommunityBased") %>
	</td>
	
	<!-- Optimist School Based -->
	<td align="center" class="formmain">
		<%= GetPerformance("OptimistSchoolBased") %>
	</td>
	
	<!-- Optimist Other Site Based -->
	<td align="center" class="formmain">
		<%= GetPerformance("OptimistOtherSiteBased") %>
	</td>	
	
	<!-- Optimist Not Partnering -->
	<td align="center" class="formmain">
		<% if Trim(GetPerformance("OptimistNotPartnering"))="1" then %>x<%else%>&nbsp;<% End If %>
	</td>		
	
	<!-- Optimist Interest -->
	<td align="center" colspan="2" class="formmain">
		<% if Trim(GetPerformance("OptimistNotPartnering"))="1" then %>
		
			<% if Trim(GetPerformance("OptimistInterest")) = "1" then %>Yes<%else%>No<% End If %>
		
		<%else%>n/a<%end if%>
	</td>		
	
</tr>


<!-- AARP Club -->
<tr>
	<td align="left" class="formmain">AARP</td>	
	
	<!-- AARP Community Based -->
	<td align="center" class="formmain">
		<%= GetPerformance("AARPCommunityBased") %>
	</td>
	
	<!-- AARP School Based -->
	<td align="center" class="formmain">
		<%= GetPerformance("AARPSchoolBased") %>
	</td>
	
	<!-- AARP Other Site Based -->
	<td align="center" class="formmain">
		<%= GetPerformance("AARPOtherSiteBased") %>
	</td>	
	
	<!-- AARP Not Partnering -->
	<td align="center" class="formmain">
		<% if Trim(GetPerformance("AARPNotPartnering"))="1" then %>x<%else%>&nbsp;<% End If %>
	</td>		
	
	<!-- AARP Interest -->
	<td align="center" colspan="2" class="formmain">
		<% if Trim(GetPerformance("AARPNotPartnering"))="1" then %>
		
			<% if Trim(GetPerformance("AARPInterest")) = "1" then %>Yes<%else%>No<% End If %>
		
		<%else%>n/a<%end if%>

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

	<td class="formMain" colspan="5" align="left">	

	<% if Trim(GetPerformance("AlphaRating")) = "0" then %>Not Applicable<%else%><%=GetPerformance("AlphaRating")%><%end if%>

	</td>
</tr>

<!-- Lions Club Rating -->
<tr>
	<td class="formMain" colspan="2">Lions Club</td>
	<td class="formMain" colspan="5" align="left">	
	<% if Trim(GetPerformance("LionsRating")) = "0" then %>Not Applicable<%else%><%=GetPerformance("LionsRating")%><%end if%>
	</td>	
</tr>

<!-- Rotary Club Rating -->
<tr>
	<td class="formMain" colspan="2">Rotary Club</td>
	<td class="formMain" colspan="5" align="left">	
	<% if Trim(GetPerformance("RotaryRating")) = "0" then %>Not Applicable<%else%><%=GetPerformance("RotaryRating")%><%end if%>
	</td>	
</tr>

<!-- Kiwanis Club Rating -->
<tr>
	<td class="formMain" colspan="2">Kiwanis Club</td>
	<td class="formMain" colspan="5" align="left">	
	<% if Trim(GetPerformance("KiwanisRating")) = "0" then %>Not Applicable<%else%><%=GetPerformance("KiwanisRating")%><%end if%>
	</td>	
</tr>

<!-- Optimist Club Rating -->
<tr>
	<td class="formMain" colspan="2">Optimist Club</td>
	<td class="formMain" colspan="5" align="left">	
	<% if Trim(GetPerformance("OptimistRating")) = "0" then %>Not Applicable<%else%><%=GetPerformance("OptimistRating")%><%end if%>
	</td>	
</tr>

<!-- AARP Rating -->
<tr>
	<td class="formMain" colspan="2">AARP</td>
	<td class="formMain" colspan="5" align="left">	
	<% if Trim(GetPerformance("AARPRating")) = "0" then %>Not Applicable<%else%><%=GetPerformance("AARPRating")%><%end if%>
	</td>	
</tr>

<!-- Alpha Phi Alpha Partnership -->
<tr>
	<td colspan="7" class="formHeaderMedium">ALPHA PHI ALPHA PARTNERSHIP</td>	
</tr>

<tr>
	<td colspan="7" align="center" class="formMain">I am partnering with the Alphas in the following ways:</td>
</tr>

<tr>
	<td colspan="7" align="left" class="formMain">
	<% if Trim(GetPerformance("AlphaFunding"))="1" then %><b>Funding:</b> Alpha chapter supports BBBS funding efforts<BR><% end if %>
	<% if Trim(GetPerformance("AlphaProgramInitiative"))="1" then %><b>Program Initiative:</b> Chapter has activities with children on waiting list<br><% End If %>
	<% if Trim(GetPerformance("AlphaLeadershipInvolvement"))="1" then %><b>Leadership Involvement:</b> Alpha serves on board, provides agency with professional skills and resources (serves as volunteer)<% End If %>
	</td>
</tr>


<!-- Chapter Locations -->

<tr>
	<td colspan="7" align="center" class="formMain">Name and location of your local Alpha Phi Alpha Chapter(s):</td>
</tr>

<tr>
	<td align="left" class="formMain">Undergraduate Chapter</td>
	<td align="left" colspan="6" class="formMain">
	Name:&nbsp;<%= GetPerformance("AlphaUndergradChapterName") %><br>
	City:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= GetPerformance("AlphaUndergradChapterCity") %><br>	
	State:&nbsp;&nbsp;<%= GetPerformance("AlphaUndergradChapterState")%>
	</td>	
</tr>

<tr>
	<td align="left" class="formMain">Alumni Chapter</td>
	<td align="left" colspan="6" class="formMain">
	Name:&nbsp;<%= GetPerformance("AlphaAlumniChapterName") %><br>
	City:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= GetPerformance("AlphaAlumniChapterCity") %><br>	
	State:&nbsp;&nbsp;<%= GetPerformance("AlphaAlumniChapterState")%>
	</td>	
</tr>

<!-- End Partnership Questionnaire -->


<% end if %>

<% If printform = "No" Then %>
		<% if FBIEdit = 0 and SBMEdit = 0 and PartnershipEdit = 0 then %>
			<tr>
				<td colspan="7" class="formMainBold"><div align="center">No "Other" Reports are Required This Month.<br><a href="monthly.asp">Click Here</a> to return.</div></td>
			</tr>
		<% end if %>

		<% if ReadOnlyLevel=0 then %>
			<% if FBIEdit <> 0 or SBMEdit <> 0 or PartnershipEdit <> 0 then %>			
				<tr>
					<td colspan="7" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold"></td>
				</tr>
			<% end if %>
		<% else %>
			<% if FBIEdit <> 0 or SBMEdit <> 0 or PartnershipEdit <> 0 then %>					
				<tr>
					<td colspan="7" class="formMain" align="center"><strong>HEY!!! Where did the <em>Edit</em> Button</strong> go?  <a href="..\helpfiles\surveyhelp.asp?HelpID=password1" onclick="NewWindow(this.href,'name','500','250','yes');return false;">Click Here</a> for an explanation.</td>
				</tr>			
			<% end if %>
		<% end if %>				
			<tr>
				<td colspan="7"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
			</tr>

				
				
<% End If %>			
				
				
		</table>
		
		

