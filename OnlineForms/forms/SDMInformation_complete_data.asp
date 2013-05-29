
<table border=0>
<tr><td align="center"><form name="frmSDMInformation" action="SDMInformation_edit.asp?y=<%= Request("y") %>" method="post">
<!--#include file="../includes/form_stamp.asp"-->
<input type="hidden" name="status" value="editOld">
	<P>
			<table width="600" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063">
				<tr> 
					<td colspan="3" align="center" valign="top" class="formSubhead">BBBS - <%= y %> Annual Agency Information (AAI)</td>
			</tr>
				
<% if printform = "No" then %>				
				
				<tr>
					<td colspan="3" class="formHeader">SDM INFORMATION</td>
				</tr>
				
<% else %>

				<tr>
					<td colspan="3" class="formIndex">SDM INFORMATION</td>
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
					<td colspan="3" class="formMainBold">Created: <%= GetSDMInformation("CreateDate") %><br>
					<% form = "SDMInformation" %> 
					<% gid = GetSDMInformation("SDMInformationID") %>
					<!--#include file="../includes/lastmodified_stamp.asp"-->
					</td>
				</tr>


				
<!-- Question Number 1 -->
				<tr>
					<td align="right" valign="top" class="formMain">1.</td>
					<td align="left" valign="top" class="formMain">What is the total number of volunteer inquiries you received in <%=y%>?<br><span class="formSubHead">A volunteer is considered to have inquired when he/she contacts the agency, expresses an interest in being a Big and provides basic contact information.  Contact includes web-based inquiries.</span></td>
					<td align="right" valign="top" class="formMain"><%=  GetSDMInformation("VolunteerInquiries") %></td>
				</tr>
				
<!-- Question Number 2 -->
				<tr>
					<td align="right" valign="top" class="formMain">2.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Volunteer In Person interviews in <%=y%>?</td>
					<td align="right" valign="top" class="formMain"><%= GetSDMInformation("VolunteerInPersonInterviews") %></td>
				</tr>
<!-- Question Number 3 -->
				<tr> 
					<td align="right" valign="top" class="formMain">3.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Volunteers that were matched in <%=y%>?</td>
					<td align="right" valign="top" class="formMain"><%= GetSDMInformation("VolunteersMatched") %></td>
				</tr>

<!-- Question Number 4 -->
				<tr>
					<td align="right" valign="top" class="formMain">4.</td>
					<td align="left" valign="top" class="formMain">Volunteer Rematch Rate<br>
					<span class="formSubHead">
					If you collect this data, please report it.<hr>
					Volunteer Rematch Rate is calculated as follows: <br><br><em>&nbsp;&nbsp;&nbsp;&nbsp;Number of Volunteers Rematched with a New Child in <%=y%><br>&nbsp;&nbsp;&nbsp;&nbsp;DIVIDED BY<br>&nbsp;&nbsp;&nbsp;&nbsp;the Total Number of Closed Matches in <%=y%>.</em>
					</span>
					</td>
					<td align="right" valign="top" class="formMain"><%= GetSDMInformation("VolunteerRematchRate") %></td>
				</tr>
				
				<tr>
					<td colspan="7" class="formHeaderSmall">YOUTH</td>
				</tr>				
				

<!-- Question Number 5 -->
				<tr>
					<td align="right" valign="top" class="formMain">5.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Youth Inquiries you received in <%=y%>?<br>
					<span class="formSubHead">
					A Youth is considered to have inquired when his/her parent or guardian contacts the agency, expresses an interest in getting a Big and provides basic contact information.
					</span>
					</td>
					<td align="right" valign="top" class="formMain"><%= GetSDMInformation("YouthInquiries") %></td>
				</tr>				
				
				
<!-- Question Number 6 -->
				<tr>
					<td align="right" valign="top" class="formMain">6.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Youth In Person Interviews in <%=y%>?</td>
					<td align="right" valign="top" class="formMain"><%= GetSDMInformation("YouthInPersonInterviews") %></td>
				</tr>
<!-- Question Number 7 -->
				<tr>
					<td align="right" valign="top" class="formMain">7.</td>
					<td align="left" valign="top" class="formMain">What is the total number of Youth who were matched in <%=y%>?</td>
					<td align="right" valign="top" class="formMain"><%= GetSDMInformation("YouthsMatched") %></td>
				</tr>

				

<% if printform = "No" then %>	

	<% if ReadOnlyLevel = 0 then %>					
		<tr>
			<td colspan="3" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold"></td>
		</tr>
		<tr>
		
	<% else %>
		<tr>
			<td colspan="9" class="formMainCentered">Where did the <strong>Edit Button</strong> go?  <a href="..\helpfiles\surveyhelp.asp?HelpID=password1" onclick="NewWindow(this.href,'name','500','250','yes');return false;">Click Here</a> for an explanation.</td>
		</tr>	
		
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