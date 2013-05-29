<table border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063" width="400">
	<form name="frmSpecialPopulations" action="SpecialPopulations_edit.asp?y=<%= Request("y") %>" method="post">
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
		
		
<% if printform="No" then %>		
		<tr>
			<td colspan="4" class="formHeader">SPECIAL POPULATIONS</td>
		</tr>
<% else %>
		<tr>
			<td colspan="4" class="formIndex">SPECIAL POPULATIONS</td>
		</tr>

<%end if %>
		
		
		
		<tr>
			<td colspan="4" class="formMainBold">Created: <%= GetSpecialPopulations("CreateDate") %><br>
		<% form = "SpecialPopulations" %> 
		<% gid = GetSpecialPopulations("SpecialPopulationsID") %>
		<!--#include file="../includes/lastmodified_stamp.asp"-->
			</td>
		</tr>
<!-- Section 1 -->
		<tr>
			<td colspan="4" align="center" valign="top" class="formMain">If your agency served any of the following <b>Special Populations</b> in this<br>ADS year with a coordinated effort to reach a <b>Targeted</b> population<br>serving at least <b>six recipients</b>, then <b>indicate quantity below</b>:</td>
		</tr>
		<tr>
			<td class="formMain">Abused/Neglected:</td>
			<td class="formMainRightJ"><%= GetSpecialPopulations("AbusedNeglected") %></td>
			<td class="formMain">Institutionalized:</td>
			<td class="formMainRightJ"><%= GetSpecialPopulations("Institutionalized") %></td>
		</tr>	
		<tr>
			<td class="formMain">Adjudicated Delinquents:</td>
			<td class="formMainRightJ"><%= GetSpecialPopulations("AdjudicatedDelinquents") %></td>
			<td class="formMain">Learning Disabled:</td>
			<td class="formMainRightJ"><%= GetSpecialPopulations("LearningDisabled") %></td>
		</tr>
		<tr>
			<td class="formMain">After School (Latchkey):</td>
			<td class="formMainRightJ"><%= GetSpecialPopulations("AfterSchool") %></td>
			<td class="formMain">Physically Disabled:</td>
			<td class="formMainRightJ"><%= GetSpecialPopulations("PhysicallyDisabled") %></td>
		</tr>
		<tr>
			<td class="formMain">AIDS Affected:</td>
			<td class="formMainRightJ"><%= GetSpecialPopulations("AIDSAffected") %></td>
			<td class="formMain">Pregnant Teen:</td>
			<td class="formMainRightJ"><%= GetSpecialPopulations("PregnantTeen") %></td>
		</tr>
		<tr>
			<td class="formMain">Deaf & Hearing Impaired:</td>
			<td class="formMainRightJ"><%= GetSpecialPopulations("DeafHearingImpaired") %></td>
			<td class="formMain">School Dropouts:</td>
			<td class="formMainRightJ"><%= GetSpecialPopulations("SchoolDropouts") %></td>
		</tr>
		<tr>
			<td class="formMain">Developmentally Disabled:</td>
			<td class="formMainRightJ"><%= GetSpecialPopulations("DevelopmentallyDisabled") %></td>
			<td class="formMain">Teen Parents (Female):</td>
			<td class="formMainRightJ"><%= GetSpecialPopulations("TeenParentsFemale") %></td>
		</tr>
		<tr>
			<td class="formMain">Foster Children:</td>
			<td class="formMainRightJ"><%= GetSpecialPopulations("FosterChildren") %></td>
			<td class="formMain">Teen Parents (Male):</td>
			<td class="formMainRightJ"><%= GetSpecialPopulations("TeenParentsMale") %></td>
		</tr>
		<tr>
			<td class="formMain">Homeless:</td>
			<td class="formMainRightJ"><%= GetSpecialPopulations("Homeless") %></td>
			<td class="formMain">Visually Impaired:</td>
			<td class="formMainRightJ"><%= GetSpecialPopulations("VisuallyImpaired") %></td>
		</tr>
		<tr>
			<td class="formMain">Incarcerated Parents:</td>
			<td class="formMainRightJ"><%= GetSpecialPopulations("IncarceratedParents") %></td>
			<td class="formMain">Other: <%= GetSpecialPopulations("OtherType") %></td>
			<td class="formMainRightJ"><%= GetSpecialPopulations("Other") %></td>
		</tr>
		
<% if printform="No" then %>

	<% if ReadOnlyLevel = 0 then %>
		<tr>
			<td colspan="4" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold"></td>
		</tr>
	<% end if %>

		<tr>
			<td colspan="4"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
		</tr>
<%end if %>
	</table>

</form>