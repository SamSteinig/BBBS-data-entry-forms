<% Dim ReadOnlyLevel
	If Session("ReadOnly") then
		ReadOnlyLevel=1
	Else
		ReadOnlyLevel=0
	End If %>
	
<!-- RESULTS TABLE STARTS HERE -->


<form name="frmExpenses" action="BudgetForecast_edit.asp?y=<%= Request("y") %>" method="post" ID="Form1">
<!--#include file="../includes/form_stamp.asp"-->
<input type="hidden" name="status" value="editOld" ID="Hidden1">
<table border="1" cellspacing="0" cellpadding="2" width = "660" bordercolordark="#003063" ID="Table2">


 

	<tr>
		<td colspan="6" align="center" class="formSubhead">BBBS - <%= y %> Annual Agency Information (AAI)</td>
	</tr>
	
	<% if printform="No" then %>
	
		<tr>
			<td colspan="6" class="formHeader">Staffing Expense Summary</td>
		</tr>
	
	<% else %>
	
		<tr>
			<td colspan="6" class="formIndex">Staffing Expense Summary</td>
		</tr>
	
	<% end if %>	
	
	<tr>
		<td colspan="6" class="formMainBold">Created: <%= GetBudget("CreateDate") %><br>
		<% form = "Budget" %> 
		<% gid = GetBudget("BudgetForecastID") %>
		<!--#include file="../includes/lastmodified_stamp.asp"-->
		</td>
	</tr>


	
	<tr>
	
		<td colspan="6">
			<table width="100%" border="1" bordercolordark="#003063" cellspacing="0" cellpadding="2" ID="Table3">	
			


	<!-- first row of table headers -->

			

				<tr>
				    <td align="left" valign="top" class="formMain">1.</td>
					<td colspan="4" width="85%"class="formMain">Overall Projected salary increase forecasted for <%= y+1 %> budget year:</td>
					<td class="formMain" valign="top" align="center"><% If isNull(GetBudget("TotalBudgetPrcnt")) = false Then %><%= GetBudget("TotalBudgetPrcnt")%><% Else %>N/A<% End If %>%</td>
				</tr>
				<tr>
				    <td align="left" valign="top" class="formMain">2.</td> 
					<td colspan="4" width="85%"class="formMain">Overall percent increase to employee Medical, Vision, Dental premiums forecasted for <%= y+1 %> budget year:</td>
					<td class="formMain" valign="top" align="center"><% If isNull(GetBudget("BenefitsBudgetPrcnt")) = false Then %><%= GetBudget("BenefitsBudgetPrcnt")%><% Else %>N/A<% End If %>%</td>
				</tr>
				<tr>
				    <td align="left" valign="top" class="formMain">3.</td>
					<td colspan="4" width="85%"class="formMain">Overall percent increase to employee MERIT cost forecasted for <%= y+1 %> budget year:</td>
					<td class="formMain" valign="top" align="center"><% If isNull(GetBudget("MeritIncreasePrcnt")) = false Then %><%= GetBudget("MeritIncreasePrcnt")%><% Else %>N/A<% End If %>%</td>
				</tr>
				<tr>
				    <td align="left" valign="top" class="formMain">4.</td>
					<td colspan="4" width="85%"class="formMain">Average number of hours reduced by exempt employees </td>
					<td class="formMain" valign="top" align="center"><% If isNull(GetBudget("ExemptReduced")) = false Then %><%= GetBudget("ExemptReduced")%><% Else %>N/A<% End If %></td>
				</tr>
                <tr>
                    <td align="left" valign="top" class="formMain">5.</td>
					<td colspan="4" width="85%"class="formMain">Average number of hours reduced by non-exempt employees </td>
					<td class="formMain" valign="top" align="center"><% If isNull(GetBudget("NonExemptReduced")) = false Then %><%= GetBudget("NonExemptReduced")%><% Else %>N/A<% End If %></td>
				</tr> 
				<tr>
				    <td align="left" valign="top" class="formMain">6.</td>
					<td colspan="4" width="85%"class="formMain">Number of employees laid off since July 1, 2008</td>
					<td class="formMain" valign="top" align="center"><% If isNull(GetBudget("Laidoff")) = false Then %><%= GetBudget("Laidoff")%><% Else %>N/A<% End If %></td>
				</tr>
				<tr> 
			        <td align="left" valign="top" class="formMain">7.</td>
			        <td colspan="3" align="left" valign="top" class="formMain">Has your agency had an across the board salary reduction program.
						<% If (GetBudget("BoardSalaryReduction") = True) Then %> <strong>Yes</strong><%else%> <strong>No</strong><% End If %><br>
						<% If (GetBudget("BoardSalaryReduction") = True) Then %> 
						<% end if %>
					</td>
				<tr>
				    <td align="left" valign="top" class="formMain">8.</td>
					<td colspan="4" width="85%"class="formMain">Minimum number of hours an employee is considered full time at your agency?</td>
					<td class="formMain" valign="top" align="center"><% If isNull(GetBudget("MinHoursFullTime")) = false Then %><%= GetBudget("MinHoursFullTime")%><% Else %>N/A<% End If %></td>
				</tr>	
					
					
				</tr>
	<% if printform="No" then %>	
	
		<% if ReadOnlyLevel=0 then %>	
			<tr>
				<td colspan="6" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold" ID="Submit1" NAME="Submit1"></td>
			</tr>
		<% else %>
			<tr>
				<td colspan="6" class="formMainCentered">Where did the <strong>Edit Button</strong> go?  <a href="..\helpfiles\surveyhelp.asp?HelpID=password1" onclick="NewWindow(this.href,'name','500','250','yes');return false;">Click Here</a> for an explanation.</td>
			</tr>							
		<% end if %>
		<tr>
			<td colspan="6" align="center"><!--#include file="../includes/contact_info.inc"--></td>
		</tr>
		
	<% end if %>

</table>
</form>

			
					 
		