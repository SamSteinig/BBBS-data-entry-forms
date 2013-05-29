<% Dim ReadOnlyLevel
	If Session("ReadOnly") then
		ReadOnlyLevel=1
	Else
		ReadOnlyLevel=0
	End If %>


<form name="frmExpenses" action="expenses_edit.asp?y=<%= Request("y") %>" method="post">
<!--#include file="../includes/form_stamp.asp"-->
<input type="hidden" name="status" value="editOld">
<table border="1" cellspacing="0" cellpadding="2" width = "660" bordercolordark="#003063">


 

	<tr>
		<td colspan="6" align="center" class="formSubheadbold">BBBS - <font color="#ff0000"><strong><%= y %> </strong></font> Annual Agency Information (AAI)</td>
	</tr>
	
	<% if printform="No" then %>
	
		<tr>
			<td colspan="6" class="formHeader">FINANCES</td>
		</tr>
	
	<% else %>
	
		<tr>
			<td colspan="6" class="formIndex">FINANCES</td>
		</tr>
	
	<% end if %>	
	
	<tr>
		<td colspan="6" class="formMainBold">Created: <%= GetExpenses("CreateDate") %><br>
		<% form = "Expenses" %> 
		<% gid = GetExpenses("ExpensesID") %>
		<!--#include file="../includes/lastmodified_stamp.asp"-->
		</td>
	</tr>
	<% if PrintForm="No" then %>
		
		<tr>
			<td colspan="6" class="formHeaderSmall">EXPENSES</td>
		</tr>	
		
	<% else %>
	
		<tr>
			<td colspan="6" class="formMainCentered"><strong>EXPENSES</strong></td>
		</tr>				
		
	<% end if %>

	<tr>
		<td class="formMain" width="24%">Salaries and Wages:</td>
		<td class="formMainRightJ" width="24%"><%= FormatCurrency(GetExpenses("SalariesWages")) %></td>
		<td colspan="2" width="4%">&nbsp;</td>		
		<td class="formMain" width="24%">Employee Benefits:</td>
		<td class="formMainRightJ" width="24%"><%= FormatCurrency(GetExpenses("EmployeeBenefits")) %></td>	
	</tr>
	<tr>
		<td class="formMain" width="24%">Liability Insurance:</td>				
		<td class="formMainRightJ" width="24%"><%= FormatCurrency(GetExpenses("Insurance")) %></td>	
		<td colspan="2" width="4%">&nbsp;</td>					
		<td class="formMain" width="24%">All Other:</td>								
		<td class="formMainRightJ" width="24%"><%= FormatCurrency(GetExpenses("Other")) %></td>			
	</tr>
	<tr>
		<td class="formMain" width="24%">Rent / Occupancy:</td>
		<td class="formMainRightJ" width="24%"><%= FormatCurrency(GetExpenses("RentOccupancy")) %></td>		
		<td colspan="2" width="4%">&nbsp;</td>				
		<td class="formMain" width="24%" bgcolor="#c0c0c0"><strong>Total Operating Expenses</strong></td>
		<td class="formMainRightJ" width="24%" bgcolor="#c0c0c0"><%= FormatCurrency(GetExpenses("Total")) %></td>		
	</tr>




	
	<tr>
	
<% if PrintForm="No" then %>
		
		<tr>
			<td colspan="6" class="formHeaderSmall">BALANCE SHEET<br>(as of December 31, <%=Y%>)</td>
		</tr>	
		
	<% else %>
	
		<tr>
			<td colspan="6" class="formMainCentered"><strong>BALANCE SHEET<br>(as of December 31, <%=Y%>)</strong></td>
		</tr>				
		
	<% end if %>	
	
		<TR>
			<TD class="formMain" colspan="6"><strong><div align="center">Assets</div></strong></TD>
		</TR>
		<tr>
			<td class="formMain" width="80%" colspan="5">Cash/Investments:</td>
			<td class="formMain" width="20%" align="right" colspan="2"><%=FormatCurrency(GetExpenses("CashInvestments"))%></td>
		</tr>
		<tr>
			<td class="formMain" width="80%" colspan="5">Receivables:</td>
			<td class="formMain" width="20%" align="right" colspan="2"><%=FormatCurrency(GetExpenses("Receivables"))%></td>							
		</tr>		
		<tr>
			<td class="formMain" width="80%" colspan="5">All Other Assets:</td>
			<td class="formMain" width="20%" align="right" colspan="2"><%=FormatCurrency(GetExpenses("AllOtherAssets"))%></td>							
		</tr>
		<tr>
			<td class="formMain" width="80%" colspan="5"  bgcolor="#c0c0c0"><strong>TOTAL ASSETS:</strong></td>
			<td class="formMain" width="20%" align="right" colspan="2"  bgcolor="#c0c0c0"><%=FormatCurrency(GetExpenses("TotalAssets"))%></td>							
		</tr>	
		<TR>
			<TD class="formMain" colspan="6"><strong><div align="center">Liabilities and Net Assets</div></strong></TD>
		</TR>
		<tr>
			<td class="formMain" width="80%" colspan="5">Short Term Liabilities:</td>
			<td class="formMain" width="20%" align="right" colspan="2"><%=FormatCurrency(GetExpenses("LiabilitiesShort"))%></td>							
		</tr>
		<tr>
			<td class="formMain" width="80%" colspan="5">Long Term Liabilities:</td>
			<td class="formMain" width="20%" align="right" colspan="2"><%=FormatCurrency(GetExpenses("LiabilitiesLong"))%></td>							
		</tr>
		<tr>
			<td class="formMain" width="80%" colspan="5">Total Liabilities:</td>
			<td class="formMain" width="20%" align="right" colspan="2"><%=FormatCurrency(GetExpenses("Liabilities"))%></td>							
		</tr>		
		<tr>
			<td class="formMain" width="80%" colspan="5">Surplus/Net Assets:</td>
			<td class="formMain" width="20%" align="right" colspan="2"><%=FormatCurrency(GetExpenses("Surplus_NetAssets"))%></td>							
		</tr>		
		<tr>
			<td class="formMain" width="80%" colspan="5" bgcolor="#c0c0c0"><strong>TOTAL LIABILITIES AND NET ASSETS:</strong></td>
			<td class="formMain" width="20%" align="right" colspan="2" bgcolor="#c0c0c0"><%=FormatCurrency(GetExpenses("TotalLiabNetAssets"))%></td>							
		</tr>				
	
		<td colspan="6">
			<table width="100%" border="1" bordercolordark="#003063" cellspacing="0" cellpadding="2">
			
			
			
			
			<% if PrintForm="No" then %>
				
				<tr>
					<td class="formHeaderSmall" colspan="6">EXPENSE BREAKDOWN BY CATEGORY - Consistent with Audited Financial Statements<br>(Enter whole numbers only)</td>
				</tr>
			<% else %>
				<tr>
					<td class="formMainCentered" colspan="6"><strong>EXPENSE BREAKDOWN BY CATEGORY<br>(Enter whole numbers only)</strong></td>
				</tr>			
			<% end if %>
			
			<tr>
				<td class="formMain" colspan="4">Program:<br>
				<font class="formSubhead"><i>Including time spent supervising program staff</i></font></td>
				<td class="formMain" valign="top" align="right" colspan="2"><%= GetExpenses("Program") %>%</td>
			</tr>			
				
			<tr>
				<td class="formMain" colspan="4">Fundraising:</td>
				<td class="formMain" valign="top" align="right" colspan="2"><%= GetExpenses("FundRaising") %>%</td>
			</tr>		
			
			<tr>
				<td class="formMain" colspan="4">Administration:<br>
				<font class="formSubhead"><i>If any administration expenses are related to program or fundraising then include those expenses in program or fundraising when calculating percentages.</i></font></td>
				<td class="formMain" valign="top" align="right" colspan="2"><%= GetExpenses("Administration") %>%</td>
			</tr>
			
			<tr>
				<td class="formMain" colspan="4" bgcolor="#c0c0c0"><strong>TOTAL</strong></td>
				<td class="formMain" valign="top" align="right" bgcolor="#c0c0c0"><%= GetExpenses("CategoryTotal") %>%</td>
			</tr>	
<!--			
			<% if PrintForm = "No"  then %>				
				
				<tr>
					<td class="formHeaderSmall" colspan="5" align="center">EXPENSE BREAKDOWN BY FUNCTION<br>(enter whole numbers only)</td>
				</tr>
			
			<% else %>
			
				<tr>
					<td class="formMainCentered" colspan="5" align="center"><strong>EXPENSE BREAKDOWN BY FUNCTION<br>(enter whole numbers only)</strong></td>
				</tr>
				
			<% end if %>			
			

			
			<tr>
				<td class="formMain" colspan="5" align="left"><em>How much total expense goes toward (must equal 100%)</em></td>
			</tr>			
			<tr>
				<td class="formMain" width="80%" colspan="4">Mentoring</td>
				<td class="formMain" align="right" width="20%"><%=GetExpenses("TotalExpenseMentoring")%>%</td>
			</tr>
			<tr>
				<td class="formMain" width="80%" colspan="4">Non-Mentoring:</td>
				<td class="formMain" align="right" width="20%"><%=GetExpenses("TotalExpenseNonMentoring")%>%</td>	
			</tr>
			<tr>
				<td class="formMain" colspan="4" bgcolor="#c0c0c0"><strong>TOTAL</strong></td>
				<td class="formMain" valign="top" align="right" bgcolor="#c0c0c0"><%= GetExpenses("ExpensesMentNonMentTotal") %>%</td>
			</tr>				
			
			<tr>
				<td class="formMain" colspan="6" align="left"><em>Estimate the percent of Mentoring Program *FTEs (Full Time Employees) that go toward the following PROGRAMS (must equal 100%)</em></td>
			</tr>			
			<tr>
				<td class="formMain" width="80%" colspan="4">Community:</td>
				<td class="formMain" width="20%" align="right"><%= GetExpenses("FTECommunity")%>%</td>
			</tr>
			<tr>
				<td class="formMain" width="80%" colspan="4">School-Based:</td>
				<td class="formMain" width="20%" align="right"><%= GetExpenses("FTESchool")%>%</td>
			</tr>
			<tr>
				<td class="formMain" width="80%" colspan="4">Other Site-Based:</td>
				<td class="formMain" width="20%" align="right"><%= GetExpenses("FTESite")%>%</td>							
			</tr>
			<tr>
				<td class="formMain" colspan="4" bgcolor="#c0c0c0"><strong>TOTAL</strong></td>
				<td class="formMain" valign="top" align="right" bgcolor="#c0c0c0"><%= GetExpenses("FTEProgramTotal") %>%</td>
			</tr>				

			<tr>
				<td class="formMain" colspan="5" align="left"><em>Estimate the percent of Mentoring Program *FTEs (Full Time Employees) that go toward the following FUNCTIONS (must equal 100%)</em></td>
			</tr>			
			<tr>
				<td class="formMain" width="80%" colspan="4">Customer Relations:</td>
				<td class="formMain" width="20%" align="right"><%= GetExpenses("FTECustomerRelations")%>%</td>
			</tr>
			<tr>
				<td class="formMain" width="80%" colspan="4">Enrollment / Matching:</td>
				<td class="formMain" width="20%" align="right"><%= GetExpenses("FTEEnrollmentMatching")%>%</td>							
			</tr>
			<tr>
				<td class="formMain" width="80%" colspan="4">Match Support:</td>
				<td class="formMain" width="20%" align="right"><%= GetExpenses("FTEMatchSupport")%>%</td>							
			</tr>
			<tr>
				<td class="formMain" colspan="4" bgcolor="#c0c0c0"><strong>TOTAL</strong></td>
				<td class="formMain" valign="top" align="right" bgcolor="#c0c0c0"><%= GetExpenses("FTEFunctionTotal") %>%</td>
			</tr>
-->
		<%If y<2007 Then%>
			
			<% if PrintForm = "No" then %>
			
				<tr>
					<td colspan="5" class="formHeaderSmall">BENEFITS - MEDICAL<BR>% paid by BBBS for Employee and Employee's Family</td>
				</tr>
				
			<% else %>
			
				<tr>
					<td colspan="5" class="formMainCentered"><strong>BENEFITS - MEDICAL<BR>% paid by BBBS for Employee and Employee's Family</td></strong>
				</tr>		
				
			<% end if %>	
			
			
				
			<tr>
				<td class="formMain" width="20%">&nbsp;</td>
				<td class="formMain" align="center" colspan="2" width="40%">Full Time</td>
				<td class="formMain" align="center" colspan="2" width="40%">Part Time</td>				
			</tr>
			
			<tr>
				<td class="formMain" align="center" width="20%">Medical</td>

				<td class="formMain" valign="middle" align="center" width="20%">
				For Employee<br>
				<%= GetExpenses("BenMedFullEmployee")%>%
				</td>
				
				<td class="formMain" valign="middle" align="center" width="20%">
				For Family<br>
				<%= GetExpenses("BenMedFullFamily")%>%
				</td>
				
				<td class="formMain" valign="middle" align="center" width="20%">
				For Employee<br>
				<%= GetExpenses("BenMedPartEmployee") %>%
				</td>		
				
				<td class="formMain" valign="top" align="center" width="20%">
				For Family<br>
				<%= GetExpenses("BenMedPartFamily") %>%
				</td>						
			
			</tr>
			
			<tr>
				<td class="formMain" align="center">Dental</td>

				<td class="formMain" valign="top" align="center">
				For Employee<br>
				<%= GetExpenses("BenDentFullEmployee") %>%
				</td>
				
				<td class="formMain" valign="top" align="center">
				For Family<br>
				<%= GetExpenses("BenDentFullFamily") %>%
				</td>
				
				<td class="formMain" valign="top" align="center">
				For Employee<br>
				<%= GetExpenses("BenDentPartEmployee") %>%
				</td>		
				
				<td class="formMain" valign="top" align="center">
				For Family<br>
				<%= GetExpenses("BenDentPartFamily") %>%
				</td>						
			
			</tr>		

			<% if PrintForm="No" then %>
				
				<tr>
					<td colspan="6" class="formHeaderSmall">BENEFITS - NON-MEDICAL<br>(check all that apply)</td>
				</tr>	
				
			<% else %>
			
				<tr>
					<td colspan="6" class="formMainCentered"><strong>BENEFITS - NON-MEDICAL<br>(check all that apply)</strong></td>
				</tr>				
				
			<% end if %>				
			
			
			<tr>
				<td class="formMain" colspan="3">&nbsp;</td>
				<td class="formMain" align="center">Full Time</td>
				<td class="formMain" align="center">Part Time</td>

			</tr>			
			
			<tr>
				<td class="formMain" colspan="3">Disability Insurance SHORT Term</td>
				<td class="formMain" align="center"><% if GetExpenses("DisInsShortTermFull")=true then%>x<%else%>&nbsp;<% end if %></td>
				<td class="formMain" align="center"><% if GetExpenses("DisInsShortTermPart")=true then%>x<%else%>&nbsp;<% end if %></td>				
			</tr>
			
			<tr>
				<td class="formMain" colspan="3">Disability Insurance LONG Term</td>				
				<td class="formMain" align="center"><% if GetExpenses("DisInsLongTermFull")=true then%>x<% else %>&nbsp;<% end if %></td>
				<td class="formMain" align="center"><% if GetExpenses("DisInsLongTermPart")=true then%>x<% else %>&nbsp;<% end if %></td>
			</tr>
			
			<tr>
				<td class="formMain" colspan="3">EAP: Employee Assistance Programs</td>
				<td class="formMain" align="center"><% if GetExpenses("EAPFull")=true then%>x<%else%>&nbsp;<% end if %></td>				
				<td class="formMain" align="center"><% if GetExpenses("EAPPart")=true then%>x<%else%>&nbsp;<% end if %></td>								
			</tr>
			
			<tr>
				<td class="formMain" colspan="3">"Flex" Pre-Tax Plan (medical, dependent)</td>			
				<td class="formMain" align="center"><% if GetExpenses("FlexFull")=true then%>x<%else%>&nbsp;<% end if %></td>				
				<td class="formMain" align="center"><% if GetExpenses("FlexPart")=true then%>x<%else%>&nbsp;<% end if %></td>								
			</tr>			
			
			<tr>
				<td class="formMain" colspan="3">Health Club</td>			
				<td class="formMain" align="center"><% if GetExpenses("HealthClubFull")=true then%>x<%else%>&nbsp;<% end if %></td>				
				<td class="formMain" align="center"><% if GetExpenses("HealthClubPart")=true then%>x<%else%>&nbsp;<% end if %></td>								
			</tr>		
			
			<tr>
				<td class="formMain" colspan="3">Life Insurance</td>			
				<td class="formMain" align="center"><% if GetExpenses("LifeInsuranceFull")=true then%>x<%else%>&nbsp;<% end if %></td>				
				<td class="formMain" align="center"><% if GetExpenses("LifeInsurancePart")=true then%>x<%else%>&nbsp;<% end if %></td>								
			</tr>								
			
			<tr>
				<td class="formMain" colspan="3">Paid Time Off (Floating Holidays, Personal)</td>			
				<td class="formMain" align="center"><% if GetExpenses("TimeOffFull")=true then%>x<%else%>&nbsp;<% end if %></td>
				<td class="formMain" align="center"><% if GetExpenses("TimeOffPart")=true then%>x<%else%>&nbsp;<% end if %></td>
			</tr>		
			
			<tr>
				<td class="formMain" colspan="3">Paid Time Off (Sick Time)</td>			
				<td class="formMain" align="center"><% if GetExpenses("TimeOffSickFull")=true then%>x<%else%>&nbsp;<% end if %></td>				
				<td class="formMain" align="center"><% if GetExpenses("TimeOffSickPart")=true then%>x<%else%>&nbsp;<% end if %></td>								
			</tr>													
					
			<tr>
				<td class="formMain" colspan="3">Paid Time Off (Vacation)</td>			
				<td class="formMain" align="center"><% if GetExpenses("TimeOffVacFull")=true then%>x<%else%>&nbsp;<% end if %></td>				
				<td class="formMain" align="center"><% if GetExpenses("TimeOffVacPart")=true then%>x<%else%>&nbsp;<% end if %></td>								
			</tr>	
			
			<tr>
				<td class="formMain" colspan="3">Professional Dues, Conferences, etc.</td>			
				<td class="formMain" align="center"><% if GetExpenses("ProfDuesFull")=true then%>x<%else%>&nbsp;<% end if %></td>				
				<td class="formMain" align="center"><% if GetExpenses("ProfDuesPart")=true then%>x<%else%>&nbsp;<% end if %></td>								
			</tr>			
			
			<tr>
				<td class="formMain" colspan="3">Retirement</td>			
				<td class="formMain" align="center"><% if GetExpenses("RetirementFull")=true then%>x<%else%>&nbsp;<% end if %></td>				
				<td class="formMain" align="center"><% if GetExpenses("RetirementPart")=true then%>x<%else%>&nbsp;<% end if %></td>								
			</tr>						
			
			<tr>
				<td class="formMain" colspan="3">Telecommuting</td>			
				<td class="formMain" align="center"><% if GetExpenses("TelecommFull")=true then%>x<%else%>&nbsp;<% end if %></td>				
				<td class="formMain" align="center"><% if GetExpenses("TelecommPart")=true then%>x<%else%>&nbsp;<% end if %></td>
			</tr>						
			
			<tr>
				<td class="formMain" colspan="3">Tuition</td>			
				<td class="formMain" align="center"><% if GetExpenses("TuitionFull")=true then%>x<%else%>&nbsp;<% end if %></td>				
				<td class="formMain" align="center"><% if GetExpenses("TuitionPart")=true then%>x<%else%>&nbsp;<% end if %></td>								
			</tr>
		<%End If%>									
						
			</table>
		</td>
	
	</tr>
	
	
		

	<% if printform="No" then %>	
	
		<% if ReadOnlyLevel=0 then %>	
			<tr>
				<td colspan="6" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold"></td>
			</tr>
		<% else %>
			<tr>
				<td colspan="9" class="formMainCentered">Where did the <strong>Edit Button</strong> go?  <a href="..\helpfiles\surveyhelp.asp?HelpID=password1" onclick="NewWindow(this.href,'name','500','250','yes');return false;">Click Here</a> for an explanation.</td>
			</tr>							
		<% end if %>
		<tr>
			<td colspan="6"><!--#include file="../includes/contact_info.inc"--></td>
		</tr>
		
	<% end if %>
</table>

</form>