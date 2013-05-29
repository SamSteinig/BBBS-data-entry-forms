<table width="550" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063">

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
		<td colspan="6" class="formHeader">Monthly Performance<br>REVENUE / EXPENSE<br><%= MonthName(Request("m"), False) & " " & Request("y") %></td>
	</tr>
			
<% else %>			

	<tr>
		<td colspan="6" class="formIndex">Monthly Performance<br>REVENUE / EXPENSE <br><%= MonthName(Request("m"), False) & " " & Request("y") %></td>
	</tr>	
			
<% end if %>

		<tr>
		
			<td colspan="6" class="formMainBold">Created: <%= GetPerformance("CreateDate") %><br>		
			
			<% form = "FinancePerformance" %> 
			<% gid = GetPerformance("FinancePerformanceID") %>
			<%= GetPerformance("FinancePerformanceID") %>
			<!--#include file="../includes/lastmodified_stamp.asp"-->
			</td>
		</tr>
		
		
		
<!-- Finance Questions -->

			<tr>
				<td valign="middle" align="center" class="formHeaderMedium" colspan="6">REVENUE</td>
			</tr>

			<tr>
				<td valign="middle" class="formMain" colspan="5">United Way</td>
				<td valign="middle" class="formMain" ><%=formatcurrency(GetPerformance("UnitedWay"))%></td>			
			</tr>	
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">Government - Federal Funding</td>
				<td valign="middle" class="formMain"><%=formatcurrency(GetPerformance("GovFederalFunding"))%></td>				
			</tr>	
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">Government - State Funding</td>
				<td valign="middle" class="formMain"><%=formatcurrency(GetPerformance("GovStateFunding"))%></td>				
			</tr>				

			<tr>
				<td valign="middle" class="formMain" colspan="5">Government - Local Funding</td>
				<td valign="middle" class="formMain"><%=formatcurrency(GetPerformance("GovLocalFunding"))%></td>				
			</tr>	
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">Foundations - Grants</td>
				<td valign="middle" class="formMain"><%=formatcurrency(GetPerformance("FoundationGrants"))%></td>				
			</tr>				
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">Corporations - Non-event Donations</td>
				<td valign="middle" class="formMain"><%=formatcurrency(GetPerformance("CorporateGifts"))%></td>				
			</tr>				
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">BBBSA (Pass-Through) Grants</td>
				<td valign="middle" class="formMain"><%=formatcurrency(GetPerformance("BBBSAGrants"))%></td>				
			</tr>		
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">Individual Giving (Non-Event)</td>
				<td valign="middle" class="formMain"><%=formatcurrency(GetPerformance("IndividualGiving"))%></td>				
			</tr>	
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">Events</td>
				<td valign="middle" class="formMain"><%=formatcurrency(GetPerformance("EventsTotal"))%></td>				
			</tr>
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">&nbsp;&nbsp;&nbsp;&nbsp;of Total Events, Portion From Individuals</td>
				<td valign="middle" class="formMain"><%=formatcurrency(GetPerformance("EventsIndiv"))%></td>				
			</tr>				
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">&nbsp;&nbsp;&nbsp;&nbsp;of Total Events, Portion From Corporations</td>
				<td valign="middle" class="formMain"><%=formatcurrency(GetPerformance("EventsCorp"))%></td>				
			</tr>							
			
													
							
			<tr>
				<td valign="middle" class="formMain" colspan="5">Dividends and Interest</td>
				<td valign="middle" class="formMain"><%=formatcurrency(GetPerformance("DividendsInterest"))%></td>				
			</tr>
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">Other</td>
				<td valign="middle" class="formMain"><%=formatcurrency(GetPerformance("Other"))%></td>				
			</tr>		
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">Total Gross Revenue</td>
				<td valign="middle" class="formMain"><%=formatcurrency(GetPerformance("TotalGross"))%></td>				
			</tr>		
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">Total Direct Expenses From Special Event Fundraising</td>
				<td valign="middle" class="formMain"><%=formatcurrency(GetPerformance("SpecEventExp"))%></td>				
			</tr>	
			
			<tr>
				<td valign="middle" class="formMain" colspan="5"><strong>TOTAL NET REVENUE</strong></td>
				<td valign="middle" class="formMain"><b><%=formatcurrency(GetPerformance("Total"))%></b></td>				
			</tr>										

			<tr>
				<td valign="middle" class="formMain" colspan="5">&nbsp;&nbsp;&nbsp;&nbsp;Of Total, Net Amount Raised Through BFKS</td>
				<td valign="middle" class="formMain"><%=formatcurrency(GetPerformance("TotalAmountBFKS"))%></td>				
			</tr>		
			
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">&nbsp;&nbsp;&nbsp;&nbsp;Of Total, Net Amount Raised Through RMM</td>
				<td valign="middle" class="formMain"><%=formatcurrency(GetPerformance("TotalAmountRMM"))%></td>				
			</tr>					
			

<%''''' Balance Sheet section added 3/23/2009 saf
	Dim iMonth
	If Request("m") <> "" Then
		iMonth = CINT(Request("m"))
	Else
	    iMonth = CINT(Request("month"))
	End If

	Select case iMonth
		Case 12  ''''' Annual Expenses section added 4/4/2009 saf
			Set Con = Server.CreateObject("ADODB.Connection")
			Con.Open "BBBSAforms", "sa","12sist12"
			query = "SELECT * FROM tbl_frmExpenses WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
			Set GetExpenses = Con.Execute(query)
			
			Dim SalariesWages, EmployeeBenefits, Insurance, Other, RentOccupancy, Total, Program, FundRaising
			Dim Administration, CategoryTotal, TotalOperatingExpense
			
			If not GetExpenses.EOF Then
				If ISNULL(GetExpenses("SalariesWages")) then 
					SalariesWages = 0
				Else
					SalariesWages = GetExpenses("SalariesWages")
				End If
				
				If ISNULL(GetExpenses("EmployeeBenefits")) then 
					EmployeeBenefits = 0
				Else
					EmployeeBenefits = GetExpenses("EmployeeBenefits")
				End If
				
				If ISNULL(GetExpenses("Insurance")) then 
					Insurance = 0
				Else
					Insurance = GetExpenses("Insurance")
				End If
				
				If ISNULL(GetExpenses("Other")) then 
					Other = 0
				Else
					Other = GetExpenses("Other")
				End If
				
				If ISNULL(GetExpenses("RentOccupancy")) then 
					RentOccupancy = 0
				Else
					RentOccupancy = GetExpenses("RentOccupancy")
				End If
				
				If ISNULL(GetExpenses("Total")) then 
					Total = 0
				Else
					Total = GetExpenses("Total")
				End If
				
				If ISNULL(GetExpenses("Program")) then 
					Program = 0
				Else
					Program = GetExpenses("Program")
				End If
				
				If ISNULL(GetExpenses("FundRaising")) then 
					FundRaising = 0
				Else
					FundRaising = GetExpenses("FundRaising")
				End If
				
				If ISNULL(GetExpenses("Administration")) then 
					Administration = 0
				Else
					Administration = GetExpenses("Administration")
				End If
				
				If ISNULL(GetExpenses("CategoryTotal")) then 
					CategoryTotal = 0
				Else
					CategoryTotal = GetExpenses("CategoryTotal")
				End If
				
				If ISNULL(GetPerformance("TotalOperatingExpense")) then 
					TotalOperatingExpense = 0
				Else
					TotalOperatingExpense = GetPerformance("TotalOperatingExpense")
				End If
				
			Else
				SalariesWages = 0
				EmployeeBenefits = 0
				Insurance = 0
				Other = 0
				RentOccupancy = 0
				Total = 0
				Program = 0
				FundRaising = 0
				Administration = 0
				CategoryTotal = 0
				TotalOperatingExpense = 0
			End If

		'''Expenses %>
		<% if PrintForm="No" then %>
				
			<tr>
				<td colspan="6" class="formHeaderSmall">ANNUAL EXPENSES</td>
			</tr>	
				
		<% else %>
	
			<tr>
				<td colspan="6" class="formMainCentered"><strong>ANNUAL EXPENSES</strong></td>
			</tr>				
				
		<% end if %>
			<tr>
				<td class="formMain" width="24%">Salaries and Wages:</td>
				<td class="formMainRightJ" width="24%"><%= FormatCurrency(SalariesWages) %></td>
				<td colspan="2" width="4%">&nbsp;</td>		
				<td class="formMain" width="24%">Employee Benefits:</td>
				<td class="formMainRightJ" width="24%"><%= FormatCurrency(EmployeeBenefits) %></td>	
			</tr>
			<tr>
				<td class="formMain" width="24%">Liability Insurance:</td>				
				<td class="formMainRightJ" width="24%"><%= FormatCurrency(Insurance) %></td>	
				<td colspan="2" width="4%">&nbsp;</td>					
				<td class="formMain" width="24%">All Other:</td>								
				<td class="formMainRightJ" width="24%"><%= FormatCurrency(Other) %></td>			
			</tr>
			<tr>
				<td class="formMain" width="24%">Rent / Occupancy:</td>
				<td class="formMainRightJ" width="24%"><%= FormatCurrency(RentOccupancy) %></td>		
				<td colspan="2" width="4%">&nbsp;</td>				
				<td class="formMain" width="24%" bgcolor="#c0c0c0"><strong>Total Operating Expenses</strong></td>
				<td class="formMainRightJ" width="24%" bgcolor="#c0c0c0"><%= FormatCurrency(Total) %></td>		
			</tr>
		
			<%'''Expense Break down
				if PrintForm="No" then %>	
					<tr>
						<td class="formHeaderSmall" colspan="6">EXPENSE BREAKDOWN BY CATEGORY - Consistent with Audited Financial Statements<br>(Enter whole numbers only)</td>
					</tr>
				<% else %>
					<tr>
						<td class="formMainCentered" colspan="6"><strong>EXPENSE BREAKDOWN BY CATEGORY<br>(Enter whole numbers only)</strong></td>
					</tr>			
				<% end if %>
				
				<tr>
					<td class="formMain" colspan="5">Program:<br>
					<font class="formSubhead"><i>Including time spent supervising program staff</i></font></td>
					<td class="formMain" valign="top" align="left" colspan="2"><%= Program %>%</td>
				</tr>			
					
				<tr>
					<td class="formMain" colspan="5">Fundraising:</td>
					<td class="formMain" valign="top" align="left" colspan="2"><%= FundRaising %>%</td>
				</tr>		
				
				<tr>
					<td class="formMain" colspan="5">Administration:<br>
					<font class="formSubhead"><i>If any administration expenses are related to program or fundraising then include those expenses in program or fundraising when calculating percentages.</i></font></td>
					<td class="formMain" valign="top" align="left" colspan="2"><%= Administration %>%</td>
				</tr>
				
				<tr>
					<td class="formMain" colspan="5" bgcolor="#c0c0c0"><strong>TOTAL</strong></td>
					<td class="formMain" valign="top" align="left" bgcolor="#c0c0c0"><%= CategoryTotal %>%</td>
				</tr>
				<%
		
		Case Else 'Monthly Expenses not yearly expenses
%>
			<tr>
				<td valign="middle" align="center" class="formHeaderMedium" colspan="6">EXPENSE</td>
			</tr>
			
			<tr>
				<td valign="middle" class="formMain" colspan="5">Total Operating Expense<br>(should <strong>NOT</strong> include expense directly related to fundraising events)</td>
				<td valign="middle" class="formMain"><%=formatcurrency(TotalOperatingExpense)%></td>								
			</tr>					
<%	End Select

	Select case iMonth
	    Case 3, 6, 9, 12
			'Get data for year, month from tbl_frmBalanceSheet
			Dim rsBalanceSheet, strSQL, CashInvestments, Receivables, AllOtherAssets, TotalAssets, STLiabilities, LTLiabilities
			Dim TotalLiabilities, Surplus_NetAssets, TotalLiabilitiesAndNetAssets
			Set Con = Server.CreateObject("ADODB.Connection")
			Con.Open "BBBSAforms", "sa","12sist12"
			strSQL = "SELECT * FROM tbl_frmBalanceSheet WHERE AgencyID=" & Session("AgencyIDN") & " AND Yr=" & Int(Request("y")) & " AND Mth=" & Int(Request("m"))
			Set rsBalanceSheet = Con.Execute(strSQL)

			If not rsBalanceSheet.EOF Then
				CashInvestments = rsBalanceSheet("CashInvestments")
				Receivables = rsBalanceSheet("Receivables")
				AllOtherAssets = rsBalanceSheet("AllOtherAssets")
				TotalAssets = rsBalanceSheet("TotalAssets")
				STLiabilities = rsBalanceSheet("ShortTermLiabilities")
				LTLiabilities = rsBalanceSheet("LongTermLiabilities")
				TotalLiabilities = rsBalanceSheet("TotalLiabilities")
				Surplus_NetAssets = rsBalanceSheet("Surplus_NetAssets")
				TotalLiabilitiesAndNetAssets = rsBalanceSheet("TotalLiabNetAssets")
			Else
				CashInvestments = 0
				Receivables = 0
				AllOtherAssets = 0
				TotalAssets = 0
				STLiabilities = 0
				LTLiabilities = 0
				TotalLiabilities = 0
				Surplus_NetAssets = 0
				TotalLiabilitiesAndNetAssets = 0
			End If
			
			'Display Data	
			'Section header
			Response.Write("<tr><td valign='middle' align='center' class='formHeaderMedium' colspan=6>ANNUAL BALANCE SHEET<br>(as of " & MonthName(Request("m")) & " 31, " & Request("y") & "</td></tr>")
			
			'Data
			Response.Write("<tr><td valign='middle' class='formMain' colspan='5'>Cash Investments</td>")
			Response.Write("    <td valign='middle' class='formMain'>" & formatcurrency(CashInvestments) & "</td></tr>")					

			Response.Write("<tr><td valign='middle' class='formMain' colspan='5'>Receivables and Other Current Assets</td>")
			Response.Write("    <td valign='middle' class='formMain'>" & formatcurrency(Receivables) & "</td></tr>")					

			Response.Write("<tr><td valign='middle' class='formMain' colspan='5'>Non Current Assets</td>")
			Response.Write("    <td valign='middle' class='formMain'>" & formatcurrency(AllOtherAssets) & "</td></tr>")					

			Response.Write("<tr><td valign='middle' class='formMain' colspan='5'>Total Assets</td>")
			Response.Write("    <td valign='middle' class='formMain'>" & formatcurrency(TotalAssets) & "</td></tr>")					

			Response.Write("<tr><td valign='middle' class='formMain' colspan='5'>Current Liabilities</td>")
			Response.Write("    <td valign='middle' class='formMain'>" & formatcurrency(STLiabilities) & "</td></tr>")					

			Response.Write("<tr><td valign='middle' class='formMain' colspan='5'>Non Current Liabilities</td>")
			Response.Write("    <td valign='middle' class='formMain'>" & formatcurrency(LTLiabilities) & "</td></tr>")					

			Response.Write("<tr><td valign='middle' class='formMain' colspan='5'>Total Liabilities</td>")
			Response.Write("    <td valign='middle' class='formMain'>" & formatcurrency(TotalLiabilities) & "</td></tr>")					

			Response.Write("<tr><td valign='middle' class='formMain' colspan='5'>Net Assets</td>")
			Response.Write("    <td valign='middle' class='formMain'>" & formatcurrency(Surplus_NetAssets) & "</td></tr>")					

			Response.Write("<tr><td valign='middle' class='formMain' colspan='5'>Total Liabilities and Net Assets</td>")
			Response.Write("    <td valign='middle' class='formMain'>" & formatcurrency(TotalLiabilitiesAndNetAssets) & "</td></tr>")					

	    Case else
	End Select
%>
			
												
			
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
		
		

