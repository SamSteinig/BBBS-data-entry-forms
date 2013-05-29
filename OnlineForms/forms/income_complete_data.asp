<form name="frmIncome" action="Income_edit.asp?y=<%= Request("y") %>" method="post">
<!--#include file="../includes/form_stamp.asp"-->
<input type="hidden" name="status" value="editOld"> 

<table border="1" cellpadding="2" cellspacing="0" bordercolordark="003063" width="400">
	<tr>
		<td colspan="2" align="center" valign="top" class="formSubhead">BBBS - <%= y %> Annual Agency Information (AAI)</td>
	</tr>

<% Dim ReadOnlyLevel
If Session("ReadOnly") then
	ReadOnlyLevel=1
Else
	ReadOnlyLevel=0
End If
%>

<% if printform = "No" then %>		

	<tr>
		<td colspan="2" class="formHeader">REVENUE</td>
	</tr>
	
<% else %>

	<tr>
		<td colspan="2" class="formIndex">REVENUE</td>
	</tr>
	
<% end if %>




	<tr>
		<td colspan="2" class="formMainBold">Created: <%= GetIncome("CreateDate") %><br>
		<% form = "Income" %> 
		<% gid = GetIncome("IncomeID") %>
		<!--#include file="../includes/lastmodified_stamp.asp"-->
		</td>
	</tr>
	<tr>
		<td class="formMain">United Way</td>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("UnitedWay")) %></td>
	</tr>
	<tr>
		<td class="formMain">Federal Funding</td>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("FederalGovernmentFunding")) %></td>
	</tr>
	<tr>
		<td class="formMain">State Funding</td>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("StateGovernmentFunding")) %></td>
	</tr>	
	<tr>
		<td class="formMain">Local Funding</td>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("LocalGovernmentFunding")) %></td>
	</tr>	
	<tr>
		<td class="formMain">Foundation Grants</td>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("FoundationGrants")) %></td>
	</tr>
	<tr>
		<td class="formMain">Corporate Gifts</td>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("CorporateGifts")) %></td>
	</tr>
	<tr>
		<td class="formMain">BBBSA Grants</td>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("BBBSAGrants")) %></td>
	</tr>	
	<tr>
		<td class="formMain">Online Donations <span class="formSubHead">(through BBBSA)</span></td>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("OnlineDonations")) %></td>
	</tr>	
	<tr>
		<td class="formMain">RMM <span class="formSubHead">(Raising More Money)</span></td>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("RMM")) %></td>
	</tr>			
	
	<tr>
		<td class="formMain">Individual Giving - Board Members <span class="formSubHead">(excluding RMM)</span></td>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("IndividualGivingBoard")) %></td>
	</tr>
	<tr>
		<td class="formMain">Individual Giving - Non Board Members <span class="formSubHead">(excluding RMM)</span></td>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("IndividualGivingNonBoard")) %></td>
	</tr>	
	
	<% if printform = "No" then %>		

		<tr>
			<td colspan="2" class="formHeaderSmall">Special Events</td>
		</tr>
	
	<% else %>

		<tr>
			<td colspan="2" class="formMain"><div align="center"><strong>Special Events</strong></div></td>
		</tr>
	
	<% end if %>	
	

	
	<tr>
		<td class="formMain">Bowl For Kids' Sake <span class="formSubHead">(BFKS)</span><br>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("BowlForKidsSake")) %></td>
	</tr>
	<tr>
		<td class="formMain">Dinner / Auctions<br>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("DinnerAuctions")) %></td>
	</tr>	
	<tr>
		<td class="formMain">Golf<br>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("Golf")) %></td>
	</tr>		
	<tr>
		<td class="formMain">Bingo<br>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("Bingo")) %></td>
	</tr>		
	<tr>
		<td class="formMain">Raffle<br>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("Raffle")) %></td>
	</tr>		
	<tr>
		<td class="formMain">Cars For Kids' Sake <span class="formSubHead">(CFKS)</span></td>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("CarsForKidsSake")) %></td>
	</tr>
	<tr>
		<td class="formMain">Other Special Events <span class="formSubHead">(Total)</span><br>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("OtherSpecialEvents")) %></td>
	</tr>	
	
	<% if PrintForm="No" then %>
		
		<tr>
			<td class="formHeaderSmall" colspan="2">&nbsp;</td>
		</tr>
		
	<% else %>
	
		<tr>
			<td class="formMain" colspan="2">&nbsp;</td>
		</tr>	
		
	<% end if %>
	
	<tr>
		<td class="formMain">Dividends &amp; Interest</td>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("DividendsInterest")) %></td>
	</tr>
	<tr>
		<td class="formMain">Other Funding:&nbsp;<strong><%= GetIncome("OtherFundingType") %></strong></td>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("OtherFunding")) %></td>
	</tr>
	<tr>
		<td class="formMainRightJ"><strong>TOTAL</strong>&nbsp;&#61;&nbsp;</td>
		<td class="formMainRightJ">&nbsp;<%= FormatCurrency(GetIncome("Total")) %></td>
	</tr>
	<tr>
		<td class="formMain">Of this <%= FormatCurrency(GetIncome("Total")) %> total, amount targeted for <strong>Non-Mentoring</strong> programs</td>
		<td class="formMainRightJ">&nbsp;<%=FormatCurrency(GetIncome("NonMentoringIncome"))%></td>
	</tr>
	
<% if printform = "No" then %>	

	<% if ReadOnlyLevel = 0 then %>
		<tr>
			<td colspan="2" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold"></td>
		</tr>
	<% else %>
		<tr>
			<td colspan="9" class="formMainCentered">Where did the <strong>Edit Button</strong> go?  <a href="..\helpfiles\surveyhelp.asp?HelpID=password1" onclick="NewWindow(this.href,'name','500','295','yes');return false;">Click Here</a> for an explanation.</td>
		</tr>			
		
	<% end if %>
	
	<tr>
		<td colspan="2"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
	</tr>
	
<% end if %>
	
	
</table>