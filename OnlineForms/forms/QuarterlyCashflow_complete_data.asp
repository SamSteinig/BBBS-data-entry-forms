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
		<td colspan="2" class="formHeader">Quarterly Balance Sheet<br>Q<%= Request("q") & " " & Request("y") %></td>
	</tr>
			
<% else %>			

	<tr>
		<td colspan="2" class="formIndex">Quarterly Balance Sheet<br>Q<%= Request("q") & " " & Request("y") %></td>
	</tr>	
			
<% end if %>

		<tr>
		
			<td colspan="2" class="formMainBold">Created: <%= GetQCF("CreateDate") %><br>		
			
			<% form = "QuarterlyCashflow" %> 
			<% gid = GetQCF("QuarterlycashflowID") %>
			<%= GetQCF("QuarterlycashflowID") %>
			<!--#include file="../includes/lastmodified_stamp.asp"-->
			</td>
		</tr>

		<tr><td valign="middle" align="center" class="formHeaderMedium" colspan="2">&nbsp;</td></tr>

		<tr><td valign="middle" class="formMain">Cash and Investments</td>
			<td valign="middle" class="formMain"><%= formatcurrency(GetQCF("CashAndInvestments")) %></td></tr>	
		<tr><td valign="middle" class="formMain">Receivables</td>
			<td valign="middle" class="formMain"><%= formatcurrency(GetQCF("Receivables")) %></td></tr>	
		<tr><td valign="middle" class="formMain">All Other Assets</td>
			<td valign="middle" class="formMain"><%= formatcurrency(GetQCF("AllOtherAssets")) %></td></tr>				
		<tr><td valign="middle" class="formMain">Total Assets</td>
			<td valign="middle" class="formMain"><%= formatcurrency(GetQCF("TotalAssets")) %></td></tr>					

		<tr><td colspan=2>&nbsp;</td>
		
		<tr><td valign="middle" class="formMain">Current Liabilities</td>
			<td valign="middle" class="formMain"><%= formatcurrency(GetQCF("CurrentLiabilities")) %></td></tr>	
		<tr><td valign="middle" class="formMain">Long-Term Liabilities</td>
			<td valign="middle" class="formMain"><%= formatcurrency(GetQCF("LongTermLiabilities")) %></td></tr>	
		<tr><td valign="middle" class="formMain">Total Liabilities</td>
			<td valign="middle" class="formMain"><%= formatcurrency(GetQCF("TotalLiabilities")) %></td></tr>					

		<tr><td colspan=2>&nbsp;</td>
		
		<tr><td valign="middle" class="formMain">Net Assets</td>
			<td valign="middle" class="formMain"><%= formatcurrency(GetQCF("NetAssets")) %></td></tr>					
		<tr><td valign="middle" class="formMain">Liabilities and Net Assets</td>
			<td valign="middle" class="formMain"><%= formatcurrency(GetQCF("LiabilitiesAndNetAssets")) %></td></tr>					

<!--		<tr><td colspan=2>&nbsp;</td>
-->		
<!--		<tr><td colspan="2" class="formHeader"><input type="button" value="Save Form" class="formMainBold" onclick="validateForm(); return false;"  onclick="TotalNet();"  id=button1 name=button1></td></tr>
-->			
<% If printform = "No" Then %>

		<% if DisplayMetrics = 1 then %>
			<% if ReadOnlyLevel=0 then %>
				<tr><td colspan="9" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold" id=submit1 name=submit1></td></tr>
			<% else %>
				<tr><td colspan="9" class="formMain">Where did the <strong>Edit Button</strong> go?  <a href="..\helpfiles\surveyhelp.asp?HelpID=password1" onclick="NewWindow(this.href,'name','500','250','yes');return false;">Click Here</a> for an explanation.</td></tr>			
			<% end if %>
				
				<tr><td colspan="9"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td></tr>				
		<% else %>
		
			<% if ReadOnlyLevel=0 then %>
				<tr><td colspan="7" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold" id=submit2 name=submit2></td></tr>
			<% else %>
				<tr><td colspan="7" class="formMain" align="center"><strong>HEY!!! Where did the <em>Edit</em> Button</strong> go?  <a href="..\helpfiles\surveyhelp.asp?HelpID=password1" onclick="NewWindow(this.href,'name','500','250','yes');return false;">Click Here</a> for an explanation.</td></tr>			
			<% end if %>				
				<tr><td colspan="7"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td></tr>
		<% end if %>
<% End If %>			
		</table>
		
		

