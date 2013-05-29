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
		<td colspan="2" class="formHeader">PERFORMANCE - SBM GRANT PROGRESS REPORT<br><%= MonthName(Request("m"), False) & " " & Request("y") %></td>
	</tr>
			
<% else %>			

	<tr>
		<td colspan="2" class="formIndex">PERFORMANCE - SBM GRANT PROGRESS REPORT<br><%= MonthName(Request("m"), False) & " " & Request("y") %></td>
	</tr>	
			
<% end if %>

		<tr>
		
			<td colspan="2" class="formMainBold">Created: <%= GetPerformance("CreateDate") %><br>		
			
			<% form = "SBMPerformance" %> 
			<% gid = GetPerformance("SBMPerformanceID") %>
			<%= GetPerformance("SBMPerformanceID") %>
			<!--#include file="../includes/lastmodified_stamp.asp"-->
			</td>
		</tr>
		
		
		
<!-- SBM Questions -->



			<tr>
				<td valign="middle" class="formMain">Amount Raised Towards Match Pledge</td>
				<td valign="middle" class="formMain" ><%=formatcurrency(GetPerformance("SBMAmountRaisedTowardsMatchPledge"))%></td>			
			</tr>			

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
		
		

