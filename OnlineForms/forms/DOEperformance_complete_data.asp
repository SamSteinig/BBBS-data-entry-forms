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

<% Dim ReadOnlyLevel
If Session("ReadOnly") then
	ReadOnlyLevel=1
Else
	ReadOnlyLevel=0
End If
%>




<% if printform = "No" then %>

	<tr>
		<td colspan="7" class="formHeader">DOE Grant Performance<br><%= MonthName(Request("m"), False) & " " & Request("y") %></td>
	</tr>
			
<% else %>			

	<tr>
		<td colspan="7" class="formIndex">DOE Grant Performance<br><%= MonthName(Request("m"), False) & " " & Request("y") %></td>
	</tr>	
			
<% end if %>

		<tr>
		
			<td colspan="7" class="formMainBold">Created: <%= GetDOEPerformance("CreateDate") %><br>		
			
			<% form = "DOEPerformance" %> 
			<% gid = GetDOEPerformance("DOEPerformanceID") %>
			<%= GetDOEPerformance("DOEPerformanceID") %>
			<!--#include file="../includes/lastmodified_stamp.asp"-->
			</td>
		</tr>
		<tr>
			<td align="center" valign="middle" class="formMain">&nbsp;</td>
			<td align="center" valign="middle" class="formMain">School Based</td>
		</tr>
		
		<!-- Matches Open/Active in the Beginning of the Month -->
		<% if y="2005" and m="1" then %>
		<% else %>
			<tr>
				<td align="center" valign="middle" class="formMain">
				OPEN/ACTIVE&nbsp;matches<br>on&nbsp;the&nbsp;<strong>FIRST</strong>&nbsp;day&nbsp;of<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				<td align="right" valign="middle" class="formMain"><%= PrevOpenSchool %></td>
			</tr>		
		<% end if %>

		<!-- Matches Closed During the Month -->
			<tr>
				<td align="center" valign="middle" class="formMain">
				Matches&nbsp;CLOSED&nbsp;during<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				<td align="right" valign="middle" class="formMain"><%= GetDOEPerformance("ClosedMatchesSchoolBased") %></td>
			</tr>		


		
		<!-- New Matches Opened During the Month -->
		<tr>	
			<% if y="2005" and m="1" then %>
				<td align="center" valign="middle" class="formMain"><b><font color="red">ONE-TIME Baseline Entry for January 2005:</b></font><br></font></b>Enter any DOE matches that existed<br><em>prior</em> to January 2005 <strong>PLUS</strong><br>any new DOE matches created <strong>DURING</strong> January 2005.</td>			
			<% else %>
				<td align="center" valign="middle" class="formMain">NEW&nbsp;matches opened<br>during&nbsp;<b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>			
			<% end if %>
			<td align="right" valign="middle" class="formMain"><%= GetDOEPerformance("NewMatchesSchoolBased") %></td>
		</tr>	
		
		


		</tr>					
	
			<tr>
				<td align="center" valign="middle" class="formMain">
				OPEN/ACTIVE&nbsp;matches<br>on&nbsp;the&nbsp;last&nbsp;day&nbsp;of<br><b><%= MonthName(Request("m"), False) & " " & Request("y") %></b></td>
				<td align="right" valign="middle" class="formMain"><%= GetDOEPerformance("OpenMatchesSchoolBased") %></td>
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
		
		

