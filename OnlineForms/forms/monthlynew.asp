<!--#include file="../includes/session_stamp.asp"-->

<% Dim ReadOnlyLevel
If Session("ReadOnly") then
	ReadOnlyLevel=1
Else
	ReadOnlyLevel=0
End If
%>

<% Dim NoDataEntered 
NoDataEntered=0
%>

<% 

' Check to see if this form was submitted. If so, redirect to appropriate form.

If Request("status") = "bounce" Then
	z = Split(Request("month"),"-")
	m = z(0)
	y = z(1)
	f = Request("forms")


	Redim x(1)
	Redim x(2)
	Redim x(3)
	Redim x(4)
	x(1) = "Performance"
	x(2) = "SDMPerformance"
	x(3) = "OtherPerformance"
	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa", "12sist12"
		query = "SELECT PerformanceID FROM tbl_frmPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(y) & " AND Month=" & Int(m)
		Set GetData = Con.Execute(query)
		If (GetData.EOF OR GetData.BOF) Then
			'show blank form in new month only if selecting core performance
			if ReadonlyLevel = 0 then
				if f = "1" then
				
				
					' Pull Previous Month's Match Info only if a new record is being created.  If editing existing record, previous month's match info
					' is pulled in the performance_complete.asp page
					
					
					Set Con = Server.CreateObject("ADODB.Connection")
					Con.Open "BBBSAforms", "sa","12sist12"
					
					if m = 1 then
						query = "SELECT * FROM tbl_frmPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & y-1 & " AND Month=" & m+11
						PrevMonth = 12
						PrevYear = y-1
					else	
						query = "SELECT * FROM tbl_frmPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & y & " AND Month=" & m-1
						PrevMonth = m - 1
						PrevYear = y
					end if
					
					Set GetPrev = Con.Execute(query)
					
					PrevOpenComm = GetPrev("OpenMatchesCommunityBased")
					PrevOpenSchool = GetPrev("OpenMatchesSchoolBased")
					PrevOpenOther = GetPrev("OpenMatchesOtherSiteBased")
					PrevOpenGroup = GetPrev("OpenMatchesGroupMentoring")
					PrevOpenSpecMent = GetPrev("OpenMatchesSpecialProgramsMentoring")
					PrevOpenSpecNonMent = GetPrev("OpenMatchesSpecialProgramsNonMentoring")
					
					GetPrev.Close
					Set GetPrev = Nothing
					
					' End Pull Previous Month's Match Info
					
						Response.Redirect(x(f) & "_edit.asp?y=" & y & "&m=" & m & "&PrevOpenComm=" & PrevOpenComm & "&PrevOpenSchool=" & PrevOpenSchool & "&PrevOpenOther=" & PrevOpenOther & "&PrevOpenGroup=" & PrevOpenGroup & "&PrevOpenSpecMent=" & PrevOpenSpecMent & "&PrevOpenSpecNonMent=" & PrevOpenSpecNonMent)

				else
					NoDataMonth=int(m)
					NoDataEntered=1
				end if
			else 
				NoDataMonth=int(m)
				NoDataEntered=1
			end if
		Else
			'show complete form w/ edit button
			'z = x(f) & "ID"
			z = "PerformanceID"
			id = GetData(z)
			Response.Redirect(x(f) & "_complete.asp?y=" & y & "&m=" & m & "&id=" & id)
		End If
		GetData.Close
		Set GetData = Nothing	
	Con.Close
	Set Con = Nothing
End If
%>


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


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Monthly Performance Forms</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
	<% ' <!--#include file="../includes/top_nav_forms_monthly.inc"--><!-- include file has </head> and <body> tags --><br>     %>
	<!--#include file="../includes/surveytitle.inc"-->
	
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr>
			<td width="220" valign="top"><img src="../includes/images/photos_football.jpg" alt="" width="220" height="477" border="0"></td>
			<td width="100%" valign="top">
			<br><br>
			<span class="formIndex">Monthly Performance Forms</span>
			
			<%
				' Get all dates for performance forms that have been previously entered for the respective agency.
				' Add the next year/month to the drop down list
			
				set Con = Server.CreateObject("ADODB.Connection")
				Con.Open "BBBSAforms", "B3SAWeb", "6monkey6"
				sql = "p_getPerformanceFormDates '" & Session("AgencyIDN") & "'"
				Set rs = Con.Execute(sql)
				
				' Check to see if there are any results. If not, just display the current 
				' month and year.
			
				dim isRsEmpty
				dim thisMonth
				dim thisYear
				dim thisMonthName
				dim currentDate
				currentDate = Date
				isRsEmpty = false
				dim setNextMonth
				setNextMonth = true
					
				
				if rs.BOF and rs.EOF then
					
					' Populate thisMonth and thisYear with the last month and year
					' ( BBBSA does not want the current month to show )
					
					isRsEmpty		= true
					thisMonth		= month(currentDate) - 1
					thisMonthName	= monthname(thisMonth)
					thisYear		= year(currentDate)
				
				else	
					
					' Add the next month before the result set based on the most recent date 
					' in the resultset. But, do not add the current month if it would be 
					' the next month after the most recent result set. (The current month should
					' not show )
					
					' get the most recent date in the result set
					rs.MoveFirst
					dim newDate
					dim rsMonth
					dim rsYear
					rsMonth = rs("Month")
					rsYear = rs("Year")
					newDate = rsMonth & "-1-" & rsYear
					newDate = DateValue(newDate)
					newDate = DateAdd("m", 1, newDate)
					thisMonth = month(newDate)
					thisMonthName = monthname(thisMonth)
					thisYear = year(newDate)
					
					' determine if the next month/year in the series would be the current month/year
					
					dim currentMonth
					dim currentYear
					currentMonth = month(currentDate)
					currentYear = year(currentDate)
					if (currentMonth = thisMonth) and (currentYear = thisYear) then 
						setNextMonth = false
					end if
					
					' BEGIN APRIL 2003 SPECIAL BRANCH
					' The following branch is a work around just in case someone has already entered 
					' incomplete April 2003 data (the month/year when this logic to not show the current 
					' month/year in the drop down list was added. )
					
					dim showFirstRecordInResultSet
					showFirstRecordInResultSet = true
					
					if (rsMonth = currentMonth) and (rsYear = currentYear) then
						setNextMonth = false
						showFirstRecordInResultSet = false
					end if
					
					' END APRIL 2003 SPECIAL BRANCH
				end if 
				
			%>
			
			<form method="post" action="monthly.asp">
				<input type="hidden" name="status" value="bounce">
				<select name="forms" size=1 class="formMain">
					<option value="1" class="formMain">Core Business</option>
					<option value="2" class="formMain">SDM Metrics</option>
					<option value="3" class="formMain">Other Reports</option>
				</select>&nbsp;
				<select name="month" size=1 class="formMain">
					<% 	
						if setNextMonth = true then
							Response.Write "<option	value=""" & thisMonth & "-" & thisYear & """ class=""formMain"">" & thisYear & "-" & thisMonthName & "</option>"
						end if
						
						if isRsEmpty = false then
							rs.MoveFirst
							
							' BEGIN APRIL 2003 SPECIAL BRANCH
							' The following branch is a work around just in case someone has already entered 
							' incomplete April 2003 data (the month/year when this logic to not show the current 
							' month/year in the drop down list was added. )
							
							if showFirstRecordInResultSet = false then
								rs.MoveNext
							end if
							
							' END APRIL 2003 SPECIAL BRANCH
							
							
							while not rs.EOF
								thisMonth = rs("Month") 
								thisMonthName = MonthName(thisMonth)
								thisYear = rs("Year")
								Response.Write "<option	value=""" & thisMonth & "-" & thisYear & """ class=""formMain"">" & thisYear & "-" & thisMonthName & "</option>"	
								rs.MoveNext
							wend
						end if 
					%>
				</select>
				&nbsp;<input type="submit" value="Go" class="formMainBold">
			</form>
			<%
			if NoDataEntered = 1 then %>
			
			<span class="formMain">
			
			
				<br><b><em>No Data Has Been Entered for the month of <%=monthname(NoDataMonth)%></em></b>
				<br><br>
				<% if ReadOnlyLevel = 1 then %>
				<hr>
				I was able to enter a new month's data in the past. Why am I restricted from entering data now?<br><a href="..\helpfiles\surveyhelp.asp?HelpID=password1" onclick="NewWindow(this.href,'name','500','250','yes');return false;">Click Here</a> for an explanation.
				<hr>
				<% end if %>
				
			
			</span>	
				
				
			<% end if %>
			
			<% 
				' clean up
				rs.close
				set rs = nothing
				Con.close
				set Con = nothing
			%>
			
			<br>
			<!--#include file="../includes/contact_info.inc"-->
			<br>
			</td>
		</tr>
	</table>
			
</body>
</html>
