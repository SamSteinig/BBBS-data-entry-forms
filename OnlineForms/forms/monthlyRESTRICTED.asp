<!--#include file="../includes/session_stamp.asp"-->

<% Dim ReadOnlyLevel
If Session("ReadOnly") then
	ReadOnlyLevel=1
Else
	ReadOnlyLevel=0
End If
%>

<% Dim AdminLevel
If Session("Admin") then 
	AdminLevel = 1
Else
	AdminLevel = 0
End If
%>


<% Dim NoDataEntered 
NoDataEntered=0


' Check for AIM Agency
DIM AIMAgency
Set AIMCon = Server.CreateObject("ADODB.Connection")
AIMCon.Open "BBBSAForms","sa","12sist12"
query = "SELECT AIMReconciled from tblDemogs Where AgencyID = " & int(Session("AgencyIDN"))

Set AIMQuery = AIMCon.Execute(query)

if (AIMQuery.eof) or (AIMQuery.bof) then
	AIMAgency = "n"
else
	if AIMQuery("AIMReconciled") = 1 then
		AIMAgency = "y"
	else
		AIMAgency = "n"
	end if	
end if 

' Check for DOE Agency
DIM DOEAgency
Set DOECon = Server.CreateObject("ADODB.Connection")
DOECon.Open "BBBSAforms","sa","12sist12"
query = "SELECT * FROM tblGrants_Grantees WHERE GranteeAgencyID = " & int(Session("AgencyIDN")) & " and (GranteesGroupID = 30 or GranteesGroupID = 31) " 

Set DOEQuery = DOECon.Execute(query)
if (DOEquery.eof) then
	DOEAgency = 0
else
	DOEAgency = 1
End if


' Check for MCP Agency
DIM MCPAgency
Set MCPCon = Server.CreateObject("ADODB.Connection")
MCPCon.Open "BBBSAforms","sa","12sist12"
query = "SELECT * FROM tblGrants_Grantees WHERE GranteeAgencyID = " & int(Session("AgencyIDN")) & " and (GranteesGroupID = 33 or GranteesGroupID = 34) " 

Set MCPQuery = MCPCon.Execute(query)
if (MCPquery.eof) then
	MCPAgency = 0
else
	MCPAgency = 1
End if


' Check for SBM Agency
'DIM SBMAgency
'Set SBMCon = Server.CreateObject("ADODB.Connection")
'SBMCon.Open "BBBSAforms","sa","12sist12"
'query = "SELECT SBM FROM tbl_AgencyInfo WHERE AgencyID = '" & Session("AgencyIDN") & "' and SBM = -1  " 
'Set SBMQuery = SBMCon.Execute(query)
'if (SBMquery.eof) then
'	SBMAgency = 0
'else
'	SBMAgency = 1
'End if




' Check to see if core form was submitted. If so, redirect to appropriate form.

If Request("CoreStatus") = "Bounce"  or Request("SDMStatus") = "Bounce"  or Request("SBMStatus") = "Bounce"  or Request("FinanceStatus") = "Bounce" or Request("DOEStatus") = "Bounce" or Request("MCPStatus") = "Bounce" Then
	z = Split(Request("month"),"-")
	m = z(0)
	y = z(1)
	
	If Request("CoreStatus") = "Bounce" Then
		f = Request("CoreForms")
		Redim x(1)
		x(1) = "Performance"
	Else
		If Request("SDMStatus") = "Bounce" Then
			f = Request("SDMForms")
			Redim x(2)
			x(2) = "SDMPerformance"
		Else
			If Request("SBMStatus") = "Bounce" Then
				f = Request("SBMForms")
				Redim x(3)
				x(3) = "SBMPerformance"
			Else
				If Request("FinanceStatus") = "Bounce" Then
					f = Request("FinanceForms")
					Redim x(4)
					x(4) = "FinancePerformance"
				Else
					If Request("DOEStatus") = "Bounce" Then
						f = Request("DOEForms")
						DOE="yes"
						Redim x(5)
						x(5) = "DOEPerformance"
					Else
						If Request("MCPStatus") = "Bounce" Then
							f = Request("MCPForms")
							MCP="yes"
							Redim x(6)
							x(6) = "MCPPerformance"					
						End If
					End If					
				End If
			End If		
		End If
	End If	



	Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa", "12sist12"
		query = "SELECT " & x(f) & "ID FROM tbl_frm" & x(f) & " WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(y) & " AND Month=" & Int(m)
		Set GetData = Con.Execute(query)
		If (GetData.EOF OR GetData.BOF) Then
			'show blank form in new month only if selecting core performance
			if ReadonlyLevel = 0 then
													
				' Pull Previous Month's Match Info only if a new record is being created, and only if performance is being selected.  If editing existing record, previous month's match info
				' is pulled in the performance_complete.asp page
					

				if x(f) = "Performance" or x(f) = "DOEPerformance"  or x(f) = "MCPPerformance" then
					Set Con = Server.CreateObject("ADODB.Connection")
					Con.Open "BBBSAforms", "sa","12sist12"
					
					if m = 1 then
						query = "SELECT * FROM tbl_frm" & x(f) & " WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & y-1 & " AND Month=" & m+11
						PrevMonth = 12
						PrevYear = y-1
					else	
						query = "SELECT * FROM tbl_frm" & x(f) & " WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & y & " AND Month=" & m-1
						PrevMonth = m - 1
						PrevYear = y
					end if
					
					Set GetPrev = Con.Execute(query)
					
					if GetPrev.BOF and GetPrev.EOF then
						PrevOpenComm = 0
						PrevOpenSchool = 0
						PrevOpenOther = 0
						PrevOpenGroup = 0
						PrevOpenSpecMent = 0
						PrevOpenSpecNonMent = 0
					else
					
					' Check to see which core performance form is selected.  Different forms use different fields
					
						Select Case x(f)
							Case "Performance"
								PrevOpenComm = GetPrev("OpenMatchesCommunityBased")
								PrevOpenSchool = GetPrev("OpenMatchesSchoolBased")
								PrevOpenOther = GetPrev("OpenMatchesOtherSiteBased")						
								PrevOpenGroup = GetPrev("OpenMatchesGroupMentoring")
								PrevOpenSpecMent = GetPrev("OpenMatchesSpecialProgramsMentoring")
								PrevOpenSpecNonMent = GetPrev("OpenMatchesSpecialProgramsNonMentoring")			
							Case "DOEPerformance"				
								PrevOpenSchool = GetPrev("OpenMatchesSchoolBased")									
							Case "MCPPerformance"
								PrevOpenComm = GetPrev("OpenMatchesCommunityBased")
								PrevOpenSchool = GetPrev("OpenMatchesSchoolBased")
								PrevOpenOther = GetPrev("OpenMatchesOtherSiteBased")						
						End Select							
					
					end if
					
					
					GetPrev.Close
					Set GetPrev = Nothing
				
				
				
					' End Pull Previous Month's Match Info
					
					'Pull Year to Date Info only from Core Performance form and only if a new record is being created and only if month > 1
					
				if x(f)="Performance" then					
											
						dim CommunityYTD
						CommunityYTD = 0
						dim SchoolYTD
						SchoolYTD = 0
						dim OtherYTD
						OtherYTD = 0
						dim RevenueYTD
						RevenueYTD = 0
						
					if m > 1 then
	
					 	Set Con = Server.CreateObject("ADODB.Connection")
						Con.Open "BBBSAforms", "sa","12sist12"
						query = "SELECT * FROM tbl_frmPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & y & " ORDER BY month"
	
						Set GetYTD = Con.Execute(query)	
	
						' Get First Month's Data
						CommunityYTD = GetYTD("ClosedMatchesCommunityBased") 
						SchoolYTD = GetYTD("ClosedMatchesSchoolBased")
						OtherYTD = GetYTD("ClosedMatchesOtherSiteBased") 
						RevenueYTD = GetYTD("Revenue")	
											
						
						count = 0
						for count = 1 to m - 1
								CommunityYTD = CommunityYTD + GetYTD("ClosedMatchesCommunityBased") 
								SchoolYTD = SchoolYTD + GetYTD("ClosedMatchesSchoolBased")
								OtherYTD = OtherYTD + GetYTD("ClosedMatchesOtherSiteBased") 
								RevenueYTD = RevenueYTD + GetYTD("Revenue")
								GetYTD.MoveNext()	
						next		
						
						GetYTD.Close
						Set GetYTD = Nothing		
						
					 	Set Con = Server.CreateObject("ADODB.Connection")
						Con.Open "BBBSAforms", "sa","12sist12"
						query = "SELECT * FROM tbl_frmPerformance WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & y & " AND Month=" & m-1
					
						Set GetYTD = Con.Execute(query)
						CommunityYTD = CommunityYTD + GetYTD("OpenMatchesCommunityBased")		
						SchoolYTD = SchoolYTD + GetYTD("OpenMatchesSchoolBased")	
						OtherYTD = OtherYTD + GetYTD("OpenMatchesOtherSiteBased")	
						
						
						
						GetYTD.Close
						Set GetYTD = Nothing						
							
						
					end if	
					'End Pull Year to Date Info
				end if
			end if
				
			Response.Redirect(x(f) & "_edit.asp?y=" & y & "&m=" & m & "&PrevOpenComm=" & PrevOpenComm & "&PrevOpenSchool=" & PrevOpenSchool & "&PrevOpenOther=" & PrevOpenOther & "&PrevOpenGroup=" & PrevOpenGroup & "&PrevOpenSpecMent=" & PrevOpenSpecMent & "&PrevOpenSpecNonMent=" & PrevOpenSpecNonMent & "&CommunityYTD=" & CommunityYTD & "&SchoolYTD=" & SchoolYTD & "&OtherYTD=" & OtherYTD & "&RevenueYTD=" & RevenueYTD)


			else 
				NoDataMonth=int(m)
				NoDataEntered=1
			end if
		Else		
			'show complete form w/ edit button
			z = x(f) & "ID"
			'z = "PerformanceID"
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
			<td width="220" valign="top">
			<img src="../includes/images/photos_football.jpg" alt="" width="220" height="477" border="0">
			<br><a href="FormAdminLogin.asp">...</a>
			</td>
			<td width="100%" valign="top">
			<br><br>
			<span class="formIndex">Monthly Performance Forms</span><br><br>
			
			<%
				' Get all dates for Performance forms that have been previously entered for the respective agency.
				' Add the next year/month to the drop down list
			
				set Con = Server.CreateObject("ADODB.Connection")
				Con.Open "BBBSAforms", "sa", "12sist12"
'				sql = "p_getPerformanceFormDatesTEMP '" & Session("AgencyIDN") & "'"   ' for Pulling All Years Wendy
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


			<table width = 400 cellpadding="3" cellspacing="2" border="1" bordercolor="#800080">
			
			<tr>
			<form method="post" action="monthly.asp">
				<input type="hidden" name="CoreStatus" value="Bounce">
				<input type="hidden" name="CoreForms" value="1">
				<td align="left" bgcolor="#cccccc" class="formMainBold">
				&nbsp;Core Business
				</td>
				<% if AimAgency = "n"  or AdminLevel = 1 then %>
				<td align="left" bgcolor="#cccccc" class="formMainBold">
				
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
					<td align="left" bgcolor="#cccccc">&nbsp;<input type="submit" value="Go" class="formMainBold"></td>
				<% else %>
					<td align="left" colspan="2" bgcolor="#cccccc" class="formMainBold">				
					AIM Agency - No Longer Required
				<% end if %>	
				</td>


			</form>
			</tr>


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

		
<!-- Trigger to turn off SDM Metrics Form on February 1, 2007 -->			
<% if date() < #02/01/2007# then %>
<!-- SDM Metrics Form -->

			<%
				' Get all dates for Performance forms that have been previously entered for the respective agency.
				' Add the next year/month to the drop down list
			
				set Con = Server.CreateObject("ADODB.Connection")
				Con.Open "BBBSAforms", "sa", "12sist12"
				sql = "p_getSDMPerformanceFormDates '" & Session("AgencyIDN") & "'"
				Set rs = Con.Execute(sql)
				
				' Check to see if there are any results. If not, just display the current 
				' month and year.
			
'				dim isRsEmpty
'				dim thisMonth
'				dim thisYear
'				dim thisMonthName
'				dim currentDate
				currentDate = Date
				isRsEmpty = false
'				dim setNextMonth
				setNextMonth = true
					
				
				if rs.BOF and rs.EOF then
					
					' Populate thisMonth and thisYear with the last month and year
					' ( BBBSA does not want the current month to show )
					
					isRsEmpty		= true
'					thisMonth		= month(currentDate) - 1
					thisMonth = 1
					thisMonthName	= monthname(thisMonth)
'					thisYear		= year(currentDate)
					thisYear = 2004
				
				else	
					
					' Add the next month before the result set based on the most recent date 
					' in the resultset. But, do not add the current month if it would be 
					' the next month after the most recent result set. (The current month should
					' not show )
					
					' get the most recent date in the result set
					rs.MoveFirst
'					dim newDate
'					dim rsMonth
'					dim rsYear
					rsMonth = rs("Month")
					rsYear = rs("Year")
					newDate = rsMonth & "-1-" & rsYear
					newDate = DateValue(newDate)
					newDate = DateAdd("m", 1, newDate)
					thisMonth = month(newDate)
					thisMonthName = monthname(thisMonth)
					thisYear = year(newDate)
					
					' determine if the next month/year in the series would be the current month/year
					
'					dim currentMonth
'					dim currentYear
					currentMonth = month(currentDate)
					currentYear = year(currentDate)
					if (currentMonth = thisMonth) and (currentYear = thisYear) then 
						setNextMonth = false
					end if
					
					' BEGIN APRIL 2003 SPECIAL BRANCH
					' The following branch is a work around just in case someone has already entered 
					' incomplete April 2003 data (the month/year when this logic to not show the current 
					' month/year in the drop down list was added. )
					
'					dim showFirstRecordInResultSet
					showFirstRecordInResultSet = true
					
					if (rsMonth = currentMonth) and (rsYear = currentYear) then
						setNextMonth = false
						showFirstRecordInResultSet = false
					end if
					
					' END APRIL 2003 SPECIAL BRANCH
				end if 
				
			%>
			

			
			<tr>
			<form method="post" action="monthly.asp">
				<input type="hidden" name="SDMStatus" value="Bounce">
				<input type="hidden" name="SDMForms" value="2">
				<td align="left" bgcolor="#cccccc" class="formMainBold">				
				&nbsp;SDM Metric Components
				</td>
				<% if AIMAgency = "n" or AdminLevel = 1 then %>
					<td align="left" bgcolor="#cccccc" class="formMainBold">				
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
					</td>
					<td align="left" bgcolor="#cccccc">				
					&nbsp;<input type="submit" value="Go" class="formMainBold">
				<% else %>
					<td align="left" bgcolor="#cccccc" colspan="2" class="formMainBold">				
					AIM Agency - Not Required
				<% end if %>				
				</td>
				</tr>
			</form>
			
			
<!-- End SDM Metrics Form -->

<!-- SDM Closure Metrics -->

			<tr>
			<form method="post" action="SDMClosureMetrics_complete.asp">
				<input type="hidden" name="SortField" value="MatchID">
				<input type="hidden" name="SortDirection" value="ASC">
			
				<td  align="left" bgcolor="#cccccc" class="formMainBold">				
				&nbsp;SDM Closure Metrics
				</td>
				<% if AIMAgency = "n" or AdminLevel = 1 then %>					
					<td  align="left" bgcolor="#cccccc" class="formMainBold">				
					&nbsp
					</td>
					<td align="left" bgcolor="#cccccc">
					&nbsp;<input type="submit" value="Go" class="formMainBold">
					</td>
				<% else %>
					<td  align="left" colspan="2" bgcolor="#cccccc" class="formMainBold">				
					AIM Agency - Not Required
					</td>
				<% end if %>
					
			</tr>
			</form>
<% end if %>

			
<!-- Quarterly Balance Sheet -->
<!-- DO NOT LAUNCH UNTIL 04/01/2007 - S.M.S. 
			<tr>
			<form method="post" action="quarterly.asp">
			
				<td  align="left" bgcolor="#cccccc" class="formMainBold">				
				&nbsp;Quarterly Balance Sheet <font color="#ff0000"><b> **NEW!**</b></font>
				</td>
			
				<td  align="left" bgcolor="#cccccc" class="formMainBold">				
				&nbsp
				</td>
				<td align="left" bgcolor="#cccccc">
				&nbsp;<input type="submit" value="Go" class="formMainBold">
				</td>


					
			</tr>
			</form>
-->				



<!-- Finance Form -->

			<%
				' Get all dates for Performance forms that have been previously entered for the respective agency.
				' Add the next year/month to the drop down list
			
				set Con = Server.CreateObject("ADODB.Connection")
				Con.Open "BBBSAforms", "sa", "12sist12"
				sql = "p_getFinancePerformanceFormDates '" & Session("AgencyIDN") & "'"
				Set rs = Con.Execute(sql)
				
				' Check to see if there are any results. If not, just display the current 
				' month and year.
			
'				dim isRsEmpty
'				dim thisMonth
'				dim thisYear
'				dim thisMonthName
'				dim currentDate
				currentDate = Date
				isRsEmpty = false
'				dim setNextMonth
				setNextMonth = true
					
				
				if rs.BOF and rs.EOF then
					
					' Populate thisMonth and thisYear with the last month and year
					' ( BBBSA does not want the current month to show )
					
					isRsEmpty		= true
					thisMonth		= month(currentDate) - 1
'					thisMonth = 1
					thisMonthName	= monthname(thisMonth)
					thisYear		= year(currentDate)
'					thisYear = 2005
				
				else	
					
					' Add the next month before the result set based on the most recent date 
					' in the resultset. But, do not add the current month if it would be 
					' the next month after the most recent result set. (The current month should
					' not show )
					
					' get the most recent date in the result set
					rs.MoveFirst
'					dim newDate
'					dim rsMonth
'					dim rsYear
					rsMonth = rs("Month")
					rsYear = rs("Year")
					newDate = rsMonth & "-1-" & rsYear
					newDate = DateValue(newDate)
					newDate = DateAdd("m", 1, newDate)
					thisMonth = month(newDate)
					thisMonthName = monthname(thisMonth)
					thisYear = year(newDate)
					
					' determine if the next month/year in the series would be the current month/year
					
'					dim currentMonth
'					dim currentYear
					currentMonth = month(currentDate)
					currentYear = year(currentDate)
					if (currentMonth = thisMonth) and (currentYear = thisYear) then 
						setNextMonth = false
					end if
					
					

				end if 
				
			%>
			

			
			<tr>
			<form method="post" action="monthly.asp">
				<input type="hidden" name="FinanceStatus" value="Bounce">
				<input type="hidden" name="FinanceForms" value="4">
				<td align="left" bgcolor="#cccccc" class="formMainBold">				
				&nbsp;Revenue / Expense
				</td>
				<td align="left" bgcolor="#cccccc">				
				<select name="month" size=1 class="formMain">
					<% 	
						if setNextMonth = true then
							Response.Write "<option	value=""" & thisMonth & "-" & thisYear & """ class=""formMain"">" & thisYear & "-" & thisMonthName & "</option>"
						end if
						
						if isRsEmpty = false then
							rs.MoveFirst
							
							
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
				</td>
				<td align="left" bgcolor="#cccccc">				
				&nbsp;<input type="submit" value="Go" class="formMainBold">
				</td>
				</tr>
			</form>
			
			
<!-- End Finance Form -->


<!-- DOE Form -->
	<% if DOEAgency = 1 then %>

			<%
				' Get all dates for Performance forms that have been previously entered for the respective agency.
				' Add the next year/month to the drop down list
			
				set Con = Server.CreateObject("ADODB.Connection")
				Con.Open "BBBSAforms", "sa", "12sist12"
				sql = "p_getDOEPerformanceFormDates '" & Session("AgencyIDN") & "'"
				Set rs = Con.Execute(sql)
				
				' Check to see if there are any results. If not, just display the current 
				' month and year.
			

				currentDate = Date
				isRsEmpty = false
				setNextMonth = true
					
				
				if rs.BOF and rs.EOF then
					
					' Populate thisMonth and thisYear with the last month and year
					' ( BBBSA does not want the current month to show )
					
					isRsEmpty		= true
'					thisMonth		= month(currentDate) - 1
					thisMonth = 1
					thisMonthName	= monthname(thisMonth)
'					thisYear		= year(currentDate)
					thisYear = 2005
				
				else	
					
					' Add the next month before the result set based on the most recent date 
					' in the resultset. But, do not add the current month if it would be 
					' the next month after the most recent result set. (The current month should
					' not show )
					
					' get the most recent date in the result set
					rs.MoveFirst
'					dim newDate
'					dim rsMonth
'					dim rsYear
					rsMonth = rs("Month")
					rsYear = rs("Year")
					newDate = rsMonth & "-1-" & rsYear
					newDate = DateValue(newDate)
					newDate = DateAdd("m", 1, newDate)
					thisMonth = month(newDate)
					thisMonthName = monthname(thisMonth)
					thisYear = year(newDate)
					
					' determine if the next month/year in the series would be the current month/year
					
'					dim currentMonth
'					dim currentYear
					currentMonth = month(currentDate)
					currentYear = year(currentDate)
					if (currentMonth = thisMonth) and (currentYear = thisYear) then 
						setNextMonth = false
					end if
					
					

				end if 
				
			%>
			

			
			<tr>
			<form method="post" action="monthly.asp">
				<input type="hidden" name="DOEStatus" value="Bounce">
				<input type="hidden" name="DOEForms" value="5">
				<td align="left" bgcolor="#cccccc" class="formMainBold">				
				&nbsp;DOE Grant Performance
				</td>
				<td align="left" bgcolor="#cccccc">				
				<select name="month" size=1 class="formMain">
					<% 	
						if setNextMonth = true then
							Response.Write "<option	value=""" & thisMonth & "-" & thisYear & """ class=""formMain"">" & thisYear & "-" & thisMonthName & "</option>"
						end if
						
						if isRsEmpty = false then
							rs.MoveFirst
							
							
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
				</td>
				<td align="left" bgcolor="#cccccc">				
				&nbsp;<input type="submit" value="Go" class="formMainBold">
				</td>
				</tr>
			</form>
			
	<% end if %>			
<!-- End DOE Form -->



<!-- MCP Form -->
	<% if MCPAgency = 1 then %>

			<%
				' Get all dates for Performance forms that have been previously entered for the respective agency.
				' Add the next year/month to the drop down list
			
				set Con = Server.CreateObject("ADODB.Connection")
				Con.Open "BBBSAforms", "sa", "12sist12"
				sql = "p_getMCPPerformanceFormDates '" & Session("AgencyIDN") & "'"
				Set rs = Con.Execute(sql)
				
				' Check to see if there are any results. If not, just display the current 
				' month and year.
			

				currentDate = Date
				isRsEmpty = false
				setNextMonth = true
					
				
				if rs.BOF and rs.EOF then
					
					' Populate thisMonth and thisYear with the last month and year
					' ( BBBSA does not want the current month to show )
					
					isRsEmpty		= true
'					thisMonth		= month(currentDate) - 1
					thisMonth = 1
					thisMonthName	= monthname(thisMonth)
'					thisYear		= year(currentDate)
					thisYear = 2005
				
				else	
					
					' Add the next month before the result set based on the most recent date 
					' in the resultset. But, do not add the current month if it would be 
					' the next month after the most recent result set. (The current month should
					' not show )
					
					' get the most recent date in the result set
					rs.MoveFirst
'					dim newDate
'					dim rsMonth
'					dim rsYear
					rsMonth = rs("Month")
					rsYear = rs("Year")
					newDate = rsMonth & "-1-" & rsYear
					newDate = DateValue(newDate)
					newDate = DateAdd("m", 1, newDate)
					thisMonth = month(newDate)
					thisMonthName = monthname(thisMonth)
					thisYear = year(newDate)
					
					' determine if the next month/year in the series would be the current month/year
					
'					dim currentMonth
'					dim currentYear
					currentMonth = month(currentDate)
					currentYear = year(currentDate)
					if (currentMonth = thisMonth) and (currentYear = thisYear) then 
						setNextMonth = false
					end if
					
					

				end if 
				
			%>
			

			
			<tr>
			<form method="post" action="monthly.asp">
				<input type="hidden" name="MCPStatus" value="Bounce">
				<input type="hidden" name="MCPForms" value="6">
				<td align="left" bgcolor="#cccccc" class="formMainBold">				
				&nbsp;MCP Grant Performance
				</td>
				<td align="left" bgcolor="#cccccc">				
				<select name="month" size=1 class="formMain">
					<% 	
						if setNextMonth = true then
							Response.Write "<option	value=""" & thisMonth & "-" & thisYear & """ class=""formMain"">" & thisYear & "-" & thisMonthName & "</option>"
						end if
						
						if isRsEmpty = false then
							rs.MoveFirst
							
							
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
				</td>
				<td align="left" bgcolor="#cccccc">				
				&nbsp;<input type="submit" value="Go" class="formMainBold">
				</td>
				</tr>
			</form>
			
	<% end if %>			
<!-- End MCP Form -->


			
			
<!-- SBM Grant Progress Report Form -->
	<% If SBMAgency=1 then %>
			<%
				' Get all dates for Performance forms that have been previously entered for the respective agency.
				' Add the next year/month to the drop down list
			
				set Con = Server.CreateObject("ADODB.Connection")
				Con.Open "BBBSAforms", "sa", "12sist12"
				sql = "p_getSBMPerformanceFormDates '" & Session("AgencyIDN") & "'"
				Set rs = Con.Execute(sql)
				
				' Check to see if there are any results. If not, just display the current 
				' month and year.
			
'				dim isRsEmpty
'				dim thisMonth
'				dim thisYear
'				dim thisMonthName
'				dim currentDate
				currentDate = Date
				isRsEmpty = false
'				dim setNextMonth
				setNextMonth = true
					
				
				if rs.BOF and rs.EOF then
					
					' Populate thisMonth and thisYear with the last month and year
					' ( BBBSA does not want the current month to show )
					
					isRsEmpty		= true
'					thisMonth		= month(currentDate) - 1
					thisMonth		= 12
					thisMonthName	= monthname(thisMonth)
'					thisYear		= year(currentDate)
					thisYear		= 2004                   ' [?????]
			
				
				else	
					
					' Add the next month before the result set based on the most recent date 
					' in the resultset. But, do not add the current month if it would be 
					' the next month after the most recent result set. (The current month should
					' not show )
					
					' get the most recent date in the result set
					rs.MoveFirst
					rsMonth = rs("Month")
					rsYear = rs("Year")
					newDate = rsMonth & "-1-" & rsYear
					newDate = DateValue(newDate)
					newDate = DateAdd("m", 6, newDate)
					thisMonth = month(newDate)
					thisMonthName = monthname(thisMonth)
					thisYear = year(newDate)
					
					' determine if the next month/year in the series would be the current month/year
					
					currentMonth = month(currentDate)
					currentYear = year(currentDate)
					if (currentMonth = thisMonth) and (currentYear = thisYear) then 					
						setNextMonth = false
					end if
					

				end if 
				
			%>
			

			
			<tr>

			<form method="post" action="monthly.asp">
				<input type="hidden" name="SBMStatus" value="Bounce">
				<input type="hidden" name="SBMForms" value="3">
				<td align="left" bgcolor="#cccccc" class="formMainBold">
				&nbsp;SBM Grant Progress Report
				</td>
				<td align="left" bgcolor="#cccccc">		
						<%if currentMonth-1=0 then
							currentMonth=13
						end if%>
				<select name="month" size=1 class="formMain">
					<% 
'						If there are no records in the SBM Table (agency entering for the first time) OR if the next available month is June or December, then create new record

						if isRsEmpty = true or (setNextMonth = true and (currentMonth-1 = 6 or currentMonth-1 = 12)) then
'						if SetNextMonth = true and month(date) < currentmonth then
							Response.Write "<option	value=""" & thisMonth & "-" & thisYear & """ class=""formMain"">" & thisYear & "-" & thisMonthName & "</option>"
						end if
						
						if isRsEmpty = false then
							rs.MoveFirst
							
							
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
				</td>
				<td align="left" bgcolor="#cccccc">				
				&nbsp;<input type="submit" value="Go" class="formMainBold">
				</td>
				</tr>
			</form>
			
		<% end if %>
<!-- End SBM Grant Progress Report -->				

		</table>

			<%

				' clean up
				rs.close
				set rs = nothing
				Con.close
				set Con = nothing
				
			   'Add Monthly reminder
			   ' DO NOT LAUNCH 
			   ' Select Case Month(Date)
			   '    Case 1, 4, 7, 10
			   '        Response.Write("<BR><span class=" & Chr(34) & "formIndex" & Chr(34) & "><font color=red>Please remember to complete your Quarterly Balance Sheet</font></span>")
			   ' End Select

			%>			
			
			
			
			
			<br>
			<!--#include file="../includes/contact_info.inc"-->
			<br>
			</td>
		</tr>
	</table>
			
</body>
</html>
