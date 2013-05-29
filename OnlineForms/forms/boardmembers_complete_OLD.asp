
<!--#include file="../includes/session_stamp.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Board Members</title>
	<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">
	
<% '<!--#include file="../includes/top_nav_forms_yearly.inc"--><!-- include file has </head> and <body> tags --><br>%>
<!--#include file="../includes/surveytitle.inc"-->

<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
<td width="220" valign="top"><img src="../includes/images/photos_slinky.jpg" alt="" width="220" height="477" border="0"></td>
<td width="100%" valign="top">
<br>

<% 	
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
query = "SELECT * FROM tbl_frmBoardMembers WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
Set GetBoardMembers = Con.Execute(query)
 %>	

			<table width="640" border="1" cellspacing="0" cellpadding="3" bordercolordark="#003063">
			<form name="frmBoardMembers" action="BoardMembers_edit.asp?y=<%= Request("y") %>" method="post">
			<!--#include file="../includes/form_stamp.asp"-->
			<input type="hidden" name="status" value="editOld">
				<tr> 
					<td colspan="3" align="center" valign="top" class="formSubhead">BBBS - <%= y %> Annual Agency Information (AAI)</td>
				</tr>
				<tr>
					<td colspan="3" class="formHeader">BOARD MEMBERS</td>
				</tr>
					<tr>
					<td colspan="3" class="formMainBold">Created: <%= GetBoardMembers("CreateDate") %>
		<% form = "BoardMembers" %> 
		<% gid = GetBoardMembers("BoardMembersID") %>
		<!--#include file="../includes/lastmodified_stamp.asp"-->
					</td>
					</tr>

<!-- Question Number 1 -->
				<tr>
					<td align="left" valign="top" class="formMain">1.</td>
					<td align="left" valign="top" class="formMain">Number of Board Members as of 12/31:</td>
					<td align="left" valign="top" class="formMain">&nbsp;<%= GetBoardMembers("NumberBoardMembers") %></td>
				</tr>
<!-- Question Number 2 -->
				<tr> 
					<td align="left" valign="top" class="formMain">2.</td>
					<td align="left" valign="top" class="formMain">If you have term limits for Board Members, enter number of years:</td>
					<td align="left" valign="top" class="formMain">&nbsp;<%= GetBoardMembers("TermLimitsYears") %></td>
				</tr>
<!-- Question Number 3 -->
				<tr>
					<td align="left" valign="top" class="formMain">3.</td>
					<!-- the reason there are non breaking spaces in between each word in this field is that in netscape it makes the column expand -->
					<td align="left" valign="top" class="formMain" >What&nbsp;is&nbsp;the&nbsp;average&nbsp;tenure&nbsp;of&nbsp;your&nbsp;board&nbsp;members?</td>
					<td align="left" valign="top" class="formMain"><%= GetBoardMembers("AverageTenureYears") %>&nbsp;Years&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= GetBoardMembers("AverageTenureMonths") %>&nbsp;Months&nbsp;</td>
				</tr>
<!-- Question Number 4 -->
				<tr>
					<td align="left" valign="top" class="formMain">4.</td>
					<td colspan="2" align="left" valign="top" class="formMain">What standing committees do you have? (Please check all that apply.)<br>
<!-- nested table -->
						<table width="640" border="0" cellspacing="3" cellpadding="3">
							<tr>
								<td align="left" valign="top" class="formMain">1 Personnel <% If GetBoardMembers("StandingCommitteesPersonnel") = True Then %><b>X</b><% End If %></td>
								<td align="left" valign="top" class="formMain">2 Program <% If GetBoardMembers("StandingCommitteesProgram") = True Then %><b>X</b><% End If %></td>
								<td align="left" valign="top" class="formMain">3 Executive <% If GetBoardMembers("StandingCommitteesExecutive") = True Then %><b>X</b><% End If %></td>
								<td align="left" valign="top" class="formMain">4 Fund Development <% If GetBoardMembers("StandingCommitteesFundDevelopment") = True Then %><b>X</b><% End If %></td>
							</tr>
							<tr>
								<td align="left" valign="top" class="formMain">5 Finance <% If GetBoardMembers("StandingCommitteesFinance") = True Then %><b>X</b><% End If %></td>
								<td align="left" valign="top" class="formMain">6 Public Relations <% If GetBoardMembers("StandingCommitteesPublicRelations") = True Then %><b>X</b><% End If %></td>
								<td align="left" valign="top" class="formMain">7 Strategic Planning <% If GetBoardMembers("StandingCommitteesStrategicPlanning") = True Then %><b>X</b><% End If %></td>
								<td align="left" valign="top" class="formMain">8 Board Development <% If GetBoardMembers("StandingCommitteesBoardDevelopment") = True Then %><b>X</b><% End If %></td>
							</tr>
							<tr>
								<td align="left" valign="top" class="formMain">9 Volunteer Recruitment <% If GetBoardMembers("StandingCommitteesVolunteerRecruitment") = True Then %><b>X</b><% End If %></td>
								<td colspan="3" align="left" valign="top" class="formMain">10 Other <% If GetBoardMembers("StandingCommitteesOther") = True Then %><b>X</b><% End If %> (Name): <%= GetBoardMembers("StandingCommitteesOtherText") %></td>
							</tr>
						</table>
					</td>		
				</tr>
<!-- Question Number 5 -->
				<tr>
					<td align="left" valign="top" class="formMain">5.</td>
					<td colspan="2" align="left" valign="top" class="formMain">Enter the number of FEMALE board members by ethnicity below:<br>
<!-- nested table -->
						<table width="640" border="0" cellspacing="3" cellpadding="3" align="center">
							<tr>
								<td align="left" valign="top" class="formMain">White (Not Hispanic)<br><%= GetBoardMembers("FemaleWhite") %></td>
								<td align="left" valign="top" class="formMain">Black<br><%= GetBoardMembers("FemaleBlack") %></td>
								<td align="left" valign="top" class="formMain">Hispanic<br><%= GetBoardMembers("FemaleHispanic") %></td>
								<td align="left" valign="top" class="formMain">Asian<br><%= GetBoardMembers("FemaleAsian") %></td>
								
							</tr>
							<tr>
								<td align="left" valign="top" class="formMain">Pacific Islander<br><%= GetBoardMembers("FemaleIslander") %></td>
								<td align="left" valign="top" class="formMain">Native American<br><%= GetBoardMembers("FemaleNative") %></td>
								<td align="left" valign="top" class="formMain">Multi-Racial<br><%= GetBoardMembers("FemaleMulti") %></td>
								<td align="left" valign="top" class="formMain">Unknown<br><%= GetBoardMembers("FemaleUnknown") %></td>
							</tr>
						</table> 
					</td>		
				</tr>
<!-- Question Number 6 -->
				<tr>
					<td align="left" valign="top" class="formMain">6.</td>
					<td colspan="2" align="left" valign="top" class="formMain">Enter the number of MALE board members by ethnicity below:<br>
<!-- nested table -->
						<table width="640" border="0" cellspacing="3" cellpadding="3" align="center">
							<tr>
								<td align="left" valign="top" class="formMain">White (Not Hispanic)<br><%= GetBoardMembers("MaleWhite") %></td>
								<td align="left" valign="top" class="formMain">Black<br><%= GetBoardMembers("MaleBlack") %></td>
								<td align="left" valign="top" class="formMain">Hispanic<br><%= GetBoardMembers("MaleHispanic") %></td>
								<td align="left" valign="top" class="formMain">Asian<br><%= GetBoardMembers("MaleAsian") %></td>
							</tr>
							<tr>
								<td align="left" valign="top" class="formMain">Pacific Islander<br><%= GetBoardMembers("MaleIslander") %></td>
								<td align="left" valign="top" class="formMain">Native American<br><%= GetBoardMembers("MaleNative") %></td>
								<td align="left" valign="top" class="formMain">Multi-Racial<br><%= GetBoardMembers("MaleMulti") %></td>
								<td align="left" valign="top" class="formMain">Unknown<br><%= GetBoardMembers("MaleUnknown") %></td>
							</tr>
						</table>
					</td>		
				</tr>
<!-- Question Number 7 -->
				<tr>
					<td align="left" valign="top" class="formMain">7.</td>
					<td colspan="2" align="left" valign="top" class="formMain">Frequency of board meetings.<br>
<!-- nested table -->
						<table width="640" border="0" cellspacing="3" cellpadding="3" align="center">
							<tr>
								<td align="center" valign="top" class="formMain">Monthly <% If GetBoardMembers("FrequencyMonthly") = True Then %><b>X</b><% End If %></td>
								<td align="center" valign="top" class="formMain">Every 2 Months <% If GetBoardMembers("FrequencyTwoMonths") = True Then %><b>X</b><% End If %></td>
								<td align="center" valign="top" class="formMain">Quarterly <% If GetBoardMembers("FrequencyQuarterly") = True Then %><b>X</b><% End If %></td>
								<td align="center" valign="top" class="formMain">Other <% If GetBoardMembers("FrequencyOther") = True Then %><b>X</b><% End If %> <%= GetBoardMembers("FrequencyOtherText") %></td>
							</tr> 
						</table> 
					</td>		
				</tr>
<!-- Question Number 8 --> 
				<tr> 
					<td align="left" valign="top" class="formMain">8.</td>
					<td colspan="3" align="left" valign="top" class="formMain">Are all board members expected to:<br> 
						a. <% If GetBoardMembers("MoneyMinimum") = True Then %><b>X</b><% End If %> Make a miniumum annual financial commitment? If checked, indicate the amount:&nbsp;&nbsp;&nbsp;<%=  FormatCurrency(GetBoardMembers("MoneyMinimumAmount")) %><br>
						b. <% If GetBoardMembers("MoneyInKind") = True Then %><b>X</b><% End If %> Make either a monetary or in-kind contribution - no specified amount<br>
						c. <% If GetBoardMembers("MoneyNotExpected") = True Then %><b>X</b><% End If %> Not expected, but encouraged<br>
						d. <% If GetBoardMembers("MoneyNoPolicy") = True Then %><b>X</b><% End If %> No policy or expectations
					</td>
				</tr>
<!-- Question Number 9 -->
				<tr> 
					<td align="left" valign="top" class="formMain">9.</td>
					<td align="left" valign="top" class="formMain">How much money did your board members contribute this past year?</td>
					<td align="left" valign="top" class="formMain">&nbsp;&nbsp;<%= FormatCurrency(GetBoardMembers("YearlyContribution")) %></td>
				</tr>	
<!-- Question Number 10 -->
				<tr>
					<td align="left" valign="top" class="formMain">10.</td>
					<td colspan="2" align="left" valign="top" class="formMain">Please enter the number of board members by professional skills/expertise below.<br>
<!-- nested table -->
						<table width="640" border="0" cellspacing="1" cellpadding="1" align="center">
							<tr>
								<td align="left" valign="top" class="formMain">Finance/Accounting/Banking<br><%= GetBoardMembers("SkillsFinance") %></td>
								<td align="left" valign="top" class="formMain">Legal<br><%= GetBoardMembers("SkillsLegal") %></td>
								<td align="left" valign="top" class="formMain">Public Relations<sup>1</sup><br><%= GetBoardMembers("SkillsPublicRelations") %></td>
								<td colspan="2" align="left" valign="top" class="formMain">Human Services Practitioner<sup>2</sup><br><%= GetBoardMembers("SkillsHumanServicesPractitioner") %></td>								
							</tr>
							<tr>
								<td align="left" valign="top" class="formMain">Human Services Administrator<sup>2</sup><br><%= GetBoardMembers("SkillsHumanServicesAdministrator") %></td>
								<td align="left" valign="top" class="formMain">Full Time College/H.S. Student<br><%= GetBoardMembers("SkillsFullTimeStudent") %></td>
								<td align="left" valign="top" class="formMain">Human Resources<br><%= GetBoardMembers("SkillsHumanResources") %></td>
								<td colspan="2" align="left" valign="top" class="formMain">Corporate CEO<br><%= GetBoardMembers("SkillsCorporateCEO") %></td>
							</tr>
							<tr>
								<td align="left" valign="top" class="formMain">Other Corporate Officer<br><%= GetBoardMembers("SkillsOtherCorporateOfficer") %></td>
								<td align="left" valign="top" class="formMain">Insurance/Sales<br><%= GetBoardMembers("SkillsInsurance") %></td>
								<td align="left" valign="top" class="formMain">Small Business Owner<br><%= GetBoardMembers("SkillsSmallBusiness") %></td><br>
								<td colspan="2" align="left" valign="top" class="formMain">Big<br><%= GetBoardMembers("SkillsBig") %></td>								
							</tr>
							<tr>
								<td align="left" valign="top" class="formMain">Parent of Little<br><%= GetBoardMembers("SkillsParentLittle") %></td>
								<td align="left" valign="top" class="formMain">Little<br><%= GetBoardMembers("SkillsLittle") %></td>
								<td align="left" valign="top" class="formMain">Local Government<br><%= GetBoardMembers("SkillsLocalgovernment") %></td>
								<td align="left" valign="top" class="formMain">Other<br><%= GetBoardMembers("SkillsOther") %></td>
								<td align="left" valign="top" class="formMain">Unknown<br><%= GetBoardMembers("SkillsUnknown") %></td>

						</table> 
					</td>		
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td colspan="2" class="formSubHead">(1) Public Relations Includes: Marketing, Communications, Graphic Design<br>(2) Human Services Includes: Teacher, Psychologist, Social Worker</td>
				</tr>
				<tr>
					<td colspan="3" class="formHeader"><input type="submit" value="Edit Form" class="formMainBold"></td>
				</tr>
				<tr>
					<td colspan="3"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
				</tr>

			</table>	
			

<% 
GetBoardMembers.Close
Set GetBoardMembers = Nothing
Con.Close
Set Con = Nothing
 %>
</form>
	<p>&nbsp;</p>
	<p>&nbsp;</p>   	
</td>
</tr>
</table>
</body>
</html>
