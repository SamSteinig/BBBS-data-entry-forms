<!-- RESULTS TABLE STARTS HERE -->
		<% if printform="No" then %>
			<table border="1" cellpadding="2" cellspacing="0" bordercolordark="003063" width="660">
		<% else %>
			<table border="1" cellpadding="2" cellspacing="0" bordercolordark="003063" width="600">		
		<% end if %>
			<tr>
				<td colspan="8" align="center" class="formSubhead">BBBS - <%= Request("y") %> Annual Agency Information (AAI)</td>
			</tr>
			<tr>
				<td colspan="8" align="center" class="formHeader">STAFF</td>
			</tr>
			<tr>
				<td colspan="8" bgcolor="#ffffff"><p>&nbsp;</p></td>
			</tr>
			
			<% if printform="No" then %>
								 
			<form name="frmStaff" action="staff_edit.asp?y=<%= Request("y") %>" method="post" ID="Form1">
			<!--#include file="../includes/form_stamp.asp"-->
			<input type="hidden" name="status" value="newStaff" ID="Hidden1">
						<tr>
			                <td colspan="8" class="formHeader"><input type="submit" value="Add Additional Staff Member" class="formMainBold" ID="Submit1" NAME="Submit1"></td>
			       		</tr>
						
			</form>
			
			<% end if %>			

<!-- first row of table headers -->
			<tr>
			<% if printform="No" then %>		
				<td rowspan="3" class="formHeaderSmall">#</td>
				<td class="formHeaderSmall">Birth Year:</td>
				<td class="formHeaderSmall">Position:</td>
				<td class="formHeaderSmall">Race:</td>
				<td class="formHeaderSmall">Gender:</td>
				<td class="formHeaderSmall" colspan="2">Compensation<br>(Base Salary)</td>
				<!--<td rowspan="3" class="formHeaderSmall">Edit/Delete</td>  Remove Delete function by Diane 6/19/2008 slp-->
				<td rowspan="3" class="formHeaderSmall">Edit</td>
			<% else %>
				<td rowspan="3" class="formMainBold"><div align="center">#</div></td>
				<td class="formMainBold"><div align="center">Birth Year:</div></td>
				<td class="formMainBold"><div align="center">Position:</div></td>
				<td class="formMainBold"><div align="center">Race:</div></td>
				<td class="formMainBold"><div align="center">Gender:</div></td>
				<td class="formMainBold" colspan="2"><div align="center">Compensation<br>(Base Salary)</div></td>
			<% end if %>
			</tr>
<!-- second row of table headers -->
			<tr>
			<% if printform="No" then %>			
				<td class="formHeaderSmall">Hired (M/Y):</td>
				<td class="formHeaderSmall">Position Start:</td>
				<td class="formHeaderSmall">Status:</td>
				<td class="formHeaderSmall">Hrs/Wk:</td>
				<td class="formHeaderSmall" colspan="2">Compensation<br>(Bonus/Incentives)</td>
			<% else %>
				<td class="formMainBold"><div align="center">Hired (M/Y):</div></td>
				<td class="formMainBold"><div align="center">Position Start:</div></td>
				<td class="formMainBold"><div align="center">Status:</div></td>
				<td class="formMainBold"><div align="center">Hrs/Wk:</div></td>
				<td class="formMainBold" colspan="2"><div align="center">Compensation<br>(Bonus/Incentives)</div></td>		
			<% end if %>
			</tr>
<!-- third row of table headers -->
			<tr>
			<% if printform="No" then %>			
				<td class="formHeaderSmall">Years in BBBS:</td>
				<td class="formHeaderSmall">Employee Name:</td>
				<td class="formHeaderSmall">Education:</td>
				<td class="formHeaderSmall">EverABig</td>
				<td class="formHeaderSmall" colspan="2">Total Compensation<br>(Salary+Bonus/Incentives)</td>
			<% else %>
				<td class="formMainBold"><div align="center">Years in BBBS:</div></td>
				<td class="formMainBold"><div align="center">Employee Name:</div></td>
				<td class="formMainBold"><div align="center">Education:</div></td>
				<td class="formMainBold"><div align="center">EverABig</div></td>
				<td class="formMainBold" colspan="2"><div align="center">Total Compensation<br>(Salary+Bonus/Incentives)</div></td>		
			<% end if %>
			</tr>
					<%
						ct = 1
						Set Con = Server.CreateObject("ADODB.Connection")
						Con.Open "BBBSAforms", "sa","12sist12"
						query = "SELECT * FROM tbl_frmStaff WHERE AgencyID='" & Session("AgencyIDN") & "' AND Year=" & Int(Request("y"))
						Set GetStaff = Con.Execute(query)
					%>
<script language="JavaScript">
<!-- 
				function confirmDelete(row)
	{
		if (confirm("Are you sure you want to delete this record?"))
		{
			location = "staff_edit.asp?status=deleteRow&row="+ row +"&y=<%= Request("y") %>";
			//alert("Record deleted.");
		}
		else
		{
		return false;
		}
	}		
// -->
</script>	
					<%
						If GetStaff.EOF OR GetStaff.BOF Then
					%>
					<tr>
	                <td colspan="8" class="formMainBold">No Staff Members To List</td>
    		   		</tr>
					<%
						Else
						GetStaff.MoveFirst
						Do Until GetStaff.EOF
					 %>
<!-- first row of results -->


			<tr>
				<td rowspan="3" class="formMain"><%= ct %></td>
				<td class="formMain" align="center"><%= GetStaff("BirthYear") %></td>
					<% 
					query = "SELECT position FROM tbl_StaffPosition WHERE code=" & Int(GetStaff("position"))
					Set GetCode = Con.Execute(query)
					 %>
					<td class="formMain" align="center">
					<% If GetCode.EOF OR GetCode.BOF Then %>
						<i>Unlisted</i>
					<% else %> 
						<%= GetCode("position") %>
					<% end if %>
					</td>
					<% 
					GetCode.Close
					Set GetCode = Nothing
					 %>
					
					<% 
					query = "SELECT race FROM tbl_StaffRace WHERE code=" & Int(GetStaff("race"))
					Set GetCode = Con.Execute(query)
					 %>
				<td class="formMain" align="center"><%= GetCode("race") %></td>
					<% 
					GetCode.Close
					Set GetCode = Nothing
					 %>
				  
				<td class="formMain" align="center"><%= UCase(GetStaff("sex")) %></td>
				
<% 'added 12/21/2006 saf per Stu "add Employee Name"  %>
				<td class="formMainRightJ" colspan=2>&nbsp;<%If IsNull(GetStaff("basesalary")) Then%><%= FormatCurrency(GetStaff("yearlysalary")) %><%Else%><%= FormatCurrency(GetStaff("basesalary")) %><%End If%>&nbsp;</td>
				
				<% if printform="No" then %>
				<td align="center" class="formMain" rowspan="3"><a href="staff_edit.asp?status=editRow&row=<%= GetStaff("StaffID") %>&y=<%= Request("y") %>">Edit Record</a><br></td>
		 <!--   row version prior too 2008 AAI. Remove Delete function per Diane 6/19/2008			<td align="right" class="formMain" rowspan="3"><a href="staff_edit.asp?status=editRow&row=<%= GetStaff("StaffID") %>&y=<%= Request("y") %>">Edit Record</a><br><br><a href="#" onclick="confirmDelete(<%= GetStaff("StaffID") %>);return false;">Delete Record</a></td>
		 <!--	row version prior to 2007		<td align="right" class="formMain" rowspan="3"><a href="staff_edit.asp?status=editRow&row=<%= GetStaff("StaffID") %>&y=<%= Request("y") %>">Edit Record</a><br><br><a href="staff_edit.asp?status=deleteRow&row=<%= GetStaff("StaffID") %>&y=<%= Request("y") %>">Delete Record</td> -->
				<% end if %>
			</tr>	
			
			<tr>
				<td class="formMain" align="center"><%= (GetStaff("Monthstart")) & "/" & GetStaff("yearstart") %></td>
				<td class="formMain" align="center">&nbsp;<% If not ISNULL(GetStaff("PositionStartDate")) Then Response.Write(FormatDateTime(GetStaff("PositionStartDate"),2)) %>&nbsp;</td>
				<td class="formMain" align="center"><% If GetStaff("monthend") = 0 Then %>Still Employed<% Else %><%= MonthName(GetStaff("monthend")) %><% End If %></td>
				<td class="formMain" align="center"><%If IsNull(GetStaff("hoursweek")) Then%>N/A<%Else%><%= GetStaff("hoursweek")%><%End If%></td>
				<td colspan="2" class="formMainRightJ"><%If IsNull(GetStaff("bonussalary")) Then%>0<%Else%><%= FormatCurrency(GetStaff("bonussalary"))%><%End If%>&nbsp;</td>				
			</tr>
			
			<tr>
				<td class="formMain" align="center"><%If IsNull(GetStaff("YearsInNetwork")) Then%>N/A<%Else%><%= (GetStaff("YearsInNetwork"))%><%End If%></td>
				<td class="formMain" align="center">&nbsp;<%If IsNull(GetStaff("EmployeeName")) Then%>N/A<%Else%><%= (GetStaff("EmployeeName"))%><%End If%>&nbsp;</td>
				<% query = "SELECT education FROM tbl_StaffEducation WHERE code=" & Int(GetStaff("Education"))
					Set GetCode = Con.Execute(query)
					 %>
				<td class="formMain" align="center"><%= GetCode("Education") %></td>
					<% 
					GetCode.Close
					Set GetCode = Nothing
					 %>
              <td class="formMain" align="center"><%If IsNull(GetStaff("EverABIG")) Then%>N/A<%ElseIf(GetStaff("EverABIG")) = 1 Then %>Yes<%Else%>No<%End If%></td>
				<td colspan="2" class="formMainRightJ"><%= FormatCurrency(GetStaff("yearlysalary")) %>&nbsp;</td>				
			</tr>
			
			<tr>
                <td colspan="8" class="formHeader"><img src="../images/spacer.gif" width="1" height="5" alt="" border="0"></td>
       		</tr>

					<% 
							GetStaff.MoveNext
							ct = ct + 1
						Loop
						GetStaff.Close
						Set GetStaff = Nothing
						Con.Close
						Set Con = Nothing
					End If
					 %>

			<% if printform="No" then %>
								 
			<form name="frmStaff" action="staff_edit.asp?y=<%= Request("y") %>" method="post">
			<!--#include file="../includes/form_stamp.asp"-->
			<input type="hidden" name="status" value="newStaff">
						<tr>
			                <td colspan="8" class="formHeader"><input type="submit" value="Add Additional Staff Member" class="formMainBold"></td>
			       		</tr>
						
						<tr>
							<td colspan="8"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
						</tr>
			</form>
			
			<% end if %>
					 
		</table>