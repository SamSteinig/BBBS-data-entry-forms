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

<!-- first row of table headers -->
			<tr>
			<% if printform="No" then %>		
				<td rowspan="2" class="formHeaderSmall">#</td>
				<td class="formHeaderSmall">Birth Year:</td>
				<td class="formHeaderSmall">Position:</td>
				<td class="formHeaderSmall">Race:</td>
				<td class="formHeaderSmall">Gender:</td>
				<td rowspan="2" class="formHeaderSmall" colspan="2">Compensation<br>(Salary+Bonus/Incentives)</td>
				<td rowspan="2" class="formHeaderSmall">Edit/Delete</td>
			<% else %>
				<td rowspan="2" class="formMainBold"><div align="center">#</div></td>
				<td class="formMainBold"><div align="center">Birth Year:</div></td>
				<td class="formMainBold"><div align="center">Position:</div></td>
				<td class="formMainBold"><div align="center">Race:</div></td>
				<td class="formMainBold"><div align="center">Gender:</div></td>
				<td rowspan="2" class="formMainBold" colspan="2"><div align="center">Compensation<br>(Salary+Bonus/Incentives)</div></td>		
			<% end if %>
			</tr>
<!-- second row of table headers -->
			<tr>
			<% if printform="No" then %>			
				<td class="formHeaderSmall">Month Start:</td>
				<td class="formHeaderSmall">Year Start:</td>
				<td class="formHeaderSmall">Month End:</td>
				<td class="formHeaderSmall">Hrs/Wk:</td>
			<% else %>
				<td class="formMainBold"><div align="center">Month Start:</div></td>
				<td class="formMainBold"><div align="center">Year Start:</div></td>
				<td class="formMainBold"><div align="center">Month End:</div></td>
				<td class="formMainBold"><div align="center">Hrs/Wk:</div></td>
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
				function confirmDelete()
	{
		if (confirm("Are you sure you want to delete this record?"))
		{
			location = "staff_edit.asp?status=deleteRow&row=<%= GetStaff("StaffID") %>&y=<%= Request("y") %>";
			alert("Record deleted.");
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
				<td rowspan="2" class="formMain"><%= ct %></td>
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
				<td class="formMain" align="center" colspan="2">&nbsp;</td>

				<% if printform="No" then %>
		<!-- 			<td align="right" class="formMain" rowspan="2"><a href="staff_edit.asp?status=editRow&row=<%= GetStaff("StaffID") %>&y=<%= Request("y") %>">Edit Record</a><br><a href="#" onclick="confirmDelete();return false;">Delete Record</a></td>				-->
		 			<td align="right" class="formMain" rowspan="2"><a href="staff_edit.asp?status=editRow&row=<%= GetStaff("StaffID") %>&y=<%= Request("y") %>">Edit Record</a><br><a href="staff_edit.asp?status=deleteRow&row=<%= GetStaff("StaffID") %>&y=<%= Request("y") %>">Delete Record</td>
				<% end if %>
			</tr>	
			
			<tr>
				<td class="formMain" align="center"><%= MonthName(GetStaff("Monthstart")) %></td>
				<td class="formMain" align="center"><%= GetStaff("yearstart") %></td>
				<td class="formMain" align="center"><% If GetStaff("monthend") = 0 Then %>Still Employed<% Else %><%= MonthName(GetStaff("monthend")) %><% End If %></td>
				<td class="formMain" align="center"><%= GetStaff("hoursweek") %></td>
				<td colspan="2" class="formMainRightJ"><%= FormatCurrency(GetStaff("yearlysalary")) %></td>				
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