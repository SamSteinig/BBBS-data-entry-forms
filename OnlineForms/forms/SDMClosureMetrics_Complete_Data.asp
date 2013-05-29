<script language="JavaScript">
<!-- 
	function confirmDelete()
	{
		if (confirm("Are you sure you want to delete this record?"))
		{
			location = "SDMClosureMetrics_Delete.asp?row=" + row + "&AgencyID=<%=Session("AgencyIDN")%>"
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

	Dim SortField
	SortField = request("SortField")
	
	Dim SortDirection
	SortDirection = request("SortDirection")
	
	Dim MatchIDSearch
	MatchIDSearch = request("MatchIDSearch")
	
%>

<!-- RESULTS TABLE STARTS HERE -->

			<table border="1" cellpadding="2" cellspacing="0" bordercolordark="003063" width="725">

			<tr>
				<td colspan="7" align="center" class="formSubhead">Monthly Performance</td>
			</tr>
			<tr>
				<td colspan="7" align="center" class="formHeader">SDM Closure Metrics</td>
			</tr>
			<tr>
				<td colspan="4" bgcolor="#ffffff" class="formSubhead">Current Sort Column:&nbsp;<%=SortField%><br>Current Sort Direction:&nbsp;<% if SortDirection="ASC" then %>Ascending<%else%>Descending<%end if%><br>To change sort order, click on the up or down arrows next to the column titles.</td>
				<td colspan="3" align="right" valign="bottom" class="formMain">
				
				<form name="frmSDMClosureMetricsSearch" action="SDMClosureMetrics_complete.asp?SortField=<%=SortField%>&SortDirection=<%=SortDirection%>" method="post">
					&nbsp;Match ID:&nbsp;<input type="text" name="MatchIDSearch" size="20" value="">
					<input type="submit" value="Search" class="formMainBold">
				</form>
				</td>
			</tr>		

			


			
<!-- first row of table headers -->
			<tr>

				<td class="formHeaderSmall">#</td>
				<td class="formHeaderSmall">
				Match ID:&nbsp;
				<a href="SDMClosureMetrics_Complete.asp?SortField=MatchID&SortDirection=ASC"><img src="../images/sort_up.gif" alt="" width="9" height="5" border="0"></a>
				<a href="SDMClosureMetrics_Complete.asp?SortField=MatchID&SortDirection=DESC"><img src="../images/sort_down.gif" alt="" width="9" height="5" border="0"></a>								
				</td>
				<td class="formHeaderSmall">
				Match Open Date:&nbsp;
				<a href="SDMClosureMetrics_Complete.asp?SortField=MatchStartDate&SortDirection=ASC"><img src="../images/sort_up.gif" alt="" width="9" height="5" border="0"></a>
				<a href="SDMClosureMetrics_Complete.asp?SortField=MatchStartDate&SortDirection=DESC"><img src="../images/sort_down.gif" alt="" width="9" height="5" border="0"></a>				
				</td>
				<td class="formHeaderSmall">
				Match Close Date:&nbsp;
				<a href="SDMClosureMetrics_Complete.asp?SortField=MatchEndDate&SortDirection=ASC"><img src="../images/sort_up.gif" alt="" width="9" height="5" border="0"></a>
				<a href="SDMClosureMetrics_Complete.asp?SortField=MatchEndDate&SortDirection=DESC"><img src="../images/sort_down.gif" alt="" width="9" height="5" border="0"></a>								
				</td>
				<td class="formHeaderSmall">
				Match Type:&nbsp;
				<a href="SDMClosureMetrics_Complete.asp?SortField=MatchType&SortDirection=ASC"><img src="../images/sort_up.gif" alt="" width="9" height="5" border="0"></a>
				<a href="SDMClosureMetrics_Complete.asp?SortField=MatchType&SortDirection=DESC"><img src="../images/sort_down.gif" alt="" width="9" height="5" border="0"></a>								
				</td>
				<td class="formHeaderSmall" colspan="2">Edit/Delete</td>
			</tr>
			
			

	<%
	
		ct = 1
		Set Con = Server.CreateObject("ADODB.Connection")
		Con.Open "BBBSAforms", "sa","12sist12"
		if MatchIDSearch = "" then
			query = "SELECT * FROM tbl_frmSDMClosureMetrics WHERE AgencyID='" & Session("AgencyIDN") & "' ORDER BY " & SortField & " " & SortDirection
		else
			query = "SELECT * FROM tbl_frmSDMClosureMetrics WHERE AgencyID='" & Session("AgencyIDN") & "' AND MatchID LIKE '%" & MatchIDSearch & "%'"
		end if		
		Set GetSDMClosureMetrics = Con.Execute(query)
	%>

	<%
		If GetSDMClosureMetrics.EOF OR GetSDMClosureMetrics.BOF Then
	%>
		<tr>
             <td colspan="8" class="formMainBold"><br>&nbsp;No Closure Metrics Records to List.  <% if MatchIDSearch <> "" then %><a href="SDMClosureMetrics_complete.asp?SortField=<%=SortField%>&SortDirection=<%=SortDirection%>">Cancel Search</a><% end if %><br>&nbsp;</td>
		</tr>
	<%
		Else 
		GetSDMClosureMetrics.MoveFirst
		Do Until GetSDMClosureMetrics.EOF
	 %>
<!-- first row of results -->


			<tr>
				<td class="formMain"><%= ct %></td>
				<td class="formMain" align="center"><%= GetSDMClosureMetrics("MatchID") %></td>
				<td class="formMain" align="center"><%= GetSDMClosureMetrics("MatchStartDate") %></td>
				<td class="formMain" align="center"><%= GetSDMClosureMetrics("MatchEndDate") %></td>				
				<td class="formMain" align="center">
				
				<%
				Select Case GetSDMClosureMetrics("MatchType")
					Case 1
						MatchType = "Community"
					Case 2
						MatchType = "School"
					Case 3 
						MatchType = "Other Site"
				End Select
				%>	
				
				<%= MatchType %>
				
				
				</td>								


	 			<td align="right" class="formMain"  valign="middle" width="60">
					<form name="frmSDMClosureMetricsEdit" action="SDMClosureMetrics_edit.asp?row=<%=GetSDMClosureMetrics("SDMClosureMetricsID")%>&AgencyID=<%=Session(AgencyIDN)%>&y=0&SortField=<%=SortField%>&SortDirection=<%=SortDirection%>" method="post">
						<input type="hidden" name="status" value="editRow">
						<input type="submit" value="Edit" class="formMainBold">
					</form>	
				</td>
				<td align="right" class="formMain" valign="middle" width="60">
			       
					<form name="frmSDMClosureMetricsDeleteConfirm" action="SDMClosureMetrics_delete_confirm.asp?row=<%=GetSDMClosureMetrics("SDMClosureMetricsID")%>&AgencyID=<%=Session(AgencyIDN)%>&MatchID=<%= GetSDMClosureMetrics("MatchID")%>&SortField=<%=SortField%>&SortDirection=<%=SortDirection%>" method="post">
						<input type="submit" value="Delete" class="formMainBold">
					</form>
				</td>
			</tr>	
			
			
			
			<tr>
                <td colspan="8" class="formHeader"><img src="../images/spacer.gif" width="1" height="5" alt="" border="0"></td>
       		</tr>

					<% 
							GetSDMClosureMetrics.MoveNext
							ct = ct + 1
						Loop
						GetSDMClosureMetrics.Close
						Set GetSDMClosureMetrics = Nothing
						Con.Close
						Set Con = Nothing
					End If
					 %>

			<% if printform="No" then %>
								 
			<form name="frmSDMClosureMetrics" action="SDMClosureMetrics_edit.asp?y=0&SortField=<%=SortField%>&SortDirection=<%=SortDirection%>" method="post">
			<!--#include file="../includes/form_stamp.asp"-->
			<input type="hidden" name="status" value="newStaff">
						<tr>
			                <td colspan="8" class="formHeader"><input type="submit" value="Add New Closure Metrics Record" class="formMainBold"></td>
			       		</tr>
						
						<tr>
							<td colspan="8"><div align="center"><!--#include file="../includes/contact_info.inc"--></div></td>
						</tr>
			</form>
			
			<% end if %>
					 
		</table>