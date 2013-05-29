<%@ Language=VBScript %>
<%
	' COPYRIGHT 2003 Philip Dearmore, all rights reserved.
	' The HTML, Javascript, and writing on this page are all copyrighted and may not be
	' redistributed in any form except by ASPTools.Biz, or any of the web sites that
	' ASPTools.Biz uploads to for distribution.
%>
<HTML>
<HEAD>
<TITLE>ASPTools.Biz ASP Date Picker Component</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">


<SCRIPT LANGUAGE=javascript>
<!--

// OpenCalendar
// The OpenCalendar function can be used as a shortcut to the actual window.open statement that 
// ultimately must be issued.  OpenCalendar includes a rudimentary popup blocker detection
// system that unfortunately won't always work because of a wide variety of methods employed
// to block popup windows.  You may want to inform your guests that popup blockers should be
// disabled for your site for proper functioning.

function OpenCalendar(inVariable, inField, inCallback, inDefault, inDayStyle, inBlnClose, inBlnWeekends)
{
	var strURL;
	var hWnd; 
	
	if (inVariable) strURL = "inVariable=" + inVariable;
	else if (inField) strURL = "inField=" + inField;
	else if (inCallback) strURL = "inCallback=" + inCallback;
	else strURL = "invalid=True";
	
	if (inDefault) strURL += "&inDefault=" + inDefault;
	if (inBlnClose) strURL += "&inBlnClose=" + inBlnClose;
	if (inBlnWeekends) strURL += "&inBlnWeekends=" + inBlnWeekends;
	if (inDayStyle) strURL += "&inDayStyle=" + inDayStyle;
	
	strURL += "&inFormat=" + document.forms[0].selFormat.value;
	strURL += "&inDelimiter=" + document.DatePicker.txtDelim.value;
	
//-->	hWnd = window.open( "ASPDatePicker.asp?" + strURL, "ASPDatePicker",
//-->		"menubar=no,toolbar=no,location=no,scrollbars=no,resizable=no,status=no,width=210,height=150");

	hWnd = window.open( "ASPDatePicker.asp?" + strURL, "ASPDatePicker");

		
	if (!hWnd) alert("Please disable popup blockers to use the Date Picker calendar.");
}

//-->
</SCRIPT>

</HEAD>
<BODY topmargin="0" leftmargin="2" rightmargin="2">
<FORM name="DatePicker">
<TABLE border="0" cellpadding="5" cellspacing="0">
	<TR>
		<TD colspan="3" valign="top">
			<DIV class="label">
				ASP Date Picker
			</DIV>
			<DIV width="100%">
					<TABLE border="0" cellpadding="2" cellspacing="0">
						<TR><TD valign="top">
				(Runs on an IIS Server, compatible w/ NS6+, IE4+ and Opera 6+)
				<P>
							The ASPTools ASP Date Picker is a DHTML calendar presented in a pop-up window
							on demand from code in your application.  It is normally called from a
							window.open() Javascript command.  Options are passed in the querystring.
							After the calendar is opened, it communicates
							the selected date back to the calling page in whichever of three methods makes
							the most sense from <I>your</I> standpoint.  The three methods are illustrated
							below.  You can use the "View Source" option to see how these are accomplished.
						</TD>
						<TD valign="top" align="right">
							<B>Date Format</B><BR><BR>
							DatePicker defaults to mm/dd/yyyy format.  You can also choose
							from many other format combinations with a custom delimiter.  Set the
							format here for each of the three methods below.<BR><BR>
							<NOBR>Format:
							<SELECT name="selFormat">
								<OPTION value="0">mm dd yyyy</OPTION>
								<OPTION value="1">mm dd yy</OPTION>
								<OPTION value="2">dd mm yyyy</OPTION>
								<OPTION value="3">dd mm yy</OPTION>
								<OPTION value="4">yyyy mm dd</OPTION>
								<OPTION value="5">yyyy dd mm</OPTION>
								<OPTION value="6">yy dd mm</OPTION>
								<OPTION value="7">yy mm dd</OPTION>
							</SELECT>
							</NOBR>
							<NOBR>
							Delimiter:
							<INPUT type="text" value="/" name="txtDelim" size="2" maxlength="1">
							</NOBR>
						</TD>
						</TR>
					</TABLE>
				</P>
			</DIV>
		</TD>
	</TR>
	<TR>
		<TD valign="top" width="33%">
			<DIV class="label">Method One</DIV>
			<DIV>
				<strong>Javascript Variable</strong>
				<P>
					You can create a Javascript variable and pass it to the ASP Date Picker
					in the inVariable querystring parameter.  In the example below, the value
					of the bottom button is set by the strDate variable.  If you choose a 
					date, then click the Update button, you will see it has been updated with
					the last date selected.  You can do this multiple times.
				</P>
				<CENTER>
					<SCRIPT LANGUAGE=javascript>
					<!--
						var strDate = "Click ^ First";
					//-->
					</SCRIPT>
					<INPUT type="Button" value="Open Calendar" onClick="OpenCalendar('strDate');" style="width:100px;font-size:8pt;">
					<BR>
					<INPUT type="Button" value="Update" name="VarDate" onClick="this.value = strDate;" style="width:100px;font-size:8pt;">
				</CENTER>
			</DIV>
		</TD>
		<TD valign="top" width="33%">
			<DIV class="label">Method Two</DIV>
			<DIV>
				<strong>Form Field</strong>
				<P>
					An HTML form field can be passed in the inField querystring parameter.  The
					field will be set to the selected date. The name of the field should be passed
					as a string, preceeded by the name of the form (in this case "DatePicker.FldDate").
				</P>
				<CENTER>
					<INPUT type="Button" value="Open Calendar" onClick="OpenCalendar(null,'DatePicker.FldDate');" style="width:100px;font-size:8pt;">
					<BR>
					<INPUT type="Text" name="FldDate" style="width:100px;font-size:8pt;">
				</CENTER>
			</DIV>
		</TD>
		<TD valign="top" width="33%">
			<DIV class="label">Method Three</DIV>
			<DIV>
				<strong>Callback Function</strong>
				<P>
					The name of a function can be passed in the inCallback querystring
					parameter.  This function will be called with the selected date as an argument.
					This has potential to give you more
					flexibility with the results of a chosen date.  The function below passes
					the callback, then executes code to change the status bar.  Notice that the
					calendar stays open when selected a date... This isn't necessary to the 
					Callback method&mdash;this is just to demonstrate that the calendar can remain 
					open by selecting the appropriate option.
				</P>
				<CENTER>
					<SCRIPT LANGUAGE=javascript>
					<!--
						function CB (strDate)
						{
							window.status = "The date selected was " + strDate;
						}
					//-->
					</SCRIPT>
					<INPUT type="Button" value="Open Calendar" onClick="OpenCalendar(null, null, 'CB', null, null, 'False');" style="width:100px;font-size:8pt;">
				</CENTER>
			</DIV>
		</TD>
	</TR>	
	<TR>
		<TD colspan="3">
			<DIV class="label">
				The "Non-Popup-Window" (IFrame) Method
			</DIV>
			<DIV width="100%">
				<TABLE border="0" cellpadding="0" cellspacing="0">
					<TR>
						<TD valign="top" align="center">
							<IFRAME width="210" height="150" src="ASPDatePicker.asp?inField=DatePicker.FldIframe&inBlnIframe=True" frameborder="0" allowtransparency="true" scrolling="no">
								<P align="center">This IFRAME tag displays an inline Date Picker calendar, but your browser does not support IFRAMES.</P> 
							</IFRAME>
							<BR>
							<INPUT type="text" name="FldIframe" style="width:100px;font-size:8pt;text-align:center;">
						</TD>
						<TD valign="top" style="padding-left: 35px;">
							<BR><BR>
							Using IFrames, supported by IE3+, NS6+ and Opera 4+, you can embed the
							ASPTools Date Picker Calendar in a web page.  This is simply a matter of
							inserting the proper IFrame tag (you can view the source of the tag to
							the left) and adding an extra parameter to the URL that calls the
							Calendar that looks like "inBlnIframe=True".  The IFRAME method supports
							other date formats as well by setting the inFormat variable in the src
							attribute.
						</TD>
					</TR>
				</TABLE>
			</DIV>
		</TD>
	</TR>
</TABLE>
</FORM>
<BR><BR>
</BODY>
</HTML>
