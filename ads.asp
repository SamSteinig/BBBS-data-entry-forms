<%@LANGUAGE="VBSCRIPT"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<!--#include file="../Connections/NAD_BE.asp" -->
<!--#INCLUDE FILE="../security/passwordpro/check_user_inc.asp"-->
<%

dim referrer
referrer = "../surveys/ads.asp"
%>
<html>
<head>
	<title>BBBS :: Surveys & Reports: Program Evaluation Surveys</title>
	
	<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
	<!--
	
	if(window.event + "" == "undefined") event = null;
	function HM_f_PopUp(){return false};
	function HM_f_PopDown(){return false};
	popUp = HM_f_PopUp;
	popDown = HM_f_PopDown;
	
	//-->
	</SCRIPT>
	
	<SCRIPT LANGUAGE="JavaScript1.2" SRC="../media/scripts/HM_Loader.js" TYPE='text/javascript'></SCRIPT>
	
	<SCRIPT LANGUAGE="JavaScript1.2" SRC="../media/scripts/tool_tip.js" TYPE='text/javascript'></SCRIPT>
	
	<LINK rel=STYLESHEET href = "../media/scripts/bbbsa.css" Type = "text/css">
	
	<!--#include file="../media/inc/mouseover.inc"-->
	
	<!-- stupid number generator script -->
	<SCRIPT LANGUAGE="JavaScript1.2">
	<!--
	var numcount = 6
	day = new Date()
	seed = day.getTime()
	ran = parseInt(((seed - (parseInt(seed/1000,10) * 1000))/10)/100*numcount + 1,10)
	
	if (ran == (1))    
	count=("1") 
	if (ran == (2))
	count=("2") 
	if (ran == (3))
	count=("3")
	if (ran == (4))
	count=("4")
	if (ran == (5))
	count=("5")
	if (ran == (6))
	count=("6")
	// -->
	</SCRIPT>
	
</head>

<BODY BACKGROUND="../media/images/bground.gif" BGCOLOR="#ffffff" LINK="#00368f" ALINK="#ff9e11" VLINK="#00368f" TOPMARGIN="0" LEFTMARGIN="0" MARGINHEIGHT="0" MARGINWIDTH="0"><!-- header and menubar --><!--#include file="../media/inc/menu_home.inc"--><!-- table for shadow -->
<TABLE WIDTH="100%" BORDER="0" CELLPADDING="0" CELLSPACING="0">
	<TR>
		<TD WIDTH="10%"><IMG height=6 src="../media/images/shadow_piece1.gif" width=177 border=0></TD>
		<TD BACKGROUND="../media/images/shadow_piece2.gif" WIDTH="90%"><IMG height=1 src="../media/images/spacer.gif" width=1 border=0></TD>
	</TR>
</TABLE><!-- main table -->
<TABLE WIDTH="100%" BORDER="0" CELLPADDING="0" CELLSPACING="0">
	<TR><!-- left side content cell -->
		<TD VALIGN="top" WIDTH="10%">
		<IMG height=10 src="../media/images/spacer.gif" width=1 border=0><BR>
		
			<TABLE WIDTH="100%" BORDER="0" CELLPADDING="4" CELLSPACING="0">
				<TR BGCOLOR="#bb9d1c">
					<TD VALIGN="center" ALIGN="right" WIDTH="10%"></TD>
					<TD VALIGN="center" WIDTH="90%" CLASS="menutitle">Related Links</TD>		
				</TR>
				
				<TR>
					<TD VALIGN="center" ALIGN="right" WIDTH="10%"></TD>
					<TD VALIGN="center" WIDTH="90%"><IMG height=8 src="../media/images/sm_arrow.gif" width=4 border=0>&nbsp;<A class=cool href="./ads.asp">ADS Survey</A></TD>		
				</TR>
				
				<TR>
					<TD VALIGN="center" ALIGN="right" WIDTH="10%"></TD>
					<TD VALIGN="center" WIDTH="90%"><IMG height=4 src="../media/images/sm_arrow2.gif" width=8 border=0>&nbsp;<SPAN class=menupick>Program Evaluation Surveys</SPAN></TD>		
				</TR>
				
			</TABLE>
			
			
			<BR><!--#include file="../media/inc/community_tools.inc"-->
			
		</TD><!-- end left side content cell --><!-- spacer cell -->
		<TD VALIGN="top" WIDTH="1%"><IMG height=1 src="../media/images/spacer.gif" width=12 border=0></TD><!-- main content cell -->
		<TD VALIGN="top" WIDTH="88%"><!-- search table -->
			<TABLE WIDTH="549" BORDER="0" CELLPADDING="0" CELLSPACING="0">
				<TR>
					<TD VALIGN="top"><IMG height=10 src="../media/images/spacer.gif" width=1 border=0></TD>		
				</TR>
				
				<TR>
					<TD VALIGN="top"><IMG height=55 alt="Collecting Information and Sharing Experiences" src="../media/images/header_surveys.gif" width=553 border=0 ></TD>		
				</TR>
				
				<TR>
					<TD VALIGN="top"><IMG height=10 src="../media/images/spacer.gif" width=1 border=0></TD>		
				</TR>
				
				<TR>
					<TD VALIGN="top" CLASS="text">
						<P><FONT face="Arial" size="3"><b><u>Please Note: These Pages Have Moved!</u></b></FONT></P>
						<P><a href="../MyAgency/ads.asp" class="topstory">The ADS/Monthly Surveys Have Moved to Your MY AGENCY SECTION. Please Update Your Bookmarks!</a></P>
					</TD>		
				</TR>
				
				<TR>
					<TD VALIGN="top"><IMG height=20 src="../media/images/spacer.gif" width=1 border=0></TD>		
				</TR>				
			</TABLE>					
		</TD><!-- end main content cell --><!-- spacer cell -->
		<TD VALIGN="top" WIDTH="1%"><IMG height=1 src="../media/images/spacer.gif" width=5 border=0></TD>		
	</TR>
</TABLE>

</BODY>
</html>
