<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../../Connections/DBConn2.asp" -->
<%
Dim rsNews
Dim rsNews_numRows

Set rsNews = Server.CreateObject("ADODB.Recordset")
rsNews.ActiveConnection = MM_DBConn2_STRING
rsNews.Source = "SELECT *  FROM dbo.News"
rsNews.CursorType = 0
rsNews.CursorLocation = 2
rsNews.LockType = 1
rsNews.Open()

rsNews_numRows = 0
%>
<!-- AspUpload Code samples: UploadScript3.asp -->
<!-- Invoked by Form3.asp -->
<!-- Copyright (c) 2001 Persits Software, Inc. -->
<!-- http://www.persits.com -->


<HTML>
<BODY BGCOLOR="#FFFFFF">
<%
	Set Upload = Server.CreateObject("Persits.Upload")

	' Uncomment this line if unique file name generation is necessary
	' Upload.OverwriteFiles = False 

	' We must call Upload.Save or SaveVirtual before we can use Upload.Form!
	Upload.Save "D:\Inetpub\wwwroot\SieradNew\ASP\www\ADMIN\img"
%>

<h3>Upload Results</H3>
<TABLE CELLSPACING=0 CELLPADDING=2 BORDER=1>
<TR><TH>Form Item</TH><TH>Value(s)</TH></TR>
	<TR>
		<TD>Form_text</TD>
		<TD>
		<% = Upload.Form("p_title") %>&nbsp;
		</TD>
	</TR>
	<TR>
		<TD>Form_check</TD>
		<TD>
		<% = Upload.Form("p_lead") %>
		</TD>
	</TR>
	<TR>
		<TD>Form_select</TD>
		<TD>
		<% = Upload.Form("p_content") %>&nbsp;
		</TD>
	</TR>
	<TR>
		<TD>Form_radio</TD>
		<TD>
		<%
		
		'SUB Add()	
		dim id,thumbnail,image,clip, head,content
		dim nomor
		'id = Upload.Form("h_id") 
		'nomor = 1
		login = Session("updateby")
		
		Set File = Upload.Files("p_thumbnail")
		If File Is Nothing Then
		   thumbnail = "../../ADMIN/img/none.gif"
		   Response.Write "p_thumbnail KOSONG" & "<BR>"
		else
			thumbnail= "../../ADMIN/img/" & File.filename
			Response.Write "p_thumbnail ISI" & "<BR>"
		End If 

		Set File = Upload.Files("p_picture")
		If File Is Nothing Then
		   image = "../../ADMIN/img/none.gif"
		   Response.Write "p_image KOSONG" & "<BR>"
		else
			image= "../../ADMIN/img/" & File.filename
			Response.Write "p_image ISI" & "<BR>"
		End If 

		
		'if Session("Langu")="INA" then				
		''SQL = "Update Activities Set head = '"&Upload.Form("p_title")&"', clip = '"&Upload.Form("p_lead")&"', content = '"&Upload.Form("p_content")&"', thumbnail = '"&thumbnail&"', image = '"&image&"'" 
  		''SQL = SQL & " Where id = '"&id&"'" 
			SQL = "Insert Into News (updateby, title, clip, content, thumbnail, image,Lang) Values('"&Login&"', '"&Upload.Form("p_title")&"', '"&Upload.Form("p_lead")&"', '"&Upload.Form("p_content")&"', '"&thumbnail&"', '"&image&"','"&Upload.Form("sLang")&"')"
			''response.write SQL
			Set rsNews = Server.CreateObject("ADODB.Recordset")
			rsNews.ActiveConnection = MM_DBConn2_STRING
			rsNews.Source = SQL
			rsNews.CursorType = 0
			rsNews.CursorLocation = 2
			rsNews.LockType = 1
			rsNews.Open()
			rsNews_numRows = 0
			'response.redirect "br_News.asp"
		'Else
		'	SQL = "Insert Into EnNews (updateby, head, clip, content, thumbnail, image) Values('"&Login&"', '"&Upload.Form("p_title")&"', '"&Upload.Form("p_lead")&"', '"&Upload.Form("p_content")&"', '"&thumbnail&"','"&image&"')"
		'	'response.write SQL
			'Set rsEnNews = Server.CreateObject("ADODB.Recordset")
			'rsEnNews.ActiveConnection = MM_AIConnect_STRING
			'rsEnNews.Source = SQL
			'rsEnNews.CursorType = 0
			'rsEnNews.CursorLocation = 2
			'rsEnNews.LockType = 1
			'rsEnNews.Open()			
			'rsEnNews_numRows = 0
			'response.redirect "En_br_News.asp"
		'End if
		'end sub
		
		
			
		'call Add()
		response.redirect "br_News.asp"
		'response.Write sql
		
 %>
      &nbsp;
		</TD>
	</TR>
	<TR>
		<TD>Form_file</TD>
		
    <TD>&nbsp; </TD>
	</TR>
</TABLE>

</BODY>
</HTML>
<%
rsNews.Close()
Set rsNews = Nothing
%>
