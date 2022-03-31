<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%> 
<!--#include file="../../Connections/simConn.asp" -->
<%
Dim rsArea__MMColParam
rsArea__MMColParam = "1"
If (Request.Form("ID_Area") <> "") Then 
  rsArea__MMColParam = Request.Form("ID_Area")
End If
%>
<%
Dim rsArea
Dim rsArea_numRows

Set rsArea = Server.CreateObject("ADODB.Recordset")
rsArea.ActiveConnection = MM_simConn_STRING
rsArea.Source = "SELECT * FROM dbo.Sales_Group WHERE ID_Area = '" + Replace(rsArea__MMColParam, "'", "''") + "'"
rsArea.CursorType = 0
rsArea.CursorLocation = 2
rsArea.LockType = 1
rsArea.Open()

rsArea_numRows = 0
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
rsArea_numRows = rsArea_numRows + Repeat2__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsArea_total
Dim rsArea_first
Dim rsArea_last

' set the record count
rsArea_total = rsArea.RecordCount

' set the number of rows displayed on this page
If (rsArea_numRows < 0) Then
  rsArea_numRows = rsArea_total
Elseif (rsArea_numRows = 0) Then
  rsArea_numRows = 1
End If

' set the first and last displayed record
rsArea_first = 1
rsArea_last  = rsArea_first + rsArea_numRows - 1

' if we have the correct record count, check the other stats
If (rsArea_total <> -1) Then
  If (rsArea_first > rsArea_total) Then
    rsArea_first = rsArea_total
  End If
  If (rsArea_last > rsArea_total) Then
    rsArea_last = rsArea_total
  End If
  If (rsArea_numRows > rsArea_total) Then
    rsArea_numRows = rsArea_total
  End If
End If
%>

<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsArea_total = -1) Then

  ' count the total records by iterating through the recordset
  rsArea_total=0
  While (Not rsArea.EOF)
    rsArea_total = rsArea_total + 1
    rsArea.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsArea.CursorType > 0) Then
    rsArea.MoveFirst
  Else
    rsArea.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsArea_numRows < 0 Or rsArea_numRows > rsArea_total) Then
    rsArea_numRows = rsArea_total
  End If

  ' set the first and last displayed record
  rsArea_first = 1
  rsArea_last = rsArea_first + rsArea_numRows - 1
  
  If (rsArea_first > rsArea_total) Then
    rsArea_first = rsArea_total
  End If
  If (rsArea_last > rsArea_total) Then
    rsArea_last = rsArea_total
  End If

End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>::Pencarian ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.style1 {font-size: 18px}
.style4 {font-size: 14px}
-->
</style>
<script language="JavaScript" type="text/JavaScript">
<!--

function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
<link href="../../Capex/css/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style5 {color: #FFFFFF}
.style7 {color: #FFFFFF; font-weight: bold; }
.style9 {
	font-family: verdana;
	font-size: 12px;
	font-weight: bold;
}
-->
</style>
</head>

<body bgcolor="#FFFFFF" background="../../Capex/Image/bg.gif">
<div align="center"> 
  <table width="756" border="0">
    <tr> 
      <td width="750" colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="178"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="178" height="93">
                <param name="movie" value="../../img/logo.swf">
                <param name="quality" value="high">
                <embed src="../../img/logo.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="178" height="93"></embed></object></td>
            <td width="150" background="../../img/bg_top2.jpg">&nbsp;</td>
            <td width="211" background="../../img/bg_top.jpg">&nbsp;</td>
            <td background="../../img/bg_top3.jpg">&nbsp;</td>
            <td width="469"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="469" height="93">
                <param name="movie" value="../../img/tagline.swf">
                <param name="quality" value="high">
                <embed src="../../img/tagline.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="469" height="93"></embed></object></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td colspan="2"><h3 align="center"><font color="#6699FF">.:: Pencarian ::. 
          </font></h3></td>
    </tr>
    <tr>
      <td colspan="2">
<form name="form1" method="post" action="SearchIDareafor_SA.asp">
          <font color="#FF6600"><strong>Pencarian ID Area : </strong> </font> 
          <input name="ID_Area" type="text" id="ID_Area">
          <input type="submit" name="Submit" value="Cari">
        </form></td>
    </tr>
    <tr bgcolor="#669900"> 
      <td colspan="2"> <div align="center"><span class="style7">Welcome <%= Session("updateuser") %></span></div></td>
    </tr>
    <tr bgcolor="#FF9900"> 
      <td colspan="2"><div align="center"><font color="#0000A0"><strong>..:: Records 
          <%=(rsArea_first)%> to <%=(rsArea_last)%> of <%=(rsArea_total)%> ::.. </strong></font></div></td>
    </tr>
  </table>
  <br>
  <% If Not rsArea.EOF Or Not rsArea.BOF Then %>
  <table width="100%" border="1" align="center" cellspacing="0" bordercolor="#FFFFFF">
    <tr bgcolor="#006600"> 
      <td width="17"> <div align="center" class="style5">NO</div></td>
      <td width="97"> <div align="center" class="style5">ID Area</div></td>
      <td width="121"> <div align="center" class="style5">Area</div></td>
      <td width="116"> <div align="center" class="style5">Area Manager</div></td>
      <td width="113"> <div align="center" class="style5">Ass AM</div></td>
    </tr>
    <% 
While ((Repeat2__numRows <> 0) AND (NOT rsArea.EOF)) 
%>
    <tr bgcolor="#CCCCCC"> 
      <td><div align="center"><font color="#006600"><%=(Repeat1__index + 1)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsArea.Fields.Item("ID_Area").Value)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsArea.Fields.Item("Area").Value)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsArea.Fields.Item("AM").Value)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsArea.Fields.Item("Ass_AM").Value)%></font></div></td>
    </tr>
    <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  rsArea.MoveNext()
Wend
%>
  </table>
  <% End If ' end Not rsArea.EOF Or NOT rsArea.BOF %>
  <p>&nbsp; </p>
  <table width="750" border="1">
    <tr> 
      <td bordercolor="#FFFFCC" bgcolor="#669900">&nbsp;</td>
    </tr>
  </table>
  <br>
  <table width="600" border="0">
    <tr> 
      <td>&nbsp;</td>
    </tr>
  </table>
</div>
</body>
</html>
<%
rsArea.Close()
Set rsArea = Nothing
%>
