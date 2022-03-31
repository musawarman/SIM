<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%> 
<!--#include file="../../Connections/simConn.asp" -->
<%
Dim rsProdPesan__MMColParam
rsProdPesan__MMColParam = "1"
If (Request.Form("Nama_Produk") <> "") Then 
  rsProdPesan__MMColParam = Request.Form("Nama_Produk")
End If
%>
<%
Dim rsProdPesan
Dim rsProdPesan_numRows

Set rsProdPesan = Server.CreateObject("ADODB.Recordset")
rsProdPesan.ActiveConnection = MM_simConn_STRING
rsProdPesan.Source = "SELECT * FROM dbo.tb_ProdukPesan WHERE Nama_Produk = '" + Replace(rsProdPesan__MMColParam, "'", "''") + "'"
rsProdPesan.CursorType = 0
rsProdPesan.CursorLocation = 2
rsProdPesan.LockType = 1
rsProdPesan.Open()

rsProdPesan_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsProdPesan_numRows = rsProdPesan_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsProdPesan_total
Dim rsProdPesan_first
Dim rsProdPesan_last

' set the record count
rsProdPesan_total = rsProdPesan.RecordCount

' set the number of rows displayed on this page
If (rsProdPesan_numRows < 0) Then
  rsProdPesan_numRows = rsProdPesan_total
Elseif (rsProdPesan_numRows = 0) Then
  rsProdPesan_numRows = 1
End If

' set the first and last displayed record
rsProdPesan_first = 1
rsProdPesan_last  = rsProdPesan_first + rsProdPesan_numRows - 1

' if we have the correct record count, check the other stats
If (rsProdPesan_total <> -1) Then
  If (rsProdPesan_first > rsProdPesan_total) Then
    rsProdPesan_first = rsProdPesan_total
  End If
  If (rsProdPesan_last > rsProdPesan_total) Then
    rsProdPesan_last = rsProdPesan_total
  End If
  If (rsProdPesan_numRows > rsProdPesan_total) Then
    rsProdPesan_numRows = rsProdPesan_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsProdPesan_total = -1) Then

  ' count the total records by iterating through the recordset
  rsProdPesan_total=0
  While (Not rsProdPesan.EOF)
    rsProdPesan_total = rsProdPesan_total + 1
    rsProdPesan.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsProdPesan.CursorType > 0) Then
    rsProdPesan.MoveFirst
  Else
    rsProdPesan.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsProdPesan_numRows < 0 Or rsProdPesan_numRows > rsProdPesan_total) Then
    rsProdPesan_numRows = rsProdPesan_total
  End If

  ' set the first and last displayed record
  rsProdPesan_first = 1
  rsProdPesan_last = rsProdPesan_first + rsProdPesan_numRows - 1
  
  If (rsProdPesan_first > rsProdPesan_total) Then
    rsProdPesan_first = rsProdPesan_total
  End If
  If (rsProdPesan_last > rsProdPesan_total) Then
    rsProdPesan_last = rsProdPesan_total
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
<form name="form1" method="post" action="SearchNamaProduk.asp">
          <font color="#FF6600"><strong>Pencarian Nama Produk : </strong> </font> 
          <input name="Nama_Produk" type="text" id="Nama_Produk">
          <input type="submit" name="Submit" value="Cari">
        </form></td>
    </tr>
    <tr bgcolor="#669900"> 
      <td colspan="2"> <div align="center"><span class="style7">Welcome <%= Session("updateuser") %></span></div></td>
    </tr>
    <tr bgcolor="#FF9900"> 
      <td colspan="2"><div align="center"><font color="#0000A0"><strong>..:: Records 
          <%=(rsProdPesan_first)%> to <%=(rsProdPesan_last)%> of <%=(rsProdPesan_total)%> ::.. </strong></font></div></td>
    </tr>
  </table>
  <br>
  <% If Not rsProdPesan.EOF Or Not rsProdPesan.BOF Then %>
  <table border="1" align="center" cellspacing="0" bordercolor="#FFFFFF">
    <tr bgcolor="#006600"> 
      <td width="17"> <div align="center" class="style5">NO</div></td>
      <td width="97"> <div align="center" class="style5">ProdukID</div></td>
      <td width="121"> <div align="center" class="style5">UserID</div></td>
      <td width="116"> <div align="center" class="style5">Category</div></td>
      <td width="113"> <div align="center" class="style5">Nama Produk</div></td>
      <td width="121"> <div align="center" class="style5">Kuantitas</div></td>
      <td width="116"> <div align="center" class="style5">Tanggal Pesan</div></td>
      <td width="113"> <div align="center" class="style5">Status</div></td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rsProdPesan.EOF)) %>
    <tr bgcolor="#CCCCCC"> 
      <td><div align="center"><font color="#006600"><%=(Repeat1__index + 1)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsProdPesan.Fields.Item("ProdukID").Value)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsProdPesan.Fields.Item("UserID").Value)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsProdPesan.Fields.Item("Category").Value)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsProdPesan.Fields.Item("Nama_Produk").Value)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsProdPesan.Fields.Item("Kuantitas").Value)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsProdPesan.Fields.Item("Tanggal_Pesan").Value)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsProdPesan.Fields.Item("Status").Value)%></font></div></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsProdPesan.MoveNext()
Wend
%>
  </table>
  <% End If ' end Not rsProdPesan.EOF Or NOT rsProdPesan.BOF %>
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
rsProdPesan.Close()
Set rsProdPesan = Nothing
%>
