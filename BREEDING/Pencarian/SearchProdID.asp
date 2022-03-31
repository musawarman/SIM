<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%> 
<!--#include file="../../Connections/simConn.asp" -->
<%
Dim rsProdPesan__MMColParam
rsProdPesan__MMColParam = "1"
If (Request.Form("ProdukID") <> "") Then 
  rsProdPesan__MMColParam = Request.Form("ProdukID")
End If
%>
<%
Dim rsProdPesan
Dim rsProdPesan_numRows

Set rsProdPesan = Server.CreateObject("ADODB.Recordset")
rsProdPesan.ActiveConnection = MM_simConn_STRING
rsProdPesan.Source = "SELECT * FROM dbo.tb_ProdukPesan WHERE ProdukID = '" + Replace(rsProdPesan__MMColParam, "'", "''") + "'"
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
<%
Dim MM_paramName 
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
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
<form name="form1" method="post" action="SearchProdID.asp">
          <font color="#FF6600"><strong>Pencarian Produk ID: </strong> </font> 
          <input name="ProdukID" type="text" id="ProdukID">
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
      <td width="121"> <div align="center" class="style5">Kuantitas</div></td>
      <td width="116"> <div align="center" class="style5">Tanggal Pesan</div></td>
      <td width="113"> <div align="center" class="style5">Status</div></td>
      <td width="113"> <div align="center" class="style5">Aksi</div></td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rsProdPesan.EOF)) %>
    <tr bgcolor="#CCCCCC"> 
      <td><div align="center"><font color="#006600"><%=(Repeat1__index + 1)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsProdPesan.Fields.Item("ProdukID").Value)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsProdPesan.Fields.Item("UserID").Value)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsProdPesan.Fields.Item("Kuantitas").Value)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsProdPesan.Fields.Item("Tanggal_Pesan").Value)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsProdPesan.Fields.Item("Status").Value)%></font></div></td>
      <td><div align="center"><font color="#006600"><A HREF="mod_status.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProdukID=" & rsProdPesan.Fields.Item("ProdukID").Value %>">UbahStatus</A></font></div></td>
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
