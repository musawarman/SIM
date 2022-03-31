<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%> 
<!--#include file="../../Connections/simConn.asp" -->
<%
Dim rsSupplier__MMColParam
rsSupplier__MMColParam = "1"
If (Request.Form("Company_Name") <> "") Then 
  rsSupplier__MMColParam = Request.Form("Company_Name")
End If
%>
<%
Dim rsSupplier
Dim rsSupplier_numRows

Set rsSupplier = Server.CreateObject("ADODB.Recordset")
rsSupplier.ActiveConnection = MM_simConn_STRING
rsSupplier.Source = "SELECT * FROM dbo.Supplier WHERE Company_Name = '" + Replace(rsSupplier__MMColParam, "'", "''") + "'"
rsSupplier.CursorType = 0
rsSupplier.CursorLocation = 2
rsSupplier.LockType = 1
rsSupplier.Open()

rsSupplier_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsSupplier_numRows = rsSupplier_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsSupplier_total
Dim rsSupplier_first
Dim rsSupplier_last

' set the record count
rsSupplier_total = rsSupplier.RecordCount

' set the number of rows displayed on this page
If (rsSupplier_numRows < 0) Then
  rsSupplier_numRows = rsSupplier_total
Elseif (rsSupplier_numRows = 0) Then
  rsSupplier_numRows = 1
End If

' set the first and last displayed record
rsSupplier_first = 1
rsSupplier_last  = rsSupplier_first + rsSupplier_numRows - 1

' if we have the correct record count, check the other stats
If (rsSupplier_total <> -1) Then
  If (rsSupplier_first > rsSupplier_total) Then
    rsSupplier_first = rsSupplier_total
  End If
  If (rsSupplier_last > rsSupplier_total) Then
    rsSupplier_last = rsSupplier_total
  End If
  If (rsSupplier_numRows > rsSupplier_total) Then
    rsSupplier_numRows = rsSupplier_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsSupplier_total = -1) Then

  ' count the total records by iterating through the recordset
  rsSupplier_total=0
  While (Not rsSupplier.EOF)
    rsSupplier_total = rsSupplier_total + 1
    rsSupplier.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsSupplier.CursorType > 0) Then
    rsSupplier.MoveFirst
  Else
    rsSupplier.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsSupplier_numRows < 0 Or rsSupplier_numRows > rsSupplier_total) Then
    rsSupplier_numRows = rsSupplier_total
  End If

  ' set the first and last displayed record
  rsSupplier_first = 1
  rsSupplier_last = rsSupplier_first + rsSupplier_numRows - 1
  
  If (rsSupplier_first > rsSupplier_total) Then
    rsSupplier_first = rsSupplier_total
  End If
  If (rsSupplier_last > rsSupplier_total) Then
    rsSupplier_last = rsSupplier_total
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
<form name="form1" method="post" action="SearchcompanyName.asp">
          <font color="#FF6600"><strong>Pencarian Company Name: </strong> </font> 
          <input name="Company_Name" type="text" id="Company_Name">
          <input type="submit" name="Submit" value="Cari">
        </form></td>
    </tr>
    <tr bgcolor="#669900"> 
      <td colspan="2"> <div align="center"><span class="style7">Welcome <%= Session("updateuser") %></span></div></td>
    </tr>
    <tr bgcolor="#FF9900"> 
      <td colspan="2"><div align="center"><font color="#0000A0"><strong>..:: Records 
          <%=(rsSupplier_first)%> to <%=(rsSupplier_last)%> of <%=(rsSupplier_total)%> ::.. </strong></font></div></td>
    </tr>
  </table>
  <br>
  <% If Not rsSupplier.EOF Or Not rsSupplier.BOF Then %>
  <table width="100%" border="1" align="center" cellspacing="0" bordercolor="#FFFFFF">
    <tr bgcolor="#006600"> 
      <td width="17"> <div align="center" class="style5">NO</div></td>
      <td width="97"> <div align="center" class="style5">SupplierID</div></td>
      <td width="121"> <div align="center" class="style5">Supplier Name</div></td>
      <td width="116"> <div align="center" class="style5">Company Name</div></td>
      <td width="113"> <div align="center" class="style5">Contact Name</div></td>
      <td width="121"> <div align="center" class="style5">Address</div></td>
      <td width="113"> <div align="center" class="style5">City</div></td>
      <td width="121"> <div align="center" class="style5">Region</div></td>
      <td> <div align="center" class="style5">Phone</div></td>
      <td> <div align="center" class="style5">Home Page</div></td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rsSupplier.EOF)) %>
    <tr bgcolor="#CCCCCC"> 
      <td><div align="center"><font color="#006600"><%=(Repeat1__index + 1)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsSupplier.Fields.Item("SupplierID").Value)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsSupplier.Fields.Item("Supplier_Name").Value)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsSupplier.Fields.Item("Company_Name").Value)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsSupplier.Fields.Item("Contact_Name").Value)%></font></div></td>
      <td><div align="left"><font color="#006600"><%=(rsSupplier.Fields.Item("Address").Value)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsSupplier.Fields.Item("City").Value)%></font></div></td>
      <td><div align="center"><font color="#006600"><%=(rsSupplier.Fields.Item("Region").Value)%></font></div></td>
      <td><div align="left"><%=(rsSupplier.Fields.Item("Phone").Value)%></div>
        <div align="center"><font color="#006600"></font></div></td>
      <td><div align="left"></div>
        <div align="center"><font color="#006600"><%=(rsSupplier.Fields.Item("Home_Page").Value)%></font></div></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsSupplier.MoveNext()
Wend
%>
  </table>
  <% End If ' end Not rsSupplier.EOF Or NOT rsSupplier.BOF %>
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
rsSupplier.Close()
Set rsSupplier = Nothing
%>
