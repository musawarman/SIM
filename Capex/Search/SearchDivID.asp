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
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="8"
MM_authFailedURL="failed.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (false Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
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

Repeat1__numRows = 10
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
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = rsDivisi
MM_rsCount   = rsDivisi_total
MM_size      = rsDivisi_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsDivisi_first = MM_offset + 1
rsDivisi_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsDivisi_first > MM_rsCount) Then
    rsDivisi_first = MM_rsCount
  End If
  If (rsDivisi_last > MM_rsCount) Then
    rsDivisi_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
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
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = MM_keepMove & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>::List Divisi ::</title>
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
<link href="../css/style.css" rel="stylesheet" type="text/css">
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

<body bgcolor="#FFFFFF" background="../Image/bg.gif">
<div align="center"> 
  <table width="756" border="0">
    <tr> 
      <td width="750" colspan="2"><img src="../Image/banner2.gif" width="750" height="100"></td>
    </tr>
    <tr> 
      <td colspan="2"><h3 align="center"><font color="#6699FF">.:: Pencarian ::. 
          </font></h3></td>
    </tr>
    <tr>
      <td colspan="2">
<form name="form1" method="post" action="SearchDivID.asp">
          <font color="#FF6600"><strong>Pencarian Produk ID: </strong> </font> 
          <input name="ProdukID" type="text" id="ProdukID">
          <input type="submit" name="Submit" value="Search">
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
      <td><div align="center"><a href="../MainMenu/MasterBudget.asp" target="_parent">Master Budget</a> | <a href="../MainMenu/MasterCompany.asp" target="_parent">Master Company</a> | <a href="../MainMenu/MasterCurrency.asp" target="_parent">Master Currency </a> | <a href="../MainMenu/MasterDivisi.asp" target="_parent">Master Divisi </a>| <a href="../MainMenu/MasterUser.asp" target="_parent">Master User </a> | <a href="../MainMenu/MasterVendor.asp" target="_parent">Master Vendor </a></div></td>
    </tr>
  </table>
</div>
</body>
</html>
<%
rsProdPesan.Close()
Set rsProdPesan = Nothing
%>