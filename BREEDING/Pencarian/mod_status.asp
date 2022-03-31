<%@LANGUAGE="VBSCRIPT"%>

 
<!--#include file="../../Connections/simConn.asp" -->

<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_simConn_STRING
  MM_editTable = "dbo.tb_ProdukPesan"
  MM_editColumn = "ProdukID"
  MM_recordId = "'" + Request.Form("MM_recordId") + "'"
  MM_editRedirectUrl = "confirm_status.asp"
  MM_fieldsStr  = "Produk ID|value|Nama Produk|value|Quantity Per Unit|value|UOM|value|HargaPerDOC|value|TotalHarga|value|tanggal|value|select|value|textarea|value|hiddenField|value"
  MM_columnsStr = "ProdukID|',none,''|Nama_Produk|',none,''|Kuantitas|none,none,NULL|SatuanBerat|',none,''|Price|none,none,NULL|QtyPrice|none,none,NULL|Tanggal_Pesan|',none,NULL|Status|',none,''|Status_Desc|',none,''|UpdateUser|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim rsProdPesan__MMColParam
rsProdPesan__MMColParam = "1"
If (Request.QueryString("ProdukID") <> "") Then 
  rsProdPesan__MMColParam = Request.QueryString("ProdukID")
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
      <td colspan="2"><h3 align="center"><font color="#6699FF">.:: Ubah Status 
          ::. </font></h3></td>
    </tr>
    <tr>
      <td colspan="2">&nbsp; </td>
    </tr>
    <tr bgcolor="#669900"> 
      <td colspan="2"> <div align="center"><span class="style7">Welcome <%= Session("updateuser") %></span></div></td>
    </tr>
    <tr bgcolor="#FF9900"> 
      <td colspan="2"><div align="center"><font color="#0000A0"></font></div></td>
    </tr>
  </table>
  <form name="form1" method="POST" action="<%=MM_editAction%>">
    
    <table width="53%" border="0" align="center" bordercolor="#003399">
      
      <tr> 
        <td colspan="3"> <div align="center"><img src="../../img/garis1.gif" width="500" height="2"></div></td>
      </tr>
      <tr bordercolor="1"> 
        <td> <div align="right">Produk ID </div></td>
        <td width="1%"> <div align="center">:</div></td>
        <td><div align="left"> 
            <input name="Produk ID" id="Produk ID" value="<%=(rsProdPesan.Fields.Item("ProdukID").Value)%>" size="40" readonly="text">
          </div></td>
      </tr>
      <tr bordercolor="1"> 
        <td><div align="right">Nama Produk </div></td>
        <td> <div align="center">:</div></td>
        <td><div align="left"> 
            <input name="Nama Produk" id="Nama Produk" value="<%=(rsProdPesan.Fields.Item("Nama_Produk").Value)%>" size="40" readonly="text">
          </div></td>
      </tr>
      <tr bordercolor="1"> 
        <td><div align="right">Kuantitas </div></td>
        <td> <div align="center">:</div></td>
        <td><div align="left"> 
            <input name="Quantity Per Unit" id="Quantity Per Unit" value="<%=(rsProdPesan.Fields.Item("Kuantitas").Value)%>" readonly="text">
          </div></td>
      </tr>
      <tr bordercolor="1"> 
        <td><div align="right">Satuan Berat </div></td>
        <td><div align="center">:</div></td>
        <td><div align="left"> 
            <input name="UOM" id="UOM2" value="<%=(rsProdPesan.Fields.Item("SatuanBerat").Value)%>" readonly="text">
          </div></td>
      </tr>
      <tr bordercolor="1"> 
        <td><div align="right">Harga Per DOC </div></td>
        <td> <div align="center">:</div></td>
        <td> <div align="left"> 
            <input name="HargaPerDOC" id="HargaPerDOC2" value="<%=(rsProdPesan.Fields.Item("Price").Value)%>" size="40" readonly="text">
          </div></td>
      </tr>
      <tr bordercolor="1"> 
        <td><div align="right">Total Harga </div></td>
        <td> <div align="center">:</div></td>
        <td> <div align="left"> 
            <input name="TotalHarga" id="TotalHarga2" value="<%=(rsProdPesan.Fields.Item("QtyPrice").Value)%>" size="40" readonly="text">
          </div></td>
      </tr>
      <tr bordercolor="1"> 
        <td><div align="right">Tanggal Pesan </div></td>
        <td> <div align="center">:</div></td>
        <td> <div align="left"> 
            <input name="tanggal" id="tanggal" value="<%=(rsProdPesan.Fields.Item("Tanggal_Pesan").Value)%>" readonly="text">
          </div></td>
      </tr>
      <tr bordercolor="1"> 
        <td height="20"> <div align="right">Pilih Status *</div></td>
        <td> <div align="center">:</div></td>
        <td> <div align="left"> 
            <select name="select">
              <option>DOC Dibatalkan</option>
              <option>DOC Ditunda</option>
              <option>Sedang Dalam Proses</option>
              <option>DOC Telah Diterima</option>
            </select>
          </div></td>
      </tr>
      <tr bordercolor="1"> 
        <td height="20"> <div align="right">Keterangan *</div></td>
        <td> <div align="center">:</div></td>
        <td> <div align="left"> 
            <textarea name="textarea" cols="40" rows="10"></textarea>
          </div></td>
      </tr>
      <tr bordercolor="1"> 
        <td width="28%"> <div align="right"> 
            <input type="reset" name="Reset" value="Reset">
            <input type="submit" name="Submit2" value="Ubah !">
            <input name="hiddenField" type="hidden" value="<%= Session("updateuser") %>">
          </div></td>
        <td> <div align="center"></div></td>
        <td width="71%"> <div align="left"> </div></td>
      </tr>
      <tr> 
        <td colspan="3"> <div align="center"><img src="../../img/garis1.gif" width="500" height="2"></div></td>
      </tr>
    </table>
  
<input type="hidden" name="MM_update" value="form1">
    <input type="hidden" name="MM_recordId" value="<%= rsProdPesan.Fields.Item("ProdukID").Value %>">
  </form>
  <br>
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
