<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%> 
<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "index.asp"
  ' redirect with URL parameters (remove the "MM_Logoutnow" query param).
  if (MM_logoutRedirectPage = "") Then MM_logoutRedirectPage = CStr(Request.ServerVariables("URL"))
  If (InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
    MM_newQS = "?"
    For Each Item In Request.QueryString
      If (Item <> "MM_Logoutnow") Then
        If (Len(MM_newQS) > 1) Then MM_newQS = MM_newQS & "&"
        MM_newQS = MM_newQS & Item & "=" & Server.URLencode(Request.QueryString(Item))
      End If
    Next
    if (Len(MM_newQS) > 1) Then MM_logoutRedirectPage = MM_logoutRedirectPage & MM_newQS
  End If
  Response.Redirect(MM_logoutRedirectPage)
End If
%>
<!--#include file="Connections/simConn.asp" -->
<%

Dim getID__NamaPemesan
getID__NamaPemesan = ""
if(Session("UpdateUser") <> "") then getID__NamaPemesan = Session("UpdateUser")

%>
<%
Dim rsProdukPesan__MMColParam
rsProdukPesan__MMColParam = "1"
If (Request.Form("UserID") <> "") Then 
  rsProdukPesan__MMColParam = Request.Form("UserID")
End If
%>
<%
Dim rsProdukPesan
Dim rsProdukPesan_numRows

Set rsProdukPesan = Server.CreateObject("ADODB.Recordset")
rsProdukPesan.ActiveConnection = MM_simConn_STRING
rsProdukPesan.Source = "SELECT *  FROM dbo.tb_ProdukPesan  WHERE UserID = '" + Replace(rsProdukPesan__MMColParam, "'", "''") + "' and status <> 'DOC Telah Diterima'  ORDER BY UpdateDate ASC"
rsProdukPesan.CursorType = 0
rsProdukPesan.CursorLocation = 2
rsProdukPesan.LockType = 1
rsProdukPesan.Open()

rsProdukPesan_numRows = 0
%>
<%

set getID = Server.CreateObject("ADODB.Command")
getID.ActiveConnection = MM_simConn_STRING
getID.CommandText = "dbo.P_AmbilID"
getID.CommandType = 4
getID.CommandTimeout = 0
getID.Prepared = true
getID.Parameters.Append getID.CreateParameter("@RETURN_VALUE", 3, 4)
getID.Parameters.Append getID.CreateParameter("@NamaPemesan", 200, 1,50,getID__NamaPemesan)
getID.Parameters.Append getID.CreateParameter("@UserID", 200, 2,100)
getID.Execute()

%>
<%
Dim rsProdPesan__MMColParam
rsProdPesan__MMColParam = "1"
If (Request.Form("UserID") <> "") Then 
  rsProdPesan__MMColParam = Request.Form("UserID")
End If
%>
<%
Dim rsProdPesan
Dim rsProdPesan_numRows

Set rsProdPesan = Server.CreateObject("ADODB.Recordset")
rsProdPesan.ActiveConnection = MM_simConn_STRING
rsProdPesan.Source = "SELECT *  FROM dbo.tb_ProdukPesan  WHERE UserID = '" + Replace(rsProdPesan__MMColParam, "'", "''") + "' and status ='DOC Telah Diterima' "
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
rsProdukPesan_numRows = rsProdukPesan_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
rsProdPesan_numRows = rsProdPesan_numRows + Repeat2__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsProdukPesan_total
Dim rsProdukPesan_first
Dim rsProdukPesan_last

' set the record count
rsProdukPesan_total = rsProdukPesan.RecordCount

' set the number of rows displayed on this page
If (rsProdukPesan_numRows < 0) Then
  rsProdukPesan_numRows = rsProdukPesan_total
Elseif (rsProdukPesan_numRows = 0) Then
  rsProdukPesan_numRows = 1
End If

' set the first and last displayed record
rsProdukPesan_first = 1
rsProdukPesan_last  = rsProdukPesan_first + rsProdukPesan_numRows - 1

' if we have the correct record count, check the other stats
If (rsProdukPesan_total <> -1) Then
  If (rsProdukPesan_first > rsProdukPesan_total) Then
    rsProdukPesan_first = rsProdukPesan_total
  End If
  If (rsProdukPesan_last > rsProdukPesan_total) Then
    rsProdukPesan_last = rsProdukPesan_total
  End If
  If (rsProdukPesan_numRows > rsProdukPesan_total) Then
    rsProdukPesan_numRows = rsProdukPesan_total
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

Set MM_rs    = rsProdukPesan
MM_rsCount   = rsProdukPesan_total
MM_size      = rsProdukPesan_numRows
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
rsProdukPesan_first = MM_offset + 1
rsProdukPesan_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsProdukPesan_first > MM_rsCount) Then
    rsProdukPesan_first = MM_rsCount
  End If
  If (rsProdukPesan_last > MM_rsCount) Then
    rsProdukPesan_last = MM_rsCount
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
<html>
<head>
<title>Activities :: Sierad </title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="Capex/css/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function MM_displayStatusMsg(msgStr) { //v1.0
  status=msgStr;
  document.MM_returnValue = true;
}
//-->
</script>
</head>

<body topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('img/depan_on.gif','img/ID_on.gif','img/laporan_on.gif')">
<table width="800" border="1" align="center" bordercolor="#006600">
  <tr bordercolor="#006600"> 
    <td colspan="2"> <div align="left"><img src="Capex/Image/sieradonline.gif" width="222" height="85"> 
      </div>
      <div align="right"><font color="#006600">Date : 
        <script name="current" src="GeneratedItems/current.js" language="JavaScript1.2"></script>
        </font></div></td>
  </tr>
  <tr bordercolor="#006600" bgcolor="#CCCCCC"> 
    <td> <div align="left"><font color="#006600">Selamat Datang </font><font color="#006600"><%= Session("UpdateUser") %></font></div></td>
    <td width="300"> <div align="center"><font color="#009900"><a href="contact.asp"><font color="#006600">Hubungi 
        Kami</font></a></font><font color="#FF0000">&nbsp; </font>| <a href="karir.asp"><font color="#006600">Karir 
        </font></a>| <a href="link.asp"><font color="#006600">Links </font></a>| 
        <font color="#006600"><a href="<%= MM_Logout %>">Log Out</a></font></div></td>
  </tr>
</table>
<table width="800" border="0" align="center" bordercolor="#FF6600" bgcolor="#006600">
  <tr>
    <td><div align="center"><img src="BREEDING/img/spacer.gif" width="795" height="10"></div></td>
  </tr>
</table>
<table width="800" border="2" align="center" bordercolor="#CCCCCC" bgcolor="#006600">
  <tr> 
    <td width="150" height="23"><div align="center"><a href="main_page.asp" onMouseOver="MM_swapImage('Image1','','img/depan_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/depan.gif" name="Image1" width="150" height="20" border="0" id="Image1"></a></div></td>
    <td rowspan="5" bgcolor="#006600"> <div align="left"> 
        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="323" height="122" align="middle">
          <param name="movie" value="Capex/Animasi/anakayam.swf">
          <param name="quality" value="high">
          <param name="SCALE" value="exactfit">
          <embed src="Capex/Animasi/anakayam.swf" width="323" height="122" align="middle" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" scale="exactfit"></embed></object>
        <font color="#FFFFFF"></font><font color="#FFFFFF"></font></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="id_anda.asp" onMouseOver="MM_swapImage('Image2','','img/ID_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/ID.gif" name="Image2" width="150" height="20" border="0" id="Image2"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="komentar.asp" onMouseOver="MM_swapImage('Image3','','img/laporan_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/laporan.gif" name="Image3" width="150" height="20" border="0" id="Image3"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"></div></td>
  </tr>
  <tr> 
    <td height="24"> <div align="center"></div></td>
  </tr>
</table>
<table width="800" border="0" align="center" bordercolor="#FF6600" bgcolor="#006600">
  <tr>
    <td><div align="center"><img src="BREEDING/img/spacer.gif" width="795" height="10"></div></td>
  </tr>
</table>
<table width="800" border="1" align="center" background="business/img/bg.gif">
  <tr>
    <td height="200"> 
      <div align="center"> 
        <table width="100%" border="2" bordercolor="#009966">
          
          
          
          
          <tr bordercolor="1"> 
            <td width="54%" height="16"><p><font color="#336666">ID Anda telah 
                terdaftar dalam database kami, Anda bisa melihat status pemesanan 
                anda dengan meng-klik tombol dibawah ini.</font></p>
              <p><font color="#336666">ID Anda :</font></p>
              <form name="form1" method="post" action="">
                <p>
                  <INPUT                        name=UserID
                        style="BORDER-RIGHT: #999999 1px solid; BORDER-TOP: #333333 1px solid; FONT-SIZE: 11px; BORDER-LEFT: #333333 1px solid; WIDTH: 150px; COLOR: #000000; BORDER-BOTTOM: #999999 1px solid; HEIGHT: 18px; BACKGROUND-COLOR: #ffecec" value="<%= getID.Parameters.Item("@UserID").Value %>" readonly ="text">
                  <input type="submit" name="Submit" value="Lihat Status">
                </p>
                <p>&nbsp;</p>
              </form></td>
          </tr>
        </table>
        
        
        <% If Not rsProdukPesan.EOF Or Not rsProdukPesan.BOF Then %>
        <table width="740" height="197" border="1">
          <tr bgcolor="#FF9900"> 
            <td bgcolor="#006600"> <div align="center"><font color="#FFFFFF"> 
                Status Pencarian : Sukses</font></div>
              </td>
          </tr>
          <tr bgcolor="#663366"> 
            <td height="173"> <p><font color="#FFFFFF">DOC yang telah diterima 
                :</font></p>
              <table width="762" border="1">
                <tr bgcolor="#6666FF"> 
                  <td> <div align="center"><font color="#FFFFFF">No</font></div></td>
                  <td> <div align="center"><font color="#FFFFFF">Produk ID</font></div></td>
                  <td> <div align="center"><font color="#FFFFFF">Nama Produk </font></div></td>
                  <td> <div align="center"><font color="#FFFFFF">Kuantitas</font></div></td>
                  <td> <div align="center"><font color="#FFFFFF">Satuan</font></div></td>
                  <td> <div align="center"><font color="#FFFFFF">Status</font></div></td>
                </tr>
                <% 
While ((Repeat2__numRows <> 0) AND (NOT rsProdPesan.EOF)) 
%>
                <tr> 
                  <td> 
                    <div align="center"><font color="#FFFFFF"><%=(Repeat2__index + 1)%></font></div></td>
                  <td> 
                    <div align="center"><font color="#FFFFFF"><%=(rsProdPesan.Fields.Item("ProdukID").Value)%></font></div></td>
                  <td> 
                    <div align="center"><font color="#FFFFFF"><%=(rsProdPesan.Fields.Item("Nama_Produk").Value)%></font></div></td>
                  <td> 
                    <div align="center"><font color="#FFFFFF"><%=(rsProdPesan.Fields.Item("Kuantitas").Value)%></font></div></td>
                  <td> 
                    <div align="center"><font color="#FFFFFF"><%=(rsProdPesan.Fields.Item("SatuanBerat").Value)%></font></div></td>
                  <td> 
                    <div align="center"><font color="#FFFFFF"><%=(rsProdPesan.Fields.Item("Status").Value)%></font></div></td>
                </tr>
                <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  rsProdPesan.MoveNext()
Wend
%>
              </table>
              <p>&nbsp;</p>
              <p><font color="#FFFFFF">Menunggu konfirmasi :</font></p>
              <table width="780" height="44" border="1" align="center" bordercolor="#00FFFF">
                <tr bgcolor="#6666FF"> 
                  <td width="18"><div align="center"><font color="#FFFFFF">No 
                      </font></div></td>
                  <td width="131"> <div align="center"><font color="#FFFFFF">Produk 
                      ID </font></div></td>
                  <td width="123"> <div align="center"><font color="#FFFFFF">Nama 
                      Produk </font></div></td>
                  <td width="134"> <div align="center"><font color="#FFFFFF">Kuantitas</font></div></td>
                  <td width="164"><div align="center"><font color="#FFFFFF">Status</font></div></td>
                  <td width="164"><div align="center"><font color="#FFFFFF">Harga</font></div></td>
                </tr>
                <% While ((Repeat1__numRows <> 0) AND (NOT rsProdukPesan.EOF)) %>
                <tr bgcolor="#CCCCCC"> 
                  <td height="22"> 
                    <div align="center"><%=(Repeat1__index + 1)%></div></td>
                  <td><div align="left"><A HREF="ubah_status.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ProdukID=" & rsProdukPesan.Fields.Item("ProdukID").Value %>"><%=(rsProdukPesan.Fields.Item("ProdukID").Value)%></A> </div></td>
                  <td><div align="center"><%=(rsProdukPesan.Fields.Item("Nama_Produk").Value)%></div></td>
                  <td><div align="center"><%=(rsProdukPesan.Fields.Item("Kuantitas").Value)%></div></td>
                  <td><div align="center"><font color="#FF0000"><strong><%=(rsProdukPesan.Fields.Item("Status").Value)%></strong> </font> </div></td>
                  <td><div align="center"><%=(rsProdukPesan.Fields.Item("QtyPrice").Value)%></div></td>
                </tr>
                <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsProdukPesan.MoveNext()
Wend
%>
              </table>
              <table border="0" width="600" align="center">
                <tr> 
                  <td width="42%" height="45" align="center"> <% If MM_offset <> 0 Then %> <a href="<%=MM_moveFirst%>"><img src="img/first.gif" name="first" width="75" height="25" border=0 align="right" onMouseOver="MM_displayStatusMsg('rz : list company -&gt; first page');return document.MM_returnValue" onload=""></a> 
                    <% End If ' end MM_offset <> 0 %> </td>
                  <td width="8%" align="center"> 
                    <% If MM_offset <> 0 Then %>
                    <a href="<%=MM_movePrev%>"><img src="img/previous.gif" name="previous" width="75" height="25" border=0 align="right" onMouseOver="MM_displayStatusMsg('rz : list company -&gt; previous page');return document.MM_returnValue" onload=""></a> 
                    <% End If ' end MM_offset <> 0 %>
                  </td>
                  <td width="8%" align="center"> 
                    <% If Not MM_atTotal Then %>
                    <a href="<%=MM_moveNext%>"><img src="img/next.gif" name="next" width="75" height="25" border=0 align="left" onMouseOver="MM_displayStatusMsg('rz : list company -&gt; next page');return document.MM_returnValue" onload=""></a> 
                    <% End If ' end Not MM_atTotal %>
                  </td>
                  <td width="42%" align="center"> 
                    <% If Not MM_atTotal Then %>
                    <a href="<%=MM_moveLast%>"><img src="img/last.gif" width="75" height="25" border=0 align="left" onMouseOver="MM_displayStatusMsg('rz : list company -&gt; last page');return document.MM_returnValue"></a> 
                    <% End If ' end Not MM_atTotal %>
                  </td>
                </tr>
              </table>
              <p><font color="#FFFFFF">Silahkan klik Produk ID untuk mengetahui 
                status pemesanan Anda !</font> </p>
              
              
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              </td>
          </tr>
        </table>
        <% End If ' end Not rsProdukPesan.EOF Or NOT rsProdukPesan.BOF %>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
      </div></td>
  </tr>
</table>
<table width="800" border="1" align="center" background="business/img/bg.gif">
  <tr>
    <td><p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p><img src="img/bannerrg.jpg" width="790" height="20"></p></td>
  </tr>
</table>
</body>
</html>
<%
rsProdukPesan.Close()
Set rsProdukPesan = Nothing
%>
<%
rsProdPesan.Close()
Set rsProdPesan = Nothing
%>
