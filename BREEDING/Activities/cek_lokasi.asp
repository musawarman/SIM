<%@LANGUAGE="VBSCRIPT"%> 
<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "../../index.asp"
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
<!--#include file="../../Connections/simConn.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="5,4,3"
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
Dim rsPemesan__MMColParam
rsPemesan__MMColParam = "1"
If (Request.Form("UserID") <> "") Then 
  rsPemesan__MMColParam = Request.Form("UserID")
End If
%>
<%
Dim rsPemesan
Dim rsPemesan_numRows

Set rsPemesan = Server.CreateObject("ADODB.Recordset")
rsPemesan.ActiveConnection = MM_simConn_STRING
rsPemesan.Source = "SELECT * FROM dbo.tb_Pemesan WHERE UserID = '" + Replace(rsPemesan__MMColParam, "'", "''") + "'"
rsPemesan.CursorType = 0
rsPemesan.CursorLocation = 2
rsPemesan.LockType = 1
rsPemesan.Open()

rsPemesan_numRows = 0
%>
<%
Dim Repeat3__numRows
Dim Repeat3__index

Repeat3__numRows = 10
Repeat3__index = 0
rsPemesanan_numRows = rsPemesanan_numRows + Repeat3__numRows
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
rsAcceptProd_first = MM_offset + 1
rsAcceptProd_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsAcceptProd_first > MM_rsCount) Then
    rsAcceptProd_first = MM_rsCount
  End If
  If (rsAcceptProd_last > MM_rsCount) Then
    rsAcceptProd_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsProdukDT_first = MM_offset + 1
rsProdukDT_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsProdukDT_first > MM_rsCount) Then
    rsProdukDT_first = MM_rsCount
  End If
  If (rsProdukDT_last > MM_rsCount) Then
    rsProdukDT_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsProdukHD_first = MM_offset + 1
rsProdukHD_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsProdukHD_first > MM_rsCount) Then
    rsProdukHD_first = MM_rsCount
  End If
  If (rsProdukHD_last > MM_rsCount) Then
    rsProdukHD_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsOrder_first = MM_offset + 1
rsOrder_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsOrder_first > MM_rsCount) Then
    rsOrder_first = MM_rsCount
  End If
  If (rsOrder_last > MM_rsCount) Then
    rsOrder_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<html>
<head>
<title>:: Sierad : Activities ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

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

<body background="../../business/img/bg.gif" leftmargin="0" topmargin="0" onLoad="MM_preloadImages('../img/home_on.gif','../img/sysman_on.gif','../img/activities_on.gif','../img/report_on.gif','../img/faq_on.gif')">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
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
</table>
<table width="100%" border="0">
  <tr>
    <td><img src="../img/spacer.gif" width="1117" height="10"></td>
  </tr>
</table>
<table width="100%" border="1" bordercolor="#009900" background="../../img/bg.gif">
  <tr> 
    <td width="14%"><div align="left"><a href="../index.asp" onMouseOver="MM_swapImage('Image2','','../img/home_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../img/home.gif" name="Image2" width="150" height="20" border="0" id="Image2"></a></div></td>
    <td rowspan="5"> <div align="center"><img src="../img/img_sierad_produce.jpg" width="333" height="120" align="left"><img src="../img/cooling.jpg" width="216" height="120"></div></td>
    <td width="0%" rowspan="5"><div align="left"><img src="../../img/garis.gif" width="1" height="120"></div></td>
    <td rowspan="5" background="../../business/img/bg.gif"> <div align="right"> 
        Tanggal : 
        <script name="current" src="../../GeneratedItems/current.js" language="JavaScript1.2"></script>
      </div>
      <p>&nbsp;</p>
      <p align="center"><font color="#009900"><a href="../contact.asp"><font color="#009900">Hubungi 
        Kami</font></a></font><font color="#FF0000"> </font>| <a href="../karir.asp"><font color="#009900">Karir 
        </font></a>| <a href="../link.asp"><font color="#009900">Links </font></a>| 
        <font color="#009900"><a href="<%= MM_Logout %>">Logout</a></font> 
      </p>
      <p align="center"><img src="../../img/garis1.gif" width="388" height="1" align="top"></p></td>
  </tr>
  <tr> 
    <td><div align="left"><a href="../../ADMIN/login.asp" onMouseOver="MM_swapImage('Image3','','../img/sysman_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../img/sysman.gif" name="Image3" width="150" height="20" border="0" id="Image3"></a></div></td>
  </tr>
  <tr> 
    <td><div align="left"><a href="activities.asp" onMouseOver="MM_swapImage('Image4','','../img/activities_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../img/activities.gif" name="Image4" width="150" height="20" border="0" id="Image4"></a></div></td>
  </tr>
  <tr> 
    <td><div align="left"><a href="../Reports/reportListing.asp" onMouseOver="MM_swapImage('Image5','','../img/report_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../img/report.gif" name="Image5" width="150" height="20" border="0" id="Image5"></a></div></td>
  </tr>
  <tr> 
    <td height="24"> <div align="left"><a href="../Pencarian/Search.asp" onMouseOver="MM_swapImage('Image1','','../img/faq_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../img/faq.gif" name="Image1" width="150" height="20" border="0" id="Image1"></a></div></td>
  </tr>
</table>
<table width="100%" border="0">
  <tr>
    <td><img src="../img/spacer.gif" width="1117" height="10"></td>
  </tr>
</table>
<table width="100%" height="580" border="0">
  <tr> 
    <td height="576"> 
      <p><font color="#993300">&gt;&gt; Lokasi Mitra</font></p>
      <% If Not rsPemesan.EOF Or Not rsPemesan.BOF Then %>
      <table width="60%" border="0" align="center">
        <tr> 
          <td width="33%"> <div align="right">ID :</div></td>
          <td width="67%"> <input name="UserID" readonly="text" value="<%=(rsPemesan.Fields.Item("UserID").Value)%>"></td>
        </tr>
        <tr> 
          <td> <div align="right">Nama Pemesan :</div></td>
          <td> <input name="textfield3" readonly="text" value="<%=(rsPemesan.Fields.Item("NamaPemesan").Value)%>"></td>
        </tr>
        <tr> 
          <td> <div align="right">First Name :</div></td>
          <td> <input name="textfield4" readonly="text" value="<%=(rsPemesan.Fields.Item("First_Name").Value)%>"> 
          </td>
        </tr>
        <tr> 
          <td> <div align="right">Last Name :</div></td>
          <td> <input name="textfield5" readonly="text" value="<%=(rsPemesan.Fields.Item("LastName").Value)%>"></td>
        </tr>
        <tr> 
          <td> <div align="right">Alamat Pengiriman :</div></td>
          <td> <input name="textfield6" readonly="text" value="<%=(rsPemesan.Fields.Item("Alamat").Value)%>" size="50"></td>
        </tr>
        <tr> 
          <td> <div align="right">Kota :</div></td>
          <td> <input name="textfield7" readonly="text" value="<%=(rsPemesan.Fields.Item("Kota").Value)%>"></td>
        </tr>
        <tr> 
          <td> <div align="right">Kode Area :</div></td>
          <td> <input name="textfield8" readonly="text" value="<%=(rsPemesan.Fields.Item("Kode_Area").Value)%>"></td>
        </tr>
        <tr> 
          <td> <div align="right">No Telpn :</div></td>
          <td> <input name="textfield9" readonly="text" value="<%=(rsPemesan.Fields.Item("No_Telepon").Value)%>"></td>
        </tr>
        <tr> 
          <td> <div align="right">Fax :</div></td>
          <td> <input name="textfield10" readonly="text" value="<%=(rsPemesan.Fields.Item("Fax").Value)%>"></td>
        </tr>
        <tr> 
          <td> <div align="right">Email :</div></td>
          <td> <input name="textfield11" readonly="text" value="<%=(rsPemesan.Fields.Item("Email").Value)%>"></td>
        </tr>
        <tr> 
          <td> <div align="right">Alamat Kandang :</div></td>
          <td> <input name="address" readonly="text" id="address2" value="<%=(rsPemesan.Fields.Item("Alamat_Kandang").Value)%>" size="50"></td>
        </tr>
        <tr> 
          <td> <div align="right">Nama Kandang :</div></td>
          <td> <input name="name" readonly="text" id="name2" value="<%=(rsPemesan.Fields.Item("Nama_Kandang").Value)%>"></td>
        </tr>
        <tr> 
          <td> <div align="right"></div></td>
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td><input name="button" type=button class=btn onClick=history.back() value=Back></td>
        </tr>
      </table>
      <br> <br> <br> <br> <br> <br> <br> <br> <table width="100%" border="2" bordercolor="#FF9900" background="../../img/bg.gif">
        <tr> 
          <td height="25"> <div align="center"> <font color="#009900"> &gt;&gt; 
              <a href="../../index.asp">DEPAN</a> | <a href="../../ADMIN/login.asp">ADMINISTRATOR</a> 
              | <a href="activities.asp">AKTIVITAS</a> | <a href="../Reports/reportListing.asp">LAPORAN</a> 
              | <a href="../Pencarian/Search.asp">PENCARIAN</a></font> <font color="#009900">&lt;&lt;</font></div></td>
        </tr>
        <tr> 
          <td height="16"> <div align="center">Web Master PT. Sierad Produce Tbk 
            </div></td>
        </tr>
      </table>
      <% End If ' end Not rsPemesan.EOF Or NOT rsPemesan.BOF %>
    </td>
  </tr>
</table>
</body>
</html>
<%
rsPemesan.Close()
Set rsPemesan = Nothing
%>
