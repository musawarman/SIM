<%@LANGUAGE="VBSCRIPT"%>

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
Dim rsProdukPesan__MMColParam
rsProdukPesan__MMColParam = "1"
If (Request.QueryString("ProdukID") <> "") Then 
  rsProdukPesan__MMColParam = Request.QueryString("ProdukID")
End If
%>
<%
Dim rsProdukPesan
Dim rsProdukPesan_numRows

Set rsProdukPesan = Server.CreateObject("ADODB.Recordset")
rsProdukPesan.ActiveConnection = MM_simConn_STRING
rsProdukPesan.Source = "SELECT * FROM dbo.tb_ProdukPesan WHERE ProdukID = '" + Replace(rsProdukPesan__MMColParam, "'", "''") + "'"
rsProdukPesan.CursorType = 0
rsProdukPesan.CursorLocation = 2
rsProdukPesan.LockType = 1
rsProdukPesan.Open()

rsProdukPesan_numRows = 0
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
        <font color="#006600"><a href="<%= MM_Logout %>">Log 
        Out</a></font></div></td>
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
    <td height="439"> 
      <div align="center"> 
        <table width="740" height="242" border="1">
          <tr bgcolor="#FF9900"> 
            <td bgcolor="#006600"> <div align="center"><font color="#FFFFFF"> 
                <img src="Icon/account.gif" width="29" height="25">Status Pemesanan</font></div></td>
          </tr>
          <tr bgcolor="#663366"> 
            <td height="203">&nbsp; <form name="form1">
                <table width="70%" border="0" align="center">
                  <tr> 
                    <td width="67%"> <div align="right"><font color="#FFFFFF">Produk ID :</font></div></td>
                    <td width="33%"> <input name="kategori" id="kategori" value="<%=(rsProdukPesan.Fields.Item("ProdukID").Value)%>" size="40" readonly="text"></td>
                  </tr>
                  <tr> 
                    <td> <div align="right"><font color="#FFFFFF">Kategori :</font></div></td>
                    <td> <input name="kategori" readonly="text" id="kategori" value="<%=(rsProdukPesan.Fields.Item("Category").Value)%>"></td>
                  </tr>
                  <tr> 
                    <td> <div align="right"><font color="#FFFFFF">Nama Produk 
                        :</font></div></td>
                    <td> <input name="NamaProd" readonly="text" id="NamaProd" value="<%=(rsProdukPesan.Fields.Item("Nama_Produk").Value)%>"></td>
                  </tr>
                  <tr> 
                    <td> <div align="right"><font color="#FFFFFF">Kuantitas :</font></div></td>
                    <td> <input name="Kuantitas" readonly="text" id="Kuantitas" value="<%=(rsProdukPesan.Fields.Item("Kuantitas").Value)%>"></td>
                  </tr>
                  <tr> 
                    <td> <div align="right"><font color="#FFFFFF">Satuan :</font></div></td>
                    <td> <input name="Satuan" readonly="text" id="Satuan" value="<%=(rsProdukPesan.Fields.Item("SatuanBerat").Value)%>"></td>
                  </tr>
                  <tr> 
                    <td> <div align="right"><font color="#FFFFFF">Status :</font></div></td>
                    <td> <input name="Status" readonly="text" id="Status" value="<%=(rsProdukPesan.Fields.Item("Status").Value)%>"></td>
                  </tr>
                  <tr> 
                    <td> <div align="right"><font color="#FFFFFF">Keterangan Status 
                        :</font></div></td>
                    <td> <textarea readonly = "textarea" name="textarea" cols="40" rows="10"><%=(rsProdukPesan.Fields.Item("Status_Desc").Value)%></textarea></td>
                  </tr>
                  <tr> 
                    <td> <div align="right">
                        <INPUT name="button" type=button class=btn onclick=history.back() value=Back>
                      </div></td>
                    <td> <input name="hiddenField" type="hidden" value="<%=(rsProdukPesan.Fields.Item("UpdateUser").Value)%>"> </td>
                  </tr>
                </table>
              </form></td>
          </tr>
        </table>
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
