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
//-->
</script>
</head>

<body background="../../business/img/bg.gif" leftmargin="0" topmargin="0" onLoad="MM_preloadImages('../img/home_on.gif','../img/sysman_on.gif','../img/activities_on.gif','../img/report_on.gif','../img/faq_on.gif','../img/send_prod_on.jpg','../img/send_appr_on.jpg','../img/hold_on.jpg','../img/accept_prod_on.jpg','../img/accept_appr_on.jpg')">
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
    <td><img src="../img/spacer.gif" width="1122" height="10"></td>
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
        <a href="<%= MM_Logout %>">Log Out</a></p>
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
    <td><img src="../img/spacer.gif" width="1122" height="10"></td>
  </tr>
</table>
<table width="100%" border="0">
  <tr> 
    <td height="14"><font color="#993300"> <img src="../../Icon/club.gif" width="29" height="25">AKTIVITAS</font></td>
  </tr>
  <tr>
    <td height="14"><img src="../img/LOOKER1.GIF" width="200" height="12"></td>
  </tr>
</table>
<table width="804" height="399" border="0" align="center" cellpadding="0" cellspacing="0" cool usegridx usegridy showgridx showgridy gridx="16" gridy="16">
  <tr height="398"> 
    <td width="48" height="398"></td>
    <td width="755" height="398" valign="top" align="left" xpos="48"> <table width="700" border="1" align="center" cellpadding="4" cellspacing="10">
        <tr bgcolor="#666666"> 
          <td height="28"> <div align="left"><a href="list_pemesanan.asp" onMouseOver="MM_swapImage('Image6','','../img/send_prod_on.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="../img/send_prod.jpg" name="Image6" width="150" height="30" border="0" id="Image6"></a></div></td>
          <td><div align="left"><a href="send_appr.asp" onMouseOver="MM_swapImage('Image7','','../img/send_appr_on.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="../img/send_appr.jpg" name="Image7" width="150" height="30" border="0" id="Image7"></a> 
            </div></td>
          <td><div align="left"><a href="hold.asp" onMouseOver="MM_swapImage('Image8','','../img/hold_on.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="../img/hold.jpg" name="Image8" width="150" height="30" border="0" id="Image8"></a></div></td>
        </tr>
        <tr> 
          <td><p><font color="#006600">Silahkan klik tombol diatas untuk produk 
              yang akan dikirimkan.</font></p>
            <p align="right">&gt;&gt; <a href="list_orderdetail.asp"><strong>detail</strong></a></p></td>
          <td><p><font color="#006600">Silahkan klik tombol diatas untuk <em>Approval 
              </em>produk yang dikirim.(<a href="detail_approval.asp">detail</a>)</font></p>
            <p align="right">&gt;&gt; <a href="orderdetail.asp"><strong>proses 
              approval</strong></a></p></td>
          <td><p><font color="#006600">Silahkan klik tombol diatas untuk penahanan 
              produk yang akan dikirim.</font></p>
            <p align="right">&gt;&gt; <a href="hold_orderdetail.asp"><strong>detail</strong></a></p></td>
        </tr>
        <tr bgcolor="#666666"> 
          <td bgcolor="#666666">
<div align="left"><a href="accept_prod.asp" onMouseOver="MM_swapImage('Image9','','../img/accept_prod_on.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="../img/accept_prod.jpg" name="Image9" width="150" height="30" border="0" id="Image9"></a></div></td>
          <td bgcolor="#666666">&nbsp;</td>
          <td><div align="left"><a href="accept_appr.asp" onMouseOver="MM_swapImage('Image10','','../img/accept_appr_on.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="../img/accept_appr.jpg" name="Image10" width="150" height="30" border="0" id="Image10"></a></div></td>
        </tr>
        <tr> 
          <td><p><font color="#006600">Silahkan klik tombol diatas untuk penerimaan 
              produk.</font></p>
            <p align="right">&gt;&gt; <a href="list_accept_prod.asp"><strong>detail</strong></a></p></td>
          <td bgcolor="#666666">&nbsp;</td>
          <td><p><font color="#006600">Silahkan klik tombol diatas untuk Approval 
              penerimaan produk.(<a href="detail_accept-approval.asp">detail</a>)</font></p>
            <p align="right">&gt;&gt; <a href="list_accept_prod_Act.asp"><strong>proses 
              approval</strong></a></p></td>
        </tr>
      </table>
      <div align="center"><img src="../img/BAR_ELEG.GIF" width="700" height="3"></div></td>
    <td width="1" height="398"><spacer type="block" width="1" height="398"></td>
  </tr>
  <tr height="1" cntrlrow> 
    <td width="48" height="1"><spacer type="block" width="48" height="1"></td>
    <td width="755" height="1"><div align="center"></div>
      <div align="center"></div>
      <spacer type="block" width="755" height="1"></td>
    <td width="1" height="1"></td>
  </tr>
</table>
<table width="100%" border="2" bordercolor="#FF9900" background="../../img/bg.gif">
  <tr> 
    <td height="25"> <div align="center"> <font color="#009900"> &gt;&gt; <a href="../../index.asp">DEPAN</a> 
        | <a href="../../ADMIN/login.asp">ADMINISTRATOR</a> | <a href="activities.asp">AKTIVITAS</a> 
        | <a href="../Reports/reportListing.asp">LAPORAN</a> | <a href="../Pencarian/Search.asp">PENCARIAN</a></font> 
        <font color="#009900">&lt;&lt;</font></div></td>
  </tr>
  <tr> 
    <td height="21"> <div align="center">Web Master PT. Sierad Produce Tbk </div></td>
  </tr>
</table>
</body>
</html>
