<%@LANGUAGE="JAVASCRIPT" CODEPAGE="1252"%>
<!--#include file="../../Connections/DBConn.asp" -->
<%
var rsNews = Server.CreateObject("ADODB.Recordset");
rsNews.ActiveConnection = MM_DBConn_STRING;
rsNews.Source = "SELECT ID,  thumbnail,tgl,Title, Clip,  Lengkap  FROM dbo.News where lang='Inggris'  ORDER BY tgl DESC";
rsNews.CursorType = 0;
rsNews.CursorLocation = 2;
rsNews.LockType = 1;
rsNews.Open();
var rsNews_numRows = 0;
%>
<%
var Repeat1__numRows = 3;
var Repeat1__index = 0;
rsNews_numRows += Repeat1__numRows;
%>
<% var MM_paramName = ""; %>
<%
// *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

// create the list of parameters which should not be maintained
var MM_removeList = "&index=";
if (MM_paramName != "") MM_removeList += "&" + MM_paramName.toLowerCase() + "=";
var MM_keepURL="",MM_keepForm="",MM_keepBoth="",MM_keepNone="";

// add the URL parameters to the MM_keepURL string
for (var items=new Enumerator(Request.QueryString); !items.atEnd(); items.moveNext()) {
  var nextItem = "&" + items.item().toLowerCase() + "=";
  if (MM_removeList.indexOf(nextItem) == -1) {
    MM_keepURL += "&" + items.item() + "=" + Server.URLencode(Request.QueryString(items.item()));
  }
}

// add the Form variables to the MM_keepForm string
for (var items=new Enumerator(Request.Form); !items.atEnd(); items.moveNext()) {
  var nextItem = "&" + items.item().toLowerCase() + "=";
  if (MM_removeList.indexOf(nextItem) == -1) {
    MM_keepForm += "&" + items.item() + "=" + Server.URLencode(Request.Form(items.item()));
  }
}

// create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL + MM_keepForm;
if (MM_keepBoth.length > 0) MM_keepBoth = MM_keepBoth.substring(1);
if (MM_keepURL.length > 0)  MM_keepURL = MM_keepURL.substring(1);
if (MM_keepForm.length > 0) MM_keepForm = MM_keepForm.substring(1);
%>
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>					
function DoDateTime(str, nNamedFormat, nLCID)				
	dim strRet								
	dim nOldLCID								
										
	strRet = str								
	If (nLCID > -1) Then							
		oldLCID = Session.LCID						
	End If									
										
	On Error Resume Next							
										
	If (nLCID > -1) Then							
		Session.LCID = nLCID						
	End If									
										
	If ((nLCID < 0) Or (Session.LCID = nLCID)) Then				
		strRet = FormatDateTime(str, nNamedFormat)			
	End If									
										
	If (nLCID > -1) Then							
		Session.LCID = oldLCID						
	End If									
										
	DoDateTime = strRet							
End Function									
</SCRIPT>									
<html>
<head>
<title>:: Sierad ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../style.css" rel="stylesheet" type="text/css">
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
//-->
</script>
</head>

<body background="../business/img/bg.gif" leftmargin="0" topmargin="0" onLoad="MM_preloadImages('img/bg_mid_on__13.jpg','img/bg_mid_on__14.jpg','../business/img/mn_on_1.jpg','../business/img/mn_on_2.jpg','../business/img/mn_on_3.jpg','../business/img/mn_on_4.jpg','../business/img/mn_on_5.jpg','../business/img/mn_on_6.jpg','../business/img/mn_on_7.jpg','../business/img/mn_on_8.jpg')">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="178"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="178" height="93">
        <param name="movie" value="../img/logo.swf">
        <param name="quality" value="high">
        <embed src="../img/logo.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="178" height="93"></embed></object></td>
    <td width="150" background="../img/bg_top2.jpg">&nbsp;</td>
    <td width="211" background="../img/bg_top.jpg">&nbsp;</td>
    <td background="../img/bg_top3.jpg">&nbsp;</td>
    <td width="469"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="469" height="93">
        <param name="movie" value="../img/tagline.swf">
        <param name="quality" value="high">
        <embed src="../img/tagline.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="469" height="93"></embed></object></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="233"><img src="img/bg_mn_top.gif" width="233" height="25"></td>
    <td width="627"><img src="img/bg_mn.jpg" width="627" height="25"></td>
    <td width="140"><img src="img/bg_mn_top_r.gif" width="140" height="25"></td>
    <td background="../careers/img/bg_top_r.gif">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="107"><img src="img/bg_mid__1.jpg" width="107" height="55"></td>
    <td width="155"><img src="img/bg_mid__2.jpg" width="155" height="55"></td>
    <td width="155"><img src="img/bg_mid__3.jpg" width="155" height="55"></td>
    <td width="155"><img src="img/bg_mid__4.jpg" width="155" height="55"></td>
    <td width="155"><img src="img/bg_mid__5.jpg" width="155" height="55"></td>
    <td width="175"><img src="img/bg_mid__6.jpg" width="175" height="55"></td>
    <td width="98"><img src="img/bg_mid__7.jpg" width="98" height="55"></td>
    <td background="../business/img/bg_yellow.gif">&nbsp;</td>
  </tr>
  <tr> 
    <td><img src="img/bg_mid__8.jpg" width="107" height="40"></td>
    <td><img src="img/bg_mid__9.jpg" width="155" height="40" border="0"></td>
    <td><img src="img/bg_mid__10.jpg" width="155" height="40"></td>
    <td><img src="img/bg_mid__11.jpg" width="155" height="40"></td>
    <td><img src="img/bg_mid__12.jpg" width="155" height="40"></td>
    <td><a href="../../contact/index.asp" onMouseOver="MM_swapImage('Image1','','img/bg_mid_on__13.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="img/bg_mid__13.jpg" alt="Indonesia Version" name="Image1" width="175" height="40" border="0" id="Image1"></a></td>
    <td><a href="#" onMouseOver="MM_swapImage('Image2','','img/bg_mid_on__14.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="img/bg_mid__14.jpg" alt="Sierad News" name="Image2" width="98" height="40" border="0" id="Image2"></a></td>
    <td background="../business/img/bg_yellow.gif">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td><img src="../img/spacer.gif" width="1" height="1"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="60" background="../business/img/bg_mid_menu.gif"><img src="../img/spacer.gif" width="60" height="1"></td>
    <td width="50"><a href="../index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image30','','../business/img/mn_on_1.jpg',1)"><img src="../business/img/mn_1.jpg" alt="Home" name="Image30" width="50" height="30" border="0"></a></td>
    <td width="170"><a href="../corporate/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image31','','../business/img/mn_on_2.jpg',1)"><img src="../business/img/mn_2.jpg" alt="Corporate Overview" name="Image31" width="170" height="30" border="0"></a></td>
    <td width="180"><a href="../business/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image32','','../business/img/mn_on_3.jpg',1)"><img src="../business/img/mn_3.jpg" alt="Business Structure" name="Image32" width="180" height="30" border="0"></a></td>
    <td width="100"><a href="../products/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image33','','../business/img/mn_on_4.jpg',1)"><img src="../business/img/mn_4.jpg" alt="Products" name="Image33" width="100" height="30" border="0"></a></td>
    <td width="70"><a href="../news/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image34','','../business/img/mn_on_5.jpg',1)"><img src="../business/img/mn_5.jpg" alt="News" name="Image34" width="70" height="30" border="0"></a></td>
    <td width="90"><a href="../careers/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image35','','../business/img/mn_on_6.jpg',1)"><img src="../business/img/mn_6.jpg" alt="Careers" name="Image35" width="90" height="30" border="0"></a></td>
    <td width="140"><a href="../report/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image36','','../business/img/mn_on_7.jpg',1)"><img src="../business/img/mn_7.jpg" alt="Annual Report" name="Image36" width="140" height="30" border="0"></a></td>
    <td width="88"><a href="../contact/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image37','','../business/img/mn_on_8.jpg',1)"><img src="../business/img/mn_8.jpg" alt="Contact Us" name="Image37" width="88" height="30" border="0"></a></td>
    <td background="../business/img/bg_mid_menu.gif">&nbsp;</td>
  </tr>
</table>
<table width="928" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td><img src="../img/spacer.gif" width="1" height="8"></td>
  </tr>
</table>
<table width="988" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="1"><img src="../img/line.gif" width="1" height="40"></td>
    <td width="49" background="../business/img/bg_line.gif"><img src="../img/spacer.gif" width="49" height="40"></td>
    <td width="584" background="../business/img/bg_line.gif">&nbsp;</td>
    <td width="353" background="../business/img/bg_line.gif"><img src="../business/img/news_ttl.gif" width="108" height="21"></td>
    <td width="1"><img src="../img/line.gif" width="1" height="40"></td>
  </tr>
</table>
<table width="928" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td><img src="../img/spacer.gif" width="1" height="1"></td>
  </tr>
</table>
<table width="988" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="1"><img src="../img/line.gif" width="1" height="25"></td>
    <td background="../business/img/bg_line_bot.gif"><img src="../img/spacer.gif" width="301" height="25"></td>
    <td width="1"><img src="../img/line.gif" width="1" height="25"></td>
  </tr>
</table>
<table width="988" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="1" background="../img/line.gif"><img src="../img/spacer.gif" width="1" height="25"></td>
    <td width="47" background="../business/img/bg_bottom.gif">&nbsp;</td>
    <td width="561" valign="top" background="../business/img/bg_bottom.gif"> <p><b>Head 
        Office</b><br>
        Jl. Kemang Raya No. 67 <br>
        Jakarta 12730<br>
        Tel. (021) 7193888 <br>
        Fax. (021) 7193889 <br>
        <br>
        <b>Breeding Division</b><br>
        Jl. Raya Parung KM 20 Parung<br>
        Bogor - Jawa Barat<br>
        Tel. (0251) 611143<br>
        Fax. (0251) 611655<br>
        <br>
        <br>
        <b>Contract Farming Divisions</b><br>
        Bojong Sari No. 26 Sawangan,<br>
        Bogor-Jawa Barat<br>
        Tel. (0251) 619789<br>
        Fax. (0251) 619790<br>
        <br>
        Jl. Raya Sidoarjo - Krian Desa Ketimang,<br>
        Wonoayu-Sidoarjo, Jawa Timur<br>
        Tel. (031) 8852804 - 06<br>
        Fax. (031) 8852810<br>
        <br>
        Jl. Patimura KM 1 Canden<br>
        Salatiga 50711 Jawa Tengah<br>
        Tel. (0298) 327353 <br>
        Fax.(0298) 313492<br>
        <br>
        <b>Feed Mill Divisions<br>
        </b>Jl. Raya Serang KM 13 Balaraja,<br>
        Tangerang - Banten <br>
        Tel. (021) 5953888 <br>
        Fax.(021) 5950150 <br>
        <br>
        Jl. Raya Sidoarjo - Krian Desa Ketimang, <br>
        Wonoayu-Sidoarjo, Jawa Timur <br>
        Tel. (031) 8852804 - 06 <br>
        Fax. (031) 8852810 <br>
        <br>
        Jl. Ir. Sutami KM 12 Desa Suka Negara Tanjung Bintang <br>
        Lampung Selatan <br>
        Tel. (0721) 351175 <br>
        Fax. (0721) 351173<br>
        <br>
        <b>Slaughterhouse Divisions<br>
        </b>Jl. Raya Parung KM 19 Parung, <br>
        Bogor-Jawa Barat <br>
        Tel. (0251) 611862 <br>
        Fax. (0251) 611079<br>
        <br>
        Jl. Patimura KM 1 Canden <br>
        Salatiga 50711 Jawa Tengah <br>
        Tel. (0298) 327353 <br>
        Fax. (0298) 313492<br>
        <br>
        <b>PT Sierad Industries</b><br>
        Poultry Equipment<br>
        Jl. Modern Industri No. 24 <br>
        Kawasan Industri Modern Cikande - Serang, Banten <br>
        Tel. (0254) 402536 <br>
        Fax. (0254) 402538<br>
        <br>
        <b>PT Sierad Biotek</b><br>
        Animal Pharmaceutical &amp; Health Care Products<br>
        Jl. M.H. Thamrin Blok A-10 No. 3 <br>
        Lippo Cikarang, Bekasi 17550 - Jawa Barat <br>
        Tel. (021) 8972264 - 66 <br>
        Fax. (021) 8972268 <br>
        <br>
        <b>PT Dwipamina Nusantara</b> <br>
        Fishmeal Product<br>
        Dusun Ketapang Desa Pengambengan Negara - Jembrana, Bali <br>
        Tel. (0365) 42147 <br>
        Fax.(0365) 42148 <br>
        <br>
        <b>PT Sierad Pangan</b> <br>
        Hartz Restaurant<br>
        Jl. Raya Kebayoran Lama No. 220 <br>
        Jakarta Selatan 12220 <br>
        Tel. (021) 5360482 <br>
        Fax. (021) 5357932<br>
        <br>
        <b>PT Wendy Citarasa</b> <br>
        Wendy's Fast Foods<br>
        Jl. Raya Kebayoran Lama No. 220 <br>
        Jakarta Selatan 12220 <br>
        Tel. (021) 5350481 <br>
        Fax. (021) 5357932<br>
        <br>
    </td>
    <td width="25" background="../business/img/bg_bottom.gif">&nbsp;</td>
    <td width="1" background="../img/line.gif"><img src="../img/spacer.gif" width="1" height="25"></td>
    <td width="24" background="../business/img/bg_bottom.gif">&nbsp;</td>
    <td width="325" valign="top" background="../business/img/bg_bottom.gif">
<div align="right"> 
        <% while ((Repeat1__numRows-- != 0) && (!rsNews.EOF)) { %>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="56" valign="top"><img src="<%=(rsNews.Fields.Item("thumbnail").Value)%>"></td>
            <td width="2" valign="top"> <div align="center"></div></td>
            <td width="252" valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td><%= DoDateTime((rsNews.Fields.Item("tgl").Value), 1, 2057) %></td>
                </tr>
                <tr> 
                  <td><strong><A HREF="../news/News_in.asp?<%= MM_keepNone + ((MM_keepNone!="")?"&":"") + "ID=" + rsNews.Fields.Item("ID").Value %>"><%=(rsNews.Fields.Item("Title").Value)%></A></strong></td>
                </tr>
                <tr> 
                  <td><%=(rsNews.Fields.Item("Clip").Value)%></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td valign="top">&nbsp;</td>
            <td>&nbsp;</td>
            <td> <div align="right"><A HREF="../news/News_in.asp?<%= MM_keepNone + ((MM_keepNone!="")?"&":"") + "ID=" + rsNews.Fields.Item("ID").Value %>"><%=(rsNews.Fields.Item("lengkap").Value)%></A></div></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td><div align="right"></div></td>
          </tr>
        </table>
        <%
  Repeat1__index++;
  rsNews.MoveNext();
}
%>
      </div></td>
    <td width="41" background="../business/img/bg_bottom.gif">&nbsp;</td>
    <td width="1" background="../img/line.gif"><img src="../img/spacer.gif" width="1" height="25"></td>
  </tr>
</table>
<table width="988" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="1"><img src="../img/line.gif" width="1" height="19"></td>
    <td background="../business/img/bg_line_bot2.gif"><img src="../img/spacer.gif" width="301" height="19"></td>
    <td width="1"><img src="../img/line.gif" width="1" height="19"></td>
  </tr>
</table>
<br>
</body>
</html>
<%
rsNews.Close();
%>
