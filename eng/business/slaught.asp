<%@LANGUAGE="JAVASCRIPT" CODEPAGE="1252"%>
<!--#include file="../../Connections/DBConn.asp" -->
<%
var rsProfile = Server.CreateObject("ADODB.Recordset");
rsProfile.ActiveConnection = MM_DBConn_STRING;
rsProfile.Source = "SELECT *  FROM dbo.profile  WHERE compid='RPA' and lang='Inggris'";
rsProfile.CursorType = 0;
rsProfile.CursorLocation = 2;
rsProfile.LockType = 1;
rsProfile.Open();
var rsProfile_numRows = 0;
%>
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
var Repeat1__numRows = 2;
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
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
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

<body background="img/bg.gif" leftmargin="0" topmargin="0" onLoad="MM_preloadImages('img/top_mn_on_01.gif','img/top_mn_on_02.gif','img/top_mn_on_03.gif','img/top_mn_on_04.gif','img/top_mn_on_05.gif','img/bg_mid_on__13.jpg','img/bg_mid_on__14.jpg','img/mn_on_1.jpg','img/mn_on_2.jpg','img/mn_on_3.jpg','img/mn_on_4.jpg','img/mn_on_5.jpg','img/mn_on_6.jpg','img/mn_on_7.jpg','img/mn_on_8.jpg')">
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
    <td width="97"><a href="feedmill.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image8','','img/top_mn_on_01.gif',1)"><img src="img/top_mn-01.gif" alt="Feedmill" name="Image8" width="97" height="25" border="0"></a></td>
    <td width="99"><a href="breeding.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image9','','img/top_mn_on_02.gif',1)"><img src="img/top_mn-02.gif" alt="Breeding" name="Image9" width="99" height="25" border="0"></a></td>
    <td width="138"><a href="slaught.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image10','','img/top_mn_on_03.gif',1)"><img src="img/top_mn-03.gif" alt="Slaughterhouse" name="Image10" width="138" height="25" border="0"></a></td>
    <td width="110"><a href="equipment.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image11','','img/top_mn_on_04.gif',1)"><img src="img/top_mn-04.gif" alt="Equipment" name="Image11" width="110" height="25" border="0"></a></td>
    <td width="183"><a href="animal.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image12','','img/top_mn_on_05.gif',1)"><img src="img/top_mn-05.gif" alt="Animal Pharmaceutical" name="Image12" width="183" height="25" border="0"></a></td>
    <td width="140"><img src="img/bg_mn_top_r.gif" width="140" height="25"></td>
    <td background="img/bg_top_r.gif">&nbsp;</td>
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
    <td background="img/bg_yellow.gif">&nbsp;</td>
  </tr>
  <tr> 
    <td><img src="img/bg_mid__8.jpg" width="107" height="40"></td>
    <td><a href="franchise.asp"><img src="img/bg_mid__92.jpg" alt="Franchise Restaurant" width="155" height="40" border="0"></a></td>
    <td><img src="img/bg_mid__10.jpg" width="155" height="40"></td>
    <td><img src="img/bg_mid__11.jpg" width="155" height="40"></td>
    <td><img src="img/bg_mid__12.jpg" width="155" height="40"></td>
    <td><a href="../../business/slaught.asp" onMouseOver="MM_swapImage('Image1','','img/bg_mid_on__13.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="img/bg_mid__13.jpg" alt="Indonesia Version" name="Image1" width="175" height="40" border="0" id="Image1"></a></td>
    <td><a href="#" onMouseOver="MM_swapImage('Image2','','img/bg_mid_on__14.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="img/bg_mid__14.jpg" alt="Sierad News" name="Image2" width="98" height="40" border="0" id="Image2"></a></td>
    <td background="img/bg_yellow.gif">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td><img src="../img/spacer.gif" width="1" height="1"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="60" background="img/bg_mid_menu.gif"><img src="../img/spacer.gif" width="60" height="1"></td>
    <td width="50"><a href="../index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image30','','img/mn_on_1.jpg',1)"><img src="img/mn_1.jpg" alt="Home" name="Image30" width="50" height="30" border="0"></a></td>
    <td width="170"><a href="../corporate/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image31','','img/mn_on_2.jpg',1)"><img src="img/mn_2.jpg" alt="Corporate Overview" name="Image31" width="170" height="30" border="0"></a></td>
    <td width="180"><a href="../business/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image32','','img/mn_on_3.jpg',1)"><img src="img/mn_3.jpg" alt="Business Structure" name="Image32" width="180" height="30" border="0"></a></td>
    <td width="100"><a href="../products/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image33','','img/mn_on_4.jpg',1)"><img src="img/mn_4.jpg" alt="Products" name="Image33" width="100" height="30" border="0"></a></td>
    <td width="70"><a href="../news/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image34','','img/mn_on_5.jpg',1)"><img src="img/mn_5.jpg" alt="News" name="Image34" width="70" height="30" border="0"></a></td>
    <td width="90"><a href="../careers/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image35','','img/mn_on_6.jpg',1)"><img src="img/mn_6.jpg" alt="Careers" name="Image35" width="90" height="30" border="0"></a></td>
    <td width="140"><a href="../report/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image36','','img/mn_on_7.jpg',1)"><img src="img/mn_7.jpg" alt="Annual Report" name="Image36" width="140" height="30" border="0"></a></td>
    <td width="88"><a href="../contact/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image37','','img/mn_on_8.jpg',1)"><img src="img/mn_8.jpg" alt="Contact Us" name="Image37" width="88" height="30" border="0"></a></td>
    <td background="img/bg_mid_menu.gif">&nbsp;</td>
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
    <td width="49" background="img/bg_line.gif"><img src="../img/spacer.gif" width="49" height="40"></td>
    <td width="584" background="img/bg_line.gif"><img src="img/slaught_ttl.gif"></td>
    <td width="353" background="img/bg_line.gif"><img src="img/news_ttl.gif" width="108" height="21"></td>
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
    <td background="img/bg_line_bot.gif"><img src="../img/spacer.gif" width="301" height="25"></td>
    <td width="1"><img src="../img/line.gif" width="1" height="25"></td>
  </tr>
</table>
<table width="988" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="1" background="../img/line.gif"><img src="../img/spacer.gif" width="1" height="25"></td>
    <td width="47" background="img/bg_bottom.gif">&nbsp;</td>
    <td width="561" valign="top" background="img/bg_bottom.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><%=(rsProfile.Fields.Item("content").Value)%></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>
      <br> 
    </td>
    <td width="25" background="img/bg_bottom.gif">&nbsp;</td>
    <td width="1" background="../img/line.gif"><img src="../img/spacer.gif" width="1" height="25"></td>
    <td width="24" background="img/bg_bottom.gif">&nbsp;</td>
    <td width="325" valign="top" background="img/bg_bottom.gif">
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
                  <td><A HREF="../news/News_in.asp?<%= MM_keepNone + ((MM_keepNone!="")?"&":"") + "ID=" + rsNews.Fields.Item("ID").Value %>"><strong><%=(rsNews.Fields.Item("Title").Value)%></strong></A></td>
                </tr>
                <tr> 
                  <td><%=(rsNews.Fields.Item("Clip").Value)%></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td><A HREF="../news/News_in.asp?<%= MM_keepNone + ((MM_keepNone!="")?"&":"") + "ID=" + rsNews.Fields.Item("ID").Value %>"> 
              <div align="right"><%=(rsNews.Fields.Item("lengkap").Value)%></div>
              </A></td>
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
    <td width="41" background="img/bg_bottom.gif">&nbsp;</td>
    <td width="1" background="../img/line.gif"><img src="../img/spacer.gif" width="1" height="25"></td>
  </tr>
</table>
<table width="988" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="1"><img src="../img/line.gif" width="1" height="19"></td>
    <td background="img/bg_line_bot2.gif"><img src="../img/spacer.gif" width="301" height="19"></td>
    <td width="1"><img src="../img/line.gif" width="1" height="19"></td>
  </tr>
</table>
<br>
</body>
</html>
<%
rsProfile.Close();
%>
<%
rsNews.Close();
%>
