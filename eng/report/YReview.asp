<%@LANGUAGE="JAVASCRIPT"%>
<!--#include file="../../Connections/DBConn.asp" -->
<%
var rsNews = Server.CreateObject("ADODB.Recordset");
rsNews.ActiveConnection = MM_DBConn_STRING;
rsNews.Source = "SELECT ID,  thumbnail,tgl,Title, Clip,  Lengkap  FROM dbo.News  WHERE lang='Inggris'  ORDER BY tgl DESC";
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
<script language="JavaScript" type="text/JavaScript">
<!--
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

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<link href="../style.css" rel="stylesheet" type="text/css">
</head>

<body background="../business/img/bg.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../business/img/top_mn_on_01.gif','../business/img/top_mn_on_02.gif','../business/img/top_mn_on_03.gif','../business/img/top_mn_on_04.gif','../business/img/top_mn_on_05.gif','../business/img/mn_on_1.jpg','../business/img/mn_on_2.jpg','../business/img/mn_on_3.jpg','../business/img/mn_on_4.jpg','../business/img/mn_on_5.jpg','../business/img/mn_on_6.jpg','../business/img/mn_on_7.jpg','../business/img/mn_on_8.jpg','img/bg_mid_on__13.jpg','img/bg_mid_on__14.jpg')">
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
        <embed src="../img/tagline.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="469" height="93"></embed></object></td>  </tr>
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
    <td><a href="../../report/yReview.asp" onMouseOver="MM_swapImage('Image1','','img/bg_mid_on__13.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="img/bg_mid__13.jpg" alt="Versi Inggris" name="Image1" width="175" height="40" border="0" id="Image1"></a></td>
    <td><a href="#" onMouseOver="MM_swapImage('Image2','','img/bg_mid_on__14.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="img/bg_mid__14.jpg" alt="Warta Sierad" name="Image2" width="98" height="40" border="0" id="Image2"></a></td>
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
    <td width="200"><a href="../corporate/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image31','','../business/img/mn_on_2.jpg',1)"><img src="../business/img/mn_2.jpg" alt="Tentang Perusahaan" name="Image31" width="200" height="30" border="0"></a></td>
    <td width="150"><a href="../business/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image32','','../business/img/mn_on_3.jpg',1)"><img src="../business/img/mn_3.jpg" alt="Struktur Bisnis" name="Image32" width="150" height="30" border="0"></a></td>
    <td width="90"><a href="../products/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image33','','../business/img/mn_on_4.jpg',1)"><img src="../business/img/mn_4.jpg" alt="Produk" name="Image33" width="90" height="30" border="0"></a></td>
    <td width="80"><a href="../news/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image34','','../business/img/mn_on_5.jpg',1)"><img src="../business/img/mn_5.jpg" alt="Berita" name="Image34" width="80" height="30" border="0"></a></td>
    <td width="80"><a href="../careers/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image35','','../business/img/mn_on_6.jpg',1)"><img src="../business/img/mn_6.jpg" alt="Karir" name="Image35" width="80" height="30" border="0"></a></td>
    <td width="170"><a href="../report/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image36','','../business/img/mn_on_7.jpg',1)"><img src="../business/img/mn_7.jpg" alt="Laporan Tahunan" name="Image36" width="170" height="30" border="0"></a></td>
    <td width="68"><a href="../contact/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image37','','../business/img/mn_on_8.jpg',1)"><img src="../business/img/mn_8.jpg" alt="Alamat" name="Image37" width="68" height="30" border="0"></a></td>
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
    <td width="584" background="../business/img/bg_line.gif">::<a href="index.asp">Our 
      Credo</a>:: | ::<a href="shareholder.asp">Report to Shareholders</a>:: | 
      ::Year in Review:: | ::<a href="outlook2003.asp">Outlook 2003</a>:: | ::<a href="Prospect.asp">Prospect</a>:: 
      |::<a href="future.asp">Our Future</a>:: | ::<a href="commisioners.asp">Commisioners 
      &amp; Directors</a>:: | ::<a href="FinancialReport.pdf">Financial Report</a>:: 
    </td>
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
    <td width="561" background="../business/img/bg_bottom.gif"> <div align="justify"><strong>Net 
        Sales </strong></div>
      <p align="justify">Total net sales rose modestly from IDR 1.308 trillion 
        to 1.316 trillion in 2002. Sales of feed prod ucts, including internal 
        consumption rose by 2.3 percent from 418,700 metric tons to 428,400 metric 
        tons. DOC sales rose by 5.4 percent from 80.52 million to 84.85 million 
        due to improvements in the productivity of existing facilities and Parent 
        Stock population as we continued to restrain any increase in production 
        capacity during the year. Unfortunately the revenue gains resulting from 
        increased sales of feeds and DOC were negated by a drop in raw material 
        sales, the cessation of live bird trading and a decline in sales generated 
        by our fast foods division due to the continued closure of unprofitable 
        outlets. The effect of the potentially adverse VAT regulation on live 
        bird sales introduced in 2001 was fully felt in 2002, as we continued 
        to withdraw from the trading of live birds.</p>
      <p align="justify"><strong>Gross Profit</strong></p>
      <p align="justify">Gross profit rose 17 percent from IDR 113.479 million 
        to IDR 132.714 billion. The increase in Gross profit is notable given 
        the only marginal growth in sales, and is due in part to the unusually 
        high prices of DOCs in 2002 on account of an unanticipated shortfall in 
        industry-wide production. As a consequence, live bird prices also rose 
        allowing us to sell our feeds at higher than planned prices. The improved 
        profit margin is also a direct result of internal productivity improvements 
        and cost reduction initiatives and is remarkable when viewed in the context 
        of the significant appreciation of the Rupiah - from IDR 10,400 to IDR 
        8,940 per USD and from IDR 7.915.68 to IDR 7.539.54 per JPY - as this 
        prompted a rapid decline in the replacement costs of imported commodities 
        and other raw materials compelling many feed mills to mark down the selling 
        prices of their higher costing inventories of raw materials to stave off 
        further losses or declines in value. </p>
      <p align="justify">Perhaps even more remarkable is the fact that higher 
        profit margins were achieved in 2002 despite the losses incured by the 
        slaughterhouse in its difficult transition from the export market to the 
        domestic market in the wake of declining export prices and stiff price 
        competition from other poultry exporters. The unusually high prices of 
        inputs, notably live birds, in the second semester caused operating losses 
        in fullfiling previously established 3-6 month supply contracts to major 
        institutional buyers at fixed prices. However, these losses are unusual 
        and not normal and such should not be repeated in 2003. Moreover, our 
        slaughterhouse operations provide hedge and windfall profits on committed 
        orders in the event of declining live bird prices.</p>
      <p align="justify"><strong>Operating Expenses</strong></p>
      <p align="justify">Operating expenses rose by IDR 9.078 billion, to IDR 
        110.858 billion in 2002 from IDR 98.372 billion the previous year due 
        primarily to a one-time increase in promotional expenditure. Other costs 
        also followed an upward trend, notably transport, fuel, and utilities 
        in response to the government's gradual withdrawal of longstanding subsidies.</p>
      <p align="justify"><strong>Operating Income</strong></p>
      <p align="justify">Operating income improved by 44.68 percent, from IDR 
        15.106 billion in 2001 to IDR 21.855 billion in 2002.</p>
      <p align="justify"><strong>Net Income</strong></p>
      <p align="justify">Income before income tax rose to IDR 127.735 billion 
        in marked contrast to a net pre-tax loss of IDR 33.763 billion in 2001. 
        This development was due to FOREX gains on account of a stronger Rupiah 
        and a reduction in fees and expenses associated with the administration 
        of the company's restructured loans. The release of net tax losses carried 
        forward and expiring in 2002 amounting to IDR 202.126 billion resulted 
        in a net loss after tax of IDR 74.369 billion, a marked improvement from 
        the IDR 300.761 billion net loss posted in 2001.</p>
      <p align="justify"><strong>Financial Ratios</strong></p>
      <p align="justify">Net shareholder's equity improved by almost 100 percent 
        from IDR 38.760 billion to IDR 74.781 billion in 2002 reflecting the reappraisal 
        of certain assets from their historical costs. Other assets, heretofore 
        recorded at their historical costs, will be revalued in the next few years 
        in order to more fairly reflect the true asset value of the company. The 
        company's financial ratios also improved significantly within the report 
        period : the current ratio improved from 2.1:1 in 2001 to 3.4:1 by year's 
        end 2002; the debt to equity ratio improved by over half, from 26.8 times 
        in 2001 to 12.6 times in 2002 and finally, total assets to total debts 
        remained largely unchanged, at 1.22:1 from 1.26:1.</p>
      <p align="justify"><strong>Net Working Capital</strong><br>
        Confirming the company&#8217;s improved financial ratios, net working 
        capital grew by 17 percent, from IDR 265 billion at the end of fiscal 
        year 2001 to IDR 311.62 billion as of the end of 2002. Cash on hand remained 
        largely unchanged at IDR 29 billion, while cash deposits on commodity 
        purchases largely with Toepfer, improved from IDR 28.13 billion to IDR 
        47.90 billion.</p>
      <p align="justify"></p></td>
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
