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
    <td><a href="../../report/shareholder.asp" onMouseOver="MM_swapImage('Image1','','img/bg_mid_on__13.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="img/bg_mid__13.jpg" alt="Versi Inggris" name="Image1" width="175" height="40" border="0" id="Image1"></a></td>
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
      Credo</a>:: | ::Report to Shareholders:: | ::<a href="YReview.asp">Year 
      in Review</a>:: | ::<a href="outlook2003.asp">Outlook 2003</a>:: | ::<a href="Prospect.asp">Prospect</a>:: 
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
    <td width="561" valign="top" background="../business/img/bg_bottom.gif"><div align="left"><img src="img/mr_AntoniusSujatakc.gif" width="92" height="153" align="left"> 
        <img src="img/mr_Budiardjokc.gif" width="93" height="154" align="left">Dear 
        Shareholders,</div>
      <p align="justify">The Recovery of the Indonesian economy, though to have 
        begun in earnest in 2001, suffered some setbacks in 2002. The contraction 
        of the U.S. economy significantly impacted the growth and performance 
        of markets worldwide. Global economic activity and already waning investor 
        confidence were further dampened by several high profile corporate scandals 
        and bankruptcies in the United States.</p>
      <p align="center"><strong></strong></p>
      <p align="justify">In effort to sustain domestic market activity the Indonesian 
        government introduced several monetary and fiscal policies and continuous 
        progress was made in bank and overseas debt restructuring efforts. A slower 
        than anticipated rebound in consumer confidence and business spending 
        coupled with poor currency and stock market performance as a result of 
        the Bali bombing effected only marginal growth in 2002 despite positive 
        sentiment at the beginning of the year.</p>
      <p align="justify">Despite these and other obstacles, the poultry industry 
        continued its inevitable evolution towards integration and we at Sierad 
        again directed our attention to the fundamentals of our business. The 
        goal adopted by the company in 2001 were pursued in 2002 as Sierad adopted 
        a three-pronged strategy for growth, which we titled &quot;growing from 
        within&quot;. Under the framework of growing from within, three areas 
        were identified for development namely human resources , marketing , and 
        distribution.</p>
      <p align="justify">In 2002 we laid the groundwork for the development of 
        our existing employees by identifying their strengths and weaknesses and 
        working to maximize their skills through training and knowledge transfer. 
        Not only did this effort result in a more competent workforce, but it 
        also served to create a team oriented work environment, conducive to nurturing 
        creativity and personal initiative.</p>
      <p align="justify">The new Sierad logo, adopted in 2001, was officially 
        launched at our annual customer gathering in June 2002 setting the stage 
        for the more aggressive marketing effort to follow. In 2002 Sierad participated 
        in the Indo Livestock exhibition in Bali and Floriade Netherlands at which 
        our marketing team showcased newly released branded premium products and 
        introduced the company to the international buyers and producers.</p>
      <p align="justify">We also focused our attention on costumer relationship 
        management using a more personalized marketing approach in an effort to 
        strengthen our distribution network. 2003 will see the re-launching of 
        several products, the introduction of a number of new products and product 
        argumentation, and the repositioning of existing products. Also in the 
        coming months we will take a more creative approach to marketing with 
        the introduction of a points reward program called Sierad Poin, and the 
        launch of new and more attractive packaging for our feed products as part 
        of our overall strategy to present the new face of Sierad.</p>
      <p align="justify">The lifeblood of any manufacturer, the development of 
        our distribution channels remains vital to our success. In 2002 Sierad 
        continued its efforts to improve its distribution network in order to 
        more effectively service its customers. These improvements were not confined 
        to the distribution of primary processed products to supermarkets and 
        fastfood outlets, but extended across the supply chain to our entire base 
        of rural and cosmopolitan consumers.</p>
      <p align="justify">The strengthening of the Rupiah at the start of the year 
        serve to reduce the cost of production inputs, which had up until 2002 
        risen steadily. Raw materials - particularly the grain comodities used 
        in the production of feeds-constitude a major component of the finished 
        product and are for the most part imported. The Rupiah appreciation at 
        the beginning of 2002 reduced the cost of many imported items in Rupiah 
        terms. However, the aggresive competition offered by a major new entrant 
        in the domestic feed milling industry coupled with reductions in the cost 
        of raw materials fueled aggressive, cutthroat pricing competition, which 
        resulted in generally thinner margins.</p>
      <p align="justify">Actual industry-wide utilization of installed capacity 
        remained at only 60 percent throughout the report period, a legacy of 
        the unbridled expansion prior to the onset of the economic crises in 1997. 
        In light of this, Sierad continued to aggressively cut cost and boost 
        productivity via stricter purchasing and handling controls in order to 
        ensure reasonable inventory levels and reduce handling costs. The company 
        also discharged further costs by implementing better control of production 
        and milling losses.</p>
      <p align="justify">Our affiliation with our primary international grain 
        vendor for commodity purchases was also strengthened through a creative 
        commodity supply arrangement. This effort augmented our funding requirements 
        that have historically sustain our growth. In 2003, Sierad will continue 
        to shore up its purchasing procedures and improve vendor relationships 
        to the end of maintaining acceptable levels of growth.</p>
      <p align="justify">The production of Day Old Chicks (DOC) increased in 2002 
        primarily as a result of internal productivity improvements as opposed 
        to any increase in capacity, allowing us to post modest growth with modest 
        additional capital expenditure. Our decision in 2002 to replace our spent 
        multiple-strain Parent Stock with a single dominant strain along with 
        the application of the single-age flock method in our farms are expected 
        to further improve production. The increased familiarity in breeding management 
        and improved performance analysis capabilities, both direct result of 
        the single strain method, coupled with the use of single age farming techniques 
        will lead to lower mortality rates and more prolific breeding.</p>
      <p align="justify">In 2002 we began the construction of a modest breeding 
        farm in East Java, which will soon be followed by the construction of 
        localized hatchery equipped with excess hatchery machines from our existing 
        hatcheries in West Java. This will allow us to hatch eggs in East Java 
        reducing our dependency on the delivery of DOCs from our hatcheries in 
        Sukabumi and Bogor. The construction of these facilities will serve to 
        support our local milling facilities, and will bring our production closer 
        to a major market enabling us to deliver healthier, better quality DOCs 
        to customers.</p>
      <p align="justify">Broiler and Layer DOC sales were major revenue contributors 
        in 2002 due to unexpected shortfall in the supply of broiler DOCs in the 
        face of increasing demand. As a natural consequence, the price of live 
        birds rose steadily throughout the year and peaked during the holy month 
        of Ramadhan, outpacing 2001 prices and industry expectations.</p>
      <p align="justify">This increase, though a positive development for our 
        kemitraan [contract growing] and commercial farming operations allowing 
        our contract growers to sell live birds produced at much higher prices, 
        had a negative impact on our slaughterhouse, which had to buy the live 
        birds at unexpectedly high prices to honor previous short term supply 
        contracts priced at considerably lower levels.</p>
      <p align="justify">In 2002, the slaughterhouse initiated a productivity 
        improvement program aimed at enhancing the quality and consistency of 
        birds procured. These efforts were met with increased yield, which we 
        are confident will further improve in the coming months.</p>
      <p align="justify">Exports of processed products to Japan from our slaughterhouse 
        that began in 2001 experienced some reversals early on in the year as 
        a result of a stronger Rupiah. Export revenues were further eroded by 
        the steep decline in exports prices due to increased regional competition 
        and short-lived difficulties in releasing exports from Japanese ports. 
        As a result of this downturn in exports, Sierad turn its attention to 
        the domestic market, opting to establish a solid foothold in the domestic 
        market before aggressively pursuing the export market.</p>
      <p align="justify">The Indonesian market remains a predominantly live bird 
        market with birds sold in traditional wet markets accounting for as much 
        as 90 percent of total industry sales to end consumers. Although the evolution 
        of the domestic market from a predominantly live bird market to a primary 
        processed market consisting of carcasses and chicken cut-ups is inevitable, 
        what remains to be seen is the pace at which this development will occur.</p>
      <p align="justify">Growing public awareness and expectation for quality, 
        hygiene and the like will in time provide the stimulus needed to increase 
        the patronage of modern slaughterhouses and their primary processed products. 
        However, aside from current consumer preferences in favor of live birds, 
        there remain other issues that must be addressed in order to better accommodate 
        increased demand for processed products, such as the shortage of cold 
        storage and refrigeration facilities at distribution points and the preference 
        of farmers to breed smaller birds which impedes efficiency in slaughtering.</p>
      <p align="justify">Throughout the year we continue to closed unprofitable 
        Wendy's outlets in order to improve our overall profitabilty. These loss-making 
        outlets were generally situated in office buildings and as such were dependent 
        on tenants for patronage. When the levels of occupancy of the buildings 
        declined - a legacy of the economic crises - the sales and profitability 
        of these outlets also declined. We will continue to monitor those stores, 
        which continue to yield marginal returns and will try to improve their 
        performance by, among others, renegotiating their leases, which now represent 
        an unreasonably large chunk of revenue. Our closer observation of this 
        facet of our business may inevitably lead to the closure of additional 
        stores rendered unprofitable on account of their locations.</p></td>
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
