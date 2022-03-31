<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/DBConn2.asp" -->
<%
Dim rsNews
Dim rsNews_numRows

Set rsNews = Server.CreateObject("ADODB.Recordset")
rsNews.ActiveConnection = MM_DBConn2_STRING
rsNews.Source = "SELECT id,tgl, Title, Clip,lengkap  FROM dbo.News  WHERE lang='Inggris'  ORDER BY tgl DESC"
rsNews.CursorType = 0
rsNews.CursorLocation = 2
rsNews.LockType = 1
rsNews.Open()

rsNews_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 2
Repeat1__index = 0
rsNews_numRows = rsNews_numRows + Repeat1__numRows
%>
<%
Dim MM_paramName 
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
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<script type='text/javascript'>
	function Pop_Go(){return}
	function PopMenu(a,b){return}
	function OutMenu(a){return}
</script>
<script type='text/javascript' src='exmplpopmenu_var.js'></script>
<script type='text/javascript' src='popmenu_com.js'></script>
<body background="img/bg.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="Pop_Go();MM_preloadImages('img/menu_on_1.jpg','img/menu_on_2.jpg','img/menu_on_3.jpg','img/menu_on_4.jpg','img/menu_on_5.jpg','img/menu_on_6.jpg','img/menu_on_7.jpg','img/pic_mid2_on-14.jpg','img/pic_mid2_on-15.jpg')">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="178"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="178" height="93">
        <param name="movie" value="img/logo.swf">
        <param name="quality" value="high">
        <embed src="img/logo.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="178" height="93"></embed></object></td>
    <td width="150" background="img/bg_top2.jpg">&nbsp;</td>
    <td width="211" background="img/bg_top.jpg">&nbsp;</td>
    <td background="img/bg_top3.jpg">&nbsp;</td>
    <td width="469"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="469" height="93">
        <param name="movie" value="img/tagline.swf">
        <param name="quality" value="high">
        <embed src="img/tagline.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="469" height="93"></embed></object></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="10" valign="top" background="img/bg_line_top.gif"><img src="img/spacer.gif" width="1" height="10"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="200"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><a href="corporate/index.asp" onMouseOut="OutMenu('PopMenu2');MM_swapImgRestore()" onMouseOver="PopMenu('PopMenu2',event);MM_swapImage('Image8','','img/menu_on_1.jpg',1)"><img src="img/menu_1.jpg" alt="Corporate Overview" name="Image8" width="200" height="31" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="business/index.asp" onMouseOut="OutMenu('PopMenu1');MM_swapImgRestore()" onMouseOver="PopMenu('PopMenu1',event);MM_swapImage('Image9','','img/menu_on_2.jpg',1)"><img src="img/menu_2.jpg" alt="Business Structure" name="Image9" width="200" height="31" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="products/index.asp" onMouseOut="OutMenu('PopMenu3');MM_swapImgRestore()" onMouseOver="PopMenu('PopMenu3',event);MM_swapImage('Image10','','img/menu_on_3.jpg',1)"><img src="img/menu_3.jpg" alt="Products" name="Image10" width="200" height="30" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="news/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image11','','img/menu_on_4.jpg',1)"><img src="img/menu_4.jpg" alt="News" name="Image11" width="200" height="31" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="careers/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image12','','img/menu_on_5.jpg',1)"><img src="img/menu_5.jpg" alt="Careers" name="Image12" width="200" height="30" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="report/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image13','','img/menu_on_6.jpg',1)"><img src="img/menu_6.jpg" alt="Annual Report" name="Image13" width="200" height="31" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="contact/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image14','','img/menu_on_7.jpg',1)"><img src="img/menu_7.jpg" alt="Contact Us" name="Image14" width="200" height="29" border="0"></a></td>
        </tr>
      </table></td>
    <td width="291"><table width="291" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="102"><img src="img/pic_mid_1.jpg" width="102" height="95"></td>
          <td width="63"><img src="img/pic_mid_2.jpg" width="63" height="95"></td>
          <td><img src="img/pic_mid_3.jpg" width="63" height="95"></td>
          <td><img src="img/pic_mid_4.jpg" width="63" height="95"></td>
        </tr>
        <tr>
          <td><img src="img/pic_mid_5.jpg" width="102" height="118"></td>
          <td><img src="img/pic_mid_6.jpg" width="63" height="118"></td>
          <td><img src="img/pic_mid_7.jpg" width="63" height="118"></td>
          <td><img src="img/pic_mid_8.jpg" width="63" height="118"></td>
        </tr>
      </table></td>
    <td width="509">
<table width="509" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="img/pic_mid2-01.jpg" width="88" height="95"></td>
          <td><img src="img/pic_mid2-02.jpg" width="88" height="95"></td>
          <td><img src="img/pic_mid2-03.jpg" width="88" height="95"></td>
          <td><img src="img/pic_mid2-04.jpg" width="120" height="95"></td>
          <td><img src="img/pic_mid2-05.jpg" width="125" height="95"></td>
        </tr>
        <tr> 
          <td><img src="img/pic_mid2-06.jpg" width="88" height="95"></td>
          <td><img src="img/pic_mid2-07.jpg" width="88" height="95"></td>
          <td><img src="img/pic_mid2-08.jpg" width="88" height="95"></td>
          <td><img src="img/pic_mid2-09.jpg" width="120" height="95"></td>
          <td><img src="img/pic_mid2-10.jpg" width="125" height="95"></td>
        </tr>
        <tr>
          <td><img src="img/pic_mid2-11.jpg" width="88" height="23"></td>
          <td><img src="img/pic_mid2-12.jpg" width="88" height="23"></td>
          <td><img src="img/pic_mid2-13.jpg" width="88" height="23"></td>
          <td><a href="../ina_index2.asp" onMouseOver="MM_swapImage('Image1','','img/pic_mid2_on-14.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="img/pic_mid2-14.jpg" alt="Indonesia Version" name="Image1" width="120" height="23" border="0" id="Image1"></a></td>
          <td><a href="#" onMouseOver="MM_swapImage('Image2','','img/pic_mid2_on-15.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="img/pic_mid2-15.jpg" alt="Sierad News" name="Image2" width="125" height="23" border="0" id="Image2"></a></td>
        </tr>
      </table></td>
    <td background="img/bg_mid_rg.gif">&nbsp;</td>
  </tr>
</table>
<table width="928" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="img/spacer.gif" width="1" height="8"></td>
  </tr>
</table>
<table width="928" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="1"><img src="img/line.gif" width="1" height="40"></td>
    <td width="16" background="img/bg_line.gif"><img src="img/spacer.gif" width="16" height="40"></td>
    <td width="119" background="img/bg_line.gif"><img src="img/front_subtitle_01.jpg" width="119" height="13"></td>
    <td width="301" background="img/bg_line.gif"><img src="img/spacer.gif" width="301" height="40"></td>
    <td width="108" background="img/bg_line.gif">&nbsp;</td>
    <td width="180" background="img/bg_line.gif"><img src="img/spacer.gif" width="180" height="40"></td>
    <td width="160" background="img/bg_line.gif"><img src="img/front_subtitle_02.jpg" width="108" height="13"></td>
    <td width="42" background="img/bg_line.gif"><img src="img/spacer.gif" width="42" height="40"></td>
    <td width="1"><img src="img/line.gif" width="1" height="40"></td>
  </tr>
</table>
<table width="928" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td><img src="img/spacer.gif" width="1" height="1"></td>
  </tr>
</table>
<table width="928" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="1"><img src="img/line.gif" width="1" height="25"></td>
    <td background="img/bg_line_bot.gif"><img src="img/spacer.gif" width="301" height="25"></td>
    <td width="1"><img src="img/line.gif" width="1" height="25"></td>
  </tr>
</table>
<table width="928" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="1" background="img/line.gif"><img src="img/spacer.gif" width="1" height="25"></td>
    <td width="16" background="img/bg_bottom.gif"><img src="img/spacer.gif" width="16" height="25"></td>
    <td width="400" valign="top" background="img/bg_bottom.gif"> 
      <p>&nbsp;&nbsp;&nbsp;&nbsp;PT Sierad Produce Tbk is the entity resulting 
        from the merger in 2001 of four companies conducting the core businesses 
        of the Sierad Group. These were PT Anwar Sierad Tbk, PT Sierad Produce 
        Tbk, and their fully owned subsidiaries, PT Sierad Feedmill and PT Sierad 
        Grains. <br>
        <br>
        &nbsp;&nbsp;&nbsp; Sierad Produce, formerly PT Betara Darma Ekspor Impor, 
        was incorporated on 6 September 1985. Its current name was adopted on 
        27 December 1996 in preparation for its successful public listing on the 
        Jakarta Stock Exchange. Its core businesses include the production of 
        primary processed and poultry feed, breeding, the production of day old 
        chicks, contract farming, slaughtering and the production of further processed 
        value-added products.</p></td>
    <td width="16" background="img/bg_bottom.gif"><img src="img/spacer.gif" width="16" height="25"></td>
    <td width="1" background="img/line.gif"><img src="img/spacer.gif" width="1" height="25"></td>
    <td width="16" background="img/bg_bottom.gif"><img src="img/spacer.gif" width="16" height="25"></td>
    <td width="274" valign="top" background="img/bg_bottom.gif"> <% 
While ((Repeat1__numRows <> 0) AND (NOT rsNews.EOF)) 
%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><em><%= DoDateTime((rsNews.Fields.Item("tgl").Value), 1, 2057) %></em></td>
        </tr>
        <tr> 
          <td><strong><A HREF="news/News_in.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "id=" & rsNews.Fields.Item("id").Value %>"><%=(rsNews.Fields.Item("Title").Value)%></A></strong></td>
        </tr>
        <tr> 
          <td><div align="left"><%=(rsNews.Fields.Item("Clip").Value)%></div></td>
        </tr>
        <tr> 
          <td><div align="right"><A HREF="news/News_in.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "id=" & rsNews.Fields.Item("id").Value %>"><%=(rsNews.Fields.Item("lengkap").Value)%></A></div></td>
        </tr>
        <tr> 
          <td> <div align="left"></div></td>
        </tr>
      </table>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsNews.MoveNext()
Wend
%> <br></td>
    <td width="16" background="img/bg_bottom.gif"><img src="img/spacer.gif" width="16" height="25"></td>
    <td width="1" background="img/line.gif"><img src="img/spacer.gif" width="1" height="25"></td>
    <td width="16" background="img/bg_bottom.gif"><img src="img/spacer.gif" width="16" height="25"></td>
    <td width="10" valign="top" background="img/bg_bottom.gif">&nbsp; </td>
    <td width="1" background="img/line.gif"><img src="img/spacer.gif" width="1" height="25"></td>
  </tr>
</table>
<table width="928" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="1"><img src="img/line.gif" width="1" height="19"></td>
    <td background="img/bg_line_bot2.gif"><img src="img/spacer.gif" width="301" height="19"></td>
    <td width="1"><img src="img/line.gif" width="1" height="19"></td>
  </tr>
</table>
<br>

</body>
</html>
<%
rsNews.Close()
Set rsNews = Nothing
%>
