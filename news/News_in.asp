<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/simConn.asp" -->
<%
Dim rsNews__MMColParam
rsNews__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  rsNews__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim rsNews
Dim rsNews_numRows

Set rsNews = Server.CreateObject("ADODB.Recordset")
rsNews.ActiveConnection = MM_simConn_STRING
rsNews.Source = "SELECT * FROM dbo.News WHERE ID = " + Replace(rsNews__MMColParam, "'", "''") + ""
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

<body background="../business/img/bg.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../business/img/top_mn_on_01.gif','../business/img/top_mn_on_02.gif','../business/img/top_mn_on_03.gif','../business/img/top_mn_on_04.gif','../business/img/top_mn_on_05.gif','../business/img/mn_on_1.jpg','../business/img/mn_on_2.jpg','../business/img/mn_on_3.jpg','../business/img/mn_on_4.jpg','../business/img/mn_on_5.jpg','../business/img/mn_on_6.jpg','../business/img/mn_on_7.jpg','../business/img/mn_on_8.jpg','../corporate/img/bg_mid_on__13.jpg','../corporate/img/bg_mid_on__14.jpg')">
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
    <td width="233"><img src="../corporate/img/bg_mn_top.gif" width="233" height="25"></td>
    <td width="100"><a href="../corporate/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image8','','../corporate/img/top_mn_on_06.gif',1)"><img src="../corporate/img/top_mn-06.gif" alt="Visi &amp; Misi" name="Image8" width="100" height="25" border="0"></a></td>
    <td width="100"><a href="../corporate/history.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image9','','../corporate/img/top_mn_on_07.gif',1)"><img src="../corporate/img/top_mn-07.gif" alt="Sejarah" name="Image9" width="100" height="25" border="0"></a></td>
    <td width="427"><img src="../corporate/img/bg_mn_top_r1.gif" width="427" height="25"></td>
    <td width="140"><img src="../business/img/bg_mn_top_r.gif" width="140" height="25"></td>
    <td background="../business/img/bg_top_r.gif">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="107"><img src="../corporate/img/bg_mid__1.jpg" width="107" height="55"></td>
    <td width="155"><img src="../corporate/img/bg_mid__2.jpg" width="155" height="55"></td>
    <td width="155"><img src="../corporate/img/bg_mid__3.jpg" width="155" height="55"></td>
    <td width="155"><img src="../corporate/img/bg_mid__4.jpg" width="155" height="55"></td>
    <td width="155"><img src="../corporate/img/bg_mid__5.jpg" width="155" height="55"></td>
    <td width="175"><img src="../corporate/img/bg_mid__6.jpg" width="175" height="55"></td>
    <td width="98"><img src="../corporate/img/bg_mid__7.jpg" width="98" height="55"></td>
    <td background="../business/img/bg_yellow.gif">&nbsp;</td>
  </tr>
  <tr>
    <td><img src="../corporate/img/bg_mid__8.jpg" width="107" height="40"></td>
    <td><img src="../corporate/img/bg_mid__9.jpg" width="155" height="40" border="0"></td>
    <td><img src="../corporate/img/bg_mid__10.jpg" width="155" height="40"></td>
    <td><img src="../corporate/img/bg_mid__11.jpg" width="155" height="40"></td>
    <td><img src="../corporate/img/bg_mid__12.jpg" width="155" height="40"></td>
    <td><a href="../eng/corporate/history.html" onMouseOver="MM_swapImage('Image1','','../corporate/img/bg_mid_on__13.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="../corporate/img/bg_mid__13.jpg" alt="Versi Inggris" name="Image1" width="175" height="40" border="0" id="Image1"></a></td>
    <td><a href="#" onMouseOver="MM_swapImage('Image2','','../corporate/img/bg_mid_on__14.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="../corporate/img/bg_mid__14.jpg" alt="Warta Sierad" name="Image2" width="98" height="40" border="0" id="Image2"></a></td>
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
    <td width="170"><a href="../report/index.html" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image36','','../business/img/mn_on_7.jpg',1)"><img src="../business/img/mn_7.jpg" alt="Laporan Tahunan" name="Image36" width="170" height="30" border="0"></a></td>
    <td width="68"><a href="../contact/index.html" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image37','','../business/img/mn_on_8.jpg',1)"><img src="../business/img/mn_8.jpg" alt="Alamat" name="Image37" width="68" height="30" border="0"></a></td>
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
    <td width="47" background="../business/img/bg_bottom.gif"><img src="../business/img/news_pic.jpg" width="115" height="83"></td>
    <td width="561" valign="top" background="../business/img/bg_bottom.gif"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><div align="justify">&nbsp;&nbsp;<strong><%=(rsNews.Fields.Item("title").Value)%></strong></div></td>
        </tr>
        <tr> 
          <td><div align="justify"><%=(rsNews.Fields.Item("lengkap").Value)%></div></td>
        </tr>
      </table>
      <p>&nbsp;</p></td>
    <td width="25" background="../business/img/bg_bottom.gif">&nbsp;</td>
    <td width="1" background="../img/line.gif"><img src="../img/spacer.gif" width="1" height="25"></td>
    <td width="24" background="../business/img/bg_bottom.gif">&nbsp;</td>
    <td width="325" valign="top" background="../business/img/bg_bottom.gif"> 
      <div align="right"> 
        <% 
While ((Repeat1__numRows <> 0) AND (NOT rsNews.EOF)) 
%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="56" valign="top">&nbsp;</td>
            <td width="2" valign="top"> <div align="center"></div></td>
            <td width="252" valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td><%= DoDateTime((rsNews.Fields.Item("Tanggal").Value), 1, 2057) %></td>
                </tr>
                <tr> 
                  <td><A HREF="News_in.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ID=" & rsNews.Fields.Item("ID").Value %>"><%=(rsNews.Fields.Item("title").Value)%></A></td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td> <div align="right"></div>
              </td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td><div align="right"></div></td>
          </tr>
        </table>
        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsNews.MoveNext()
Wend
%>
      </div></td>
    <td width="41" background="../business/img/bg_bottom.gif">&nbsp;</td>
    <td width="1" background="../img/line.gif"><img src="../img/spacer.gif" width="1" height="25"></td>
  </tr>
</table>
<table width="988" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="1"><img src="../img/line.gif" width="1" height="19"></td>
    <td background="../business/img/bg_line_bot2.gif"><img src="../img/spacer.gif" width="301" height="19">
      <INPUT name="button" type=button class=btn onclick=history.back() value=Back></td>
    <td width="1"><img src="../img/line.gif" width="1" height="19"></td>
  </tr>
</table>
<br>

</body>
</html>
<%
rsNews.Close()
Set rsNews = Nothing
%>


