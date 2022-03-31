<%@LANGUAGE="VBSCRIPT"%> 

<!--#include file="../../Capex/Connections/CapexConn.asp" -->
<html>
<head>
<title>::Pencarian ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../Capex/css/style.css" rel="stylesheet" type="text/css">
</head>

<body background="../../Capex/Image/bg.gif">
<table width="750" border="0" align="center">
  <tr> 
    <td width="916" height="20"><table width="100%" border="0" cellspacing="0" cellpadding="0">
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
      </table></td>
  </tr>
  <tr> 
    <td height="20"><div align="center"><font color="#0000FF">Tanggal :</font> 
        <%=date %> </div></td>
  </tr>
  <tr> 
    <td height="14"><div align="center"><font color="#FF0000">Selamat Datang</font> 
        <%= Session("UpdateUser") %></div></td>
  </tr>
  <tr> 
    <td height="14"><div align="center"><strong><font color="#0000FF">Silahkan 
        lakukan pencarian pada menu di samping kiri Anda</font></strong></div></td>
  </tr>
  <tr>
    <td height="14"><div align="center">
        <table width="600" border="0">
          <tr> 
            <td>&nbsp;</td>
          </tr>
        </table>
    </div></td>
  </tr>
</table>
<div align="left"> </div>
</body>
</html>

