<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/simConn.asp" -->
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Request.QueryString
MM_valUsername=CStr(Request.Form("p_username"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization="UserLevel"
  MM_redirectLoginSuccess="index.asp"
  MM_redirectLoginFailed="fail.asp"
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_simConn_STRING
  MM_rsUser.Source = "SELECT UserID, UserPassword"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM dbo.UserMS WHERE UserID='" & Replace(MM_valUsername,"'","''") &"' AND UserPassword='" & Replace(Request.Form("p_password"),"'","''") & "'"
  MM_rsUser.CursorType = 0
  MM_rsUser.CursorLocation = 2
  MM_rsUser.LockType = 3
  MM_rsUser.Open
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
	Session("UpdateUser") = Session("MM_Username")
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<html>
<head>
<title>:: Sierad : Login ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
</head>

<body background="../business/img/bg.gif" leftmargin="0" topmargin="0">
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
<table width="100%" border="0">
  <tr>
    <td><img src="img/spacer.gif" width="1125" height="10"></td>
  </tr>
</table>
<table width="100%" border="1" bordercolor="#009966">
  <tr bordercolor="1" bgcolor="#cd853f"> 
    <td width="46%" height="16" bgcolor="#009900"><font color="#FFFFFF">:: BREEDING 
      ::</font></td>
    <td width="54%" bgcolor="#009900"><font color="#FFFFFF">:: LOGIN ::</font></td>
  </tr>
  <tr bordercolor="1" bgcolor="#cd853f">
    <td height="23" bgcolor="#009900"><img src="img/LOOKER1.GIF" width="200" height="12"></td>
    <td bgcolor="#009900"><img src="img/LOOKER1.GIF" width="200" height="12"></td>
  </tr>
</table>
<table width="87%" border="0" align="center">
  <tr> 
    <td width="45%"><div align="justify">Breeding farm yang kami miliki terletak 
        jauh dari area pemukiman di daerah Jawa Barat yang sejuk. Pabrik tempat 
        menampung ayam betina dewasa dirancang sesuai sistem closed-house dengan 
        peralatan modern seperti kipas pendingin, mesin pakan otomatis dan sistem 
        nipple drinking. Sierad Produce menjalankan 12 breeding farm yang sebagian 
        besar terletak di Jawa Barat. Total produksi dari seluruh lahan tersebut 
        mencapai lebih dari 120 juta anak ayam (DOC) per tahun, memenuhi kebutuhan 
        internal bisnis Sierad Produce dan pasar eksternal baik dalam negeri maupun 
        ekspor.<br>
        <br>
        <b> PRODUK<br>
        </b>Broiler DOC <br>
        Layer DOC <br>
        Male DOC <br>
      </div></td>
    <td width="1%"><img src="../img/garis.gif" width="1" height="248"></td>
    <td width="54%"><form name="form1" method="POST" action="<%=MM_LoginAction%>">
        <TABLE cellSpacing=0 cellPadding=2 align=center border=1>
          <TBODY>
            <TR> 
              <TD class=title2><div align="center"><img src="../Icon/sign.gif" width="29" height="25">User 
                  Login </div></TD>
            </TR>
            <TR> 
              <TD class=tdverylight> <TABLE class=tblight cellSpacing=1 cellPadding=2 width="100%" 
                  border=1>
                  <TBODY>
                    <TR> 
                      <TD align=right>Username * :</TD>
                      <TD><INPUT 
                        style="BORDER-RIGHT: #999999 1px solid; BORDER-TOP: #333333 1px solid; FONT-SIZE: 11px; BORDER-LEFT: #333333 1px solid; WIDTH: 150px; COLOR: #000000; BORDER-BOTTOM: #999999 1px solid; HEIGHT: 18px; BACKGROUND-COLOR: #ffecec" 
                        name=p_username> </TD>
                    </TR>
                    <TR> 
                      <TD align=right>Password * :</TD>
                      <TD><INPUT 
                        style="BORDER-RIGHT: #999999 1px solid; BORDER-TOP: #333333 1px solid; FONT-SIZE: 11px; BORDER-LEFT: #333333 1px solid; WIDTH: 150px; COLOR: #000000; BORDER-BOTTOM: #999999 1px solid; HEIGHT: 18px; BACKGROUND-COLOR: #ffecec" 
                        type=password name=p_password> </TD>
                    </TR>
                    <TR align=middle> 
                      <TD colSpan=2 height=50><INPUT class=btn type=submit value=Login name=submit> 
                        <INPUT class=btn type=reset value=Reset name=reset> </TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
          </TBODY>
        </TABLE>
      </form></td>
  </tr>
  <tr> 
    <td colspan="3">&nbsp;</td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<table width="100%" border="0" background="../img/bg.gif">
  <tr> 
    <td height="25"> <div align="center"> <a href="../corporate/index.asp"><font color="#009900">Home</font></a><font color="#009900"> 
        |</font> <a href="../business/index.asp"><font color="#009900">Struktur 
        Bisnis</font></a> <font color="#009900">|</font> <a href="../products/index.asp"><font color="#009900">Produk</font></a><font color="#009900"> 
        |</font> <a href="../news/index.asp"><font color="#009900">Berita</font></a><font color="#009900"> 
        |</font> <a href="../careers/index.asp"><font color="#009900">Karir</font></a><font color="#009900"> 
        </font> <font color="#009900">| </font> <a href="../report/index.html"><font color ="#009900">Laporan 
        Tahunan</font></a> <font color="#009900">| <a href="../contact/index.asp"><font color="#009900">Alamat</font></a></font></div></td>
  </tr>
  <tr> 
    <td height="14"> <div align="center">Web Master PT. Sierad Produce Tbk </div></td>
  </tr>
</table>
</body>
</html>
