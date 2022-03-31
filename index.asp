<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%> 
<!--#include file="Connections/simConn.asp" -->
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Request.QueryString
MM_valUsername=CStr(Request.Form("p_username2"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization=""
  MM_redirectLoginSuccess="main_page.asp"
  MM_redirectLoginFailed="fail.asp"
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_simConn_STRING
  MM_rsUser.Source = "SELECT NamaPemesan, Password"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM dbo.tb_Pemesan WHERE NamaPemesan='" & Replace(MM_valUsername,"'","''") &"' AND Password='" & Replace(Request.Form("p_password2"),"'","''") & "'"
  MM_rsUser.CursorType = 0
  MM_rsUser.CursorLocation = 2
  MM_rsUser.LockType = 3
  MM_rsUser.Open
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
	Session("updateuser") = Session("MM_Username")
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

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
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>:: Sierad :: - Microsoft</title>




</script>
<link href="style.css" rel="stylesheet" type="text/css">

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

<script type='text/javascript'>
	function Pop_Go(){return}
	function PopMenu(a,b){return}
	function OutMenu(a){return}
</script>

<script type='text/javascript' src='exmplpopmenu_var.js'></script>
<script type='text/javascript' src='popmenu_com.js'></script>


<link href="style.css" rel="stylesheet" type="text/css">
<body background="img/bg.gif" class="bhs2" onLoad=Pop_Go();MM_preloadImages('img/menu_on_2.jpg','img/menu_on_3.jpg','img/menu_on_4.jpg','img/menu_on_5.jpg','img/menu_on_6.jpg','img/menu_on_7.jpg','img/menu_on_1.jpg')>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="178"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="178" height="93">
        <param name="movie" value="Animasi/logo.swf">
        <param name="quality" value="high"><param name="SCALE" value="exactfit">
        <embed src="Animasi/logo.swf" width="178" height="93" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" scale="exactfit"></embed></object></td>
    <td width="150" background="img/bg_top2.jpg"><img src="img/bg_top2.jpg" width="150" height="93"></td>
    <td width="211" background="img/bg_top.jpg"><img src="img/bg_top.jpg" width="211" height="93"></td>
    <td background="img/bg_top3.jpg"><img src="img/bg_top3.jpg" width="1" height="93"></td>
    <td width="469"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="469" height="93">
        <param name="movie" value="Animasi/tagline.swf">
        <param name="quality" value="high"><param name="SCALE" value="exactfit">
        <embed src="Animasi/tagline.swf" width="469" height="93" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" scale="exactfit"></embed></object></td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="10" valign="top" background="img/bg_line_top.gif"><img src="img/spacer.gif" width="1" height="10"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="200"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><a href="corporate/index.asp" onMouseOut="OutMenu('PopMenu2');MM_swapImgRestore()" onMouseOver="PopMenu('PopMenu2',event);MM_swapImage('Image8','','img/menu_on_1.jpg',1)"><img src="img/menu_1.jpg" alt="Tentang Perusahaan" name="Image8" width="200" height="31" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="business/index.asp" onMouseOut="OutMenu('PopMenu1');MM_swapImgRestore()" onMouseOver="PopMenu('PopMenu1',event);MM_swapImage('Image9','','img/menu_on_2.jpg',1)"><img src="img/menu_2.jpg" alt="Struktur Bisnis" name="Image9" width="200" height="31" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="products/index.asp" onMouseOut="OutMenu('PopMenu3');MM_swapImgRestore()" onMouseOver="PopMenu('PopMenu3',event);MM_swapImage('Image10','','img/menu_on_3.jpg',1)"><img src="img/menu_3.jpg" alt="Produk" name="Image10" width="200" height="30" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="news/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image11','','img/menu_on_4.jpg',1)"><img src="img/menu_4.jpg" alt="Berita" name="Image11" width="200" height="31" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="careers/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image12','','img/menu_on_5.jpg',1)"><img src="img/menu_5.jpg" alt="Karir" name="Image12" width="200" height="30" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="report/index.html" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image13','','img/menu_on_6.jpg',1)"><img src="img/menu_6.jpg" alt="Laporan Tahunan" name="Image13" width="200" height="31" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="contact/index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image14','','img/menu_on_7.jpg',1)"><img src="img/menu_7.jpg" alt="Alamat" name="Image14" width="200" height="29" border="0"></a></td>
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
    <td width="509"> <table width="509" border="0" cellspacing="0" cellpadding="0">
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
          <td><a href="eng/index.html" onMouseOver="MM_swapImage('Image1','','img/pic_mid2_on-14.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="img/pic_mid2-14.jpg" alt="Versi Inggris" name="Image1" width="120" height="23" border="0" id="Image1"></a></td>
          <td><a href="#" onMouseOver="MM_swapImage('Image2','','img/pic_mid2_on-15.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="img/pic_mid2-15.jpg" alt="Warta Sierad" name="Image2" width="125" height="23" border="0" id="Image2"></a></td>
        </tr>
      </table></td>
    <td background="img/bg_mid_rg.gif">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="10" valign="top" background="img/bg_line_top.gif"><img src="img/spacer.gif" width="1" height="10"></td>
  </tr>
</table>
<table width="928" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="1"><img src="img/line.gif" width="1" height="25"></td>
    <td background="img/bg_line_bot.gif"><img src="img/spacer.gif" width="301" height="25"></td>
    <td width="1"><img src="img/line.gif" width="1" height="25"></td>
  </tr>
</table>
<table width="732" height="332" border="0" align="center">
  <tr> 
    <td width="355" height="30"> <div align="left"><img src="img/ttg_sierad.gif" width="150" height="28"></div></td>
    <td width="102" rowspan="3"><img src="img/garis.gif" width="1" height="304"></td>
    <td width="3" rowspan="3"><img src="img/garis.gif" width="1" height="304"></td>
    <td width="254"><div align="left"><font color="#003333"><img src="img/ttg_sierad1.gif" width="150" height="28"></font></div></td>
  </tr>
  <tr> 
    <td height="20"><img src="img/garis1.gif" width="300" height="1" align="top"></td>
    <td><img src="img/garis1.gif" width="300" height="1"></td>
  </tr>
  <tr> 
    <td height="163"><div align="justify"><font color="#336666">PT Sierad Produce 
        Tbk adalah gabungan dari 4 perusahaan pada tahun 2001 yang bergerak di 
        satu bidang bisnis utama di bawah naungan Sierad Group. Empat perusahaan 
        ini adalah PT Anwar Sierad Tbk, PT Sierad Produce Tbk, PT Sierad Feedmill 
        dan PT Sierad Grains.<br>
        <br>
        &nbsp;&nbsp;&nbsp;&nbsp;Sierad Produce, dahulu bernama PT. Betara Darma 
        ekspor impor, berdiri pada tanggal 6 September 1985. Nama Sierad mulai 
        digunakan pada tanggal 27 Desember 1996 saat persiapan untuk public listing 
        yang cukup berhasil di Jakarta Stock Exchange. Bisnis utama perusahaan 
        ini meliputi produksi pakan ternak olahan, breeding, produksi anak ayam, 
        kemitraan, rumah potong ayam dan pembuatan produk olahan bernilai tambah 
        lainnya.</font></div></td>
    <td><div align="justify"> 
        <p><font color="#336666">Pendistribusian DOC(<em>Day Old Chick</em>) pada 
          perusahaan ini dilakukan dengan cara yang sistematis dan terencana, 
          sehingga memudahkan penjadwalan pendistribusian DOC.</font></p>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
        <p align="right"><font color="#336666">:: <a href="BREEDING/login.asp">Login</a> 
          ::</font></p>
      </div></td>
  </tr>
  <tr> 
    <td height="20"><img src="img/garis1.gif" width="300" height="1" align="top"></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><img src="img/garis1.gif" width="300" height="1" align="top"></td>
  </tr>
</table>
<table width="928" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="1"><img src="img/line.gif" width="1" height="25"></td>
    <td background="img/bg_line_bot.gif"><div align="right"><img src="img/spacer.gif" width="301" height="25"><font color="#336666"></font></div></td>
    <td width="1"><img src="img/line.gif" width="1" height="25"></td>
  </tr>
</table>
<table width="928" border="0" align="center" bordercolor="#336666">
  <tr> 
    <td colspan="3"><img src="Icon/Green.ico"><font color="#336666">Info</font></td>
  </tr>
  <tr>
    <td colspan="3"><img src="img/garis1.gif" width="120" height="1" align="top"></td>
  </tr>
  <tr> 
    <td width="456" height="270"> <p><font color="#336666">Anda ingin menjadi 
        mitra kami ! klik \<a href="daftar_pemesan.asp">daftar</a>\ untuk registrasi.</font></p>
      <p><font color="#336666">Dapatkan kemudahan dalam pemesanan DOC, produk 
        DOC yang kami sediakan antara lain :</font></p>
      <p><font color="#336666">DOC broiler(pedaging) : Produk ini didistribusikan 
        baik kepasar dan ke mitra peternak untuk diternakkan hingga menjadi ayam 
        broiler atau pedaging.</font></p>
      <p><font color="#336666">DOC layer(petelur): Produk ini didistribusikan 
        kepasar untuk diternakan menjadi ayam petelur</font></p>
      <p><font color="#336666">DOC jantan : Produk ini didistribusikan baik kepasar 
        dan ke mitra peternak</font></p>
      <p>Minimal Pemesanan DOC adalah Lima Ribu Ekor untuk setiap DOC.</p>
      <p>
        <script LANGUAGE="JavaScript">

// set speed of banner (pause in milliseconds between addition of new character)
var speed = 10 

// decrease value to increase speed (must be positive)
// set pause between completion of message and beginning of following message
var pause = 1500 

// increase value to increase pause
// set initial values
var timerID = null
var bannerRunning = false

// create array
var ar = new Array()

// assign the strings to the array's elements
ar[0] = "Selamat Datang "
ar[1] = "Sierad Produce Website"
ar[2] = "Index Page"
ar[3] = "Created By Musawarman"

// assign index of current message
var message = 0

// empty string initialization
var state = ""

// no value is currently being displayed
clearState()

// stop the banner if it is currently running
function stopBanner() {	
	// if banner is currently running	
	if (bannerRunning)		
	// stop the banner		
	clearTimeout(timerID)	
	// timer is now stopped	
	timerRunning = false
}

// start the banner
function startBanner() {	
	// make sure the banner is stopped	
	stopBanner()	
	// start the banner from the current position	
	showBanner()
}

// assign state a string of "0" characters of the length of the current message
function clearState() {	
	// initialize to empty string	
	state = ""	
	// create string of same length containing 0 digits	
	for (var i = 0; i < ar[message].length; ++i) {		
		state += "0"	
	}
}

// display the current message
function showBanner() {	
	// if the current message is done	
	if (getString()) {		
		// increment message		
		message++		
		// if new message is out of range wrap around to first message		
	if (ar.length <= message)			
		message = 0		
		// new message is first displayed as empty string		
		clearState()		
		// display next character after pause milliseconds		
		timerID = setTimeout("showBanner()", pause)	
	} 
	else {		
		// initialize to empty string		
		var str = ""		
		// built string to be displayed (only character selected thus far are displayed)		
	for (var j = 0; j < state.length; ++j) {			
		str += (state.charAt(j) == "1") ? ar[message].charAt(j) : "     "		
	}		
	// partial string is placed in status bar		
	window.status = str		
	// add another character after speed milliseconds		
	timerID = setTimeout("showBanner()", speed)	
	}
}

function getString() {	
	// set variable to true (it will stay true unless proven otherwise)	
	var full = true	
	// set variable to false if a free space is found in string (a not-displayed char)	
	for (var j = 0; j < state.length; ++j) {		
		// if character at index j of current message has not been placed in displayed string		
		if (state.charAt(j) == 0)			
		full = false	
	}	
	// return true immediately if no space found (avoid infinitive loop later)	
	if (full) return true	
	// search for random until free space found (braoken up via break statement)	
	while (1) {		
		// a random number (between 0 and state.length - 1 == message.length - 1)		
		var num = getRandom(ar[message].length)		
		// if free space found break infinitive loop		
	if (state.charAt(num) == "0")			
		break	
	}	
	// replace the 0 character with 1 character at place found	
	state = state.substring(0, num) + "1" + state.substring(num + 1, state.length)	
	// return false because the string was not full (free space was found)	
	return false
}

function getRandom(max) {	
	// create instance of current date	
	var now = new Date()		
	// create a random number (good generator)	
	var num = now.getTime() * now.getSeconds() * Math.random()	
	// cut random number to value between 0 and max - 1, inclusive	
	return num % max
}
startBanner()
// -->
</script>
      </p>
      <p>&nbsp;</p></td>
    <td width="1"><img src="img/garis.gif" width="1" height="304"></td>
    <td width="457"><p><font color="#336666">Anda sudah menjadi mitra kami ?</font></p>
      <p><font color="#336666">Silahkan isikan user name dan password pada form 
        login dibawah ini.</font></p>
      <form name="form1" method="POST" action="<%=MM_LoginAction%>">
        <TABLE cellSpacing=0 cellPadding=2 align=center border=1>
          <TBODY>
            <TR> 
              <TD width="222" class=title2><div align="center"><img src="Icon/mb.gif" width="31" height="16">User 
                  Login </div></TD>
            </TR>
            <TR> 
              <TD class=tdverylight> <TABLE width="100%" 
                  border=1 cellPadding=2 cellSpacing=1 class=tblight>
                  <TBODY>
                    <TR> 
                      <TD align=right>User ID * :</TD>
                      <TD><INPUT 
                        style="BORDER-RIGHT: #999999 1px solid; BORDER-TOP: #333333 1px solid; FONT-SIZE: 11px; BORDER-LEFT: #333333 1px solid; WIDTH: 150px; COLOR: #000000; BORDER-BOTTOM: #999999 1px solid; HEIGHT: 18px; BACKGROUND-COLOR: #ffecec" 
                        name=p_username2> </TD>
                    </TR>
                    <TR> 
                      <TD height="25" align=right>Password * :</TD>
                      <TD><INPUT 
                        style="BORDER-RIGHT: #999999 1px solid; BORDER-TOP: #333333 1px solid; FONT-SIZE: 11px; BORDER-LEFT: #333333 1px solid; WIDTH: 150px; COLOR: #000000; BORDER-BOTTOM: #999999 1px solid; HEIGHT: 18px; BACKGROUND-COLOR: #ffecec" 
                        type=password name=p_password2> </TD>
                    </TR>
                    <TR align=middle> 
                      <TD colSpan=2 height=50><INPUT class=btn type=submit value=Login name=submit2> 
                        <INPUT class=btn type=reset value=Reset name=reset2> </TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
          </TBODY>
        </TABLE>
      </form>
      <p>&nbsp;</p>
      <p>&nbsp;</p></td>
  </tr>
  <tr> 
    <td height="22" colspan="3"><div align="center"><img src="Capex/Image/bannerrg.jpg" width="920" height="20"></div></td>
  </tr>
</table>
</body>
</html>
