<%@LANGUAGE="VBSCRIPT"%>
 
<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "index.asp"
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
<html>
<head>
<title>Activities :: Sierad </title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="Capex/css/style.css" rel="stylesheet" type="text/css">
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

<body topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('img/depan_on.gif','img/ID_on.gif','img/laporan_on.gif')">
<table width="800" border="1" align="center" bordercolor="#006600">
  <tr bordercolor="#006600"> 
    <td colspan="2"> <div align="left"><img src="Capex/Image/sieradonline.gif" width="222" height="85"> 
      </div>
      <div align="right"><font color="#006600">Date : 
        <script name="current" src="GeneratedItems/current.js" language="JavaScript1.2"></script>
        </font></div></td>
  </tr>
  <tr bordercolor="#006600" bgcolor="#CCCCCC"> 
    <td> <div align="left"><font color="#006600">Selamat Datang </font><font color="#006600"><%= Session("UpdateUser") %></font></div></td>
    <td width="300"> <div align="center"><font color="#009900"><a href="contact.asp"><font color="#006600">Hubungi 
        Kami</font></a></font><font color="#FF0000">&nbsp; </font>| <a href="karir.asp"><font color="#006600">Karir 
        </font></a>| <a href="link.asp"><font color="#006600">Links </font></a>| 
        <font color="#006600"><a href="<%= MM_Logout %>">Log Out</a></font></div></td>
  </tr>
</table>
<table width="800" border="0" align="center" bordercolor="#FF6600" bgcolor="#006600">
  <tr>
    <td><div align="center"><img src="BREEDING/img/spacer.gif" width="795" height="10"></div></td>
  </tr>
</table>
<table width="800" border="2" align="center" bordercolor="#CCCCCC" bgcolor="#006600">
  <tr> 
    <td width="150" height="23"><div align="center"><a href="main_page.asp" onMouseOver="MM_swapImage('Image1','','img/depan_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/depan.gif" name="Image1" width="150" height="20" border="0" id="Image1"></a></div></td>
    <td rowspan="5" bgcolor="#006600"> <div align="left"> 
        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="323" height="122" align="middle">
          <param name="movie" value="Capex/Animasi/anakayam.swf">
          <param name="quality" value="high">
          <param name="SCALE" value="exactfit">
          <embed src="Capex/Animasi/anakayam.swf" width="323" height="122" align="middle" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" scale="exactfit"></embed></object>
        <font color="#FFFFFF"></font><font color="#FFFFFF"></font></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="id_anda.asp" onMouseOver="MM_swapImage('Image2','','img/ID_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/ID.gif" name="Image2" width="150" height="20" border="0" id="Image2"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="komentar.asp" onMouseOver="MM_swapImage('Image3','','img/laporan_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/laporan.gif" name="Image3" width="150" height="20" border="0" id="Image3"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"></div></td>
  </tr>
  <tr> 
    <td height="24"> <div align="center"></div></td>
  </tr>
</table>
<table width="800" border="0" align="center" bordercolor="#FF6600" bgcolor="#006600">
  <tr>
    <td><div align="center"><img src="BREEDING/img/spacer.gif" width="795" height="10"></div></td>
  </tr>
</table>
<table width="800" border="1" align="center" background="business/img/bg.gif">
  <tr>
    <td height="1530"> 
      <div align="center"> 
        <table width="100%" border="2" bordercolor="#009966">
          <tr bordercolor="1"> 
            <td width="54%" height="16"><table width="500" border="0" align="center">
                <tr> 
                  <td height="14" bgcolor="#006600"> <div align="center"><font color="#FF9900">PERSEPSI</font></div></td>
                </tr>
              </table>
              <table width="600" border="0" align="center">
                <tr>
                  <td><font color="#FF6600" size="+2">PERSEPSI</font> 
                    <p><br>
                      Alkisah tersebutlah dua orang peri, satu yang senior dan 
                      satu yang yunior. Kedua peri ini bertugas mengunjungi desa-desa 
                      di bumi. Suatu hari mereka mampir di suatu desa dan menginap 
                      di rumah seorang kaya yang pelit. Karena ketamakan dan kepelitannya, 
                      walau rumahnya besar dan banyak kamar mewah, kedua peri 
                      ini hanya diijinkan tidur di ruang bawah tanah di rumah 
                      tersebut.</p>
                    <p><br>
                      Saat bangun tidur pagi, kedua peri ini mendapati sebuah 
                      lubang di salah satu dinding ruang bawah tanah tersebut. 
                      Oleh peri senior, lubang itu ditambal dan ditutup dengan 
                      rapi. Setelah pergi dari rumah orang kaya itu, peri yunior 
                      kesal dengan sikap peri senior. &#8220;Kakak ini bagaimana 
                      sih? Dia jahat, tapi kakak malah membantu memperbaiki ruang 
                      bawah tanahnya?&#8221; katanya ketus.</p>
                    <p><br>
                      Peri senior menanggapinya dengan tenang, &#8220;Semuanya 
                      tidak seperti yang terlihat&#8221;.</p>
                    <p><br>
                      Malam berikutnya, keduanya menginap di rumah seorang keluarga 
                      petani miskin yang ramah dan baik hati. Keluarga petani 
                      miskin ini memberikan kamar anaknya kepada kedua peri itu 
                      dan mereka tidur sekamar.</p>
                    <p><br>
                      Pada saat bangun pagi, kedua peri ini mendapatkan keluarga 
                      petani ini sedang meratap sedih. Satu-satunya sapi perah 
                      tumpuan penghasilan mereka yang utama telah mati.</p>
                    <p><br>
                      Peri yunior tidak bisa lagi menyembunyikan kekesalannya. 
                      &#8220;Nah, sekarang kita bertemu petani yang baik hati, 
                      tapi kakak malah membiarkan sapi mereka mati&#8221; katanya 
                      pada peri senior. Peri senior kembali menjawab dengan senyum, 
                      &#8220;Semuanya tidak seperti yang terlihat&#8221;. </p>
                    <p><br>
                      Setelah memberi penghiburan, kedua peri itupun berangkat 
                      lagi. Melihat peri yunior begitu penasaran, peri seniorpun 
                      memberi penjelasan. &#8220;Pada saat menemukan lubang di 
                      ruang bawah tanah si orang kaya, saya melihat tambang emas 
                      yang tersembunyi. Orang kaya itu tidak berhati mulia, jadi 
                      saya menutupi lubang tersebut agar dia tidak menemukan tambang 
                      emas itu. Dan malam saat kita menginap di rumah petani, 
                      malaikat maut datang untuk mengambil nyawa istri si petani. 
                      Saya mencegahnya dan setelah sepakat, dia mengambil nyawa 
                      sapi sebagai gantinya. Jadi semuanya tidak seperti yang 
                      terlihat&#8217;.</p>
                    <p><br>
                      Peri yunior pun mengangguk paham.</p>
                    <p><br>
                      Kita bisa langsung menangkap pesan moral yang terkandung 
                      dari dongeng di atas. </p>
                    <p><br>
                      Sebuah dongeng yang memperlihatkan bahwa persepsi keliru 
                      sangat mudah timbul.</p>
                    <p><br>
                      Persepsi merupakan pandangan kita terhadap realitas berdasarkan 
                      hal-hal yang subyektif, misalnya pengalaman dan nilai dari 
                      kita. Persepsi selalu </p>
                    <p></p>
                    <p><br>
                      berdampingan dengan berbagai hal yang terjadi di sekeliling 
                      kita. Seringkali kita terdorong untuk mengambil kesimpulan 
                      dan bertindak berdasarkan persepsi kita, tanpa melihat sisi-sisi 
                      yang lain. Pengalamanlah yang akan mengajarkan banyak hal 
                      kepada kita mengenai persepsi kita yang dangkal.</p>
                    <p><br>
                      Orang bijak berkata bahwa tidak ada kemutlakan mengenai 
                      yang benar dan salah. Yang ada hanyalah persepsi masing-masing 
                      orang mengenai yang benar dan salah.</p>
                    <p><br>
                      Ungkapan ini memang terbukti dalam kehidupan sehari-hari 
                      dengan terjadinya begitu banyak kesalahpahaman dan perselisihan.</p>
                    <p><br>
                      Contoh paling mudah adalah persepsi tentang isu keadilan 
                      dalam dunia kerja. Perusahaan akan merasa sudah memberikan 
                      keadilan dengan adanya kesepakatan kerja awal yang disetujui 
                      dengan karyawan. Tapi setelah karyawan masuk bekerja, dengan 
                      berbagai hal yang dialami, didengar dan dipersepsikannya, 
                      dia bisa jadi akan merasa tidak diperlakukan dengan adil. 
                      Dia pun mulai menuntut. Dalam ukuran massal, demo pun digelar. 
                      Apa yang terjadi? Kedua belah pihak sama-sama merasa benar!</p>
                    <p><br>
                      Contoh lain adalah perusahaan seringkali merasa telah memenuhi 
                      kebutuhan karyawan dengan berbagai fasilitas dan tentu saja 
                      uang. Padahal karyawan juga butuh apresiasi tulus, kepercayaan, 
                      rasa aman dan sebagainya. Kalau tidak percaya akan ketulusan 
                      perusahaan, pemberian sesuatu bahkan apresiasi sekalipun 
                      justru akan dicibir, karena dipersepsikan sebagai sesuatu 
                      yang tidak tulus dan mencurigakan. Sekali lagi, sama-sama 
                      mereka benar. Contoh lain lagi, karena kurangnya pengalaman 
                      dan kedewasaan, karyawan seringkali membandingkan gajinya 
                      dengan temannya di perusahaan lain yang mempunyai posisi 
                      sama. Walau ini perbandingan apple to apple di satu sisi 
                      &#8211; bahwa mereka punya posisi sama, di sisi lain bisa 
                      jadi apple to orange &#8211; tingkat skill mereka dan kondisi 
                      perusahaan mereka bisa berbeda. Selain itu, deal mereka 
                      khan juga berbeda?</p>
                    <p></p>
                    <p>Sumber : MDI News edisi Juli 2002</p>
                    <p></p>
                    <p></p>
                    <p>Jadi bolak-baliklah semua lembaran yang dapat Anda temui 
                      sebelum mengambil kesimpulan. Cek dan ricek! Kalau memang 
                      takut salah dan ingin efektif, kuncinya adalah komunikasi 
                      yang baik dan keterbukaan. Keterbukaan akan melahirkan kepercayaan 
                      dan mempersempit celah kemungkinan pengambilan persepsi 
                      yang negatif. Kalaupun terjadi perbedaan pendapat &#8211; 
                      dengan kepercayaan di antara kedua belah pihak &#8211; win-win 
                      solution lebih bisa dicapai.</p>
                    <p><br>
                    </p></td>
                </tr>
              </table>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              </td>
          </tr>
        </table>
        <p>&nbsp;</p>
      </div></td>
  </tr>
</table>
<table width="800" border="1" align="center" background="business/img/bg.gif">
  <tr> 
    <td><img src="img/bannerrg.jpg" width="790" height="20"></td>
  </tr>
</table>
</body>
</html>
