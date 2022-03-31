<%@LANGUAGE="VBSCRIPT"%>
 
<%
' Proses Logout
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "index.asp"
  ' Pengalihan URL
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
    <td height="1692"> 
      <div align="center"> 
        <table width="100%" height="1626" border="2" bordercolor="#009966">
          <tr bordercolor="1"> 
            <td width="54%" height="1618"> 
              <table width="500" border="0" align="center">
                <tr> 
                  <td height="14" bgcolor="#006600"> <div align="center"><font color="#FF9900">PEMBOROSAN</font></div></td>
                </tr>
              </table>
              <table width="600" border="0" align="center">
                <tr>
                  <td height="1158"><font color="#FF6600" size="+2">KIAT MENGATASI 
                    PEMBOROSAN DI KANTOR</font> 
                    <p><br>
                      Kata orang waktu adalah uang. Pemborosan waktu berarti pemborosan 
                      uang pula. Tetapi ternyata tidak mudah untuk mengukur kerugian 
                      yang disebabkan oleh pemborosan waktu. Lebih buruk lagi, 
                      tidak banyak karyawan menyadari betapa mereka sebenarnya 
                      telah melakukan pemborosan waktu. Banyak dari kita yang 
                      menganggap apa yang kita lakukan adalah hal-hal yang biasa 
                      saja, padahal justru menyumbangkan pemborosan waktu yang 
                      tidak sedikit.</p>
                    <p><br>
                      Pemborosan waktu terasa pada saat kita tidak mampu menyelesaikan 
                      tugas-tugas yang semestinya kita selesaikan. Kita merasa 
                      kehabisan waktu dan bekerja seolah dikejar-kejar waktu. 
                      Padahal kita telah menyia-nyiakan sumber daya waktu tanpa 
                      kita sadari. Para karyawan yang pintar selalu mengatur pekerjaan. 
                      Berikut ini ada beberapa kiat sederhana untuk mengatasi 
                      pemborosan waktu.</p>
                    <p><br>
                      Tiba di kantor tepat waktu </p>
                    <p>Banyak karyawan tiba di kantor beberapa menit terlambat. 
                      Di kota-kota besar, lalu lintas yang padat menjadi alasan 
                      pembenaran bagi seorang karyawan untuk datang terlambat. 
                      Padahal begitu Anda tiba di kantor, Anda tidak bisa langsung 
                      kerja (beberapa wanita harus menyempatkan diri ke kamar 
                      kecil untuk merapikan riasan, menyalakan komputer, menyeduh 
                      kopi atau teh, membaca koran pagi, membuka internet dan 
                      sebagainya). Tanpa terasa Anda kehilangan waktu sekitar 
                      15 menit atau mungkin lebih setiap paginya. Keterlambatan 
                      adalah pemborosan waktu yang paling utama. Anda bisa menyelamatkan 
                      waktu yang sangat berharga ini dengan datang ke kantor beberapa 
                      menit sebelum waktunya.</p>
                    <p><br>
                      Susun prioritas penyelesaian tugas-tugas </p>
                    <p>Jangan langsung mengerjakan ini-itu, tetapi sempatkan beberapa 
                      menit di awal pagi Anda untuk menyusun daftar prioritas 
                      penyelesaian tugas-tugas. Letakkan tugas yang harus Anda 
                      kerjakan terlebih dahulu dan penting di urutan paling atas. 
                      Lalu usahakan Anda mengikuti daftar tersebut. Mungkin Anda 
                      tidak selalu bisa mengikuti daftar prioritas yang telah 
                      ada, tetapi itu merupakan alat yang baik untuk mengingatkan 
                      Anda. Karena itu jangan kehilangan fleksibilitas.</p>
                    <p><br>
                      Jangan menunda-nunda pekerjaan </p>
                    <p>Perhatikan daftar prioritas yang Anda siapkan, apakah Anda 
                      menuliskan tugas yang sama berulang-ulang setiap harinya? 
                      Bila ya, itu berarti Anda telah menunda-nunda pekerjaan. 
                      Menunda pekerjaan adalah pemborosan waktu yang tersembunyi. 
                      Bila Anda tidak mampu menyelesaikan sebuah tugas yang sulit, 
                      jangan berhenti untuk mencari-cari alasan pembenarannya, 
                      tetapi mulailah mengerjakan tugas lain yang lebih mudah. 
                      Bila Anda telah menemukan jalan keluar bagi tugas Anda yang 
                      terhenti, maka kerjakanlah. Jangan biarkan penundaan suatu 
                      pekerjaan menunda pekerjaan yang lain.</p>
                    <p><br>
                      Hindari sikap &#8216;perfectionism&#8217;</p>
                    <p>Ada orang yang tidak puas dengan lay-out table yang sudah 
                      dia buat, lalu ia mengubah-ubah dan mencari bentuk yang 
                      paling &#8220;sempurna&#8221;. Tanpa sadar ia telah menghabiskan 
                      puluhan lembar kertas untuk mencetak table tersebut. Tanpa 
                      sadar pula, keinginannya untuk &#8220;sempurna&#8221; telah 
                      menyia-nyiakan waktunya. Perfectionism, bisa menjadi pemboros 
                      waktu yang sangat besar. Kadang perfectionism melanda pada 
                      hal-hal yang kecil, seperti tanda baca, kalimat, bentuk 
                      garis dan lain-lain yang tidak substantive. Daripada mengejar 
                      kesempurnaan lebih baik Anda kerjakan sesuatu yang terbaik, 
                      lalu selesaikan dan kerjakan tugas-tugas lain.</p>
                    <p><br>
                      Awasi jam ngobrol </p>
                    <p>Semua orang tahu, ngobrol adalah pemboros waktu terbesar 
                      namun paling disukai. Mungkin pada awalnya Anda hanya berniat 
                      membicarakan beberapa hal urusan pekerjaan dengan kolega, 
                      staf atau atasan Anda. Tetapi seringkali tidak sadar Anda 
                      melencengkan percakapan ke hal-hal lain yang menghabiskan 
                      waktu lebih banyak dari pada urusan pekerjaan itu sendiri. 
                      Karena itu, sekarang perhatikan perilaku Anda yang memicu 
                      pemborosan waktu karena obrolan kecil seperti saat Anda 
                      menyeduh kopi bersama teman-teman, memfoto kopi, mengirim 
                      fax atau lain-lain.</p>
                    <p><br>
                      Kendalikan waktu Anda di telepon </p>
                    <p>Interupsi telepon juga pemborosan waktu, bila Anda tidak 
                      mampu menanganinya dengan baik. Bila Anda sedang sibuk-sibuknya 
                      mengerjakan sesuatu, mintalah operator untuk mengalihkan 
                      telepon yang masuk atau mintalah staf Anda untuk menangani 
                      telepon tersebut. Kendalikan telepon sebelum telepon mengendalikan 
                      Anda.</p>
                    <p><br>
                      Atur penanganan e-mail dan surat-surat </p>
                    <p>Kemajuan teknologi juga berdampak pada pemborosan waktu. 
                      Internet sekarang turut menyumbang kerugian waktu bagi pekerja. 
                      Juga e-mail yang masuk. Semua orang tahu bahwa berkomunikasi 
                      melalui e-mail tidak membutuhkan balasan secepat kilat. 
                      Kita tahu tidak semua orang berada pada posisi on-line, 
                      jadi mereka bisa menunggu.</p>
                    <p><br>
                      Demikian beberapa saran sederhana untuk mengatasi pemborosan 
                      waktu di kantor. Semoga bermanfaat dan sukses selalu untuk 
                      Anda. END(sumber E-mail)</p>
                    </td>
                </tr>
              </table>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              </td>
          </tr>
        </table>
        <p>&nbsp;</p>
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
