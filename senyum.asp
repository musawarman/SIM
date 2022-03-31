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
    <td height="1675"> 
      <div align="center"> 
        <table width="100%" border="2" bordercolor="#009966">
          <tr bordercolor="1"> 
            <td width="54%" height="16"><table width="500" border="0" align="center">
                <tr> 
                  <td height="14" bgcolor="#006600"> <div align="center"><font color="#FF9900">SENYUM 
                      DULU DONG</font></div></td>
                </tr>
              </table>
              <table width="600" border="0" align="center">
                <tr>
                  <td><font color="#FF6600" size="+2">SANG RAJA</font> <p><br>
                      Di sebuah hutan terdapat raja hutan (singa) yang merasa 
                      dirinya paling hebat. </p>
                    <p><br>
                      Untuk melegalisasikan kehebatannya, singa tersebut berencana 
                      untuk bertanya kepada semua penghuni hutan.</p>
                    <p><br>
                      Proses dimulai, dia memanggil semua rakyatnya (kecuali gajah, 
                      karena sedang sakit gigi) dan mulai memberikan pertanyaan. 
                    </p>
                    <p><br>
                      Pertanyaan pertama ditujukan pada gorilla, katanya : &#8220;Hai 
                      gorilla, siapakah yang paling gagah di hutan ini?&#8221;</p>
                    <p><br>
                      Gorila : &#8220;Anda tuanku&#8221;</p>
                    <p><br>
                      Mendengar jawaban itu, banggalah si singa mendengarnya.</p>
                    <p><br>
                      Kemudian dia bertanya pada banteng : &#8220;Hai banteng, 
                      siapakah yang paling gagah dan hebat di hutan ini?&#8221;</p>
                    <p><br>
                      Banteng : &#8220;Tentu saja Anda-lah tuan&#8221;</p>
                    <p><br>
                      Demikian seterusnya singa bertanya kepada seluruh binatang 
                      yang berkumpul di situ. Dan semuanya memberikan jawaban 
                      yang sama, bahwa Singa-lah yang paling kuat, hebat dan gagah.</p>
                    <p><br>
                      Maka semakin sombonglah si singa mendengar jawaban itu.</p>
                    <p><br>
                      Namun dia masih penasaran, karena belum mendengar pendapat 
                      gajah yang sedang sakit.</p>
                    <p><br>
                      Maka dengan gagahnya dia datang ke rumah gajah dan memberikan 
                      pertanyaan yang sama pula.</p>
                    <p><br>
                      Namun gajah tetap diam dan tidak mau menjawab.</p>
                    <p><br>
                      Singa kesal dan bertanya lagi dengan nada gusar.</p>
                    <p><br>
                      Di luar dugaan si singa, gajah itu langsung menghajar dan 
                      menginjak-injak singa hingga pingsan.</p>
                    <p><br>
                      Setelah dia siuman, dengan badan babak belur, singa berkata 
                      kepada gajah : </p>
                    <p><br>
                      &#8220;Gajah&#8230;gajah&#8230;kalau kamu nggak tahu jawabannya, 
                      jangan marah gitu dong&#8230;&#8230;..(kata singa sambil 
                      menahan sakit)&#8221;</p>
                    <p></p>
                    <p></p>
                    <p>AYAM KAKI TIGA</p>
                    <p></p>
                    <p><br>
                      Seorang pria, sebut saja Doni, sedang mengendarai mobilnya 
                      di jalan pinggir kota.</p>
                    <p><br>
                      Setelah beberapa saat berlalu, dari kaca spion dia melihat 
                      ada ayam berlari-lari di belakang mobilnya.</p>
                    <p><br>
                      Doni terheran-heran melihat ayam yang dapat berlari mengejar 
                      mobilnya yang berkecepatan 50 km/jam</p>
                    <p><br>
                      Doni menambah kecepatan mobilnya menjadi 60 km/jam, dan 
                      ayam itu masih terus membuntutinya, bahkan sudah berada 
                      di samping mobilnya.</p>
                    <p><br>
                      Ditambah lagi kecepatannya menjadi 72 km/jam, ternyata ayam 
                      itu berhasil melampaui mobil Doni. </p>
                    <p><br>
                      Setelah diperhatikan ayam tersebut ternyata berkaki tiga.</p>
                    <p><br>
                      Saking penasaran, Doni mengikuti ke mana larinya ayam itu</p>
                    <p><br>
                      Akhirnya dia sampai di sebuah peternakan.</p>
                    <p><br>
                      Dia keluar dari mobilnya dan melihat bahwa semua ayam di 
                      peternakan itu berkaki tiga.</p>
                    <p><br>
                      Doni bertemu dengan si empunya peternakan dan bertanya : 
                      &#8220;Apa yang terjadi dengan ayam-ayam ini?&#8221;</p>
                    <p><br>
                      Peternak itu menjawab : &#8216;Uhmmmm&#8230;karena banyaknya 
                      orang yang suka makan dengan sop kaki ayam. Untuk memenuhi 
                      banyaknya permintaan konsumen akan kaki ayam maka saya mengembangkan 
                      rekayasa genetik pada induk ayam-ayam tersebut hingga dapat 
                      mempunyai keturunan berkaki tiga.&#8221;</p>
                    <p><br>
                      Doni bertanya lagi : &#8220;Bagaimana dengan rasa ayam itu, 
                      apakah tidak berubah oleh rekayasa Anda?&#8221;</p>
                    <p><br>
                      Peternak itu menjawab : &#8220;Nah di situlah letak masalahnya&#8230;&#8230;.karena 
                      sampai saat ini saya belum berhasil menangkap satu ekor-pun 
                      untuk dijadikan sample rasa&#8230;.&#8221;</p>
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
