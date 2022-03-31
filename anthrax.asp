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
    <td height="1530"> 
      <div align="center"> 
        <table width="100%" border="2" bordercolor="#009966">
          <tr bordercolor="1"> 
            <td width="54%" height="16"><table width="500" border="0" align="center">
                <tr> 
                  <td height="14" bgcolor="#006600"> <div align="center"><font color="#FF9900">ANTHRAX</font></div></td>
                </tr>
              </table>
              <table border="0" cellpadding="7" cellspacing="7" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2">
                <tr> 
                  <td width="49%" valign="top"> <p class="MsoBodyText" style="text-indent: .25in; text-align: justify; font-family: Arial Narrow; margin-left: 0in; margin-right: 0in; margin-top: 0in; margin-bottom: .0001pt"><span class="content"><span lang="IN"><font size="-5">Nama 
                      <i>Anthrax</i> kembali akrab di telinga kita, akibat pemberitaan 
                      yang gencar oleh media massa. Tidak hanya karena kasus bom 
                      biologis di Amerika, namun juga karena kemunculannya di 
                      berbagai tempat di Indonesia, seperti kejadian di Purwakarta 
                      (Januari 2000), Bogor (Januari 2001) dan yang terbaru di 
                      Kebon Pedes, Bogor (Agustus 2002).</font></span></span></p>
                    <p class="MsoNormal" align="justify"><font size="-5"><span class="content"><i><span lang="IN" style="font-family: Arial Narrow">Anthrax 
                      </span></i><span lang="IN" style="font-family: Arial Narrow">sebenarnya 
                      sudah dikenal lama dalam sejarah manusia. Menurut catatan 
                      pertama penyakit ini sudah muncul sejak jaman Mesir Kuno. 
                      Penyakit ini pada akhirnya sudah menyebar di lima benua 
                      (Afrika, Eropa, Asia, Amerika dan Australia).</span></span></font></p>
                    <p class="MsoBodyTextIndent" style="text-align: justify; text-indent: .25in; font-family: Arial Narrow; margin-left: 0in; margin-right: 0in; margin-top: 0in; margin-bottom: .0001pt"><font size="-5"><span class="content"><span lang="IN">Di 
                      Indonesia, <i>anthrax</i> pertama kali diberitakan oleh 
                      <i>Javasche Courant</i>, karena menjangkiti kerbau di Telukbetung 
                      (Sumatra) pada tahun 1884, kemudian di Buleleng (Bali), 
                      Palembang dan Lampung (1885) serta di&nbsp; Banten (1886). 
                      Namun selama lebih dari 100 tahun penyakit <i>anthrax</i> 
                      tidak pernah terjadi lagi di Bali, sehingga Bali dinyatakan 
                      sebagai daerah bebas <i>anthrax</i> sampai saat ini.</span></span></font></p>
                    <p class="MsoBodyTextIndent" style="text-align: justify; text-indent: .25in; font-family: Arial Narrow; margin-left: 0in; margin-right: 0in; margin-top: 0in; margin-bottom: .0001pt"><font size="-5"><span class="content"><span lang="IN">Penyakit 
                      <i>anthrax</i> merupakan penyakit <i>zoonosis </i>yang artinya 
                      berpotensi menular dari hewan kepada manusia. Zoonosis umumnya 
                      terjadi di daerah yang tergolong kurang subur dan tingkat 
                      pendidikan masyarakatnya relatif rendah.&nbsp; Di daerah 
                      seperti di atas, biasanya proses pemotongan ternak dilakukan 
                      di luar Rumah Potong Hewan (RPH) tanpa pengawasan dinas 
                      peternakan.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                      </span></span></font></p>
                    <p><font size="-5"><span class="content"><span lang="IN" style="font-size: 12.0pt; font-family: Times New Roman"><img border=0 width=300 height=191 src="img/anthrax1.jpg" v:shapes="_x0000_s1039"> 
                      </span></span></font></p>
                    <p class="MsoBodyTextIndent" style="text-indent: .25in; font-size: 11.0pt; font-family: Arial Narrow; margin-left: 0in; margin-right: 0in; margin-top: 0in; margin-bottom: .0001pt" align="justify"><font size="-5"><span class="content"><i><span lang="IN" style="font-size: 12.0pt; ">Anthrax</span></i><span lang="IN" style="font-size: 12.0pt; "> 
                      adalah suatu penyakit yang disebabkan oleh kuman <i>bacillus 
                      anthracis</i> yang bersifat mematikan. Bakteri ini berbentuk 
                      batang, berukuran 1-1.5 mikron kali 3-8 mikron, bersifat 
                      <i>anaerobic</i> (tidak memerlukan udara). Apabila terjadi 
                      kontak dengan oksigen (O<sub>2</sub>), bakteri ini akan 
                      membentuk spora yang sangat tahan terhadap pengaruh lingkungan. 
                      Oleh karena itu setiap hewan yang mati dengan dugaan <i>anthrax</i> 
                      tidak boleh dilakukan autopsi. Spora <i>anthrax</i> dapat 
                      bertahan </span><span lang="IN">selama 60 tahun dalam kondisi 
                      lingkungan yang tidak subur sekalipun dan memiliki ketahanan 
                      yang hebat terhadap pengaruh panas dan bahan kimia. Di dalam 
                      tanah spora <i>anthrax</i> akan menjadi bentuk vegetatif 
                      pada kondisi yang cocok dan akan berbentuk spora bila kondisi 
                      tanahnya tidak menguntungkan. Tanah berair kurang cocok 
                      untuk spora <i>anthrax </i>tetapi tanah berkapur yang bersifat 
                      basa, merupakan lingkungan yang paling disukai oleh <i>B. 
                      anthracis</i>. Jikalau spora ini masuk ke tempat yang tidak 
                      ada udara, misalkan di dalam tubuh, dia akan kembali berubah 
                      menjadi kuman yang menyebabkan kematian.</span></span></font></p>
                    <p class="MsoNormal" style="text-align:justify;text-indent:.25in"><font size="-5"><span class="content"><span lang="IN" style="font-family: Arial Narrow">Hampir 
                      semua hewan berdarah panas dan golongan mamalia sangat peka 
                      terhadap penyakit ini. Di Indonesia, penyakit ini menjangkiti 
                      sapi, kerbau, kambing, domba, kuda dan babi. Burung pemakan 
                      <i>kadaver</i> tidak tertular, namun bertindak sebagi penyebar 
                      penyakit ke daerah lain, apabila burung tersebut membawa 
                      makanan tercemar dan terbang ke tempat lain untuk menghabiskan 
                      makanan tersebut.</span></span></font></p>
                    <p class="MsoNormal" style="text-align:justify;text-indent:.25in"><font size="-5"><span class="content"><i><span lang="IN" style="font-family: Arial Narrow">Anthrax 
                      </span></i><span lang="IN" style="font-family: Arial Narrow">akan 
                      menjadi problem dari masa ke masa apabila terjadi siklus 
                      <b>hewan-tanah-hewan</b>. Pada musim kemarau yang panjang, 
                      di areal tanah tertentu terjadi perubahan dasyat pada suatu 
                      lingkungan mikro. Kondisi seperti ini, menyebabkan <i>anthrax</i> 
                      muncul dalam bentuk wabah (epidemi). Pada musim kering, 
                      di mana rumput sangat langka, sering terjadi ternak tertular 
                      lewat rerumputan yang tercabut sampai ke akarnya. Lewat 
                      akar tersebut terbawa pula spora <i>anthrax</i>.</span></span></font></p>
                    <p class="MsoNormal" style="text-align:justify;text-indent:.25in"><font size="-5"><span class="content"><span lang="IN" style="font-family: Arial Narrow">Penularan 
                      <i>anthrax</i> pada hewan umumnya terjadi <i><u>per os</u></i><u> 
                      yaitu penularan penyakit melalui mulut</u>, karena makan 
                      dan minum dari makanan dan minuman yang tercemar. Pada sapi, 
                      kerbau dan kuda umumnya <i>anthrax</i> bersifat akut. Oleh 
                      karena itu, kematian hewan-hewan secara mendadak di daerah 
                      endemik <i>anthrax</i> tidak boleh langsung diautopsi, tetapi 
                      harus diyakinkan dulu lewat pemeriksaan darah perifer (misalnya 
                      lewat daun telinga). Bila ada dugaan <i>anthrax</i>, bangkai 
                      harus segera dikubur cukup dalam. Misalkan ada ternak yang 
                      diduga mati karena <i>anthrax</i>, kemudian dilakuan autopsi 
                      pada bangkai tersebut, maka pada saat itu sebenarnya terjadi 
                      proses penyebaran spora <i>anthrax</i> secara besar-besaran. 
                      Pada bangkai yang sudah terlanjur diautopsi, dapat ditemukan 
                      limpanya membengkak dan rapuh karena terjadi peradangan 
                      hebat dan isi limpanya berubah menjadi hitam, seperti ter.</span></span></font></p>
                    <p align="justify"><font size="-5"><span class="content"><span lang="IN" style="font-family: Arial Narrow">Terkadang 
                      ditemukan darah berwarna hitam pekat yang sulit menggumpal 
                      dari lubang anus, hidung dan telinga, sesaat sebelum hewan 
                      mati. Perkembangan bakteri <i>anthrax</i> dalam sistem limfatik 
                      relatif lamban, tetapi begitu masuk ke dalam aliran darah, 
                      bakteri ini akan berkembang dengan sangat cepat yang berlangsung 
                      terus sampai terjadi kematian. Kematian umumnya disebabkan 
                      oleh pengaruh protein toksin/racun (prototoksin) yang dihasilkan 
                      oleh bakteri ini, yang dapat menimbulkan gangguan susunan 
                      syaraf pusat berupa kelumpuhan pusat respirasi.</span></span></font></p>
                    <p><span class="content">&nbsp;</span></p></td>
                  <td width="55%" valign="top"> <p class="MsoNormal" style="text-align:justify;text-indent:.25in"><span class="content"><span class="content"><span lang="IN" style="font-family: Arial Narrow"><font size="-5">Penularan 
                      <i>anthrax</i> dari hewan ke manusia umumnya terjadi karena 
                      adanya kontak dengan hewan atau hasil hewan. Penularan <i>per 
                      os</i>, pernah terjadi di Indonesia karena dilakukan pemotongan 
                      ternak di rumah, kemudian daging ternak tersebut dibuat 
                      sate tanpa pembakaran sempurna. </font></span></span></span></p>
                    <p class="MsoNormal" style="text-align:justify;text-indent:.25in"><font size="-5"><span class="content"><span class="content"><span lang="IN" style="font-family: Arial Narrow">Dari 
                      perbedaan tempat infeksi di tubuh, <i>anthrax </i>dapat 
                      dibagi menjadi tiga jenis. Yang pertama <b><i>cutaneous 
                      anthrax</i></b>, kedua <b><i>gastro intestinal anthrax</i></b> 
                      dan yang ketiga <b><i>inhaled anthrax</i></b>.</span></span></span></font></p>
                    <p class="MsoBodyTextIndent" align="justify"><font size="-5" face="Arial Narrow"><span class="content"><b><i><span lang="IN">Cutaneous 
                      anthrax</span></i></b><span lang="IN"> disebabkan oleh infeksi 
                      melalui luka di kulit. Jenis ini meliputi &gt;95% kasus 
                      yang dilaporkan di seluruh dunia, termasuk kasus di Bogor 
                      pada Januari 2002. Spora dari binatang yang terinfeksi, 
                      misalkan di tempat penjagalan masuk ke kulit korban melalui 
                      lubang luka. Dalam waktu 1-2 hari kemudian, muncul benjolan 
                      yang gatal, disusul dengan gelembung cairan dan borok hitam 
                      di sekelilingnya. Apabila cepat diobati, 99% dapat sembuh 
                      total. Tetapi seperti kasus yang terjadi di Indonesia, biasanya 
                      hal ini kurang diperhatikan sehingga infeksi lebih lanjut 
                      ke jaringan lain melalui aliran darah, yang bisa menimbulkan 
                      kondisi lebih parah dan kematian.</span></span></font></p>
                    <p class="MsoNormal" style="text-align:justify" align="center"><font size="-5"><span class="content"><img border="0" src="img/anthrax2.jpg" width="300" height="212"></span></font></p>
                    <p align="justify"><font size="-5"><span class="content"><b><i><span lang="IN" style="font-family: Arial Narrow">Gastro 
                      intestinal anthrax</span></i></b><span lang="IN" style="font-family: Arial Narrow">, 
                      yang disebabkan oleh infeksi melalui makanan/daging yang 
                      sudah tertular.&nbsp; Spora <i>B.anthracis</i> sangat stabil, 
                      sehingga lebih baik menghindari makan daging dari ternak 
                      yang mati karena <i>anthrax.</i> Contoh kasus di Purwakarta, 
                      Jabar pada awal tahun 2000, akibat mengkonsumsi daging burung 
                      unta terinfeksi yang dijual murah, dan di Bogor pada Agustus 
                      2002, karena mengkonsumsi daging sapi perah yang terinfeksi. 
                      Gejalanya adalah sakit perut&nbsp; yang mendadak, disertai 
                      mual, muntah dan mencret berat. Bila sudah parah, akan menyebabkan 
                      pendarahan di dalam perut. Kalau tidak segera diobati, resiko 
                      kematiannya mencapai 25-60%.</span></span></font></p>
                    <p class="MsoNormal" style="text-align:justify;text-indent:.25in"><font size="-5"><span class="content"><b><i><span lang="IN" style="font-family: Arial Narrow">Inhaled 
                      anthrax</span></i></b><span lang="IN" style="font-family: Arial Narrow">, 
                      disebabkan karena spora yang terhirup oleh korban. Jadi 
                      hanya mungkin disebabkan oleh ulah manusia yang menyebarkan 
                      spora tersebut dalam tindakan teroris atau perang. <i>Anthrax</i> 
                      jenis ini sangat mematikan (90% kemungkinan tewas). Ini 
                      disebabkan karena spora <i>B. anthracis</i> langsung terbawa 
                      ke dalam tubuh melalui paru-paru dan berinteraksi dengan 
                      sel <i>macrophage</i> yang menjadi sasaran pertamanya. Gejala 
                      awalnya, setelah 1-5 hari masa inkubasi sangat mirip dengan 
                      flu biasa, seperti batuk-batuk, panas dan badan lemah. Saat 
                      kondisi makin parah, seperti sulit bernafas, sudah dipastikan 
                      korban tidak tertolong lagi karena protein racun sudah menyebar 
                      dan tidak bisa dimusnahkan dengan antibiotika apapun.</span></span></font></p>
                    <p class="MsoNormal" style="text-align:justify;text-indent:.25in"><font size="-5"><span class="content"><span class="content"><span lang="IN" style="font-family: Arial Narrow">Kematian 
                      pasien karena <i>anthrax</i> umumnya tidak terjadi jika 
                      pengobatan dilakukan pada tahap awal penyakit.</span></span></span></font></p>
                    <p class="MsoNormal" style="text-align:justify;text-indent:.25in"><font size="-5"><span class="content"><span class="content"><span lang="IN" style="font-family: Arial Narrow">Pada 
                      manusia, spesimen untuk pemeriksaan laboratorik dapat diambil 
                      dari cairan vesikel (urin), jaringan tubuh, darah (sewaktu 
                      terjadi <i>septicemia</i>) dan usapan langsung (<i>direct 
                      smear</i>) dari lesi kulit. </span></span></span></font></p>
                    <p class="MsoNormal" style="text-align:justify;text-indent:.25in"><font size="-5"><span class="content"><span class="content"><span lang="IN" style="font-family: Arial Narrow">Pengobatan 
                      <i>anthrax</i> dapat dilakukan dengan antibiotika seperti 
                      penisilin dan oksitetrasiklin, bila penyakit masih dalam 
                      tahap awal.</span></span></span></font></p>
                    <p class="MsoNormal" style="text-align:justify;text-indent:.25in"><font size="-5"><span class="content"><span class="content"><span lang="IN" style="font-family: Arial Narrow">Pencegahan 
                      penyakit <i>anthrax</i> bisa dilakukan dengan tindakan karantina, 
                      berupa larangan masuknya hewan dari daerah tertular ke daerah 
                      bebas. Sebagai contoh, hewan dan bahan asal hewan dari NTB 
                      dan NTT tidak diperbolehkan masuk ke Bali. Hal ini dilakukan 
                      untuk melokalisasi penyebaran penyakit karena bakteri <i>anthrax.</i> 
                      </span></span></span></font></p>
                    <p class="MsoNormal" style="text-align:justify;text-indent:.25in"><font size="-5"><span class="content"><i><span lang="IN" style="font-family: Arial Narrow">Anthrax</span></i><span lang="IN" style="font-family: Arial Narrow"> 
                      memang sulit dimusnahkan, tetapi bukan mustahil untuk dilakukan. 
                      Niat melenyapkan <i>anthrax </i>dari Indonesia, bukan tanpa 
                      dasar, karena sudah ada beberapa negara yang telah sukses 
                      membebaskan diri&nbsp; dari wabah ini. Misalkan Ciprus, 
                      dengan usaha kerasnya selama lebih dari 50 tahun telah mendeklarasikan 
                      dirinya sebagai negara bebas <i>anthrax</i>, yang notabene 
                      <i>anthrax</i> telah tercatat di negara ini sejak tahun 
                      1910 dan menyerang ternak terutama domba dan kambing.</span></span></font></p>
                    <p class="MsoNormal" style="text-align:justify;text-indent:.25in"><font size="-5"><span class="content"><span class="content"><span lang="IN" style="font-family: Arial Narrow">Kebijakan 
                      veteriner yang ditempuh berupa pelaksanaan program vaksinasi 
                      terhadap ternak yang beresiko terkena <i>anthrax</i>, selama 
                      kurun waktu tertentu dan pengawasan lalu lintas ternak yang 
                      sangat ketat. </span></span></span></font></p>
                    <p class="MsoNormal" style="text-align:justify;text-indent:.25in"></p>
                    <font size="-5"><span class="content"> 
                    <![if !mso]>
                    <![endif]>
                    </span> 
                    <p align="justify"><span class="content"><span class="content"><span lang="IN" style="font-family: Arial Narrow">Jadi 
                      jelas bahwa bakteri anthrax memang paling kuat bertahan 
                      di alam, namun <b>bukan merupakan penyakit yang tidak bisa 
                      dilenyapkan. </b></span></span></span></p>
                    </font> <p align="justify"><font size="-5"><i><span class="content"><span lang="IN" style="font-family: Arial Narrow"><b>(</b></span><span style="mso-bidi-font-size:12.0pt;
    font-family:&quot;Arial Narrow&quot;">END/THP, dari berbagai sumber<o:p>)</o:p></span></span></i></font><i><span class="content"><span style="mso-bidi-font-size:12.0pt;
    font-family:&quot;Arial Narrow&quot;"><o:p></o:p></span></span></i></p>
                    <p><span class="content">&nbsp;</span></p>
                    <p><span class="content">&nbsp;</span></p>
                    <p><span class="content">&nbsp;</span></p>
                    <p><span class="content">&nbsp;</span></p></td>
                </tr>
              </table>
              <p>&nbsp;</p>
              <p>&nbsp;</p></td>
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
