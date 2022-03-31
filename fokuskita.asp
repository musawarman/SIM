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
    <td height="1060"> 
      <div align="center"> 
        <table width="100%" border="2" bordercolor="#009966">
          <tr bordercolor="1"> 
            <td width="54%" height="16"><table width="500" border="0" align="center">
                <tr> 
                  <td height="14" bgcolor="#006600"> <div align="center"><font color="#FF9900">FOKUS 
                      KITA </font></div></td>
                </tr>
              </table>
              <table width="600" border="0" align="center">
                <tr>
                  <td><font color="#FF6600" size="+2">Bedah Kasus PT QSAR </font>
<p><br>
                      PT QSAR (Qurnia Subur Alam Raya), serasa masih hangat di 
                      ingatan kita. Berbagai media massa mengulas dan menghadirkannya 
                      sebagai berita hangat, karena menyangkut dana milyaran rupiah. 
                      PT QSAR yang bergerak di bidang agribisnis ini beroperasi 
                      di daerah Kadu Dampit Sukabumi. Dalam waktu yang relatif 
                      singkat, sekitar empat tahun aset perusahaan berlipat menjadi 
                      kurang lebih 600 milyar rupiah. Hal ini tidak wajar karena 
                      menurut Bob Sadino, dalam kondisi normal ukuran perusahaan 
                      agribisnis sebesar itu, baru dapat dicapai paling tidak 
                      dalam waktu 10 tahun. Prakteknya perusahaan tersebut justru 
                      ambruk. Mengapa bisa terjadi ?</p>
                    <p><br>
                      Dari berbagai media, dapat disimpulkan bahwa ambruknya PT 
                      QSAR terjadi karena hal-hal berikut ini :</p>
                    <p><br>
                      1. Tidak ada&#8221;goodwill&#8221; bahkan ada indikasi pihak 
                      pengelola membohongi investor</p>
                    <p><br>
                      Dalam proposal untuk para investor, harga jual yang ditawarkan 
                      tidak akurat dengan kondisi pasar, misalkan tomat ditawarkan 
                      Rp 5.000,- per kg (pukul rata), padahal di tingkat pengecer 
                      - kita bisa membeli dengan harga yang lebih rendah. Lagi 
                      pula tidak semua produk pertanian berkualitas seragam. Pemasaran 
                      produk yang katanya diekspor ke Malaysia dan Singapura, 
                      prakteknya hanya dijual di pasar sayur Bogor dan Cibitung. 
                      Fakta yang paling menggelikan, jalan masuk ke tempat panen 
                      pun tidak bisa dilalui oleh kontainer yang biasa digunakan 
                      untuk ekspor.</p>
                    <p><br>
                      2. Manajemen pengelola perusahaan kedodoran</p>
                    <p><br>
                      Buruknya sistem administrasi menyebabkan ada beberapa investor 
                      yang dibayar sampai dua kali dan bahkan ada yang kelebihan 
                      angka nol, dalam arti investor tersebut dibayar 10 kali 
                      lipat dari yang seharusnya. Pihak manajemen juga tidak bisa 
                      mengelola dana yang membanjir dan mengakibatkan over value. 
                      Investasi yang berlebihan justru menimbulkan kehancuran 
                      karena uang yang masuk harus dikembalikan dengan hasilnya 
                      walaupun program belum sempat dijalankan.</p>
                    <p><br>
                      3. Penguasaan teknis akan komoditi yang ditanam tidak memadai</p>
                    <p><br>
                      Pematokan harga jual yang fantastis terlepas dari niat untuk 
                      membohongi investor. Fakta yang ada, hampir semua produk 
                      pertanian tidak ada yang harga jualnya sama (flat) sepanjang 
                      tahun. Pemasaran produknya tidak diperhitungkan dengan matang 
                      dan tidak disesuaikan dengan daya serap pasar. Kendala utama 
                      agribisnis adalah produknya begitu fragile dalam arti mudah 
                      rusak karena penanganan yang tidak benar; tanamannya sangat 
                      terpengaruh cuaca, iklim, hama dan harga yang hampir selalu 
                      dikuasai oleh tengkulak.</p>
                    <p><br>
                      4. Peran pejabat pemerintah cenderung ceroboh</p>
                    <p></p>
                    <p><br>
                      Kesalahan dimulai dari pemberian perijinan yang kabarnya 
                      baru dibuat pada awal tahun 2001 (kurang lebih 2 tahun sesudah 
                      perusahaan berdiri). Seharusnya instansi terkait dapat mempelajari 
                      dulu teknik kelayakan usahanya pada saat pengajuan ijin. 
                      Selain itu banyaknya pejabat yang datang ke lokasi, foto 
                      bersama pemilik dan pengelola yang tanpa disadari dijadikan 
                      bahan promosi gratis di tabloid internal maupun media massa 
                      umum. Jadi jangan salahkan bila banyak masyarakat yang tergiur 
                      untuk menjadi investor melihat banyak pejabat yang pernah 
                      berkunjung ke lokasi.</p>
                    <p></p>
                    <p></p>
                    <p><br>
                      END/THP (dari berbagai sumber)</p>
                    <p></p></td>
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
