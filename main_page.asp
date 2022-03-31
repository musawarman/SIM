<%@LANGUAGE="VBSCRIPT" %> 
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
<!--#include file="Connections/simConn.asp" -->
<%
Dim rs_PrdName
Dim rs_PrdName_numRows

Set rs_PrdName = Server.CreateObject("ADODB.Recordset")
rs_PrdName.ActiveConnection = MM_simConn_STRING
rs_PrdName.Source = "SELECT * FROM dbo.tb_PrdName"
rs_PrdName.CursorType = 0
rs_PrdName.CursorLocation = 2
rs_PrdName.LockType = 1
rs_PrdName.Open()

rs_PrdName_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rs_PrdName_numRows = rs_PrdName_numRows + Repeat1__numRows
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
<title>:: Sierad ::</title>
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

<body background="img/bg.gif" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('img/depan_on.gif','img/ID_on.gif','img/laporan_on.gif')">
<table width="800" border="1" align="center" bordercolor="#006600" bgcolor="#FFFFFF">
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
    <td height="1165"> 
      <div align="center"> 
        <table width="800" border="0">
          <tr> 
            <td height="14" bgcolor="#006600"> <div align="center"><font color="#FF9900" size="-1">:: 
                DOC Yang Ditawarkan ::</font></div></td>
          </tr>
        </table>
        <p><font color="#006600">Silahkan lakukan pemesanan dengan mengklik order.</font></p>
        <table width="600" border="1">
          <tr bgcolor="#663366"> 
            <td> <div align="center"><font color="#CCCCCC">DOC </font></div></td>
            <td> <div align="center"><font color="#CCCCCC">Deskripsi</font></div></td>
            <td> <div align="center"><font color="#CCCCCC">Harga</font></div></td>
            <td> <div align="center"><font color="#CCCCCC">Pesan</font></div></td>
          </tr>
          <% 
While ((Repeat1__numRows <> 0) AND (NOT rs_PrdName.EOF)) 
%>
          <tr bgcolor="#FFFFFF"> 
            <td> 
              <div align="center"><font color="#006600"><%=(rs_PrdName.Fields.Item("Nama_Produk").Value)%></font></div></td>
            <td> 
              <div align="center"><font color="#006600"><%=(rs_PrdName.Fields.Item("Deskripsi").Value)%></font></div></td>
            <td> 
              <div align="center"><font color="#006600"><%=(rs_PrdName.Fields.Item("Price").Value)%></font></div></td>
            <td> 
              <div align="center"><font color="#006600"><A HREF="order_doc.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ID_PrdName=" & rs_PrdName.Fields.Item("ID_PrdName").Value %>"><img src="img/ORDER.GIF" width="65" height="20"></A></font></div></td>
          </tr>
          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_PrdName.MoveNext()
Wend
%>
        </table>
        <p>&nbsp;</p>
        <table width="800" border="0">
          <tr> 
            <td height="31" bgcolor="#006600"> 
              <div align="center"><font color="#FF9900" size="2">WARTA 
                SIERAD</font></div></td>
          </tr>
        </table>
        <table width="800" border="1" cellpadding="0" cellspacing="2" bgcolor="#FFFFFF">
          <tr> 
            <td><img src="img/sierkecil.gif" width="77" height="29" border="0"> 
            </td>
            <td><span class="judul"><font color="#006600" size="-1">Fokus Kita : Bedah Kasus PT&nbsp;QSAR</font></span></td>
          </tr>
          <tr> 
            <td bgcolor="#140101"> <div align="right"> </div></td>
            <td></td>
          </tr>
          <tr> 
            <td> <div align="right"> 
                <p></p>
              </div>
              <p></p></td>
            <td><p><span class="content">PT QSAR (Qurnia Subur Alam Raya), serasa 
                masih hangat di ingatan kita. Berbagai media massa mengulas dan 
                menghadirkannya sebagai berita hangat, </span>... </p>
              <div align="right"> 
                <p><a href="fokuskita.asp">detail......<br>
                  </a></p>
              </div></td>
          </tr>
          <tr> 
            <td bgcolor="#080707"></td>
            <td></td>
          </tr>
          <tr> 
            <td><img src="img/sierkecil.gif" width="77" height="29" border="0"></td>
            <td><span class="judul"><font color="#006600" size="-1">Persepsi<a href="berita/Persepsi.htm"> 
              </a></font></span></td>
          </tr>
          <tr> 
            <td bgcolor="#080707"></td>
            <td></td>
          </tr>
          <tr> 
            <td> </td>
            <td><span class="content">Alkisah tersebutlah dua orang peri, satu 
              yang senior dan satu yang yunior. Kedua peri ini bertugas mengunjungi 
              desa-desa di bumi. Suatu hari mereka mampir di suatu desa dan menginap 
              di rumah seorang kaya yang pelit. Karena ketamakan dan kepelitannya</span>... 
              <div align="right"> 
                <p><a href="persepsi.asp">detail......<br>
                  </a></p>
              </div></td>
          </tr>
          <tr> 
            <td><img src="img/sierkecil.gif" width="77" height="29" border="0"></td>
            <td><span class="judul"><font color="#006600" size="-1">Senyum Dulu Donk<a href="berita/senyumduludong.htm"> 
              </a></font></span></td>
          </tr>
          <tr> 
            <td bgcolor="#080707"></td>
            <td></td>
          </tr>
          <tr> 
            <td> </td>
            <td><span class="content">SANG RAJA<br>
              </span> <p></p>
              <p><span class="content">Di sebuah hutan terdapat raja hutan (singa) 
                yang merasa dirinya paling hebat.<br>
                </span></p>
              <p><span class="content">Untuk melegalisasikan kehebatannya, singa 
                tersebut berencana untuk bertanya kepada semua penghuni hutan.<br>
                </span></p>
              <p><span class="content">Proses dimulai</span>...</p>
              <div align="right"> 
                <p><a href="senyum.asp">detail......<br>
                  </a></p>
              </div></td>
          </tr>
          <tr> 
            <td><img src="img/sierkecil.gif" width="77" height="29" border="0"></td>
            <td><span class="judul"><font color="#006600" size="-1">Hati-Hati... Anthrax<a href="berita/Anthrax.htm"> 
              </a></font></span></td>
          </tr>
          <tr> 
            <td bgcolor="#080707"></td>
            <td></td>
          </tr>
          <tr> 
            <td> </td>
            <td><span class="content">Nama Anthrax kembali akrab di telinga kita, 
              akibat pemberitaan yang gencar oleh media massa. Tidak hanya karena 
              kasus bom biologis di Amerika, namun juga karena kemunculannya di 
              berbagai tempat di Indonesia, seperti</span> ... 
              <div align="right"> 
                <p><a href="anthrax.asp">detail......<br>
                  </a></p>
              </div></td>
          </tr>
          <tr> 
            <td><img src="img/sierkecil.gif" width="77" height="29" border="0"></td>
            <td><span class="judul"><font color="#006600" size="-1">Pemborosan</font></span></td>
          </tr>
          <tr> 
            <td bgcolor="#080707"></td>
            <td></td>
          </tr>
          <tr> 
            <td></td>
            <td><span class="content">Kata orang waktu adalah uang. Pemborosan 
              waktu berarti pemborosan uang pula. Tetapi ternyata tidak mudah 
              untuk mengukur kerugian yang disebabkan oleh pemborosan waktu. Lebih 
              buruk lagi, </span>... 
              <div align="right"> 
                <p><a href="boros.asp">detail......<br>
                  </a></p>
              </div></td>
          </tr>
        </table>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
        <p align="left">&nbsp;</p>
        <p align="left">&nbsp;</p>
        <p align="left"><img src="img/bannerrg.jpg" width="790" height="20"></p>
        </div></td>
  </tr>
</table>
<table width="800" border="1" align="center" background="business/img/bg.gif">
  <tr> 
    <td height="11">
<div align="center"><img src="BREEDING/img/BAR_ELEG.GIF" width="790" height="7"></div></td>
  </tr>
</table>
</body>
</html>
<%
rs_PrdName.Close()
Set rs_PrdName = Nothing
%>
