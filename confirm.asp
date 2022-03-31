<%@LANGUAGE="VBSCRIPT"%>

<!--#include file="Connections/simConn.asp" -->
<%
Dim rsPemesan
Dim rsPemesan_numRows

Set rsPemesan = Server.CreateObject("ADODB.Recordset")
rsPemesan.ActiveConnection = MM_simConn_STRING
rsPemesan.Source = "SELECT * FROM dbo.tb_Pemesan ORDER BY IDPemesan DESC"
rsPemesan.CursorType = 0
rsPemesan.CursorLocation = 2
rsPemesan.LockType = 1
rsPemesan.Open()

rsPemesan_numRows = 0
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
<script language="JavaScript">
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

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' silahkan lihat format email (ex: john@yahoo.com).\n';
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' harus berupa angka.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' tidak boleh kosong.\n'; }
  } if (errors) alert('Silahkan isikan data dengan lengkap:\n'+errors);
  document.MM_returnValue = (errors == '');
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
        <param name="quality" value="high">
        <embed src="Animasi/logo.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="178" height="93"></embed> 
      </object></td>
    <td width="150" background="img/bg_top2.jpg"><img src="img/bg_top2.jpg" width="150" height="93"></td>
    <td width="211" background="img/bg_top.jpg"><img src="img/bg_top.jpg" width="211" height="93"></td>
    <td background="img/bg_top3.jpg"><img src="img/bg_top3.jpg" width="1" height="93"></td>
    <td width="469"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="469" height="93">
        <param name="movie" value="Animasi/tagline.swf">
        <param name="quality" value="high">
        <embed src="Animasi/tagline.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="469" height="93"></embed> 
      </object></td>
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
          <td><a href="corporate/index.html" onMouseOut="OutMenu('PopMenu2');MM_swapImgRestore()" onMouseOver="PopMenu('PopMenu2',event);MM_swapImage('Image8','','img/menu_on_1.jpg',1)"><img src="img/menu_1.jpg" alt="Tentang Perusahaan" name="Image8" width="200" height="31" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="business/index.html" onMouseOut="OutMenu('PopMenu1');MM_swapImgRestore()" onMouseOver="PopMenu('PopMenu1',event);MM_swapImage('Image9','','img/menu_on_2.jpg',1)"><img src="img/menu_2.jpg" alt="Struktur Bisnis" name="Image9" width="200" height="31" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="products/index.html" onMouseOut="OutMenu('PopMenu3');MM_swapImgRestore()" onMouseOver="PopMenu('PopMenu3',event);MM_swapImage('Image10','','img/menu_on_3.jpg',1)"><img src="img/menu_3.jpg" alt="Produk" name="Image10" width="200" height="30" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="news/index.html" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image11','','img/menu_on_4.jpg',1)"><img src="img/menu_4.jpg" alt="Berita" name="Image11" width="200" height="31" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="careers/index.html" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image12','','img/menu_on_5.jpg',1)"><img src="img/menu_5.jpg" alt="Karir" name="Image12" width="200" height="30" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="report/index.html" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image13','','img/menu_on_6.jpg',1)"><img src="img/menu_6.jpg" alt="Laporan Tahunan" name="Image13" width="200" height="31" border="0"></a></td>
        </tr>
        <tr> 
          <td><a href="contact/index.html" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image14','','img/menu_on_7.jpg',1)"><img src="img/menu_7.jpg" alt="Alamat" name="Image14" width="200" height="29" border="0"></a></td>
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
<table width="928" border="0" align="center">
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="2"><form action="confirm_pemesan.asp" METHOD="POST" name="form1" onSubmit="MM_validateForm('User Name','','R','Password','','R','Alamat','','R','Kota','','R','Kode Area','','NisNum','No Telepon','','NisNum','ContactPerson','','R','Contact Title','','R','Email','','NisEmail');return document.MM_returnValue">
        <table width="500" border="0" align="center" bordercolor="#009900">
          <tr> 
            <td colspan="3"><div align="left">DATA YANG ANDA MASUKAN :</div></td>
          </tr>
          <tr> 
            <td width="319"><div align="right">User ID *</div></td>
            <td width="8"><div align="center">:</div></td>
            <td width="586"><input name="User Name" type="text" id="User Name" value="<%=(rsPemesan.Fields.Item("NamaPemesan").Value)%>"></td>
          </tr>
          <tr> 
            <td><div align="right">Password *</div></td>
            <td><div align="center">:</div></td>
            <td><input name="Password" type="text" id="Password" value="<%=(rsPemesan.Fields.Item("Password").Value)%>">
              maksimal 50 char</td>
          </tr>
          <tr> 
            <td><div align="right">First Name *</div></td>
            <td><div align="right">:</div></td>
            <td><div align="left"> 
                <input name="FirstName" type="text" id="FirstName" value="<%=(rsPemesan.Fields.Item("First_Name").Value)%>">
              </div></td>
          </tr>
          <tr> 
            <td><div align="right">Last Name *</div></td>
            <td><div align="right">:</div></td>
            <td><div align="left"> 
                <input name="LastName" type="text" id="LastName" value="<%=(rsPemesan.Fields.Item("LastName").Value)%>">
              </div></td>
          </tr>
          <tr> 
            <td><div align="right">Alamat *</div></td>
            <td><div align="center">:</div></td>
            <td><input name="Alamat" type="text" id="Alamat" value="<%=(rsPemesan.Fields.Item("Alamat").Value)%>" size="40"></td>
          </tr>
          <tr> 
            <td><div align="right">Kota *</div></td>
            <td><div align="center">:</div></td>
            <td><input name="Kota" type="text" id="Kota" value="<%=(rsPemesan.Fields.Item("Kota").Value)%>"></td>
          </tr>
          <tr> 
            <td><div align="right">Kode Area *</div></td>
            <td><div align="center">:</div></td>
            <td><input name="Kode Area" type="text" id="Kode Area" value="<%=(rsPemesan.Fields.Item("Kode_Area").Value)%>"></td>
          </tr>
          <tr> 
            <td><div align="right">No Telepon *</div></td>
            <td><div align="center">:</div></td>
            <td><input name="No Telepon" type="text" id="No Telepon" value="<%=(rsPemesan.Fields.Item("No_Telepon").Value)%>"></td>
          </tr>
          <tr> 
            <td><div align="right">Contact Person *</div></td>
            <td><div align="center">:</div></td>
            <td><input name="ContactPerson" type="text" id="ContactPerson" value="<%=(rsPemesan.Fields.Item("Contact_Person").Value)%>"></td>
          </tr>
          <tr> 
            <td><div align="right">Contact Title *</div></td>
            <td><div align="center">:</div></td>
            <td><input name="Contact Title" type="text" id="Contact Title" value="<%=(rsPemesan.Fields.Item("Contact_Title").Value)%>"></td>
          </tr>
          <tr> 
            <td><div align="right">Fax *</div></td>
            <td><div align="center">:</div></td>
            <td><input name="Fax" type="text" id="Fax" value="<%=(rsPemesan.Fields.Item("Fax").Value)%>"> 
            </td>
          </tr>
          <tr> 
            <td><div align="right">Email *</div></td>
            <td><div align="center">:</div></td>
            <td><input name="Email" type="text" id="Email" value="<%=(rsPemesan.Fields.Item("Email").Value)%>">
              contoh:john@yahoo.com</td>
          </tr>
          <tr> 
            <td><div align="right"></div></td>
            <td><div align="center"></div></td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td> <div align="center"> 
                <input type="submit" name="Submit2" value="OK">
              </div></td>
            <td><div align="center"></div></td>
            <td>(*) harus diisi </td>
          </tr>
          <tr> 
            <td> <div align="right"></div></td>
            <td><div align="center"></div></td>
            <td>&nbsp;</td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
<table width="928" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="1"><img src="img/line.gif" width="1" height="25"></td>
    <td background="img/bg_line_bot.gif"><div align="right"><img src="img/spacer.gif" width="301" height="25"><font color="#336666"></font></div></td>
    <td width="1"><img src="img/line.gif" width="1" height="25"></td>
  </tr>
</table>
</body>
</html>
<%
rsPemesan.Close()
Set rsPemesan = Nothing
%>
