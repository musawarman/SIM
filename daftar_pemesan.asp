<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%> 
<!--#include file="Connections/simConn.asp" -->
<%
' Deklarasi Variabel

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If


MM_abortEdit = false


MM_editQuery = ""
%>
<%

MM_flag="MM_insert"
If (CStr(Request(MM_flag)) <> "") Then
  MM_dupKeyRedirect="fail_pemesan.asp"
  MM_rsKeyConnection=MM_simConn_STRING
  MM_dupKeyUsernameValue = CStr(Request.Form("User Name"))
  MM_dupKeySQL="SELECT NamaPemesan FROM dbo.tb_Pemesan WHERE NamaPemesan='" & MM_dupKeyUsernameValue & "'"
  MM_adodbRecordset="ADODB.Recordset"
  set MM_rsKey=Server.CreateObject(MM_adodbRecordset)
  MM_rsKey.ActiveConnection=MM_rsKeyConnection
  MM_rsKey.Source=MM_dupKeySQL
  MM_rsKey.CursorType=0
  MM_rsKey.CursorLocation=2
  MM_rsKey.LockType=3
  MM_rsKey.Open
  If Not MM_rsKey.EOF Or Not MM_rsKey.BOF Then 

    MM_qsChar = "?"
    If (InStr(1,MM_dupKeyRedirect,"?") >= 1) Then MM_qsChar = "&"
    MM_dupKeyRedirect = MM_dupKeyRedirect & MM_qsChar & "requsername=" & MM_dupKeyUsernameValue
    Response.Redirect(MM_dupKeyRedirect)
  End If
  MM_rsKey.Close
End If
%>
<%
' Masukkan Data

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_simConn_STRING
  MM_editTable = "dbo.tb_Pemesan"
  MM_editRedirectUrl = "confirm.asp"
  MM_fieldsStr  = "User Name|value|Password|value|FirstName|value|LastName|value|Alamat|value|Kota|value|Kode Area|value|No Telepon|value|ContactPerson|value|Contact Title|value|Fax|value|Email|value|Alamat_Kandang|value|Nama Kandang|value|Kapasitas|value"
  MM_columnsStr = "NamaPemesan|',none,''|Password|',none,''|First_Name|',none,''|LastName|',none,''|Alamat|',none,''|Kota|',none,''|Kode_Area|',none,''|No_Telepon|',none,''|Contact_Person|',none,''|Contact_Title|',none,''|Fax|',none,''|Email|',none,''|Alamat_Kandang|',none,''|Nama_Kandang|',none,''|Kapasitas_Kandang|none,none,NULL"


  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  

  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next


  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%


Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then


  MM_tableValues = ""
  MM_dbValues = ""
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then

    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim rsPemesan
Dim rsPemesan_numRows

Set rsPemesan = Server.CreateObject("ADODB.Recordset")
rsPemesan.ActiveConnection = MM_simConn_STRING
rsPemesan.Source = "SELECT * FROM dbo.tb_Pemesan"
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
    <td height="20"><img src="Icon/FOLDER7.ICO"><font color="#006600">PENDAFTARAN</font></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td height="20"><img src="img/garis_ijo.gif" width="120" height="1"></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="2"><form action="<%=MM_editAction%>" method="POST" name="form1" onSubmit="MM_validateForm('User Name','','R','FirstName','','R','LastName','','R','Alamat','','R','Kota','','R','Kode Area','','NisNum','No Telepon','','NisNum','ContactPerson','','R','Contact Title','','R','Email','','NisEmail','Kapasitas','','RisNum','Password','','R');return document.MM_returnValue">
        <table width="540" border="1" align="center" bordercolor="#009900">
          <tr> 
            <td colspan="3"><div align="center"></div></td>
          </tr>
          <tr> 
            <td width="319"><div align="right"><font color="#006600">User ID *</font></div></td>
            <td width="8"><div align="center"><font color="#006600">:</font></div></td>
            <td width="586"><input name="User Name" type="text" id="User Name"></td>
          </tr>
          <tr> 
            <td><div align="right"><font color="#006600">Password *</font></div></td>
            <td><div align="center"><font color="#006600">:</font></div></td>
            <td><input name="Password" type="password" id="Password">
              maksimal 50 char</td>
          </tr>
          <tr> 
            <td><div align="right"><font color="#006600">First Name *</font></div></td>
            <td><div align="right"><font color="#006600">:</font></div></td>
            <td><div align="left"> 
                <input name="FirstName" type="text" id="FirstName">
              </div></td>
          </tr>
          <tr> 
            <td><div align="right"><font color="#006600">Last Name *</font></div></td>
            <td><div align="right"><font color="#006600">:</font></div></td>
            <td><div align="left"> 
                <input name="LastName" type="text" id="LastName">
              </div></td>
          </tr>
          <tr> 
            <td><div align="right"><font color="#006600">Alamat *</font></div></td>
            <td><div align="center"><font color="#006600">:</font></div></td>
            <td><input name="Alamat" type="text" id="Alamat" size="40"></td>
          </tr>
          <tr> 
            <td><div align="right"><font color="#006600">Kota *</font></div></td>
            <td><div align="center"><font color="#006600">:</font></div></td>
            <td><input name="Kota" type="text" id="Kota"></td>
          </tr>
          <tr> 
            <td><div align="right"><font color="#006600">Kode Area *</font></div></td>
            <td><div align="center"><font color="#006600">:</font></div></td>
            <td><input name="Kode Area" type="text" id="Kode Area"></td>
          </tr>
          <tr> 
            <td><div align="right"><font color="#006600">No Telepon *</font></div></td>
            <td><div align="center"><font color="#006600">:</font></div></td>
            <td><input name="No Telepon" type="text" id="No Telepon"></td>
          </tr>
          <tr> 
            <td><div align="right"><font color="#006600">Contact Person *</font></div></td>
            <td><div align="center"><font color="#006600">:</font></div></td>
            <td><input name="ContactPerson" type="text" id="ContactPerson"></td>
          </tr>
          <tr> 
            <td><div align="right"><font color="#006600">Contact Title *</font></div></td>
            <td><div align="center"><font color="#006600">:</font></div></td>
            <td><input name="Contact Title" type="text" id="Contact Title"></td>
          </tr>
          <tr> 
            <td><div align="right"><font color="#006600">Fax </font></div></td>
            <td><div align="center"><font color="#006600">:</font></div></td>
            <td><input name="Fax" type="text" id="Fax"></td>
          </tr>
          <tr> 
            <td height="26"> <div align="right"><font color="#006600">Email </font></div></td>
            <td><div align="center"><font color="#006600">:</font></div></td>
            <td><input name="Email" type="text" id="Email" size="30">
              contoh:john@yahoo.com</td>
          </tr>
          <tr> 
            <td><div align="right"><font color="#006600">Alamat Kandang</font></div></td>
            <td><div align="center"><font color="#006600">:</font></div></td>
            <td> <input name="Alamat_Kandang" type="text" id="Alamat_Kandang" size="50"></td>
          </tr>
          <tr> 
            <td><div align="right"><font color="#006600">Nama Kandang</font></div></td>
            <td><font color="#006600">:</font></td>
            <td> <input name="Nama Kandang" type="text" id="Nama Kandang"></td>
          </tr>
          <tr> 
            <td> <div align="right"><font color="#006600">Kapasitas Kandang </font></div></td>
            <td><font color="#006600">:</font></td>
            <td> <input name="Kapasitas" type="text" id="Kapasitas"></td>
          </tr>
          <tr> 
            <td><div align="right"> 
                <input type="submit" name="Submit2" value="Kirim ">
                <input type="reset" name="Reset" value="Reset">
              </div></td>
            <td><div align="center"></div></td>
            <td>(*) harus diisi </td>
          </tr>
          <tr> 
            <td colspan="3"> <div align="right"></div>
              <div align="center"></div></td>
          </tr>
        </table>
        <p> 
          <input type="hidden" name="MM_insert" value="form1">
        </p>
        <p>&nbsp; </p>
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
