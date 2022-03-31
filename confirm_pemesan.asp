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

' Gagal edit
MM_abortEdit = false


MM_editQuery = ""
%>
<%
' Update Rekord

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_simConn_STRING
  MM_editTable = "dbo.tb_Pemesan"
  MM_editColumn = "IDPemesan"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "index.asp"
  MM_fieldsStr  = "p_username|value"
  MM_columnsStr = "UserID|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
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
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
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
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
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
    <td colspan="2"><p><font color="#993300">&gt;&gt; Corfirmation</font></p>
      <table width="100%" border="2" bordercolor="#009966">
        <tr bordercolor="1"> 
          <td width="54%" height="16"><p><font color="#336666">Terima Kasih anda 
              telah terdaftar menjadi mitra kami.</font></p>
            <p><font color="#336666">ID Anda </font></p>
            <form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
              <font color="#336666"> 
              <INPUT 
                        name=p_username 
                        style="BORDER-RIGHT: #999999 1px solid; BORDER-TOP: #333333 1px solid; FONT-SIZE: 11px; BORDER-LEFT: #333333 1px solid; WIDTH: 150px; COLOR: #000000; BORDER-BOTTOM: #999999 1px solid; HEIGHT: 18px; BACKGROUND-COLOR: #ffecec" value="<%= response.write("00")%><%=(rsPemesan.Fields.Item("IDPemesan").Value)%>-<%=(rsPemesan.Fields.Item("NamaPemesan").Value)%>-<%= date %>" readonly="text">
              <input type="submit" name="Submit" value="OK">
              </font> 
              <input type="hidden" name="MM_recordId" value="<%= rsPemesan.Fields.Item("IDPemesan").Value %>">
              <input type="hidden" name="MM_update" value="form1">
            </form></td>
        </tr>
      </table>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      </td>
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
