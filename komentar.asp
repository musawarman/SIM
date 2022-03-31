<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%> 
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
' *** Edit Operations: declare variables

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

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_simConn_STRING
  MM_editTable = "dbo.Message"
  MM_editRedirectUrl = "daftar_guestbook.asp"
  MM_fieldsStr  = "Nama|value|Email|value|Topik|value|Pesan|value"
  MM_columnsStr = "Name|',none,''|Email|',none,''|Topik|',none,''|Message|',none,''"

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
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
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
      ElseIf (MM_delim = "'") Then  ' escape quotes
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
    ' execute the insert
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
Dim rsPesan
Dim rsPesan_numRows

Set rsPesan = Server.CreateObject("ADODB.Recordset")
rsPesan.ActiveConnection = MM_simConn_STRING
rsPesan.Source = "SELECT * FROM dbo.Message"
rsPesan.CursorType = 0
rsPesan.CursorLocation = 2
rsPesan.LockType = 1
rsPesan.Open()

rsPesan_numRows = 0
%>
<html>
<head>
<title>Komentar Anda :: Sierad</title>
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

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' format email (misal : john@yahoo.com).\n';
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' tidak boleh kosong.\n'; }
  } if (errors) alert('Silahkan isikan data dengan lengkap :\n'+errors);
  document.MM_returnValue = (errors == '');
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
    <td><p><img src="Icon/reg.gif" width="29" height="25"><font color="#006600">Komentar 
        Anda </font><font color="#006600"> | <a href="daftar_guestbook.asp">Lihat 
        Daftar</a> |</font></p>
      <form action="<%=MM_editAction%>" method="POST" name="form1" onSubmit="MM_validateForm('Nama','','R','Email','','NisEmail','Topik','','R','Pesan','','R');return document.MM_returnValue">
        <table width="60%" border="0">
          <tr> 
            <td colspan="2"><div align="center"><img src="img/garis1.gif" width="500" height="2"></div></td>
          </tr>
          <tr> 
            <td width="21%" height="25"> <div align="right">Nama :</div></td>
            <td width="79%"><input name="Nama" type="text" id="Nama"></td>
          </tr>
          <tr> 
            <td><div align="right">Email :</div></td>
            <td><input name="Email" type="text" id="Email"></td>
          </tr>
          <tr> 
            <td><div align="right">Topik :</div></td>
            <td><input name="Topik" type="text" id="Topik"> </td>
          </tr>
          <tr> 
            <td><div align="right">Pesan</div></td>
            <td><textarea name="Pesan" cols="40" rows="10" id="Pesan"></textarea></td>
          </tr>
          <tr> 
            <td><div align="right"> 
                <input type="submit" name="Submit" value="Kirim !">
              </div></td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td height="20" colspan="2"> <div align="center"><img src="img/garis1.gif" width="500" height="2"></div></td>
          </tr>
        </table>
        <input type="hidden" name="MM_insert" value="form1">
      </form>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p><img src="img/bannerrg.jpg" width="790" height="20"></p></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
<%
rsPesan.Close()
Set rsPesan = Nothing
%>
