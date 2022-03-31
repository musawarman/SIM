<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../../Connections/simConn.asp" -->
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
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_simConn_STRING
  MM_editTable = "dbo.ProductsKirimHD"
  MM_editColumn = "OrderID"
  MM_recordId = "'" + Request.Form("MM_recordId") + "'"
  MM_editRedirectUrl = "list_order.asp"
  MM_fieldsStr  = "OrderID|value|Tujuan|value"
  MM_columnsStr = "OrderID|',none,''|SupplierID|',none,''"

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
<%
Dim rsOrder
Dim rsOrder_numRows

Set rsOrder = Server.CreateObject("ADODB.Recordset")
rsOrder.ActiveConnection = MM_simConn_STRING
rsOrder.Source = "SELECT * FROM dbo.KirimMS ORDER BY UpdateDate DESC"
rsOrder.CursorType = 0
rsOrder.CursorLocation = 2
rsOrder.LockType = 1
rsOrder.Open()

rsOrder_numRows = 0
%>
<%
Dim rsProdHD
Dim rsProdHD_numRows

Set rsProdHD = Server.CreateObject("ADODB.Recordset")
rsProdHD.ActiveConnection = MM_simConn_STRING
rsProdHD.Source = "SELECT OrderID, SupplierID FROM dbo.ProductsKirimHD"
rsProdHD.CursorType = 0
rsProdHD.CursorLocation = 2
rsProdHD.LockType = 1
rsProdHD.Open()

rsProdHD_numRows = 0
%>
<html>
<head>
<title>:: Sierad : Activities ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

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
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
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

<body background="../../business/img/bg.gif" leftmargin="0" topmargin="0" onLoad="MM_preloadImages('../img/home_on.gif','../img/sysman_on.gif','../img/activities_on.gif','../img/report_on.gif','../img/faq_on.gif')">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="178"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="178" height="93">
        <param name="movie" value="../../img/logo.swf">
        <param name="quality" value="high">
        <embed src="../../img/logo.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="178" height="93"></embed></object></td>
    <td width="150" background="../../img/bg_top2.jpg">&nbsp;</td>
    <td width="211" background="../../img/bg_top.jpg">&nbsp;</td>
    <td background="../../img/bg_top3.jpg">&nbsp;</td>
    <td width="469"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="469" height="93">
        <param name="movie" value="../../img/tagline.swf">
        <param name="quality" value="high">
        <embed src="../../img/tagline.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="469" height="93"></embed></object></td>
  </tr>
</table>
<table width="100%" border="0">
  <tr>
    <td><img src="../img/spacer.gif" width="1125" height="10"></td>
  </tr>
</table>
<table width="100%" border="1" bordercolor="#009900" background="../../img/bg.gif">
  <tr> 
    <td width="14%"><div align="left"><a href="../index.asp" onMouseOver="MM_swapImage('Image2','','../img/home_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../img/home.gif" name="Image2" width="150" height="20" border="0" id="Image2"></a></div></td>
    <td rowspan="5"> <div align="center"><img src="../img/img_sierad_produce.jpg" width="333" height="120" align="left"><img src="../img/cooling.jpg" width="216" height="120"></div></td>
    <td width="0%" rowspan="5"><div align="left"><img src="../../img/garis.gif" width="1" height="120"></div></td>
    <td rowspan="5" background="../../business/img/bg.gif"> <div align="right"> 
        Tanggal : 
        <script name="current" src="../../GeneratedItems/current.js" language="JavaScript1.2"></script>
      </div>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p align="center"><font color="#009900"><a href="../contact.asp"><font color="#009900">Hubungi 
        Kami</font></a></font><font color="#FF0000"> </font>| <a href="../karir.asp"><font color="#009900">Karir 
        </font></a>| <a href="../link.asp"><font color="#009900">Links </font></a><img src="../../img/garis1.gif" width="390" height="1" align="top"></p></td>
  </tr>
  <tr> 
    <td><div align="left"><a href="../../ADMIN/login.asp" onMouseOver="MM_swapImage('Image3','','../img/sysman_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../img/sysman.gif" name="Image3" width="150" height="20" border="0" id="Image3"></a></div></td>
  </tr>
  <tr> 
    <td><div align="left"><a href="activities.asp" onMouseOver="MM_swapImage('Image4','','../img/activities_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../img/activities.gif" name="Image4" width="150" height="20" border="0" id="Image4"></a></div></td>
  </tr>
  <tr> 
    <td><div align="left"><a href="../Reports/reportListing.asp" onMouseOver="MM_swapImage('Image5','','../img/report_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../img/report.gif" name="Image5" width="150" height="20" border="0" id="Image5"></a></div></td>
  </tr>
  <tr> 
    <td height="24"> <div align="left"><a href="../Pencarian/Search.asp" onMouseOver="MM_swapImage('Image1','','../img/faq_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../img/faq.gif" name="Image1" width="150" height="20" border="0" id="Image1"></a></div></td>
  </tr>
</table>
<table width="100%" border="0">
  <tr>
    <td><img src="../img/spacer.gif" width="1125" height="10"></td>
  </tr>
</table>
<table width="100%" border="0">
  <tr>
    <td height="14"><p><font color="#993300">&gt;&gt; Pengiriman Barang </font></p>
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1" onSubmit="MM_validateForm('OrderID','','R','Order Date','','R','Product Name','','R','Destination Address','','R','Tujuan','','R','Phone Address','','R','Ship No','','R','Ship Name','','R','Status','','R');return document.MM_returnValue">
        <table width="53%" border="0" align="center" bordercolor="#003399">
          <tr bordercolor="1"> 
            <td colspan="3"> <div align="right"></div>
              <div align="center"></div>
              <font color="#006600"> Data Yang Anda Masukkan :</font></td>
          </tr>
          <tr> 
            <td colspan="3"> <div align="center"><img src="../../img/garis1.gif" width="500" height="2"></div></td>
          </tr>
          <tr bordercolor="1"> 
            <td> <div align="right">Order ID *</div></td>
            <td width="1%"> <div align="center">:</div></td>
            <td><input name="OrderID" id="OrderID" value="<%=(rsOrder.Fields.Item("OrderID").Value)%>" readonly="text"> 
            </td>
          </tr>
          <tr bordercolor="1"> 
            <td><div align="right">Order Date *</div></td>
            <td> <div align="center">:</div></td>
            <td><input name="Order Date" readonly="text" id="Order Date" value="<%=(rsOrder.Fields.Item("OrderDate").Value)%>"></td>
          </tr>
          <tr bordercolor="1"> 
            <td><div align="right">Product Name *</div></td>
            <td> <div align="center">:</div></td>
            <td><input name="Product Name" type="text" id="Product Name" value="<%=(rsOrder.Fields.Item("ProductName").Value)%>"></td>
          </tr>
          <tr bordercolor="1"> 
            <td><div align="right">Destination Address *</div></td>
            <td> <div align="center">:</div></td>
            <td><input name="Destination Address" type="text" id="Destination Address" value="<%=(rsOrder.Fields.Item("DestinationAddress").Value)%>" size="50"></td>
          </tr>
          <tr bordercolor="1"> 
            <td><div align="right">Tujuan Pengiriman *</div></td>
            <td>:</td>
            <td><input name="Tujuan" type="text" id="Tujuan" value="<%=(rsOrder.Fields.Item("Tujuan").Value)%>" size="40"></td>
          </tr>
          <tr bordercolor="1"> 
            <td><div align="right">Phone Address *</div></td>
            <td>:</td>
            <td><input name="Phone Address" type="text" id="Phone Address" value="<%=(rsOrder.Fields.Item("PhoneAddress").Value)%>"></td>
          </tr>
          <tr bordercolor="1"> 
            <td><div align="right">Kota *</div></td>
            <td>:</td>
            <td><input name="Kota" type="text" id="Kota" value="<%=(rsOrder.Fields.Item("City").Value)%>"></td>
          </tr>
          <tr bordercolor="1"> 
            <td><div align="right">Area Manager *</div></td>
            <td>:</td>
            <td><input name="Area Manager" type="text" id="Area Manager" value="<%=(rsOrder.Fields.Item("Area_Manager").Value)%>" size="30"></td>
          </tr>
          <tr bordercolor="1"> 
            <td><div align="right">Ship Via *</div></td>
            <td> <div align="center">:</div></td>
            <td><input name="ship Via" type="text" id="ship Via" value="<%=(rsOrder.Fields.Item("ShipVia").Value)%>"></td>
          </tr>
          <tr bordercolor="1"> 
            <td><div align="right">Ship No *</div></td>
            <td> <div align="center">:</div></td>
            <td><input name="Ship No" type="text" id="Ship No" value="<%=(rsOrder.Fields.Item("ShipNo").Value)%>"></td>
          </tr>
          <tr bordercolor="1"> 
            <td><div align="right">Ship Name *</div></td>
            <td> <div align="center">:</div></td>
            <td><input name="Ship Name" type="text" id="Ship Name" value="<%=(rsOrder.Fields.Item("ShipName").Value)%>"> 
            </td>
          </tr>
          <tr bordercolor="1"> 
            <td><div align="right">Status *</div></td>
            <td> <div align="center">:</div></td>
            <td><input name="Status" readonly="text" id="Status" value="<%=response.write("Unapproved")%>"></td>
          </tr>
          <tr bordercolor="1"> 
            <td><div align="right">ID Pengirim *</div></td>
            <td> <div align="center">:</div></td>
            <td><input name="IDPengirim" type="text" id="IDPengirim" value="<%=(rsOrder.Fields.Item("IDPengirim").Value)%>"></td>
          </tr>
          <tr bordercolor="1"> 
            <td width="28%"> <div align="right"> 
                <input type="submit" name="Submit2" value="Submit">
                <input name="hiddenField" type="hidden" value="<%= Session("updateuser") %>">
              </div></td>
            <td>&nbsp;</td>
            <td width="71%"> <div align="left"> </div></td>
          </tr>
          <tr> 
            <td colspan="3"> <div align="center"><img src="../../img/garis1.gif" width="500" height="2"></div></td>
          </tr>
        </table>
        <input type="hidden" name="MM_update" value="form1">
        <input type="hidden" name="MM_recordId" value="<%= rsOrder.Fields.Item("OrderID").Value %>">
      </form>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p></td>
  </tr>
</table>
<table width="100%" border="2" bordercolor="#FF9900" background="../../img/bg.gif">
  <tr> 
    <td height="25"> <div align="center"> <font color="#009900"> &gt;&gt; <a href="../../index.asp">DEPAN</a> 
        | <a href="../../ADMIN/login.asp">ADMINISTRATOR</a> | <a href="activities.asp">AKTIVITAS</a> 
        | <a href="../Reports/reportListing.asp">LAPORAN</a> | <a href="../Pencarian/Search.asp">PENCARIAN</a></font> 
        <font color="#009900">&lt;&lt;</font></div></td>
  </tr>
  <tr> 
    <td height="21"> <div align="center">Web Master PT. Sierad Produce Tbk </div></td>
  </tr>
</table>
</body>
</html>
<%
rsOrder.Close()
Set rsOrder = Nothing
%>
<%
rsProdHD.Close()
Set rsProdHD = Nothing
%>
