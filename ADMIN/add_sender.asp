<%@LANGUAGE="VBSCRIPT"%> 
<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "../index.asp"
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
<!--#include file="../Connections/simConn.asp" -->
<%
Dim rsSupplier
Dim rsSupplier_numRows

Set rsSupplier = Server.CreateObject("ADODB.Recordset")
rsSupplier.ActiveConnection = MM_simConn_STRING
rsSupplier.Source = "SELECT SupplierID FROM dbo.Supplier"
rsSupplier.CursorType = 0
rsSupplier.CursorLocation = 2
rsSupplier.LockType = 1
rsSupplier.Open()

rsSupplier_numRows = 0
%>
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Request.QueryString
MM_valUsername=CStr(Request.Form("SupplierID"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization=""
  MM_redirectLoginSuccess="add_pengirim.asp"
  MM_redirectLoginFailed="fail.asp"
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_simConn_STRING
  MM_rsUser.Source = "SELECT SupplierID, SupplierID"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM dbo.Supplier WHERE SupplierID='" & Replace(MM_valUsername,"'","''") &"' AND SupplierID='" & Replace(Request.Form("SupplierID"),"'","''") & "'"
  MM_rsUser.CursorType = 0
  MM_rsUser.CursorLocation = 2
  MM_rsUser.LockType = 3
  MM_rsUser.Open
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
	Session("Updatesupplier") = Session("MM_Username")
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<html>
<head>
<title>--Administration :: Sierad Produce</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../style.css" rel="stylesheet" type="text/css">
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
//-->
</script>
</head>

<body background="../img/bg.gif" onLoad="MM_preloadImages('img/adduser_on.gif','img/addnews_on.gif','img/listuser_on.gif','img/listnews_on.gif','img/addcompany_on.gif','img/kirimms_on.gif','img/addsender_on.gif','img/addproduct_on.gif','img/addshippers_on.gif','img/addsupplier_on.gif','img/listcompany_on.gif','img/listdivision_on.gif','img/listsender_on.gif','img/listproduct_on.gif','img/listshippers_on.gif','img/listsupplier_on.gif','img/listkirimms_on.gif','img/adddivisi_on.gif')">
<table width="100%" height="95" border="0" align="center">
  <tr>
    <td bgcolor="#FFFFFF"><div align="left"><img src="img/bann.jpg" width="500" height="100"></div></td>
  </tr>
</table>
<table width="100%" border="0">
  <tr>
    <td><div align="right">tanggal : <%= date		   %></div></td>
  </tr>
</table>
<table width="100%" border="0">
  <tr>
    <td><img src="../img/garis1.gif" width="1105" height="2"></td>
  </tr>
</table>
<table width="150" border="0" bgcolor="#CCCCCC">
  <tr> 
    <td bgcolor="#009966"> <div align="center"><font color="#FFFFFF">:: MENU ::</font></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="add_users.asp" onMouseOver="MM_swapImage('Image1','','img/adduser_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/adduser.jpg" name="Image1" width="150" height="30" border="0" id="Image1"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="add_news.asp" onMouseOver="MM_swapImage('Image3','','img/addnews_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/addnews.jpg" name="Image3" width="150" height="30" border="0" id="Image3"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="add_company.asp" onMouseOver="MM_swapImage('Image11','','img/addcompany_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/addcompany.jpg" name="Image11" width="150" height="30" border="0" id="Image11"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="add_divisi.asp" onMouseOver="MM_swapImage('Image12','','img/adddivisi_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/adddivisi.jpg" name="Image12" width="150" height="30" border="0" id="Image12"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="#" onMouseOver="MM_swapImage('Image13','','img/kirimms_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/kirimms.jpg" name="Image13" width="150" height="30" border="0" id="Image13"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="add_sender.asp" onMouseOver="MM_swapImage('Image14','','img/addsender_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/addsender.jpg" name="Image14" width="150" height="30" border="0" id="Image14"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="#" onMouseOver="MM_swapImage('Image15','','img/addproduct_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/addproduct.jpg" name="Image15" width="150" height="30" border="0" id="Image15"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="#" onMouseOver="MM_swapImage('Image16','','img/addshippers_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/addshippers.jpg" name="Image16" width="150" height="30" border="0" id="Image16"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="add_supplier.asp" onMouseOver="MM_swapImage('Image17','','img/addsupplier_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/addsupplier.jpg" name="Image17" width="150" height="30" border="0" id="Image17"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><img src="../img/garis1.gif" width="187" height="2"></div></td>
  </tr>
</table>
<table width="150" border="0" bgcolor="#CCCCCC">
  <tr> 
    <td><div align="center"><a href="br_users.asp" onMouseOver="MM_swapImage('Image6','','img/listuser_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/list_user.jpg" name="Image6" width="150" height="30" border="0" id="Image6"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="br_news.asp" onMouseOver="MM_swapImage('Image7','','img/listnews_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/list_news.jpg" name="Image7" width="150" height="30" border="0" id="Image7"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="br_company.asp" onMouseOver="MM_swapImage('Image18','','img/listcompany_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/list_company.jpg" name="Image18" width="150" height="30" border="0" id="Image18"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="br_divisi.asp" onMouseOver="MM_swapImage('Image19','','img/listdivision_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/list_division.jpg" name="Image19" width="150" height="30" border="0" id="Image19"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="br_kirimms.asp" onMouseOver="MM_swapImage('Image20','','img/listkirimms_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/list_kirimms.jpg" name="Image20" width="150" height="30" border="0" id="Image20"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="br_sender.asp" onMouseOver="MM_swapImage('Image21','','img/listsender_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/list_sender.jpg" name="Image21" width="150" height="30" border="0" id="Image21"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="br_product.asp" onMouseOver="MM_swapImage('Image22','','img/listproduct_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/list_product.jpg" name="Image22" width="150" height="30" border="0" id="Image22"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="#" onMouseOver="MM_swapImage('Image23','','img/listshippers_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/list_shippers.jpg" name="Image23" width="150" height="30" border="0" id="Image23"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="br_supplier.asp" onMouseOver="MM_swapImage('Image24','','img/listsupplier_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="img/list_supplier.jpg" name="Image24" width="150" height="30" border="0" id="Image24"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><img src="../img/garis1.gif" width="187" height="2"></div></td>
  </tr>
</table>
<table width="150" border="0" bgcolor="#CCCCCC">
  <tr> 
    <td><div align="center">)) <a href="<%= MM_Logout %>">logout</a> 
        ((</div></td>
  </tr>
  <tr> 
    <td><div align="center"><img src="../img/garis1.gif" width="187" height="2"></div></td>
  </tr>
</table>
<div id="Layer1" style="position:absolute; left:204px; top:148px; width:916px; height:371px; z-index:1"> 
  <table width="915" height="18" border="0" bgcolor="#CCCCCC">
    <tr> 
      <td height="14" bgcolor="#009966"> <div align="left"><font color="#FFFFFF">.: 
          ADD SENDER :.</font></div></td>
    </tr>
  </table>
  <form ACTION="<%=MM_LoginAction%>" METHOD="POST" name="form1">
    
    <table width="53%" border="0" align="center">
      <tr> 
        <td colspan="3"> <div align="center"><img src="../img/garis1.gif" width="500" height="2"></div></td>
      </tr>
      <tr> 
        <td colspan="3"><div align="center"><font color="#FF0000">Silahkan pilih 
            dari Supplier mana Produk akan dikirim !</font></div></td>
      </tr>
      <tr> 
        <td width="48%"> <div align="right">Supplier ID *</div></td>
        <td width="0%"><div align="center">:</div></td>
        <td width="52%"> <select name="SupplierID" size="1" id="SupplierID">
            <%
While (NOT rsSupplier.EOF)
%>
            <option value="<%=(rsSupplier.Fields.Item("SupplierID").Value)%>"><%=(rsSupplier.Fields.Item("SupplierID").Value)%></option>
            <%
  rsSupplier.MoveNext()
Wend
If (rsSupplier.CursorType > 0) Then
  rsSupplier.MoveFirst
Else
  rsSupplier.Requery
End If
%>
          </select></td>
      </tr>
      <tr bordercolor="1"> 
        <td> <div align="right">
            <input type="submit" name="Submit2" value="OK">
            <input name="hiddenField" type="hidden" value="<%= Session("updateuser") %>">
          </div></td>
        <td>&nbsp;</td>
        <td> <div align="left"> </div></td>
      </tr>
      <tr> 
        <td colspan="3"> <div align="center"><img src="../img/garis1.gif" width="500" height="2"></div></td>
      </tr>
    </table>
  </form>
  <table width="53%" border="0">
    <tr> 
      <td><div align="right"> 
          <INPUT name="button" type=button class=btn onclick=history.back() value=Back>
        </div></td>
    </tr>
  </table>
  <p>&nbsp;</p>
</div>
</body>
</html>
<%
rsSupplier.Close()
Set rsSupplier = Nothing
%>