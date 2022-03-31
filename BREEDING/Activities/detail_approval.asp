<%@LANGUAGE="VBSCRIPT"%>
 
<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "../../index.asp"
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
<!--#include file="../../Connections/simConn.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="5,4,3"
MM_authFailedURL="failed.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (false Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<%
Dim rsKirimMS
Dim rsKirimMS_numRows

Set rsKirimMS = Server.CreateObject("ADODB.Recordset")
rsKirimMS.ActiveConnection = MM_simConn_STRING
rsKirimMS.Source = "SELECT * FROM dbo.KirimMS ORDER BY UpdateDate ASC"
rsKirimMS.CursorType = 0
rsKirimMS.CursorLocation = 2
rsKirimMS.LockType = 1
rsKirimMS.Open()

rsKirimMS_numRows = 0
%>
<%
Dim rsKirimAppr__MMColParam
rsKirimAppr__MMColParam = "1"
If (Request.Form("OrderID") <> "") Then 
  rsKirimAppr__MMColParam = Request.Form("OrderID")
End If
%>
<%
Dim rsKirimAppr
Dim rsKirimAppr_numRows

Set rsKirimAppr = Server.CreateObject("ADODB.Recordset")
rsKirimAppr.ActiveConnection = MM_simConn_STRING
rsKirimAppr.Source = "SELECT * FROM dbo.KirimAppr WHERE OrderID = '" + Replace(rsKirimAppr__MMColParam, "'", "''") + "'"
rsKirimAppr.CursorType = 0
rsKirimAppr.CursorLocation = 2
rsKirimAppr.LockType = 1
rsKirimAppr.Open()

rsKirimAppr_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rsKirimMS_numRows = rsKirimMS_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsKirimMS_total
Dim rsKirimMS_first
Dim rsKirimMS_last

' set the record count
rsKirimMS_total = rsKirimMS.RecordCount

' set the number of rows displayed on this page
If (rsKirimMS_numRows < 0) Then
  rsKirimMS_numRows = rsKirimMS_total
Elseif (rsKirimMS_numRows = 0) Then
  rsKirimMS_numRows = 1
End If

' set the first and last displayed record
rsKirimMS_first = 1
rsKirimMS_last  = rsKirimMS_first + rsKirimMS_numRows - 1

' if we have the correct record count, check the other stats
If (rsKirimMS_total <> -1) Then
  If (rsKirimMS_first > rsKirimMS_total) Then
    rsKirimMS_first = rsKirimMS_total
  End If
  If (rsKirimMS_last > rsKirimMS_total) Then
    rsKirimMS_last = rsKirimMS_total
  End If
  If (rsKirimMS_numRows > rsKirimMS_total) Then
    rsKirimMS_numRows = rsKirimMS_total
  End If
End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = rsKirimMS
MM_rsCount   = rsKirimMS_total
MM_size      = rsKirimMS_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsKirimMS_first = MM_offset + 1
rsKirimMS_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsKirimMS_first > MM_rsCount) Then
    rsKirimMS_first = MM_rsCount
  End If
  If (rsKirimMS_last > MM_rsCount) Then
    rsKirimMS_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
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
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = MM_keepMove & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<%
Dim Repeat3__numRows
Dim Repeat3__index

Repeat3__numRows = 10
Repeat3__index = 0
rsSupplier_numRows = rsSupplier_numRows + Repeat3__numRows
%>

<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsKirimappr_first = MM_offset + 1
rsKirimappr_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsKirimappr_first > MM_rsCount) Then
    rsKirimappr_first = MM_rsCount
  End If
  If (rsKirimappr_last > MM_rsCount) Then
    rsKirimappr_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
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

function MM_displayStatusMsg(msgStr) { //v1.0
  status=msgStr;
  document.MM_returnValue = true;
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
    <td><img src="../img/spacer.gif" width="1117" height="10"></td>
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
      <p align="center"><font color="#009900"><a href="../contact.asp"><font color="#009900">Hubungi 
        Kami</font></a></font><font color="#FF0000"> </font>| <a href="../karir.asp"><font color="#009900">Karir 
        </font></a>| <a href="../link.asp"><font color="#009900">Links </font></a>| 
        <font color="#009900"><a href="<%= MM_Logout %>">Logout</a></font></p>
      <p align="center"><img src="../../img/garis1.gif" width="388" height="1" align="top"></p>
      </td>
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
    <td><img src="../img/spacer.gif" width="1117" height="10"></td>
  </tr>
</table>
<table width="100%" border="0">
  <tr>
    <td height="14"><p><font color="#993300">&gt;&gt; List Approval</font></p>
      <table width="60%" border="1" align="center" bordercolor="#FF9900">
        <tr bgcolor="#6666FF"> 
          <td><div align="center"><font color="#FFFFFF">No</font></div></td>
          <td> <div align="center"><font color="#FFFFFF">Order ID</font></div></td>
          <td> <div align="center"><font color="#FFFFFF">Produk ID</font></div></td>
          <td> <div align="center"><font color="#FFFFFF">ID Area</font></div></td>
          <td> <div align="center"><font color="#FFFFFF">Status</font></div></td>
        </tr>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT rsKirimMS.EOF)) 
%>
        <tr bgcolor="#CCCCCC"> 
          <td height="22"> 
            <div align="center"><font color="#006600"><%=(Repeat1__index + 1)%></font></div></td>
          <td> 
            <div align="center"> 
              <form name="form1" method="post" action="detail_approval.asp">
                <input name="OrderID" readonly="text" id="OrderID" value="<%=(rsKirimMS.Fields.Item("OrderID").Value)%>">
                <input type="submit" name="Submit" value="cek">
              </form>
              <font color="#006600"></font></div></td>
          <td> <div align="center"><font color="#006600"><%=(rsKirimMS.Fields.Item("ProductsID").Value)%></font></div></td>
          <td> 
            <div align="center"><font color="#006600"><%=(rsKirimMS.Fields.Item("Sales_Area").Value)%></font></div></td>
          <td> 
            <div align="center"><font color="#006600"><%=(rsKirimMS.Fields.Item("Status").Value)%></font></div></td>
        </tr>
        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsKirimMS.MoveNext()
Wend
%>
      </table>
      <table border="0" width="330" align="center">
        <tr> 
          <td width="42%" height="65" align="center"> <% If MM_offset <> 0 Then %>
            <A HREF="<%=MM_moveFirst%>"><img src="../../img/first.gif" width="75" height="25" border=0 align="right" onMouseOver="MM_displayStatusMsg('rz : list divisi -&gt; first page');return document.MM_returnValue"></A> 
            <% End If ' end MM_offset <> 0 %> </td>
          <td width="8%" align="center"> 
            <% If MM_offset <> 0 Then %>
            <A HREF="<%=MM_movePrev%>"><img src="../../img/previous.gif" width="75" height="25" border=0 align="right" onMouseOver="MM_displayStatusMsg('rz : list divisi -&gt; previous page');return document.MM_returnValue"></A> 
            <% End If ' end MM_offset <> 0 %>
          </td>
          <td width="8%" align="center"> 
            <% If Not MM_atTotal Then %>
            <A HREF="<%=MM_moveNext%>"><img src="../../img/next.gif" width="75" height="25" border=0 align="left" onMouseOver="MM_displayStatusMsg('rz : list divisi -&gt;next page');return document.MM_returnValue"></A> 
            <% End If ' end Not MM_atTotal %>
          </td>
          <td width="42%" align="center"> 
            <% If Not MM_atTotal Then %>
            <A HREF="<%=MM_moveLast%>"><img src="../../img/last.gif" width="75" height="25" border=0 align="left" onMouseOver="MM_displayStatusMsg('rz : list divisi -&gt; last page');return document.MM_returnValue"></A> 
            <% End If ' end Not MM_atTotal %>
          </td>
        </tr>
      </table>
      <p align="center">&nbsp;</p>
      <% If Not rsKirimAppr.EOF Or Not rsKirimAppr.BOF Then %>
      <table width="40%" border="1" align="center" bordercolor="#FF9900">
        <tr bgcolor="#6666FF"> 
          <td> <div align="center"><font color="#FFFFFF">Order ID</font></div></td>
          <td> <div align="center"><font color="#FFFFFF">Tanggal Approval</font></div></td>
          <td> <div align="center"><font color="#FFFFFF">Approved By</font></div></td>
        </tr>
        <tr bgcolor="#CCCCCC"> 
          <td height="26"> 
            <div align="center"><font color="#006600"></font></div>
            
            <input name="OrderID" readonly="text" id="OrderID" value="<%=(rsKirimAppr.Fields.Item("OrderID").Value)%>"></td>
          <td> <div align="center"><font color="#006600">
              <input name="textfield" readonly="text" value="<%=(rsKirimAppr.Fields.Item("TanggalAppr").Value)%>">
              </font></div></td>
          <td> <div align="center"> <font color="#006600">
              <input name="Contact_Name" readonly="text" id="Contact_Name" value="<%=(rsKirimAppr.Fields.Item("UserID").Value)%>">
              </font></div></td>
        </tr>
      </table>
      <% End If ' end Not rsKirimAppr.EOF Or NOT rsKirimAppr.BOF %>
      <p>&nbsp;</p>
      
      
      
      
       
      <p align="center">&nbsp;</p>
      </td>
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
rsKirimMS.Close()
Set rsKirimMS = Nothing
%>
<%
rsKirimAppr.Close()
Set rsKirimAppr = Nothing
%>
