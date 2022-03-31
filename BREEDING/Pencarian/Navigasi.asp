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

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>:: Pencarian ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
<!--
function mmLoadMenus() {
  if (window.mm_menu_0806111527_0) return;
          window.mm_menu_0806111527_0 = new Menu("root",106,17,"",11,"#FFFF00","#FF0000","#330066","#33FF99","left","middle",3,0,500,-5,7,true,true,true,0,true,true);
  mm_menu_0806111527_0.addMenuItem("Budget&nbsp;ID","window.open('SearchBudgetID.asp', 'mainFrame');");
  mm_menu_0806111527_0.addMenuItem("Budget&nbsp;Name","window.open('SearchBudgetName.asp', 'mainFrame');");
  mm_menu_0806111527_0.addMenuItem("Total&nbsp;Budget","window.open('SearchBudTotal.asp', 'mainFrame');");
  mm_menu_0806111527_0.addMenuItem("Saldo&nbsp;Budget","window.open('SearchBudSaldo.asp', 'mainFrame');");
  mm_menu_0806111527_0.addMenuItem("Post&nbsp;By","window.open('SearchBudPostBy.asp', 'mainFrame');");
   mm_menu_0806111527_0.fontWeight="bold";
   mm_menu_0806111527_0.hideOnMouseOut=true;
   mm_menu_0806111527_0.bgColor='#555555';
   mm_menu_0806111527_0.menuBorder=1;
   mm_menu_0806111527_0.menuLiteBgColor='';
   mm_menu_0806111527_0.menuBorderBgColor='#FFFF00';
  window.mm_menu_0806112043_0 = new Menu("root",122,17,"",11,"#FFFF00","#FF0000","#330066","#33FF99","left","middle",3,0,500,-5,7,true,true,true,0,true,true);
  mm_menu_0806112043_0.addMenuItem("Company&nbsp;ID","window.open('SearchCompID.asp', 'mainFrame');");
  mm_menu_0806112043_0.addMenuItem("Company&nbsp;Name","window.open('SearchCompName.asp', 'mainFrame');");
  mm_menu_0806112043_0.addMenuItem("Post&nbsp;By","window.open('SearchCompPostBy.asp', 'mainFrame');");
   mm_menu_0806112043_0.fontWeight="bold";
   mm_menu_0806112043_0.hideOnMouseOut=true;
   mm_menu_0806112043_0.bgColor='#555555';
   mm_menu_0806112043_0.menuBorder=1;
   mm_menu_0806112043_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0806112043_0.menuBorderBgColor='#FFFF00';
window.mm_menu_0806112209_0 = new Menu("root",119,17,"",11,"#FFFF00","#FF0000","#330066","#33FF99","left","middle",3,0,500,-5,7,true,true,true,0,true,true);
  mm_menu_0806112209_0.addMenuItem("Currency&nbsp;ID","window.open('SearchCurrID.asp', 'mainFrame');");
  mm_menu_0806112209_0.addMenuItem("Currency&nbsp;Name","window.open('SearchCurrName.asp', 'mainFrame');");
  mm_menu_0806112209_0.addMenuItem("Post&nbsp;By","window.open('SearchCurrPostBy.asp', 'mainFrame');");
   mm_menu_0806112209_0.fontWeight="bold";
   mm_menu_0806112209_0.hideOnMouseOut=true;
   mm_menu_0806112209_0.bgColor='#555555';
   mm_menu_0806112209_0.menuBorder=1;
   mm_menu_0806112209_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0806112209_0.menuBorderBgColor='#FFFF00';
  window.mm_menu_0806112240_0 = new Menu("root",101,17,"",11,"#FFFF00","#FF0000","#330066","#33FF99","left","middle",3,0,500,-5,7,true,true,true,0,true,true);
  mm_menu_0806112240_0.addMenuItem("Divisi&nbsp;ID","window.open('SearchDivID.asp', 'mainFrame');");
  mm_menu_0806112240_0.addMenuItem("Company&nbsp;ID","window.open('SearchDivComID.asp', 'mainFrame');");
  mm_menu_0806112240_0.addMenuItem("Divisi&nbsp;Name","window.open('SearchDivName.asp', 'mainFrame');");
  mm_menu_0806112240_0.addMenuItem("Post&nbsp;By","window.open('SearchDivPostBy.asp', 'mainFrame');");
   mm_menu_0806112240_0.fontWeight="bold";
   mm_menu_0806112240_0.hideOnMouseOut=true;
   mm_menu_0806112240_0.bgColor='#555555';
   mm_menu_0806112240_0.menuBorder=1;
   mm_menu_0806112240_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0806112240_0.menuBorderBgColor='#FFFF00';
  window.mm_menu_0806112351_0 = new Menu("root",92,17,"",11,"#FFFF00","#FF0000","#330066","#33FF99","left","middle",3,0,500,-5,7,true,true,true,0,true,true);
  mm_menu_0806112351_0.addMenuItem("Jabatan","window.open('SearchUsrJabatan.asp', 'mainFrame');");
  mm_menu_0806112351_0.addMenuItem("User&nbsp;ID","window.open('SearchUsrID.asp', 'mainFrame');");
  mm_menu_0806112351_0.addMenuItem("User&nbsp;Name","window.open('SearchUsrName.asp', 'mainFrame');");
  mm_menu_0806112351_0.addMenuItem("User&nbsp;Level","window.open('SearchUsrLevel.asp', 'mainFrame');");
  mm_menu_0806112351_0.addMenuItem("User&nbsp;Status","window.open('SearchUsrStat.asp', 'mainFrame');");
  mm_menu_0806112351_0.addMenuItem("Post&nbsp;By","window.open('SearchUsrPostBy.asp', 'mainFrame');");
   mm_menu_0806112351_0.fontWeight="bold";
   mm_menu_0806112351_0.hideOnMouseOut=true;
   mm_menu_0806112351_0.bgColor='#555555';
   mm_menu_0806112351_0.menuBorder=1;
   mm_menu_0806112351_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0806112351_0.menuBorderBgColor='#FFFF00';
  window.mm_menu_0806112648_0 = new Menu("root",116,17,"",11,"#FFFF00","#FF0000","#330066","#33FF99","left","middle",3,0,500,-5,7,true,true,true,0,true,true);
  mm_menu_0806112648_0.addMenuItem("Vendor&nbsp;ID","window.open('SearchVendorID.asp', 'mainFrame');");
  mm_menu_0806112648_0.addMenuItem("Vendor&nbsp;Name","window.open('SearchVenName.asp', 'mainFrame');");
  mm_menu_0806112648_0.addMenuItem("Contact&nbsp;Person","window.open('SearchVenCp.asp', 'mainFrame');");
  mm_menu_0806112648_0.addMenuItem("Post&nbsp;By","window.open('SearchVenPostBy.asp', 'mainFrame');");
   mm_menu_0806112648_0.fontWeight="bold";
   mm_menu_0806112648_0.hideOnMouseOut=true;
   mm_menu_0806112648_0.bgColor='#555555';
   mm_menu_0806112648_0.menuBorder=1;
   mm_menu_0806112648_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0806112648_0.menuBorderBgColor='#FFFF00';

              window.mm_menu_0825195750_0 = new Menu("root",108,17,"",11,"#FFFF00","#FF0000","#330066","#33FF99","left","middle",3,0,500,-5,7,true,true,true,0,true,true);
  mm_menu_0825195750_0.addMenuItem("Produk&nbsp;ID","window.open('SearchProdID.asp', 'mainFrame');");
  mm_menu_0825195750_0.addMenuItem("User&nbsp;ID","window.open('SearchUserID.asp', 'mainFrame');");
  mm_menu_0825195750_0.addMenuItem("Nama&nbsp;Produk","window.open('SearchNamaProduk.asp', 'mainFrame');");
   mm_menu_0825195750_0.fontWeight="bold";
   mm_menu_0825195750_0.hideOnMouseOut=true;
   mm_menu_0825195750_0.bgColor='#555555';
   mm_menu_0825195750_0.menuBorder=1;
   mm_menu_0825195750_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0825195750_0.menuBorderBgColor='#FFFF00';

      window.mm_menu_0825205744_0 = new Menu("root",96,17,"",11,"#FFFF00","#FF0000","#330066","#33FF99","left","middle",3,0,500,-5,7,true,true,true,0,true,true);
  mm_menu_0825205744_0.addMenuItem("User&nbsp;ID","window.open('SearchUserID_Mitra.asp', 'mainFrame');");
  mm_menu_0825205744_0.addMenuItem("Nama&nbsp;Mitra","window.open('SearchNama_Mitra.asp', 'mainFrame');");
  mm_menu_0825205744_0.addMenuItem("Kota","window.open('Searchcity.asp', 'mainFrame');");
   mm_menu_0825205744_0.fontWeight="bold";
   mm_menu_0825205744_0.hideOnMouseOut=true;
   mm_menu_0825205744_0.bgColor='#555555';
   mm_menu_0825205744_0.menuBorder=1;
   mm_menu_0825205744_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0825205744_0.menuBorderBgColor='#FFFF00';

        window.mm_menu_0828074303_0 = new Menu("root",122,17,"",11,"#FFFF00","#FF0000","#330066","#33FF99","left","middle",3,0,500,-5,7,true,true,true,0,true,true);
  mm_menu_0828074303_0.addMenuItem("Supplier&nbsp;ID","window.open('SearchsupplierID.asp', 'mainFrame');");
  mm_menu_0828074303_0.addMenuItem("Company&nbsp;Name","window.open('SearchcompanyName.asp', 'mainFrame');");
  mm_menu_0828074303_0.addMenuItem("Supplier&nbsp;Name","window.open('SearchSupplierName.asp', 'mainFrame');");
  mm_menu_0828074303_0.addMenuItem("City","window.open('Searchcity_supplier.asp', 'mainFrame');");
  mm_menu_0828074303_0.addMenuItem("Region","window.open('Search_region.asp', 'mainFrame');");
   mm_menu_0828074303_0.fontWeight="bold";
   mm_menu_0828074303_0.hideOnMouseOut=true;
   mm_menu_0828074303_0.bgColor='#555555';
   mm_menu_0828074303_0.menuBorder=1;
   mm_menu_0828074303_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0828074303_0.menuBorderBgColor='#FFFF00';

      window.mm_menu_0909191624_0 = new Menu("root",72,17,"",11,"#FFFF00","#FF0000","#330066","#33FF99","left","middle",3,0,500,-5,7,true,true,true,0,true,true);
  mm_menu_0909191624_0.addMenuItem("ID&nbsp;Area","window.open('SearchIDareafor_SA.asp', 'mainFrame');");
  mm_menu_0909191624_0.addMenuItem("Area","window.open('Searchareafor_SA.asp', 'mainFrame');");
   mm_menu_0909191624_0.fontWeight="bold";
   mm_menu_0909191624_0.hideOnMouseOut=true;
   mm_menu_0909191624_0.bgColor='#555555';
   mm_menu_0909191624_0.menuBorder=1;
   mm_menu_0909191624_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0909191624_0.menuBorderBgColor='#FFFF00';

mm_menu_0909191624_0.writeMenus();
} // mmLoadMenus()

function MM_displayStatusMsg(msgStr) { //v1.0
  status=msgStr;
  document.MM_returnValue = true;
}
//-->
</script>
<script language="JavaScript" src="../../Capex/Search/mm_menu.js"></script>
<link href="../../style.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#330066" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<p> 
  <script language="JavaScript1.2">mmLoadMenus();</script>
</p>
<p>&nbsp; </p>
<p align="center"><font color="#FFFF00"><strong>Search By<br>
  _________________ </strong></font></p>
<p> <a href="javascript:;" onMouseOver="MM_showMenu(window.mm_menu_0825195750_0,0,25,null,'image1');MM_displayStatusMsg('ms.Prod');return document.MM_returnValue" onMouseOut="MM_startTimeout();"><img src="../img/prod_search.jpg" name="image1" width="125" height="25" border="0" id="image1"></a><br>
  <br>
  <a href="javascript:;" onMouseOver="MM_displayStatusMsg('ms.Mitra');MM_showMenu(window.mm_menu_0825205744_0,0,25,null,'image2');return document.MM_returnValue" onMouseOut="MM_startTimeout();"><img src="../img/mitra.jpg" name="image2" width="125" height="25" border="0" id="image2"></a> 
  <br>
  <br>
  <a href="javascript:;" onMouseOver="MM_displayStatusMsg('ms.Supplier');MM_showMenu(window.mm_menu_0828074303_0,0,25,null,'image3');return document.MM_returnValue" onMouseOut="MM_startTimeout();"><img src="../img/supplier.jpg" name="image3" width="125" height="25" border="0" id="image3"></a> 
  <br>
  <br>
  <a href="javascript:;" onMouseOver="MM_showMenu(window.mm_menu_0909191624_0,0,25,null,'image4')" onMouseOut="MM_startTimeout();"><img src="../img/Area_search.jpg" name="image4" width="125" height="25" border="0" id="image4"></a><br>
  <br>
  <br>
  <br>
  <br>
  <font color="#FFFF00">_________________</font></p>
<p align="center">| <a href="<%= MM_Logout %>" target="_parent">logout</a> |</p>
<p>&nbsp;</p>
</body>
</html>
