<%@LANGUAGE="JAVASCRIPT" CODEPAGE="1252"%>
<%
// *** Logout the current user.
MM_Logout = String(Request.ServerVariables("URL")) + "?MM_Logoutnow=1";
if (String(Request("MM_Logoutnow"))=="1") {
  Session.Contents.Remove("MM_Username");
  Session.Contents.Remove("MM_UserAuthorization");
  var MM_logoutRedirectPage = "login.asp";
  // redirect with URL parameters (remove the "MM_Logoutnow" query param).
  if (MM_logoutRedirectPage == "") MM_logoutRedirectPage = String(Request.ServerVariables("URL"));
  if (String(MM_logoutRedirectPage).indexOf("?") == -1 && Request.QueryString != "") {
    var MM_newQS = "?";
    for (var items=new Enumerator(Request.QueryString); !items.atEnd(); items.moveNext()) {
      if (String(items.item()) != "MM_Logoutnow") {
        if (MM_newQS.length > 1) MM_newQS += "&";
        MM_newQS += items.item() + "=" + Server.URLencode(Request.QueryString(items.item()));
      }
    }
    if (MM_newQS.length > 1) MM_logoutRedirectPage += MM_newQS;
  }
  Response.Redirect(MM_logoutRedirectPage);
}
%>
<!--#include file="../../Connections/DBConn.asp" -->
<%
// *** Restrict Access To Page: Grant or deny access to this page
var MM_authorizedUsers="5,6,4";
var MM_authFailedURL="Failed.asp";
var MM_grantAccess=false;
if (String(Session("MM_Username")) != "undefined") {
  if (false || (String(Session("MM_UserAuthorization"))=="") || (MM_authorizedUsers.indexOf(String(Session("MM_UserAuthorization"))) >=0)) {
    MM_grantAccess = true;
  }
}
if (!MM_grantAccess) {
  var MM_qsChar = "?";
  if (MM_authFailedURL.indexOf("?") >= 0) MM_qsChar = "&";
  var MM_referrer = Request.ServerVariables("URL");
  if (String(Request.QueryString()).length > 0) MM_referrer = MM_referrer + "?" + String(Request.QueryString());
  MM_authFailedURL = MM_authFailedURL + MM_qsChar + "accessdenied=" + Server.URLEncode(MM_referrer);
  Response.Redirect(MM_authFailedURL);
}
%>
<%
var rsNews__MMColParam = "1";
if (String(Request.QueryString("ID")) != "undefined" && 
    String(Request.QueryString("ID")) != "") { 
  rsNews__MMColParam = String(Request.QueryString("ID"));
}
%>
<%
var rsNews = Server.CreateObject("ADODB.Recordset");
rsNews.ActiveConnection = MM_DBConn_STRING;
rsNews.Source = "SELECT *  FROM dbo.News  WHERE ID = "+ rsNews__MMColParam.replace(/'/g, "''") + "";
rsNews.CursorType = 0;
rsNews.CursorLocation = 2;
rsNews.LockType = 1;
rsNews.Open();
var rsNews_numRows = 0;
%>
<html><!-- InstanceBegin template="/Templates/Admin2.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" --> 
<title>:: Website Administration ::</title>
<!-- InstanceEndEditable --> 
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable --> 
<link href="../../ADMIN/css/style.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0">
<!-- memanggil file header and menu-->
<!-- BEGIN HEADER -->
<TABLE cellSpacing=0 cellPadding=1 width="100%" border=0>
  <TBODY>
    <TR align=middle> 
      <TD class=tdheader height=75>..:: Sierad Produce Website Administration 
        ::..</TD>
    </TR>
    <TR> 
      <TD height=1><IMG 
      src="../../ADMIN/css/spacer.gif" 
      width=1 height=1 class="btn_file"></TD>
    </TR>
  </TBODY>
</TABLE>
<!-- END HEADER -->
<!-- BEGIN MENU -->
<TABLE class=tbtitle cellSpacing=0 cellPadding=1 width=200 align=left 
  border=0>
  <TBODY>
    <TR> 
      <TD class=tdtitle> <TABLE cellSpacing=0 cellPadding=2 width="100%" border=0>
          <TBODY>
            <TR> 
              <TD class=tdtitle>Menu</TD>
            </TR>
            <TR> 
              <TD class=tdverylight> <TABLE class=tdlight cellSpacing=1 cellPadding=2 width="100%" 
            border=0>
                  <TBODY>
                    <TR> 
                      <TD class=tdlight>&nbsp;</TD>
                    </TR>
                    <TR> 
                      <TD class=tdlight>User Management<BR> 
                        <LI type=square><A 
                  href="../../ADMIN/br_User.asp">Browse user</A> </LI></TD>
                    </TR>
                    <TR> 
                      <TD class=tdlight>Corporate Overview<BR> 
                        <LI type=square><A 
                  href="../../ADMIN/br_Visi.asp">Vision &amp; Mision</A> </LI>
                        <LI type=square><a href="../../ADMIN/br_sejarah.asp">History 
                          Background</a> </LI></TD>
                    </TR>
                    <TR> 
                      <TD height="29" valign="top" class=tdlight>Business Structure<BR>
                        <LI type=square><a href="../../ADMIN/br_bisnis.asp">Browse 
                          Business Structure</a></LI>
                        </TD>
                    </TR>
                    <TR> 
                      <TD class=tdlight>News Management<BR> 
                        <LI type=square>Browse News <a href="../../ADMIN/br_News.asp">[INA]</a> 
                          <a href="br_News.asp">[EN] </a></LI></TD>
                    </TR>
                    <TR> 
                      <TD class=tdlight>Careers Management<BR> 
                        <LI type=square><A class=linked1 
                  href="../../ADMIN/br_Careers.asp">Careers </A> </LI></TD>
                    </TR>
                    <TR> 
                      <TD class=tdlight>Product<BR> 
                        <LI type=square>Browse Product <a href="../../ADMIN/br_product.asp">[INA]</a> 
                          <a href="br_product.asp">[EN]</a> 
                      </TD>
                    </TR>
                    <TR> 
                      <TD class=tdlight>FAQ Management (Disable)<BR> 
                        <LI type=square>Browse FAQ </LI></TD>
                    </TR>
                    <TR> 
                      <TD class=tdlight>Annual Report (Disable)<BR> 
                        <LI type=square>Browse Annual Report</TD>
                    </TR>
                    <TR> 
                      <TD class=tdlight>Contact Management<BR> 
                        <LI type=square><A 
                  href="../../ADMIN/br_Contact.asp">Browse Contact Us</A> </LI></TD>
                    </TR>
                    <TR> 
                      <TD class=tdlight><!-- InstanceBeginEditable name="LogRegion" --> 
                        <p><a href="<%= MM_Logout %>">LogOut</a></p>
                        <!-- InstanceEndEditable --></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
  </TBODY>
</TABLE>
<!-- END MENU -->
<TABLE class=tbtitle cellSpacing=0 cellPadding=1 width=* border=0>
  <TBODY>
    <TR> 
      <TD width=1000 valign="top" class=tdtitle> <!-- InstanceBeginEditable name="EditRegion3" --> 
        <TABLE border=0 cellPadding=2 cellSpacing=0 width="100%"><TBODY>
          <TR> 
            <TD class=tdtitle>Modify News (English Version)</TD>
          </TR>
          <TR vAlign=top> 
            <TD class=tdverylight> <form action="upload_Edit_news.asp" method="post" enctype="multipart/form-data" name="form1">
                <TABLE border=0 cellPadding=2 cellSpacing=1 width="100%">
                  <TBODY>
                    <TR> 
                      <TD align=right class=tddark width=125>Title *</TD>
                      <TD class=tddark width=5>:</TD>
                      <TD class=tdlight vAlign=top> <INPUT 
                  name=p_title class=btn_text value="<%=(rsNews.Fields.Item("Title").Value)%>" size=40>
                        <input name="h_id" type="hidden" id="h_id" value="<%=(rsNews.Fields.Item("ID").Value)%>"> </TD>
                    </TR>
                    <TR> 
                      <TD align=right class=tddark vAlign=top>Lead * </TD>
                      <TD class=tddark vAlign=top>:</TD>
                      <TD class=tdlight vAlign=top><TEXTAREA class=btn_textarea cols=34 name=p_lead rows=5><%=(rsNews.Fields.Item("Clip").Value)%></TEXTAREA> 
                      </TD>
                    </TR>
                    <TR> 
                      <TD align=right class=tddark vAlign=top>Content * </TD>
                      <TD class=tddark vAlign=top>:</TD>
                      <TD class=tdlight><TEXTAREA class=btn_textarea cols=34 name=p_content rows=10><%=(rsNews.Fields.Item("Content").Value)%></TEXTAREA> 
                      </TD>
                    </TR>
                    <TR> 
                      <TD align=right class=tddark vAlign=top>Thumbnail &nbsp;</TD>
                      <TD class=tddark vAlign=top>:</TD>
                      <TD class=tdlight vAlign=top><INPUT class=btn_file 
                  name=p_thumbnail type=file> &nbsp;Max. 65 x 65 px <BR>
                        <img src="<%=(rsNews.Fields.Item("thumbnail").Value)%>"> 
                        <input name="h_thumbnail" type="hidden" id="h_thumbnail" value="<%=(rsNews.Fields.Item("thumbnail").Value)%>"> </TD>
                    </TR>
                    <TR> 
                      <TD align=right class=tddark vAlign=top>Picture &nbsp;</TD>
                      <TD class=tddark vAlign=top>:</TD>
                      <TD class=tdlight vAlign=top><p>
                          <INPUT class=btn_file 
                  name=p_picture type=file>
                          &nbsp;Max. 100 x 100 px <BR>
                          <img src="<%=(rsNews.Fields.Item("Image").Value)%>"> 
                          <input name="h_image" type="hidden" id="h_image" value="<%=(rsNews.Fields.Item("Image").Value)%>">
                        </p>
                        <p>
                          <select name="select">
                            <option value="Inggris" <%=(("Inggris" == (rsNews.Fields.Item("Lang").Value))?"SELECTED":"")%>>Inggris</option>
                            <% 
while (!rsNews.EOF) {
%>
                            <option value="<%=(rsNews.Fields.Item("Lang").Value)%>" <%=((rsNews.Fields.Item("Lang").Value == (rsNews.Fields.Item("Lang").Value))?"SELECTED":"")%> ><%=(rsNews.Fields.Item("Lang").Value)%></option>
                            <%
  rsNews.MoveNext();
}
if (rsNews.CursorType > 0) {
  if (!rsNews.BOF) rsNews.MoveFirst();
} else {
  rsNews.Requery();
}
%>
                          </select>
                        </p></TD>
                    </TR>
                    <TR> 
                      <TD class=tdbrowsetitle colSpan=3 vAlign=center><input style="BACKGROUND-COLOR: #ffecec; BORDER-BOTTOM: #999999 1px solid; BORDER-LEFT: #333333 1px solid; BORDER-RIGHT: #999999 1px solid; BORDER-TOP: #333333 1px solid; COLOR: #000000; FONT-SIZE: 11px; HEIGHT: 18px; WIDTH: 50px" type="submit" name="Submit" value="Submit"> 
                        <input style="BACKGROUND-COLOR: #ffecec; BORDER-BOTTOM: #999999 1px solid; BORDER-LEFT: #333333 1px solid; BORDER-RIGHT: #999999 1px solid; BORDER-TOP: #333333 1px solid; COLOR: #000000; FONT-SIZE: 11px; HEIGHT: 18px; WIDTH: 50px" type="reset" name="Submit2" value="Reset"> 
                      </TD>
                    </TR>
                    <TR> 
                      <TD class=tdbrowsetitle colSpan=3>*) Must be filled..</TD>
                    </TR>
                </TABLE>
              </form>
              <BR> <TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
                <FORM>
                  <TBODY>
                    <TR align=middle> 
                      <TD><INPUT name="button" type=button style="BACKGROUND-COLOR: #ffecec; BORDER-BOTTOM: #999999 1px solid; BORDER-LEFT: #333333 1px solid; BORDER-RIGHT: #999999 1px solid; BORDER-TOP: #333333 1px solid; COLOR: #000000; FONT-SIZE: 11px; HEIGHT: 18px; WIDTH: 50px" onclick=history.back() value=Back> 
                      </TD>
                    </TR>
                </FORM></TBODY>
              </TABLE>
              <BR></TD>
          </TR>
        </TABLE>
        <!-- InstanceEndEditable --></TD>
    </TR>
  </TBODY>
</TABLE>
</body>
<!-- InstanceEnd --></html>
<%
rsNews.Close();
%>
