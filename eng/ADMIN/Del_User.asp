<%@LANGUAGE="JAVASCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/DBConn.asp" -->
<%
// *** Edit Operations: declare variables

// set the form action variable
var MM_editAction = Request.ServerVariables("SCRIPT_NAME");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

// boolean to abort record edit
var MM_abortEdit = false;

// query string to execute
var MM_editQuery = "";
%>
<%
// *** Delete Record: declare variables

if (String(Request("MM_delete")) == "form1" &&
    String(Request("MM_recordId")) != "undefined") {

  var MM_editConnection = MM_DBConn_STRING;
  var MM_editTable = "dbo.members";
  var MM_editColumn = "memberID";
  var MM_recordId = "" + Request.Form("MM_recordId") + "";
  var MM_editRedirectUrl = "br_User.asp";

  // append the query string to the redirect URL
  if (MM_editRedirectUrl && Request.QueryString && Request.QueryString.Count > 0) {
    MM_editRedirectUrl += ((MM_editRedirectUrl.indexOf('?') == -1)?"?":"&") + Request.QueryString;
  }
}
%>
<%
// *** Delete Record: construct a sql delete statement and execute it

if (String(Request("MM_delete")) != "undefined" &&
    String(Request("MM_recordId")) != "undefined") {

  // create the sql delete statement
  MM_editQuery = "delete from " + MM_editTable + " where " + MM_editColumn + " = " + MM_recordId;

  if (!MM_abortEdit) {
    // execute the delete
    var MM_editCmd = Server.CreateObject('ADODB.Command');
    MM_editCmd.ActiveConnection = MM_editConnection;
    MM_editCmd.CommandText = MM_editQuery;
    MM_editCmd.Execute();
    MM_editCmd.ActiveConnection.Close();

    if (MM_editRedirectUrl) {
      Response.Redirect(MM_editRedirectUrl);
    }
  }

}
%>
<%
var rsUser__MMColParam = "1";
if (String(Request.QueryString("memberID")) != "undefined" && 
    String(Request.QueryString("memberID")) != "") { 
  rsUser__MMColParam = String(Request.QueryString("memberID"));
}
%>
<%
var rsUser = Server.CreateObject("ADODB.Recordset");
rsUser.ActiveConnection = MM_DBConn_STRING;
rsUser.Source = "SELECT * FROM dbo.members WHERE memberID = "+ rsUser__MMColParam.replace(/'/g, "''") + "";
rsUser.CursorType = 0;
rsUser.CursorLocation = 2;
rsUser.LockType = 1;
rsUser.Open();
var rsUser_numRows = 0;
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
                        <p>LogOut</p>
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
        <table cellspacing=0 cellpadding=2 width="100%" border=0>
          <tbody>
            <tr> 
              <td class=tdtitle>Delete Record User</td>
            </tr>
            <tr valign=top> 
              <td class=tdlight> <form ACTION="<%=MM_editAction%>" name="form1" method="POST">
                  
                  <TABLE border=0 cellPadding=2 cellSpacing=1 width="100%">
                    <TBODY>
                      <TR> 
                        <TD width=130 align=right valign="top" class=tddark><div align="right">Member 
                            ID *</div></TD>
                        <TD width=9 valign="top" class=tddark>:</TD>
                        <TD width="426" valign="top" class=tdlight><%=(rsUser.Fields.Item("memberID").Value)%> </TD>
                      </TR>
                      <TR> 
                        <TD align=right valign="top" class=tddark>Dept ID * </TD>
                        <TD valign="top" class=tddark>:</TD>
                        <TD valign="top" class=tdlight><%=(rsUser.Fields.Item("DeptID").Value)%> </TD>
                      </TR>
                      <TR> 
                        <TD align=right valign="top" class=tddark>Name *</TD>
                        <TD valign="top" class=tddark>:</TD>
                        <TD valign="top" class=tdlight><%=(rsUser.Fields.Item("Name").Value)%></TD>
                      </TR>
                      <TR> 
                        <TD align=right valign="top" class=tddark>Login *</TD>
                        <TD valign="top" class=tddark>:</TD>
                        <TD valign="top" class=tdlight><%=(rsUser.Fields.Item("Login").Value)%></TD>
                      </TR>
                      <TR> 
                        <TD align=right valign="top" class=tddark>Password *</TD>
                        <TD valign="top" class=tddark>:</TD>
                        <TD valign="top" class=tdlight><%=(rsUser.Fields.Item("Password").Value)%></TD>
                      </TR>
                      <TR valign="top"> 
                        <TD align=right class=tddark>Level ID * </TD>
                        <TD class=tddark>:</TD>
                        <TD class=tdlight><%=(rsUser.Fields.Item("levelID").Value)%></TD>
                      </TR>
                      <TR> 
                        <TD class=tdbrowsetitle colSpan=3 vAlign=center><input style="BACKGROUND-COLOR: #ffecec; BORDER-BOTTOM: #999999 1px solid; BORDER-LEFT: #333333 1px solid; BORDER-RIGHT: #999999 1px solid; BORDER-TOP: #333333 1px solid; COLOR: #000000; FONT-SIZE: 11px; HEIGHT: 18px; WIDTH: 50px" type="submit" name="Submit" value="Delete"> 
                        </TD>
                      </TR>
                      <TR> 
                        <TD class=tdbrowsetitle colSpan=3>*) Must be filled..</TD>
                      </TR>
                  </TABLE>
                  <p> 
                    <input type="hidden" name="MM_delete" value="form1">
                  </p>
                
                  <input type="hidden" name="MM_recordId" value="<%=(rsUser.Fields.Item("memberID").Value)%>">
                </form>
                <TABLE cellSpacing=0 cellPadding=2 width="100%" border=0>
                  <FORM  >
                    <TBODY>
                      <TR align=middle> 
                        <TD><INPUT name="button2" type=button class=btn onclick=history.back() value=Back> 
                        </TD>
                      </TR>
                  </FORM>
                </TABLE>
                <br></td>
            </tr>
          </tbody>
        </table>
        <!-- InstanceEndEditable --></TD>
    </TR>
  </TBODY>
</TABLE>
</body>
<!-- InstanceEnd --></html>
<%
rsUser.Close();
%>
