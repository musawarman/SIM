<%@LANGUAGE="JAVASCRIPT"%>
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
// *** Insert Record: set variables

if (String(Request("MM_insert")) == "form1") {

  var MM_editConnection = MM_DBConn_STRING;
  var MM_editTable  = "dbo.Profile";
  var MM_editRedirectUrl = "br_bisnis.asp";
  var MM_fieldsStr = "p_title|value|p_lead|value|p_content|value|select|value|hiddenField|value|select2|value";
  var MM_columnsStr = "Name|',none,''|Lead|',none,''|content|',none,''|CompId|',none,''|updateby|',none,''|Lang|',none,''";

  // create the MM_fields and MM_columns arrays
  var MM_fields = MM_fieldsStr.split("|");
  var MM_columns = MM_columnsStr.split("|");
  
  // set the form values
  for (var i=0; i+1 < MM_fields.length; i+=2) {
    MM_fields[i+1] = String(Request.Form(MM_fields[i]));
  }

  // append the query string to the redirect URL
  if (MM_editRedirectUrl && Request.QueryString && Request.QueryString.Count > 0) {
    MM_editRedirectUrl += ((MM_editRedirectUrl.indexOf('?') == -1)?"?":"&") + Request.QueryString;
  }
}
%>
<%
// *** Insert Record: construct a sql insert statement and execute it

if (String(Request("MM_insert")) != "undefined") {

  // create the sql insert statement
  var MM_tableValues = "", MM_dbValues = "";
  for (var i=0; i+1 < MM_fields.length; i+=2) {
    var formVal = MM_fields[i+1];
    var MM_typesArray = MM_columns[i+1].split(",");
    var delim =    (MM_typesArray[0] != "none") ? MM_typesArray[0] : "";
    var altVal =   (MM_typesArray[1] != "none") ? MM_typesArray[1] : "";
    var emptyVal = (MM_typesArray[2] != "none") ? MM_typesArray[2] : "";
    if (formVal == "" || formVal == "undefined") {
      formVal = emptyVal;
    } else {
      if (altVal != "") {
        formVal = altVal;
      } else if (delim == "'") { // escape quotes
        formVal = "'" + formVal.replace(/'/g,"''") + "'";
      } else {
        formVal = delim + formVal + delim;
      }
    }
    MM_tableValues += ((i != 0) ? "," : "") + MM_columns[i];
    MM_dbValues += ((i != 0) ? "," : "") + formVal;
  }
  MM_editQuery = "insert into " + MM_editTable + " (" + MM_tableValues + ") values (" + MM_dbValues + ")";

  if (!MM_abortEdit) {
    // execute the insert
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
var rsProfile = Server.CreateObject("ADODB.Recordset");
rsProfile.ActiveConnection = MM_DBConn_STRING;
rsProfile.Source = "SELECT * FROM dbo.Profile";
rsProfile.CursorType = 0;
rsProfile.CursorLocation = 2;
rsProfile.LockType = 1;
rsProfile.Open();
var rsProfile_numRows = 0;
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
            <TD class=tdtitle>Modify Profile</TD>
          </TR>
          <TR vAlign=top> 
            <TD class=tdverylight> <form ACTION="<%=MM_editAction%>" method="POST" name="form1">
                <TABLE border=0 cellPadding=2 cellSpacing=1 width="100%">
                  <TBODY>
                    <TR> 
                      <TD align=right class=tddark width=125>Company *</TD>
                      <TD class=tddark width=5>:</TD>
                      <TD class=tdlight vAlign=top> <INPUT 
                  name=p_title class=btn_text size=40> </TD>
                    </TR>
                    <TR> 
                      <TD align=right class=tddark vAlign=top>Leading * </TD>
                      <TD class=tddark vAlign=top>:</TD>
                      <TD class=tdlight vAlign=top><TEXTAREA class=btn_textarea cols=34 name=p_lead rows=5></TEXTAREA> 
                      </TD>
                    </TR>
                    <TR> 
                      <TD align=right class=tddark vAlign=top>Content * </TD>
                      <TD class=tddark vAlign=top>:</TD>
                      <TD class=tdlight><TEXTAREA class=btn_textarea cols=34 name=p_content rows=10></TEXTAREA> 
                      </TD>
                    </TR>
                    <TR> 
                      <TD align=right class=tddark vAlign=top>Division *</TD>
                      <TD class=tddark vAlign=top>&nbsp;</TD>
                      <TD class=tdlight vAlign=top> <select name="select">
                          <option value="Biotek">Biotek</option>
                          <option value="RPA">RPA</option>
                          <option value="Breeding">Breeding</option>
                          <option value="Feedmill">Feedmill</option>
                          <option value="Industri">Industri</option>
                          <option value="DMN">DMN</option>
                          <option value="Hartz">Hartz</option>
                          <option value="Wendys">Wendys</option>
                        </select> <input name="hiddenField" type="hidden" value="<%= Session("updateby") %>"> 
                      </TD>
                    </TR>
                    <TR> 
                      <TD align=right class=tddark vAlign=top>Language *</TD>
                      <TD class=tddark vAlign=top>&nbsp;</TD>
                      <TD class=tdlight vAlign=top><select name="select2">
                          <option value="Indonesia">Indonesia</option>
                          <option value="Inggris">Inggris</option>
                        </select></TD>
                    </TR>
                    <TR> 
                      <TD align=right class=tddark vAlign=top>&nbsp;</TD>
                      <TD class=tddark vAlign=top>&nbsp;</TD>
                      <TD class=tdlight vAlign=top>&nbsp; </TD>
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
                <input type="hidden" name="MM_insert" value="form1">
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
rsProfile.Close();
%>
