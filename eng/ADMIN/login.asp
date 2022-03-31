<%@LANGUAGE="JAVASCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/DBConn.asp" -->
<%
var rsMember__MMColParam = "1";
if (String(Request.QueryString("ADM")) != "undefined" && 
    String(Request.QueryString("ADM")) != "") { 
  rsMember__MMColParam = String(Request.QueryString("ADM"));
}
%>
<%
var rsMember = Server.CreateObject("ADODB.Recordset");
rsMember.ActiveConnection = MM_DBConn_STRING;
rsMember.Source = "SELECT memberID, DeptID, Name, Login, Password, levelID FROM dbo.members WHERE DeptID = '"+ rsMember__MMColParam.replace(/'/g, "''") + "'";
rsMember.CursorType = 0;
rsMember.CursorLocation = 2;
rsMember.LockType = 1;
rsMember.Open();
var rsMember_numRows = 0;
%>
<%
// *** Validate request to log in to this site.
var MM_LoginAction = Request.ServerVariables("URL");
if (Request.QueryString!="") MM_LoginAction += "?" + Request.QueryString;
var MM_valUsername=String(Request.Form("p_username"));
if (MM_valUsername != "undefined") {
  var MM_fldUserAuthorization="levelID";
  var MM_redirectLoginSuccess="Admin.asp";
  var MM_redirectLoginFailed="login.asp";
  var MM_flag="ADODB.Recordset";
  var MM_rsUser = Server.CreateObject(MM_flag);
  MM_rsUser.ActiveConnection = MM_DBConn_STRING;
  MM_rsUser.Source = "SELECT Login, Password";
  if (MM_fldUserAuthorization != "") MM_rsUser.Source += "," + MM_fldUserAuthorization;
  MM_rsUser.Source += " FROM dbo.members WHERE Login='" + MM_valUsername.replace(/'/g, "''") + "' AND Password='" + String(Request.Form("p_password")).replace(/'/g, "''") + "'";
  MM_rsUser.CursorType = 0;
  MM_rsUser.CursorLocation = 2;
  MM_rsUser.LockType = 3;
  MM_rsUser.Open();
  if (!MM_rsUser.EOF || !MM_rsUser.BOF) {
    // username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername;
	Session("updateBy") = Session("MM_Username");
    if (MM_fldUserAuthorization != "") {
      Session("MM_UserAuthorization") = String(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value);
    } else {
      Session("MM_UserAuthorization") = "";
    }
    if (String(Request.QueryString("accessdenied")) != "undefined" && true) {
      MM_redirectLoginSuccess = Request.QueryString("accessdenied");
    }
    MM_rsUser.Close();
    Response.Redirect(MM_redirectLoginSuccess);
  }
  MM_rsUser.Close();
  Response.Redirect(MM_redirectLoginFailed);
}
%>
<html>
<head>
<title>:: Sierad Produce ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/style.css" rel="stylesheet" type="text/css">
</head>

<body>
<table cellspacing=0 cellpadding=2 width="100%" border=0>
  <tbody>
    <tr> 
      <td class=tdtitle>Welcome </td>
    </tr>
    <tr valign=top> 
      <td class=tdlight><form name="form1" method="POST" action="<%=MM_LoginAction%>">
          <TABLE class=tbtitle cellSpacing=0 cellPadding=1 align=center border=0>
            <TBODY>
              <TR> 
                <TD> <TABLE cellSpacing=0 cellPadding=2 align=center border=0>
                    <TBODY>
                      <TR> 
                        <TD class=title2>Administrator Login </TD>
                      </TR>
                      <TR> 
                        <TD class=tdverylight> <TABLE class=tblight cellSpacing=1 cellPadding=2 width="100%" 
                  border=0>
                            <TBODY>
                              <TR> 
                                <TD align=right>Username * :</TD>
                                <TD><INPUT 
                        style="BORDER-RIGHT: #999999 1px solid; BORDER-TOP: #333333 1px solid; FONT-SIZE: 11px; BORDER-LEFT: #333333 1px solid; WIDTH: 150px; COLOR: #000000; BORDER-BOTTOM: #999999 1px solid; HEIGHT: 18px; BACKGROUND-COLOR: #ffecec" 
                        name=p_username> </TD>
                              </TR>
                              <TR> 
                                <TD align=right>Password * :</TD>
                                <TD><INPUT 
                        style="BORDER-RIGHT: #999999 1px solid; BORDER-TOP: #333333 1px solid; FONT-SIZE: 11px; BORDER-LEFT: #333333 1px solid; WIDTH: 150px; COLOR: #000000; BORDER-BOTTOM: #999999 1px solid; HEIGHT: 18px; BACKGROUND-COLOR: #ffecec" 
                        type=password name=p_password> </TD>
                              </TR>
                              <TR align=middle> 
                                <TD colSpan=2 height=50><INPUT class=btn type=submit value=Login name=submit> 
                                  <INPUT class=btn type=reset value=Reset name=reset> 
                                  <INPUT type=hidden value=login name=act> </TD>
                              </TR>
                            </TBODY>
                          </TABLE></TD>
                      </TR>
                    </TBODY>
                  </TABLE></TD>
              </TR>
            </TBODY>
          </TABLE>
        </form>
        <br> <br> </td>
    </tr>
  </tbody>
</table>
</body>
</html>
<%

rsMember.Close();
%>
