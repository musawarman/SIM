<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/simConn.asp" -->
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Request.QueryString
MM_valUsername=CStr(Request.Form("p_username"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization="UserLevel"
  MM_redirectLoginSuccess="administrator.asp"
  MM_redirectLoginFailed="login.asp"
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_simConn_STRING
  MM_rsUser.Source = "SELECT UserID, UserPassword"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM dbo.UserMS WHERE UserID='" & Replace(MM_valUsername,"'","''") &"' AND UserPassword='" & Replace(Request.Form("p_password"),"'","''") & "'"
  MM_rsUser.CursorType = 0
  MM_rsUser.CursorLocation = 2
  MM_rsUser.LockType = 3
  MM_rsUser.Open
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
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
<title>:: Login ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/style.css" rel="stylesheet" type="text/css">
</head>

<body>
<table cellspacing=0 cellpadding=2 width="100%" border=0>
  <tbody>
    <tr> 
      <td class=tdtitle>SELAMAT DATANG</td>
    </tr>
    <tr valign=top> 
      <td class=tdlight><form name="form1" method="POST" action="<%=MM_LoginAction%>">
          <TABLE class=tbtitle cellSpacing=0 cellPadding=1 align=center border=0>
            <TBODY>
              <TR> 
                <TD> <TABLE cellSpacing=0 cellPadding=2 align=center border=0>
                    <TBODY>
                      <TR> 
                        <TD class=title2><img src="../Icon/admin.gif" width="29" height="25">Administrator 
                          Login </TD>
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
