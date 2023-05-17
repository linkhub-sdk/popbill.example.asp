<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 홈택스연동 인증을 위해 팝빌에 현금영수증 자료조회 부서사용자 계정을 등록합니다.
    ' - https://developers.popbill.com/reference/htcashbill/asp/api/cert#RegistDeptUser
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    '홈택스에서 생성한 현금영수증 부서사용자 아이디
    deptUserID = "userid"

    '홈택스에서 생성한 현금영수증 부서사용자 비밀번호
    deptUserPWD = "pwd"

    ' 팝빌회원 아이디
    userID = "testkorea"

    On Error Resume Next

    Set Presponse = m_HTCashbillService.RegistDeptUser(testCorpNum, deptUserID, deptUserPWD, userID)

    If Err.Number <> 0 then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else
        code = Presponse.code
        message = Presponse.message
    End If


    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>부서사용자 계정등록</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>