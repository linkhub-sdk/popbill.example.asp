<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 연동회원의 회사정보를 확인합니다.
    ' - https://developers.popbill.com/reference/htcashbill/asp/api/member#QuitMember
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    '탈퇴 사유
    Set m_QuitReason = New QuitReason

    m_QuitReason.quitReason = "탈퇴사유"

    '팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    Set response = m_HTCashbillService.QuitMember(testCorpNum, m_QuitReason, UserID)

    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        Err.Clears
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>환불 가능 포인트 조회</legend>
                <%
                    If code = 0 Then
                %>
                    <fieldset class="fieldset2">
                        <legend> CorpInfo </legend>
                            <ul>
                                <li> code (응답 코드) : <%=response.code%></li>
                                <li> message (응답 메시지) : <%=response.message%></li>
                            </ul>
                        </fieldset>
                <%
                    Else
                %>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>
                <%
                    End If
                %>
            </fieldset>
         </div>
    </body>
</html>