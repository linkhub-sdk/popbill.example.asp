<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 연동회원의 포인트 환불신청내역을 확인합니다.
    ' - https://developers.popbill.com/reference/accountcheck/asp/api/point#GetRefundableBalance
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    '팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    Set RefundHistoryResult = m_AccountCheckService.GetRefundHistory(testCorpNum, UserID)

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
                <legend>연동회원 포인트 환불내역 확인</legend>
                <%
                    If code = 0 Then
                %>
                    <fieldset class="fieldset2">
                        <legend> CorpInfo </legend>
                            <ul>
                                <li> code (응답 코드) : <%=code%></li>
                                <li> total (총 검색결과 건수) : <%=total%></li>
                                <li> perPage (페이지당 검색개수) : <%=perPage%></li>
                                <li> pageNum (페이지 번호) : <%=pageNum%></li>
                                <li> perCount (페이지 개수) : <%=perCount%></li>

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