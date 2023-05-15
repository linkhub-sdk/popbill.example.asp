<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 연동회원의 포인트 결제내역을 확인합니다.
    ' - https://developers.popbill.com/reference/accountcheck/asp/api/point#GetPaymentHistory
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    ' 조회 기간의 시작일자 (형식 : yyyyMMdd)
    SDate = "20230501"

    ' 조회 기간의 종료일자 (형식 : yyyyMMdd)
    EDate = "20230530"

    ' 목록 페이지번호
    Page = 1

    ' 페이지당 표시할 목록 개수
    PerPage = 500

    '팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    Set paymentHistoryResult = m_AccountCheckService.GetPaymentHistory(testCorpNum, SDate, EDate, Page, PerPage, UserID)

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
                <legend>연동회원 포인트 결제내역 확인</legend>
                <%
                    If code = 0 Then
                %>
                    <fieldset class="fieldset2">
                        <legend> CorpInfo </legend>
                            <ul>
                                <li> refundableBalance (환불 가능 포인트) : <%=refundableBalance%></li>
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