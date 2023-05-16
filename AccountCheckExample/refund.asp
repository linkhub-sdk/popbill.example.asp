<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 연동회원 포인트를 환불 신청합니다.
    ' - https://developers.popbill.com/reference/accountcheck/asp/api/point#Refund
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    Set RefundForm = New RefundForm
    '담당자명
    RefundForm, "contactName", ""

    '담당자 연락처
    RefundForm, "tel", ""

    '환불 신청 포인트
    RefundForm, "requestPoint", ""

    '은행명
    RefundForm, "accountBank", ""

    '계좌번호
    RefundForm, "accountNum", ""

    '예금주명
    RefundForm, "accountName", ""

    '환불사유
    RefundForm, "reason", ""


    '팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    Set refundResponse = m_AccountCheckService.Refund(testCorpNum, RefundForm, UserID)

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
                <legend>연동회원 포인트 환불신청</legend>
                <%
                    If code = 0 Then
                %>
                    <fieldset class="fieldset2">
                        <legend> refundResponse </legend>
                            <ul>
                                <li> code (응답 코드) : <%=refundResponse.code%></li>
                                <li> message (응답 메시지) : <%=refundResponse.message%></li>
                                <li> refundCode (환불코드) : <%=refundResponse.refundCode%></li>
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