<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 연동회원 포인트 충전을 위해 무통장입금을 신청합니다.
    ' - https://developers.popbill.com/reference/cashbill/asp/api/point#PaymentRequest
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    Dim m_PaymentForm : Set m_PaymentForm = New PaymentForm

    '담당자명
    m_PaymentForm.SettlerName = "담당자"

    '담당자 이메일
    m_PaymentForm.SettlerEmail = "email_damdang@email.com"


    '담당자 휴대폰
    m_PaymentForm.NotifyHP = "010-1234-1234"

    '입금자명
    m_PaymentForm.PaymentName = "입금자"

    '결제금액
    m_PaymentForm.SettleCost = "10000"

    '팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_CashbillService.PaymentRequest(testCorpNum, m_PaymentForm, UserID)

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
                        <legend> PaymentResponse </legend>
                            <ul>
                                <li> code (응답코드) : <%=result.code%></li>
                                <li> message (응답메시지) : <%=result.message%></li>
                                <li> settleCode (정산코드) : <%=result.settleCode%></li>
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