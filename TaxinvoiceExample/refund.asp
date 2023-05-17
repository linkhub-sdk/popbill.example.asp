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
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/point#Refund
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    Dim m_RefundForm : Set m_RefundForm = New RefundForm
    '담당자명
    m_RefundForm.ContactUame = "담당자 이름"

    '담당자 연락처
    m_RefundForm.TEL = "010-1234-1234"

    '환불 신청 포인트
    m_RefundForm.RequestPoint = "1000"

    '은행명
    m_RefundForm.AccountBank = "신한"

    '계좌번호
    m_RefundForm.AccountNum = "110-1234-12345"

    '예금주명
    m_RefundForm.AccountName = "예금주_테스트"

    '환불사유
    m_RefundForm.Reason = "환불하겠습니다"


    '팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_TaxinvoiceService.Refund(testCorpNum, m_RefundForm, UserID)

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
                        <legend> result </legend>
                            <ul>
                                <li> code (응답 코드) : <%=result.code%></li>
                                <li> message (응답 메시지) : <%=result.message%></li>
                                <li> refundCode (환불코드) : <%=result.refundCode%></li>
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