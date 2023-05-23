<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 연동회원 포인트 무통장 입금신청내역 1건을 확인합니다.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/point#GetSettleResult
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    '정산코드
    SettleCode = "202305120000000035"

    '팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_TaxinvoiceService.GetSettleResult(testCorpNum, SettleCode, UserID)

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
                <legend>연동회원 무통장 입금신청 정보확인</legend>
                <%
                    If code = 0 Then
                %>
                    <fieldset class="fieldset2">
                        <legend> PaymentHistory </legend>
                        <ul>
                            <li>productType (결제 내용) : <%= result.productType %></li>
                            <li>productName (결제 상품명) : <%= result.productName %></li>
                            <li>settleType (결제유형) : <%= result.settleType %></li>
                            <li>settlerName (담당자명) : <%= result.settlerName %></li>
                            <li>settlerEmail (담당자메일) : <%= result.settlerEmail %></li>
                            <li>settleCost (결제금액) : <%= result.settleCost %></li>
                            <li>settlePoint (충전포인트) : <%= result.settlePoint %></li>
                            <li>settleState (결제상태) : <%= result.settleState %></li>
                            <li>regDT (등록일시 ) : <%= result.regDT %></li>
                            <li>stateDT (상태일시 ) : <%= result.stateDT %></li>
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