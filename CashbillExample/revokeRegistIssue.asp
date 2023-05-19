<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 취소 현금영수증 데이터를 팝빌에 저장과 동시에 발행하여 "발행완료" 상태로 처리합니다.
    ' - 취소 현금영수증의 금액은 원본 금액을 넘을 수 없습니다.
    ' - 현금영수증 국세청 전송 정책 [https://developers.popbill.com/guide/cashbill/asp/introduction/policy-of-send-to-nts]
    ' - 취소 현금영수증 발행 시 구매자 메일주소로 발행 안내 베일이 전송되니 유의하시기 바랍니다.
    ' - https://developers.popbill.com/reference/cashbill/asp/api/issue#RevokeRegistIssue
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    ' 팝빌회원 아이디
    userID = "testkorea"

    ' 문서번호, 가맹점 사업자번호 단위 고유번호 할당, 1~24자리 영문,숫자조합으로 중복없이 구성.
    mgtKey = "20220720-ASP-102"

    ' 원본 현금영수증 국세청승인번호
    orgConfirmNum = "TB0000102"

    ' 원본 현금영수증 거래일자
    orgTradeDate = "20221108"

    ' 발행안내 문자 전송여부
    smssendYN = False

    ' 메모
    memo = "즉시발행 메모"

    ' 안내메일 제목, 공백처리시 기본양식으로 전송
    emailSubject = ""

    ' 거래일시, 날짜(yyyyMMddHHmmss)
    ' 당일, 전일만 가능, 미입력시 기본값 발행일시 처리
    tradeDT = ""

    On Error Resume Next

    Set Presponse = m_CashbillService.RevokeRegistIssue(testCorpNum, mgtKey, orgConfirmNum, orgTradeDate, smssendYN, memo, userID, emailSubject, tradeDT)

    If Err.Number <> 0 then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else
        code = Presponse.code
        message = Presponse.message
        confirmNum = Presponse.confirmNum
        tradeDate = Presponse.tradeDate
        tradeDT = Presponse.tradeDT
    End If

    On Error GoTo 0

%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>취소현금영수증 즉시발행</legend>
                <ul>
                    <li> Response.code : <%=code%> </li>
                    <li> Response.message : <%=message%> </li>
                    <% If confirmNum <> "" Then %>
                    <li> Response.confirmNum : <%=confirmNum%> </li>
                    <% End If %>
                    <% If tradeDate <> "" Then %>
                    <li> Response.tradeDate : <%=tradeDate%> </li>
                    <% End If %>
                    <% If tradeDT <> "" Then %>
                    <li> Response.tradeDT : <%=tradeDT%> </li>
                    <% End If %>
                </ul>
            </fieldset>
         </div>
    </body>
</html>