<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 작성된 (부분)취소 현금영수증 데이터를 팝빌에 저장과 동시에 발행하여 "발행완료" 상태로 처리합니다.
    ' - 취소 현금영수증의 금액은 원본 금액을 넘을 수 없습니다.
    ' - 현금영수증 국세청 전송 정책 [https://docs.popbill.com/cashbill/ntsSendPolicy?lang=asp]
    ' - "발행완료"된 취소 현금영수증은 국세청 전송 이전에 발행취소(cancelIssue API) 함수로 국세청 신고 대상에서 제외할 수 있습니다.
    ' - 취소 현금영수증 발행 시 구매자 메일주소로 발행 안내 베일이 전송되니 유의하시기 바랍니다.
    ' - https://docs.popbill.com/cashbill/asp/api#RevokeRegistIssue_Part
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"	

    ' 팝빌회원 아이디
    userID = "testkorea"				 

    ' 문서번호, 가맹점 사업자번호 단위 고유번호 할당, 1~24자리 영문,숫자조합으로 중복없이 구성.
    mgtKey = "20171115-01"

    ' 원본 현금영수증 국세청승인번호
    orgConfirmNum = "820116333"

    ' 원본 현금영수증 거래일자
    orgTradeDate = "20170711"

    ' 안내 문자 전송여부 , true / false 중 택 1
    ' └ true = 전송 , false = 미전송
    ' └ 원본 현금영수증의 구매자(고객)의 휴대폰번호 문자 전송
    smssendYN = False

    ' 메모
    memo = "즉시발행 메모"
    
    ' 현금영수증 취소유형 , true / false 중 택 1
    ' └ true = 부분 취소 , false = 전체 취소
    ' └ 미입력시 기본값 false 처리
    isPartCancel = True

    ' 취소사유 , 1 / 2 / 3 중 택 1
    ' └ 1 = 거래취소 , 2 = 오류발급취소 , 3 = 기타
    ' └ 미입력시 기본값 1 처리
    cancelType = 1

    ' [취소] 공급가액
    ' - 현금영수증 취소유형이 true 인 경우 취소할 공급가액 입력
    ' - 현금영수증 취소유형이 false 인 경우 미입력
    supplyCost = "5000"

    ' [취소] 부가세
    ' - 현금영수증 취소유형이 true 인 경우 취소할 부가세 입력
    ' - 현금영수증 취소유형이 false 인 경우 미입력
    tax = "500"

    ' [취소] 봉사료
    ' - 현금영수증 취소유형이 true 인 경우 취소할 봉사료 입력
    ' - 현금영수증 취소유형이 false 인 경우 미입력
    serviceFee = "0"
    
    ' [취소] 거래금액 (공급가액+부가세+봉사료)
    ' - 현금영수증 취소유형이 true 인 경우 취소할 거래금액 입력
    ' - 현금영수증 취소유형이 false 인 경우 미입력
    totalAmount = "5500"

    On Error Resume Next

    Set Presponse = m_CashbillService.RevokeRegistIssue_Part(testCorpNum, mgtKey, orgConfirmNum, orgTradeDate, smssendYN, memo, userID, _
        isPartCancel, cancelType, supplyCost, tax, serviceFee, totalAmount)

    If Err.Number <> 0 then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else 
        code = Presponse.code
        message = Presponse.message
        confirmNum = Presponse.confirmNum
        tradeDate = Presponse.tradeDate		
    End If

    On Error GoTo 0 

%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>(부분) 취소현금영수증 즉시발행</legend>
                <ul>
                    <li> Response.code : <%=code%> </li>
                    <li> Response.message : <%=message%> </li>
                    <% If confirmNum <> "" Then %>
                    <li> Response.confirmNum : <%=confirmNum%> </li>
                    <% End If %>
                    <% If tradeDate <> "" Then %>
                    <li> Response.tradeDate : <%=tradeDate%> </li>
                    <% End If %>
                </ul>
            </fieldset>
         </div>
    </body>
</html>