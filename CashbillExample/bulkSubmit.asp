<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 최대 100건의 현금영수증 발행을 한번의 요청으로 접수합니다.
    ' - https://developers.popbill.com/reference/cashbill/asp/api/issue#BulkSubmit
    '**************************************************************

    ' 팝빌회원 사업자번호
    CorpNum = "1234567890"

    ' 제출아이디, 최대 36자리 (영문, 숫자, "-" 조합)
    SubmitID = "20220720-ASP-BULK001"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    Dim cashbillList(99)
    for i = 0 to 99
        ' 현금영수증 정보 객체 생성
        Set CashbillObj = New Cashbill

        CashbillObj.mgtKey = SubmitID + CStr(i)

        ' 문서형태, [승인거래, 취소거래] 중 기재
        CashbillObj.tradeType = "승인거래"

        ' [취소거래시 필수] 원본 현금영수증 국세청승인번호
        CashbillObj.orgConfirmNum = ""

        ' [취소거래시 필수] 원본 현금영수증 거래일자
        CashbillObj.orgTradeDate = ""

        ' 거래구분, [소득공제용, 지출증빙용] 중 기재
        CashbillObj.tradeUsage = "소득공제용"

        ' 거래유형, [일반, 도서공연, 대중교통] 중 기재
        ' 미입력시 기본값 '일반' 처리
        CashbillObj.tradeOpt = "일반"

        ' 과세형태, [과세, 비과세] 중 기재
        CashbillObj.taxationType = "과세"

        ' 공급가액
        CashbillObj.supplyCost = "10000"

        ' 부가세
        CashbillObj.tax = "1000"

        ' 봉사료
        CashbillObj.serviceFee = "0"

        ' 합계금액, 공급가액 + 봉사료 + 세액
        CashbillObj.totalAmount = "11000"

        ' 가맹점 사업자번호, "-" 제외 10자리
        CashbillObj.franchiseCorpNum = CorpNum

        ' 가맹점 종사업장 식별번호
        CashbillObj.franchiseTaxRegID = ""

        ' 가맹점 상호
        CashbillObj.franchiseCorpName = "가맹점 상호"

        ' 가맹점 대표자 성명
        CashbillObj.franchiseCEOName = "가맹점 대표자"

        ' 가맹점 주소
        CashbillObj.franchiseAddr = "가맹점 주소"

        ' 가맹점 전화번호
        CashbillObj.franchiseTEL = "070-1234-1234"

        ' 식별번호, 거래구분에 따라 작성
        ' └ 소득공제용 - 주민등록/휴대폰/카드번호(현금영수증 카드)/자진발급용 번호(010-000-1234) 기재가능
        ' └ 지출증빙용 - 사업자번호/주민등록/휴대폰/카드번호(현금영수증 카드) 기재가능
        ' └ 주민등록번호 13자리, 휴대폰번호 10~11자리, 카드번호 13~19자리, 사업자번호 10자리 입력 가능
        CashbillObj.identityNum = "0100001234"

        ' 주문고객명
        CashbillObj.customerName = "고객명"

        ' 주문상품명
        CashbillObj.itemName = "상품명"

        ' 주문번호
        CashbillObj.orderNumber = "주문번호"

        ' 이메일
        ' 팝빌 개발환경에서 테스트하는 경우에도 안내 메일이 전송되므로,
        ' 실제 거래처의 메일주소가 기재되지 않도록 주의
        CashbillObj.email = ""

        ' 휴대폰
        CashbillObj.hp = ""

        ' 발행안내문자 전송여부
        ' 안내문자 전송시 포인트가 차감되며, 전송실패시 환불처리됩니다.
        CashbillObj.smssendYN = False

        ' 거래일시, 날짜(yyyyMMddHHmmss)
        ' 당일, 전일만 가능, 미입력시 기본값 발행일시 처리
        CashbillObj.tradeDT = "20221108000000"

        Set cashbillList(i) =  CashbillObj
    Next

    On Error Resume Next

    Set Presponse = m_CashbillService.BulkSubmit(CorpNum, SubmitID, cashbillList, UserID)

    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        receiptID = ""
        Err.Clears
    Else
        code = Presponse.code
        message =Presponse.message
        receiptID = Presponse.receiptID
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>현금영수증 초대량 접수</legend>
                <ul>
                    <li>응답코드 (Response.code) : <%=code%> </li>
                    <li>응답메시지 (Response.message) : <%=message%> </li>
                    <% If receiptID <> "" Then %>
                    <li>접수아이디 (Response.receiptID) : <%=receiptID%> </li>
                    <% End If %>
                </ul>
            </fieldset>
        </div>
    </body>
</html>
