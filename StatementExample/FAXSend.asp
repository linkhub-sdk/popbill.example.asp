<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 전자명세서를 팩스로 전송하는 함수로, 팝빌에 데이터를 저장하는 과정이 없습니다.
    ' - 팝빌 사이트 [문자·팩스] > [팩스] > [전송내역] 메뉴에서 전송결과를 확인 할 수 있습니다.
    ' - 함수 호출시 포인트가 과금됩니다.
    ' - 선팩스 전송 요청 시 작성한 문서번호는 팩스로 전송되는 파일명에 사용됩니다.
    ' - 팩스 전송결과를 확인하기 위해서는 선팩스 전송 요청 시 반환받은 접수번호를 이용하여 팩스 API의 전송내역 확인 (GetFaxResult API) 함수를 이용하면 됩니다.
    ' - https://developers.popbill.com/reference/statement/asp/api/etc#FAXSend
    '**************************************************************

    ' 팝빌회원 사업자번호, "-"제외 10자리
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 발신번호
    sendNum = "07043042991"

    ' 수신팩스번호
    ReceiveNum = "010111222"

    ' 문서번호
    mgtKey = "20220720-ASP-002"


    ' 전자명세서 객체 생성
    Set newStatement = New Statement

    ' 기재상 작성일자, 날짜형식(yyyyMMdd)
    newStatement.writeDate = "20220720"

    ' {영수, 청구, 없음} 중 기재
    newStatement.purposeType = "영수"

    ' 과세형태, {과세, 영세, 면세} 중 기재
    newStatement.taxType = "과세"

    ' 맞춤양식코드, 공백처리시 기본양식으로 작성
    newStatement.formCode = ""

    ' 명세서 종류코드 - 121(거래명세서), 122(청구서), 123(견적서) 124(발주서), 125(입금표), 126(영수증)
    newStatement.itemCode = "121"

    ' 문서번호, 숫자, 영문, '-', '_' 조합 (최대24자리)으로 사업자별로 중복되지 않도록 구성
    newStatement.mgtKey = mgtKey



    '**************************************************************
    '                        발신자 정보
    '**************************************************************

    ' 발신자 사업자번호, '-' 제외 10자리
    newStatement.senderCorpNum = CorpNum

    ' 발신자 종사업장 식별번호, 필요시 기재, 형식은 숫자 4자리
    newStatement.senderTaxRegID = ""

    ' 발신자 상호
    newStatement.senderCorpName = "발신자 상호"

    ' 발신자 대표자성명
    newStatement.senderCEOName = "발신자"" 대표자 성명"

    ' 발신자 주소
    newStatement.senderAddr = "발신자 주소"

    ' 발신자 종목
    newStatement.senderBizClass = "발신자 종목"

    ' 발신자 업태
    newStatement.senderBizType = "발신자 업태,업태2"

    ' 발신자 담당자 성명
    newStatement.senderContactName = "발신자 담당자명"

    ' 발신자 메일주소
    newStatement.senderEmail = ""

    ' 발신자 연락처
    newStatement.senderTEL = ""

    ' 발신자 휴대폰번호
    newStatement.senderHP = ""



    '**************************************************************
    '                      수신자 정보
    '**************************************************************

    ' 수신자 사업자번호, '-' 제외 10자리
    newStatement.receiverCorpNum = "8888888888"

    ' 수신자 상호
    newStatement.receiverCorpName = "수신자 상호"

    ' 수신자 대표자 성명
    newStatement.receiverCEOName = "수신자 대표자 성명"

    ' 수신자 주소
    newStatement.receiverAddr = "수신자 주소"

    ' 수신자 종목
    newStatement.receiverBizClass = "수신자 종목"

    ' 수신자 업태
    newStatement.receiverBizType = "수신자 업태"

    ' 수신자 담당자명
    newStatement.receiverContactName = "수신자 담당자명"

    ' 수신자 메일주소
    newStatement.receiverEmail = ""

    ' 수신자 연락처
    newStatement.receiverTEL = ""

    ' 수신자 휴대폰번호
    newStatement.receiverHP = ""



    '**************************************************************
    '                      전자명세서 기재사항
    '**************************************************************

    ' 공급가액 합계
    newStatement.supplyCostTotal = "100000"

    ' 세액 합계
    newStatement.taxTotal = "10000"

    ' 합계금액, 공급가액 합계 + 세액 합계
    newStatement.totalAmount = "110000"

    ' 기재 상 일련번호 항목
    newStatement.serialNum = "123"

    ' 기재 상 비고 항목
    newStatement.remark1 = "비고1"
    newStatement.remark2 = "비고2"
    newStatement.remark3 = "비고3"

    ' 사업자등록증 이미지 첨부여부  (true / false 중 택 1)
    ' └ true = 첨부 , false = 미첨부(기본값)
    ' - 팝빌 사이트 또는 인감 및 첨부문서 등록 팝업 URL (GetSealURL API) 함수를 이용하여 등록
    newStatement.businessLicenseYN = False

    ' 통장사본 이미지 첨부여부  (true / false 중 택 1)
    ' └ true = 첨부 , false = 미첨부(기본값)
    ' - 팝빌 사이트 또는 인감 및 첨부문서 등록 팝업 URL (GetSealURL API) 함수를 이용하여 등록
    newStatement.bankBookYN = False

    ' 발행시 알림문자 전송여부
    newStatement.smssendYN = True

    '**************************************************************
    '                      전자명세서 상세(품목)
    '**************************************************************

    Set newDetail = New StatementDetail

    newDetail.serialNum = "1"             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20220720"   '거래일자  yyyyMMdd
    newDetail.itemName = "품명"
    newDetail.spec = "규격"
    newDetail.unit = "단위"
    newDetail.qty = "1" '수량           ' 소숫점 2자리까지 문자열로 기재가능
    newDetail.unitCost = "100000"       ' 소숫점 2자리까지 문자열로 기재가능
    newDetail.supplyCost = "100000"
    newDetail.tax = "10000"
    newDetail.remark = "비고"
    newDetail.spare1 = "spare1"
    newDetail.spare2 = "spare2"
    newDetail.spare3 = "spare3"
    newDetail.spare4 = "spare4"
    newDetail.spare5 = "spare5"

    newStatement.AddDetail newDetail

    Set newDetail = New StatementDetail

    newDetail.serialNum = "2"             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20220720"   '거래일자  yyyyMMdd
    newDetail.itemName = "품명"
    newDetail.spec = "규격"
    newDetail.unit = "단위"
    newDetail.qty = "1" '수량           ' 소숫점 2자리까지 문자열로 기재가능
    newDetail.unitCost = "100000"       ' 소숫점 2자리까지 문자열로 기재가능
    newDetail.supplyCost = "100000"
    newDetail.tax = "10000"
    newDetail.remark = "비고"
    newDetail.spare1 = "spare1"
    newDetail.spare2 = "spare2"
    newDetail.spare3 = "spare3"
    newDetail.spare4 = "spare4"
    newDetail.spare5 = "spare5"

    newStatement.AddDetail newDetail


    '**************************************************************
    '                       전자명세서 추가속성
    '**************************************************************

    newStatement.propertyBag.Set "Balance", "150000"
    newStatement.propertyBag.Set "CBalance", "100000"


    On Error Resume Next

    Presponse = m_StatementService.FAXSend(CorpNum, newStatement, sendNum, ReceiveNum, UserID)

    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else
        code = Presponse.code
        message = Presponse.message
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>선팩스 전송</legend>
                <ul>
                    <% If code = 0 Then %>
                        <li>recepitNum : <%=Presponse%> </li>
                    <% Else %>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    <% End If %>
                </ul>
            </fieldset>
        </div>
    </body>
</html>