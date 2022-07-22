<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 최대 100건의 세금계산서 발행을 한번의 요청으로 접수합니다.
    ' - 세금계산서 발행을 위해서 공급자의 인증서가 팝빌 인증서버에 사전등록 되어야 합니다.
    '   └ 위수탁발행의 경우, 수탁자의 인증서 등록이 필요합니다.
    ' - https://docs.popbill.com/taxinvoice/asp/api#BulkSubmit
    '**************************************************************
    
    ' 팝빌회원 사업자번호
    testCorpNum = "1234567890"

    ' 제출아이디, 최대 36자리 (영문, 숫자, "-" 조합)
    SubmitID = "ASP-BULK-TEST102"

    ' 지연발행 강제여부  (true / false 중 택 1)
    ' └ true = 가능 , false = 불가능
    ' - 발행마감일이 지난 세금계산서를 발행하는 경우, 가산세가 부과될 수 있습니다.
    ' - 가산세가 부과되더라도 발행을 해야하는 경우에는 forceIssue의 값을
    '   true로 선언하여 발행(Issue API)를 호출하시면 됩니다.
    forceIssue = False

    ' 팝빌회원 아이디
    userID = "testkorea"
   
    
    Dim taxinvoiceList(99)  
    for i = 0 to 99
        ' 세금계산서 정보 객체 생성
        Set newTaxinvoice = New Taxinvoice

        ' 작성일자, 날짜형식(yyyyMMdd)
        newTaxinvoice.writeDate = "20210601"

        ' {정과금, 역과금} 중 기재, '역과금'은 역발행 프로세스에서만 이용가능
        newTaxinvoice.chargeDirection = "정과금"
        
        ' 발행형태, {정발행, 역발행, 위수탁} 중 기재
        newTaxinvoice.issueType = "정발행"

        ' {영수, 청구, 없음} 중 기재 
        newTaxinvoice.purposeType = "영수"
        
        ' 과세형태, {과세, 영세, 면세} 중 기재
        newTaxinvoice.taxType = "과세"
        

        '**************************************************************
        '                       공급자 정보
        '**************************************************************

        ' 공급자 사업자번호, '-' 제외 10자리
        newTaxinvoice.invoicerCorpNum = "1234567890"

        ' 공급자 종사업자 식별번호. 필요시 숫자 4자리 기재
        newTaxinvoice.invoicerTaxRegID = ""

        ' 공급자 상호
        newTaxinvoice.invoicerCorpName = "공급자 상호" + CStr(i)

        ' 공급자 문서번호, 1~24자리 (숫자, 영문, '-', '_') 조합으로
        ' 사업자 별로 중복되지 않도록 구성
        newTaxinvoice.invoicerMgtKey = "20210601-ASP-Bulk0" + CStr(i)

        ' 공급자 대표자 성명
        newTaxinvoice.invoicerCEOName = "공급자 대표자 성명"
        
        ' 공급자 주소
        newTaxinvoice.invoicerAddr = "공급자 주소"
        
        ' 공급자 종목
        newTaxinvoice.invoicerBizClass = "공급자 종목"
        
        ' 공급자 업태
        newTaxinvoice.invoicerBizType = "공급자 업태,업태2"
        
        ' 공급자 담당자명
        newTaxinvoice.invoicerContactName = "공급자 담당자명"
        
        ' 공급자 담당자 메일주소 
        newTaxinvoice.invoicerEmail = ""
        
        ' 공급자 담당자 연락처 
        newTaxinvoice.invoicerTEL = ""
        
        ' 공급자 담당자 휴대폰번호
        newTaxinvoice.invoicerHP = ""

        ' 발행 안내 문자 전송여부 (true / false 중 택 1)
        ' └ true = 전송 , false = 미전송
        ' └ 공급받는자 (주)담당자 휴대폰번호 {invoiceeHP1} 값으로 문자 전송
        ' - 전송 시 포인트 차감되며, 전송실패시 환불처리
        newTaxinvoice.invoicerSMSSendYN = False
        


        '**************************************************************
        '                     공급받는자 정보
        '**************************************************************

        ' 공급받는자 구분, [사업자, 개인, 외국인] 중 기재
        newTaxinvoice.invoiceeType = "사업자"

        ' 공급받는자 사업자번호
        ' - {invoiceeType}이 "사업자" 인 경우, 사업자번호 (하이픈 ('-') 제외 10자리)
        ' - {invoiceeType}이 "개인" 인 경우, 주민등록번호 (하이픈 ('-') 제외 13자리)
        ' - {invoiceeType}이 "외국인" 인 경우, "9999999999999" (하이픈 ('-') 제외 13자리)
        newTaxinvoice.invoiceeCorpNum = "8888888888"

        ' 공급받는자 종사업자 식별번호. 필요시 숫자 4자리 기재	
        newTaxinvoice.invoiceeTaxRegID = ""
        
        ' 공급자받는자 상호
        newTaxinvoice.invoiceeCorpName = "공급받는자 상호" + CStr(i)

        ' [역발행시 필수] 공급받는자 문서번호(역발행시 필수)
        newTaxinvoice.invoiceeMgtKey = ""

        ' 공급받는자 대표자 성명
        newTaxinvoice.invoiceeCEOName = "공급받는자 대표자 성명"
        
        ' 공급받는자 주소
        newTaxinvoice.invoiceeAddr = "공급받는자 주소"
        
        ' 공급받는자 종목
        newTaxinvoice.invoiceeBizClass = "공급받는자 종목"
        
        ' 공급받는자 업태
        newTaxinvoice.invoiceeBizType = "공급받는자 업태"
        
        ' 공급받는자 담당자명
        newTaxinvoice.invoiceeContactName1 = "공급받는자 담당자명"
        
        ' 공급받는자 담당자 메일주소
        ' 팝빌 개발환경에서 테스트하는 경우에도 안내 메일이 전송되므로,
        ' 실제 거래처의 메일주소가 기재되지 않도록 주의
        newTaxinvoice.invoiceeEmail1 = ""
        
        ' 공급받는자 연락처
        newTaxinvoice.invoiceeTEL1 = ""
        
        ' 공급받는자 휴대폰번호
        newTaxinvoice.invoiceeHP1 = ""

        ' 역발행시 공급자에게 발행안내문자 전송여부
        newTaxinvoice.invoiceeSMSSendYN = False


        '**************************************************************
        '                       세금계산서 정보
        '**************************************************************

        ' 공급가액 합계
        newTaxinvoice.supplyCostTotal = "100000"

        ' 세액 합계
        newTaxinvoice.taxTotal = "10000"

        ' 합계금액, 공급가액 합계 + 세액합계
        newTaxinvoice.totalAmount = "110000"
        
        ' 기재 상 '일련번호' 항목
        newTaxinvoice.serialNum = "123"

        ' 기재 상 '권' 항목, 최대값 32767
        newTaxinvoice.kwon = "1"

        ' 기재 상 '호' 항목, 최대값 32767
        newTaxinvoice.ho = "1"

        ' 기재 상 '현금' 항목
        newTaxinvoice.cash = ""
        
        ' 기재 상 '수표' 항목
        newTaxinvoice.chkBill = ""

        ' 기재 상 '어음' 항목
        newTaxinvoice.note = ""
        
        ' 기재 상 '외상미수금' 항목
        newTaxinvoice.credit = ""

        ' 비고
        ' {invoiceeType}이 "외국인" 이면 remark1 필수
        ' - 외국인 등록번호 또는 여권번호 입력
        newTaxinvoice.remark1 = "비고1"
        newTaxinvoice.remark2 = "비고2"
        newTaxinvoice.remark3 = "비고3"

        ' 사업자등록증 이미지 첨부여부  (true / false 중 택 1)
        ' └ true = 첨부 , false = 미첨부(기본값)
        ' - 팝빌 사이트 또는 인감 및 첨부문서 등록 팝업 URL (GetSealURL API) 함수를 이용하여 등록
        newTaxinvoice.businessLicenseYN = False 

        ' 통장사본 이미지 첨부여부  (true / false 중 택 1)
        ' └ true = 첨부 , false = 미첨부(기본값)
        ' - 팝빌 사이트 또는 인감 및 첨부문서 등록 팝업 URL (GetSealURL API) 함수를 이용하여 등록
        newTaxinvoice.bankBookYN = False
    
        
        
        '**************************************************************
        '         수정세금계산서 정보 (수정세금계산서 작성시에만 기재
        ' - 수정세금계산서 관련 정보는 연동매뉴얼 또는 개발가이드 링크 참조
        ' - [참고] 수정세금계산서 작성방법 안내 - https://docs.popbill.com/taxinvoice/modify?lang=asp
        '**************************************************************

        ' [수정세금계산서 발행시 필수] 수정사유코드, 수정사유에 따라 1~6중 선택기재
        newTaxinvoice.modifyCode = ""

        ' [수정세금계산서 발행시 필수] 원본세금계산서의 국세청 승인번호 기재
        newTaxinvoice.orgNTSConfirmNum = ""


        '**************************************************************
        '                         상세항목(품목) 정보
        '**************************************************************
        Set newDetail = New TaxinvoiceDetail
        newDetail.serialNum = 1             '일련번호 1부터 순차 기재
        newDetail.purchaseDT = "20210601"   '거래일자  yyyyMMdd
        newDetail.itemName = "품명1번"
        newDetail.spec = "규격"
        newDetail.qty = "1" '수량           ' 소숫점 2자리까지 문자열로 기재가능
        newDetail.unitCost = "50000"       ' 소숫점 2자리까지 문자열로 기재가능
        newDetail.supplyCost = "50000"
        newDetail.tax = "5000"
        newDetail.remark = "비고"

        newTaxinvoice.AddDetail newDetail

        Set newDetail = New TaxinvoiceDetail
        newDetail.serialNum = 2             '일련번호 1부터 순차 기재
        newDetail.purchaseDT = "20210601"   '거래일자  yyyyMMdd
        newDetail.itemName = "품명2번"
        newDetail.spec = "규격"
        newDetail.qty = "1" '수량           ' 소숫점 2자리까지 문자열로 기재가능
        newDetail.unitCost = "50000"       ' 소숫점 2자리까지 문자열로 기재가능
        newDetail.supplyCost = "50000"
        newDetail.tax = "5000"
        newDetail.remark = "비고"
        
        newTaxinvoice.AddDetail newDetail
    


        '**************************************************************
        '                            추가담당자 정보
        ' - 세금계산서 발행안내 메일을 수신받을 공급받는자 담당자가 다수인 경우
        '   담당자 정보를 추가하여 발행안내메일을 다수에게 전송할 수 있습니다.
        '**************************************************************
        set newContact = New Contact
        newContact.serialNum = 1
        newContact.contactName = "담당자1 성명"
        newContact.email = ""   
        newTaxinvoice.AddContact newContact
        
        set newContact = New Contact
        newContact.serialNum = 2
        newContact.contactName = "담당자2 성명"
        newContact.email = ""
        newTaxinvoice.AddContact newContact
        
        Set taxinvoiceList(i) =  newTaxinvoice
    Next

    On Error Resume Next
    
    Set Presponse = m_TaxinvoiceService.BulkSubmit(testCorpNum, SubmitID, taxinvoiceList, forceIssue, userID)
    
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
                <legend>세금계산서 초대량 접수</legend>
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