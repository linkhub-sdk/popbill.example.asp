<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 수집 상태 확인(GetJobState API) 함수를 통해 상태 정보가 확인된 작업아이디를 활용하여 수집된 전자세금계산서 매입/매출 내역을 조회합니다.
    ' - https://docs.popbill.com/httaxinvoice/asp/api#Search
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"	

    ' 팝빌회원 아이디
    UserID = "testkorea"
    
    ' 수집 요청(requestJob) 시 반환받은 작업아이디(jobID)
    JobID = "019102415000000014"

    ' 문서형태 배열 ("N" 와 "M" 중 선택, 다중 선택 가능)
    ' └ N = 일반 , M = 수정
    ' - 미입력 시 전체조회
    Dim TIType(2) 
    TIType(0) = "N"
    TIType(1) = "M"

    ' 과세형태 배열 ("T" , "N" , "Z" 중 선택, 다중 선택 가능)
    ' └ T = 과세, N = 면세, Z = 영세
    ' - 미입력 시 전체조회
    Dim TaxType(3)
    TaxType(0) = "T"
    TaxType(1) = "N"
    TaxType(2) = "Z"
    
    ' 발행목적 배열 ("R" , "C", "N" 중 선택, 다중 선택 가능)
    ' └ R = 영수, C = 청구, N = 없음
    ' - 미입력 시 전체조회
    Dim PurposeType(3)
    PurposeType(0) = "R"
    PurposeType(1) = "C"
    PurposeType(2) = "N"

    ' 종사업장번호 유무 (null , "0" , "1" 중 택 1)
    ' - null = 전체 , 0 = 없음, 1 = 있음
    TaxRegIDYN = ""

    ' 종사업장번호의 주체 ("S" , "B" , "T" 중 택 1)
    ' └ S = 공급자 , B = 공급받는자 , T = 수탁자
    ' - 미입력시 전체조회
    TaxRegIDType = "S"

    ' 종사업장번호
    ' 다수기재시 콤마(",")로 구분하여 구성 ex ) "0001,0002"
    ' - 미입력시 전체조회
    TaxRegID = ""
    
    ' 페이지 번호 
    Page  = 1

    ' 페이지당 목록개수
    PerPage = 10

    ' 정렬방항, D-내림차순, A-오름차순
    Order = "D"

    ' 거래처 상호 / 사업자번호 (사업자) / 주민등록번호 (개인) / "9999999999999" (외국인) 중 검색하고자 하는 정보 입력
    ' - 사업자번호 / 주민등록번호는 하이픈('-')을 제외한 숫자만 입력
    ' - 미입력시 전체조회
    SearchString = ""

    On Error Resume Next

    Set result = m_HTTaxinvoiceService.Search(testCorpNum, JobID, TIType, TaxType, PurposeType, _	
                                TaxRegIDYN, TaxRegIDType, TaxRegID, Page, PerPage, Order, UserID, SearchString)

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
                <legend>수집 결과 조회</legend>
                <%
                    If code = 0 Then
                %>
                    <ul>
                        <li> code (응답코드) : <%=result.code%> </li>
                        <li> message  (응답메시지) : <%=result.message%> </li>
                        <li> total (총 검색결과 건수) : <%=result.total%> </li>
                        <li> perPage (페이지당 검색개수) : <%=result.perPage%> </li>
                        <li> pageNum (페이지 번호) : <%=result.pageNum%> </li>
                        <li> pageCount (페이지 개수) : <%=result.pageCount%> </li>
                    </ul>

                <%
                    For i=0 To UBound(result.list) -1 
                %>
                    <fieldset class="fieldset2">					
                        <legend>세금계산서 정보 [ <%=i+1%> / <%= UBound(result.list) %> ] </legend>
                            <ul>
                                            
                                <li> ntsconfirmNum (국세청승인번호) : <%= result.list(i).ntsconfirmNum %></li>
                                <li> writeDate (작성일자) : <%= result.list(i).writeDate %></li>
                                <li> issueDate (발행일자) : <%= result.list(i).issueDate %></li>
                                <li> sendDate (전송일자) : <%= result.list(i).sendDate %></li>
                                <li> taxType (과세형태) : <%= result.list(i).taxType %></li>
                                <li> invoiceType (매입/매출) : <%= result.list(i).invoiceType %></li>
                                <li> purposeType (영수/청구) : <%= result.list(i).purposeType %></li>
                                <li> supplyCostTotal (공급가액 합계) : <%= result.list(i).supplyCostTotal %></li>
                                <li> taxTotal (세액 합계) : <%= result.list(i).taxTotal %></li>
                                <li> totalAmount (합계금액) : <%= result.list(i).totalAmount %></li>
                                <li> remark1 (비고) : <%= result.list(i).remark1 %></li>						
                                <li> purchaseDate (거래일자) : <%= result.list(i).purchaseDate %></li>
                                <li> itemName (품명) : <%= result.list(i).itemName %></li>
                                <li> spec (규격) : <%= result.list(i).spec %></li>
                                <li> qty (수량) : <%= result.list(i).qty %></li>
                                <li> unitCost (단가) : <%= result.list(i).unitCost %></li>
                                <li> supplyCost (공급가액) : <%= result.list(i).supplyCost %></li>
                                <li> tax (세액) : <%= result.list(i).tax %></li>
                                <li> remark (비고) : <%= result.list(i).remark %></li>
                                <li> modifyYN (수정 전자세금계산서 여부 ) : <%= result.list(i).modifyYN %></li>
                                <li> orgNTSConfirmNum (원본 전자세금계산서 국세청승인번호) : <%= result.list(i).orgNTSConfirmNum %></li>
                                <br/>
                                <p><b>공급자 정보</b></p>
                                <li> invoicerCorpNum (사업자번호) : <%= result.list(i).invoicerCorpNum %></li>
                                <li> invoicerTaxRegID (종사업장번호) : <%= result.list(i).invoicerTaxRegID %></li>
                                <li> invoicerCorpName (상호) : <%= result.list(i).invoicerCorpName %></li>
                                <li> invoicerCEOName (대표자 성명) : <%= result.list(i).invoicerCEOName %></li>
                                <li> invoicerEmail (담당자 이메일) : <%= result.list(i).invoicerEmail %></li>
                                <br/>
                                <p><b>공급받는자 정보</b></p>
                                <li> invoiceeCorpNum (사업자번호) : <%= result.list(i).invoiceeCorpNum %></li>
                                <li> invoiceeType (공급받는자 구분) : <%= result.list(i).invoiceeType %></li>
                                <li> invoiceeTaxRegID (종사업장번호) : <%= result.list(i).invoiceeTaxRegID %></li>
                                <li> invoiceeCorpName (상호) : <%= result.list(i).invoiceeCorpName %></li>
                                <li> invoiceeCEOName (대표자 성명) : <%= result.list(i).invoiceeCEOName %></li>
                                <li> invoiceeEmail1 (담당자 이메일) : <%= result.list(i).invoiceeEmail1 %></li>
                                <br/>
                                <p><b>수탁자 정보</b></p>
                                <li> trusteeCorpNum (사업자번호) : <%= result.list(i).trusteeCorpNum %></li>
                                <li> trusteeTaxRegID (종사업장번호) : <%= result.list(i).trusteeTaxRegID %></li>
                                <li> trusteeCorpName (상호) : <%= result.list(i).trusteeCorpName %></li>
                                <li> trusteeCEOName (대표자 성명) : <%= result.list(i).trusteeCEOName %></li>
                                <li> trusteeEmail (담당자 이메일) : <%= result.list(i).trusteeEmail %></li>

                            </ul>
                        </fieldset>
                <%
                        Next					
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

