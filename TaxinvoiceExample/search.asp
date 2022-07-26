<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 검색조건을 사용하여 세금계산서 목록을 조회합니다. (조회기간 단위 : 최대 6개월)
    ' - https://docs.popbill.com/taxinvoice/asp/api#Search
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외 10자리
    testCorpNum = "1234567890"
    
    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
    KeyType = "SELL"

    ' 일자 유형 ("R" , "W" , "I" 중 택 1)
    ' └ R = 등록일자 , W = 작성일자 , I = 발행일자
    DType = "W"
    
    ' 시작일자, yyyyMMdd
    SDate = "20220701"

    ' 종료일자, yyyyMMdd
    EDate = "20220720"
    
    ' 상태코드 배열 (2,3번째 자리에 와일드카드(*) 사용 가능)
    ' - 미입력시 전체조회
    Dim State(2)
    State(0) = "3**"
    State(1) = "6**"

    
    ' 문서 유형 배열 ("N" , "M" 중 선택, 다중 선택 가능)
    ' - N = 일반 세금계산서 , M = 수정 세금계산서
    ' - 미입력시 전체조회
    Dim TIType(2)
    TIType(0) = "N"
    TIType(1) = "M"

    ' 과세형태 배열 ("T" , "N" , "Z" 중 선택, 다중 선택 가능)
    ' - T = 과세 , N = 면세 , Z = 영세
    ' - 미입력시 전체조회
    Dim TaxType(3)
    TaxType(0) = "T"
    TaxType(1) = "N"
    TaxType(2) = "Z"

    ' 발행형태 배열 ("N" , "R" , "T" 중 선택, 다중 선택 가능)
    ' - N = 정발행 , R = 역발행 , T = 위수탁발행
    ' - 미입력시 전체조회
    Dim IssueType(3)
    IssueType(0) = "N"
    IssueType(1) = "R"
    IssueType(2) = "T"

    ' 등록유형 배열 ("P" , "H" 중 선택, 다중 선택 가능)
    ' - P = 팝빌, H = 홈택스 또는 외부ASP
    ' - 미입력시 전체조회
    Dim RegType(2)
    RegType(0) = "P"
    RegType(1) = "H"

    ' 공급받는자 휴폐업상태 배열 ("N" , "0" , "1" , "2" , "3" , "4" 중 선택, 다중 선택 가능)
    ' - N = 미확인 , 0 = 미등록 , 1 = 사업 , 2 = 폐업 , 3 = 휴업 , 4 = 확인실패
    ' - 미입력시 전체조회
    Dim CloseDownState(5)
    CloseDownState(0) = "N"
    CloseDownState(1) = "0"
    CloseDownState(2) = "1"
    CloseDownState(3) = "2"
    CloseDownState(4) = "3"

    ' 지연발행 여부 (null , true , false 중 택 1)
    ' - null = 전체조회 , true = 지연발행 , false = 정상발행
    LateOnly = null		

    ' 정렬방향, A-오름차순, D-내림차순
    Order = "D"

    ' 페이지 번호
    Page = 1

    ' 페이지당 검색갯수, 최대 1000
    PerPage = 5

    ' 종사업장번호의 주체 ("S" , "B" , "T" 중 택 1)
    ' └ S = 공급자 , B = 공급받는자 , T = 수탁자
    ' - 미입력시 전체조회
    TaxRegIDType = "S"

    ' 종사업장번호 유무 (null , "0" , "1" 중 택 1)
    ' - null = 전체 , 0 = 없음, 1 = 있음
    TaxRegIDYN = ""
    
    ' 종사업장번호
    ' 다수기재시 콤마(",")로 구분하여 구성 ex ) "0001,0002"
    ' - 미입력시 전체조회
    TaxRegID = ""

    ' 거래처 상호 / 사업자번호 (사업자) / 주민등록번호 (개인) / "9999999999999" (외국인) 중 검색하고자 하는 정보 입력
    ' - 사업자번호 / 주민등록번호는 하이픈('-')을 제외한 숫자만 입력
    ' - 미입력시 전체조회
    QString = ""

    ' 세금계산서의 문서번호 / 국세청 승인번호 중 검색하고자 하는 정보 입력
    ' - 미입력시 전체조회
    MgtKey = ""

    ' 연동문서 여부 (null , "0" , "1" 중 택 1)
    ' - null = 전체조회 , 0 = 일반문서 , 1 = 연동문서
    ' - 일반문서 : 세금계산서 작성 시 API가 아닌 팝빌 사이트를 통해 등록한 문서
    ' - 연동문서 : 세금계산서 작성 시 API를 통해 등록한 문서
    InterOPYN = ""

    On Error Resume Next

    Set result = m_TaxinvoiceService.Search(testCorpNum, KeyType, DType, SDate, EDate, State, TIType, TaxType, _
                        IssueType, RegType, CloseDownState, LateOnly, Order, Page, PerPage, TaxRegIDType, TaxRegIDYN, _
                        TaxRegID, QString, MgtKey, InterOPYN, UsreID)

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
                <%
                    If code = 0 Then
                %>
                        <legend>세금계산서 목록조회</legend>
                        <ul>
                            <li> code (응답코드) : <%=result.code%></li>
                            <li> message (응답메시지) : <%=result.message%></li>
                            <li> total (총 검색결과 건수) : <%=result.total%></li>
                            <li> pageNum (페이지 번호) : <%=result.pageNum%></li>
                            <li> perPage (페이지당 목록개수) : <%=result.perPage%></li>
                            <li> pageCount (페이지 개수) : <%=result.pageCount%></li>
                        </ul>
                        <%
                            For i=0 To UBound(result.list) -1
                        %>
                            <fieldset class="fieldset2">					
                                <legend>  세금계산서 상태/요약정보 [ <%=i+1%> / <%=UBound(result.list)%> ]</legend>
                                    <ul>
                                        <li> itemKey (세금계산서 아이템키) :  <%=result.list(i).itemKey%> </li>
                                        <li> taxType (과세형태) :  <%=result.list(i).taxType%> </li>
                                        <li> writeDate (작성일자) :  <%=result.list(i).writeDate%> </li>
                                        <li> regDT (임시저장 일자) :  <%=result.list(i).regDT%> </li>
                                        <li> issueType (발행형태) :  <%=result.list(i).issueType %> </li>
                                        <li> supplyCostTotal (공급가액 합계) :  <%=result.list(i).supplyCostTotal%> </li>
                                        <li> taxTotal (세액 합계) :  <%=result.list(i).taxTotal%> </li>
                                        <li> purposeType (영수/청구) :  <%=result.list(i).purposeType%> </li>
                                        <li> issueDT (발행일시) :  <%=result.list(i).issueDT%> </li>
                                        <li> lateIssueYN (지연발행 여부) :  <%=result.list(i).lateIssueYN%> </li>
                                        <li> preIssueDT (발행예정일시) :  <%=result.list(i).preIssueDT%> </li>
                                        <li> openYN (개봉 여부) :  <%=result.list(i).openYN%> </li>
                                        <li> openDT (개봉일시) :  <%=result.list(i).openDT%> </li>
                                        <li> stateMemo (상태메모) :  <%=result.list(i).stateMemo%> </li>
                                        <li> stateCode (상태코드) :  <%=result.list(i).stateCode%> </li>
                                        <li> stateDT (상태 변경일시) :  <%=result.list(i).stateDT%> </li>
                                        <li> ntsconfirmNum (국세청 승인번호) :  <%=result.list(i).ntsconfirmNum %> </li>
                                        <li> ntsresult (국세청 전송결과) :  <%=result.list(i).ntsresult%> </li>
                                        <li> ntssendDT (국세청 전송일시) :  <%=result.list(i).ntssendDT%> </li>
                                        <li> ntsresultDT  (국세청 결과 수신일시) :  <%=result.list(i).ntsresultDT%> </li>
                                        <li> ntssendErrCode (전송실패 사유코드) :  <%=result.list(i).ntssendErrCode%> </li>
                                        <li> modifyCode (수정사유코드) : <%=result.list(i).modifyCode%></li> 
                                        <li> interOPYN (연동문서여부) :  <%=result.list(i).interOPYN%> </li>
                                        <li> invoicerCorpName (공급자 상호) :  <%=result.list(i).invoicerCorpName%> </li>
                                        <li> invoicerCorpNum (공급자 사업자번호) :  <%=result.list(i).invoicerCorpNum%> </li>
                                        <li> invoicerMgtKey (공급자 문서번호) :  <%=result.list(i).invoicerMgtKey%> </li>
                                        <li> invoicerPrintYN (공급자 인쇄여부) :  <%=result.list(i).invoicerPrintYN%> </li>
                                        <li> invoiceeCorpName (공급받는자 상호) :  <%=result.list(i).invoiceeCorpName%> </li>
                                        <li> invoiceeCorpNum (공급받는자 사업자번호) :  <%=result.list(i).invoiceeCorpNum%> </li>
                                        <li> invoiceeMgtKey (공급받는자 문서번호) :  <%=result.list(i).invoiceeMgtKey%> </li>
                                        <li> invoiceePrintYN (공급받는자 인쇄여부) :  <%=result.list(i).invoiceePrintYN%> </li>
                                        <li> closeDownState (공급받는자 휴폐업상태) :  <%=result.list(i).closeDownState%> </li>
                                        <li> closeDownStateDate (공급받는자 휴폐업일자) :  <%=result.list(i).closeDownStateDate%> </li>
                                    </ul>
                                </fieldset>
                <%
                        Next
                    Else
                %>
                </fieldset>
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
