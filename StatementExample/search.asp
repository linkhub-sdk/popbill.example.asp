<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 검색조건을 사용하여 전자명세서 목록을 조회합니다. (조회기간 단위 : 최대 6개월)
    ' - https://developers.popbill.com/reference/statement/asp/api/info#Search
    '**************************************************************

    ' 팝빌회원 사업자번호
    testCorpNum = "1234567890"

    ' 검색일자 유형 ("R" , "W" , "I" 중 택 1)
    ' - R = 등록일자 , W = 작성일자 , I = 발행일자
    DType = "W"

    ' 시작일자, yyyyMMdd
    SDate = "20220701"

    ' 종료일자, yyyyMMdd
    EDate = "20220720"

    ' 전자명세서 상태코드 배열 (2,3번째 자리에 와일드카드(*) 사용 가능)
    ' - 미입력시 전체조회
    Dim State(2)
    State(0) = "2**"
    State(1) = "3**"

    ' 전자명세서 문서유형 배열 (121 , 122 , 123 , 124 , 125 , 126 중 선택. 다중 선택 가능)
    ' 121 = 명세서 , 122 = 청구서 , 123 = 견적서
    ' 124 = 발주서 , 125 = 입금표 , 126 = 영수증
    Dim ItemCode(6)
    ItemCode(0) = "121"
    ItemCode(1) = "122"
    ItemCode(2) = "123"
    ItemCode(3) = "124"
    ItemCode(4) = "125"
    ItemCode(5) = "126"

    ' 정렬방향, A-오름차순, D-내림차순
    Order = "D"

    ' 페이지 번호
    Page = 1

    ' 페이지당 검색개수
    PerPage = 20

    ' 통합검색어, 거래처 상호명 또는 거래처 사업자번호로 조회
    ' - 미입력시 전체조회
    SQuery = ""

    On Error Resume Next

    Set result = m_StatementService.Search(testCorpNum, DType, SDate, EDate, State, ItemCode, Order, Page, PerPage, SQuery)

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
                <legend>전자명세서 목록조회</legend>
                    <% If code = 0 Then %>
                    <ul>
                        <li> code(응답 코드) : <%=result.code%></li>
                        <li> message(응답 메시지) : <%=result.message%></li>
                        <li> total(총 검색결과 건수) : <%=result.total%></li>
                        <li> perPage(페이지당 검색개수) : <%=result.perPage%></li>
                        <li> pageNum(페이지 번호) : <%=result.pageNum%></li>
                        <li> pageCount(페이지 개수) : <%=result.pageCount%></li>
                    </ul>

                    <% For i=0 To UBound(result.list)-1 %>

                        <fieldset class="fieldset2">
                            <legend> 전자명세서 조회결과 [ <%=i+1%> / <%=UBound(result.list)%> ] </legend>
                            <ul>
                                <li> itemKey(아이템키) : <%=result.list(i).itemKey%></li>
                                <li> itemCode(문서종류코드) : <%=result.list(i).itemCode %></li>
                                <li> itemKey(팝빌번호) : <%=result.list(i).itemKey %></li>
                                <li> invoiceNum(팝빌승인번호) : <%=result.list(i).invoiceNum %></li>
                                <li> mgtKey(파트너 문서번호) : <%=result.list(i).mgtKey %></li>
                                <li> taxType(세금형태) : <%=result.list(i).taxType %></li>
                                <li> writeDate(작성일자) : <%=result.list(i).writeDate %></li>
                                <li> regDT(등록일시) : <%=result.list(i).regDT %></li>
                                <li> senderCorpName(발신자 상호) : <%=result.list(i).senderCorpName %></li>
                                <li> senderCorpNum(발신자 사업자번호) : <%=result.list(i).senderCorpNum %></li>
                                <li> senderPrintYN(발신자 인쇄여부) : <%=result.list(i).senderPrintYN %></li>
                                <li> receiverCorpName(수신자 상호) : <%=result.list(i).receiverCorpName %></li>
                                <li> receiverCorpNum(수신자 사업자번호) : <%=result.list(i).receiverCorpNum %></li>
                                <li> receiverPrintYN(수신자 인쇄여부) : <%=result.list(i).receiverPrintYN %></li>
                                <li> supplyCostTotal(공급가액 합계) : <%=result.list(i).supplyCostTotal %></li>
                                <li> taxTotal(세액 합계) : <%=result.list(i).taxTotal %></li>
                                <li> purposeType(영수/청구) : <%=result.list(i).purposeType %></li>
                                <li> issueDT(발행일시) : <%=result.list(i).issueDT %></li>
                                <li> stateCode(상태코드) : <%=result.list(i).stateCode %></li>
                                <li> stateDT(상태 변경일시) : <%=result.list(i).stateDT %></li>
                                <li> stateMemo(상태메모) : <%=result.list(i).stateMemo %></li>
                                <li> openYN(메일 개봉 여부) : <%=result.list(i).openYN %></li>
                                <li> openDT(개봉 일시) : <%=result.list(i).openDT %></li>
                            </ul>
                        </fieldset>
                    <%
                        Next
                        Else
                    %>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    <% End If %>
            </fieldset>
        </div>
    </body>
</html>