<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 수집 상태 확인(GetJobState API) 함수를 통해 상태 정보 확인된 작업아이디를 활용하여 현금영수증 매입/매출 내역을 조회합니다.
    ' - https://developers.popbill.com/reference/htcashbill/asp/api/search#Search
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = ""

    ' 수집 요청(requestJob) 시 반환받은 작업아이디(jobID)
    JobID = "018100815000000002"

    ' 문서형태 배열 ("N" 와 "C" 중 선택, 다중 선택 가능)
    ' └ N = 일반 현금영수증 , C = 취소현금영수증
    ' - 미입력 시 전체조회
    Dim TradeType(2)
    TradeType(0) = "N"
    TradeType(1) = "C"

    ' 거래구분 배열 ("P" 와 "C" 중 선택, 다중 선택 가능)
    ' └ P = 소득공제용 , C = 지출증빙용
    ' - 미입력 시 전체조회
    Dim TradeUsage(2)
    TradeUsage(0) = "P"
    TradeUsage(1) = "C"

    ' 페이지 번호
    Page  = 1

    ' 페이지당 목록개수
    PerPage = 10

    ' 정렬방항, D-내림차순, A-오름차순
    Order = "D"

    On Error Resume Next

    Set result = m_HTCashbillService.Search(testCorpNum, JobID, TradeType, TradeUsage, Page, PerPage, Order, UserID)

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
                        <legend>ListActiveJob [ <%=i+1%> / <%= UBound(result.list) %> ] </legend>
                            <ul>
                                <li> ntsconfirmNum (국세청승인번호) : <%= result.list(i).ntsconfirmNum %></li>
                                <li> tradeDate (거래일자) : <%= result.list(i).tradeDate %></li>
                                <li> tradeDT (거래일시) : <%= result.list(i).tradeDT %></li>
                                <li> tradeType (문서형태) : <%= result.list(i).tradeType %></li>
                                <li> tradeUsage (거래구분) : <%= result.list(i).tradeUsage %></li>
                                <li> totalAmount (거래금액) : <%= result.list(i).totalAmount %></li>
                                <li> supplyCost (공급가액) : <%= result.list(i).supplyCost %></li>
                                <li> tax (부가세) : <%= result.list(i).tax %></li>
                                <li> serviceFee (봉사료) : <%= result.list(i).serviceFee %></li>
                                <li> invoiceType (매입/매출) : <%= result.list(i).invoiceType %></li>
                                <li> franchiseCorpNum (발행자 사업자번호) : <%= result.list(i).franchiseCorpNum %></li>
                                <li> franchiseCorpName (발행자 상호) : <%= result.list(i).franchiseCorpName %></li>
                                <li> franchiseCorpType (발행자 사업자유형) : <%= result.list(i).franchiseCorpType %></li>
                                <li> identityNum (식별번호) : <%= result.list(i).identityNum %></li>
                                <li> identityNumType (식별번호유형) : <%= result.list(i).identityNumType %></li>
                                <li> customerName (고객명) : <%= result.list(i).customerName %></li>
                                <li> cardOwnerName (카드소유자명) : <%= result.list(i).cardOwnerName %></li>
                                <li> deductionType (공제유형) : <%= result.list(i).deductionType %></li>
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
