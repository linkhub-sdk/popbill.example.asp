<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 검색조건을 사용하여 팩스전송 내역을 조회합니다. (조회기간 단위 : 최대 2개월)
    ' - 팩스 접수일시로부터 2개월 이내 접수건만 조회할 수 있습니다.
    ' - https://developers.popbill.com/reference/fax/asp/api/info#Search
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    '시작일자, yyyyMMdd
    SDate = "20220701"

    '종료일자, yyyyMMdd
    EDate = "20220720"

    ' 전송상태 배열 ("1" , "2" , "3" , "4" 중 선택, 다중 선택 가능)
    ' └ 1 = 대기 , 2 = 성공 , 3 = 실패 , 4 = 취소
    ' - 미입력 시 전체조회
    Dim State(4)
    State(0) = "1"
    State(1) = "2"
    State(2) = "3"
    State(3) = "4"

    ' 예약여부 (false , true 중 택 1)
    ' false = 전체조회, true = 예약전송건 조회
    ' 미입력시 기본값 false 처리
    ReserveYN = False

    ' 개인조회 여부 (false , true 중 택 1)
    ' false = 접수한 팩스 전체 조회 (관리자권한)
    ' true = 해당 담당자 계정으로 접수한 팩스만 조회 (개인권한)
    ' 미입력시 기본값 false 처리
    SenderOnlyYN = False

    '정렬발향, A-오름차순, D-내림차순
    Order = "D"

    '페이지 번호
    Page = 1

    '페이지당 검색개수
    PerPage = 20

    ' 조회하고자 하는 발신자명 또는 수신자명
    ' - 미입력시 전체조회
    QString = ""

    On Error Resume Next

    Set result = m_FaxService.Search(CorpNum, SDate, EDate, State, ReserveYN, SenderOnlyYN, Order, Page, PerPage, QString)

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
                <legend>팩스전송 전송내역 조회 </legend>
                    <ul>
                        <li> code (응답코드) : <%=result.code%></li>
                        <li> total (총 검색결과 건수) : <%=result.total%></li>
                        <li> pageNum (페이지 번호) : <%=result.pageNum%></li>
                        <li> perPage (페이지당 목록개수) : <%=result.perPage%></li>
                        <li> pageCount (페이지 개수) : <%=result.pageCount%></li>
                        <li> message (응답메시지) : <%=result.message%></li>
                    </ul>
                <% If code = 0 Then
                        For i=0 To UBound(result.list)-1
                %>
                    <fieldset class="fieldset2">
                            <legend> 팩스 전송결과 [ <%=i+1%> /  <%=UBound(result.list)%> ] </legend>
                            <ul>
                                <li>state (전송상태 코드) : <%=result.list(i).state%> </li>
                                <li>result (전송결과 코드) : <%=result.list(i).result%> </li>
                                <li>sendNum (발신번호) : <%=result.list(i).sendNum%> </li>
                                <li>senderName (발신자명) : <%=result.list(i).senderName%> </li>
                                <li>ReceiveNum (수신번호) : <%=result.list(i).ReceiveNum%> </li>
                                <li>ReceiveNumType (수신번호 유형) : <%=result.list(i).ReceiveNumType%> </li>
                                <li>receiveName (수신자명) : <%=result.list(i).receiveName%> </li>
                                <li>title (팩스 제목) : <%=result.list(i).title %> </li>
                                <li>sendPageCnt (페이지수) : <%=result.list(i).sendPageCnt%></li>
                                <li>successPageCnt (성공 페이지수) : <%=result.list(i).successPageCnt%></li>
                                <li>failPageCnt (실패 페이지수) : <%=result.list(i).failPageCnt%></li>
                                <li>refundPageCnt (환불 페이지수) : <%=result.list(i).refundPageCnt%></li>
                                <li>cancelPageCnt (취소 페이지수) : <%=result.list(i).cancelPageCnt%></li>
                                <li>reserveDT (예약시간) : <%=result.list(i).reserveDT%></li>
                                <li>sendDT (발송시간) : <%=result.list(i).sendDT%></li>
                                <li>receiptDT (전송 접수시간) : <%=result.list(i).receiptDT%></li>
                                <li>fileNames (전송파일명 배열) : <%=result.list(i).fileNames%></li>
                                <li>interOPRefKey (파트너 지정키) : <%=result.list(i).interOPRefKey%> </li>
                                <li>ReceiptNum (접수번호) : <%=result.list(i).ReceiptNum%> </li>
                                <li>RequestNum (요청번호) : <%=result.list(i).RequestNum%> </li>
                                <li>chargePageCnt (과금 페이지수) : <%=result.list(i).chargePageCnt%> </li>
                                <li>tiffFileSize (변환파일용량 (단위 : byte)) : <%=result.list(i).tiffFileSize%> </li>
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
                <%	End If	%>

            </fieldset>
        </div>
    </body>
</html>