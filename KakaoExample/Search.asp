
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 검색조건을 사용하여 카카오톡 전송내역 목록을 조회합니다. (조회기간 단위 : 최대 2개월)
    ' - 카카오톡 접수일시로부터 6개월 이내 접수건만 조회할 수 있습니다.
    ' - https://developers.popbill.com/reference/kakaotalk/asp/api/info#Search
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    '시작일자
    SDate = "20220701"

    '종료일자
    EDate = "20220720"

    ' 전송상태 배열 ("0" , "1" , "2" , "3" , "4" , "5" 중 선택, 다중 선택 가능)
    ' └ 0 = 전송대기 , 1 = 전송중 , 2 = 전송성공 , 3 = 대체문자 전송 , 4 = 전송실패 , 5 = 전송취소
    ' - 미입력 시 전체조회
    Dim State(6)
    State(0) = "0"
    State(1) = "1"
    State(2) = "2"
    State(3) = "3"
    State(4) = "4"
    State(5) = "5"

    ' 검색대상 배열 ("ATS", "FTS", "FMS" 중 선택, 다중 선택 가능)
    ' └ ATS = 알림톡 , FTS = 친구톡(텍스트) , FMS = 친구톡(이미지)
    ' - 미입력 시 전체조회
    Dim Item(3)
    Item(0) = "ATS"
    Item(1) = "FTS"
    Item(2) = "FMS"

    ' 전송유형별 조회 (null , "0" , "1" 중 택 1)
    ' └ null = 전체 , 0 = 즉시전송건 , 1 = 예약전송건
    ' - 미입력 시 전체조회
    ReserveYN = ""

    ' 사용자권한별 조회 (true / false 중 택 1)
    ' └ false = 접수한 카카오톡 전체 조회 (관리자권한)
    ' └ true = 해당 담당자 계정으로 접수한 카카오톡만 조회 (개인권한)
    ' 미입력시 기본값 false 처리
    SenderYN = False

    ' 정렬방향, D-내림차순, A-오름차순
    Order = "D"

    ' 페이지 번호
    Page = 1

    PerPage = 30

    ' 조회하고자 하는 수신자명
    ' - 미입력시 전체조회
    QString = ""

    On Error Resume Next

    Set resultObj = m_KakaoService.Search(testCorpNum, SDate, EDate, State, Item, ReserveYN, SenderYN, Order, Page, PerPage, QString)

    If Err.Number <> 0 then
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

                    <legend>카카오톡 전송내역 조회 </legend>
                    <ul>
                    <% If code = 0 Then %>
                            <li> code (응답코드) : <%=resultObj.code%></li>
                            <li> message (응답메시지) : <%=resultObj.message%></li>
                            <li> total (총 검색결과 건수) : <%=resultObj.total%></li>
                            <li> pageNum (페이지 번호) : <%=resultObj.pageNum%></li>
                            <li> pageCount (페이지 개수) : <%=resultObj.pageCount%></li>
                            <li> perPage (페이지당 검색개수) : <%=resultObj.perPage%></li>
                    </ul>
                        <%
                            For i=0 To UBound(resultObj.list) -1
                        %>
                            <fieldset class="fieldset2">
                                <legend> 카카오톡 전송결과 [ <%=i+1%> / <%= UBound(resultObj.list)%> ] </legend>
                                <ul>
                                    <li>state (전송상태 코드) : <%=resultObj.list(i).state%> </li>
                                    <li>sendDT (전송일시) : <%=resultObj.list(i).sendDT%> </li>
                                    <li>result (전송결과 코드) : <%=resultObj.list(i).result%> </li>
                                    <li>resultDT (전송결과 수신일시) : <%=resultObj.list(i).resultDT%> </li>
                                    <li>contentType (카카오톡 유형) : <%=resultObj.list(i).contentType%> </li>
                                    <li>receiveNum (수신번호) : <%=resultObj.list(i).receiveNum%> </li>
                                    <li>receiveName (수신자명) : <%=resultObj.list(i).receiveName%> </li>
                                    <li>content (알림톡/친구톡 내용) : <%=resultObj.list(i).content%> </li>
                                    <li>altSubject (대체문자 제목) : <%=resultObj.list(i).altSubject%></li>
                                    <li>altContent (대체문자 내용) : <%=resultObj.list(i).altContent%></li>
                                    <li>altContentType (대체문자 전송타입) : <%=resultObj.list(i).altContentType%> </li>
                                    <li>altSendDT (대체문자 전송일시) : <%=resultObj.list(i).altSendDT%> </li>
                                    <li>altResult (대체문자 전송결과 코드) : <%=resultObj.list(i).altResult%> </li>
                                    <li>altResultDT (대체문자 전송결과 수신일시) : <%=resultObj.list(i).altResultDT%> </li>
                                    <li>receiptNum (접수번호) : <%=resultObj.list(i).receiptNum%> </li>
                                    <li>requestNum (요청번호) : <%=resultObj.list(i).requestNum%> </li>
                                    <li>interOPRefKey (파트너 지정키) : <%=resultObj.list(i).interOPRefKey%> </li>
                                </ul>
                            </fieldset>
                        <%
                            Next
                        Else
                        %>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    <% End If %>

            </fieldset>
        </div>
    </body>
</html>