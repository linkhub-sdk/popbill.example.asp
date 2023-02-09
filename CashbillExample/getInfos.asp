<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 다수건의 현금영수증 상태 및 요약 정보를 확인합니다. (1회 호출 시 최대 1,000건 확인 가능)
    ' - 리턴값 'CashbillInfo'의 변수 'stateCode'를 통해 현금영수증의 상태코드를 확인합니다.
    ' - 현금영수증 상태코드 : [https://developers.popbill.com/reference/cashbill/asp/response-code#state-code]
    ' - https://developers.popbill.com/reference/cashbill/asp/api/info#GetInfos
    '**************************************************************
    
    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"	 

    ' 팝빌회원 아이디
    userID = "testkorea"		 

    ' 조회할 현금영수증 문서번호 배열, 최대 1000건
    Dim mgtKeyList(3) 
    MgtKeyList(0) = "20220720-ASP-001"
    MgtKeyList(1) = "20220720-ASP-002"
    MgtKeyList(2) = "20220720-ASP-003"

    On Error Resume Next
    
    Set Presponse = m_CashbillService.GetInfos(testCorpNum, mgtKeyList, userID)

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
                <legend>현금영수증 상태 대량 확인</legend>
                <ul>
                    <% If code = 0 Then 
                        For i=0 To Presponse.Count-1 %>
                        <fieldset class="fieldset2">
                            <legend> 현금영수증 조회 결과 [<%=i+1%>]</legend>
                            <ul>
                                <li>itemKey (현금영수증 아이템키) : <%=Presponse.Item(i).itemKey%></li>
                                <li>mgtKey (문서번호) : <%=Presponse.Item(i).mgtKey%></li>
                                <li>tradeDate (거래일자) : <%=Presponse.Item(i).tradeDate%></li>
                                <li>tradeDT (거래일시) : <%=Presponse.Item(i).tradeDT%></li>
                                <li>tradeType (문서형태) : <%=Presponse.Item(i).tradeType%></li>
                                <li>tradeUsage (거래구분) : <%=Presponse.Item(i).tradeUsage%></li>
                                <li>tradeOpt (거래유형) : <%=Presponse.Item(i).tradeOpt%></li>
                                <li>taxationType (과세형태) : <%=Presponse.Item(i).taxationType%></li>
                                <li>totalAmount (거래금액) : <%=Presponse.Item(i).totalAmount%></li>
                                <li>issueDT (발행일시) : <%=Presponse.Item(i).issueDT%></li>
                                <li>regDT (등록일시) : <%=Presponse.Item(i).regDT%></li>
                                <li>stateMemo (상태메모) : <%=Presponse.Item(i).stateMemo%></li>
                                <li>stateCode (상태코드) : <%=Presponse.Item(i).stateCode%></li>
                                <li>stateDT (상태변경일시) : <%=Presponse.Item(i).stateDT%></li>
                                <li>identityNum (식별번호) : <%=Presponse.Item(i).identityNum%></li>
                                <li>itemName (상품명) : <%=Presponse.Item(i).itemName%></li>
                                <li>customerName (고객명) : <%=Presponse.Item(i).customerName%></li>
                                <li>confirmNum (국세청 승인번호) : <%=Presponse.Item(i).confirmNum%></li>
                                <li>orgConfirmNum (원본 현금영수증 국세청승인번호) : <%=Presponse.Item(i).orgConfirmNum%></li>
                                <li>orgTradeDate (원본 현금영수증 거래일자) : <%=Presponse.Item(i).orgTradeDate%></li>
                                <li>ntssendDT (국세청 전송일시) : <%=Presponse.Item(i).ntssendDT%></li>
                                <li>ntsresultDT (국세청 처리결과 수신일시) : <%=Presponse.Item(i).ntsResultDT%></li>
                                <li>ntsresultCode (국세청 처리결과 상태코드) : <%=Presponse.Item(i).ntsResultCode%></li>
                                <li>ntsresultMessage (국세청 처리결과 메시지) : <%=Presponse.Item(i).ntsResultMessage%></li>
                                <li>printYN (인쇄여부) : <%=Presponse.Item(i).printYN%></li>
                            </ul>
                            </fieldset>
                    <%	Next
                        Else %>
                        <li> Response.code : <%=code%> </li>
                        <li> Response.message : <%=message%> </li>
                    <% End If%> 
                    
                </ul>
            </fieldset>
         </div>
    </body>
</html>