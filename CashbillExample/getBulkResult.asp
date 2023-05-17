<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 접수시 기재한 SubmitID를 사용하여 현금영수증 접수결과를 확인합니다.
    ' - 개별 현금영수증 처리상태는 접수상태(txState)가 완료(2) 시 반환됩니다.
    ' - https://developers.popbill.com/reference/cashbill/asp/api/issue#GetBulkResult
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    ' 제출아이디, 최대 36자리 (영문, 숫자, "-" 조합)
    SubmitID = "20221109-ASP-BULK001"

    ' 팝빌회원아이디
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_CashbillService.GetBulkResult(testCorpNum, SubmitID, UserID)

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
                <legend>초대량 접수 결과 확인</legend>
                <%
                    If code = 0 Then
                %>
                    <ul>
                        <li> code (응답코드) :  <%=result.code%> </li>
                        <li> message (응답메시지) :  <%=result.message%> </li>
                        <li> submitID (제출아이디) :  <%=result.submitID%> </li>
                        <li> submitCount (현금영수증 접수 건수) :  <%=result.submitCount%> </li>
                        <li> successCount (현금영수증 발행 성공 건수) : <%=result.successCount%></li>
                        <li> failCount (현금영수증 발행 실패 건수) :  <%=result.failCount %> </li>
                        <li> txState (접수상태코드) :  <%=result.txState%> </li>
                        <li> txResultCode (접수 결과코드) :  <%=result.txResultCode%> </li>
                        <li> txStartDT (발행처리 시작일시) :  <%=result.txStartDT%> </li>
                        <li> txEndDT (발행처리 완료일시	) :  <%=result.txEndDT%> </li>
                        <li> receiptDT (접수일시) :  <%=result.receiptDT%> </li>
                        <li> receiptID (접수아이디) :  <%=result.receiptID%> </li>
                    </ul>
                    <%   Dim i
                        For i=0 To UBound(result.issueResult) -1
                     %>
                     <fieldset class="fieldset2">
                        <legend>  issueResult (발행 결과) [ <%=i+1%> / <%=UBound(result.issueResult)%> ]</legend>
                        <ul>
                            <li> mgtKey (문서번호) : <%=result.issueResult(i).mgtKey %>
                            <li> code (응답코드) : <%=result.issueResult(i).code %>
                            <li> message (응답메시지) : <%=result.issueResult(i).message %>
                            <li> confirmNum (국세청승인번호) : <%=result.issueResult(i).confirmNum %>
                            <li> tradeDate (거래일자) : <%=result.issueResult(i).tradeDate %>
                            <li> tradeDT (거래일시) : <%=result.issueResult(i).tradeDT %>
                        </ul>
                    </fieldset>
                     <% Next %>
                    <%
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