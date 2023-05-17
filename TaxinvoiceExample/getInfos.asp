<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 다수건의 세금계산서 상태 및 요약 정보를 확인합니다. (1회 호출 시 최대 1,000건 확인 가능)
    ' - 리턴값 'TaxinvoiceInfo'의 변수 'stateCode'를 통해 세금계산서의 상태코드를 확인합니다.
    ' - 세금계산서 상태코드 [https://developers.popbill.com/reference/taxinvoice/asp/response-code#state-code]
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/info#GetInfos
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    ' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
    KeyType = "SELL"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 세금계산서 문서번호 배열, 최대 1000건
    Dim MgtKeyList(3)
    MgtKeyList(0) = "20220720-ASP-001"
    MgtKeyList(1) = "20220720-ASP-002"
    MgtKeyList(2) = "20220720-ASP-003"

    On Error Resume Next

    Set result = m_TaxinvoiceService.GetInfos(testCorpNum, KeyType, MgtKeyList, UserID)

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
                <legend>세금계산서 상태/요약 정보 확인 - 대량</legend>
                <%
                    If code = 0 Then
                        For i=0 To result.Count-1
                %>
                            <fieldset class="fieldset2">
                                <legend> TaxinvoiceResult : <%=i+1%> </legend>
                                    <ul>
                                        <li> itemKey (세금계산서 아이템키) :  <%=result.Item(i).itemKey%> </li>
                                        <li> taxType (과세형태) :  <%=result.Item(i).taxType%> </li>
                                        <li> writeDate (작성일자) :  <%=result.Item(i).writeDate%> </li>
                                        <li> regDT (임시저장 일자) :  <%=result.Item(i).regDT%> </li>
                                        <li> issueType (발행형태) :  <%=result.Item(i).issueType %> </li>
                                        <li> supplyCostTotal (공급가액 합계) :  <%=result.Item(i).supplyCostTotal%> </li>
                                        <li> taxTotal (세액 합계) :  <%=result.Item(i).taxTotal%> </li>
                                        <li> purposeType (영수/청구) :  <%=result.Item(i).purposeType%> </li>
                                        <li> issueDT (발행일시) :  <%=result.Item(i).issueDT%> </li>
                                        <li> lateIssueYN (지연발행 여부) :  <%=result.Item(i).lateIssueYN%> </li>
                                        <li> preIssueDT (발행예정일시) :  <%=result.Item(i).preIssueDT%> </li>
                                        <li> openYN (개봉 여부) :  <%=result.Item(i).openYN%> </li>
                                        <li> openDT (개봉일시) :  <%=result.Item(i).openDT%> </li>
                                        <li> stateMemo (상태메모) :  <%=result.Item(i).stateMemo%> </li>
                                        <li> stateCode (상태코드) :  <%=result.Item(i).stateCode%> </li>
                                        <li> stateDT (상태 변경일시) :  <%=result.Item(i).stateDT%> </li>
                                        <li> ntsconfirmNum (국세청 승인번호) :  <%=result.Item(i).ntsconfirmNum %> </li>
                                        <li> ntsresult (국세청 전송결과) :  <%=result.Item(i).ntsresult%> </li>
                                        <li> ntssendDT (국세청 전송일시) :  <%=result.Item(i).ntssendDT%> </li>
                                        <li> ntsresultDT  (국세청 결과 수신일시) :  <%=result.Item(i).ntsresultDT%> </li>
                                        <li> ntssendErrCode (전송실패 사유코드) :  <%=result.Item(i).ntssendErrCode%> </li>
                                        <li> modifyCode (수정사유코드) : <%=result.Item(i).modifyCode%></li>
                                        <li> interOPYN (연동문서여부) :  <%=result.Item(i).interOPYN%> </li>
                                        <li> invoicerCorpName (공급자 상호) :  <%=result.Item(i).invoicerCorpName%> </li>
                                        <li> invoicerCorpNum (공급자 사업자번호) :  <%=result.Item(i).invoicerCorpNum%> </li>
                                        <li> invoicerMgtKey (공급자 문서번호) :  <%=result.Item(i).invoicerMgtKey%> </li>
                                        <li> invoicerPrintYN (공급자 인쇄여부) :  <%=result.Item(i).invoicerPrintYN%> </li>
                                        <li> invoiceeCorpName (공급받는자 상호) :  <%=result.Item(i).invoiceeCorpName%> </li>
                                        <li> invoiceeCorpNum (공급받는자 사업자번호) :  <%=result.Item(i).invoiceeCorpNum%> </li>
                                        <li> invoiceeMgtKey (공급받는자 문서번호) :  <%=result.Item(i).invoiceeMgtKey%> </li>
                                        <li> invoiceePrintYN (공급받는자 인쇄여부) :  <%=result.Item(i).invoiceePrintYN%> </li>
                                        <li> closeDownState (공급받는자 휴폐업상태) :  <%=result.Item(i).closeDownState%> </li>
                                        <li> closeDownStateDate (공급받는자 휴폐업일자) :  <%=result.Item(i).closeDownStateDate%> </li>
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