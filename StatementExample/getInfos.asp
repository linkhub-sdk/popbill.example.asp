<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 다수건의 전자명세서 상태 및 요약정보 확인합니다. (1회 호출 시 최대 1,000건 확인 가능)
    ' - https://developers.popbill.com/reference/statement/asp/api/info#GetInfos
    '**************************************************************

    ' 팝빌회원 사업자번호
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 명세서 코드 - 121(거래명세서), 122(청구서), 123(견적서), 124(발주서), 125(입금표), 126(영수증)
    itemCode = "121"

    ' 문서번호 배열, 최대 1000건
    Dim mgtKeyList(2)
    mgtKeyList(0) = "20220720-ASP-001"
    mgtKeyList(1) = "20220720-ASP-002"

    On Error Resume Next

    Set result = m_StatementService.GetInfos(CorpNum, itemCode, mgtKeyList, UserID)

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
                <legend>전자명세서 상태/요약정보 확인 - 대량 </legend>
                <ul>
                    <% If code = 0 Then
                        For i=0 To result.Count-1 %>

                        <fieldset class="fieldset2">
                            <legend> 전자명세서 정보 [<%=i+1%>] </legend>
                            <ul>
                                <li> itemCode(문서종류코드) : <%=result.Item(i).itemCode %></li>
                                <li> itemKey(팝빌번호) : <%=result.Item(i).itemKey %></li>
                                <li> invoiceNum(팝빌승인번호) : <%=result.Item(i).invoiceNum %></li>
                                <li> mgtKey(파트너 문서번호) : <%=result.Item(i).mgtKey %></li>
                                <li> taxType(세금형태) : <%=result.Item(i).taxType %></li>
                                <li> writeDate(작성일자) : <%=result.Item(i).writeDate %></li>
                                <li> regDT(등록일시) : <%=result.Item(i).regDT %></li>
                                <li> senderCorpName(발신자 상호) : <%=result.Item(i).senderCorpName %></li>
                                <li> senderCorpNum(발신자 사업자번호) : <%=result.Item(i).senderCorpNum %></li>
                                <li> senderPrintYN(발신자 인쇄여부) : <%=result.Item(i).senderPrintYN %></li>
                                <li> receiverCorpName(수신자 상호) : <%=result.Item(i).receiverCorpName %></li>
                                <li> receiverCorpNum(수신자 사업자번호) : <%=result.Item(i).receiverCorpNum %></li>
                                <li> receiverPrintYN(수신자 인쇄여부) : <%=result.Item(i).receiverPrintYN %></li>
                                <li> supplyCostTotal(공급가액 합계) : <%=result.Item(i).supplyCostTotal %></li>
                                <li> taxTotal(세액 합계) : <%=result.Item(i).taxTotal %></li>
                                <li> purposeType(영수/청구) : <%=result.Item(i).purposeType %></li>
                                <li> issueDT(발행일시) : <%=result.Item(i).issueDT %></li>
                                <li> stateCode(상태코드) : <%=result.Item(i).stateCode %></li>
                                <li> stateDT(상태 변경일시) : <%=result.Item(i).stateDT %></li>
                                <li> stateMemo(상태메모) : <%=result.Item(i).stateMemo %></li>
                                <li> openYN(메일 개봉 여부) : <%=result.Item(i).openYN %></li>
                                <li> openDT(개봉 일시) : <%=result.Item(i).openDT %></li>
                            </ul>
                        </fieldset>
                    <%
                        Next
                        Else
                    %>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    <% End If %>
                </ul>
            </fieldset>
        </div>
    </body>
</html>
