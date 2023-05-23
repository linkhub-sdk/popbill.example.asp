<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 세금계산서 1건의 상세정보를 XML로 반환합니다.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/info#GetXML
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외 10자리
    CorpNum = "1234567890"

    ' 세금계산서 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
    KeyType = "SELL"

    ' 문서번호
    MgtKey = "20220720-ASP-001"

    On Error Resume Next

    Set taxXML = m_TaxinvoiceService.GetXML(CorpNum, KeyType, MgtKey)

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
                <legend>상세정보 확인 - XML</legend>
                <%

                    If code = 0 Then
                %>
                <ul>
                    <li>code (응답코드) : <%=taxXML.code%></li>
                    <li>message (응답메시지) : <%=taxXML.message%></li>
                    <li>retObject (전자세금계산서 XML문서) : <%=Replace(taxXML.retObject, "<", "&lt;")%></li>
                    <!-- Browser에서 xml문서를 출력하기 위해 '<' &lt로 치환하였습니다. -->
                </ul>

                <%
                    Else
                %>
                    <ul>
                        <li>Response.dcode : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    </ul>
                <%
                    End If
                %>
            </fieldset>
        </div>
    </body>
</html>