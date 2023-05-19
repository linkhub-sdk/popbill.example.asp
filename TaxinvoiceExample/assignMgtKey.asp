<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 팝빌 사이트를 통해 발행하여 문서번호가 부여되지 않은 세금계산서에 문서번호를 할당합니다.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/etc#AssignMgtKey
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    ' 세금계산서 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
    mgtKeyType= "SELL"

    ' 세금계산서 아이템키, 문서 목록조회(Search) API의 반환항목중 ItemKey 참조
    itemKey = "018082116393500001"

    ' 할당할 문서번호, 숫자, 영문 '-', '_' 조합으로 1~24자리까지
    ' 사업자번호별 중복없는 고유번호 할당
    mgtKey = "20220720-ASP-001"

    On Error Resume Next

    Set Presponse = m_TaxinvoiceService.AssignMgtKey(testCorpNum, mgtKeyType, itemKey, mgtKey)

    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else
        code = Presponse.code
        message =Presponse.message
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>문서번호 할당 </legend>
                <ul>
                    <li> Response.code : <%=code%></li>
                    <li> Response.message : <%=message%></li>
                </ul>
            </fieldset>
        </div>
    </body>
</html>