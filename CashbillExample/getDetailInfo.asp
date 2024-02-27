<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 현금영수증 1건의 상세정보를 조회합니다.
    ' - https://developers.popbill.com/reference/cashbill/asp/api/info#GetDetailInfo
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 문서번호
    mgtKey = "20220720-ASP-001"

    On Error Resume Next

    Set Presponse = m_CashbillService.GetDetailInfo(CorpNum, mgtKey, UserID)

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
                <legend>현금영수증 상세정보 확인</legend>
                <ul>
                    <% If code = 0 Then %>
                        <fieldset class="fieldset2">
                            <ul>
                                <li>mgtKey (문서번호) : <%=Presponse.mgtKey%></li>
                                <li>confirmNum (국세청승인번호) : <%=Presponse.confirmNum%></li>
                                <li>orgConfirmNum (원본 현금영수증 국세청승인번호) : <%=Presponse.orgConfirmNum%></li>
                                <li>orgTradeDate (원본 현금영수증 거래일자) : <%=Presponse.orgTradeDate%></li>
                                <li>tradeDate (거래일자) : <%=Presponse.tradeDate%></li>
                                <li>tradeDT (거래일시) : <%=Presponse.tradeDT%></li>
                                <li>tradeType (문서형태) : <%=Presponse.tradeType %></li>
                                <li>tradeUsage (거래구분) : <%=Presponse.tradeUsage%></li>
                                <li>tradeOpt (거래유형) : <%=Presponse.tradeOpt %></li>
                                <li>taxationType (과세형태) : <%=Presponse.taxationType%></li>
                                <li>totalAmount (거래금액) : <%=Presponse.totalAmount%></li>
                                <li>supplyCost (공급가액) : <%=Presponse.supplyCost%></li>
                                <li>tax (부가세) : <%=Presponse.tax %></li>
                                <li>serviceFee (봉사료) : <%=Presponse.serviceFee%></li>
                                <li>franchiseCorpNum (가맹점 사업자번호) : <%=Presponse.franchiseCorpNum%></li>
                                <li>franchiseTaxRegID (가맹점 종사업장 식별번호) : <%=Presponse.franchiseTaxRegID%></li>
                                <li>franchiseCorpName (가맹점 상호) : <%=Presponse.franchiseCorpName%></li>
                                <li>franchiseCEOName (가맹점 대표자명) : <%=Presponse.franchiseCEOName%></li>
                                <li>franchiseAddr (가맹점 주소) : <%=Presponse.franchiseAddr%></li>
                                <li>franchiseTEL (가맹점 전화번호) : <%=Presponse.franchiseTEL %></li>
                                <li>identityNum (식별번호) : <%=Presponse.identityNum%></li>
                                <li>customerName (주문자명) : <%=Presponse.customerName%></li>
                                <li>itemName (주문상품명) : <%=Presponse.itemName%></li>
                                <li>orderNumber (주문번호) : <%=Presponse.orderNumber%></li>
                                <li>email (이메일) : <%=Presponse.email%></li>
                                <li>hp (휴대폰) : <%=Presponse.hp%></li>
                                <li>smssendYN (알림문자 전송여부) : <%=Presponse.smssendYN%></li>
                                <li>cancelType (취소사유) : <%=Presponse.cancelType %></li>
                            </ul>
                        </fieldset>
                    <%	Else %>
                        <li> Response.code : <%=code%> </li>
                        <li> Response.message : <%=message%> </li>
                    <% End If%>

                </ul>
            </fieldset>
        </div>
    </body>
</html>
