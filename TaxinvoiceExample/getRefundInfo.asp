<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 환불 신청 정보를 확인합니다.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/point#GetRefundInfo
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

	'환불 코드
	refundCode = "023040000017"

    '팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    Dim refundHistory: Set refundHistory = m_TaxinvoiceService.GetRefundInfo(testCorpNum, refundCode, UserID)

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
                <legend>환불 신청 상태 조회</legend>
                <%
                    If code = 0 Then
                %>
                    <fieldset class="fieldset2">
                        <legend> GetRefundInfo </legend>
                            <ul>
                                <li> reqDT (신청 일시) : <%=refundHistory.reqDT%></li>
                                <li> requestPoint (환불 신청포인트) : <%=refundHistory.requestPoint%></li>
                                <li> accountBank (환불계좌 은행명) : <%=refundHistory.accountBank%></li>
                                <li> accountNum (환불계좌번호) : <%=refundHistory.accountNum%></li>
                                <li> accountName (환불계좌 예금주명) : <%=refundHistory.accountName%></li>
                                <li> state (상태) : <%=refundHistory.state%></li>
                                <li> reason (환불사유) : <%=refundHistory.reason%></li>
                            </ul>
                        </fieldset>
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