<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 계좌의 정액제 해지를 요청합니다.
    ' - https://developers.popbill.com/reference/easyfinbank/asp/api/manage#CloseBankAccount
    '**************************************************************

    ' 팝빌회원 사업자번호
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"


    ' 기관코드
    ' 산업은행-0002 / 기업은행-0003 / 국민은행-0004 /수협은행-0007 / 농협은행-0011 / 우리은행-0020
    ' SC은행-0023 / 대구은행-0031 / 부산은행-0032 / 광주은행-0034 / 제주은행-0035 / 전북은행-0037
    ' 경남은행-0039 / 새마을금고-0045 / 신협은행-0048 / 우체국-0071 / KEB하나은행-0081 / 신한은행-0088 /씨티은행-0027
    BankCode = ""

    ' 계좌번호 하이픈('-') 제외
    AccountNumber = ""

    ' 해지유형, "일반", "중도" 중 택 1
    ' - 일반(일반해지) - 해지 요청일이 포함된 정액제 이용기간 만료 후 해지
    ' - 중도(중도해지) - 해지 요청시 즉시 정지되고 팝빌 담당자가 승인시 해지
    ' └ 중도일 경우, 정액제 잔여기간은 일할로 계산되어 포인트 환불(무료 이용기간에 해지하면 전액 환불)
    CloseType = "중도"

    On Error Resume Next
        Set Presponse = m_EasyFinBankService.CloseBankAccount(CorpNum, BankCode, AccountNumber, CloseType, UserID)

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
                <legend>계좌 정액제 해지신청</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
        </div>
    </body>
</html>
