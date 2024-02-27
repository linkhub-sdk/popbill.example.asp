<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 계좌조회 서비스를 이용할 계좌를 팝빌에 등록합니다.
    ' - https://developers.popbill.com/reference/easyfinbank/asp/api/manage#GetBankAccountInfo
    '**************************************************************

    ' 팝빌회원 사업자번호
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 계좌정보 객체 생성
    Set infoObj = New EasyFinBankAccountForm

    ' 기관코드
    ' 산업은행-0002 / 기업은행-0003 / 국민은행-0004 /수협은행-0007 / 농협은행-0011 / 우리은행-0020
    ' SC은행-0023 / 대구은행-0031 / 부산은행-0032 / 광주은행-0034 / 제주은행-0035 / 전북은행-0037
    ' 경남은행-0039 / 새마을금고-0045 / 신협은행-0048 / 우체국-0071 / KEB하나은행-0081 / 신한은행-0088 /씨티은행-0027
    infoObj.BankCode = ""

    ' 계좌번호 하이픈('-') 제외
    infoObj.AccountNumber = ""

    ' 계좌비밀번호
    infoObj.AccountPWD = ""

    ' 계좌유형, "법인" 또는 "개인" 입력
    infoObj.AccountType = ""

    ' 예금주 식별정보 (‘-‘ 제외)
    ' 계좌유형이 "법인"인 경우 : 사업자번호(10자리)
    ' 계좌유형이 "개인"인 경우 : 예금주 생년월일 (6자리-YYMMDD)
    infoObj.IdentityNumber = ""

    ' 계좌 별칭
    infoObj.AccountName = ""

    ' 인터넷뱅킹 아이디 (국민은행 필수)
    infoObj.BankID = ""

    ' 조회전용 계정 아이디 (대구은행, 신협, 신한은행 필수)
    infoObj.FastID = ""

    ' 조회전용 계정 비밀번호 (대구은행, 신협, 신한은행 필수
    infoObj.FastPWD = ""

    ' 결제기간(개월), 1~12 입력가능, 미기재시 기본값(1) 처리
    ' - 파트너 과금방식의 경우 입력값에 관계없이 1개월 처리
    infoObj.UsePeriod = ""

    ' 메모
    infoObj.Memo = ""

    On Error Resume Next
        Set Presponse = m_EasyFinBankService.RegistBankAccount(CorpNum, infoObj, UserID)

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
                <legend>계좌 등록</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
        </div>
    </body>
</html>
