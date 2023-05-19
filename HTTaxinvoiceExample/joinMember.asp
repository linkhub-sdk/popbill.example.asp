<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->

<%
    '**************************************************************
    ' 사용자를 연동회원으로 가입처리합니다.
    ' - https://developers.popbill.com/reference/httaxinvoice/asp/api/member#JoinMember
    '**************************************************************

    ' 회원정보 객체 생성
    Set joinInfo = New JoinForm

    ' 링크아이디
    joinInfo.LinkID = "TESTER"

    ' 사업자번호, "-"제외 10자리
    joinInfo.CorpNum = "1234567890"

    ' 대표자성명
    joinInfo.CEOName = "대표자성명"

    ' 상호명
    joinInfo.CorpName =  "상호"

    ' 주소
    joinInfo.Addr =   "주소"

    ' 업태
    joinInfo.BizType =  "업태"

    ' 종목
    joinInfo.BizClass = "종목"

    ' 아이디 (6자 이상 20자 미만)
    joinInfo.ID =  "userid"

    ' 비밀번호 (8자 이상 20자 이하) 영문, 숫자 ,특수문자 조합
    joinInfo.Password =  "asdf1234!@#$"

    ' 담당자명
    joinInfo.ContactName = "담당자명"

    ' 담당자연락처
    joinInfo.ContactTEL = ""

    ' 담당자 이메일
    joinInfo.ContactEmail = ""

    On Error Resume Next

    Set Presponse = m_HTTaxinvoiceService.JoinMember(joinInfo)

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
                <legend>연동회원 가입요청</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
        </div>
    </body>
</html>