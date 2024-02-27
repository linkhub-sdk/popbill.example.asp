<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 연동회원의 회사정보를 수정합니다..
    ' - https://developers.popbill.com/reference/easyfinbank/asp/api/member#UpdateCorpInfo
    '**************************************************************

    '팝빌회원 사업자번호
    CorpNum = "1234567890"

    '팝빌회원 아이디
    UserID = "testkorea"

    Set infoObj = New CorpInfo

    ' 대표자명
    infoObj.ceoname = "링크허브 대표자"

    ' 상호
    infoObj.corpName = "링크허브"

    ' 주소
    infoObj.addr	= "주소수정"

    ' 업태
    infoObj.bizType = "업태정보"

    ' 종목
    infoObj.bizClass = "종목정보"

    On Error Resume Next
    Set Presponse = m_EasyFinBankService.UpdateCorpInfo(CorpNum, infoObj, UserID)

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
                <legend>회사정보 수정</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
        </div>
    </body>
</html>
