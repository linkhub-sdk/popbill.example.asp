<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 사업자번호를 조회하여 연동회원 가입여부를 확인합니다.
    ' - https://developers.popbill.com/reference/httaxinvoice/asp/api/member#CheckIsMember
    '**************************************************************

    ' 사업자번호 ("-"제외)
    CorpNum = "1231212312"

    On Error Resume Next

    Set Presponse = m_HTTaxinvoiceService.CheckIsMember(CorpNum,LinkID)

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
                <legend>연동회원사 가입 여부 확인 결과</legend>
                <ul>
                    <li>Response.code : <%=CStr(code)%></li>
                    <li>Response.message : <%=message%></li>
                </ul>
            </fieldset>
        </div>
    </body>
</html>