<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="../Example.css" media="screen" />
        <title>기업정보조회 API SDK ASP Example.</title>
    </head>
    <!--#include file="common.asp"-->
    <%
        '**************************************************************
        ' 사업자번호 1건에 대한 기업정보를 확인합니다.
        ' - https://developers.popbill.com/reference/bizinfocheck/asp/api/check#CheckBizInfo
        '**************************************************************
        '팝빌회원 사업자번호
        MemberCorpNum = "1234567890"

        '조회할 사업자번호
        CheckCorpNum = "6798700433"

        ' 팝빌회원 아이디
        UserID = "testkorea"

        On Error Resume Next
            Set result = m_BizInfoCheckService.checkBizInfo(MemberCorpNum, CheckCorpNum, UserID )

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
                <legend>기업정보조회</legend>
            <%
                If Not IsEmpty(result) Then

            %>

                <ul>
                    <li>corpNum (사업자번호) : <%= result.corpNum%></li>
                    <li>companyRegNum (법인번호): <%=result.companyRegNum%></li>
                    <li>checkDT (확인일시) : <%=result.checkDT%></li>
                    <li>corpName (상호): <%=result.corpName%></li>
                    <li>corpCode (기업형태코드): <%=result.corpCode%></li>
                    <li>corpScaleCode (기업규모코드): <%=result.corpScaleCode%></li>
                    <li>personCorpCode (개인법인코드): <%=result.personCorpCode%></li>
                    <li>headOfficeCode (본점지점코드) : <%=result.headOfficeCode%></li>
                    <li>industryCode (산업코드) : <%=result.industryCode%></li>
                    <li>establishCode (설립구분코드) : <%=result.establishCode%></li>
                    <li>establishDate (설립일자) : <%=result.establishDate%></li>
                    <li>CEOName (대표자명) : <%=result.ceoname%></li>
                    <li>workPlaceCode (사업장구분코드): <%=result.workPlaceCode%></li>
                    <li>addrCode (주소구분코드) : <%=result.addrCode%></li>
                    <li>zipCode (우편번호) : <%=result.zipCode%></li>
                    <li>addr (주소) : <%=result.addr%></li>
                    <li>addrDetail (상세주소) : <%=result.addrDetail%></li>
                    <li>enAddr (영문주소) : <%=result.enAddr%></li>
                    <li>bizClass (업종) : <%=result.bizClass%></li>
                    <li>bizType (업태) : <%=result.bizType%></li>
                    <li>result (결과코드) : <%=result.result%></li>
                    <li>resultMessage (결과메시지) : <%=result.resultMessage%></li>
                    <li>closeDownTaxType (사업자과세유형) : <%=result.closeDownTaxType%></li>
                    <li>closeDownTaxTypeDate (과세유형전환일자):<%=result.closeDownTaxTypeDate%></li>
                    <li>closeDownState (휴폐업상태) : <%=result.closeDownState%></li>
                    <li>closeDownStateDate (휴폐업일자) : <%=result.closeDownStateDate%></li>
                </ul>

            <%
                End If
                If Not IsEmpty(code) then
            %>

            <ul>
                <li>Response.code : <%= code %> </li>
                <li>Response.message : <%= message %></li>
            </ul>
            <%
                End If
            %>

            </fieldset>
    </body>
</html>
