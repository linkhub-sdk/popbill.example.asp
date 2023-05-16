<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="../Example.css" media="screen" />
        <title>���������ȸ API SDK ASP Example.</title>
    </head>
    <!--#include file="common.asp"-->
    <%
        '**************************************************************
        ' ����ڹ�ȣ 1�ǿ� ���� ��������� Ȯ���մϴ�.
        ' - https://developers.popbill.com/reference/bizinfocheck/asp/api/check#CheckBizInfo
        '**************************************************************
        '�˺�ȸ�� ����ڹ�ȣ
        MemberCorpNum = "1234567890"

        '��ȸ�� ����ڹ�ȣ
        CheckCorpNum = "6798700433"

        ' �˺�ȸ�� ���̵�
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
                <legend>���������ȸ</legend>
            <%
                If Not IsEmpty(result) Then

            %>

                <ul>
                    <li>corpNum (����ڹ�ȣ) : <%= result.corpNum%></li>
                    <li>companyRegNum (���ι�ȣ): <%=result.companyRegNum%></li>
                    <li>checkDT (Ȯ���Ͻ�) : <%=result.checkDT%></li>
                    <li>corpName (��ȣ): <%=result.corpName%></li>
                    <li>corpCode (��������ڵ�): <%=result.corpCode%></li>
                    <li>corpScaleCode (����Ը��ڵ�): <%=result.corpScaleCode%></li>
                    <li>personCorpCode (���ι����ڵ�): <%=result.personCorpCode%></li>
                    <li>headOfficeCode (���������ڵ�) : <%=result.headOfficeCode%></li>
                    <li>industryCode (����ڵ�) : <%=result.industryCode%></li>
                    <li>establishCode (���������ڵ�) : <%=result.establishCode%></li>
                    <li>establishDate (��������) : <%=result.establishDate%></li>
                    <li>CEOName (��ǥ�ڸ�) : <%=result.ceoname%></li>
                    <li>workPlaceCode (����屸���ڵ�): <%=result.workPlaceCode%></li>
                    <li>addrCode (�ּұ����ڵ�) : <%=result.addrCode%></li>
                    <li>zipCode (������ȣ) : <%=result.zipCode%></li>
                    <li>addr (�ּ�) : <%=result.addr%></li>
                    <li>addrDetail (���ּ�) : <%=result.addrDetail%></li>
                    <li>enAddr (�����ּ�) : <%=result.enAddr%></li>
                    <li>bizClass (����) : <%=result.bizClass%></li>
                    <li>bizType (����) : <%=result.bizType%></li>
                    <li>result (����ڵ�) : <%=result.result%></li>
                    <li>resultMessage (����޽���) : <%=result.resultMessage%></li>
                    <li>closeDownTaxType (����ڰ�������) : <%=result.closeDownTaxType%></li>
                    <li>closeDownTaxTypeDate (����������ȯ����):<%=result.closeDownTaxTypeDate%></li>
                    <li>closeDownState (���������) : <%=result.closeDownState%></li>
                    <li>closeDownStateDate (���������) : <%=result.closeDownStateDate%></li>
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