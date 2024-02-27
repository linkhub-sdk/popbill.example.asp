<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
    <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
    <title>�˺� SDK ASP Example.</title>
</head>
<body>
<div id="content">
    <p class="heading1">�˺� ���� API SDK ASP Example.</p>
    <br/>
    <fieldset class="fieldset1">
        <legend>�߽Ź�ȣ �������</legend>
        <ul>
            <li><a href="checkSenderNumber.asp">checkSenderNumber</a> - �߽Ź�ȣ ��Ͽ��� Ȯ��</li>
            <li><a href="getSenderNumberMgtURL.asp">getSenderNumberMgtURL</a> - �߽Ź�ȣ ���� �˾� URL</li>
            <li><a href="getSenderNumberList.asp">getSenderNumberList</a> - �߽Ź�ȣ ��� Ȯ��</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>���� ����</legend>
        <ul>
            <li><a href="sendSMS.asp">sendSMS</a> - �ܹ� ����</li>
            <li><a href="sendSMS_Multi.asp">sendSMS</a> - �ܹ� ���� [�뷮]</li>
            <li><a href="sendLMS.asp">sendLMS</a> - �幮 ����</li>
            <li><a href="sendLMS_Multi.asp">sendLMS</a> - �幮 ���� [�뷮]</li>
            <li><a href="sendMMS.asp">sendMMS</a> - ���� ����</li>
            <li><a href="sendMMS_Multi.asp">sendMMS</a> - ���� ���� [�뷮]</li>
            <li><a href="sendXMS.asp">sendXMS</a> - �ܹ�/�幮 �ڵ��ν� ����</li>
            <li><a href="sendXMS_Multi.asp">sendXMS</a> - �ܹ�/�幮 �ڵ��ν� ���� [�뷮]</li>
            <li><a href="cancelReserve.asp">cancelReserve</a> - �������� ���</li>
            <li><a href="cancelReserveRN.asp">cancelReserveRN</a> - �������� ��� (��û��ȣ �Ҵ�)</li>
            <li><a href="cancelReservebyRCV.asp">cancelReservebyRCV</a> - �������� ��� (������ȣ, ���Ź�ȣ)</li>
            <li><a href="cancelReserveRNbyRCV.asp">cancelReserveRNbyRCV</a> - �������� ��� (��û��ȣ, ���Ź�ȣ)</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>����Ȯ��</legend>
        <ul>
            <li><a href="getMessages.asp">getMessages</a> - ���۳��� Ȯ��</li>
            <li><a href="getMessagesRN.asp">getMessagesRN</a> - ���۳��� Ȯ�� (��û��ȣ �Ҵ�)</li>
            <li><a href="search.asp">search</a> - ���۳��� ��� ��ȸ</li>
            <li><a href="getSentListURL.asp">getSentListURL</a> - ���� ���۳��� �˾� URL</li>
            <li><a href="getAutoDenyList.asp">getAutoDenyList</a> - 080 ���Űź� ��� Ȯ��</li>
            <li><a href="CheckAutoDenyNumber.asp">CheckAutoDenyNumber</a> - 080 ��ȣ Ȯ��</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>����Ʈ ����</legend>
        <ul>
            <li><a href="getBalance.asp">getBalance</a> - ����ȸ�� �ܿ�����Ʈ Ȯ��</li>
            <li><a href="getChargeURL.asp">getChargeURL</a> - ����ȸ�� ����Ʈ���� URL</li>
            <li><a href="getPaymentURL.asp">GetPaymentURL</a> - ����ȸ�� ����Ʈ ���系�� URL</li>
            <li><a href="getUseHistoryURL.asp">GetUseHistoryURL</a> - ����ȸ�� ����Ʈ ��볻�� URL</li>
            <li><a href="getPartnerBalance.asp">getPartnerBalance</a> - ��Ʈ�� �ܿ�����Ʈ Ȯ��</li>
            <li><a href="getPartnerURL.asp">getPartnerURL</a> - ��Ʈ�� ����Ʈ���� URL</li>
            <li><a href="getChargeInfo.asp">getChargeInfo</a> - �������� Ȯ��</li>
            <li><a href="getUnitCost.asp">getUnitCost</a> - ���� �ܰ� Ȯ��</li>
            <li><a href="paymentRequest.asp">paymentRequest</a> - ����ȸ�� ������ �Աݽ�û</li>
            <li><a href="getSettleResult.asp">getSettleResult</a> - ����ȸ�� ������ �Աݽ�û Ȯ��</li>
            <li><a href="getPaymentHistory.asp">getPaymentHistory</a> - ����ȸ�� ����Ʈ �������� Ȯ��</li>
            <li><a href="getUseHistory.asp">getUseHistory</a> - ����ȸ�� ����Ʈ ��볻�� Ȯ��</li>
            <li><a href="refund.asp">refund</a> - ����ȸ�� ����Ʈ ȯ�ҽ�û</li>
            <li><a href="getRefundHistory.asp">getRefundHistory</a> - ����ȸ�� ����Ʈ ȯ�ҳ��� Ȯ��</li>
            <li><a href="getRefundInfo.asp">getRefundInfo</a> - ȯ�� ��û ���� ��ȸ</li>
			<li><a href="getRefundableBalance.asp">getRefundableBalance</a> - ȯ�� ���� ����Ʈ ��ȸ</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>ȸ������</legend>
        <ul>
            <li><a href="checkIsMember.asp">checkIsMember</a> - ����ȸ�� ���Կ��� Ȯ��</li>
            <li><a href="checkID.asp">checkID</a> - ���̵� �ߺ� Ȯ��</li>
            <li><a href="joinMember.asp">joinMember</a> - ����ȸ�� �ű԰���</li>
            <li><a href="getAccessURL.asp">getAccessURL</a> - �˺� �α��� URL</li>
            <li><a href="getCorpInfo.asp">getCorpInfo</a> - ȸ������ Ȯ��</li>
            <li><a href="updateCorpInfo.asp">updateCorpInfo</a> - ȸ������ ����</li>
            <li><a href="registContact.asp">registContact</a> - ����� ���</li>
            <li><a href="getContactInfo.asp">getContactInfo</a> - ����� ���� Ȯ��</li>
            <li><a href="listContact.asp">listContact</a> - ����� ��� Ȯ��</li>
            <li><a href="updateContact.asp">updateContact</a> - ����� ���� ����</li>
            <li><a href="quitMember.asp">quitMember</a> - �˺�ȸ�� Ż��</li>
        </ul>
    </fieldset>
</div>
</body>
</html>
