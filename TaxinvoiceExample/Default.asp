<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		
		<title>�˺� SDK ASP Example.</title>
	</head>

	<body>

		<div id="content">

			<p class="heading1">�˺� ���ݰ�꼭 SDK ASP Example.</p>
			
			<br/>

			<fieldset class="fieldset1">
				<legend>�˺� �⺻ API</legend>

				<fieldset class="fieldset2">
					<legend>ȸ���� ����</legend>
					<ul>					
						<li><a href="checkIsMember.asp">checkCorpIsMember</a> - ����ȸ���� ���� ���� Ȯ��</li>
						<li><a href="checkID.asp">checkID</a> - ���̵� �ߺ�Ȯ��</li>
						<li><a href="joinMember.asp">joinMember</a> - ����ȸ���� ���� ��û</li>
						<li><a href="getBalance.asp">getBalance</a> - ����ȸ���� �ܿ�����Ʈ Ȯ��</li>
						<li><a href="getPartnerBalance.asp">getPartnerBalance</a> - ��Ʈ�� �ܿ�����Ʈ Ȯ��</li>
						<li><a href="getPopbillURL.asp">getPopbillURL</a> - �˺� SSO URL ��û</li>
						<li><a href="listContact.asp">listContact</a> - ����� ��� ��ȸ</li>
						<li><a href="updateContact.asp">updateContact</a> - ����� ���� ����</li>
						<li><a href="registContact.asp">registContact</a> - ����� �߰�</li>
						<li><a href="updateCorpInfo.asp">updateCorpInfo</a> - ȸ������ ����</li>
						<li><a href="getCorpInfo.asp">getCorpInfo</a> - ȸ������ Ȯ��</li>
					</ul>
				</fieldset>

			</fieldset>
			
			<br />
			
			<fieldset class="fieldset1">
				<legend>���ڼ��ݰ�꼭 ���� API</legend>
				
				<fieldset class="fieldset2">
					<legend>���/����/Ȯ��/����</legend>
					<ul>
						<li><a href="checkMgtKeyInUse.asp">checkMgtKeyInUse</a> - ����������ȣ�� ���/��뿩�� Ȯ��</li>
						<li><a href="registIssue.asp">registIssue</a> - ���ݰ�꼭 ��ù���</li>
						<li><a href="register.asp">register</a> - ���ݰ�꼭 ���</li>
						<li><a href="update.asp">update</a> - ���ݰ�꼭 ����</li>
						<li><a href="search.asp">search</a> - ���ݰ�꼭 ��� ��ȸ</li>
						<li><a href="getInfo.asp">getInfo</a> - ���ݰ�꼭 ����/��� ���� Ȯ��</li>
						<li><a href="getInfos.asp">getInfos</a> - �ٷ�(�ִ� 1000��)�� ���ݰ�꼭 ����/��� ���� Ȯ��</li>
						<li><a href="getDetailInfo.asp">getDetailInfo</a> - ���ݰ�꼭 �� ���� Ȯ��</li>
						<li><a href="delete.asp">delete</a> - ���ݰ�꼭 ����</li>
						<li><a href="getLogs.asp">getLogs</a> - ���ݰ�꼭 �����̷� Ȯ��</li>
						<li><a href="attachFile.asp">attachFile</a> - ���ݰ�꼭 ÷������ �߰�</li>
						<li><a href="getFiles.asp">getFiles</a> - ���ݰ�꼭 ÷������ ���Ȯ��</li>
						<li><a href="deleteFile.asp">deleteFile</a> - ���ݰ�꼭 ÷������ 1�� ����</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>ó�� ���μ���</legend>
					<ul>
						<li><a href="send.asp">send</a> - ������/����Ź ���ݰ�꼭 ���࿹�� ó��</li>
						<li><a href="cancelSend.asp">cancelSend</a> - ������/����Ź ���ݰ�꼭 ���࿹�� ��� ó��</li>
						<li><a href="accept.asp">accept</a> - ������/����Ź ���ݰ�꼭 ���࿹���� ���� ���޹޴����� ���� ó��</li>
						<li><a href="deny.asp">deny</a> - ������/����Ź ���ݰ�꼭 ���࿹���� ���� ���޹޴����� �ź� ó��</li>
						<li><a href="issue.asp">issue</a> - ���ݰ�꼭 ���� ó��</li>
						<li><a href="cancelIssue.asp">cancelIssue</a> - ���ݰ�꼭 ������� ó�� (����û ������������ ��� ����)</li>
						<li><a href="request.asp">request</a> - ���ݰ�꼭 ��)�����û ó��.</li>
						<li><a href="cancelRequest.asp">cancelRequest</a> - ���ݰ�꼭 ��)�����û ��� ó��.</li>
						<li><a href="refuse.asp">refuse</a> - ���ݰ�꼭 ��)�����û�� ���� �������� ����ź� ó��.</li>
						<li><a href="sendToNTS.asp">sendToNTS</a> - ����� ���ݰ�꼭�� ����û ������� ��û.</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>�ΰ� ���</legend>
					<ul>
						<li><a href="sendEmail.asp">sendEmail</a> - ó�� ���μ����� ���� �̸��� ������ ��û</li>
						<li><a href="sendSMS.asp">sendSMS</a> - ���࿹��/����/��)�����û �� ���� ���ڸ޽��� �ȳ� ������ ��û.</li>
						<li><a href="sendFAX.asp">sendFAX</a> - ���ݰ�꼭�� ���� �ѽ� ���� ��û..</li>
						<li><a href="attachStatement.asp">attachStatement</a> - ���ڸ��� ÷��</li>
						<li><a href="detachStatement.asp">detachStatement</a> - ���ڸ��� ÷������</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>�˺� ���ݰ�꼭 SSO URL ���</legend>
					<ul>
						<li><a href="getURL.asp">getURL</a> - ���ݰ�꼭 ���� SSO URL Ȯ��</li>
						<li><a href="getPopUpURL.asp">getPopUpURL</a> - �ش� ���ݰ�꼭�� �˺� ȭ���� ǥ���ϴ� URL Ȯ��</li>
						<li><a href="getPrintURL.asp">getPrintURL</a> - �ش� ���ݰ�꼭�� �˺� �μ� ȭ���� ǥ���ϴ� URL Ȯ��</li>
						<li><a href="getMassPrintURL.asp">getMassPrintURL</a> - �ٷ�(�ִ�100��)�� ���ݰ�꼭 �μ� ȭ���� ǥ���ϴ� URL Ȯ��</li>
						<li><a href="getEPrintURL.asp">getEPrintURL</a> - �ش� ���ݰ�꼭�� ���޹޴��ڿ� �˺� �μ� ȭ���� ǥ���ϴ� URL Ȯ��</li>
						<li><a href="getMailURL.asp">getMailURL</a> - �ش� ���ݰ�꼭�� ���۸��ϻ��� "����" ��ư�� �ش��ϴ� URL Ȯ��</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>��Ÿ</legend>
					<ul>
						<li><a href="getUnitCost.asp">getUnitCost</a> - ���ݰ�꼭 ���� �ܰ� Ȯ��</li>
						<li><a href="getCertificateExpireDate.asp">getCertificateExpireDate</a> - ����ȸ���� ����� ������������ �����Ͻ� Ȯ��</li>
						<li><a href="getEmailPublicKeys.asp">getEmailPublicKeys</a> - Email ������ ���� ��뷮 �������� �̸��� ��� Ȯ��</li>
					</ul>
				</fieldset>
			</fieldset>
		 </div>
	</body>
</html>