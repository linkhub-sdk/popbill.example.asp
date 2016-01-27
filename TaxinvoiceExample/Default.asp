<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		
		<title>팝빌 SDK ASP Example.</title>
	</head>

	<body>

		<div id="content">

			<p class="heading1">팝빌 세금계산서 SDK ASP Example.</p>
			
			<br/>

			<fieldset class="fieldset1">
				<legend>팝빌 기본 API</legend>

				<fieldset class="fieldset2">
					<legend>회원사 정보</legend>
					<ul>					
						<li><a href="checkIsMember.asp">checkCorpIsMember</a> - 연동회원사 가입 여부 확인</li>
						<li><a href="checkID.asp">checkID</a> - 아이디 중복확인</li>
						<li><a href="joinMember.asp">joinMember</a> - 연동회원사 가입 요청</li>
						<li><a href="getBalance.asp">getBalance</a> - 연동회원사 잔여포인트 확인</li>
						<li><a href="getPartnerBalance.asp">getPartnerBalance</a> - 파트너 잔여포인트 확인</li>
						<li><a href="getPopbillURL.asp">getPopbillURL</a> - 팝빌 SSO URL 요청</li>
						<li><a href="listContact.asp">listContact</a> - 담당자 목록 조회</li>
						<li><a href="updateContact.asp">updateContact</a> - 담당자 정보 수정</li>
						<li><a href="registContact.asp">registContact</a> - 담당자 추가</li>
						<li><a href="updateCorpInfo.asp">updateCorpInfo</a> - 회사정보 수정</li>
						<li><a href="getCorpInfo.asp">getCorpInfo</a> - 회사정보 확인</li>
					</ul>
				</fieldset>

			</fieldset>
			
			<br />
			
			<fieldset class="fieldset1">
				<legend>전자세금계산서 관련 API</legend>
				
				<fieldset class="fieldset2">
					<legend>등록/수정/확인/삭제</legend>
					<ul>
						<li><a href="checkMgtKeyInUse.asp">checkMgtKeyInUse</a> - 연동관리번호의 등록/사용여부 확인</li>
						<li><a href="registIssue.asp">registIssue</a> - 세금계산서 즉시발행</li>
						<li><a href="register.asp">register</a> - 세금계산서 등록</li>
						<li><a href="update.asp">update</a> - 세금계산서 수정</li>
						<li><a href="search.asp">search</a> - 세금계산서 목록 조회</li>
						<li><a href="getInfo.asp">getInfo</a> - 세금계산서 상태/요약 정보 확인</li>
						<li><a href="getInfos.asp">getInfos</a> - 다량(최대 1000건)의 세금계산서 상태/요약 정보 확인</li>
						<li><a href="getDetailInfo.asp">getDetailInfo</a> - 세금계산서 상세 정보 확인</li>
						<li><a href="delete.asp">delete</a> - 세금계산서 삭제</li>
						<li><a href="getLogs.asp">getLogs</a> - 세금계산서 문서이력 확인</li>
						<li><a href="attachFile.asp">attachFile</a> - 세금계산서 첨부파일 추가</li>
						<li><a href="getFiles.asp">getFiles</a> - 세금계산서 첨부파일 목록확인</li>
						<li><a href="deleteFile.asp">deleteFile</a> - 세금계산서 첨부파일 1개 삭제</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>처리 프로세스</legend>
					<ul>
						<li><a href="send.asp">send</a> - 정발행/위수탁 세금계산서 발행예정 처리</li>
						<li><a href="cancelSend.asp">cancelSend</a> - 정발행/위수탁 세금계산서 발행예정 취소 처리</li>
						<li><a href="accept.asp">accept</a> - 정발행/위수탁 세금계산서 발행예정에 대한 공급받는자의 승인 처리</li>
						<li><a href="deny.asp">deny</a> - 정발행/위수탁 세금계산서 발행예정에 대한 공급받는자의 거부 처리</li>
						<li><a href="issue.asp">issue</a> - 세금계산서 발행 처리</li>
						<li><a href="cancelIssue.asp">cancelIssue</a> - 세금계산서 발행취소 처리 (국세청 전송전까지만 취소 가능)</li>
						<li><a href="request.asp">request</a> - 세금계산서 역)발행요청 처리.</li>
						<li><a href="cancelRequest.asp">cancelRequest</a> - 세금계산서 역)발행요청 취소 처리.</li>
						<li><a href="refuse.asp">refuse</a> - 세금계산서 역)발행요청에 대한 공급자의 발행거부 처리.</li>
						<li><a href="sendToNTS.asp">sendToNTS</a> - 발행된 세금계산서의 국세청 즉시전송 요청.</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>부가 기능</legend>
					<ul>
						<li><a href="sendEmail.asp">sendEmail</a> - 처리 프로세스에 대한 이메일 재전송 요청</li>
						<li><a href="sendSMS.asp">sendSMS</a> - 발행예정/발행/역)발행요청 에 대한 문자메시지 안내 재전송 요청.</li>
						<li><a href="sendFAX.asp">sendFAX</a> - 세금계산서에 대한 팩스 전송 요청..</li>
						<li><a href="attachStatement.asp">attachStatement</a> - 전자명세서 첨부</li>
						<li><a href="detachStatement.asp">detachStatement</a> - 전자명세서 첨부해제</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>팝빌 세금계산서 SSO URL 기능</legend>
					<ul>
						<li><a href="getURL.asp">getURL</a> - 세금계산서 관련 SSO URL 확인</li>
						<li><a href="getPopUpURL.asp">getPopUpURL</a> - 해당 세금계산서의 팝빌 화면을 표시하는 URL 확인</li>
						<li><a href="getPrintURL.asp">getPrintURL</a> - 해당 세금계산서의 팝빌 인쇄 화면을 표시하는 URL 확인</li>
						<li><a href="getMassPrintURL.asp">getMassPrintURL</a> - 다량(최대100건)의 세금계산서 인쇄 화면을 표시하는 URL 확인</li>
						<li><a href="getEPrintURL.asp">getEPrintURL</a> - 해당 세금계산서의 공급받는자용 팝빌 인쇄 화면을 표시하는 URL 확인</li>
						<li><a href="getMailURL.asp">getMailURL</a> - 해당 세금계산서의 전송메일상의 "보기" 버튼에 해당하는 URL 확인</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>기타</legend>
					<ul>
						<li><a href="getUnitCost.asp">getUnitCost</a> - 세금계산서 발행 단가 확인</li>
						<li><a href="getCertificateExpireDate.asp">getCertificateExpireDate</a> - 연동회원이 등록한 공인인증서의 만료일시 확인</li>
						<li><a href="getEmailPublicKeys.asp">getEmailPublicKeys</a> - Email 유통을 위한 대용량 연계사업자 이메일 목록 확인</li>
					</ul>
				</fieldset>
			</fieldset>
		 </div>
	</body>
</html>