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
					<legend>회원 정보</legend>
					<ul>					
						<li><a href="checkIsMember.asp">checkIsMember</a> - 연동회원 가입여부</li>
						<li><a href="checkID.asp">checkID</a> - 아이디 중복확인</li>
						<li><a href="joinMember.asp">joinMember</a> - 연동회원 가입 요청</li>
						<li><a href="getChargeInfo.asp">getChargeInfo</a> - 과금정보 확인</li>
						<li><a href="getBalance.asp">getBalance</a> - 연동회원 잔여포인트 확인</li>
						<li><a href="getPopbillURL.asp">getPopbillURL</a> - 팝빌 SSO URL 요청 (로그인/포인트충전/공인인증서 등록)</li>
						<li><a href="getPartnerBalance.asp">getPartnerBalance</a> - 파트너 잔여포인트 확인</li>
						<li><a href="getPartnerURL.asp">getPartnerURL</a> - 파트너 포인트 충전 URL 확인</li>
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
						<li><a href="checkMgtKeyInUse.asp">checkMgtKeyInUse</a> - 문서관리번호 사용여부 확인</li>
						<li><a href="registIssue.asp">registIssue</a> - 세금계산서 즉시발행</li>
						<li><a href="register.asp">register</a> - 세금계산서 임시저장</li>
						<li><a href="update.asp">update</a> - 세금계산서 수정</li>
						<li><a href="search.asp">search</a> - 세금계산서 목록 조회</li>
						<li><a href="getInfo.asp">getInfo</a> - 세금계산서 상태/요약 정보 확인</li>
						<li><a href="getInfos.asp">getInfos</a> - 세금계산서 상태/요약 정보 확인 - 대량</li>
						<li><a href="getDetailInfo.asp">getDetailInfo</a> - 세금계산서 상세 정보 확인</li>
						<li><a href="delete.asp">delete</a> - 세금계산서 삭제</li>
						<li><a href="getLogs.asp">getLogs</a> - 세금계산서 상태정보 변경이력 확인</li>
						<li><a href="attachFile.asp">attachFile</a> - 첨부파일 추가</li>
						<li><a href="getFiles.asp">getFiles</a> - 세금계산서 첨부파일 목록확인</li>
						<li><a href="deleteFile.asp">deleteFile</a> - 첨부파일 삭제</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>처리 프로세스</legend>
					<ul>
						<li><a href="send.asp">send</a> - 세금계산서 발행예정</li>
						<li><a href="cancelSend.asp">cancelSend</a> - 세금계산서 발행예정 취소</li>
						<li><a href="accept.asp">accept</a> - 세금계산서 발행예정 승인</li>
						<li><a href="deny.asp">deny</a> - 세금계산서 발행예정 거부 </li>
						<li><a href="issue.asp">issue</a> - 세금계산서 발행</li>
						<li><a href="cancelIssue.asp">cancelIssue</a> - 세금계산서 발행취소</li>
						<li><a href="request.asp">request</a> - 세금계산서 역발행요청</li>
						<li><a href="cancelRequest.asp">cancelRequest</a> - 세금계산서 역발행요청 취소</li>
						<li><a href="refuse.asp">refuse</a> - 세금계산서 역발행요청 거부</li>
						<li><a href="sendToNTS.asp">sendToNTS</a> - 세금계산서 국세청 즉시전송</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>부가 기능</legend>
					<ul>
						<li><a href="sendEmail.asp">sendEmail</a> - 발행 안내메일 전송</li>
						<li><a href="sendSMS.asp">sendSMS</a> - 안내문자 메시지 전송</li>
						<li><a href="sendFAX.asp">sendFAX</a> - 세금계산서 팩스 전송</li>
						<li><a href="attachStatement.asp">attachStatement</a> - 전자명세서 첨부</li>
						<li><a href="detachStatement.asp">detachStatement</a> - 전자명세서 첨부해제</li>
						<li><a href="assignMgtKey.asp">assignMgtKey</a> - 관리번호 할당 </li>
						<li><a href="listEmailConfig.asp">listEmailConfig</a> - 세금계산서 알림메일 전송목록 조회 </li>
						<li><a href="updateEmailConfig.asp">updateEmailConfig</a> - 세금계산서 알림메일 전송 설정 수정 </li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>팝빌 세금계산서 SSO URL 기능</legend>
					<ul>
						<li><a href="getURL.asp">getURL</a> - 세금계산서 관련 SSO URL 확인</li>
						<li><a href="getPopUpURL.asp">getPopUpURL</a> - 세금계산서 보기 팝업 URL</li>
						<li><a href="getPrintURL.asp">getPrintURL</a> - 세금계산서 인쇄 팝업 URL</li>
						<li><a href="getMassPrintURL.asp">getMassPrintURL</a> - 세금계산서 인쇄 팝업 URL - 대량</li>
						<li><a href="getEPrintURL.asp">getEPrintURL</a> - 세금계산서 인쇄 팝업 URL - 공급받는자용</li>
						<li><a href="getMailURL.asp">getMailURL</a> - 세금계산서 메일링크 URL</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>기타</legend>
					<ul>
						<li><a href="getUnitCost.asp">getUnitCost</a> - 세금계산서 발행단가 확인</li>
						<li><a href="getCertificateExpireDate.asp">getCertificateExpireDate</a> - 공인인증서 만료일시 확인</li>
						<li><a href="checkCertValidation.asp">checkCertValidation</a> - 공인인증서 유효성 확인</li>
						<li><a href="getEmailPublicKeys.asp">getEmailPublicKeys</a> - 대용량 연계사업자 이메일 목록 확인</li>
					</ul>
				</fieldset>
			</fieldset>
		 </div>
	</body>
</html>