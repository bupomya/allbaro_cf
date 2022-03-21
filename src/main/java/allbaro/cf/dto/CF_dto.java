package allbaro.cf.dto;

public class CF_dto {

	private String date; // 인수/재활용일자
	private String kinds; // 폐기물 종류
	private String kindCode;
	private String companyNo;// 위탁업체 식별번호
	private String companyName;// 위탁업소
	private int inAmount;// 수집량
	private int outAmount;// 폐기물 재활용량
	private String code; // 인계서 일련번호

	public CF_dto(String date, String kinds, String kindCode, String companyNo, String companyName, int inAmount,
			int outAmount, String code) {
		this.date = date;
		this.kinds = kinds;
		this.kindCode = kindCode;
		this.companyNo = companyNo;
		this.companyName = companyName;
		this.inAmount = inAmount;
		this.outAmount = outAmount;
		this.code = code;
	}

	public String getDate() {
		return date;
	}

	public void setDate(String date) {
		this.date = date;
	}

	public String getKindCode() {
		return kindCode;
	}

	public void setKindCode(String kindCode) {
		this.kindCode = kindCode;
	}

	public String getKinds() {
		return kinds;
	}

	public void setKinds(String kinds) {
		this.kinds = kinds;
	}

	public String getCompanyNo() {
		return companyNo;
	}

	public void setCompanyNo(String companyNo) {
		this.companyNo = companyNo;
	}

	public String getCompanyName() {
		return companyName;
	}

	public void setCompanyName(String companyName) {
		this.companyName = companyName;
	}

	public int getInAmount() {
		return inAmount;
	}

	public void setInAmount(int inAmount) {
		this.inAmount = inAmount;
	}

	public int getOutAmount() {
		return outAmount;
	}

	public void setOutAmount(int outAmount) {
		this.outAmount = outAmount;
	}

	public String getCode() {
		return code;
	}

	public void setCode(String code) {
		this.code = code;
	}

}
