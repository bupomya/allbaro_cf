package allbaro.cf.dto;

public class CF_dto {

	private String inDate;
	private String outDate;
	private String kinds;
	private String companyNo;
	private String companyName;
	private int inAmount;
	private int outAmount;
	private String code;

	public CF_dto(String inDate, String outDate, String kinds, String companyNo, String companyName, int inAmount,
			int outAmount, String code) {

		super();
		this.inDate = inDate;
		this.outDate = outDate;
		this.kinds = kinds;
		this.companyNo = companyNo;
		this.companyName = companyName;
		this.inAmount = inAmount;
		this.outAmount = outAmount;
		this.code = code;
	}

	public String getInDate() {
		return inDate;
	}

	public void setInDate(String inDate) {
		this.inDate = inDate;
	}

	public String getOutDate() {
		return outDate;
	}

	public void setOutDate(String outDate) {
		this.outDate = outDate;
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
