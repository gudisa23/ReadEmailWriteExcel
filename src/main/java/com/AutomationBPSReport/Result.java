package com.AutomationBPSReport;

public class Result {
	public String getSourcetype() {
		return sourcetype;
	}
	public void setSourcetype(String sourcetype) {
		this.sourcetype = sourcetype;
	}
	public String getCount() {
		return count;
	}
	public void setCount(String count) {
		this.count = count;
	}
	public String SuccessfullInstalls;
	public String TotallInstalls;
	public String Success;
	public String sourcetype;
	public String count;
	public String getSuccessfullInstalls() {
		return SuccessfullInstalls;
	}
	public void setSuccessfullInstalls(String successfullInstalls) {
		SuccessfullInstalls = successfullInstalls;
	}
	public String getTotallInstalls() {
		return TotallInstalls;
	}
	public void setTotallInstalls(String totallInstalls) {
		TotallInstalls = totallInstalls;
	}
	public String getSuccess() {
		return Success;
	}
	public void setSuccess(String success) {
		Success = success;
	}

}
