package dto;

public class W3CData{

	String company;
	String contact;
	String country;
	
	public W3CData() {

	}
	
	public W3CData(String company, String contact, String country) {
		super();
		this.company = company;
		this.contact = contact;
		this.country = country;
	}
	
	public String getCompany() {
		return company;
	}
	public void setCompany(String company) {
		this.company = company;
	}
	public String getContact() {
		return contact;
	}
	public void setContact(String contact) {
		this.contact = contact;
	}
	public String getCountry() {
		return country;
	}
	public void setCountry(String country) {
		this.country = country;
	}
	
	
}
