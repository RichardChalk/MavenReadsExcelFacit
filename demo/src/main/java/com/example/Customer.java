package com.example;

public class Customer {
  private int customerID;
  private String firstName;
  private String lastName;
  private String phoneNumber;
  private String address;
  private String payingCustomer;
  private String doNotContact;

  // Konstruktor
  public Customer(int customerID, String firstName, String lastName,
      String phoneNumber, String address, String payingCustomer,
      String doNotContact) {
    this.customerID = customerID;
    this.firstName = firstName;
    this.lastName = lastName;
    this.phoneNumber = phoneNumber;
    this.address = address;
    this.payingCustomer = payingCustomer;
    this.doNotContact = doNotContact;
  }

  // Getters och Setters
  public int getCustomerID() {
    return customerID;
  }

  public void setCustomerID(int customerID) {
    this.customerID = customerID;
  }

  public String getFirstName() {
    return firstName;
  }

  public void setFirstName(String firstName) {
    this.firstName = firstName;
  }

  public String getLastName() {
    return lastName;
  }

  public void setLastName(String lastName) {
    this.lastName = lastName;
  }

  public String getPhoneNumber() {
    return phoneNumber;
  }

  public void setPhoneNumber(String phoneNumber) {
    this.phoneNumber = phoneNumber;
  }

  public String getAddress() {
    return address;
  }

  public void setAddress(String address) {
    this.address = address;
  }

  public String getPayingCustomer() {
    return payingCustomer;
  }

  public void setPayingCustomer(String payingCustomer) {
    this.payingCustomer = payingCustomer;
  }

  public String getDoNotContact() {
    return doNotContact;
  }

  public void setDoNotContact(String doNotContact) {
    this.doNotContact = doNotContact;
  }

  // toString-metod för att skriva ut objektet lätt
  @Override
  public String toString() {
    return "Customer{" +
        "customerID=" + customerID +
        ", firstName='" + firstName + '\'' +
        ", lastName='" + lastName + '\'' +
        ", phoneNumber='" + phoneNumber + '\'' +
        ", address='" + address + '\'' +
        ", payingCustomer='" + payingCustomer + '\'' +
        ", doNotContact='" + doNotContact + '\'' +
        '}';
  }
}
