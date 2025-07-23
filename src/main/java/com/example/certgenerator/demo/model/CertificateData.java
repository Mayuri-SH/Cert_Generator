package com.example.certgenerator.demo.model;

public class CertificateData {
    private String name;
    private String activity;
    private String date;

    public CertificateData() {
    }

    public CertificateData(String name, String activity, String date) {
        this.name = name;
        this.activity = activity;
        this.date = date;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getActivity() {
        return activity;
    }

    public void setActivity(String activity) {
        this.activity = activity;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }
}
