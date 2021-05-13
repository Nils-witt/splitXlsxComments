package de.nilswitt.oks.spltXlsxComments.model;

public class Comment {
    private int id;
    private String author;
    private String forUser;
    private String value;

    public Comment(int id, String author, String forUser, String value) {
        this.id = id;
        this.author = author;
        this.forUser = forUser;
        this.value = value;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getAuthor() {
        return author;
    }

    public void setAuthor(String author) {
        this.author = author;
    }

    public String getForUser() {
        return forUser;
    }

    public void setForUser(String forUser) {
        this.forUser = forUser;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }
}
