package com.moon.document.opra.word.entity;

/**
 * word书签
 */
public class BookMark {

    /**书签名称*/
    private String bookName;

    /**word显示的书签文本*/
    private String bookText;

    public String getBookName() {
        return bookName;
    }

    public void setBookName(String bookName) {
        this.bookName = bookName;
    }

    public String getBookText() {
        return bookText;
    }

    public void setBookText(String bookText) {
        this.bookText = bookText;
    }
}
