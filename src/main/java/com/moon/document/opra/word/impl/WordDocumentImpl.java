package com.moon.document.opra.word.impl;

import com.moon.document.util.OleOfficeUtil;
import com.moon.document.opra.word.IWordDocument;
import com.moon.document.opra.word.entity.BookMark;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import java.util.ArrayList;
import java.util.List;

/**
 * 处理word的文档对象
 */
public class WordDocumentImpl implements IWordDocument{

    /**当前word文档对象*/
    private OleAutomation oleMoonOfficeLoader;

    public WordDocumentImpl(OleAutomation oleMoonOfficeLoader){
        this.oleMoonOfficeLoader = oleMoonOfficeLoader;
    }

    public String getCurrentSelectedText() throws Exception {
        OleAutomation application = getAppOle();
        OleAutomation selection = OleOfficeUtil.getProperty(application, "Selection").getAutomation();
        String range = OleOfficeUtil.getProperty(selection, "Range").getString();
        return range;
    }

    public void insertBookMark(BookMark bookMark) throws Exception {
        OleAutomation bookMarks = getBookMarkOle();
        OleAutomation application = OleOfficeUtil.getProperty(getWordsOle(), "Application").getAutomation();
        OleAutomation selection = OleOfficeUtil.getProperty(application, "Selection").getAutomation();
        String range = OleOfficeUtil.getProperty(selection, "Range").getString();
        OleOfficeUtil.invokeMethod(bookMarks, "Add",bookMark.getBookName());
    }

    public List<BookMark> getBookMarkList() throws Exception {
        List<BookMark> bookMarkList = new ArrayList<BookMark>();
        OleAutomation bookMarks = getBookMarkOle();
        int bookmarksCount = OleOfficeUtil.getProperty(bookMarks, "Count").getInt();
        for (int i = 0;i < bookmarksCount; i++){
            OleAutomation bookmark = OleOfficeUtil.invokeMethod(bookMarks, "Item", i + 1).getAutomation();
            Variant name = OleOfficeUtil.getProperty(bookmark, "Name");
            Variant range = OleOfficeUtil.getProperty(bookmark, "Range");
            BookMark bk = new BookMark();
            bk.setBookName(name.getString());
            bk.setBookText(range.getString());
            bookMarkList.add(bk);
        }
        return bookMarkList;
    }

    public void gotoBookMark(BookMark bookMark) throws Exception {
        OleAutomation bookMarks = getBookMarkOle();
        int bookmarksCount = OleOfficeUtil.getProperty(bookMarks, "Count").getInt();
        for (int i = 0;i < bookmarksCount; i++){
            OleAutomation bookmark = OleOfficeUtil.invokeMethod(bookMarks, "Item", i + 1).getAutomation();
            Variant name = OleOfficeUtil.getProperty(bookmark, "Name");
            if (bookMark.getBookName().equals(name.getString())){//找到对应的书签
                OleAutomation range = OleOfficeUtil.getProperty(bookmark, "Range").getAutomation();
                OleOfficeUtil.invokeMethod(range,"Select");
                break;
            }
        }
    }

    public void destory() {

    }

    //以下为私有方法

    /**
     * 获取words的OLE对象
     * @return
     */
    private OleAutomation getWordsOle(){
        OleAutomation activeDocument = OleOfficeUtil.getProperty(this.oleMoonOfficeLoader, "ActiveDocument").getAutomation();
        return OleOfficeUtil.getProperty(activeDocument, "Words").getAutomation();//查找word
    }

    /**
     * 获取words应用的OLE对象
     * @return
     */
    private OleAutomation getAppOle(){
        return OleOfficeUtil.getProperty(this.getWordsOle(), "Application").getAutomation();
    }

    /**
     * 获取书签的OLE对象
     * @return
     */
    private OleAutomation getBookMarkOle(){
        OleAutomation activeDocument = OleOfficeUtil.getProperty(this.oleMoonOfficeLoader, "ActiveDocument").getAutomation();
        return OleOfficeUtil.getProperty(activeDocument, "BookMarks").getAutomation();
    }
}
