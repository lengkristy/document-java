package com.moon.document.opra.word;

import com.moon.document.opra.word.entity.BookMark;

import java.util.List;

/**
 * 处理word接口
 */
public interface IWordDocument {
    /**
     * 获取word当前选中的内容
     * @return 返回选中的内容
     * @throws Exception
     */
    String getCurrentSelectedText()throws Exception;

    /**
     * 插入书签
     * @param bookMark
     * @throws Exception
     */
    void insertBookMark(BookMark bookMark)throws Exception;

    /**
     * 获取当前文档的书签列表
     * @return
     * @throws Exception
     */
    List<BookMark> getBookMarkList()throws Exception;

    /**
     * 转到书签
     * @param bookMark
     * @throws Exception
     */
    void gotoBookMark(BookMark bookMark)throws Exception;

}
