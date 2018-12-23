package com.moon.document.opra;

import com.moon.document.opra.word.IWordDocument;

/**
 * MoonOffice文档操作接口
 */
public interface IDocument {
    /**
     * 打开word
     * @param filePath 文件路径
     * @throws Exception
     */
    void open(String filePath)throws Exception;

    /**
     * 保存当前打开的文件
     * @throws Exception
     */
    void save() throws Exception;

    /**
     * 另存为
     * @param filePath 另存为的文件路径
     * @throws Exception
     */
    void saveAs(String filePath)throws Exception;

    /**
     * 获取word文档接口
     * @return
     */
    IWordDocument getWordDocument();

    /**
     * 释放对象
     */
    void destory();
}
