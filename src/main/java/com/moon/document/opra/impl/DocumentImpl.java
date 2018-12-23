package com.moon.document.opra.impl;

import com.moon.document.opra.IDocument;
import com.moon.document.opra.word.IWordDocument;
import com.moon.document.opra.word.impl.WordDocumentImpl;
import com.moon.document.util.OleOfficeUtil;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleClientSite;

/**
 * 文档接口实现
 */
public class DocumentImpl implements IDocument{

    /**Ole容器组件*/
    private OleClientSite oleClientSite;

    /**当前Ole文档加载器*/
    private OleAutomation oleMoonOfficeLoader;

    /**word文档接口*/
    private IWordDocument wordDocument;

    private DocumentImpl(OleClientSite oleClientSite){
        this.oleClientSite = oleClientSite;
        this.oleMoonOfficeLoader = new OleAutomation(this.oleClientSite);
        //创建word文档
        this.wordDocument = new WordDocumentImpl(this.oleMoonOfficeLoader);
    }

    public synchronized static DocumentImpl createInstance(OleClientSite oleClientSite)throws Exception{
        return new DocumentImpl(oleClientSite);
    }

    public void open(String filePath) throws Exception {
        OleOfficeUtil.invokeMethod(this.oleMoonOfficeLoader,"Open",filePath);
    }

    public void save() throws Exception {
        OleOfficeUtil.invokeMethod(this.oleMoonOfficeLoader,"Save");
    }

    public void saveAs(String filePath) throws Exception {
        OleOfficeUtil.invokeMethod(this.oleMoonOfficeLoader,"Save",filePath,null,null,null);
    }

    public IWordDocument getWordDocument() {
        return this.wordDocument;
    }

    public void destory() {

    }
}
