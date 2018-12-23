package com.moon.document.component;

import com.moon.document.opra.IDocument;
import com.moon.document.util.MoonOfficeFactory;
import com.moon.document.opra.word.entity.BookMark;
import org.eclipse.swt.ole.win32.OleClientSite;
import org.eclipse.swt.widgets.Composite;

import java.util.List;

/**
 * MoonOffice组件，用于显示操作Office
 * @author 代浩然
 * @version 1.0
 * @time 2018-12-13
 */
public class MoonOfficeComponent extends OleClientSite {

    /**
     * office操作接口
     */
    private IDocument document;

    /**
     * 构造函数
     * @param parent 父组件
     * @param style
     * @param progId ocx注册id
     */
    public MoonOfficeComponent(Composite parent, int style, String progId){
        super(parent, style,progId);
    }

    /**
     * 初始化资源
     */
    public void initResource()throws Exception{
        this.document = MoonOfficeFactory.createDocument(this);
    }

    /**
     * 打开文件
     * @param filePath
     */
    public void open(String filePath)throws Exception{
        this.document.open(filePath);
    }

    /**
     * 保存当前打开的文档
     * @throws Exception
     */
    public void save()throws Exception{
        this.document.save();
    }

    /**
     * 另存为
     * @param filePath 文件路径
     * @throws Exception
     */
    public void saveAs(String filePath)throws Exception{
        this.document.saveAs(filePath);
    }

    /**
     * 返回当前word选中的内容
     * @return
     * @throws Exception
     */
    public String getCurrentWordSelectedText()throws Exception{
        return this.document.getWordDocument().getCurrentSelectedText();
    }

    /**
     * 插入书签
     * @param bookMark 书签名称
     * @return
     * @throws Exception
     */
    public void insertBookMark(BookMark bookMark)throws Exception{
        this.document.getWordDocument().insertBookMark(bookMark);
    }

    /**
     * 获取当前word所有的书签
     * @return
     * @throws Exception
     */
    public List<BookMark> getBookMarks()throws Exception{
        return this.document.getWordDocument().getBookMarkList();
    }

    /**
     * 转到书签，并且选中
     * @param bookMark
     * @throws Exception
     */
    public void gotoBookMark(BookMark bookMark)throws Exception{
        this.document.getWordDocument().gotoBookMark(bookMark);
    }

    /**
     * 释放资源
     */
    public void destory(){
        this.document.destory();
    }

}
