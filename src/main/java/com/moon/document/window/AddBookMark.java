package com.moon.document.window;

import org.eclipse.jface.dialogs.Dialog;
import org.eclipse.jface.dialogs.IDialogConstants;
import org.eclipse.swt.SWT;
import org.eclipse.swt.graphics.Point;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.*;

/**
 * 添加书签窗体
 */
public class AddBookMark extends Dialog {

    private Text bkContentText;
    private Text bkNameText;

    private String bkContent;

    private String bkName;

    protected AddBookMark(Shell parentShell,String bkContent,String bkName) {
        super(parentShell);
        this.bkContent = bkContent;
        this.bkName = bkName;
    }

    @Override
    protected Control createDialogArea(Composite parent) {
        Composite comp = new Composite(parent, SWT.None);
        comp.setLayout(new GridLayout());
        new Label(comp, SWT.None).setText("书签内容");
        bkContentText = new Text(comp, SWT.None);
        bkContentText.setText(bkContent);
        bkContentText.setEnabled(false);
        bkContentText.setSize(200,50);
        new Label(comp, SWT.None).setText("书签名称");
        bkNameText = new Text(comp, SWT.None);
        return super.createDialogArea(parent);
    }

    @Override
    protected Button createButton(Composite parent, int id, String label,
                                  boolean defaultButton) {
        // return super.createButton(parent, id, label, defaultButton);
        return null;// 重写父类的按钮，返回null，使默认按钮失效
    }

    /**
     * 改写默认的OK、cancel按钮
     */
    @Override
    protected void initializeBounds() {
        Composite compo = (Composite) getButtonBar();
        super.createButton(compo, IDialogConstants.OK_ID, "OK", false);
        super.createButton(compo, IDialogConstants.NO_ID, "NO", false);
        super.initializeBounds();
    }

    @Override
    protected Point getInitialSize() {
        return new Point(300, 200);// 重新设置大小
    }

    @Override
    protected Button getButton(int id) {
        // TODO 自动生成的方法存根
        return super.getButton(id);
    }

    public String getBkContent() {
        return bkContent;
    }

    public void setBkContent(String bkContent) {
        this.bkContent = bkContent;
    }

    public String getBkName() {
        return bkName;
    }

    public void setBkName(String bkName) {
        this.bkName = bkName;
    }
}
