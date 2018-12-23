package com.moon.document.window;

import com.moon.document.component.MoonOfficeComponent;
import com.moon.document.opra.word.entity.BookMark;
import org.eclipse.jface.viewers.TableViewer;
import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.events.SelectionListener;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.ole.win32.OLE;
import org.eclipse.swt.ole.win32.OleFrame;
import org.eclipse.swt.widgets.*;

/**
 * 主窗体
 */
public class MainWindow {

    /**
     * nmb
     */
    private Display display;
    /**
     * nimb
     */
    private Shell shell;

    /**
     * table
     */
    private Table table;
    /**
     * url
     */
    private String url;

    /**
     * gridData column count
     */
    private final int GRID_COL = 5;
    /**
     * doc frame vertical span count
     */
    private final int DOC_VERTICAL_SPAN = 3;
    /**
     * table horizontal span count
     */
    private final int TABLE_HORIZONTAL_SPAN = 4;

    /**
     * 初始化资源
     * @throws Exception
     */
    public void init() throws Exception{
        display = Display.getDefault();
        shell = new Shell(display, SWT.CLOSE | SWT.MIN);
        shell.setText("文书模板编辑器");
        GridLayout gridLayout = new GridLayout();
        gridLayout.numColumns = GRID_COL;
        shell.setLayout(gridLayout);

        Composite frame = new OleFrame(shell, SWT.NONE);
        GridData gridData = new GridData(GridData.FILL_BOTH);
        gridData.verticalSpan = DOC_VERTICAL_SPAN;
        frame.setSize(display.getClientArea().width - 400, display.getClientArea().height - 50);
        frame.setLayoutData(gridData);

        final MoonOfficeComponent moonOfficeComponent = new MoonOfficeComponent(frame, SWT.NONE, "DSOFramer.FramerControl");
        moonOfficeComponent.setSize(display.getClientArea().width - 400, display.getClientArea().height - 50);
        moonOfficeComponent.doVerb(OLE.OLEIVERB_SHOW);
        moonOfficeComponent.initResource();


        Button create = new Button(shell, SWT.PUSH);
        create.setText("打開文檔");
        create.setToolTipText("Open Document");
        gridData = new GridData(GridData.HORIZONTAL_ALIGN_FILL | GridData.VERTICAL_ALIGN_BEGINNING);
        create.setLayoutData(gridData);
        create.addSelectionListener(new SelectionAdapter() {
            public void widgetSelected(SelectionEvent event) {
                //open doc
                FileDialog fileselect = new FileDialog(shell, SWT.SINGLE);
                fileselect.setFilterNames(new String[]{"Microsoft Word"});
                fileselect.setFilterExtensions(new String[]{"*.doc;*.docx"});
                url = fileselect.open();
                try {
                    moonOfficeComponent.open(url);
                    //加载书签
                    loadBookList(moonOfficeComponent,table);
                }catch (Exception e){

                }
            }
        });

        //添加书签按钮
        Button edit = new Button(shell, SWT.PUSH);
        edit.setText("添加書籤");
        edit.setToolTipText("Add Bookmark");
        gridData = new GridData(GridData.VERTICAL_ALIGN_BEGINNING);
        edit.setLayoutData(gridData);
        edit.addSelectionListener(new SelectionAdapter() {
            public void widgetSelected(SelectionEvent event) {
                try {
                    AddBookMark addBookMark = new AddBookMark(shell, moonOfficeComponent.getCurrentWordSelectedText(), "");
                    int ret = addBookMark.open();

                }catch (Exception e){

                }
            }
        });


        Button enter = new Button(shell, SWT.PUSH);
        enter.setText("保存");
        enter.setToolTipText("Save Document");
        gridData = new GridData(GridData.VERTICAL_ALIGN_BEGINNING); //HORIZONTAL_ALIGN_END
        enter.setLayoutData(gridData);
        enter.addSelectionListener(new SelectionAdapter() {
            public void widgetSelected(SelectionEvent event) {
            }
        });


        Button temp = new Button(shell, SWT.PUSH);
        temp.setText("另存为");
        temp.setToolTipText("另存为");
        gridData = new GridData(GridData.VERTICAL_ALIGN_BEGINNING);
        temp.setLayoutData(gridData);
        temp.addSelectionListener(new SelectionAdapter() {
            public void widgetSelected(SelectionEvent event) {

            }
        });

        gridData = new GridData();
        gridData.horizontalSpan = TABLE_HORIZONTAL_SPAN;
        gridData.horizontalAlignment = SWT.FILL;
        gridData.grabExcessHorizontalSpace = true;
        gridData.grabExcessVerticalSpace = true;
        gridData.verticalAlignment = SWT.FILL;
        //SWT.NULTI代表可以选择多行,SWT.FULL_SELECTION代表可以整行选择,SWT.BORDER边框，SWT.V_SCROLL ,SWT.H_SCROLL滚动条
        TableViewer tableViewer = new TableViewer(shell, SWT.BORDER | SWT.MULTI | SWT.FULL_SELECTION | SWT.BORDER | SWT.V_SCROLL | SWT.H_SCROLL);
        table = tableViewer.getTable();
        table.setLayoutData(gridData);
        table.setLinesVisible(true);
        table.setHeaderVisible(false);

        table.addSelectionListener(new SelectionListener() {
            public void widgetSelected(SelectionEvent selectionEvent) {
                try {
                    //书签跳转
                    TableItem item = (TableItem) selectionEvent.item;
                    String bkName = item.getText(0);
                    BookMark bookMark = new BookMark();
                    bookMark.setBookName(bkName);
                    moonOfficeComponent.gotoBookMark(bookMark);
                }catch (Exception e){

                }
            }

            public void widgetDefaultSelected(SelectionEvent selectionEvent) {
                int l = 0;
            }
        });
    }


    /**
     * 加载表格
     * @param moonOfficeComponent
     * @param table
     */
    private static void loadBookList(final MoonOfficeComponent moonOfficeComponent,Table table){
        try {
            table.removeAll();
            TableItem item;
            java.util.List<BookMark> bookMarkList = moonOfficeComponent.getBookMarks();
            for (BookMark bookMark:bookMarkList) {
                item = new TableItem(table, SWT.NONE);
                item.setText(new String[]{bookMark.getBookName(), bookMark.getBookText()});
            }
        }catch (Exception e){

        }
    }

    /**
     * 窗体运行
     */
    public void run(String[] args)throws Exception{
        init();

        shell.pack();
        shell.open();
        while (!shell.isDisposed()) {
            if (!display.readAndDispatch()) {
                display.sleep();
            }
        }
    }
}
