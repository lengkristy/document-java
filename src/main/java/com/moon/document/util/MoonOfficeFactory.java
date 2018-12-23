package com.moon.document.util;

import com.moon.document.opra.IDocument;
import com.moon.document.opra.impl.DocumentImpl;
import org.eclipse.swt.ole.win32.OleClientSite;

/**
 * office工厂
 */
public class MoonOfficeFactory {

    /**
     * 创建文档对象
     * @param oleClientSite
     * @return
     */
    public static IDocument createDocument(OleClientSite oleClientSite) throws Exception{
        return DocumentImpl.createInstance(oleClientSite);
    }

}
