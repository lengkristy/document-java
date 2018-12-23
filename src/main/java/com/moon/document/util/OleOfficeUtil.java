package com.moon.document.util;

import org.eclipse.swt.internal.ole.win32.IDispatch;
import org.eclipse.swt.internal.ole.win32.IUnknown;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

/**
 * Office Ole工具类
 */
public class OleOfficeUtil {

    /**
     * 反射调用Ole对象的方法
     * @param oleAutomation
     * @param method
     * @param values
     * @return
     */
    public static Variant invokeMethod(OleAutomation oleAutomation,String method,Object... values)throws Exception{
        int[] ids = oleAutomation.getIDsOfNames(new String[]{method});
        if (ids != null && ids.length > 0){
            int iMethod = oleAutomation.getIDsOfNames(new String[]{method})[0];
            Variant[] variants = new Variant[values.length];
            for (int i = 0; i < values.length; i++) {
                Variant variant = createVariant(values[i]);
                variants[i] = variant;
            }
            Variant pVarResult = oleAutomation.invoke(iMethod, variants);
            //release object
            for (Variant variant : variants) {
                variant.dispose();
            }
            return pVarResult;
        }else{
            return null;
        }
    }

    /**
     * 获取Ole对象的属性
     * @param oleAutomation
     * @param propertyName
     * @return
     */
    public static Variant getProperty(OleAutomation oleAutomation,String propertyName){
        int[] ids = oleAutomation.getIDsOfNames(new String[]{propertyName});
        if (ids != null && ids.length > 0) {
            int index = oleAutomation.getIDsOfNames(new String[]{propertyName})[0];
            return oleAutomation.getProperty(index);
        }else{
            return null;
        }
    }


    /**
     * Convert Java Primary type or String to Variant
     *
     * @param value java type data
     * @return Variant
     */
    private static Variant createVariant(Object value) {
        Variant variant = null;

        if (value == null) {
            variant = new Variant();
        } else if (value instanceof Boolean) {
            variant = new Variant((Boolean) value);
        } else if (value instanceof Short) {
            variant = new Variant((Short) value);
        } else if (value instanceof Integer) {
            variant = new Variant((Integer) value);
        } else if (value instanceof Float) {
            variant = new Variant((Float) value);
        } else if (value instanceof String) {
            variant = new Variant(value.toString());
        } else if (value instanceof IDispatch) {
            variant = new Variant((IDispatch) value);
        } else if (value instanceof IUnknown) {
            variant = new Variant((IUnknown) value);
        } else if (value instanceof OleAutomation) {
            variant = new Variant((OleAutomation) value);
        }
        return variant;
    }
}
