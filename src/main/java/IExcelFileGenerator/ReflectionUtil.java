package IExcelFileGenerator;

import java.lang.reflect.Method;

import IExcelFileGeneratorExceptions.IExcelFileGeneratorException;

public final class ReflectionUtil {
	
	public static final String GET_PREFIX = "get";
	public static final String IS_PREFIX = "is";
	
	
	private ReflectionUtil() {
		
	}
	
	public static Method getGetterMethod(final Object object,String methodName) {
		
		String getterFieldName = getGetterFieldName(methodName,GET_PREFIX);
		
		Method getterMethod = getActualMethod(object,getterFieldName);
		
		if(getterMethod == null) {
			
			getterFieldName = getGetterFieldName(methodName,IS_PREFIX);
			
			getterMethod = getActualMethod(object,getterFieldName);
			
		}
		
		if(getterMethod == null) {
			throw new IExcelFileGeneratorException("cannot find getter field");
		}
		return getterMethod;
		
	}
	
	private static Method getActualMethod(final Object object,String getterFieldName) {
		
		for(Method method:object.getClass().getMethods()) {
			if(method.getName().equalsIgnoreCase(getterFieldName) && method.getParameterTypes().length == 0) {
				return method;
			}
		}
		return null;
		
	}
	private static String getGetterFieldName(String methodName,String prefix) {
		return prefix + methodName.substring(0,1).toUpperCase() + methodName.substring(1);
		
	}

}
