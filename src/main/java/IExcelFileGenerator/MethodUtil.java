package IExcelFileGenerator;

import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.Map;

public class MethodUtil {
	
	public static final Map<String,Method> getterFieldNamesMap = new HashMap<>();
	public static final Map<String,Method> setterFieldNamesMap = new HashMap<>();

	public Method getGetMethod(final Object object,String header) {
		
		Method getMethod = getterFieldNamesMap.get(header);
		
		if(getMethod == null) {
			getMethod = ReflectionUtil.getGetterMethod(object, header);
			getterFieldNamesMap.put(header, getMethod);
		}
		return getMethod;
		
	}
}
