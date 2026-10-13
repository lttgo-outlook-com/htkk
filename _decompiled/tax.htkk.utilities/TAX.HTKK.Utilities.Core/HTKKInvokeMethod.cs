using System;
using System.Collections.Generic;
using System.Reflection;

namespace TAX.HTKK.Utilities.Core;

public class HTKKInvokeMethod
{
	public static List<string> GetListMethods(string assemblyName, string namespaceName, string className)
	{
		try
		{
			List<string> list = new List<string>();
			Type type = Type.GetType(namespaceName + "." + className + "," + assemblyName);
			MethodInfo[] methods = type.GetMethods(BindingFlags.Static | BindingFlags.Public);
			MethodInfo[] array = methods;
			foreach (MethodInfo methodInfo in array)
			{
				list.Add(methodInfo.Name);
			}
			return list;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static object InvokeStringMethod(string assemblyName, string namespaceName, string className, string methodName, object[] stringParam)
	{
		try
		{
			Type type = Type.GetType(namespaceName + "." + className + "," + assemblyName);
			return type.InvokeMember(methodName, BindingFlags.Static | BindingFlags.Public | BindingFlags.InvokeMethod, null, null, stringParam);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}
}
