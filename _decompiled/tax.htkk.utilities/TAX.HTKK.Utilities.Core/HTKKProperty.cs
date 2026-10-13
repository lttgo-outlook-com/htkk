using System;
using System.Reflection;

namespace TAX.HTKK.Utilities.Core;

public class HTKKProperty
{
	public static object GetPropertyValue(object SourceData, string propName)
	{
		try
		{
			return SourceData.GetType().GetProperty(propName).GetValue(SourceData, null);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}
}
