using System;
using System.Reflection;
using TAX.HTKK.Utilities.Core;

namespace TAX.HTKK.Utilities.Entity;

public class DataLocationInfo
{
	public string ColStart { get; set; }

	public string RowStart { get; set; }

	public string ColEnd { get; set; }

	public DataLocationInfo()
	{
		ColStart = string.Empty;
		RowStart = string.Empty;
		ColEnd = string.Empty;
	}

	public DataLocationInfo(string strDataLocation)
	{
		try
		{
			string[] array = strDataLocation.Split(new char[1] { '_' });
			if (array.Length > 1)
			{
				ColStart = array[0];
			}
			string[] array2 = array[1].Split(new char[1] { ':' });
			RowStart = array2[0];
			ColEnd = array2[1];
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
