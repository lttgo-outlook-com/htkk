using System;
using TAX.HTKK.Utilities.Core;

namespace TAX.HTKK.Utilities.Entity;

public class HTKKException : Exception
{
	public HTKKException()
	{
	}

	public HTKKException(string fnName, Exception inner)
		: base(fnName, inner)
	{
		string message = string.Format("{0} - {1}: {2}." + Environment.NewLine, DateTime.Now.ToString("dd/MM/yyyy"), fnName, inner.StackTrace);
		if (HTKKConfig.ENABLE_LOGGING.Equals("1"))
		{
			HTKKFile.SimpleLogException(message);
		}
	}
}
