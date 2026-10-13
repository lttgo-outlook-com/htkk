using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;

[DebuggerStepThrough]
[GeneratedCode("wsdl", "2.0.50727.3038")]
[DesignerCategory("code")]
public class AdminGetBlockCompletedEventArgs : AsyncCompletedEventArgs
{
	private object[] results;

	public byte[] Result
	{
		get
		{
			RaiseExceptionIfNecessary();
			return (byte[])results[0];
		}
	}

	public string err_msg
	{
		get
		{
			RaiseExceptionIfNecessary();
			return (string)results[1];
		}
	}

	internal AdminGetBlockCompletedEventArgs(object[] results, Exception exception, bool cancelled, object userState)
		: base(exception, cancelled, userState)
	{
		this.results = results;
	}
}
