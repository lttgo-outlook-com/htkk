using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;

[DebuggerStepThrough]
[DesignerCategory("code")]
[GeneratedCode("wsdl", "2.0.50727.3038")]
public class GetBlockCompletedEventArgs : AsyncCompletedEventArgs
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

	public byte[] block_signature
	{
		get
		{
			RaiseExceptionIfNecessary();
			return (byte[])results[1];
		}
	}

	public string err_msg
	{
		get
		{
			RaiseExceptionIfNecessary();
			return (string)results[2];
		}
	}

	internal GetBlockCompletedEventArgs(object[] results, Exception exception, bool cancelled, object userState)
		: base(exception, cancelled, userState)
	{
		this.results = results;
	}
}
