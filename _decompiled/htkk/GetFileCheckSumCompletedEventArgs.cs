using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;

[DesignerCategory("code")]
[GeneratedCode("wsdl", "2.0.50727.3038")]
[DebuggerStepThrough]
public class GetFileCheckSumCompletedEventArgs : AsyncCompletedEventArgs
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

	internal GetFileCheckSumCompletedEventArgs(object[] results, Exception exception, bool cancelled, object userState)
		: base(exception, cancelled, userState)
	{
		this.results = results;
	}
}
