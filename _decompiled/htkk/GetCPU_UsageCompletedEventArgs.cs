using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;

[DesignerCategory("code")]
[GeneratedCode("wsdl", "2.0.50727.3038")]
[DebuggerStepThrough]
public class GetCPU_UsageCompletedEventArgs : AsyncCompletedEventArgs
{
	private object[] results;

	public float Result
	{
		get
		{
			RaiseExceptionIfNecessary();
			return (float)results[0];
		}
	}

	internal GetCPU_UsageCompletedEventArgs(object[] results, Exception exception, bool cancelled, object userState)
		: base(exception, cancelled, userState)
	{
		this.results = results;
	}
}
