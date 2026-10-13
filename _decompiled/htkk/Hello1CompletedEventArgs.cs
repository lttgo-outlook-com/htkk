using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;

[DebuggerStepThrough]
[DesignerCategory("code")]
[GeneratedCode("wsdl", "2.0.50727.3038")]
public class Hello1CompletedEventArgs : AsyncCompletedEventArgs
{
	private object[] results;

	public string Result
	{
		get
		{
			RaiseExceptionIfNecessary();
			return (string)results[0];
		}
	}

	internal Hello1CompletedEventArgs(object[] results, Exception exception, bool cancelled, object userState)
		: base(exception, cancelled, userState)
	{
		this.results = results;
	}
}
