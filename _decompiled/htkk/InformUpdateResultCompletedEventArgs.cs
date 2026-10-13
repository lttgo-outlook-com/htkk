using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;

[GeneratedCode("wsdl", "2.0.50727.3038")]
[DebuggerStepThrough]
[DesignerCategory("code")]
public class InformUpdateResultCompletedEventArgs : AsyncCompletedEventArgs
{
	private object[] results;

	public int Result
	{
		get
		{
			RaiseExceptionIfNecessary();
			return (int)results[0];
		}
	}

	internal InformUpdateResultCompletedEventArgs(object[] results, Exception exception, bool cancelled, object userState)
		: base(exception, cancelled, userState)
	{
		this.results = results;
	}
}
