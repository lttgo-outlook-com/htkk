using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;

[GeneratedCode("wsdl", "2.0.50727.3038")]
[DebuggerStepThrough]
[DesignerCategory("code")]
public class FindServerCompletedEventArgs : AsyncCompletedEventArgs
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

	public string file_server_uri
	{
		get
		{
			RaiseExceptionIfNecessary();
			return (string)results[1];
		}
	}

	public string lastest_version_code
	{
		get
		{
			RaiseExceptionIfNecessary();
			return (string)results[2];
		}
	}

	internal FindServerCompletedEventArgs(object[] results, Exception exception, bool cancelled, object userState)
		: base(exception, cancelled, userState)
	{
		this.results = results;
	}
}
