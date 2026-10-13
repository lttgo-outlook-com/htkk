using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;

[GeneratedCode("wsdl", "2.0.50727.3038")]
[DebuggerStepThrough]
[DesignerCategory("code")]
public class GetLastestVersionCodeCompletedEventArgs : AsyncCompletedEventArgs
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

	public string version_code
	{
		get
		{
			RaiseExceptionIfNecessary();
			return (string)results[1];
		}
	}

	public string download_info_str
	{
		get
		{
			RaiseExceptionIfNecessary();
			return (string)results[2];
		}
	}

	public string public_key
	{
		get
		{
			RaiseExceptionIfNecessary();
			return (string)results[3];
		}
	}

	public byte[] signatures_list
	{
		get
		{
			RaiseExceptionIfNecessary();
			return (byte[])results[4];
		}
	}

	public int sign_len
	{
		get
		{
			RaiseExceptionIfNecessary();
			return (int)results[5];
		}
	}

	internal GetLastestVersionCodeCompletedEventArgs(object[] results, Exception exception, bool cancelled, object userState)
		: base(exception, cancelled, userState)
	{
		this.results = results;
	}
}
