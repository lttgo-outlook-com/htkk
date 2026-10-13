using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;

[GeneratedCode("wsdl", "2.0.50727.3038")]
[DebuggerStepThrough]
[DesignerCategory("code")]
public class DownloadFilesInfoCompletedEventArgs : AsyncCompletedEventArgs
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

	public string info_str
	{
		get
		{
			RaiseExceptionIfNecessary();
			return (string)results[1];
		}
	}

	public string public_key
	{
		get
		{
			RaiseExceptionIfNecessary();
			return (string)results[2];
		}
	}

	public byte[] signatures
	{
		get
		{
			RaiseExceptionIfNecessary();
			return (byte[])results[3];
		}
	}

	public int sign_len
	{
		get
		{
			RaiseExceptionIfNecessary();
			return (int)results[4];
		}
	}

	internal DownloadFilesInfoCompletedEventArgs(object[] results, Exception exception, bool cancelled, object userState)
		: base(exception, cancelled, userState)
	{
		this.results = results;
	}
}
