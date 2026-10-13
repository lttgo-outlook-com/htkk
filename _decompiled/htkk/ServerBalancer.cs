using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Threading;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.Xml.Serialization;

[DebuggerStepThrough]
[GeneratedCode("wsdl", "2.0.50727.3038")]
[DesignerCategory("code")]
[WebServiceBinding(Name = "ServerBalancerSoap", Namespace = "http://tempuri.org/")]
public class ServerBalancer : SoapHttpClientProtocol
{
	private SendOrPostCallback Hello1OperationCompleted;

	private SendOrPostCallback FindServerOperationCompleted;

	private SendOrPostCallback StartDownloadSessionOperationCompleted;

	private SendOrPostCallback GetLastestVersionCodeOperationCompleted;

	public event Hello1CompletedEventHandler Hello1Completed;

	public event FindServerCompletedEventHandler FindServerCompleted;

	public event StartDownloadSessionCompletedEventHandler StartDownloadSessionCompleted;

	public event GetLastestVersionCodeCompletedEventHandler GetLastestVersionCodeCompleted;

	public ServerBalancer(string server_url)
	{
		((WebClientProtocol)this).Url = server_url;
	}

	[SoapDocumentMethod(/*Could not decode attribute arguments.*/)]
	public string Hello1()
	{
		object[] array = ((SoapHttpClientProtocol)this).Invoke("Hello1", new object[0]);
		return (string)array[0];
	}

	public IAsyncResult BeginHello1(AsyncCallback callback, object asyncState)
	{
		return ((SoapHttpClientProtocol)this).BeginInvoke("Hello1", new object[0], callback, asyncState);
	}

	public string EndHello1(IAsyncResult asyncResult)
	{
		object[] array = ((SoapHttpClientProtocol)this).EndInvoke(asyncResult);
		return (string)array[0];
	}

	public void Hello1Async()
	{
		Hello1Async(null);
	}

	public void Hello1Async(object userState)
	{
		if (Hello1OperationCompleted == null)
		{
			Hello1OperationCompleted = OnHello1OperationCompleted;
		}
		((SoapHttpClientProtocol)this).InvokeAsync("Hello1", new object[0], Hello1OperationCompleted, userState);
	}

	private void OnHello1OperationCompleted(object arg)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Expected O, but got Unknown
		if (this.Hello1Completed != null)
		{
			InvokeCompletedEventArgs e10 = (InvokeCompletedEventArgs)arg;
			this.Hello1Completed(this, new Hello1CompletedEventArgs(e10.Results, ((AsyncCompletedEventArgs)(object)e10).Error, ((AsyncCompletedEventArgs)(object)e10).Cancelled, ((AsyncCompletedEventArgs)(object)e10).UserState));
		}
	}

	[SoapDocumentMethod(/*Could not decode attribute arguments.*/)]
	public int FindServer(out string file_server_uri, out string lastest_version_code)
	{
		object[] array = ((SoapHttpClientProtocol)this).Invoke("FindServer", new object[0]);
		file_server_uri = (string)array[1];
		lastest_version_code = (string)array[2];
		return (int)array[0];
	}

	public IAsyncResult BeginFindServer(AsyncCallback callback, object asyncState)
	{
		return ((SoapHttpClientProtocol)this).BeginInvoke("FindServer", new object[0], callback, asyncState);
	}

	public int EndFindServer(IAsyncResult asyncResult, out string file_server_uri, out string lastest_version_code)
	{
		object[] array = ((SoapHttpClientProtocol)this).EndInvoke(asyncResult);
		file_server_uri = (string)array[1];
		lastest_version_code = (string)array[2];
		return (int)array[0];
	}

	public void FindServerAsync()
	{
		FindServerAsync(null);
	}

	public void FindServerAsync(object userState)
	{
		if (FindServerOperationCompleted == null)
		{
			FindServerOperationCompleted = OnFindServerOperationCompleted;
		}
		((SoapHttpClientProtocol)this).InvokeAsync("FindServer", new object[0], FindServerOperationCompleted, userState);
	}

	private void OnFindServerOperationCompleted(object arg)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Expected O, but got Unknown
		if (this.FindServerCompleted != null)
		{
			InvokeCompletedEventArgs e10 = (InvokeCompletedEventArgs)arg;
			this.FindServerCompleted(this, new FindServerCompletedEventArgs(e10.Results, ((AsyncCompletedEventArgs)(object)e10).Error, ((AsyncCompletedEventArgs)(object)e10).Cancelled, ((AsyncCompletedEventArgs)(object)e10).UserState));
		}
	}

	[SoapDocumentMethod(/*Could not decode attribute arguments.*/)]
	public string StartDownloadSession(out int bal_mode)
	{
		object[] array = ((SoapHttpClientProtocol)this).Invoke("StartDownloadSession", new object[0]);
		bal_mode = (int)array[1];
		return (string)array[0];
	}

	public IAsyncResult BeginStartDownloadSession(AsyncCallback callback, object asyncState)
	{
		return ((SoapHttpClientProtocol)this).BeginInvoke("StartDownloadSession", new object[0], callback, asyncState);
	}

	public string EndStartDownloadSession(IAsyncResult asyncResult, out int bal_mode)
	{
		object[] array = ((SoapHttpClientProtocol)this).EndInvoke(asyncResult);
		bal_mode = (int)array[1];
		return (string)array[0];
	}

	public void StartDownloadSessionAsync()
	{
		StartDownloadSessionAsync(null);
	}

	public void StartDownloadSessionAsync(object userState)
	{
		if (StartDownloadSessionOperationCompleted == null)
		{
			StartDownloadSessionOperationCompleted = OnStartDownloadSessionOperationCompleted;
		}
		((SoapHttpClientProtocol)this).InvokeAsync("StartDownloadSession", new object[0], StartDownloadSessionOperationCompleted, userState);
	}

	private void OnStartDownloadSessionOperationCompleted(object arg)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Expected O, but got Unknown
		if (this.StartDownloadSessionCompleted != null)
		{
			InvokeCompletedEventArgs e10 = (InvokeCompletedEventArgs)arg;
			this.StartDownloadSessionCompleted(this, new StartDownloadSessionCompletedEventArgs(e10.Results, ((AsyncCompletedEventArgs)(object)e10).Error, ((AsyncCompletedEventArgs)(object)e10).Cancelled, ((AsyncCompletedEventArgs)(object)e10).UserState));
		}
	}

	[SoapDocumentMethod(/*Could not decode attribute arguments.*/)]
	public int GetLastestVersionCode(string maSoThue, string clientGuid, string curr_version, string tenCQTCapCuc, string maCQTCapCuc, string tenCQTNoiNop, string maCQTNoiNop, out string version_code, out string download_info_str, out string public_key, [XmlElement(DataType = "base64Binary")] out byte[] signatures_list, out int sign_len)
	{
		object[] array = ((SoapHttpClientProtocol)this).Invoke("GetLastestVersionCode", new object[7] { maSoThue, clientGuid, curr_version, tenCQTCapCuc, maCQTCapCuc, tenCQTNoiNop, maCQTNoiNop });
		version_code = (string)array[1];
		download_info_str = (string)array[2];
		public_key = (string)array[3];
		signatures_list = (byte[])array[4];
		sign_len = (int)array[5];
		return (int)array[0];
	}

	public IAsyncResult BeginGetLastestVersionCode(string maSoThue, string clientGuid, AsyncCallback callback, object asyncState)
	{
		return ((SoapHttpClientProtocol)this).BeginInvoke("GetLastestVersionCode", new object[2] { maSoThue, clientGuid }, callback, asyncState);
	}

	public int EndGetLastestVersionCode(IAsyncResult asyncResult, out string version_code, out string download_info_str, out string public_key, out byte[] signatures_list, out int sign_len)
	{
		object[] array = ((SoapHttpClientProtocol)this).EndInvoke(asyncResult);
		version_code = (string)array[1];
		download_info_str = (string)array[2];
		public_key = (string)array[3];
		signatures_list = (byte[])array[4];
		sign_len = (int)array[5];
		return (int)array[0];
	}

	public void GetLastestVersionCodeAsync(string maSoThue, string clientGuid)
	{
		GetLastestVersionCodeAsync(maSoThue, clientGuid, null);
	}

	public void GetLastestVersionCodeAsync(string maSoThue, string clientGuid, object userState)
	{
		if (GetLastestVersionCodeOperationCompleted == null)
		{
			GetLastestVersionCodeOperationCompleted = OnGetLastestVersionCodeOperationCompleted;
		}
		((SoapHttpClientProtocol)this).InvokeAsync("GetLastestVersionCode", new object[2] { maSoThue, clientGuid }, GetLastestVersionCodeOperationCompleted, userState);
	}

	private void OnGetLastestVersionCodeOperationCompleted(object arg)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Expected O, but got Unknown
		if (this.GetLastestVersionCodeCompleted != null)
		{
			InvokeCompletedEventArgs e10 = (InvokeCompletedEventArgs)arg;
			this.GetLastestVersionCodeCompleted(this, new GetLastestVersionCodeCompletedEventArgs(e10.Results, ((AsyncCompletedEventArgs)(object)e10).Error, ((AsyncCompletedEventArgs)(object)e10).Cancelled, ((AsyncCompletedEventArgs)(object)e10).UserState));
		}
	}

	public void CancelAsync(object userState)
	{
		((HttpWebClientProtocol)this).CancelAsync(userState);
	}
}
