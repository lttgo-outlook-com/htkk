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
[WebServiceBinding(Name = "DownloadVersionSoap", Namespace = "http://tempuri.org/")]
public class DownloadVersion : SoapHttpClientProtocol
{
	private SendOrPostCallback HelloOperationCompleted;

	private SendOrPostCallback GetNumReadingOperationCompleted;

	private SendOrPostCallback GetCPU_UsageOperationCompleted;

	private SendOrPostCallback GetFileCheckSumOperationCompleted;

	private SendOrPostCallback DownloadFilesInfoOperationCompleted;

	private SendOrPostCallback GetLatestVersionOperationCompleted;

	private SendOrPostCallback AdminGetBlockOperationCompleted;

	private SendOrPostCallback InformUpdateResultOperationCompleted;

	private SendOrPostCallback GetBlockOperationCompleted;

	public event HelloCompletedEventHandler HelloCompleted;

	public event GetNumReadingCompletedEventHandler GetNumReadingCompleted;

	public event GetCPU_UsageCompletedEventHandler GetCPU_UsageCompleted;

	public event GetFileCheckSumCompletedEventHandler GetFileCheckSumCompleted;

	public event DownloadFilesInfoCompletedEventHandler DownloadFilesInfoCompleted;

	public event GetLatestVersionCompletedEventHandler GetLatestVersionCompleted;

	public event AdminGetBlockCompletedEventHandler AdminGetBlockCompleted;

	public event InformUpdateResultCompletedEventHandler InformUpdateResultCompleted;

	public event GetBlockCompletedEventHandler GetBlockCompleted;

	public DownloadVersion(string server_url)
	{
		((WebClientProtocol)this).Url = server_url;
	}

	[SoapDocumentMethod(/*Could not decode attribute arguments.*/)]
	public string Hello()
	{
		object[] array = ((SoapHttpClientProtocol)this).Invoke("Hello", new object[0]);
		return (string)array[0];
	}

	public IAsyncResult BeginHello(AsyncCallback callback, object asyncState)
	{
		return ((SoapHttpClientProtocol)this).BeginInvoke("Hello", new object[0], callback, asyncState);
	}

	public string EndHello(IAsyncResult asyncResult)
	{
		object[] array = ((SoapHttpClientProtocol)this).EndInvoke(asyncResult);
		return (string)array[0];
	}

	public void HelloAsync()
	{
		HelloAsync(null);
	}

	public void HelloAsync(object userState)
	{
		if (HelloOperationCompleted == null)
		{
			HelloOperationCompleted = OnHelloOperationCompleted;
		}
		((SoapHttpClientProtocol)this).InvokeAsync("Hello", new object[0], HelloOperationCompleted, userState);
	}

	private void OnHelloOperationCompleted(object arg)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Expected O, but got Unknown
		if (this.HelloCompleted != null)
		{
			InvokeCompletedEventArgs e10 = (InvokeCompletedEventArgs)arg;
			this.HelloCompleted(this, new HelloCompletedEventArgs(e10.Results, ((AsyncCompletedEventArgs)(object)e10).Error, ((AsyncCompletedEventArgs)(object)e10).Cancelled, ((AsyncCompletedEventArgs)(object)e10).UserState));
		}
	}

	[SoapDocumentMethod(/*Could not decode attribute arguments.*/)]
	public int GetNumReading()
	{
		object[] array = ((SoapHttpClientProtocol)this).Invoke("GetNumReading", new object[0]);
		return (int)array[0];
	}

	public IAsyncResult BeginGetNumReading(AsyncCallback callback, object asyncState)
	{
		return ((SoapHttpClientProtocol)this).BeginInvoke("GetNumReading", new object[0], callback, asyncState);
	}

	public int EndGetNumReading(IAsyncResult asyncResult)
	{
		object[] array = ((SoapHttpClientProtocol)this).EndInvoke(asyncResult);
		return (int)array[0];
	}

	public void GetNumReadingAsync()
	{
		GetNumReadingAsync(null);
	}

	public void GetNumReadingAsync(object userState)
	{
		if (GetNumReadingOperationCompleted == null)
		{
			GetNumReadingOperationCompleted = OnGetNumReadingOperationCompleted;
		}
		((SoapHttpClientProtocol)this).InvokeAsync("GetNumReading", new object[0], GetNumReadingOperationCompleted, userState);
	}

	private void OnGetNumReadingOperationCompleted(object arg)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Expected O, but got Unknown
		if (this.GetNumReadingCompleted != null)
		{
			InvokeCompletedEventArgs e10 = (InvokeCompletedEventArgs)arg;
			this.GetNumReadingCompleted(this, new GetNumReadingCompletedEventArgs(e10.Results, ((AsyncCompletedEventArgs)(object)e10).Error, ((AsyncCompletedEventArgs)(object)e10).Cancelled, ((AsyncCompletedEventArgs)(object)e10).UserState));
		}
	}

	[SoapDocumentMethod(/*Could not decode attribute arguments.*/)]
	public float GetCPU_Usage()
	{
		object[] array = ((SoapHttpClientProtocol)this).Invoke("GetCPU_Usage", new object[0]);
		return (float)array[0];
	}

	public IAsyncResult BeginGetCPU_Usage(AsyncCallback callback, object asyncState)
	{
		return ((SoapHttpClientProtocol)this).BeginInvoke("GetCPU_Usage", new object[0], callback, asyncState);
	}

	public float EndGetCPU_Usage(IAsyncResult asyncResult)
	{
		object[] array = ((SoapHttpClientProtocol)this).EndInvoke(asyncResult);
		return (float)array[0];
	}

	public void GetCPU_UsageAsync()
	{
		GetCPU_UsageAsync(null);
	}

	public void GetCPU_UsageAsync(object userState)
	{
		if (GetCPU_UsageOperationCompleted == null)
		{
			GetCPU_UsageOperationCompleted = OnGetCPU_UsageOperationCompleted;
		}
		((SoapHttpClientProtocol)this).InvokeAsync("GetCPU_Usage", new object[0], GetCPU_UsageOperationCompleted, userState);
	}

	private void OnGetCPU_UsageOperationCompleted(object arg)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Expected O, but got Unknown
		if (this.GetCPU_UsageCompleted != null)
		{
			InvokeCompletedEventArgs e10 = (InvokeCompletedEventArgs)arg;
			this.GetCPU_UsageCompleted(this, new GetCPU_UsageCompletedEventArgs(e10.Results, ((AsyncCompletedEventArgs)(object)e10).Error, ((AsyncCompletedEventArgs)(object)e10).Cancelled, ((AsyncCompletedEventArgs)(object)e10).UserState));
		}
	}

	[SoapDocumentMethod(/*Could not decode attribute arguments.*/)]
	[return: XmlElement(DataType = "base64Binary")]
	public byte[] GetFileCheckSum(int file_ind)
	{
		object[] array = ((SoapHttpClientProtocol)this).Invoke("GetFileCheckSum", new object[1] { file_ind });
		return (byte[])array[0];
	}

	public IAsyncResult BeginGetFileCheckSum(int file_ind, AsyncCallback callback, object asyncState)
	{
		return ((SoapHttpClientProtocol)this).BeginInvoke("GetFileCheckSum", new object[1] { file_ind }, callback, asyncState);
	}

	public byte[] EndGetFileCheckSum(IAsyncResult asyncResult)
	{
		object[] array = ((SoapHttpClientProtocol)this).EndInvoke(asyncResult);
		return (byte[])array[0];
	}

	public void GetFileCheckSumAsync(int file_ind)
	{
		GetFileCheckSumAsync(file_ind, null);
	}

	public void GetFileCheckSumAsync(int file_ind, object userState)
	{
		if (GetFileCheckSumOperationCompleted == null)
		{
			GetFileCheckSumOperationCompleted = OnGetFileCheckSumOperationCompleted;
		}
		((SoapHttpClientProtocol)this).InvokeAsync("GetFileCheckSum", new object[1] { file_ind }, GetFileCheckSumOperationCompleted, userState);
	}

	private void OnGetFileCheckSumOperationCompleted(object arg)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Expected O, but got Unknown
		if (this.GetFileCheckSumCompleted != null)
		{
			InvokeCompletedEventArgs e10 = (InvokeCompletedEventArgs)arg;
			this.GetFileCheckSumCompleted(this, new GetFileCheckSumCompletedEventArgs(e10.Results, ((AsyncCompletedEventArgs)(object)e10).Error, ((AsyncCompletedEventArgs)(object)e10).Cancelled, ((AsyncCompletedEventArgs)(object)e10).UserState));
		}
	}

	[SoapDocumentMethod(/*Could not decode attribute arguments.*/)]
	public int DownloadFilesInfo(out string info_str, out string public_key, [XmlElement(DataType = "base64Binary")] out byte[] signatures, out int sign_len)
	{
		object[] array = ((SoapHttpClientProtocol)this).Invoke("DownloadFilesInfo", new object[0]);
		info_str = (string)array[1];
		public_key = (string)array[2];
		signatures = (byte[])array[3];
		sign_len = (int)array[4];
		return (int)array[0];
	}

	public IAsyncResult BeginDownloadFilesInfo(AsyncCallback callback, object asyncState)
	{
		return ((SoapHttpClientProtocol)this).BeginInvoke("DownloadFilesInfo", new object[0], callback, asyncState);
	}

	public int EndDownloadFilesInfo(IAsyncResult asyncResult, out string info_str, out string public_key, out byte[] signatures, out int sign_len)
	{
		object[] array = ((SoapHttpClientProtocol)this).EndInvoke(asyncResult);
		info_str = (string)array[1];
		public_key = (string)array[2];
		signatures = (byte[])array[3];
		sign_len = (int)array[4];
		return (int)array[0];
	}

	public void DownloadFilesInfoAsync()
	{
		DownloadFilesInfoAsync(null);
	}

	public void DownloadFilesInfoAsync(object userState)
	{
		if (DownloadFilesInfoOperationCompleted == null)
		{
			DownloadFilesInfoOperationCompleted = OnDownloadFilesInfoOperationCompleted;
		}
		((SoapHttpClientProtocol)this).InvokeAsync("DownloadFilesInfo", new object[0], DownloadFilesInfoOperationCompleted, userState);
	}

	private void OnDownloadFilesInfoOperationCompleted(object arg)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Expected O, but got Unknown
		if (this.DownloadFilesInfoCompleted != null)
		{
			InvokeCompletedEventArgs e10 = (InvokeCompletedEventArgs)arg;
			this.DownloadFilesInfoCompleted(this, new DownloadFilesInfoCompletedEventArgs(e10.Results, ((AsyncCompletedEventArgs)(object)e10).Error, ((AsyncCompletedEventArgs)(object)e10).Cancelled, ((AsyncCompletedEventArgs)(object)e10).UserState));
		}
	}

	[SoapDocumentMethod(/*Could not decode attribute arguments.*/)]
	public string GetLatestVersion()
	{
		object[] array = ((SoapHttpClientProtocol)this).Invoke("GetLatestVersion", new object[0]);
		return (string)array[0];
	}

	public IAsyncResult BeginGetLatestVersion(AsyncCallback callback, object asyncState)
	{
		return ((SoapHttpClientProtocol)this).BeginInvoke("GetLatestVersion", new object[0], callback, asyncState);
	}

	public string EndGetLatestVersion(IAsyncResult asyncResult)
	{
		object[] array = ((SoapHttpClientProtocol)this).EndInvoke(asyncResult);
		return (string)array[0];
	}

	public void GetLatestVersionAsync()
	{
		GetLatestVersionAsync(null);
	}

	public void GetLatestVersionAsync(object userState)
	{
		if (GetLatestVersionOperationCompleted == null)
		{
			GetLatestVersionOperationCompleted = OnGetLatestVersionOperationCompleted;
		}
		((SoapHttpClientProtocol)this).InvokeAsync("GetLatestVersion", new object[0], GetLatestVersionOperationCompleted, userState);
	}

	private void OnGetLatestVersionOperationCompleted(object arg)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Expected O, but got Unknown
		if (this.GetLatestVersionCompleted != null)
		{
			InvokeCompletedEventArgs e10 = (InvokeCompletedEventArgs)arg;
			this.GetLatestVersionCompleted(this, new GetLatestVersionCompletedEventArgs(e10.Results, ((AsyncCompletedEventArgs)(object)e10).Error, ((AsyncCompletedEventArgs)(object)e10).Cancelled, ((AsyncCompletedEventArgs)(object)e10).UserState));
		}
	}

	[SoapDocumentMethod(/*Could not decode attribute arguments.*/)]
	[return: XmlElement(DataType = "base64Binary")]
	public byte[] AdminGetBlock(int file_id, int block_ind, out string err_msg)
	{
		object[] array = ((SoapHttpClientProtocol)this).Invoke("AdminGetBlock", new object[2] { file_id, block_ind });
		err_msg = (string)array[1];
		return (byte[])array[0];
	}

	public IAsyncResult BeginAdminGetBlock(int file_id, int block_ind, AsyncCallback callback, object asyncState)
	{
		return ((SoapHttpClientProtocol)this).BeginInvoke("AdminGetBlock", new object[2] { file_id, block_ind }, callback, asyncState);
	}

	public byte[] EndAdminGetBlock(IAsyncResult asyncResult, out string err_msg)
	{
		object[] array = ((SoapHttpClientProtocol)this).EndInvoke(asyncResult);
		err_msg = (string)array[1];
		return (byte[])array[0];
	}

	public void AdminGetBlockAsync(int file_id, int block_ind)
	{
		AdminGetBlockAsync(file_id, block_ind, null);
	}

	public void AdminGetBlockAsync(int file_id, int block_ind, object userState)
	{
		if (AdminGetBlockOperationCompleted == null)
		{
			AdminGetBlockOperationCompleted = OnAdminGetBlockOperationCompleted;
		}
		((SoapHttpClientProtocol)this).InvokeAsync("AdminGetBlock", new object[2] { file_id, block_ind }, AdminGetBlockOperationCompleted, userState);
	}

	private void OnAdminGetBlockOperationCompleted(object arg)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Expected O, but got Unknown
		if (this.AdminGetBlockCompleted != null)
		{
			InvokeCompletedEventArgs e10 = (InvokeCompletedEventArgs)arg;
			this.AdminGetBlockCompleted(this, new AdminGetBlockCompletedEventArgs(e10.Results, ((AsyncCompletedEventArgs)(object)e10).Error, ((AsyncCompletedEventArgs)(object)e10).Cancelled, ((AsyncCompletedEventArgs)(object)e10).UserState));
		}
	}

	[SoapDocumentMethod(/*Could not decode attribute arguments.*/)]
	public int InformUpdateResult(bool successfull, string session_id, string clientGuid, string maSoThue, string client_curr_version, int size)
	{
		object[] array = ((SoapHttpClientProtocol)this).Invoke("InformUpdateResult", new object[6] { successfull, session_id, clientGuid, maSoThue, client_curr_version, size });
		return (int)array[0];
	}

	public IAsyncResult BeginInformUpdateResult(bool successfull, string session_id, string clientGuid, string maSoThue, string client_curr_version, int size, AsyncCallback callback, object asyncState)
	{
		return ((SoapHttpClientProtocol)this).BeginInvoke("InformUpdateResult", new object[6] { successfull, session_id, clientGuid, maSoThue, client_curr_version, size }, callback, asyncState);
	}

	public int EndInformUpdateResult(IAsyncResult asyncResult)
	{
		object[] array = ((SoapHttpClientProtocol)this).EndInvoke(asyncResult);
		return (int)array[0];
	}

	public void InformUpdateResultAsync(bool successfull, string session_id, string clientGuid, string maSoThue, string client_curr_version, int size)
	{
		InformUpdateResultAsync(successfull, session_id, clientGuid, maSoThue, client_curr_version, size, null);
	}

	public void InformUpdateResultAsync(bool successfull, string session_id, string clientGuid, string maSoThue, string client_curr_version, int size, object userState)
	{
		if (InformUpdateResultOperationCompleted == null)
		{
			InformUpdateResultOperationCompleted = OnInformUpdateResultOperationCompleted;
		}
		((SoapHttpClientProtocol)this).InvokeAsync("InformUpdateResult", new object[6] { successfull, session_id, clientGuid, maSoThue, client_curr_version, size }, InformUpdateResultOperationCompleted, userState);
	}

	private void OnInformUpdateResultOperationCompleted(object arg)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Expected O, but got Unknown
		if (this.InformUpdateResultCompleted != null)
		{
			InvokeCompletedEventArgs e10 = (InvokeCompletedEventArgs)arg;
			this.InformUpdateResultCompleted(this, new InformUpdateResultCompletedEventArgs(e10.Results, ((AsyncCompletedEventArgs)(object)e10).Error, ((AsyncCompletedEventArgs)(object)e10).Cancelled, ((AsyncCompletedEventArgs)(object)e10).UserState));
		}
	}

	[SoapDocumentMethod(/*Could not decode attribute arguments.*/)]
	[return: XmlElement(DataType = "base64Binary")]
	public byte[] GetBlock(string session_id, string clientGuid, string maSoThue, int size, string client_curr_version, int file_id, int block_ind, bool is_last_file, bool is_first_block, bool is_first_file, [XmlElement(DataType = "base64Binary")] out byte[] block_signature, out string err_msg)
	{
		object[] array = ((SoapHttpClientProtocol)this).Invoke("GetBlock", new object[10] { session_id, clientGuid, maSoThue, size, client_curr_version, file_id, block_ind, is_last_file, is_first_block, is_first_file });
		block_signature = (byte[])array[1];
		err_msg = (string)array[2];
		return (byte[])array[0];
	}

	public IAsyncResult BeginGetBlock(string session_id, string clientGuid, string maSoThue, int size, string client_curr_version, int file_id, int block_ind, bool is_last_file, bool is_first_block, bool is_first_file, AsyncCallback callback, object asyncState)
	{
		return ((SoapHttpClientProtocol)this).BeginInvoke("GetBlock", new object[10] { session_id, clientGuid, maSoThue, size, client_curr_version, file_id, block_ind, is_last_file, is_first_block, is_first_file }, callback, asyncState);
	}

	public byte[] EndGetBlock(IAsyncResult asyncResult, out byte[] block_signature, out string err_msg)
	{
		object[] array = ((SoapHttpClientProtocol)this).EndInvoke(asyncResult);
		block_signature = (byte[])array[1];
		err_msg = (string)array[2];
		return (byte[])array[0];
	}

	public void GetBlockAsync(string session_id, string clientGuid, string maSoThue, int size, string client_curr_version, int file_id, int block_ind, bool is_last_file, bool is_first_block, bool is_first_file)
	{
		GetBlockAsync(session_id, clientGuid, maSoThue, size, client_curr_version, file_id, block_ind, is_last_file, is_first_block, is_first_file, null);
	}

	public void GetBlockAsync(string session_id, string clientGuid, string maSoThue, int size, string client_curr_version, int file_id, int block_ind, bool is_last_file, bool is_first_block, bool is_first_file, object userState)
	{
		if (GetBlockOperationCompleted == null)
		{
			GetBlockOperationCompleted = OnGetBlockOperationCompleted;
		}
		((SoapHttpClientProtocol)this).InvokeAsync("GetBlock", new object[10] { session_id, clientGuid, maSoThue, size, client_curr_version, file_id, block_ind, is_last_file, is_first_block, is_first_file }, GetBlockOperationCompleted, userState);
	}

	private void OnGetBlockOperationCompleted(object arg)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Expected O, but got Unknown
		if (this.GetBlockCompleted != null)
		{
			InvokeCompletedEventArgs e10 = (InvokeCompletedEventArgs)arg;
			this.GetBlockCompleted(this, new GetBlockCompletedEventArgs(e10.Results, ((AsyncCompletedEventArgs)(object)e10).Error, ((AsyncCompletedEventArgs)(object)e10).Cancelled, ((AsyncCompletedEventArgs)(object)e10).UserState));
		}
	}

	public void CancelAsync(object userState)
	{
		((HttpWebClientProtocol)this).CancelAsync(userState);
	}
}
