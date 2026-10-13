using System;
using System.Net;
using System.Web.Services.Protocols;
using Foundation.Base;

namespace FileTransfer;

internal class FileServerConnector : BaseObject
{
	private DownloadVersion download_version;

	public FileServerConnector(string file_server_uri)
	{
		if (file_server_uri == null)
		{
			is_null = true;
			download_version = null;
		}
		else
		{
			download_version = new DownloadVersion(file_server_uri);
		}
	}

	public int SetProxy(string proxy_addr, int proxy_port)
	{
		if (download_version == null || is_null || proxy_addr == null)
		{
			return -1;
		}
		int result = 0;
		try
		{
			IWebProxy proxy = new WebProxy(proxy_addr, proxy_port);
			((HttpWebClientProtocol)download_version).Proxy = proxy;
		}
		catch
		{
			result = -1;
		}
		return result;
	}

	public int SetProxy(string proxy_addr, int proxy_port, string user_name, string passwd)
	{
		if (download_version == null || is_null || proxy_addr == null || user_name == null || passwd == null)
		{
			return -1;
		}
		int result = 0;
		try
		{
			IWebProxy webProxy = new WebProxy(proxy_addr, proxy_port);
			webProxy.Credentials = new NetworkCredential(user_name, passwd);
			webProxy.Credentials = new NetworkCredential("userName", "password");
			((HttpWebClientProtocol)download_version).Proxy = webProxy;
		}
		catch
		{
			result = -1;
		}
		return result;
	}

	public int DownloadBlock(string session_id, string clientGuid, string tax_code, int size, string curr_version, int file_ind, int block_ind, bool is_last_file, bool is_first_block, bool is_first_file, out byte[] bytes, out byte[] blk_sign)
	{
		if (is_null || download_version == null)
		{
			bytes = null;
			blk_sign = null;
			return -1;
		}
		int result = 0;
		try
		{
			bytes = download_version.GetBlock(session_id, clientGuid, tax_code, size, curr_version, file_ind, block_ind, is_last_file, is_first_block, is_first_file, out blk_sign, out err_msg);
		}
		catch (Exception ex)
		{
			bytes = null;
			err_msg = ex.Message;
			blk_sign = null;
			result = -1;
		}
		return result;
	}

	public int SendUpdateResult(bool successfull, string session_id, string clientGuid, string tax_code, string client_version, int size)
	{
		if (session_id == null)
		{
			return -1;
		}
		int num = 0;
		try
		{
			num = download_version.InformUpdateResult(successfull, session_id, clientGuid, tax_code, client_version, size);
		}
		catch
		{
			num = -1;
		}
		return num;
	}
}
