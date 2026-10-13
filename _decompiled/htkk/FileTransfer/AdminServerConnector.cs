using System;
using System.Net;
using System.Web.Services.Protocols;
using Foundation.Base;

namespace FileTransfer;

internal class AdminServerConnector : BaseObject
{
	private const char DOWNLOAD_INF_DELIMITER = '\n';

	private const char INF_FIELD_DELIMITER = ';';

	private ServerBalancer server_balancer;

	public AdminServerConnector(string admin_server_uri)
	{
		if (admin_server_uri == null)
		{
			is_null = true;
			server_balancer = null;
			return;
		}
		try
		{
			server_balancer = new ServerBalancer(admin_server_uri);
		}
		catch
		{
			is_null = true;
			server_balancer = null;
		}
	}

	public int SetProxy(string proxy_addr, int proxy_port)
	{
		if (server_balancer == null || is_null || proxy_addr == null)
		{
			return -1;
		}
		int result = 0;
		try
		{
			IWebProxy proxy = new WebProxy(proxy_addr, proxy_port);
			((HttpWebClientProtocol)server_balancer).Proxy = proxy;
		}
		catch
		{
			result = -1;
		}
		return result;
	}

	public int SetProxy(string proxy_addr, int proxy_port, string user_name, string passwd)
	{
		if (server_balancer == null || is_null || proxy_addr == null || user_name == null || passwd == null)
		{
			return -1;
		}
		int result = 0;
		try
		{
			IWebProxy webProxy = new WebProxy(proxy_addr, proxy_port);
			webProxy.Credentials = new NetworkCredential(user_name, passwd);
			webProxy.Credentials = new NetworkCredential("userName", "password");
			((HttpWebClientProtocol)server_balancer).Proxy = webProxy;
		}
		catch
		{
			result = -1;
		}
		return result;
	}

	public void LookupServer(out string file_server_uri, out string lastest_version)
	{
		try
		{
			if (server_balancer.FindServer(out file_server_uri, out lastest_version) != 0)
			{
				file_server_uri = null;
				lastest_version = null;
			}
		}
		catch (Exception ex)
		{
			err_msg = ex.Message;
			file_server_uri = null;
			lastest_version = null;
		}
	}

	public string StartDownloadSession(out int balancing_mode)
	{
		if (is_null || server_balancer == null)
		{
			balancing_mode = -1;
			return null;
		}
		string result;
		try
		{
			result = server_balancer.StartDownloadSession(out balancing_mode);
		}
		catch (Exception ex)
		{
			balancing_mode = -1;
			err_msg = ex.Message;
			result = null;
		}
		return result;
	}

	public int GetLastestVersionInfo(string clientGuid, string maSoThue, string curr_version, string tenCQTCapCuc, string maCQTCapCuc, string tenCQTNoiNop, string maCQTNoiNop, out DownloadFileInfo[] file_info_list, out string public_key, out string version_code)
	{
		version_code = null;
		file_info_list = null;
		public_key = null;
		if (is_null || server_balancer == null)
		{
			return -1;
		}
		int result = 0;
		try
		{
			if (server_balancer.GetLastestVersionCode(maSoThue, clientGuid, curr_version, tenCQTCapCuc, maCQTCapCuc, tenCQTNoiNop, maCQTNoiNop, out var version_code2, out var download_info_str, out public_key, out var signatures_list, out var sign_len) != 0)
			{
				result = -1;
			}
			else
			{
				file_info_list = DownloadFileInfo.LoadListFromStr(download_info_str, '\n', ';', out errno);
				if (file_info_list == null)
				{
					result = -1;
				}
				else
				{
					version_code = version_code2;
					int num = 0;
					try
					{
						for (int i = 0; i < file_info_list.Length; i++)
						{
							byte[] array = new byte[sign_len];
							for (int j = 0; j < sign_len; j++)
							{
								array[j] = signatures_list[j + num];
							}
							file_info_list[i].SetFileSignature(array);
							num += sign_len;
						}
					}
					catch
					{
						result = -1;
					}
				}
			}
		}
		catch (Exception)
		{
			result = -1;
		}
		return result;
	}
}
