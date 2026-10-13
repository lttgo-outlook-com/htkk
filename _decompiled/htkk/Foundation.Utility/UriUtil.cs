using System;
using Foundation.Base;

namespace Foundation.Utility;

internal class UriUtil : BaseObject
{
	public static int DetectSmtpMailAddr(string uri_str, out string mail_addr)
	{
		int result = 0;
		try
		{
			Uri uri = new Uri(uri_str);
			string host = uri.Host;
			string userInfo = uri.UserInfo;
			mail_addr = userInfo + "@" + host;
		}
		catch
		{
			mail_addr = "";
			result = -1;
		}
		return result;
	}

	public static int DetectPop3HostInfo(string uri_str, out string pop3_host, out int pop3_port)
	{
		int result = 0;
		try
		{
			Uri uri = new Uri(uri_str);
			pop3_host = uri.Host;
			pop3_port = uri.Port;
		}
		catch
		{
			pop3_host = "";
			pop3_port = 0;
			result = -1;
		}
		return result;
	}

	public static int DetectFtpHostInfo(string uri_str, out string ftp_host, out int ftp_port, out string ftp_folder)
	{
		int result = 0;
		try
		{
			Uri uri = new Uri(uri_str);
			ftp_host = uri.Host;
			ftp_port = uri.Port;
			ftp_folder = uri.AbsolutePath;
		}
		catch
		{
			ftp_host = "";
			ftp_port = 0;
			ftp_folder = "";
			result = -1;
		}
		return result;
	}

	public static int DetectFileHostInfo(string uri_str, out string folder_path)
	{
		int result = 0;
		try
		{
			Uri uri = new Uri(uri_str);
			folder_path = uri.LocalPath;
		}
		catch
		{
			folder_path = null;
			result = -1;
		}
		return result;
	}

	public static int DetectWebServiceHostInfo(string uri_str, out string box_name)
	{
		int result = 0;
		try
		{
			Uri uri = new Uri(uri_str);
			box_name = uri.Host;
		}
		catch
		{
			box_name = null;
			result = -1;
		}
		return result;
	}

	public static int DetectJmsHostInfo(string uri_str, out string jms_server, out int jms_port, out string queue_name)
	{
		int result = 0;
		try
		{
			Uri uri = new Uri(uri_str);
			jms_server = uri.Host;
			jms_port = uri.Port;
			queue_name = uri.AbsolutePath;
		}
		catch
		{
			jms_server = null;
			queue_name = null;
			jms_port = 0;
			result = -1;
		}
		return result;
	}

	public static string ConvertFileUriToLocalPath(string uri_str)
	{
		try
		{
			Uri uri = new Uri(uri_str);
			return uri.LocalPath;
		}
		catch
		{
			return null;
		}
	}
}
