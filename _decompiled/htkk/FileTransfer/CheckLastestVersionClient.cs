using Foundation.Base;

namespace FileTransfer;

internal class CheckLastestVersionClient : BaseObject
{
	public const int EN_URI_NULL = -1;

	public const int EN_CHECK_FROM_SERVER_FAILED = -2;

	private string admin_endpoint_uri;

	public CheckLastestVersionClient(string admin_endpoint_uri)
	{
		this.admin_endpoint_uri = admin_endpoint_uri;
	}

	public int DoCheck(string clientGuid, string maSoThue, string curr_version, string tenCQTCapCuc, string maCQTCapCuc, string tenCQTNoiNop, string maCQTNoiNop, out string lastest_version_code, out DownloadFileInfo[] down_file_info_list, out string public_key)
	{
		if (admin_endpoint_uri == null)
		{
			lastest_version_code = null;
			down_file_info_list = null;
			public_key = null;
			return -1;
		}
		AdminServerConnector adminServerConnector = new AdminServerConnector(admin_endpoint_uri);
		if (adminServerConnector.GetLastestVersionInfo(clientGuid, maSoThue, curr_version, tenCQTCapCuc, maCQTCapCuc, tenCQTNoiNop, maCQTNoiNop, out down_file_info_list, out public_key, out lastest_version_code) != 0)
		{
			errno = -2;
			lastest_version_code = null;
			down_file_info_list = null;
			return -1;
		}
		return 0;
	}
}
