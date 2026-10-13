#define DEBUG
using System;
using System.ComponentModel;
using System.Diagnostics;
using Foundation.Base;
using Foundation.Utility;

namespace FileTransfer;

internal class GetVerFilesClient : BaseObject
{
	public const int EN_MAKE_BUFFER_FAILED = -1;

	public const int EN_MEMBERS_NULL = -2;

	private const int EN_PARAMS_NULL = -3;

	public const int EN_IND_NOT_IN_RANGE = -4;

	public const int EN_LAST_IND = -5;

	public const int EN_INF_LIST_NULL = -6;

	public const int EN_OPEN_BUFFER_FAILED = -7;

	public const int EN_FIND_DOWNLOADED_FAILED = -8;

	public const int EN_FILE_SIGN_NOT_VERIFIED = -9;

	public const int EN_COUNT_DOWNED_BLK_FAILED = -10;

	public const int EN_DONWLOAD_FILE_SERVER_FAILED = -11;

	public const int EN_MERGE_FILE_FAILED = -12;

	public const int EN_VERSION_CHANGED = -13;

	public const int EN_CANCELED_BY_USER = -14;

	public const int EN_LOOKUP_ADMIN_SERVER_FAILED = -15;

	public const int EN_COUNT_TOLTAL_FAILED = -16;

	public const int EN_NO_SESSION_ID = -17;

	public const int EN_BLOCK_SIGN_NOT_VERIFIED = -18;

	public const int EN_SAVE_BLOCK_FAILED = -19;

	public const int EN_GET_FRM_SERVER_FAILED = -20;

	public const int EN_BYTES_NULL = -21;

	private const int BAL_MOD_BLOCK = 0;

	private const int BAL_MOD_FILE = 1;

	private const int BAL_MOD_SESSION = 2;

	private string admin_endpoint_uri;

	private int retry_num;

	private BlockBuffer block_buffer;

	private string buffer_folder;

	private string version_code;

	private DownloadFileInfo[] dfile_info_list;

	private string public_key;

	private string proxy_addr;

	private int proxy_port;

	private string proxy_user;

	private string proxy_passwd;

	private BackgroundWorker worker;

	private DoWorkEventArgs event_args;

	private int curr_file_ind = -1;

	private int curr_block_ind = -1;

	private int total_blocks;

	private DownloadProgress down_prog;

	private int downloaded_blocks_num;

	private int downloaded_size;

	public GetVerFilesClient(string admin_endpoint_uri, int retry_num, string buffer_folder, BackgroundWorker worker, DoWorkEventArgs event_args)
	{
		this.admin_endpoint_uri = admin_endpoint_uri;
		this.buffer_folder = buffer_folder;
		this.retry_num = retry_num;
		this.worker = worker;
		this.event_args = event_args;
		proxy_addr = "";
		proxy_port = -1;
		proxy_user = "";
		proxy_passwd = "";
	}

	public void SetProxy(string proxy_addr, int proxy_port)
	{
		this.proxy_addr = proxy_addr;
		this.proxy_port = proxy_port;
	}

	public void SetProxy(string proxy_addr, int proxy_port, string user_name, string passwd)
	{
		this.proxy_addr = proxy_addr;
		this.proxy_port = proxy_port;
		proxy_user = user_name;
		proxy_passwd = passwd;
	}

	public int SetPublicKey(string public_key)
	{
		if (public_key == null)
		{
			return -1;
		}
		this.public_key = public_key;
		return 0;
	}

	public int GetCurrFileInd()
	{
		return curr_file_ind;
	}

	public int GetBlockInd()
	{
		return curr_block_ind;
	}

	private int CheckParamsAndMemnber(string version_code, DownloadFileInfo[] dfile_info_list)
	{
		if (dfile_info_list == null || version_code == null)
		{
			errno = -3;
			return -1;
		}
		int num = 0;
		for (int i = 0; i < dfile_info_list.Length; i++)
		{
			if (dfile_info_list[i] == null)
			{
				errno = -3;
				num = -1;
				break;
			}
		}
		if (num != 0)
		{
			return num;
		}
		if (admin_endpoint_uri == null || buffer_folder == null || retry_num <= 0)
		{
			errno = -2;
			num = -1;
		}
		return num;
	}

	private int DownloadAllFiles(string file_server_uri, int balacing_mode, string session_id, string clientGuid, string tax_code, int total_size, string curr_version, DownloadFileInfo[] dfile_info_list, out int last_downloaded_ind)
	{
		int result = 0;
		Debug.Print("No file and block downloaded");
		down_prog.SetPrevNumFiles(0);
		down_prog.SetPrevSize(0);
		down_prog.SetTotalNumFiles(dfile_info_list.Length);
		downloaded_blocks_num = 0;
		block_buffer.ClearBuffer();
		worker.ReportProgress(downloaded_blocks_num * 100 / total_blocks, down_prog);
		bool flag = false;
		last_downloaded_ind = -1;
		for (int i = 0; i < dfile_info_list.Length; i++)
		{
			if (worker.CancellationPending)
			{
				event_args.Cancel = true;
				result = -1;
				break;
			}
			int numBlocks = dfile_info_list[i].GetNumBlocks();
			bool is_last_file = false;
			if (i == dfile_info_list.Length - 1)
			{
				is_last_file = true;
			}
			bool is_first_file = false;
			if (i == 0)
			{
				is_first_file = true;
			}
			bool flag2 = false;
			string file_server_uri2;
			if (file_server_uri == null)
			{
				if (balacing_mode == 1)
				{
					TryingLookupServer(out file_server_uri2, out var _);
					if (file_server_uri2 == null)
					{
						flag2 = true;
					}
				}
				else
				{
					file_server_uri2 = null;
				}
			}
			else
			{
				file_server_uri2 = file_server_uri;
			}
			if (flag2)
			{
				errno = -15;
				result = -1;
				flag = true;
			}
			else
			{
				Debug.Print("Number of blocks: " + numBlocks);
				if (DownloadBlocks(file_server_uri2, session_id, clientGuid, tax_code, total_size, curr_version, is_last_file, is_first_file, dfile_info_list[i], numBlocks, 0, retry_num) != 0)
				{
					result = -1;
					flag = true;
				}
				else if (block_buffer.MakeFileFromBlocks(dfile_info_list[i], dfile_info_list[i].GetNumBlocks()) != 0)
				{
					flag = true;
					errno = -12;
					err_msg = block_buffer.GetErrorMsg();
					result = -1;
				}
				else
				{
					down_prog.IncreaseDownloadedNumFiles();
					worker.ReportProgress(downloaded_blocks_num * 100 / total_blocks, down_prog);
				}
			}
			if (flag)
			{
				break;
			}
			last_downloaded_ind = i;
		}
		return result;
	}

	private int ResumeDonwload(string file_server_uri, int balancing_mode, string session_id, string clientGuid, string tax_code, int total_size, string curr_version, DownloadFileInfo[] dfile_info_list, int last_files_list_index, int downloading_file_index, int last_block_ind, out int last_downloaded_ind)
	{
		last_downloaded_ind = -1;
		last_downloaded_ind = -1;
		int num = 0;
		if (CountBlocksDownloaded(last_files_list_index, last_block_ind, out downloaded_blocks_num, out downloaded_size) != 0)
		{
			errno = -10;
			block_buffer.ClearBuffer();
			return -1;
		}
		down_prog.SetPrevSize(downloaded_size);
		down_prog.SetDownloadedSize(downloaded_size);
		down_prog.SetPrevNumFiles(last_files_list_index);
		down_prog.SetDownloadedNumFiles(last_files_list_index + 1);
		down_prog.SetTotalNumFiles(dfile_info_list.Length);
		worker.ReportProgress(downloaded_blocks_num * 100 / total_blocks, down_prog);
		int size = total_size - downloaded_size;
		int num2;
		try
		{
			num2 = dfile_info_list[downloading_file_index].GetNumBlocks();
		}
		catch (Exception ex)
		{
			err_msg = ex.Message;
			errno = -10000;
			num2 = -1;
		}
		if (num2 <= 0)
		{
			return -1;
		}
		string file_server_uri2;
		string lastest_version;
		if (num2 - 1 <= last_block_ind)
		{
			if (block_buffer.MakeFileFromBlocks(dfile_info_list[downloading_file_index], num2) != 0)
			{
				errno = -12;
				err_msg = block_buffer.GetErrorMsg();
				num = -1;
			}
		}
		else
		{
			bool is_last_file = false;
			if (downloading_file_index == dfile_info_list.Length - 1)
			{
				is_last_file = true;
			}
			bool flag = false;
			if (file_server_uri == null)
			{
				if (balancing_mode == 1)
				{
					TryingLookupServer(out file_server_uri2, out lastest_version);
					if (file_server_uri2 == null)
					{
						flag = true;
					}
				}
				else
				{
					file_server_uri2 = null;
				}
			}
			else
			{
				file_server_uri2 = file_server_uri;
			}
			if (flag)
			{
				errno = -15;
				num = -1;
			}
			else if (DownloadBlocks(file_server_uri2, session_id, clientGuid, tax_code, size, curr_version, is_last_file, is_first_file: true, dfile_info_list[downloading_file_index], num2, last_block_ind + 1, retry_num) != 0)
			{
				num = -1;
			}
			else if (block_buffer.MakeFileFromBlocks(dfile_info_list[downloading_file_index], num2) != 0)
			{
				errno = -12;
				err_msg = block_buffer.GetErrorMsg();
				num = -1;
			}
		}
		if (num != 0)
		{
			return num;
		}
		down_prog.IncreaseDownloadedNumFiles();
		worker.ReportProgress(downloaded_blocks_num * 100 / total_blocks, down_prog);
		bool flag2 = false;
		for (int i = last_files_list_index + 2; i < dfile_info_list.Length; i++)
		{
			if (worker.CancellationPending)
			{
				event_args.Cancel = true;
				flag2 = true;
				num = -1;
				continue;
			}
			bool is_last_file = false;
			if (i == dfile_info_list.Length - 1)
			{
				is_last_file = true;
			}
			Debug.Print("Downloading file number: " + i);
			int numBlocks = dfile_info_list[i].GetNumBlocks();
			bool flag = false;
			if (file_server_uri == null)
			{
				if (balancing_mode == 1)
				{
					TryingLookupServer(out file_server_uri2, out lastest_version);
					if (file_server_uri2 == null)
					{
						flag = true;
					}
				}
				else
				{
					file_server_uri2 = null;
				}
			}
			else
			{
				file_server_uri2 = file_server_uri;
			}
			if (flag)
			{
				errno = -15;
				num = -1;
				flag2 = true;
			}
			else if (DownloadBlocks(file_server_uri2, session_id, clientGuid, tax_code, total_size, curr_version, is_last_file, is_first_file: false, dfile_info_list[i], numBlocks, 0, retry_num) != 0)
			{
				num = -1;
				flag2 = true;
			}
			else if (block_buffer.MakeFileFromBlocks(dfile_info_list[i], numBlocks) != 0)
			{
				errno = -12;
				err_msg = block_buffer.GetErrorMsg();
				num = -1;
				flag2 = true;
			}
			else
			{
				down_prog.IncreaseDownloadedNumFiles();
				worker.ReportProgress(downloaded_blocks_num * 100 / total_blocks, down_prog);
			}
			if (flag2)
			{
				break;
			}
			last_downloaded_ind = i;
		}
		return num;
	}

	public int SendUpdateResult(bool successfull, string session_id, string clientGuid, string tax_code, string client_version, int size)
	{
		TryingLookupServer(out var file_server_uri, out var _);
		if (file_server_uri == null)
		{
			return -1;
		}
		FileServerConnector fileServerConnector = new FileServerConnector(file_server_uri);
		if (proxy_addr != "")
		{
			if (proxy_user != "")
			{
				fileServerConnector.SetProxy(proxy_addr, proxy_port, proxy_user, proxy_passwd);
			}
			else
			{
				fileServerConnector.SetProxy(proxy_addr, proxy_port);
			}
		}
		return fileServerConnector.SendUpdateResult(successfull, session_id, clientGuid, tax_code, client_version, size);
	}

	public int DownloadNewVersion(string version_code, string curr_version, string clientGuid, string tax_code, DownloadFileInfo[] dfile_info_list, out string[] temp_file_paths, out string download_session_id, out int pack_size)
	{
		temp_file_paths = null;
		download_session_id = "";
		pack_size = 0;
		if (CheckParamsAndMemnber(version_code, dfile_info_list) != 0)
		{
			errno = -3;
			return -1;
		}
		block_buffer = new BlockBuffer(buffer_folder, version_code);
		if (block_buffer.Open() != 0)
		{
			errno = -7;
			return -1;
		}
		this.version_code = version_code;
		this.dfile_info_list = dfile_info_list;
		block_buffer.CheckExistingDownloaded(out var downloading_file_id, out var file_id_list, out var block_index_list);
		if (downloading_file_id < -1)
		{
			errno = -8;
			block_buffer.ClearBuffer();
			return -1;
		}
		total_blocks = 0;
		int num = 0;
		try
		{
			for (int i = 0; i < dfile_info_list.Length; i++)
			{
				total_blocks += dfile_info_list[i].GetNumBlocks();
				num += (int)dfile_info_list[i].GetFileSize();
			}
		}
		catch
		{
			total_blocks = -1;
		}
		if (total_blocks < 0)
		{
			errno = -16;
			return -1;
		}
		pack_size = num;
		down_prog = new DownloadProgress();
		down_prog.SetTotalSize(num);
		if (file_id_list == null)
		{
			block_buffer.ClearFileBuffer();
		}
		if (block_index_list == null)
		{
			block_buffer.ClearBlockBuffer();
		}
		bool flag = false;
		if (CheckFileIdList(file_id_list, downloading_file_id, out var last_downloaded_index, out var downloading_index) != 0)
		{
			flag = true;
		}
		else if (last_downloaded_index < 0 && block_index_list == null)
		{
			flag = true;
		}
		StartSession(out var session_id, out var balancing_mode);
		if (session_id == null)
		{
			errno = -17;
			return -1;
		}
		download_session_id = session_id;
		if (balancing_mode != 0 && balancing_mode != 1 && balancing_mode != 2)
		{
			balancing_mode = 1;
		}
		string file_server_uri = null;
		bool flag2 = false;
		if (balancing_mode == 2)
		{
			TryingLookupServer(out file_server_uri, out var _);
			if (file_server_uri == null)
			{
				flag2 = true;
			}
		}
		if (flag2)
		{
			errno = -15;
			return -1;
		}
		int num2 = 0;
		string[] array = new string[dfile_info_list.Length];
		int last_downloaded_ind;
		if (flag)
		{
			block_buffer.ClearBuffer();
			num2 = DownloadAllFiles(file_server_uri, balancing_mode, session_id, clientGuid, tax_code, num, curr_version, dfile_info_list, out last_downloaded_ind);
		}
		else if (last_downloaded_index == dfile_info_list.Length - 1)
		{
			Debug.Print("All files downloaded");
			num2 = 0;
		}
		else
		{
			int last_block_ind = ((block_index_list != null) ? (block_index_list.Length - 1) : (-1));
			num2 = ResumeDonwload(file_server_uri, balancing_mode, session_id, clientGuid, tax_code, num, curr_version, dfile_info_list, last_downloaded_index, downloading_index, last_block_ind, out last_downloaded_ind);
		}
		if (num2 == 0)
		{
			int[] array2 = new int[dfile_info_list.Length];
			for (int i = 0; i < array2.Length; i++)
			{
				array2[i] = dfile_info_list[i].GetFileId();
			}
			temp_file_paths = block_buffer.BufFilePaths(array2);
			if (temp_file_paths == null)
			{
				num2 = -1;
			}
			else if (CheckFileSignatures(dfile_info_list, temp_file_paths) != 0)
			{
				errno = -9;
				block_buffer.ClearBuffer();
				num2 = -1;
			}
		}
		return num2;
	}

	private int CheckFileIdList(int[] id_list, int downloading_id, out int last_downloaded_index, out int downloading_index)
	{
		last_downloaded_index = -1;
		downloading_index = 0;
		int num = 0;
		if (id_list == null)
		{
			if (downloading_id >= 0)
			{
				if (dfile_info_list[0] == null)
				{
					num = -1;
				}
				else if (dfile_info_list[0].GetFileId() != downloading_id)
				{
					num = -1;
				}
				else
				{
					last_downloaded_index = -1;
					downloading_index = 0;
				}
			}
			else
			{
				last_downloaded_index = -1;
				downloading_index = 0;
			}
		}
		else
		{
			try
			{
				int num2 = id_list.Length;
				if (num2 > dfile_info_list.Length)
				{
					num = -1;
				}
				else
				{
					for (int i = 0; i < num2; i++)
					{
						if (dfile_info_list[i].GetFileId() != id_list[i])
						{
							num = -1;
							break;
						}
					}
					if (num == 0)
					{
						if (downloading_id >= 0)
						{
							if (dfile_info_list[num2].GetFileId() != downloading_id)
							{
								num = -1;
							}
							else
							{
								last_downloaded_index = num2 - 1;
								downloading_index = num2;
							}
						}
						else
						{
							last_downloaded_index = num2 - 1;
							downloading_index = num2;
						}
					}
				}
			}
			catch
			{
				num = 0;
			}
		}
		return num;
	}

	private int CheckFileSignatures(DownloadFileInfo[] dfile_info_list, string[] temp_file_paths)
	{
		int result = 0;
		try
		{
			for (int i = 0; i < temp_file_paths.Length; i++)
			{
				byte[] fileSignature = dfile_info_list[i].GetFileSignature();
				if (!RSAUtil.VerifyFileSignature(temp_file_paths[i], fileSignature, public_key))
				{
					result = -1;
					break;
				}
			}
		}
		catch
		{
			result = -1;
		}
		return result;
	}

	private int CountBlocksDownloaded(int last_file_list_ind, int last_block_ind, out int num_blocks, out int blocks_size)
	{
		int num = 0;
		int num2 = 0;
		int result = 0;
		try
		{
			for (int i = 0; i <= last_file_list_ind; i++)
			{
				num += dfile_info_list[i].GetNumBlocks();
				num2 += (int)dfile_info_list[i].GetFileSize();
			}
			num = num + last_block_ind + 1;
			int num3;
			try
			{
				num3 = (last_block_ind + 1) * dfile_info_list[last_file_list_ind + 1].GetBlockLen();
			}
			catch
			{
				num3 = 0;
			}
			num2 += num3;
		}
		catch
		{
			num = -1;
			num2 = -1;
			result = -1;
		}
		num_blocks = num;
		blocks_size = num2;
		return result;
	}

	private int DownloadBlocks(string file_server_uri, string session_id, string clientGuid, string tax_code, int size, string curr_version, bool is_last_file, bool is_first_file, DownloadFileInfo file_info, int file_num_blocks, int begin_block_ind, int retry_num)
	{
		int result = 0;
		down_prog.SetDownFname(file_info.GetFileName());
		worker.ReportProgress(1, down_prog);
		for (int i = begin_block_ind; i < file_num_blocks; i++)
		{
			if (worker.CancellationPending)
			{
				errno = -14;
				event_args.Cancel = true;
				result = -1;
				break;
			}
			curr_block_ind = i;
			string file_server_uri2;
			if (file_server_uri == null)
			{
				TryingLookupServer(out file_server_uri2, out var _);
			}
			else
			{
				file_server_uri2 = file_server_uri;
			}
			if (file_server_uri2 == null)
			{
				errno = -15;
				result = -1;
				break;
			}
			bool is_first_block = false;
			if (i == begin_block_ind)
			{
				is_first_block = true;
			}
			if (TryingDownloadSingleBlock(session_id, clientGuid, tax_code, size, curr_version, file_server_uri2, file_info, i, is_last_file, is_first_block, is_first_file) != 0)
			{
				result = -1;
				break;
			}
		}
		return result;
	}

	private void TryStartSesseion(out string session_id, out int balancing_mode)
	{
		int num = 0;
		bool flag = false;
		session_id = null;
		balancing_mode = -1;
		while (num <= retry_num && !flag)
		{
			StartSession(out session_id, out balancing_mode);
			if (session_id == null)
			{
				flag = true;
			}
			else
			{
				num++;
			}
		}
		if (!flag)
		{
			session_id = null;
		}
	}

	private void StartSession(out string session_id, out int balancing_mode)
	{
		AdminServerConnector adminServerConnector = new AdminServerConnector(admin_endpoint_uri);
		if (proxy_addr != "")
		{
			if (proxy_user != "")
			{
				adminServerConnector.SetProxy(proxy_addr, proxy_port, proxy_user, proxy_passwd);
			}
			else
			{
				adminServerConnector.SetProxy(proxy_addr, proxy_port);
			}
		}
		session_id = adminServerConnector.StartDownloadSession(out balancing_mode);
		err_msg = adminServerConnector.GetErrorMsg();
	}

	private void TryingLookupServer(out string file_server_uri, out string lastest_version)
	{
		int num = 0;
		bool flag = false;
		file_server_uri = null;
		lastest_version = null;
		while (num <= retry_num && !flag)
		{
			LookupFileServer(out file_server_uri, out lastest_version);
			if (file_server_uri != null && lastest_version != null)
			{
				flag = true;
			}
			else
			{
				num++;
			}
		}
		if (!flag)
		{
			file_server_uri = null;
			lastest_version = null;
		}
	}

	private void LookupFileServer(out string file_server_uri, out string lastest_version)
	{
		Debug.Print("Lookup File server");
		AdminServerConnector adminServerConnector = new AdminServerConnector(admin_endpoint_uri);
		if (proxy_addr != "")
		{
			if (proxy_user != "")
			{
				adminServerConnector.SetProxy(proxy_addr, proxy_port, proxy_user, proxy_passwd);
			}
			else
			{
				adminServerConnector.SetProxy(proxy_addr, proxy_port);
			}
		}
		adminServerConnector.LookupServer(out file_server_uri, out lastest_version);
		err_msg = adminServerConnector.GetErrorMsg();
	}

	private int TryingDownloadSingleBlock(string session_id, string clientGuid, string tax_code, int size, string curr_version, string file_server_uri, DownloadFileInfo file_info, int block_ind, bool is_last_file, bool is_first_block, bool is_first_file)
	{
		int result = 0;
		int num = 0;
		bool flag = false;
		while (num <= retry_num && !flag)
		{
			if (DownloadSingleBlock(session_id, clientGuid, tax_code, size, curr_version, file_server_uri, file_info, block_ind, is_last_file, is_first_block, is_first_file) == 0)
			{
				flag = true;
			}
			else
			{
				num++;
			}
		}
		if (!flag)
		{
			result = -1;
		}
		return result;
	}

	private int DownloadSingleBlock(string session_id, string clientGuid, string tax_code, int size, string curr_version, string file_server_uri, DownloadFileInfo file_info, int block_ind, bool is_last_file, bool is_first_block, bool is_first_file)
	{
		int result = 0;
		try
		{
			FileServerConnector fileServerConnector = new FileServerConnector(file_server_uri);
			if (proxy_addr != "")
			{
				if (proxy_user != "")
				{
					fileServerConnector.SetProxy(proxy_addr, proxy_port, proxy_user, proxy_passwd);
				}
				else
				{
					fileServerConnector.SetProxy(proxy_addr, proxy_port);
				}
			}
			int fileId = file_info.GetFileId();
			if (fileServerConnector.DownloadBlock(session_id, clientGuid, tax_code, size, curr_version, fileId, block_ind, is_last_file, is_first_block, is_first_file, out var bytes, out var blk_sign) != 0)
			{
				err_msg = fileServerConnector.GetErrorMsg();
				errno = -20;
				result = -1;
			}
			else if (bytes == null || blk_sign == null)
			{
				errno = -21;
				result = -1;
			}
			else if (RSAUtil.VerifySignature(bytes, blk_sign, public_key))
			{
				if (block_buffer.SaveBlockIntoFile(bytes, fileId, block_ind) != 0)
				{
					errno = -19;
					result = -1;
				}
				else
				{
					downloaded_blocks_num++;
					downloaded_size += bytes.Length;
					int percentProgress = downloaded_blocks_num * 100 / total_blocks;
					down_prog.SetDownloadedSize(downloaded_size);
					worker.ReportProgress(percentProgress, down_prog);
				}
			}
			else
			{
				errno = -18;
				result = -1;
			}
		}
		catch (Exception ex)
		{
			Debug.Print(ex.Message);
			result = -1;
		}
		return result;
	}
}
