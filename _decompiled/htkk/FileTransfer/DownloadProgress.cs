namespace FileTransfer;

internal class DownloadProgress
{
	private int total_num_files;

	private int prev_num_downloaded_files;

	private int just_num_downloaded_files;

	private int total_size;

	private int just_downloaded_size;

	private int prev_just_downloaded_size;

	private string downloading_fname;

	private int progress_result;

	private string err_msg;

	public DownloadProgress()
	{
		err_msg = "";
		downloading_fname = "";
	}

	public DownloadProgress(int total_num_files, int prev_num_downloaded_files, int just_num_downloaded_files, int total_size, int prev_just_downloaded_size, int just_downloaded_size)
	{
		this.total_num_files = total_num_files;
		this.just_num_downloaded_files = just_num_downloaded_files;
		this.total_size = total_size;
		this.just_downloaded_size = just_downloaded_size;
		this.prev_just_downloaded_size = prev_just_downloaded_size;
		this.prev_num_downloaded_files = prev_num_downloaded_files;
		err_msg = "";
	}

	public void SetProgressReturn(int ret)
	{
		progress_result = ret;
	}

	public int GetProgressReturn()
	{
		return progress_result;
	}

	public void SetErrMsg(string msg)
	{
		if (msg == null)
		{
			msg = "";
		}
		err_msg = msg;
	}

	public string GetErrMsg()
	{
		return err_msg;
	}

	public void SetDownFname(string fname)
	{
		downloading_fname = fname;
	}

	public string GetDownFname()
	{
		return downloading_fname;
	}

	public void SetTotalSize(int total_size)
	{
		this.total_size = total_size;
	}

	public int GetTotalSize()
	{
		return total_size;
	}

	public void SetTotalNumFiles(int total_num_files)
	{
		this.total_num_files = total_num_files;
	}

	public int GetTotalNumFiles()
	{
		return total_num_files;
	}

	public void SetPrevNumFiles(int prev_num_downloaded_files)
	{
		this.prev_num_downloaded_files = prev_num_downloaded_files;
	}

	public void SetPrevSize(int prev_just_downloaded_size)
	{
		this.prev_just_downloaded_size = prev_just_downloaded_size;
	}

	public int GetDownloadedSize()
	{
		return just_downloaded_size;
	}

	public void SetDownloadedSize(int just_downloaded_size)
	{
		this.just_downloaded_size = just_downloaded_size;
	}

	public void SetDownloadedNumFiles(int just_num_downloaded_files)
	{
		this.just_num_downloaded_files = just_num_downloaded_files;
	}

	public int GetDownloadedNumFiles()
	{
		return just_num_downloaded_files;
	}

	public void IncreaseDownloadedNumFiles()
	{
		just_num_downloaded_files++;
	}
}
