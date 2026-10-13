namespace Foundation.Base;

public class BaseObject
{
	public const int EN_NO_ERROR = 0;

	public const int EN_ERROR = -10000;

	public const int EN_UNKNOWN = -10001;

	protected int errno;

	protected string err_msg;

	protected string monitor_msg;

	protected string tag;

	protected bool is_null;

	public BaseObject()
	{
		err_msg = "";
		monitor_msg = "";
		tag = "";
	}

	public int GetErrNo()
	{
		return errno;
	}

	public bool IsNullObject()
	{
		return is_null;
	}

	public void SetTag(string tag)
	{
		if (tag == null)
		{
			this.tag = "";
		}
		else
		{
			this.tag = tag;
		}
	}

	public string GetTag()
	{
		return tag;
	}

	public string GetErrorMsg()
	{
		if (err_msg == null)
		{
			return "";
		}
		return err_msg;
	}

	protected void SetError(int errno, string err_msg)
	{
		this.errno = errno;
		if (err_msg == null)
		{
			err_msg = "";
		}
		this.err_msg = err_msg;
	}

	protected void SetErrorNo(int errno)
	{
		this.errno = errno;
		err_msg = "";
	}

	protected void SetErrorMsg(string err_msg)
	{
		if (err_msg == null)
		{
			err_msg = "";
		}
		this.err_msg = err_msg;
		errno = -10000;
	}
}
