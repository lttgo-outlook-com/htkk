namespace Foundation.Base;

internal class FunctionReturn
{
	public const int RET_OK = 0;

	public const int RET_FAILED = -1;

	private string err_msg;

	public void SetErrMsg(string err_msg)
	{
		this.err_msg = err_msg;
	}

	public string GetErrMsg()
	{
		return err_msg;
	}
}
