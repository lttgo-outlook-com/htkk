using System;
using Foundation.Base;

namespace Foundation.Utility;

internal class IdGenerator : BaseObject
{
	private const int DEFAULT_MAX_COUNT = 99;

	private const int MAX_FOR_MAX_COUNT = 9999;

	private object id_lock;

	private int count;

	private int max_count;

	private int max_count_len;

	public IdGenerator(int max_count)
	{
		count = 0;
		if (max_count <= 0)
		{
			this.max_count = 99;
		}
		else if (max_count > 9999)
		{
			max_count = 9999;
		}
		else
		{
			this.max_count = max_count;
		}
		max_count_len = max_count.ToString().Length;
		id_lock = new object();
	}

	public int GenerateNewMsgId(out string new_id)
	{
		int num = 0;
		string text = "";
		try
		{
			text = $"{Guid.NewGuid()}";
		}
		catch (Exception ex)
		{
			num = -1;
			err_msg = ex.Message;
		}
		if (num != 0)
		{
			new_id = "";
			return -1;
		}
		int num2;
		lock (id_lock)
		{
			num2 = GetCount();
		}
		try
		{
			int length = num2.ToString().Length;
			string text2 = "";
			if (length < max_count_len)
			{
				int num3 = max_count_len - length;
				for (int i = 0; i < num3; i++)
				{
					text2 += "0";
				}
			}
			new_id = text2 + num2 + "_" + text;
		}
		catch
		{
			new_id = "";
			num = -1;
		}
		return num;
	}

	public int GenerateNewMsgId(out string new_id, out int curr_count)
	{
		int num = 0;
		string text = "";
		try
		{
			text = $"{Guid.NewGuid()}";
		}
		catch (Exception ex)
		{
			num = -1;
			err_msg = ex.Message;
		}
		if (num != 0)
		{
			new_id = "";
			curr_count = -1;
			return -1;
		}
		int num2;
		lock (id_lock)
		{
			num2 = GetCount();
		}
		try
		{
			new_id = num2 + "_" + text;
			curr_count = num2;
		}
		catch
		{
			new_id = "";
			curr_count = -1;
			num = -1;
		}
		return 0;
	}

	private int GetCount()
	{
		int result = count;
		if (count >= max_count)
		{
			count = 0;
		}
		else
		{
			count++;
		}
		return result;
	}
}
