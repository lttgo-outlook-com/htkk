using System;
using System.IO;
using System.Windows.Forms;

namespace Foundation;

public class Log
{
	public static void LogEvent(string message)
	{
		string path = DateTime.Now.ToString("yyyy-MM-dd") + ".log";
		if (!Directory.Exists(Application.StartupPath + "\\Log"))
		{
			Directory.CreateDirectory(Application.StartupPath + "\\Log");
		}
		string path2 = Path.Combine(Application.StartupPath + "\\Log", path);
		try
		{
			using StreamWriter streamWriter = new StreamWriter(path2, append: true);
			streamWriter.WriteLine(message);
		}
		catch
		{
		}
	}

	public static void LogEvent(string source, Exception ex)
	{
		string empty = string.Empty;
		empty += "------------------------------------------------------------------------------------------------------------------------------------";
		string text = empty;
		empty = text + Environment.NewLine + "  \"Date\":\"" + DateTime.Now.ToString() + "\",";
		text = empty;
		empty = text + Environment.NewLine + "  \"Source\":\"[" + source + "]\",";
		text = empty;
		empty = text + Environment.NewLine + "  \"Message\":\"" + ex.Message.Replace("\\", "\\\\").Replace("\r\n", "\\r\\n").Replace("\"", "'") + "\",";
		text = empty;
		empty = text + Environment.NewLine + "  \"StackTrace\":\"" + ex.StackTrace.Replace("\\", "\\\\").Replace("\r\n", "\\r\\n").Replace("\"", "'") + "\",";
		empty = empty + Environment.NewLine + "------------------------------------------------------------------------------------------------------------------------------------";
		string path = DateTime.Now.ToString("yyyy-MM-dd") + ".log";
		if (!Directory.Exists(Application.StartupPath + "\\Log"))
		{
			Directory.CreateDirectory(Application.StartupPath + "\\Log");
		}
		string path2 = Path.Combine(Application.StartupPath + "\\Log", path);
		try
		{
			using StreamWriter streamWriter = new StreamWriter(path2, append: true);
			streamWriter.WriteLine(empty);
		}
		catch
		{
		}
	}
}
