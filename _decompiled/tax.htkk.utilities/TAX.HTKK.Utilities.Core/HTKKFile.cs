using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Utilities.Core;

public class HTKKFile
{
	public static string ErrorMessage = "";

	public static bool InVisualStudio()
	{
		Process[] processes = Process.GetProcesses();
		Process[] array = processes;
		foreach (Process process in array)
		{
			if (process.ProcessName.Equals("devenv"))
			{
				return true;
			}
		}
		return false;
	}

	public static void Log(string fnName, Exception ex)
	{
		string empty = string.Empty;
		empty = empty + Environment.NewLine + "Product name :" + HTKKConfig.PRODUCT_NAME;
		empty = empty + Environment.NewLine + "Version :" + HTKKConfig.VERSION;
		string text = empty;
		empty = text + Environment.NewLine + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + ":" + fnName;
		empty = empty + Environment.NewLine + "Message :" + ex.Message;
		empty = empty + Environment.NewLine + "StackTrace : " + ex.StackTrace;
		DateTime now = DateTime.Now;
		StreamWriter streamWriter = null;
		string path = now.ToString("yyyy-MM-dd") + ".log";
		if (!Directory.Exists(HTKKConfig.LOGPATH))
		{
			Directory.CreateDirectory(HTKKConfig.LOGPATH);
		}
		string path2 = Path.Combine(HTKKConfig.LOGPATH, path);
		try
		{
			streamWriter = new StreamWriter(path2, append: true);
			streamWriter.WriteLine(empty);
		}
		catch
		{
		}
		finally
		{
			if (streamWriter != null)
			{
				streamWriter.Dispose();
				streamWriter.Close();
			}
		}
	}

	public static void Log(string message)
	{
		DateTime now = DateTime.Now;
		StreamWriter streamWriter = null;
		string path = now.ToString("yyyy-MM-dd") + ".log";
		if (!Directory.Exists(HTKKConfig.LOGPATH))
		{
			Directory.CreateDirectory(HTKKConfig.LOGPATH);
		}
		string path2 = Path.Combine(HTKKConfig.LOGPATH, path);
		try
		{
			streamWriter = new StreamWriter(path2, append: true);
			streamWriter.WriteLine(message);
		}
		catch
		{
		}
		finally
		{
			if (streamWriter != null)
			{
				streamWriter.Dispose();
				streamWriter.Close();
			}
		}
	}

	public static void SimpleLogException(string message)
	{
		StreamWriter streamWriter = null;
		string path = DateTime.Now.ToString("yyyy-MM-dd") + ".log";
		if (!Directory.Exists(HTKKConfig.LOGPATH))
		{
			Directory.CreateDirectory(HTKKConfig.LOGPATH);
		}
		string path2 = Path.Combine(HTKKConfig.LOGPATH, path);
		try
		{
			streamWriter = new StreamWriter(path2, append: true);
			streamWriter.WriteLine(message);
		}
		catch
		{
		}
		finally
		{
			if (streamWriter != null)
			{
				streamWriter.Dispose();
				streamWriter.Close();
			}
		}
	}

	public static void LogRun(string fnName, string ex)
	{
		string empty = string.Empty;
		string text = empty;
		empty = text + Environment.NewLine + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + ": " + fnName + " - " + ex;
		DateTime now = DateTime.Now;
		StreamWriter streamWriter = null;
		string path = "log_run_" + now.ToString("yyyy-MM-dd") + ".log";
		if (!Directory.Exists(HTKKConfig.LOGPATH))
		{
			Directory.CreateDirectory(HTKKConfig.LOGPATH);
		}
		string path2 = Path.Combine(HTKKConfig.LOGPATH, path);
		try
		{
			streamWriter = new StreamWriter(path2, append: true);
			streamWriter.WriteLine(empty);
		}
		catch
		{
		}
		finally
		{
			if (streamWriter != null)
			{
				streamWriter.Dispose();
				streamWriter.Close();
			}
		}
	}

	public static void saveFile(string pathFolder, string fileName, string content)
	{
		StreamWriter streamWriter = null;
		string path = Path.Combine(pathFolder, fileName);
		try
		{
			streamWriter = new StreamWriter(path, append: true);
			streamWriter.WriteLine(content);
		}
		catch
		{
		}
		finally
		{
			if (streamWriter != null)
			{
				streamWriter.Dispose();
				streamWriter.Close();
			}
		}
	}

	public static string GetFileName(string pathFile)
	{
		try
		{
			return "";
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static bool WriteFile()
	{
		try
		{
			return true;
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static string LoadFile(string pathFile)
	{
		try
		{
			StreamReader streamReader = new StreamReader(pathFile);
			string text = streamReader.ReadToEnd();
			string[] array = Regex.Split("<!DOCTYPE Sections SYSTEM \"..\\..\\InterfaceTemplates\\Schema.dtd\">###<!DOCTYPE Sections SYSTEM \"..\\..\\InterfaceTemplates\\Schema.dtd\"[]>", "###");
			for (int i = 0; i < array.Length; i++)
			{
				text = text.Replace(array[i], "");
			}
			streamReader.Close();
			return text;
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static bool Export(string maToKhai, string fileName, string pathFile)
	{
		try
		{
			return true;
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool Import(string maToKhai, string pathFile)
	{
		try
		{
			return true;
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool Backup(string fileName, string pathFile)
	{
		try
		{
			return true;
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool Restore(string pathFile)
	{
		try
		{
			return true;
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool CopyFile(string sSourcePath, string sPathEnd, string logPath)
	{
		try
		{
			File.SetAttributes(sSourcePath, FileAttributes.Normal);
			FileInfo fileInfo = new FileInfo(sSourcePath);
			fileInfo.CopyTo(sPathEnd, overwrite: true);
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
		return true;
	}

	public static bool DeleteFile(string filePath)
	{
		try
		{
			if (File.Exists(filePath))
			{
				using FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
				fileStream.Dispose();
				File.SetAttributes(filePath, FileAttributes.Normal);
				File.Delete(filePath);
			}
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
		return true;
	}

	public static DataTable GetDataTableFilesInDirectory(string pathDirectory)
	{
		try
		{
			DataTable dataTable = new DataTable();
			dataTable.Columns.Add("FileName", typeof(string));
			dataTable.Columns.Add("PathFile", typeof(string));
			if (Directory.Exists(pathDirectory))
			{
				DirectoryInfo directoryInfo = new DirectoryInfo(pathDirectory);
				FileInfo[] files = directoryInfo.GetFiles("*.xml");
				foreach (FileInfo fileInfo in files)
				{
					DataRow dataRow = dataTable.NewRow();
					dataRow["FileName"] = fileInfo.Name;
					dataRow["PathFile"] = fileInfo.FullName;
					dataTable.Rows.Add(dataRow);
				}
			}
			return dataTable;
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static int ColumnLettersToIndex(string column)
	{
		try
		{
			int num = 0;
			int length = column.Length;
			for (int i = 0; i < length; i++)
			{
				num += (column[length - i - 1] - 65 + 1) * (int)Math.Pow(26.0, i);
			}
			return num;
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
			return -1;
		}
	}

	public static string ColumnIndexToLetters(int value)
	{
		string text = string.Empty;
		while (--value >= 0)
		{
			text = (char)(65 + value % 26) + text;
			value /= 26;
		}
		return text;
	}

	public static string GetSession(int lengthSession)
	{
		string text = "0123456789";
		char[] array = new char[lengthSession];
		Random random = new Random();
		for (int i = 0; i < array.Length; i++)
		{
			array[i] = text[random.Next(text.Length)];
		}
		return new string(array);
	}

	public static string GetMessage(string IDmsg)
	{
		try
		{
			XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
			return xMLMessage.GetMessage(IDmsg);
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
			return ex.ToString();
		}
	}

	public static string GetMessage(SheetView sView, int[] cell)
	{
		try
		{
			object value = sView.FpSpread.Sheets[sView.FpSpread.Sheets.Count - 1].Cells[cell[1], cell[0]].Value;
			return (value == null) ? "" : value.ToString();
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetMessage(FpSpread winFpSpread, int idxRow, int idxCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets["Header"];
			string text = "";
			return sheetView.Cells[idxRow, idxCol].Text;
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
			return ex.ToString();
		}
	}

	public static void SetNodeBackGroundPink(SheetView sView, int rowIndex, int colIndex, string msg, int rowSTT = 0)
	{
		try
		{
			sView.Cells[rowIndex, colIndex].BackColor = Color.Pink;
			sView.SetNote(rowIndex, colIndex, msg);
			sView.Cells[rowIndex, colIndex].NoteStyle = NoteStyle.PopupNote;
			sView.Cells[rowIndex, colIndex].NoteIndicatorSize = new Size(5, 5);
			string message = $"[{sView.SheetName}: {rowSTT}] {msg}: {sView.Cells[rowIndex, colIndex].Text}";
			LogErrorBigData(message);
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetNodeBackGroundPink2(SheetView sView, int rowIndex, int colIndex, string msg, string rowSTT = "")
	{
		try
		{
			sView.Cells[rowIndex, colIndex].BackColor = Color.Pink;
			sView.SetNote(rowIndex, colIndex, msg);
			sView.Cells[rowIndex, colIndex].NoteStyle = NoteStyle.PopupNote;
			sView.Cells[rowIndex, colIndex].NoteIndicatorSize = new Size(5, 5);
			string message = $"[{sView.SheetName}: {rowSTT}] {msg}: {sView.Cells[rowIndex, colIndex].Text}";
			LogErrorBigData(message);
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void RemoveValidateBackGround(SheetView sView, int rowIndex, int colIndex)
	{
		try
		{
			sView.Cells[rowIndex, colIndex].NoteStyle = NoteStyle.Hidden;
			sView.Cells[rowIndex, colIndex].ResetBackColor();
			sView.Cells[rowIndex, colIndex].ResetNote();
			sView.Cells[rowIndex, colIndex].BackColor = Color.FromArgb(233, 245, 254);
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void RemoveValidateBackGround1(SheetView sView, int rowIndex, int colIndex)
	{
		try
		{
			sView.Cells[rowIndex, colIndex].NoteStyle = NoteStyle.Hidden;
			sView.Cells[rowIndex, colIndex].ResetBackColor();
			sView.Cells[rowIndex, colIndex].ResetNote();
			sView.Cells[rowIndex, colIndex].BackColor = Color.FromArgb(240, 240, 240);
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetNodeBackGroundYellow(SheetView sView, int rowIndex, int colIndex, string msg)
	{
		try
		{
			if (sView.Cells[rowIndex, colIndex].BackColor != Color.Pink)
			{
				sView.Cells[rowIndex, colIndex].BackColor = Color.LemonChiffon;
				sView.SetNote(rowIndex, colIndex, msg);
				sView.Cells[rowIndex, colIndex].NoteStyle = NoteStyle.PopupNote;
				sView.Cells[rowIndex, colIndex].NoteIndicatorSize = new Size(5, 5);
			}
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void RemoveValidate(SheetView sView, int rowIndex, int colIndex, bool testColor = false)
	{
		try
		{
			if (sView.Cells[rowIndex, colIndex].Locked && testColor)
			{
				sView.Cells[rowIndex, colIndex].BackColor = Color.FromArgb(233, 245, 254);
			}
			else
			{
				sView.Cells[rowIndex, colIndex].BackColor = Color.White;
			}
			sView.Cells[rowIndex, colIndex].NoteStyle = NoteStyle.Hidden;
			sView.Cells[rowIndex, colIndex].ResetNote();
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void Create_Folder(string pathFolder)
	{
		try
		{
			if (!Directory.Exists(pathFolder))
			{
				Directory.CreateDirectory(pathFolder);
			}
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void Delete_Folder(string pathFolder)
	{
		try
		{
			if (Directory.Exists(pathFolder))
			{
				DirectoryInfo directoryInfo = new DirectoryInfo(pathFolder);
				directoryInfo.Delete(recursive: true);
			}
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static bool IsEmptyFolder(string pathFolder)
	{
		bool result = true;
		try
		{
			if (Directory.Exists(pathFolder))
			{
				DirectoryInfo directoryInfo = new DirectoryInfo(pathFolder);
				FileInfo[] files = directoryInfo.GetFiles("*.xml");
				if (files.Length > 0)
				{
					result = false;
				}
			}
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static string GetValueByAttributes(string pathFileXml, string pathAttributes)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		try
		{
			XmlDocument val = new XmlDocument();
			if (!File.Exists(pathFileXml))
			{
				return "";
			}
			val.Load(pathFileXml);
			XmlNodeList val2 = ((XmlNode)val).SelectNodes(pathAttributes);
			if (val2.Count > 0)
			{
				return ((XmlNode)val2[0].Attributes["Value"]).Value;
			}
			return "";
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public static string GetValueByNodeName(string pathFileXml, string pathNode)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		try
		{
			XmlDocument val = new XmlDocument();
			val.LoadXml(LoadFile(pathFileXml));
			XmlNode val2 = ((XmlNode)val).SelectSingleNode(pathNode);
			if (val2 != null)
			{
				return val2.InnerText;
			}
			return "";
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public static string GetStrByLengh(string strInput, int strLengh)
	{
		try
		{
			string text = strInput.Trim();
			if (HTKKFormater.isNullString(text))
			{
				return null;
			}
			if (text.Length > strLengh)
			{
				return text;
			}
			return text.PadLeft(strLengh, '0');
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static bool IsFileLocked(FileInfo file)
	{
		FileStream fileStream = null;
		try
		{
			fileStream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
		}
		catch (IOException)
		{
			return true;
		}
		finally
		{
			fileStream?.Close();
		}
		return false;
	}

	public static void RegisterDll(string filePath)
	{
		try
		{
			Process process = new Process();
			process.StartInfo.FileName = "regsvr32.exe";
			process.StartInfo.Arguments = filePath;
			process.StartInfo.UseShellExecute = false;
			process.StartInfo.CreateNoWindow = true;
			process.StartInfo.RedirectStandardOutput = true;
			process.Start();
			process.WaitForExit();
			process.Close();
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void AddOrUpdate(Dictionary<object, object> dic, object key, object newValue)
	{
		try
		{
			if (dic.TryGetValue(key, out var _))
			{
				dic[key] = newValue;
			}
			else
			{
				dic.Add(key, newValue);
			}
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void AddOrUpdateDicStr(Dictionary<string, string> dic, string key, string newValue)
	{
		try
		{
			if (dic.TryGetValue(key, out var _))
			{
				dic[key] = newValue;
			}
			else
			{
				dic.Add(key, newValue);
			}
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static string ReadFile(string pathFile)
	{
		try
		{
			return File.ReadAllText(pathFile);
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public static Dictionary<string, string> GetConfigFile()
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Expected O, but got Unknown
		//IL_0048: Unknown result type (might be due to invalid IL or missing references)
		//IL_004e: Expected O, but got Unknown
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		try
		{
			XmlDocument val = new XmlDocument();
			val.LoadXml(LoadFile("Config.xml"));
			XmlNodeList childNodes = ((XmlNode)val).SelectNodes("/Root")[0].ChildNodes;
			foreach (XmlNode item in childNodes)
			{
				XmlNode val2 = item;
				AddOrUpdateDicStr(dictionary, val2.Name, val2.InnerText);
			}
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return dictionary;
	}

	public static void EjecutarCMD(string Comando)
	{
		try
		{
			ProcessStartInfo processStartInfo = new ProcessStartInfo("cmd.exe");
			processStartInfo.Arguments = $"/c {Comando}";
			processStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
			processStartInfo.Verb = "runas";
			Process.Start(processStartInfo);
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static string getCellFormat(string id, int indexSheet, int indexTable, int typeBg)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Expected O, but got Unknown
		string result = "";
		try
		{
			XmlDocument val = new XmlDocument();
			val.LoadXml(LoadFile("Cell_Dynamic_Format.xml"));
			XmlNodeList val2 = null;
			switch (typeBg)
			{
			case 1:
				val2 = ((XmlNode)val).SelectNodes("/Tkhais/Tkhai[@id='" + id + "']/Sheet[@index='" + indexSheet + "'][@TableIndex='" + indexTable + "']/White");
				break;
			case 2:
				val2 = ((XmlNode)val).SelectNodes("/Tkhais/Tkhai[@id='" + id + "']/Sheet[@index='" + indexSheet + "'][@TableIndex='" + indexTable + "']/Bg233245254");
				break;
			case 3:
				val2 = ((XmlNode)val).SelectNodes("/Tkhais/Tkhai[@id='" + id + "']/Sheet[@index='" + indexSheet + "'][@TableIndex='" + indexTable + "']/NoWhite");
				break;
			}
			if (val2 == null || val2.Count != 1)
			{
				return null;
			}
			result = val2[0].InnerText;
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static void LogErrorBigData(string message)
	{
		DateTime now = DateTime.Now;
		StreamWriter streamWriter = null;
		string path = "error_big_data.log";
		if (!Directory.Exists(HTKKConfig.LOGPATH))
		{
			Directory.CreateDirectory(HTKKConfig.LOGPATH);
		}
		string path2 = Path.Combine(HTKKConfig.LOGPATH, path);
		try
		{
			streamWriter = new StreamWriter(path2, append: true);
			streamWriter.WriteLine(message);
		}
		catch
		{
		}
		finally
		{
			if (streamWriter != null)
			{
				streamWriter.Dispose();
				streamWriter.Close();
			}
		}
	}

	public static void deleteLogErrorBigData()
	{
		try
		{
			string text = HTKKConfig.LOGPATH + "\\error_big_data.log";
			if (File.Exists(text))
			{
				DeleteFile(text);
			}
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static string readLogErrorBigData()
	{
		string result = "";
		try
		{
			string text = HTKKConfig.LOGPATH + "\\error_big_data.log";
			if (File.Exists(text))
			{
				result = ReadFile(text);
			}
		}
		catch (Exception ex)
		{
			Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static string addSpace(int num)
	{
		return "";
	}
}
