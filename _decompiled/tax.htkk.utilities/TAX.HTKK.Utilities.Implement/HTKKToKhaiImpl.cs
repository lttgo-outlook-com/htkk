using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Xml;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Utilities.Implement;

public class HTKKToKhaiImpl : HTKKBase<HTKKToKhaiInfo>
{
	private class MethodParameters
	{
		public DateTime ThreadCreated { get; set; }

		public int ThreadId { get; set; }

		public int RowIndexEnd { get; set; }

		public int RowIndex { get; set; }

		public List<string> ListColumnName { get; set; }
	}

	public static Dictionary<string, List<DictInfo>> DictDataObject = new Dictionary<string, List<DictInfo>>();

	private static int _threadCounter = 0;

	public bool WriteXml(HTKKToKhaiInfo info, string xmlContent, string pathFile)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		//IL_0101: Unknown result type (might be due to invalid IL or missing references)
		//IL_0107: Expected O, but got Unknown
		try
		{
			XmlDocument val = new XmlDocument();
			val.LoadXml(xmlContent);
			if (info.ListConfigCells != null && info.ListConfigCells.Count > 0)
			{
				foreach (KeyValuePair<string, string> listConfigCell in info.ListConfigCells)
				{
					XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section/Cells/Cell[@CellID='" + listConfigCell.Key.Split(new char[1] { ',' })[0] + "']");
					if (val2 != null && val2.Count > 0)
					{
						((XmlNode)val2[0].Attributes["Value"]).Value = listConfigCell.Value;
					}
				}
			}
			val.DocumentElement.SetAttribute("Version", info.Version);
			XmlTextWriter val3 = new XmlTextWriter(pathFile, Encoding.UTF8);
			try
			{
				val3.Formatting = (Formatting)1;
				val.Save((XmlWriter)(object)val3);
			}
			finally
			{
				((IDisposable)val3)?.Dispose();
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public Dictionary<string, List<DictInfo>> BindDataToDict(string pathDataFile, string strAttributes)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Expected O, but got Unknown
		//IL_0056: Unknown result type (might be due to invalid IL or missing references)
		//IL_005d: Expected O, but got Unknown
		//IL_0096: Unknown result type (might be due to invalid IL or missing references)
		//IL_009d: Expected O, but got Unknown
		Dictionary<string, List<DictInfo>> dictionary = new Dictionary<string, List<DictInfo>>();
		try
		{
			string text = HTKKFile.LoadFile(pathDataFile);
			XmlDocument val = new XmlDocument();
			val.LoadXml(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cells");
			if (elementsByTagName.Count > 0)
			{
				for (int i = 0; i < elementsByTagName.Count; i++)
				{
					string text2 = string.Empty;
					List<DictInfo> list = new List<DictInfo>();
					XmlDocument val2 = new XmlDocument();
					val2.LoadXml(elementsByTagName[i].OuterXml);
					XmlNodeList elementsByTagName2 = val2.GetElementsByTagName("Cell");
					foreach (XmlNode item2 in elementsByTagName2)
					{
						XmlNode val3 = item2;
						if (string.IsNullOrEmpty(text2))
						{
							text2 = ((XmlNode)val3.Attributes[strAttributes]).Value.Split(new char[1] { '_' })[1];
						}
						DictInfo dictInfo = new DictInfo();
						dictInfo.Key = ((XmlNode)val3.Attributes[strAttributes]).Value;
						dictInfo.Value = ((XmlNode)val3.Attributes["Value"]).Value;
						DictInfo item = dictInfo;
						list.Add(item);
					}
					dictionary.Add(text2, list);
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public object[,] BindDataToArrayObject(string pathDataFile)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Expected O, but got Unknown
		//IL_004d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0054: Expected O, but got Unknown
		object[,] array = null;
		try
		{
			int num = 0;
			int num2 = 0;
			string text = HTKKFile.LoadFile(pathDataFile);
			XmlDocument val = new XmlDocument();
			val.LoadXml(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cells");
			if (elementsByTagName.Count > 0)
			{
				for (int i = 0; i < elementsByTagName.Count; i++)
				{
					XmlDocument val2 = new XmlDocument();
					val2.LoadXml(elementsByTagName[i].OuterXml);
					XmlNodeList elementsByTagName2 = val2.GetElementsByTagName("Cell");
					if (num == 0 && num2 == 0)
					{
						num = elementsByTagName.Count;
						num2 = elementsByTagName2.Count;
						array = new object[num, num2];
					}
					for (int j = 0; j < elementsByTagName2.Count; j++)
					{
						array[i, j] = ((XmlNode)elementsByTagName2[j].Attributes["Value"]).Value;
					}
				}
			}
			return array;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void AppendCellSetingToDictData(ref int countRows, ref Dictionary<string, List<DictInfo>> dictData, int rowIndexEnd)
	{
		try
		{
			DictDataObject.Clear();
			int num = 0;
			string empty = string.Empty;
			List<string> list = new List<string>();
			foreach (KeyValuePair<string, List<DictInfo>> dictDatum in dictData)
			{
				num = int.Parse(dictDatum.Key);
				foreach (DictInfo item in dictDatum.Value)
				{
					list.Add(item.Key.Split(new char[1] { '_' })[0]);
				}
			}
			for (int i = num + 1; i < rowIndexEnd; i++)
			{
				List<DictInfo> list2 = new List<DictInfo>();
				foreach (string item2 in list)
				{
					DictInfo dictInfo = new DictInfo();
					dictInfo.Key = item2 + "_" + i;
					dictInfo.Value = string.Empty;
					DictInfo current2 = dictInfo;
					list2.Add(current2);
				}
				dictData.Add(i.ToString(), list2);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private static void CreateThread(int threadId, int rowIndex, int rowIndexEnd, List<string> listColumnName)
	{
		try
		{
			Console.WriteLine("Creating Thread {0}", threadId);
			MethodParameters methodParameters = new MethodParameters();
			methodParameters.ThreadId = threadId;
			methodParameters.ThreadCreated = DateTime.Now;
			methodParameters.RowIndex = rowIndex;
			methodParameters.RowIndexEnd = rowIndexEnd;
			methodParameters.ListColumnName = listColumnName;
			Interlocked.Increment(ref _threadCounter);
			ThreadPool.QueueUserWorkItem(LongRunningProcess, methodParameters);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private static void LongRunningProcess(object input)
	{
		MethodParameters methodParameters = (MethodParameters)input;
		Random random = new Random(((MethodParameters)input).ThreadId);
		try
		{
			for (int i = methodParameters.RowIndex; i < methodParameters.RowIndexEnd; i++)
			{
				List<DictInfo> list = new List<DictInfo>();
				foreach (string item2 in methodParameters.ListColumnName)
				{
					DictInfo dictInfo = new DictInfo();
					dictInfo.Key = item2 + "_" + i;
					dictInfo.Value = string.Empty;
					DictInfo item = dictInfo;
					list.Add(item);
				}
				DictDataObject.Add(i.ToString(), list);
			}
			Interlocked.Decrement(ref _threadCounter);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private static void StopProcessing()
	{
		try
		{
			while (_threadCounter > 0)
			{
				Thread.Sleep(500);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
