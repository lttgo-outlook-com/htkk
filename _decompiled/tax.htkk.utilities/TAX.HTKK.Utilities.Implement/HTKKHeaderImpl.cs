using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Xml;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Utilities.Implement;

public class HTKKHeaderImpl : HTKKBase<HTKKHeaderInfo>
{
	public bool WriteXml(HTKKHeaderInfo info, string pathFile)
	{
		try
		{
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public HTKKHeaderInfo LoadInfo(string xmlContent)
	{
		try
		{
			HTKKHeaderInfo info = new HTKKHeaderInfo();
			HTKKBase<HTKKHeaderInfo>.XmlToObject(xmlContent, ref info);
			return info;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDictHeaderInfo(string maSoThue, string CellID)
	{
		//IL_0038: Unknown result type (might be due to invalid IL or missing references)
		//IL_003e: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string text = HTKKConfig.DATAFILES + "\\" + maSoThue + "\\Header.xml";
			if (File.Exists(text))
			{
				string text2 = HTKKFile.LoadFile(text);
				XmlDocument val = new XmlDocument();
				val.LoadXml(text2);
				XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
				for (int i = 0; i < elementsByTagName.Count; i++)
				{
					string value = ((XmlNode)elementsByTagName[i].Attributes[CellID]).Value;
					string value2 = ((XmlNode)elementsByTagName[i].Attributes["Value"]).Value;
					dictionary.Add(value, value2);
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

	public void SetHeaderInfo(ref HTKKHeaderInfo info)
	{
		//IL_0031: Unknown result type (might be due to invalid IL or missing references)
		//IL_0037: Expected O, but got Unknown
		//IL_0068: Unknown result type (might be due to invalid IL or missing references)
		//IL_006f: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.DATAFILES + "\\" + info.mst + "\\Header.xml";
			if (!File.Exists(text))
			{
				return;
			}
			XmlDocument val = new XmlDocument();
			val.LoadXml(HTKKFile.LoadFile(text));
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				if (!string.IsNullOrEmpty(((XmlNode)val2.Attributes["PropertyInfo"]).Value))
				{
					string value = ((XmlNode)val2.Attributes["PropertyInfo"]).Value;
					string value2 = ((XmlNode)val2.Attributes["Value"]).Value;
					dictionary.Add(value, value2);
				}
			}
			Type typeFromHandle = typeof(HTKKHeaderInfo);
			PropertyInfo[] properties = typeFromHandle.GetProperties();
			PropertyInfo[] array = properties;
			foreach (PropertyInfo propertyInfo in array)
			{
				if (dictionary.ContainsKey(propertyInfo.Name))
				{
					propertyInfo.SetValue(info, dictionary[propertyInfo.Name], null);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			info = null;
		}
	}

	public void SetPropertyHeader(ref HTKKHeaderInfo info, string maSoThue)
	{
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Expected O, but got Unknown
		//IL_006b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0072: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.DATAFILES + "\\" + maSoThue + "\\HEADER.xml";
			if (!File.Exists(text))
			{
				return;
			}
			info.mst = maSoThue;
			XmlDocument val = new XmlDocument();
			val.LoadXml(HTKKFile.LoadFile(text));
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				if (!string.IsNullOrEmpty(((XmlNode)val2.Attributes["PropertyInfo"]).Value))
				{
					string value = ((XmlNode)val2.Attributes["PropertyInfo"]).Value;
					string value2 = ((XmlNode)val2.Attributes["Value"]).Value;
					dictionary.Add(value, value2);
				}
			}
			Type typeFromHandle = typeof(HTKKHeaderInfo);
			PropertyInfo[] properties = typeFromHandle.GetProperties();
			PropertyInfo[] array = properties;
			foreach (PropertyInfo propertyInfo in array)
			{
				if (dictionary.ContainsKey(propertyInfo.Name))
				{
					propertyInfo.SetValue(info, dictionary[propertyInfo.Name], null);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			info = null;
		}
	}
}
