using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class PrcessExportXML
{
	public static void SetThongTinDichVu(ref string xmlContent)
	{
		try
		{
			TTinDVuInfo tTinDVuInfo = new TTinDVuInfo();
			tTinDVuInfo.MaDVu = HTKKConfig.MA_DICH_VU;
			tTinDVuInfo.TenDVu = HTKKConfig.TEN_DICH_VU;
			tTinDVuInfo.PbanDVu = HTKKConfig.VERSION;
			tTinDVuInfo.TtinNhaCCapDVu = HTKKConfig.THONG_TIN_NHA_CUNG_CAP;
			PropertyInfo[] properties = typeof(TTinDVuInfo).GetProperties();
			foreach (PropertyInfo propertyInfo in properties)
			{
				string text = propertyInfo.Name.Substring(0, 1).ToLower() + propertyInfo.Name.Substring(1, propertyInfo.Name.Length - 1);
				string text2 = HTKKProperty.GetPropertyValue(tTinDVuInfo, propertyInfo.Name).ToString();
				xmlContent = Regex.Replace(xmlContent, "<" + text + "></" + text + ">", "<" + text + ">" + text2 + "</" + text + ">");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static Dictionary<string, List<DictInfo>> GetThongTinToKhaiFromXml(XmlDocument xmlData, XmlDocument xmlConfigCellIndex, HTKKSheetInfo sheetInfo)
	{
		//IL_002f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0035: Expected O, but got Unknown
		//IL_0084: Unknown result type (might be due to invalid IL or missing references)
		//IL_008b: Expected O, but got Unknown
		//IL_00cb: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d2: Expected O, but got Unknown
		try
		{
			Dictionary<string, List<DictInfo>> dictionary = new Dictionary<string, List<DictInfo>>();
			XmlNodeList elementsByTagName = xmlConfigCellIndex.GetElementsByTagName("Static");
			List<DictInfo> list = new List<DictInfo>();
			foreach (XmlNode item2 in elementsByTagName)
			{
				XmlNode val = item2;
				string value = ((XmlNode)val.Attributes["GroupName"]).Value;
				if (!(value != "TKhaiThue"))
				{
					continue;
				}
				XmlNodeList childNodes = val.ChildNodes;
				foreach (XmlNode item3 in childNodes)
				{
					XmlNode val2 = item3;
					string value2 = string.Empty;
					if (string.IsNullOrEmpty(val2.InnerText))
					{
						continue;
					}
					XmlNodeList elementsByTagName2 = xmlData.GetElementsByTagName(val2.Name);
					foreach (XmlNode item4 in elementsByTagName2)
					{
						XmlNode val3 = item4;
						if (val3.LastChild != null)
						{
							value2 = val3.LastChild.InnerText;
						}
					}
					DictInfo dictInfo = new DictInfo();
					dictInfo.Key = val2.InnerText;
					dictInfo.Value = value2;
					DictInfo item = dictInfo;
					list.Add(item);
				}
			}
			dictionary.Add(sheetInfo.SheetIndex.ToString(), list);
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static object[,] GetThongTinToKhaiFromXml(XmlDocument xmlData, XmlDocument xmlConfigCellIndex)
	{
		//IL_008d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0094: Expected O, but got Unknown
		//IL_00b1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b8: Expected O, but got Unknown
		//IL_01de: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e5: Expected O, but got Unknown
		//IL_0228: Unknown result type (might be due to invalid IL or missing references)
		//IL_022f: Expected O, but got Unknown
		try
		{
			List<DictInfo> list = new List<DictInfo>();
			XmlNodeList elementsByTagName = xmlData.GetElementsByTagName("CoDieuChinhSoLieu");
			if (elementsByTagName.Count == 0)
			{
				DictInfo dictInfo = new DictInfo();
				dictInfo.Key = "check";
				dictInfo.Value = string.Empty;
				DictInfo item = dictInfo;
				list.Add(item);
			}
			string name = xmlConfigCellIndex.GetElementsByTagName("Dynamic")[0].FirstChild.Name;
			XmlNodeList elementsByTagName2 = xmlConfigCellIndex.GetElementsByTagName(name);
			foreach (XmlNode item2 in elementsByTagName2)
			{
				XmlNode val = item2;
				XmlNodeList childNodes = val.ChildNodes;
				foreach (XmlNode item3 in childNodes)
				{
					XmlNode val2 = item3;
					DictInfo dictInfo2 = new DictInfo();
					dictInfo2.Key = val2.Name;
					dictInfo2.Value = val2.InnerText;
					DictInfo item = dictInfo2;
					list.Add(item);
				}
			}
			object[,] array = null;
			string value = ((XmlNode)xmlConfigCellIndex.GetElementsByTagName("Dynamic")[0].Attributes["GroupName"]).Value;
			if (xmlData.GetElementsByTagName(value)[0].FirstChild != null)
			{
				string name2 = xmlData.GetElementsByTagName(value)[0].FirstChild.Name;
				XmlNodeList elementsByTagName3 = xmlData.GetElementsByTagName(name2);
				array = new object[elementsByTagName3.Count, list.Count];
				for (int i = 0; i < elementsByTagName3.Count; i++)
				{
					XmlNode val3 = elementsByTagName3[i];
					XmlDocument val4 = new XmlDocument();
					val4.LoadXml(val3.OuterXml);
					for (int j = 0; j < list.Count; j++)
					{
						DictInfo item = list[j];
						XmlNodeList elementsByTagName4 = val4.GetElementsByTagName(item.Key);
						foreach (XmlNode item4 in elementsByTagName4)
						{
							XmlNode val5 = item4;
							if (val5.LastChild != null)
							{
								string innerText = val5.LastChild.InnerText;
								array[i, j] = innerText;
							}
						}
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
}
