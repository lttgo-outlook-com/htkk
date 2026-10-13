using System;
using System.Collections.Generic;
using System.Reflection;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace HTKK;

public class WinValidateToKhai
{
	public static void WinValidateSheet(HTKKSheetInfo sheetInfo, List<ValidateInfo> listValidate, FpSpread winFpSpread)
	{
		try
		{
			Dictionary<string, List<DictInfo>> dictionary = ValidateSheet(listValidate, winFpSpread);
			if (dictionary == null)
			{
				return;
			}
			if (!WinSettings.WinListMessageErorr.ContainsKey(sheetInfo.SheetIndex))
			{
				WinSettings.WinListMessageErorr.Add(sheetInfo.SheetIndex, dictionary);
			}
			foreach (KeyValuePair<string, List<DictInfo>> item in dictionary)
			{
				WinSpreadProcess.WinSetBackColor(winFpSpread.ActiveSheet, item.Key);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private static Dictionary<string, List<DictInfo>> ValidateSheet(List<ValidateInfo> listValidate, FpSpread winFpSpread)
	{
		try
		{
			Dictionary<string, List<DictInfo>> dictionary = new Dictionary<string, List<DictInfo>>();
			if (listValidate != null)
			{
				List<ReturnMessageInfo> list = new List<ReturnMessageInfo>();
				foreach (ValidateInfo item3 in listValidate)
				{
					if (!WinSettings.DictTypeValidate.ContainsKey(item3.TypeValidate))
					{
						continue;
					}
					string text = WinSpreadProcess.WinGetText(winFpSpread.ActiveSheet, item3.CellIndex);
					if (!string.IsNullOrEmpty(text))
					{
						if (item3.TypeValidate == "T012")
						{
							ReturnMessageInfo returnMessageInfo = new ReturnMessageInfo();
							returnMessageInfo.CellIndex = item3.CellIndex;
							returnMessageInfo.ValueCellIndex = text;
							returnMessageInfo.MessageError = item3.TypeValidate;
							returnMessageInfo.ValueMessageError = ((XmlNode)((XmlNode)WinSettings.XmlMessage).SelectNodes("/Message/Infor[@ID='" + item3.MessageError + "']")[0].Attributes["Msg"]).Value;
							ReturnMessageInfo item = returnMessageInfo;
							list.Add(item);
							continue;
						}
						if (!string.IsNullOrEmpty(item3.ListCellRange))
						{
							string[] array = item3.ListCellRange.Split(new char[1] { ',' });
							string[] array2 = array;
							foreach (string text2 in array2)
							{
								string newValue = WinSpreadProcess.WinGetText(winFpSpread.ActiveSheet, text2);
								item3.Range = item3.Range.Replace(text2, newValue);
							}
						}
						object[] stringParam = new object[2] { text, item3.Range };
						object obj = HTKKInvokeMethod.InvokeStringMethod("TAX.HTKK.Business", "TAX.HTKK.Business", "ValidateBase", WinSettings.DictTypeValidate[item3.TypeValidate], stringParam);
						if (!bool.Parse(obj.ToString()))
						{
							DictInfo dictInfo = new DictInfo();
							dictInfo.Key = item3.MessageError;
							dictInfo.Value = ((XmlNode)((XmlNode)WinSettings.XmlMessage).SelectNodes("/Message/Infor[@ID='" + item3.MessageError + "']")[0].Attributes["Msg"]).Value;
							DictInfo item2 = dictInfo;
							if (dictionary.ContainsKey(item3.CellIndex))
							{
								dictionary[item3.CellIndex].Add(item2);
								continue;
							}
							List<DictInfo> list2 = new List<DictInfo>();
							list2.Add(item2);
							dictionary.Add(item3.CellIndex, list2);
						}
					}
					else
					{
						DictInfo dictInfo2 = new DictInfo();
						dictInfo2.Key = "0359";
						dictInfo2.Value = ((XmlNode)((XmlNode)WinSettings.XmlMessage).SelectNodes("/Message/Infor[@ID='0359']")[0].Attributes["Msg"]).Value;
						DictInfo item2 = dictInfo2;
						if (dictionary.ContainsKey(item3.CellIndex))
						{
							dictionary[item3.CellIndex].Add(item2);
							continue;
						}
						List<DictInfo> list2 = new List<DictInfo>();
						list2.Add(item2);
						dictionary.Add(item3.CellIndex, list2);
					}
				}
				if (list.Count > 0)
				{
					object[] stringParam = new object[1] { list };
					object obj = HTKKInvokeMethod.InvokeStringMethod("TAX.HTKK.Business", "TAX.HTKK.Business", "ValidateBase", WinSettings.DictTypeValidate["T012"], stringParam);
					if (obj != null)
					{
						list = (List<ReturnMessageInfo>)obj;
						foreach (ReturnMessageInfo item4 in list)
						{
							DictInfo dictInfo3 = new DictInfo();
							dictInfo3.Key = item4.MessageError;
							dictInfo3.Value = item4.ValueMessageError;
							DictInfo item2 = dictInfo3;
							if (dictionary.ContainsKey(item4.CellIndex))
							{
								dictionary[item4.CellIndex].Add(item2);
								continue;
							}
							List<DictInfo> list2 = new List<DictInfo>();
							list2.Add(item2);
							dictionary.Add(item4.CellIndex, list2);
						}
					}
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
}
