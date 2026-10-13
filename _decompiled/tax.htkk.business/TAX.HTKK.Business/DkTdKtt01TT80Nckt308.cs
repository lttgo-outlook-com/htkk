using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkTdKtt01TT80Nckt308
{
	private Dictionary<string, string[]> dicDMNganhNgheKD = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMNganhNgheKDTmp = new Dictionary<string, string>();

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void PreBindData(FpSpread winFpSpread, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			if (!ToKhaiInfo.SheetInfos[0].IsReadDataFile)
			{
				int[] array = GeneralFunctions.ArrayCellId("J_33");
				sheetView.Cells[array[1], array[0]].Value = DateTime.Now.Year;
			}
			int[] array2 = GeneralFunctions.ArrayCellId("I_26");
			int[] array3 = GeneralFunctions.ArrayCellId("I_28");
			int[] array4 = GeneralFunctions.ArrayCellId("C_32");
			bool flag = ConvertHelpers.ConvertToBool(sheetView.Cells[array2[1], array2[0]].Value);
			bool flag2 = ConvertHelpers.ConvertToBool(sheetView.Cells[array3[1], array3[0]].Value);
			string text = "Cơ sở chúng tôi đề nghị được áp dụng kỳ tính thuế: ";
			if (flag && flag2)
			{
				text += "Giá trị gia tăng, Thu nhập cá nhân";
			}
			else if (flag)
			{
				text += "Giá trị gia tăng";
			}
			else if (flag2)
			{
				text += "Thu nhập cá nhân";
			}
			sheetView.Cells[array4[1], array4[0]].Text = text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("T_56");
			string text = sheetView.Cells[array[1], array[0]].Text;
			HTKKFile.RemoveValidate(sheetView, array[1], array[0]);
			if (!string.IsNullOrEmpty(text))
			{
				DateTime dateTime = ConvertHelpers.ConvertToDateTime(text, "dd/MM/yyyy");
				if (dateTime != DateTime.MinValue && dateTime > DateTime.Now.Date)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17")));
					value = 2;
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[0].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[0].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[0].BackColor = Color.White;
				break;
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		//IL_012c: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("I_26");
			int[] array2 = GeneralFunctions.ArrayCellId("I_28");
			int[] array3 = GeneralFunctions.ArrayCellId("C_32");
			bool flag = ConvertHelpers.ConvertToBool(sheetView.Cells[array[1], array[0]].Value);
			bool flag2 = ConvertHelpers.ConvertToBool(sheetView.Cells[array2[1], array2[0]].Value);
			string text = "Cơ sở chúng tôi đề nghị được áp dụng kỳ tính thuế: ";
			if (idxCol == array[0] && (idxRow == array[1] || idxRow == array2[1]))
			{
				if (flag && flag2)
				{
					text += "Giá trị gia tăng, Thu nhập cá nhân";
				}
				else if (flag)
				{
					text += "Giá trị gia tăng";
				}
				else if (flag2)
				{
					text += "Thu nhập cá nhân";
				}
				if (idxRow == array2[1] && flag2)
				{
					frmMessage frmMessage = new frmMessage("0639", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
				}
			}
			sheetView.Cells[array3[1], array3[0]].Text = text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public Dictionary<string, string> GetDMNganhNgheKD(string pathFile)
	{
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Expected O, but got Unknown
		//IL_0076: Unknown result type (might be due to invalid IL or missing references)
		//IL_007d: Expected O, but got Unknown
		try
		{
			if (dicDMNganhNgheKDTmp.Count > 0)
			{
				return dicDMNganhNgheKDTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(pathFile))
			{
				return null;
			}
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 2 && array[2] == "1")
				{
					string key = array[0];
					string value2 = array[1];
					dictionary2.Add(key, value2);
					dictionary.Add(array[1], array);
				}
			}
			dicDMNganhNgheKD = dictionary;
			dicDMNganhNgheKDTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			GeneralFunctions.FillDataMstCu(xmlNode);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
