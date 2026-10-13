using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Gtgt04TT80Nckt251
{
	private int[] cell11c = GeneralFunctions.ArrayCellId("G_27");

	private int[] cell11cMa = GeneralFunctions.ArrayCellId("AJ_27");

	private int[] cell11b = GeneralFunctions.ArrayCellId("S_27");

	private int[] cell11bMa = GeneralFunctions.ArrayCellId("AK_27");

	private int[] cell11a = GeneralFunctions.ArrayCellId("G_29");

	private int[] cell11aMa = GeneralFunctions.ArrayCellId("AJ_29");

	private int[] cell11DiaChi = GeneralFunctions.ArrayCellId("Q_25");

	private int[] cellct09 = GeneralFunctions.ArrayCellId("C_21");

	private int[] cellct10 = GeneralFunctions.ArrayCellId("N_23");

	private int[] cellCoPl = GeneralFunctions.ArrayCellId("AK_11");

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			_ToKhaiInfo = ToKhaiInfo;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			GeneralFunctions.ThayDoiTinhTheoCQT(headerInfo.maCQTCapCuc, headerInfo.maCQTNoiNop, sheetView, PeriodInfo, cell11c, cell11cMa, cell11b, cell11bMa, cell11a, cell11aMa, cell11DiaChi, cellct09, cellct10);
			sheetView.Cells[cellCoPl[1], cellCoPl[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			string text = "K_47###Q_47###Q_48###Q_49###Q_50";
			string text2 = "###";
			string[] array = text.Split(new string[1] { text2 }, StringSplitOptions.None);
			if (PeriodInfo.ThuHo.Equals("TH"))
			{
				sheetView.Cells[6, HTKKFile.ColumnLettersToIndex("H") - 1].Value = true;
				for (int i = 0; i < array.Length; i++)
				{
					int[] array2 = GeneralFunctions.ArrayCellId(array[i]);
					sheetView.Cells[array2[1], array2[0]].Locked = true;
				}
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				int column = HTKKFile.ColumnLettersToIndex("L") - 1;
				int column2 = HTKKFile.ColumnLettersToIndex("C") - 1;
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					sheetView.Cells[dataLocationTemplateRowStart + i, column].Locked = false;
					sheetView.Cells[dataLocationTemplateRowStart + i, column2].Locked = false;
				}
			}
			else
			{
				sheetView.Cells[6, HTKKFile.ColumnLettersToIndex("H") - 1].Value = false;
				for (int i = 0; i < array.Length; i++)
				{
					int[] array2 = GeneralFunctions.ArrayCellId(array[i]);
					sheetView.Cells[array2[1], array2[0]].Locked = false;
				}
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				int column = HTKKFile.ColumnLettersToIndex("L") - 1;
				int column2 = HTKKFile.ColumnLettersToIndex("C") - 1;
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					sheetView.Cells[dataLocationTemplateRowStart + i, column].Locked = true;
					sheetView.Cells[dataLocationTemplateRowStart + i, column2].Locked = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_1(SheetView sView, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			if (PeriodInfo.ThuHo == "TH")
			{
				string key = "10";
				string value = "10 %";
				dictionary.Add(key, value);
				string[] array = new string[dictionary.Count];
				string[] array2 = new string[dictionary.Count];
				int num = 0;
				foreach (KeyValuePair<string, string> item in dictionary)
				{
					string[] array4;
					string[] array3 = (array4 = array);
					int num2 = num;
					nint num3 = num2;
					array3[num2] = array4[num3] + item.Value;
					string[] array5 = (array4 = array2);
					int num4 = num;
					num3 = num4;
					array5[num4] = array4[num3] + item.Key;
					num++;
				}
				ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
				comboBoxCellType.Items = array;
				comboBoxCellType.ItemData = array2;
				sView.Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].CellType = comboBoxCellType;
			}
			int[] array6 = GeneralFunctions.ArrayCellId("C_1");
			if (sView.Cells[array6[1], array6[0]].CellType.ToString() != "ButtonCellType")
			{
				ButtonCellType buttonCellType = new ButtonCellType();
				buttonCellType.TwoState = false;
				buttonCellType.Text = "Mẫu tải bảng kê";
				buttonCellType.ShadowSize = 0;
				buttonCellType.TextColor = Color.Blue;
				sView.Cells[array6[1], array6[0]].CellType = buttonCellType;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 0 && idxCol == 2)
			{
				ShowBangKe("Bang_Ke_04_ND15_GTGT_TT80.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void ShowBangKe(string sFileName)
	{
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			string text = Path.GetTempPath() + sFileName;
			File.Copy(HTKKConfig.INTERFACE_TEMPLATES + "\\excel\\" + sFileName, text, overwrite: true);
			Process.Start(text);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			frmMessage frmMessage = new frmMessage("0433", "3", "", "Thông báo");
			((Form)frmMessage).ShowDialog();
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			bool result = true;
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] cell = GeneralFunctions.ArrayCellId("E_26");
			int[] cell2 = GeneralFunctions.ArrayCellId("E_27");
			int[] cell3 = GeneralFunctions.ArrayCellId("E_28");
			int[] cell4 = GeneralFunctions.ArrayCellId("E_29");
			int[] cell5 = GeneralFunctions.ArrayCellId("E_30");
			int[] cell6 = GeneralFunctions.ArrayCellId("E_31");
			int[] cell7 = GeneralFunctions.ArrayCellId("E_17");
			string message = HTKKFile.GetMessage(sheetView, cell);
			string message2 = HTKKFile.GetMessage(sheetView, cell2);
			string message3 = HTKKFile.GetMessage(sheetView, cell3);
			string message4 = HTKKFile.GetMessage(sheetView, cell7);
			string message5 = HTKKFile.GetMessage(sheetView, cell4);
			string message6 = HTKKFile.GetMessage(sheetView, cell5);
			string message7 = HTKKFile.GetMessage(sheetView, cell6);
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			if (GeneralFunctions.CheckTinhTheoCQT(headerInfo.maCQTCapCuc, headerInfo.maCQTNoiNop, sheetView, PeriodInfo, cell11c, cell11cMa, cell11b, cell11bMa, cell11a, cell11aMa, cell11DiaChi))
			{
				value = 2;
			}
			int[] array = GeneralFunctions.ArrayCellId("N_23");
			string text = sheetView.Cells[array[1], array[0]].Text;
			HTKKFile.RemoveValidate(sheetView, array[1], array[0]);
			if (!string.IsNullOrEmpty(text))
			{
				if (text.Length == 13 || text.Length == 14)
				{
					if (!ValidateBase.CheckMaSoThue(text))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], message2);
						value = 2;
					}
				}
				else if (text.Length == 5)
				{
					if (!HTKKFormater.IsNumber(text))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], message3);
						value = 2;
					}
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], message);
					value = 2;
				}
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num2 = HTKKFile.ColumnLettersToIndex("C") - 1;
			HTKKFile.RemoveValidate(sheetView, num + 1, num2);
			HTKKFile.RemoveValidate(sheetView, num, num2);
			if (hTKKCMCTableInfo.RowsValues.Count == 2)
			{
				string text2 = sheetView.Cells[num, num2].Text;
				string text3 = sheetView.Cells[num + 1, num2].Text;
				if (text2 == text3)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num, num2, message7);
					HTKKFile.SetNodeBackGroundPink(sheetView, num + 1, num2, message7);
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
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[0];
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKFile.RemoveValidateBackGround(sView, 46, HTKKFile.ColumnLettersToIndex("Z") - 1);
			if (ToKhaiInfo.SheetInfos[1] == null)
			{
				return;
			}
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			int num = HTKKFile.ColumnLettersToIndex("C") - 1;
			int num2 = HTKKFile.ColumnLettersToIndex("D") - 1;
			int num3 = HTKKFile.ColumnLettersToIndex("E") - 1;
			int column = HTKKFile.ColumnLettersToIndex("G") - 1;
			int column2 = HTKKFile.ColumnLettersToIndex("I") - 1;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20"));
			int num4 = 0;
			double num5 = 0.0;
			double num6 = 0.0;
			double num7 = 0.0;
			double num8 = 0.0;
			for (int i = 0; i < tables.Count; i++)
			{
				num4 = tables[i].Location.DataLocationTemplateRowStart - 1;
				HTKKCMCRowInfo rowsInfo = tables[i].RowsInfo;
				List<HTKKCMCRowValue> rowsValues = tables[i].RowsValues;
				num5 = 0.0;
				num6 = 0.0;
				num7 = 0.0;
				num8 = 0.0;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					HTKKFile.RemoveValidate(sheetView, num4 + j, num);
					HTKKFile.RemoveValidate(sheetView, num4 + j, num2);
					HTKKFile.RemoveValidate(sheetView, num4 + j, num3);
					string text = sheetView.Cells[num4 + j, num].Text;
					double num9 = double.Parse(sheetView.Cells[num4 + j, num2].Value.ToString());
					string text2 = sheetView.Cells[num4 + j, num3].Text;
					string text3 = sheetView.Cells[num4 + j, column2].Text;
					string obj = sheetView.Cells[num4 + j, column].Value.ToString();
					switch (text3)
					{
					case "1":
						num5 += ConvertHelpers.ConvertToDouble(obj);
						break;
					case "2":
						num6 += ConvertHelpers.ConvertToDouble(obj);
						break;
					case "3":
						num7 += ConvertHelpers.ConvertToDouble(obj);
						break;
					case "5":
						num8 += ConvertHelpers.ConvertToDouble(obj);
						break;
					}
					if (string.IsNullOrEmpty(text))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num4 + j, num, message);
						value = 2;
					}
					if (num9 == 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num4 + j, num2, message);
						value = 2;
					}
					if (string.IsNullOrEmpty(text2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num4 + j, num3, message);
						value = 2;
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[1].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[1].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[1].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public int GetDefaultYear()
	{
		return DateTime.Now.Year;
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			if (iRow >= 47 && iRow < 51)
			{
				TinhLaiChiTieu23_25_27_29Default(winFpSpread, ToKhaiInfo);
				TinhLaiChiTieu23_25_27_29TheoPL(winFpSpread, _ToKhaiInfo);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int num = HTKKFile.ColumnLettersToIndex("E");
			int num2 = HTKKFile.ColumnLettersToIndex("I");
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			string text = sheetView.Cells[(int)iRow - 1, num - 1].Text;
			if (PeriodInfo.ThuHo == "TH")
			{
				if (text == "10 %")
				{
					sheetView.Cells[(int)iRow - 1, num2 - 1].Value = "10";
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num2 - 1].Value = "0";
				}
			}
			TinhLaiChiTieu23_25_27_29TheoPL(winFpSpread, _ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			SheetView sheetView = winFpSpread.Sheets[1];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
			int column = HTKKFile.ColumnLettersToIndex("C") - 1;
			int num = HTKKFile.ColumnLettersToIndex("D") - 1;
			int column2 = HTKKFile.ColumnLettersToIndex("E") - 1;
			int column3 = HTKKFile.ColumnLettersToIndex("I") - 1;
			int num2 = 0;
			for (int i = 0; i < tables.Count; i++)
			{
				num2 = tables[i].Location.DataLocationTemplateRowStart - 1;
				HTKKCMCRowInfo rowsInfo = tables[i].RowsInfo;
				List<HTKKCMCRowValue> rowsValues = tables[i].RowsValues;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					string text = sheetView.Cells[num2 + j, column].Text;
					string text2 = sheetView.Cells[num2 + j, column3].Text;
					sheetView.Cells[num2 + j, column2].Text = getTenThueSuat(text2);
				}
			}
			TinhLaiChiTieu23_25_27_29TheoPL(winFpSpread, _ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private string getTenThueSuat(string maThueSuat)
	{
		//IL_0051: Unknown result type (might be due to invalid IL or missing references)
		//IL_0058: Expected O, but got Unknown
		string result = "";
		XmlDataDocument xmlDataDocument = new XmlDataDocument();
		FileStream fileStream = new FileStream(HTKKConfig.INTERFACE_INI + "\\DM_thue_suat_GTGT.xml", FileMode.Open, FileAccess.Read);
		((XmlDocument)xmlDataDocument).Load((Stream)fileStream);
		XmlNodeList elementsByTagName = ((XmlDocument)xmlDataDocument).GetElementsByTagName("Item");
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		foreach (XmlNode item in elementsByTagName)
		{
			XmlNode val = item;
			string value = ((XmlNode)val.Attributes["Value"]).Value;
			string[] array = Regex.Split(value, "###");
			if (array.Length >= 2 && maThueSuat == array[0])
			{
				result = array[1];
				break;
			}
		}
		return result;
	}

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		SheetView sheetView = winFpSpread.Sheets[1];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
		int num = HTKKFile.ColumnLettersToIndex("E");
		int num2 = HTKKFile.ColumnLettersToIndex("I");
		if (hTKKCMCSheetInfo == null)
		{
			return;
		}
		for (int i = 0; i < hTKKCMCSheetInfo.Tables[0].RowsValues.Count; i++)
		{
			string text = sheetView.Cells[hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart + i - 1, num2 - 1].Text.Trim();
			if (PeriodInfo.ThuHo == "TH")
			{
				if (text == "10")
				{
					sheetView.Cells[hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart + i - 1, num - 1].Value = "10%";
					continue;
				}
				sheetView.Cells[hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart + i - 1, num - 1].Value = "";
				sheetView.Cells[hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart + i - 1, num2 - 1].Text = "0";
			}
			else if (ConvertHelpers.ConvertToDouble(text) == 1.0 || ConvertHelpers.ConvertToDouble(text) == 2.0 || ConvertHelpers.ConvertToDouble(text) == 3.0 || ConvertHelpers.ConvertToDouble(text) == 5.0)
			{
				sheetView.Cells[hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart + i - 1, num - 1].Value = text + "%";
			}
			else
			{
				sheetView.Cells[hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart + i - 1, num - 1].Value = "";
				sheetView.Cells[hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart + i - 1, num2 - 1].Text = "0";
			}
		}
		TinhLaiChiTieu23_25_27_29TheoPL(winFpSpread, _ToKhaiInfo);
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellCoPl[1], cellCoPl[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			TinhLaiChiTieu23_25_27_29TheoPL(winFpSpread, _ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_1(FpSpread winFpSpread, int actionType)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellCoPl[1], cellCoPl[0]].Value = ((_ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			TinhLaiChiTieu23_25_27_29TheoPL(winFpSpread, _ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhLaiChiTieu23_25_27_29TheoPL(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			SheetView sheetView2 = winFpSpread.Sheets[1];
			int num = HTKKFile.ColumnLettersToIndex("C") - 1;
			int num2 = HTKKFile.ColumnLettersToIndex("D") - 1;
			int num3 = HTKKFile.ColumnLettersToIndex("E") - 1;
			int column = HTKKFile.ColumnLettersToIndex("G") - 1;
			int column2 = HTKKFile.ColumnLettersToIndex("I") - 1;
			int num4 = HTKKFile.ColumnLettersToIndex("Z");
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
			double num5 = 0.0;
			double num6 = 0.0;
			double num7 = 0.0;
			double num8 = 0.0;
			for (int i = 0; i < tables.Count; i++)
			{
				int num9 = tables[i].Location.DataLocationTemplateRowStart - 1;
				HTKKCMCRowInfo rowsInfo = tables[i].RowsInfo;
				List<HTKKCMCRowValue> rowsValues = tables[i].RowsValues;
				num5 = 0.0;
				num6 = 0.0;
				num7 = 0.0;
				num8 = 0.0;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					string text = sheetView2.Cells[num9 + j, column2].Text;
					string obj = sheetView2.Cells[num9 + j, column].Value.ToString();
					switch (text)
					{
					case "1":
						num5 += ConvertHelpers.ConvertToDouble(obj);
						break;
					case "2":
						num6 += ConvertHelpers.ConvertToDouble(obj);
						break;
					case "3":
						num7 += ConvertHelpers.ConvertToDouble(obj);
						break;
					case "5":
						num8 += ConvertHelpers.ConvertToDouble(obj);
						break;
					}
				}
			}
			double num10 = ConvertHelpers.ConvertToDouble(sheetView.Cells[46, HTKKFile.ColumnLettersToIndex("Z") - 1].Value);
			double num11 = ConvertHelpers.ConvertToDouble(sheetView.Cells[46, HTKKFile.ColumnLettersToIndex("AJ") - 1].Value);
			double num12 = ConvertHelpers.ConvertToDouble(sheetView.Cells[47, HTKKFile.ColumnLettersToIndex("Z") - 1].Value);
			double num13 = ConvertHelpers.ConvertToDouble(sheetView.Cells[47, HTKKFile.ColumnLettersToIndex("AJ") - 1].Value);
			double num14 = ConvertHelpers.ConvertToDouble(sheetView.Cells[48, HTKKFile.ColumnLettersToIndex("Z") - 1].Value);
			double num15 = ConvertHelpers.ConvertToDouble(sheetView.Cells[48, HTKKFile.ColumnLettersToIndex("AJ") - 1].Value);
			double num16 = ConvertHelpers.ConvertToDouble(sheetView.Cells[49, HTKKFile.ColumnLettersToIndex("Z") - 1].Value);
			double num17 = ConvertHelpers.ConvertToDouble(sheetView.Cells[49, HTKKFile.ColumnLettersToIndex("AJ") - 1].Value);
			double num18 = num11 - num5;
			double num19 = num13 - num8;
			double num20 = num15 - num7;
			double num21 = num17 - num6;
			sheetView.Cells[46, num4 - 1].Value = num18;
			sheetView.Cells[47, num4 - 1].Value = num19;
			sheetView.Cells[48, num4 - 1].Value = num20;
			sheetView.Cells[49, num4 - 1].Value = num21;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhLaiChiTieu23_25_27_29Default(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("Z");
			int num2 = HTKKFile.ColumnLettersToIndex("AJ");
			SheetView sheetView = winFpSpread.Sheets[0];
			sheetView.Cells[46, num - 1].Value = sheetView.Cells[46, num2 - 1].Value;
			sheetView.Cells[47, num - 1].Value = sheetView.Cells[47, num2 - 1].Value;
			sheetView.Cells[48, num - 1].Value = sheetView.Cells[48, num2 - 1].Value;
			sheetView.Cells[49, num - 1].Value = sheetView.Cells[49, num2 - 1].Value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
