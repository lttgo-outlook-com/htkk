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

public class Gtgt04TT89Nckt309
{
	private int[] cellCt22TK = GeneralFunctions.ArrayCellId("Q_87");

	private int[] cellCt24TK = GeneralFunctions.ArrayCellId("Q_88");

	private int[] cellCt26TK = GeneralFunctions.ArrayCellId("Q_89");

	private int[] cellCt28TK = GeneralFunctions.ArrayCellId("Q_90");

	private int[] cellCt23TK = GeneralFunctions.ArrayCellId("Z_87");

	private int[] cellCt25TK = GeneralFunctions.ArrayCellId("Z_88");

	private int[] cellCt27TK = GeneralFunctions.ArrayCellId("Z_89");

	private int[] cellCt29TK = GeneralFunctions.ArrayCellId("Z_90");

	private int colCongThucTK = HTKKFile.ColumnLettersToIndex("AW");

	private int[] cell11a_TK = GeneralFunctions.ArrayCellId("G_69");

	private int[] cell11aMa_TK = GeneralFunctions.ArrayCellId("AT_69");

	private int[] cellct09_TK = GeneralFunctions.ArrayCellId("C_61");

	private int[] cellct10_TK = GeneralFunctions.ArrayCellId("N_63");

	private int[] cellCoPl01_1 = GeneralFunctions.ArrayCellId("AT_41");

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	private HTKKPeriodInfo _PeriodInfo = null;

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			_ToKhaiInfo = ToKhaiInfo;
			_PeriodInfo = PeriodInfo;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			if (headerInfo.maTinhNNT != PeriodInfo.MaTinh)
			{
				GeneralFunctions.lockUnlockMultiCellsStatic(sheetView, sInfo, cellct09_TK[1] + 1, cell11a_TK[1] + 1, isLock: false, isDeleteData: false);
			}
			else
			{
				GeneralFunctions.lockUnlockMultiCellsStatic(sheetView, sInfo, cellct09_TK[1] + 1, cell11a_TK[1] + 1, isLock: true, isDeleteData: true);
			}
			sheetView.Cells[cellCoPl01_1[1], cellCoPl01_1[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
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
				string value = "10";
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
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[1];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					sView.Cells[dataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex("E") - 1].CellType = comboBoxCellType;
					sView.Cells[dataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex("E") - 1].Text = "10";
					sView.Cells[dataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex("I") - 1].Text = "10";
				}
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
			array6 = GeneralFunctions.ArrayCellId("B_2");
			string text = "";
			if (PeriodInfo.Nam.Equals("2024"))
			{
				text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 142/2024/QH15";
			}
			else if (PeriodInfo.Nam.Equals("2025") && ((ConvertHelpers.ConvertToInt(PeriodInfo.Thang) <= 6 && !string.IsNullOrEmpty(PeriodInfo.Thang)) || (ConvertHelpers.ConvertToInt(PeriodInfo.Quy) <= 2 && !string.IsNullOrEmpty(PeriodInfo.Quy))))
			{
				text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 174/2024/QH15";
			}
			else if ((PeriodInfo.Nam.Equals("2025") && (ConvertHelpers.ConvertToInt(PeriodInfo.Thang) > 6 || ConvertHelpers.ConvertToInt(PeriodInfo.Quy) > 2)) || PeriodInfo.Nam.Equals("2026"))
			{
				text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 204/2025/QH15";
			}
			sView.Cells[array6[1], array6[0]].Text = text;
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
				ShowBangKe("Bang_Ke_04_GiamThue_GTGT_NQ142_GTGT_TT89.xls");
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
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("C");
			if (iRow >= cellCt22TK[1] + 1 && iRow <= cellCt28TK[1] + 1)
			{
				TinhLaiChiTieu23_25_27_29TheoPL01_1(winFpSpread, _ToKhaiInfo);
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
			int num = HTKKFile.ColumnLettersToIndex("D");
			int num2 = HTKKFile.ColumnLettersToIndex("E");
			int num3 = HTKKFile.ColumnLettersToIndex("I");
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			string text = sheetView.Cells[(int)iRow - 1, num2 - 1].Text;
			if (PeriodInfo.ThuHo == "TH")
			{
				if (text == "10")
				{
					sheetView.Cells[(int)iRow - 1, num3 - 1].Value = "10";
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num3 - 1].Value = "0";
				}
			}
			else
			{
				if (iCol == num)
				{
					string text2 = sheetView.Cells[(int)iRow - 1, num3 - 1].Text.Trim();
					if (!string.IsNullOrEmpty(text2))
					{
						TinhLaiChiTieu22_den_28_TK(winFpSpread, ToKhaiInfo, PeriodInfo, text2);
					}
				}
				if (iCol == num2)
				{
					string text2 = sheetView.Cells[(int)iRow - 1, num3 - 1].Text.Trim();
					if (!string.IsNullOrEmpty(text2))
					{
						TinhLaiChiTieu22_den_28_TK(winFpSpread, ToKhaiInfo, PeriodInfo, "");
					}
				}
			}
			TinhLaiChiTieu23_25_27_29TheoPL01_1(winFpSpread, _ToKhaiInfo);
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
			if (hTKKCMCSheetInfo != null)
			{
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
			TinhLaiThueSuat_PL01_1(winFpSpread, ToKhaiInfo);
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
		if (hTKKCMCSheetInfo != null)
		{
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
		TinhLaiThueSuat_PL01_1(winFpSpread, ToKhaiInfo);
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[cellCoPl01_1[1], cellCoPl01_1[0]].Value = ((ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			TinhLaiChiTieu23_25_27_29TheoPL01_1(winFpSpread, _ToKhaiInfo);
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
			winFpSpread.Sheets[0].Cells[cellCoPl01_1[1], cellCoPl01_1[0]].Value = ((_ToKhaiInfo.SheetInfos[1] != null) ? 1 : 0);
			TinhLaiChiTieu23_25_27_29TheoPL01_1(winFpSpread, _ToKhaiInfo);
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
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
			double num4 = 0.0;
			double num5 = 0.0;
			double num6 = 0.0;
			double num7 = 0.0;
			for (int i = 0; i < tables.Count; i++)
			{
				int num8 = tables[i].Location.DataLocationTemplateRowStart - 1;
				HTKKCMCRowInfo rowsInfo = tables[i].RowsInfo;
				List<HTKKCMCRowValue> rowsValues = tables[i].RowsValues;
				num4 = 0.0;
				num5 = 0.0;
				num6 = 0.0;
				num7 = 0.0;
				for (int j = 0; j < rowsValues.Count; j++)
				{
					string text = sheetView2.Cells[num8 + j, column2].Text;
					string obj = sheetView2.Cells[num8 + j, column].Value.ToString();
					switch (text)
					{
					case "1":
						num4 += ConvertHelpers.ConvertToDouble(obj);
						break;
					case "2":
						num5 += ConvertHelpers.ConvertToDouble(obj);
						break;
					case "3":
						num6 += ConvertHelpers.ConvertToDouble(obj);
						break;
					case "5":
						num7 += ConvertHelpers.ConvertToDouble(obj);
						break;
					}
				}
			}
			double num9 = ConvertHelpers.ConvertToDouble(sheetView.Cells[cellCt23TK[1], cellCt23TK[0]].Value);
			double num10 = ConvertHelpers.ConvertToDouble(sheetView.Cells[cellCt23TK[1], colCongThucTK - 1].Value);
			double num11 = ConvertHelpers.ConvertToDouble(sheetView.Cells[cellCt25TK[1], cellCt25TK[0]].Value);
			double num12 = ConvertHelpers.ConvertToDouble(sheetView.Cells[cellCt25TK[1], colCongThucTK - 1].Value);
			double num13 = ConvertHelpers.ConvertToDouble(sheetView.Cells[cellCt27TK[1], cellCt27TK[0]].Value);
			double num14 = ConvertHelpers.ConvertToDouble(sheetView.Cells[cellCt27TK[1], colCongThucTK - 1].Value);
			double num15 = ConvertHelpers.ConvertToDouble(sheetView.Cells[cellCt29TK[1], cellCt29TK[0]].Value);
			double num16 = ConvertHelpers.ConvertToDouble(sheetView.Cells[cellCt29TK[1], colCongThucTK - 1].Value);
			double num17 = num10 - num4;
			double num18 = num12 - num7;
			double num19 = num14 - num6;
			double num20 = num16 - num5;
			sheetView.Cells[cellCt23TK[1], cellCt23TK[0]].Value = num17;
			sheetView.Cells[cellCt25TK[1], cellCt25TK[0]].Value = num18;
			sheetView.Cells[cellCt27TK[1], cellCt27TK[0]].Value = num19;
			sheetView.Cells[cellCt29TK[1], cellCt29TK[0]].Value = num20;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhLaiChiTieu23_25_27_29TheoPL01_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo != null)
			{
				SheetView sheetView2 = winFpSpread.Sheets[1];
				sheetView.Cells[cellCt23TK[1], cellCt23TK[0]].Value = ConvertHelpers.ConvertToDouble(sheetView.Cells[cellCt23TK[1], colCongThucTK - 1].Value);
				sheetView.Cells[cellCt25TK[1], cellCt25TK[0]].Value = ConvertHelpers.ConvertToDouble(sheetView.Cells[cellCt25TK[1], colCongThucTK - 1].Value);
				sheetView.Cells[cellCt27TK[1], cellCt27TK[0]].Value = ConvertHelpers.ConvertToDouble(sheetView.Cells[cellCt27TK[1], colCongThucTK - 1].Value);
				sheetView.Cells[cellCt29TK[1], cellCt29TK[0]].Value = ConvertHelpers.ConvertToDouble(sheetView.Cells[cellCt29TK[1], colCongThucTK - 1].Value);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhLaiThueSuat_PL01_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		SheetView sheetView = winFpSpread.Sheets[1];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
		if (hTKKCMCSheetInfo == null)
		{
			return;
		}
		int num = HTKKFile.ColumnLettersToIndex("E");
		int num2 = HTKKFile.ColumnLettersToIndex("I");
		HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[1];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
		{
			string value = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text.Trim();
			if (string.IsNullOrEmpty(value))
			{
				string text = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text))
				{
					sheetView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text = text;
				}
			}
		}
		TinhLaiChiTieu22_den_28_TK(winFpSpread, ToKhaiInfo, _PeriodInfo, "");
		TinhLaiChiTieu23_25_27_29TheoPL01_1(winFpSpread, _ToKhaiInfo);
	}

	public void TinhLaiChiTieu22_den_28_TK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo, string valueMaThueSuat)
	{
		int[] array = GeneralFunctions.ArrayCellId("I_9");
		int[] array2 = GeneralFunctions.ArrayCellId("J_9");
		int[] array3 = GeneralFunctions.ArrayCellId("K_9");
		int[] array4 = GeneralFunctions.ArrayCellId("L_9");
		SheetView sheetView = winFpSpread.Sheets[0];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
		if (hTKKCMCSheetInfo != null && !PeriodInfo.ThuHo.Equals("TH"))
		{
			SheetView sheetView2 = winFpSpread.Sheets[1];
			if (valueMaThueSuat == "1" || valueMaThueSuat == "")
			{
				sheetView.Cells[cellCt22TK[1], cellCt22TK[0]].Value = sheetView2.Cells[array[1], array[0]].Value;
			}
			if (valueMaThueSuat == "5" || valueMaThueSuat == "")
			{
				sheetView.Cells[cellCt24TK[1], cellCt24TK[0]].Value = sheetView2.Cells[array4[1], array4[0]].Value;
			}
			if (valueMaThueSuat == "3" || valueMaThueSuat == "")
			{
				sheetView.Cells[cellCt26TK[1], cellCt26TK[0]].Value = sheetView2.Cells[array3[1], array3[0]].Value;
			}
			if (valueMaThueSuat == "2" || valueMaThueSuat == "")
			{
				sheetView.Cells[cellCt28TK[1], cellCt28TK[0]].Value = sheetView2.Cells[array2[1], array2[0]].Value;
			}
		}
	}
}
