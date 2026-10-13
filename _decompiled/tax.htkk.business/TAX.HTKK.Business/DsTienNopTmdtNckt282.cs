using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DsTienNopTmdtNckt282
{
	private const string So_TK = "1";

	private const string So_QD = "2";

	private const string So_TB = "3";

	private const string Loai_MST = "1000";

	private const string Loai_Ho_Chieu = "1020";

	private const string Loai_CCCD = "2080";

	private const string Loai_Giay_DKKD = "3000";

	private const string KY_THANG = "M";

	private const string KY_QUY = "Q";

	private const string KY_NAM = "Y";

	private const string KY_LAN_PHAT_SINH = "D";

	private const int Row_ngay_QD = 38;

	private const int Row_ngay_TB = 40;

	private const int Row_loai_ky = 33;

	private const int Row_Ct05 = 28;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("E_9");
			sView.Cells[array[1], array[0]].Text = PeriodInfo.LanXuatBan;
			int[] array2 = GeneralFunctions.ArrayCellId("E_13");
			sView.Cells[array2[1], array2[0]].Text = PeriodInfo.MaHSDuAnDT;
			int num = HTKKFile.ColumnLettersToIndex("D");
			int num2 = HTKKFile.ColumnLettersToIndex("E");
			int num3 = HTKKFile.ColumnLettersToIndex("L");
			int num4 = HTKKFile.ColumnLettersToIndex("G");
			int num5 = HTKKFile.ColumnLettersToIndex("I");
			string text = sView.Cells[27, num3 - 1].Text;
			if (text == "1")
			{
				sView.Cells[32, num2 - 1].Locked = false;
				sView.Cells[32, num5 - 1].Locked = false;
			}
			else
			{
				sView.Cells[32, num2 - 1].Locked = true;
				sView.Cells[32, num5 - 1].Locked = true;
			}
			if (text == "2")
			{
				sView.Cells[37, num2 - 1].Locked = false;
			}
			else
			{
				sView.Cells[37, num2 - 1].Locked = true;
			}
			if (text == "3")
			{
				sView.Cells[39, num2 - 1].Locked = false;
			}
			else
			{
				sView.Cells[39, num2 - 1].Locked = true;
			}
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
			int num = HTKKFile.ColumnLettersToIndex("D");
			int num2 = HTKKFile.ColumnLettersToIndex("L");
			int num3 = HTKKFile.ColumnLettersToIndex("G");
			int num4 = HTKKFile.ColumnLettersToIndex("E");
			int num5 = HTKKFile.ColumnLettersToIndex("I");
			if (iCol == num && iRow == 28)
			{
				string text = sheetView.Cells[27, num2 - 1].Text;
				if (text == "1")
				{
					sheetView.Cells[32, num2 - 1].Text = "M";
					sheetView.Cells[32, num4 - 1].Text = "Tháng";
					sheetView.Cells[32, num5 - 1].Text = "";
					sheetView.Cells[32, num4 - 1].Locked = true;
					sheetView.Cells[32, num5 - 1].Locked = false;
					MaskCellType maskCellType = new MaskCellType();
					maskCellType.MaskChar = '.';
					maskCellType.Mask = "00/0000";
					sheetView.Cells[32, num5 - 1].Text = "";
					sheetView.Cells[32, num5 - 1].ResetCellType();
					sheetView.Cells[32, num5 - 1].CellType = maskCellType;
				}
				else
				{
					sheetView.Cells[32, num2 - 1].Text = "";
					sheetView.Cells[32, num4 - 1].Text = "";
					sheetView.Cells[32, num5 - 1].Text = "";
					sheetView.Cells[32, num4 - 1].Locked = true;
					sheetView.Cells[32, num5 - 1].Locked = true;
				}
				if (text == "2")
				{
					sheetView.Cells[37, num4 - 1].Text = "";
					sheetView.Cells[37, num4 - 1].Locked = false;
				}
				else
				{
					sheetView.Cells[37, num4 - 1].Text = "";
					sheetView.Cells[37, num4 - 1].Locked = true;
				}
				if (text == "3")
				{
					sheetView.Cells[39, num4 - 1].Text = "";
					sheetView.Cells[39, num4 - 1].Locked = false;
				}
				else
				{
					sheetView.Cells[39, num4 - 1].Text = "";
					sheetView.Cells[39, num4 - 1].Locked = true;
				}
			}
			if (iCol == num4 && iRow == 33)
			{
				string text2 = sheetView.Cells[32, num2 - 1].Text;
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.MaskChar = '.';
				switch (text2)
				{
				case "Y":
					maskCellType.Mask = "0000";
					break;
				case "Q":
					maskCellType.Mask = "0/0000";
					break;
				case "M":
					maskCellType.Mask = "00/0000";
					break;
				case "D":
					maskCellType.Mask = "00/00/0000";
					break;
				}
				sheetView.Cells[32, num5 - 1].Text = "";
				sheetView.Cells[32, num5 - 1].ResetCellType();
				sheetView.Cells[32, num5 - 1].CellType = maskCellType;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			if ((iCol == num || iCol == num4 || iCol == num5) && iRow >= 70 && iRow <= hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count)
			{
				string text3 = sheetView.Cells[(int)iRow - 1, num - 1].Text;
				string text4 = sheetView.Cells[(int)iRow - 1, num4 - 1].Text;
				string text5 = sheetView.Cells[(int)iRow - 1, num5 - 1].Text;
				if (!string.IsNullOrEmpty(text3) || !string.IsNullOrEmpty(text4) || ConvertHelpers.ConvertToDouble(text5) != 0.0)
				{
					sheetView.Cells[(int)iRow - 1, num3 - 1].Text = "1701";
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num3 - 1].Text = "";
				}
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[0].Tables[1];
			if ((iCol == num || iCol == num4 || iCol == num5) && iRow >= hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart && iRow <= hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart + hTKKCMCTableInfo2.RowsValues.Count)
			{
				string text3 = sheetView.Cells[(int)iRow - 1, num - 1].Text;
				string text4 = sheetView.Cells[(int)iRow - 1, num4 - 1].Text;
				string text5 = sheetView.Cells[(int)iRow - 1, num5 - 1].Text;
				if (!string.IsNullOrEmpty(text3) || !string.IsNullOrEmpty(text4) || ConvertHelpers.ConvertToDouble(text5) != 0.0)
				{
					sheetView.Cells[(int)iRow - 1, num3 - 1].Text = "1003";
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num3 - 1].Text = "";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("C") - 1;
			int column = HTKKFile.ColumnLettersToIndex("L") - 1;
			int odrColMST = HTKKFile.ColumnLettersToIndex("D") - 1;
			int num2 = HTKKFile.ColumnLettersToIndex("E") - 1;
			int num3 = HTKKFile.ColumnLettersToIndex("G") - 1;
			int num4 = HTKKFile.ColumnLettersToIndex("I") - 1;
			int[] array = GeneralFunctions.ArrayCellId("E_17");
			string text = sheetView.Cells[array[1], array[0]].Text.Trim();
			bool success = false;
			bool flag = false;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[14, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[19, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[28, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[29, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[30, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[33, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text11 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[35, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			DateTime dateTime = HTKKFormater.ConvertToDateTime(text, out success);
			HTKKFile.RemoveValidate(sheetView, array[1], array[0]);
			if (!string.IsNullOrEmpty(text))
			{
				if (!success)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], HTKKFile.GetMessage("0071"));
					value = 2;
				}
				else if ((dateTime - DateTime.Now.Date).TotalDays > 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], text3);
					value = 2;
				}
			}
			int[] array2 = GeneralFunctions.ArrayCellId("L_28");
			int[] array3 = GeneralFunctions.ArrayCellId("E_33");
			int[] array4 = GeneralFunctions.ArrayCellId("L_33");
			int[] array5 = GeneralFunctions.ArrayCellId("I_33");
			int[] array6 = GeneralFunctions.ArrayCellId("E_38");
			int[] array7 = GeneralFunctions.ArrayCellId("E_40");
			string text12 = sheetView.Cells[array2[1], array2[0]].Text.Trim();
			string value2 = sheetView.Cells[array3[1], array3[0]].Text.Trim();
			string text13 = sheetView.Cells[array4[1], array4[0]].Text.Trim();
			string text14 = sheetView.Cells[array5[1], array5[0]].Text.Trim();
			string text15 = sheetView.Cells[array6[1], array6[0]].Text.Trim();
			string text16 = sheetView.Cells[array7[1], array7[0]].Text.Trim();
			HTKKFile.RemoveValidate(sheetView, array3[1], array3[0]);
			HTKKFile.RemoveValidate(sheetView, array5[1], array5[0]);
			HTKKFile.RemoveValidate(sheetView, array6[1], array6[0]);
			HTKKFile.RemoveValidate(sheetView, array7[1], array7[0]);
			switch (text12)
			{
			case "1":
				if (string.IsNullOrEmpty(value2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], text2);
					value = 2;
				}
				if (string.IsNullOrEmpty(text14))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array5[1], array5[0], text2);
					value = 2;
					break;
				}
				switch (text13)
				{
				case "M":
					if (ConvertHelpers.ConvertToDateTime("01/" + text14, "dd/MM/yyyy") == DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array5[1], array5[0], text7);
						value = 2;
					}
					break;
				case "D":
					if (ConvertHelpers.ConvertToDateTime(text14, "dd/MM/yyyy") == DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array5[1], array5[0], text7);
						value = 2;
					}
					break;
				case "Q":
					if (ConvertHelpers.ConvertToDateTime("01/" + (ConvertHelpers.ConvertToInt(text14.Split(new char[1] { '/' })[0].Replace("Q", "")) * 3 - 2).ToString().PadLeft(2, '0') + "/" + text14.Split(new char[1] { '/' })[1], "dd/MM/yyyy") == DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array5[1], array5[0], text7);
						value = 2;
					}
					break;
				}
				break;
			case "2":
				if (string.IsNullOrEmpty(text15))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array6[1], array6[0], text2);
					value = 2;
				}
				else if (ConvertHelpers.ConvertToDateTime(text15, "dd/MM/yyyy") == DateTime.MinValue)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array6[1], array6[0], text8);
					value = 2;
				}
				break;
			case "3":
				if (string.IsNullOrEmpty(text16))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array7[1], array7[0], text2);
					value = 2;
				}
				else if (ConvertHelpers.ConvertToDateTime(text16, "dd/MM/yyyy") == DateTime.MinValue)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array7[1], array7[0], text9);
					value = 2;
				}
				break;
			}
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int num5 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				string[] array8 = new string[7];
				int num6 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == odrColMST + 1);
				for (int num7 = 0; num7 < hTKKCMCTableInfo.RowsValues.Count; num7++)
				{
					array8[0] = ConvertHelpers.ConvertToString(sheetView.Cells[num5 + num7, num].Value);
					array8[1] = ConvertHelpers.ConvertToString(sheetView.Cells[num5 + num7, odrColMST].Value);
					array8[2] = ConvertHelpers.ConvertToString(sheetView.Cells[num5 + num7, num2].Value);
					array8[3] = ConvertHelpers.ConvertToString(sheetView.Cells[num5 + num7, num4].Value);
					array8[4] = ConvertHelpers.ConvertToString(sheetView.Cells[num5 + num7, column].Value);
					array8[5] = ConvertHelpers.ConvertToString(sheetView.Cells[num5 + num7, num3].Value);
					HTKKFile.RemoveValidate(sheetView, num5 + num7, num);
					HTKKFile.RemoveValidate(sheetView, num5 + num7, odrColMST);
					HTKKFile.RemoveValidate(sheetView, num5 + num7, num2);
					HTKKFile.RemoveValidate(sheetView, num5 + num7, num3);
					HTKKFile.RemoveValidate(sheetView, num5 + num7, num4);
					if (string.IsNullOrEmpty(array8[0]) && string.IsNullOrEmpty(array8[1]) && string.IsNullOrEmpty(array8[2]) && !(array8[3] != "0") && string.IsNullOrEmpty(array8[4]) && string.IsNullOrEmpty(array8[5]))
					{
						continue;
					}
					flag = true;
					if (string.IsNullOrEmpty(array8[0]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num5 + num7, num, text2, num7 + 1);
						value = 2;
					}
					if (string.IsNullOrEmpty(array8[1]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num5 + num7, odrColMST, text2, num7 + 1);
						value = 2;
					}
					else if (array8[4] == "1000")
					{
						if (!ValidateBase.CheckMaSoThueCCCD(array8[1]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num5 + num7, odrColMST, text5, num7 + 1);
							value = 2;
						}
					}
					else if (array8[4] == "2080")
					{
						if (array8[1].Length != 12)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num5 + num7, odrColMST, text11, num7 + 1);
							value = 2;
						}
						else if (!ValidateBase.CheckMaSoThueCCCD(array8[1]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num5 + num7, odrColMST, text11, num7 + 1);
							value = 2;
						}
					}
					if (string.IsNullOrEmpty(array8[2]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num5 + num7, num2, text2, num7 + 1);
						value = 2;
					}
					if (string.IsNullOrEmpty(array8[5]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num5 + num7, num3, text2, num7 + 1);
						value = 2;
					}
					if (ConvertHelpers.ConvertToDouble(array8[3]) == 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num5 + num7, num4, text2, num7 + 1);
						value = 2;
					}
				}
			}
			if (!flag)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int num5 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				HTKKFile.SetNodeBackGroundPink(sheetView, num5, odrColMST, text2, 1);
				value = 2;
				HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[1];
				int rowIndex = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart - 1;
				HTKKFile.SetNodeBackGroundPink(sheetView, rowIndex, odrColMST, text2, 1);
				value = 2;
			}
			switch (value)
			{
			case 2:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.Pink;
				break;
			case 1:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.Yellow;
				break;
			default:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			HTKKFile.deleteLogErrorBigData();
			bool result = true;
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				switch (item.Value)
				{
				case 2:
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
					break;
				case 1:
					if (winFpSpread.TabStrip[key].BackColor == Color.White)
					{
						winFpSpread.TabStrip[key].BackColor = Color.Yellow;
					}
					break;
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			if (winFpSpread.ActiveSheet.ActiveRowIndex == 0 && activeColumnIndex == 2)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_Ds_tien_nop_TMDT.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("L_28");
			int[] array2 = GeneralFunctions.ArrayCellId("D_28");
			string text = winFpSpread.Sheets[0].Cells[array[1], array[0]].Text;
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			dictionary.Add("1", "Số tờ khai");
			dictionary.Add("2", "Số quyết định");
			dictionary.Add("3", "Số thông báo");
			dictionary.Add("4", "Mã định danh hồ sơ (ID)");
			Dictionary<string, string> dictionary2 = dictionary;
			string text2 = dictionary2[text];
			winFpSpread.Sheets[0].Cells[array2[1], array2[0]].Text = text2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int column = HTKKFile.ColumnLettersToIndex("B") - 1;
			int column2 = HTKKFile.ColumnLettersToIndex("G") - 1;
			string text = "";
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				switch (i)
				{
				case 0:
					text = "1701";
					break;
				case 1:
					text = "1003";
					break;
				}
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				object[,] array = new object[hTKKCMCTableInfo.RowsValues.Count, 1];
				object[,] array2 = new object[hTKKCMCTableInfo.RowsValues.Count, 1];
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					array[j, 0] = Convert.ToString(j + 1);
					array2[j, 0] = text;
				}
				if (((Control)sheetView.FpSpread).InvokeRequired)
				{
					sheetView.SetArray(hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1, column, array);
					sheetView.SetArray(hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1, column2, array2);
				}
				else
				{
					sheetView.SetArray(hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1, column, array);
					sheetView.SetArray(hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1, column2, array2);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
