using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DsTienNopTncnNckt278
{
	private const string So_TK = "1";

	private const string So_QD = "2";

	private const string So_TB = "3";

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
			int[] array = GeneralFunctions.ArrayCellId("H_9");
			sView.Cells[array[1], array[0]].Text = PeriodInfo.LanXuatBan;
			int[] array2 = GeneralFunctions.ArrayCellId("D_13");
			sView.Cells[array2[1], array2[0]].Text = PeriodInfo.MaHSDuAnDT;
			int num = HTKKFile.ColumnLettersToIndex("D");
			int num2 = HTKKFile.ColumnLettersToIndex("K");
			int num3 = HTKKFile.ColumnLettersToIndex("H");
			int num4 = HTKKFile.ColumnLettersToIndex("C");
			int num5 = HTKKFile.ColumnLettersToIndex("F");
			string text = sView.Cells[27, num2 - 1].Text;
			if (text == "1")
			{
				sView.Cells[32, num - 1].Locked = false;
				sView.Cells[32, num3 - 1].Locked = false;
			}
			else
			{
				sView.Cells[32, num - 1].Locked = true;
				sView.Cells[32, num3 - 1].Locked = true;
			}
			if (text == "2")
			{
				sView.Cells[37, num - 1].Locked = false;
			}
			else
			{
				sView.Cells[37, num - 1].Locked = true;
			}
			if (text == "3")
			{
				sView.Cells[39, num - 1].Locked = false;
			}
			else
			{
				sView.Cells[39, num - 1].Locked = true;
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
			int num2 = HTKKFile.ColumnLettersToIndex("K");
			int num3 = HTKKFile.ColumnLettersToIndex("H");
			int num4 = HTKKFile.ColumnLettersToIndex("C");
			int num5 = HTKKFile.ColumnLettersToIndex("F");
			if (iCol == num4 && iRow == 28)
			{
				string text = sheetView.Cells[27, num2 - 1].Text;
				if (text == "1")
				{
					sheetView.Cells[32, num2 - 1].Text = "";
					sheetView.Cells[32, num - 1].Text = "";
					sheetView.Cells[32, num3 - 1].Text = "";
					sheetView.Cells[32, num - 1].Locked = false;
					sheetView.Cells[32, num3 - 1].Locked = false;
				}
				else
				{
					sheetView.Cells[32, num2 - 1].Text = "";
					sheetView.Cells[32, num - 1].Text = "";
					sheetView.Cells[32, num3 - 1].Text = "";
					sheetView.Cells[32, num - 1].Locked = true;
					sheetView.Cells[32, num3 - 1].Locked = true;
				}
				if (text == "2")
				{
					sheetView.Cells[37, num - 1].Text = "";
					sheetView.Cells[37, num - 1].Locked = false;
				}
				else
				{
					sheetView.Cells[37, num - 1].Text = "";
					sheetView.Cells[37, num - 1].Locked = true;
				}
				if (text == "3")
				{
					sheetView.Cells[39, num - 1].Text = "";
					sheetView.Cells[39, num - 1].Locked = false;
				}
				else
				{
					sheetView.Cells[39, num - 1].Text = "";
					sheetView.Cells[39, num - 1].Locked = true;
				}
			}
			if (iCol == num && iRow == 33)
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
				sheetView.Cells[32, num3 - 1].Text = "";
				sheetView.Cells[32, num3 - 1].ResetCellType();
				sheetView.Cells[32, num3 - 1].CellType = maskCellType;
			}
			if (iCol == num5 && iRow >= 69)
			{
				sheetView.Cells[(int)iRow - 1, num3 - 1].Formula = "F" + iRow;
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
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string[] array = new string[7];
			int odrColMST = HTKKFile.ColumnLettersToIndex("C") - 1;
			int num2 = HTKKFile.ColumnLettersToIndex("D") - 1;
			int num3 = HTKKFile.ColumnLettersToIndex("F") - 1;
			int num4 = HTKKFile.ColumnLettersToIndex("H") - 1;
			int num5 = HTKKFile.ColumnLettersToIndex("I") - 1;
			int[] array2 = GeneralFunctions.ArrayCellId("D_17");
			string text = sheetView.Cells[array2[1], array2[0]].Text.Trim();
			bool success = false;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[14, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[19, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[28, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[29, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[30, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[33, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			int indexCol = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == odrColMST + 1);
			DateTime dateTime = HTKKFormater.ConvertToDateTime(text, out success);
			HTKKFile.RemoveValidate(sheetView, array2[1], array2[0]);
			if (string.IsNullOrEmpty(text))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], text2);
				value = 2;
			}
			else if (!success)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], HTKKFile.GetMessage("0071"));
				value = 2;
			}
			else if ((dateTime - DateTime.Now.Date).TotalDays > 0.0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], text3);
				value = 2;
			}
			int[] array3 = GeneralFunctions.ArrayCellId("K_28");
			int[] array4 = GeneralFunctions.ArrayCellId("D_33");
			int[] array5 = GeneralFunctions.ArrayCellId("K_33");
			int[] array6 = GeneralFunctions.ArrayCellId("H_33");
			int[] array7 = GeneralFunctions.ArrayCellId("D_38");
			int[] array8 = GeneralFunctions.ArrayCellId("D_40");
			string text11 = sheetView.Cells[array3[1], array3[0]].Text.Trim();
			string value2 = sheetView.Cells[array4[1], array4[0]].Text.Trim();
			string text12 = sheetView.Cells[array5[1], array5[0]].Text.Trim();
			string text13 = sheetView.Cells[array6[1], array6[0]].Text.Trim();
			string text14 = sheetView.Cells[array7[1], array7[0]].Text.Trim();
			string text15 = sheetView.Cells[array8[1], array8[0]].Text.Trim();
			HTKKFile.RemoveValidate(sheetView, array4[1], array4[0]);
			HTKKFile.RemoveValidate(sheetView, array6[1], array6[0]);
			HTKKFile.RemoveValidate(sheetView, array7[1], array7[0]);
			HTKKFile.RemoveValidate(sheetView, array8[1], array8[0]);
			switch (text11)
			{
			case "1":
				if (string.IsNullOrEmpty(value2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array4[1], array4[0], text2);
					value = 2;
				}
				if (string.IsNullOrEmpty(text13))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array6[1], array6[0], text2);
					value = 2;
					break;
				}
				switch (text12)
				{
				case "M":
					if (ConvertHelpers.ConvertToDateTime("01/" + text13, "dd/MM/yyyy") == DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array6[1], array6[0], text7);
						value = 2;
					}
					break;
				case "D":
					if (ConvertHelpers.ConvertToDateTime(text13, "dd/MM/yyyy") == DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array6[1], array6[0], text7);
						value = 2;
					}
					break;
				case "Q":
					if (ConvertHelpers.ConvertToDateTime("01/" + (ConvertHelpers.ConvertToInt(text13.Split(new char[1] { '/' })[0].Replace("Q", "")) * 3 - 2).ToString().PadLeft(2, '0') + "/" + text13.Split(new char[1] { '/' })[1], "dd/MM/yyyy") == DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array6[1], array6[0], text7);
						value = 2;
					}
					break;
				}
				break;
			case "2":
				if (string.IsNullOrEmpty(text14))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array7[1], array7[0], text2);
					value = 2;
				}
				else if (ConvertHelpers.ConvertToDateTime(text14, "dd/MM/yyyy") == DateTime.MinValue)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array7[1], array7[0], text8);
					value = 2;
				}
				break;
			case "3":
				if (string.IsNullOrEmpty(text15))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array8[1], array8[0], text2);
					value = 2;
				}
				else if (ConvertHelpers.ConvertToDateTime(text15, "dd/MM/yyyy") == DateTime.MinValue)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array8[1], array8[0], text9);
					value = 2;
				}
				break;
			}
			for (int num6 = 0; num6 < hTKKCMCTableInfo.RowsValues.Count; num6++)
			{
				array[0] = ConvertHelpers.ConvertToString(sheetView.Cells[num + num6, odrColMST].Value);
				array[1] = ConvertHelpers.ConvertToString(sheetView.Cells[num + num6, num2].Value);
				array[2] = ConvertHelpers.ConvertToString(sheetView.Cells[num + num6, num3].Value);
				array[3] = ConvertHelpers.ConvertToString(sheetView.Cells[num + num6, num4].Value);
				array[4] = ConvertHelpers.ConvertToString(sheetView.Cells[num + num6, num5].Value);
				HTKKFile.RemoveValidate(sheetView, num + num6, odrColMST);
				HTKKFile.RemoveValidate(sheetView, num + num6, num2);
				HTKKFile.RemoveValidate(sheetView, num + num6, num3);
				HTKKFile.RemoveValidate(sheetView, num + num6, num4);
				HTKKFile.RemoveValidate(sheetView, num + num6, num5);
				if (string.IsNullOrEmpty(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num6, odrColMST, text2, num6 + 1);
					value = 2;
				}
				if (!string.IsNullOrEmpty(array[0]))
				{
					List<HTKKCMCRowValue> rowIndexInTableByCompare = GeneralFunctions.GetRowIndexInTableByCompare(hTKKCMCTableInfo, indexCol, array[0]);
					if (rowIndexInTableByCompare.Count > 1)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + num6, odrColMST, text4, num6 + 1);
						value = 2;
					}
				}
				if (!string.IsNullOrEmpty(array[0]) && !ValidateBase.CheckMaSoThueCCCD(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num6, odrColMST, text5, num6 + 1);
					value = 2;
				}
				if (string.IsNullOrEmpty(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num6, num2, text2, num6 + 1);
					value = 2;
				}
				if (ConvertHelpers.ConvertToDouble(array[2]) == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num6, num3, text2, num6 + 1);
					value = 2;
				}
				else if (ConvertHelpers.ConvertToDouble(array[2]) < 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num6, num3, text6, num6 + 1);
					value = 2;
				}
				else if (ConvertHelpers.ConvertToDouble(array[2]) < ConvertHelpers.ConvertToDouble(array[3]) + ConvertHelpers.ConvertToDouble(array[4]))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, num + num6, num3, text10);
				}
				if (ConvertHelpers.ConvertToDouble(array[3]) < 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num6, num4, text6, num6 + 1);
					value = 2;
				}
				if (ConvertHelpers.ConvertToDouble(array[4]) < 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + num6, num5, text6, num6 + 1);
					value = 2;
				}
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
				GeneralFunctions.ShowBangKe("Bang_Ke_Ds_tien_nop_TNCN.xls");
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
			int[] array = GeneralFunctions.ArrayCellId("K_28");
			int[] array2 = GeneralFunctions.ArrayCellId("C_28");
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
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int column = HTKKFile.ColumnLettersToIndex("B") - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				sheetView.Cells[num + i, column].Value = i + 1;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
