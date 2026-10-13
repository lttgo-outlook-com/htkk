using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Ph01TT80Nckt251
{
	private string DEFAULT_LOAI_TIEN_KHAC = "0,00";

	private int col2_Tb1 = HTKKFile.ColumnLettersToIndex("C");

	private int col3_Tb1 = HTKKFile.ColumnLettersToIndex("K");

	private int col4_Tb1 = HTKKFile.ColumnLettersToIndex("L");

	private int col5_Tb1 = HTKKFile.ColumnLettersToIndex("O");

	private int col6_Tb1 = HTKKFile.ColumnLettersToIndex("Q");

	private int col7_Tb1 = HTKKFile.ColumnLettersToIndex("T");

	private int col3Ma_Tb1 = HTKKFile.ColumnLettersToIndex("AE");

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string value = hTKKCMCTableInfo.RowsValues[i].RowValue[4];
				string value2 = hTKKCMCTableInfo.RowsValues[i].RowValue[6];
				string value3 = hTKKCMCTableInfo.RowsValues[i].RowValue[7];
				string text = hTKKCMCTableInfo.RowsValues[i].RowValue[3];
				HTKKCMCValidate validate = hTKKCMCTableInfo.RowsInfo.Cells[4].Validate;
				HTKKCMCValidate validate2 = hTKKCMCTableInfo.RowsInfo.Cells[6].Validate;
				HTKKCMCValidate validate3 = hTKKCMCTableInfo.RowsInfo.Cells[7].Validate;
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col4_Tb1, value, validate, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col6_Tb1, value2, validate2, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col7_Tb1, value3, validate3, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			_ToKhaiInfo = ToKhaiInfo;
			CapNhatLaiCellType(sView, ToKhaiInfo);
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("AB");
			int num2 = HTKKFile.ColumnLettersToIndex("AE");
			int num3 = HTKKFile.ColumnLettersToIndex("C");
			int num4 = HTKKFile.ColumnLettersToIndex("T");
			int num5 = HTKKFile.ColumnLettersToIndex("AE");
			SheetView sheetView = winFpSpread.Sheets[0];
			double num6 = Convert.ToDouble(sheetView.Cells[(int)iRow - 1, col4_Tb1 - 1].Value);
			double num7 = Convert.ToDouble(sheetView.Cells[(int)iRow - 1, col5_Tb1 - 1].Value);
			double num8 = Convert.ToDouble(sheetView.Cells[(int)iRow - 1, col6_Tb1 - 1].Value);
			double num9 = Convert.ToDouble(sheetView.Cells[(int)iRow - 1, col7_Tb1 - 1].Value);
			double num10 = Convert.ToDouble(sheetView.Cells[(int)iRow - 1, num - 1].Value);
			if (iCol == col3_Tb1)
			{
				string text = sheetView.Cells[(int)iRow - 1, col3_Tb1 - 1].Value.ToString().Trim();
				HTKKCMCValidate validate = hTKKCMCSheetInfo.Tables[0].RowsInfo.Cells[4].Validate;
				HTKKCMCValidate validate2 = hTKKCMCSheetInfo.Tables[0].RowsInfo.Cells[6].Validate;
				HTKKCMCValidate validate3 = hTKKCMCSheetInfo.Tables[0].RowsInfo.Cells[7].Validate;
				GeneralFunctions.InsertNumberCell(sheetView, (int)iRow, col4_Tb1, num6.ToString(), validate, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				GeneralFunctions.InsertNumberCell(sheetView, (int)iRow, col6_Tb1, num8.ToString(), validate2, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				GeneralFunctions.InsertNumberCell(sheetView, (int)iRow, col7_Tb1, num9.ToString(), validate3, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			}
			if (iCol == col4_Tb1 || iCol == col5_Tb1)
			{
				sheetView.Cells[(int)iRow - 1, col7_Tb1 - 1].Value = num10;
			}
			TongTheoLoaiTien(winFpSpread, _ToKhaiInfo);
			GeneralFunctions.fpSpReindexingTableTemplate(sheetView, hTKKCMCSheetInfo.Tables[0]);
			GeneralFunctions.SaveAllCellTypeOfSheet(sheetView, ToKhaiInfo.SheetInfos[0], 0);
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
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			return true;
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
			int[] cell = GeneralFunctions.ArrayCellId("E_22");
			string message = HTKKFile.GetMessage(sheetView, cell);
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25"));
			List<string> list = new List<string>();
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sheetView, num + i, col2_Tb1 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, col3_Tb1 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, col5_Tb1 - 1);
				string text = sheetView.Cells[num + i, col2_Tb1 - 1].Text;
				string text2 = sheetView.Cells[num + i, col3_Tb1 - 1].Text;
				string text3 = sheetView.Cells[num + i, col3Ma_Tb1 - 1].Text;
				string text4 = sheetView.Cells[num + i, col5_Tb1 - 1].Text;
				string item = text + text4;
				if (!list.Contains(item))
				{
					list.Add(item);
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col2_Tb1 - 1, message2);
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col5_Tb1 - 1, message2);
					value = 2;
				}
				if (string.IsNullOrEmpty(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col2_Tb1 - 1, message);
					value = 2;
				}
				if (string.IsNullOrEmpty(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col3_Tb1 - 1, message);
					value = 2;
				}
				if (!string.IsNullOrEmpty(text2) && string.IsNullOrEmpty(text3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col3_Tb1 - 1, message3);
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool UpdateFormula_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		TongTheoLoaiTien(winFpSpread, ToKhaiInfo);
		return true;
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			TongTheoLoaiTien(winFpSpread, _ToKhaiInfo);
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
			TongTheoLoaiTien(winFpSpread, _ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		CapNhatLaiCellType(winFpSpread.Sheets[0], ToKhaiInfo);
	}

	public void TongTheoLoaiTien(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		int num = HTKKFile.ColumnLettersToIndex("AB");
		int colMaFrom = HTKKFile.ColumnLettersToIndex("AE");
		int colTenTo = HTKKFile.ColumnLettersToIndex("C");
		int num2 = HTKKFile.ColumnLettersToIndex("T");
		int colMaTo = HTKKFile.ColumnLettersToIndex("AE");
		SheetView sheetView = winFpSpread.Sheets[0];
		Dictionary<int, int> dictionary = new Dictionary<int, int>();
		dictionary.Add(col7_Tb1, col7_Tb1);
		GeneralFunctions.CopyPaseTableToTable(hTKKCMCSheetInfo.Tables[0], hTKKCMCSheetInfo.Tables[1], sheetView, sheetView, dictionary, colMaFrom, colMaTo, col3_Tb1, colTenTo, "", hTKKCMCSheetInfo, -1, "", usingCellType: true);
	}

	public void CapNhatLaiCellType(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
		{
			string value = hTKKCMCTableInfo.RowsValues[i].RowValue[4];
			string value2 = hTKKCMCTableInfo.RowsValues[i].RowValue[6];
			string value3 = hTKKCMCTableInfo.RowsValues[i].RowValue[7];
			string text = sView.Cells[dataLocationTemplateRowStart + i - 1, col3Ma_Tb1 - 1].Value.ToString().Trim();
			HTKKCMCValidate validate = hTKKCMCTableInfo.RowsInfo.Cells[4].Validate;
			HTKKCMCValidate validate2 = hTKKCMCTableInfo.RowsInfo.Cells[6].Validate;
			HTKKCMCValidate validate3 = hTKKCMCTableInfo.RowsInfo.Cells[7].Validate;
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col4_Tb1, value, validate, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col6_Tb1, value2, validate2, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col7_Tb1, value3, validate3, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
		}
		HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[0].Tables[1];
		int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
		{
			string value3 = hTKKCMCTableInfo2.RowsValues[i].RowValue[2];
			string text = sView.Cells[dataLocationTemplateRowStart2 + i - 1, col3Ma_Tb1 - 1].Value.ToString().Trim();
			HTKKCMCValidate validate4 = hTKKCMCTableInfo2.RowsInfo.Cells[2].Validate;
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart2 + i, col7_Tb1, value3, validate4, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
		}
		GeneralFunctions.SaveAllCellTypeOfSheet(sView, ToKhaiInfo.SheetInfos[0], 0);
	}
}
