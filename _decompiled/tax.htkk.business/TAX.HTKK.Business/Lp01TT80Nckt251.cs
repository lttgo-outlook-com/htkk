using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Lp01TT80Nckt251
{
	private string DEFAULT_LOAI_TIEN_KHAC = "0,00";

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	private int col2_Tb1 = HTKKFile.ColumnLettersToIndex("C");

	private int col3_Tb1 = HTKKFile.ColumnLettersToIndex("N");

	private int col4_Tb1 = HTKKFile.ColumnLettersToIndex("O");

	private int col3Ma_Tb1 = HTKKFile.ColumnLettersToIndex("V");

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int iCol = HTKKFile.ColumnLettersToIndex("O");
			SheetView sView = winFpSpread.Sheets[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string value = hTKKCMCTableInfo.RowsValues[i].RowValue[4];
				string text = hTKKCMCTableInfo.RowsValues[i].RowValue[6];
				HTKKCMCValidate validate = hTKKCMCTableInfo.RowsInfo.Cells[4].Validate;
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, iCol, value, validate, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[0].Tables[1];
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
			{
				string value = hTKKCMCTableInfo2.RowsValues[i].RowValue[2];
				string text = hTKKCMCTableInfo2.RowsValues[i].RowValue[3];
				HTKKCMCValidate validate = hTKKCMCTableInfo2.RowsInfo.Cells[2].Validate;
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart2 + i, iCol, value, validate, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
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
			int num = HTKKFile.ColumnLettersToIndex("V");
			int num2 = HTKKFile.ColumnLettersToIndex("O");
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
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("O");
			if (iCol == col2_Tb1 || iCol == col3_Tb1)
			{
				List<int> list = new List<int>();
				list.Add(col2_Tb1);
				list.Add(col3_Tb1);
				List<string> list2 = new List<string>();
				list2.Add(sheetView.Cells[(int)iRow - 1, col2_Tb1 - 1].Text.Trim());
				list2.Add(sheetView.Cells[(int)iRow - 1, col3_Tb1 - 1].Text.Trim());
				List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, hTKKCMCTableInfo, list, list2);
				if (rowIndexInSViewByCompares.Count > 1)
				{
					string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24"));
					HTKKFile.SetNodeBackGroundPink(sheetView, (int)iRow - 1, col2_Tb1 - 1, message);
					HTKKFile.SetNodeBackGroundPink(sheetView, (int)iRow - 1, col3_Tb1 - 1, message);
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, (int)iRow - 1, col2_Tb1 - 1);
					HTKKFile.RemoveValidate(sheetView, (int)iRow - 1, col3_Tb1 - 1);
				}
			}
			if (iCol == col3_Tb1)
			{
				string value = sheetView.Cells[(int)iRow - 1, col4_Tb1 - 1].Value.ToString().Trim();
				string text = sheetView.Cells[(int)iRow - 1, col3Ma_Tb1 - 1].Value.ToString().Trim();
				HTKKCMCValidate validate = hTKKCMCTableInfo.RowsInfo.Cells[4].Validate;
				string defaultValue = ((text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				GeneralFunctions.InsertNumberCell(sheetView, (int)iRow, col4_Tb1, value, validate, defaultValue);
			}
			CongTheoLoaiTien(winFpSpread, _ToKhaiInfo);
			GeneralFunctions.SaveAllCellTypeOfSheet(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0], 0);
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
			int num2 = HTKKFile.ColumnLettersToIndex("C") - 1;
			int num3 = HTKKFile.ColumnLettersToIndex("N") - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sheetView, num + i, num2);
				HTKKFile.RemoveValidate(sheetView, num + i, num3);
				HTKKFile.RemoveValidate(sheetView, num + i, col2_Tb1 - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, col3_Tb1 - 1);
				string text = sheetView.Cells[num + i, num2].Text;
				string text2 = sheetView.Cells[num + i, num3].Text;
				if (string.IsNullOrEmpty(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num2, message);
					value = 2;
				}
				if (string.IsNullOrEmpty(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, num3, message);
					value = 2;
				}
				List<int> list = new List<int>();
				list.Add(col2_Tb1);
				list.Add(col3_Tb1);
				List<string> list2 = new List<string>();
				list2.Add(sheetView.Cells[num + i, col2_Tb1 - 1].Text.Trim());
				list2.Add(sheetView.Cells[num + i, col3_Tb1 - 1].Text.Trim());
				List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, hTKKCMCTableInfo, list, list2);
				if (rowIndexInSViewByCompares.Count > 1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col2_Tb1 - 1, message2);
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, col3_Tb1 - 1, message2);
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
		CongTheoLoaiTien(winFpSpread, _ToKhaiInfo);
		return true;
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			CongTheoLoaiTien(winFpSpread, _ToKhaiInfo);
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
			CongTheoLoaiTien(winFpSpread, _ToKhaiInfo);
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

	public void CapNhatLaiCellType(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		int num = HTKKFile.ColumnLettersToIndex("V");
		int num2 = HTKKFile.ColumnLettersToIndex("O");
		HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
		{
			string text = sView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Value.ToString().Trim();
			string text2 = sView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Value.ToString().Trim();
			HTKKCMCValidate validate = hTKKCMCTableInfo.RowsInfo.Cells[4].Validate;
			string defaultValue = ((text2 == "" || text2 == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, num2, text.ToString(), validate, defaultValue);
		}
		HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[0].Tables[1];
		int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
		{
			string text = hTKKCMCTableInfo2.RowsValues[i].RowValue[2];
			string text2 = sView.Cells[dataLocationTemplateRowStart2 + i - 1, num - 1].Value.ToString().Trim();
			HTKKCMCValidate validate = hTKKCMCTableInfo2.RowsInfo.Cells[2].Validate;
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart2 + i, num2, text, validate, (text2 == "" || text2 == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
		}
		GeneralFunctions.SaveAllCellTypeOfSheet(sView, ToKhaiInfo.SheetInfos[0], 0);
	}

	public void CongTheoLoaiTien(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		int colTenFrom = HTKKFile.ColumnLettersToIndex("N");
		int num = HTKKFile.ColumnLettersToIndex("O");
		int num2 = HTKKFile.ColumnLettersToIndex("V");
		int colTenTo = HTKKFile.ColumnLettersToIndex("C");
		int num3 = HTKKFile.ColumnLettersToIndex("O");
		Dictionary<int, int> dictionary = new Dictionary<int, int>();
		dictionary.Add(num, num);
		GeneralFunctions.CopyPaseTableToTable(hTKKCMCSheetInfo.Tables[0], hTKKCMCSheetInfo.Tables[1], sheetView, sheetView, dictionary, num2, num2, colTenFrom, colTenTo, "", hTKKCMCSheetInfo, -1, "", usingCellType: true);
	}
}
