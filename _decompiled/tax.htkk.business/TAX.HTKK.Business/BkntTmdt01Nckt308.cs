using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BkntTmdt01Nckt308
{
	private const string So_TK = "1";

	private const string So_QD = "2";

	private const string So_TB = "3";

	private string tieuMuc1 = "1749";

	private string tieuMuc2 = "1099";

	private string tieuMuc3 = "1701";

	private string tieuMuc4 = "1003";

	private int[] cellLoaiKyTen = GeneralFunctions.ArrayCellId("D_75");

	private int[] cellLoaiKyMa = GeneralFunctions.ArrayCellId("O_75");

	private int[] cellCt05 = GeneralFunctions.ArrayCellId("O_66");

	private int[] cellkyThue = GeneralFunctions.ArrayCellId("F_75");

	private int[] cellNgayQD = GeneralFunctions.ArrayCellId("D_77");

	private int[] cellNgayTB = GeneralFunctions.ArrayCellId("D_79");

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("O_50");
			sView.Cells[array[1], array[0]].Text = PeriodInfo.LanXuatBan;
			int[] array2 = GeneralFunctions.ArrayCellId("E_38");
			sView.Cells[array2[1], array2[0]].Text = PeriodInfo.MaHSDuAnDT;
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
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			if (iRow == cellCt05[1] + 1)
			{
				string text = sheetView.Cells[cellCt05[1], cellCt05[0]].Text;
				if (text == "1")
				{
					sheetView.Cells[cellLoaiKyMa[1], cellLoaiKyMa[0]].Text = "M";
					sheetView.Cells[cellLoaiKyTen[1], cellLoaiKyTen[0]].Text = "Tháng";
					GeneralFunctions.LockUnLockCol(sheetView, cellLoaiKyTen[1] + 1, cellLoaiKyTen[0] + 1, locked: true, deleteData: false);
					GeneralFunctions.LockUnLockCol(sheetView, cellkyThue[1] + 1, cellkyThue[0] + 1, locked: false, deleteData: true);
					MaskCellType maskCellType = new MaskCellType();
					maskCellType.MaskChar = '.';
					maskCellType.Mask = "00/0000";
					sheetView.Cells[cellkyThue[1], cellkyThue[0]].Text = "";
					sheetView.Cells[cellkyThue[1], cellkyThue[0]].ResetCellType();
					sheetView.Cells[cellkyThue[1], cellkyThue[0]].CellType = maskCellType;
				}
				else
				{
					sheetView.Cells[cellLoaiKyMa[1], cellLoaiKyMa[0]].Text = "";
					GeneralFunctions.LockUnLockCol(sheetView, cellLoaiKyTen[1] + 1, cellLoaiKyTen[0] + 1, locked: true, deleteData: true);
					GeneralFunctions.LockUnLockCol(sheetView, cellkyThue[1] + 1, cellkyThue[0] + 1, locked: true, deleteData: true);
				}
				if (text == "2")
				{
					GeneralFunctions.LockUnLockCol(sheetView, cellNgayQD[1] + 1, cellNgayQD[0] + 1, locked: false, deleteData: true);
				}
				else
				{
					GeneralFunctions.LockUnLockCol(sheetView, cellNgayQD[1] + 1, cellNgayQD[0] + 1, locked: true, deleteData: true);
				}
				if (text == "3")
				{
					GeneralFunctions.LockUnLockCol(sheetView, cellNgayTB[1] + 1, cellNgayTB[0] + 1, locked: false, deleteData: true);
				}
				else
				{
					GeneralFunctions.LockUnLockCol(sheetView, cellNgayTB[1] + 1, cellNgayTB[0] + 1, locked: true, deleteData: true);
				}
			}
			if (iRow == cellLoaiKyTen[1] + 1 && iCol == cellLoaiKyTen[0] + 1)
			{
				string text2 = sheetView.Cells[cellLoaiKyMa[1], cellLoaiKyMa[0]].Text;
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
				sheetView.Cells[cellkyThue[1], cellkyThue[0]].Text = "";
				sheetView.Cells[cellkyThue[1], cellkyThue[0]].ResetCellType();
				sheetView.Cells[cellkyThue[1], cellkyThue[0]].CellType = maskCellType;
			}
			int num = GeneralFunctions.DetectIndexOfTable(sInfo, (int)iRow);
			if (num >= 0)
			{
				GeneralFunctions.UpdateFromSViewToSInfoStaticLarge(sheetView, (int)iRow, (int)iCol, sInfo);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
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
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[11, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[12, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			int num = HTKKFile.ColumnLettersToIndex("B");
			int colCt2 = HTKKFile.ColumnLettersToIndex("D");
			int colCt3 = HTKKFile.ColumnLettersToIndex("E");
			int colCt4 = HTKKFile.ColumnLettersToIndex("F");
			int colCt5 = HTKKFile.ColumnLettersToIndex("G");
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				int index = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt2);
				int index2 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt3);
				int index3 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt4);
				int index4 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == colCt5);
				for (int num2 = 0; num2 < table.RowsValues.Count; num2++)
				{
					int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
					foreach (HTKKCMCTableCellInfo cell in table.RowsInfo.Cells)
					{
						if (num < cell.CellTemplate.ColIndex && cell.CellTemplate.ColIndex <= colCt5)
						{
							HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + num2 - 1, cell.CellTemplate.ColIndex - 1);
						}
					}
					bool flag = false;
					if (GeneralFunctions.HaveToValidateRow(table.RowsInfo, table.RowsValues[num2].RowValue.ToArray()))
					{
						string text3 = table.RowsValues[num2].RowValue[index];
						string value2 = table.RowsValues[num2].RowValue[index2];
						string value3 = table.RowsValues[num2].RowValue[index3];
						string text4 = table.RowsValues[num2].RowValue[index4];
						if ((table.TableName == "table_1" || table.TableName == "table_2") && string.IsNullOrEmpty(value2))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colCt3 - 1, text, num2 + 1);
							value = 2;
						}
						if ((table.TableName == "table_1" || table.TableName == "table_2") && string.IsNullOrEmpty(value3))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colCt4 - 1, text, num2 + 1);
							value = 2;
						}
						if ((table.TableName == "table_1" || table.TableName == "table_2") && text4 == "0")
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colCt5 - 1, text, num2 + 1);
							value = 2;
						}
						if (!string.IsNullOrEmpty(text3) && !ValidateBase.CheckMaSoThueCCCD(text3))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + num2 - 1, colCt2 - 1, text2, num2 + 1);
							value = 2;
						}
					}
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
			return false;
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			if (activeRowIndex == 30 && activeColumnIndex == 2)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_01_BKNT_TMDT.xls");
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
			int num = HTKKFile.ColumnLettersToIndex("O");
			int num2 = 0;
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < table.RowsValues.Count; i++)
				{
					sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text = num2 switch
					{
						3 => tieuMuc4, 
						2 => tieuMuc3, 
						1 => tieuMuc2, 
						0 => tieuMuc1, 
						_ => "", 
					};
				}
				num2++;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
