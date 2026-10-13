using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tbh01TT80Nd373Nckt308
{
	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
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

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[1];
			int num = HTKKFile.ColumnLettersToIndex("D");
			int colChilden = HTKKFile.ColumnLettersToIndex("BL");
			if (iCol == num)
			{
				int num2 = GeneralFunctions.DetectIndexOfTable(sInfo, (int)iRow);
				if (num2 >= 0)
				{
					GeneralFunctions.SetLaiDongCon(sView, sInfo, num2, -1, num, colChilden, 1);
				}
			}
			if (iCol == HTKKFile.ColumnLettersToIndex("AV"))
			{
				tinhTongCt11_pl(sView, sInfo);
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
			if (idxRow == 12 && idxCol == 2)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_01_1_TBH_ND373.xls");
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
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				switch (item.Value)
				{
				case 2:
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
					break;
				case 1:
					winFpSpread.TabStrip[key].BackColor = Color.Yellow;
					break;
				default:
					winFpSpread.TabStrip[key].BackColor = Color.White;
					break;
				}
			}
			if (winFpSpread.TabStrip[0].BackColor == Color.Pink || winFpSpread.TabStrip[1].BackColor == Color.Pink)
			{
				return false;
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			switch (0)
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

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int num = 0;
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[14, HTKKFile.ColumnLettersToIndex("E") - 1].Text.Trim();
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				foreach (HTKKCMCCellInfo staticCell in table.Header.StaticCells)
				{
					HTKKFile.RemoveValidate(sheetView, staticCell.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("D") - 1);
					string text2 = sheetView.Cells[staticCell.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Text;
					if (string.IsNullOrEmpty(text2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, staticCell.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("D") - 1, text);
						num = 2;
					}
				}
			}
			foreach (HTKKCMCTableInfo table2 in hTKKCMCSheetInfo.Tables)
			{
				List<int> allColumnsInTable = GeneralFunctions.getAllColumnsInTable(table2);
				int index = table2.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "ct1");
				int index2 = table2.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "ct10");
				int index3 = table2.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "path_header_ct2");
				int index4 = table2.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "path_key_khbs");
				List<int> list = new List<int>();
				list.Add(allColumnsInTable[index]);
				list.Add(allColumnsInTable[index2]);
				list.Add(allColumnsInTable[index3]);
				list.Add(allColumnsInTable[index4]);
				allColumnsInTable.RemoveAt(index4);
				allColumnsInTable.RemoveAt(index3);
				allColumnsInTable.RemoveAt(index2);
				allColumnsInTable.RemoveAt(index);
				int num2 = GeneralFunctions.checkValidateTableEmpty(sheetView, table2, allColumnsInTable, text, list);
				if (num2 > 0)
				{
					num = num2;
				}
			}
			switch (num)
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
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		int[] array = GeneralFunctions.ArrayCellId("S_60");
		int[] array2 = GeneralFunctions.ArrayCellId("AJ_60");
		sheetView.Cells[array[1], array[0]].Value = sheetView.Cells[array2[1], array2[0]].Value;
	}

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			winFpSpread.ActiveSheet.AutoCalculation = false;
			winFpSpread.SuspendLayout();
			if (((Control)winFpSpread).InvokeRequired)
			{
				((Control)winFpSpread).Invoke((Delegate)(System.Action)delegate
				{
					HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[1];
					SheetView sView = winFpSpread.Sheets[1];
					tinhTongCt11_pl(sView, sInfo);
				});
			}
			winFpSpread.ActiveSheet.AutoCalculation = true;
			winFpSpread.ResumeLayout();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void tinhTongCt11_pl(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		double num = 0.0;
		foreach (HTKKCMCTableInfo table in sInfo.Tables)
		{
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			int num2 = HTKKFile.ColumnLettersToIndex("AV");
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				object value = sView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Value;
				if (value != null)
				{
					num += Convert.ToDouble(value);
				}
			}
		}
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "PL_01-1_TBH_TB/tongCong");
		sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = num;
	}
}
