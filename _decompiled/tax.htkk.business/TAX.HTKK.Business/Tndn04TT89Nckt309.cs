using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tndn04TT89Nckt309
{
	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("CR_15");
			int[] array2 = GeneralFunctions.ArrayCellId("S_27");
			int[] array3 = GeneralFunctions.ArrayCellId("CR_27");
			int[] array4 = GeneralFunctions.ArrayCellId("CS_11");
			int[] array5 = GeneralFunctions.ArrayCellId("CT_11");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			DateTime dateTime = DateTime.ParseExact(PeriodInfo.TuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			DateTime value = DateTime.ParseExact(PeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			double num = Math.Round((double)dateTime.Subtract(value).Days / 30.4375, 3);
			sView.Cells[array[1], array[0]].Value = num;
			if (PeriodInfo.KieuKy == "D")
			{
				string pathCt15 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct15";
				string pathCt16 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct16";
				string pathCt17 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct17";
				string pathCt18 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct18";
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCt15);
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCt16);
				HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCt17);
				HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCt18);
				sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Locked = true;
				sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = 0;
				sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Locked = true;
				sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = 0;
				sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Formula = "";
				sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Formula = "";
			}
			if (PeriodInfo.KieuKy == "Y" && !string.IsNullOrEmpty(PeriodInfo.MaQuyetToan))
			{
				sView.Cells[array2[1], array2[0]].Value = PeriodInfo.TenQuyetToan;
				sView.Cells[array3[1], array3[0]].Value = PeriodInfo.MaQuyetToan;
			}
			if (!string.IsNullOrEmpty(HeaderInfo.NgayBatDauNamTaiChinh))
			{
				sView.Cells[array4[1], array4[0]].Text = HeaderInfo.NgayBatDauNamTaiChinh;
				sView.Cells[array5[1], array5[0]].Text = GeneralFunctions.GetNgayTaiChinhDen(HeaderInfo.NgayBatDauNamTaiChinh);
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
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("J");
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26"));
			foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
			{
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < table.RowsValues.Count; i++)
				{
					string text = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num - 1);
					if (!string.IsNullOrEmpty(text))
					{
						List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, table, num, text);
						if (rowIndexInSViewByCompare.Count > 1)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num - 1, message);
							value = 2;
						}
						string text2 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
						if (text2.StartsWith("-----") || text2.StartsWith("---"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num - 1, message2);
							value = 2;
						}
					}
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

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		SheetView sheetView = winFpSpread.Sheets[0];
		HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == "04_TNDN_1");
		if (iRow < hTKKCMCTableInfo.Location.DataLocationTemplateRowStart || iRow > hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count)
		{
			return;
		}
		int num = HTKKFile.ColumnLettersToIndex("J");
		int num2 = 0;
		if (iCol == num)
		{
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26"));
			string text = sheetView.Cells[(int)iRow - 1, num - 1].Text.Trim();
			if (text.StartsWith("-----") || text.StartsWith("---"))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, (int)iRow - 1, num - 1, message);
				num2 = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, (int)iRow - 1, num - 1);
			}
			if (num2 == 2)
			{
				winFpSpread.TabStrip[0].BackColor = Color.Pink;
			}
			else
			{
				winFpSpread.TabStrip[0].BackColor = Color.White;
			}
		}
	}
}
