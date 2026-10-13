using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class CcttDlvc01Nckt305
{
	private string nameColEnd = "BQ";

	private int[] cellLoaiToKhai = GeneralFunctions.ArrayCellId("BJ_25");

	private bool dataChange = false;

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int num = 0;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[num];
			if (string.IsNullOrEmpty(PeriodInfo.SoLan) && !string.IsNullOrEmpty(PeriodInfo.LanXuatBan))
			{
				PeriodInfo.SoLan = PeriodInfo.LanXuatBan;
				TTinChung.soLan = PeriodInfo.SoLan;
			}
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			GeneralFunctions._page = 0;
			GeneralFunctions.dataInfoOld = null;
			GeneralFunctions.showPageLargeData1(winFpSpread, num, hTKKCMCSheetInfo, GeneralFunctions._page, nameColEnd, "U8");
			setLaiGiaTriKhoiTao(winFpSpread.Sheets[num], hTKKCMCSheetInfo, PeriodInfo, TTinChung);
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
			dataChange = true;
			int num = 0;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sView = winFpSpread.Sheets[0];
			if (!GeneralFunctions.UpdateFromSViewToSInfoLarge(sView, (int)iRow, -1, hTKKCMCSheetInfo.Tables[0].RowsInfo, hTKKCMCSheetInfo, 0))
			{
				GeneralFunctions.UpdateFromSViewToSInfoStaticLarge(sView, (int)iRow, (int)iCol, hTKKCMCSheetInfo);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int iRow, int iCol)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			if (activeRowIndex == 4 && activeColumnIndex == 1)
			{
				GeneralFunctions.ShowBangKe("BangKe_01_CCTT_DLVC_305.xlsx");
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

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int num = 0;
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[num];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[num];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_12"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13"));
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			StringBuilder stringBuilder = new StringBuilder();
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			List<HTKKCMCTableCellInfo> cells = hTKKCMCTableInfo.RowsInfo.Cells;
			int index = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "CT02");
			int index2 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "CT04");
			int num2 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "CT06_ma");
			int num3 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "CT06_ten");
			int index3 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "CT07");
			int index4 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "CT10");
			int num4 = cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == "CT08");
			string tableName = "01_CCTT_DVVC";
			string text = "Chỉ tiêu ";
			for (int num5 = 0; num5 < hTKKCMCTableInfo.RowsValues.Count; num5++)
			{
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				HTKKCMCRowValue hTKKCMCRowValue = rowsValues[num5];
				hTKKCMCRowValue.RowValue = hTKKCMCRowValue.RowValue.GetRange(0, cells.Count);
				hTKKCMCRowValue.RowValue.Take(hTKKCMCTableInfo.RowsInfo.Cells.Count).ToList();
				string value2 = rowsValues[num5].RowValue[index];
				if (string.IsNullOrEmpty(value2))
				{
					hTKKCMCRowValue.RowValue.Add("error");
					hTKKCMCRowValue.RowValue.Add(cells[index].CellTemplate.ColIndex.ToString());
					string text2 = text + "(2) " + message;
					hTKKCMCRowValue.RowValue.Add(text2);
					value = 2;
					hTKKCMCRowValue.RowValue.Add(value.ToString());
					GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text2, hTKKCMCRowValue.RowValue[0]);
				}
				string value3 = rowsValues[num5].RowValue[index2];
				if (string.IsNullOrEmpty(value3))
				{
					hTKKCMCRowValue.RowValue.Add("error");
					hTKKCMCRowValue.RowValue.Add(cells[index2].CellTemplate.ColIndex.ToString());
					string text2 = text + "(4) " + message;
					hTKKCMCRowValue.RowValue.Add(text2);
					value = 2;
					hTKKCMCRowValue.RowValue.Add(value.ToString());
					GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text2, hTKKCMCRowValue.RowValue[0]);
				}
				string value4 = rowsValues[num5].RowValue[index3];
				if (string.IsNullOrEmpty(value4))
				{
					hTKKCMCRowValue.RowValue.Add("error");
					hTKKCMCRowValue.RowValue.Add(cells[index3].CellTemplate.ColIndex.ToString());
					string text2 = text + "(7) " + message;
					hTKKCMCRowValue.RowValue.Add(text2);
					value = 2;
					hTKKCMCRowValue.RowValue.Add(value.ToString());
					GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text2, hTKKCMCRowValue.RowValue[0]);
				}
				string text3 = rowsValues[num5].RowValue[index4];
				if (string.IsNullOrEmpty(text3))
				{
					hTKKCMCRowValue.RowValue.Add("error");
					hTKKCMCRowValue.RowValue.Add(cells[index4].CellTemplate.ColIndex.ToString());
					string text2 = text + "(10) " + message;
					hTKKCMCRowValue.RowValue.Add(text2);
					value = 2;
					hTKKCMCRowValue.RowValue.Add(value.ToString());
					GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text2, hTKKCMCRowValue.RowValue[0]);
				}
				else if (!GeneralFunctions.IsNumeric(text3))
				{
					hTKKCMCRowValue.RowValue.Add("error");
					hTKKCMCRowValue.RowValue.Add(cells[index4].CellTemplate.ColIndex.ToString());
					string text2 = text + "(10) " + message2;
					hTKKCMCRowValue.RowValue.Add(text2);
					value = 2;
					hTKKCMCRowValue.RowValue.Add(value.ToString());
					GeneralFunctions.LogErrorBigData(sheetView.SheetName, tableName, text2, hTKKCMCRowValue.RowValue[0]);
				}
			}
			if (value == 2)
			{
				hTKKCMCTableInfo.RowsValues = hTKKCMCTableInfo.RowsValues.OrderByDescending((HTKKCMCRowValue row) => row.RowValue.Count > cells.Count && row.RowValue[cells.Count] == "error" && row.RowValue[cells.Count + 3] == 2.ToString()).ToList();
			}
			else
			{
				hTKKCMCTableInfo.RowsValues = hTKKCMCTableInfo.RowsValues.OrderBy((HTKKCMCRowValue x) => int.Parse(x.RowValue[0])).ToList();
			}
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			GeneralFunctions._page = (int)Math.Round((decimal)dataLocationTemplateRowStart2 / (decimal)GeneralFunctions._pageSize, 0) - 1;
			GeneralFunctions.showPageLargeData1(winFpSpread, num, hTKKCMCSheetInfo, GeneralFunctions._page, nameColEnd, "U8");
			setLaiGiaTriKhoiTao(winFpSpread.Sheets[num], hTKKCMCSheetInfo, PeriodInfo, TTinChung);
			if (value == 2)
			{
				GeneralFunctions.showRowsErrorLarge(sheetView);
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

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		int num = 0;
		HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[num];
		dataChange = true;
		winFpSpread.SuspendLayout();
		winFpSpread.ActiveSheet.AutoCalculation = false;
		SheetView activeSheet = winFpSpread.ActiveSheet;
		GeneralFunctions._page = 0;
		GeneralFunctions.showPageLargeData1(winFpSpread, num, sInfo, GeneralFunctions._page, nameColEnd, "U8");
		setLaiGiaTriKhoiTao(winFpSpread.Sheets[num], sInfo, PeriodInfo, TTinChung);
		winFpSpread.ResumeLayout();
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int num = 0;
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[num];
			dataChange = true;
			winFpSpread.SuspendLayout();
			winFpSpread.ActiveSheet.AutoCalculation = false;
			GeneralFunctions._page = 0;
			GeneralFunctions.showPageLargeData1(winFpSpread, num, sInfo, GeneralFunctions._page, nameColEnd, "U8");
			setLaiGiaTriKhoiTao(winFpSpread.Sheets[num], sInfo, PeriodInfo, TTinChung);
			winFpSpread.ActiveSheet.AutoCalculation = true;
			winFpSpread.ResumeLayout();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpadatePage_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int page)
	{
		try
		{
			int num = 0;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[num];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			GeneralFunctions.showPageLargeData1(winFpSpread, num, hTKKCMCSheetInfo, page, nameColEnd, "U8");
			setLaiGiaTriKhoiTao(winFpSpread.Sheets[num], hTKKCMCSheetInfo, PeriodInfo, TTinChung);
			GeneralFunctions.showRowsErrorLarge(winFpSpread.Sheets[num]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool UpdateFormula_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		int num = 0;
		HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[num];
		GeneralFunctions.showPageLargeData1(winFpSpread, num, ToKhaiInfo.SheetInfos[num], GeneralFunctions._page, nameColEnd, "U8");
		setLaiGiaTriKhoiTao(winFpSpread.Sheets[num], sInfo, PeriodInfo, TTinChung);
		return true;
	}

	public void setLaiGiaTriKhoiTao(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKPeriodInfo PeriodInfo, HTKKCMCTTinChungInfo TTinChung)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KinhGui/File");
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KinhGui/maCQT");
		HTKKCMCCellInfo hTKKCMCCellInfo3 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KinhGui/tenCQT");
		HTKKCMCCellInfo hTKKCMCCellInfo4 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KinhGui/Loaitokhai");
		hTKKCMCCellInfo.Value = PeriodInfo.LanXuatBan;
		hTKKCMCCellInfo2.Value = TTinChung.maCQTNoiNop;
		hTKKCMCCellInfo3.Value = TTinChung.tenCQTNoiNop;
		hTKKCMCCellInfo4.Value = ((PeriodInfo.LoaiToKhai == "T") ? "04" : "01");
		sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text = ((PeriodInfo.LoaiToKhai == "T") ? "04" : "01");
		sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = PeriodInfo.LanXuatBan;
		sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text = TTinChung.maCQTNoiNop;
		sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text = TTinChung.tenCQTNoiNop;
	}
}
