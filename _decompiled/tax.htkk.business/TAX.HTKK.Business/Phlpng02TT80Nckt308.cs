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

public class Phlpng02TT80Nckt308
{
	private string DEFAULT_LOAI_TIEN_KHAC = "0,00";

	private string DON_VI_TIEN_VND = "VND";

	private int[] cellTenQuyetToan = GeneralFunctions.ArrayCellId("H_19");

	private int[] cellMaQuyetToan = GeneralFunctions.ArrayCellId("AF_19");

	private int col2a_Tb1 = HTKKFile.ColumnLettersToIndex("C");

	private int col2aMa_Tb1 = HTKKFile.ColumnLettersToIndex("AF");

	private int col2b_Tb1 = HTKKFile.ColumnLettersToIndex("I");

	private int col3_Tb1 = HTKKFile.ColumnLettersToIndex("L");

	private int col4_Tb1 = HTKKFile.ColumnLettersToIndex("O");

	private int col5_Tb1 = HTKKFile.ColumnLettersToIndex("R");

	private int col6_Tb1 = HTKKFile.ColumnLettersToIndex("T");

	private int col7_Tb1 = HTKKFile.ColumnLettersToIndex("W");

	private int col8_Tb1 = HTKKFile.ColumnLettersToIndex("AB");

	private int col9_Tb1 = HTKKFile.ColumnLettersToIndex("AC");

	private int col3Ma_Tb1 = HTKKFile.ColumnLettersToIndex("AG");

	private int col7Formula_Tb1 = HTKKFile.ColumnLettersToIndex("AH");

	private int colMaTien_Tb2 = HTKKFile.ColumnLettersToIndex("AG");

	private int colTien_Tb2 = HTKKFile.ColumnLettersToIndex("C");

	private int col2_Tb1_Pl = HTKKFile.ColumnLettersToIndex("C");

	private int col3a_Tb1_Pl = HTKKFile.ColumnLettersToIndex("D");

	private int col3b_Tb1_Pl = HTKKFile.ColumnLettersToIndex("J");

	private int col4_Tb1_Pl = HTKKFile.ColumnLettersToIndex("M");

	private int col4Ma_Tb1_Pl = HTKKFile.ColumnLettersToIndex("AG");

	private int col5_Tb1_Pl = HTKKFile.ColumnLettersToIndex("P");

	private int col6_Tb1_Pl = HTKKFile.ColumnLettersToIndex("S");

	private int col7_Tb1_Pl = HTKKFile.ColumnLettersToIndex("U");

	private int col8_Tb1_Pl = HTKKFile.ColumnLettersToIndex("X");

	private int col3aMa_Tb1_Pl = HTKKFile.ColumnLettersToIndex("AF");

	private int col3bMa_Tb1_Pl = HTKKFile.ColumnLettersToIndex("AG");

	private int colTenLoaiPhi_Tb2_Pl = HTKKFile.ColumnLettersToIndex("C");

	private int colMaLoaiPhi_Tb2_Pl = HTKKFile.ColumnLettersToIndex("AF");

	private int colMaTien_Tb2_Pl = HTKKFile.ColumnLettersToIndex("AG");

	private int colTien_Tb2_Pl = HTKKFile.ColumnLettersToIndex("M");

	private int colMaTien_Tb3_Pl = HTKKFile.ColumnLettersToIndex("AG");

	private int colTien_Tb3_Pl = HTKKFile.ColumnLettersToIndex("C");

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			CapNhatLaiCellTypeTuDataFileToKhai(winFpSpread, ToKhaiInfo, 0);
			CapNhatLaiCellTypeTuDataFilePL(winFpSpread, ToKhaiInfo, 1);
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
			if (!string.IsNullOrEmpty(PeriodInfo.MaQuyetToan))
			{
				sView.Cells[cellTenQuyetToan[1], cellTenQuyetToan[0]].Value = PeriodInfo.TenQuyetToan;
				sView.Cells[cellMaQuyetToan[1], cellMaQuyetToan[0]].Value = PeriodInfo.MaQuyetToan;
			}
			_ToKhaiInfo = ToKhaiInfo;
			CapNhatLaiCellTypeTuSViewToKhai(sView, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			CapNhatLaiCellTypeTuSViewPl(sView, ToKhaiInfo);
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
			SheetView sheetView = winFpSpread.Sheets[0];
			if (iCol == col4_Tb1 || iCol == col5_Tb1)
			{
				sheetView.Cells[(int)iRow - 1, col7_Tb1 - 1].Value = sheetView.Cells[(int)iRow - 1, col7Formula_Tb1 - 1].Value;
			}
			if (iCol == col2a_Tb1)
			{
				sheetView.Cells[(int)iRow - 1, col2b_Tb1 - 1].Value = sheetView.Cells[(int)iRow - 1, col2aMa_Tb1 - 1].Value;
			}
			TongTheoLoaiTienToKhai(winFpSpread, _ToKhaiInfo);
			CapNhatLaiCellTypeTuSViewToKhai(winFpSpread.Sheets[0], ToKhaiInfo);
			GeneralFunctions.fpSpReindexingTableTemplate(sheetView, hTKKCMCSheetInfo.Tables[0]);
			GeneralFunctions.SaveAllCellTypeOfSheet(sheetView, ToKhaiInfo.SheetInfos[0], 0);
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			SheetView sheetView = winFpSpread.Sheets[1];
			if (iCol == col3a_Tb1_Pl)
			{
				sheetView.Cells[(int)iRow - 1, col3b_Tb1_Pl - 1].Value = sheetView.Cells[(int)iRow - 1, col3aMa_Tb1_Pl - 1].Value;
			}
			TongTheoLoaiPhiPhuLuc(winFpSpread, ToKhaiInfo);
			TongTheoLoaiTienPhuLuc(winFpSpread, ToKhaiInfo);
			CapNhatLaiCellTypeTuSViewPl(winFpSpread.Sheets[1], ToKhaiInfo);
			GeneralFunctions.fpSpReindexingTableTemplate(sheetView, hTKKCMCSheetInfo.Tables[0]);
			GeneralFunctions.SaveAllCellTypeOfSheet(sheetView, ToKhaiInfo.SheetInfos[0], 0);
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
			if (idxRow == 6 && idxCol == 1)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_02_1_PHLPNG_TT80.xls");
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
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
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
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			SheetView sView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo2.Tables[1];
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			int[] cell = GeneralFunctions.ArrayCellId("E_22");
			string message = HTKKFile.GetMessage(sheetView, cell);
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24"));
			string message3 = HTKKFile.GetMessage(winFpSpread.Sheets[winFpSpread.Sheets.Count - 1], GeneralFunctions.ArrayCellId("E_24"));
			string message4 = HTKKFile.GetMessage(winFpSpread.Sheets[winFpSpread.Sheets.Count - 1], GeneralFunctions.ArrayCellId("E_25"));
			int num = HTKKFile.ColumnLettersToIndex("X");
			int item = HTKKFile.ColumnLettersToIndex("AG");
			int item2 = HTKKFile.ColumnLettersToIndex("AF");
			int num2 = HTKKFile.ColumnLettersToIndex("W");
			int num3 = HTKKFile.ColumnLettersToIndex("AG");
			int num4 = HTKKFile.ColumnLettersToIndex("I");
			List<string> list = new List<string>();
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, col2a_Tb1 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, col3_Tb1 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, col5_Tb1 - 1);
				HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, col7_Tb1 - 1);
				string text = sheetView.Cells[dataLocationTemplateRowStart + i - 1, col2a_Tb1 - 1].Text;
				string text2 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, col3_Tb1 - 1].Text;
				string text3 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, col5_Tb1 - 1].Text;
				string item3 = text + text3;
				if (!list.Contains(item3))
				{
					list.Add(item3);
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, col2a_Tb1 - 1, message2);
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, col5_Tb1 - 1, message2);
					value = 2;
				}
				if (string.IsNullOrEmpty(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, col2a_Tb1 - 1, message);
					value = 2;
				}
				if (string.IsNullOrEmpty(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, col3_Tb1 - 1, message);
					value = 2;
				}
				if (ToKhaiInfo.SheetInfos[1] == null)
				{
					continue;
				}
				double num5 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + i - 1, col7_Tb1 - 1].Value);
				string text4 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Text.Trim();
				string text5 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Text.Trim();
				List<int> list2 = new List<int>();
				list2.Add(item2);
				list2.Add(item);
				List<string> list3 = new List<string>();
				list3.Add(text5);
				list3.Add(text4);
				if (!string.IsNullOrEmpty(text4))
				{
					List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sView, hTKKCMCTableInfo2, list2, list3);
					if (rowIndexInSViewByCompares.Count <= 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, col7_Tb1 - 1, message4 + " tiểu mục " + text5 + " loại tiền " + text4);
						value = 2;
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			SheetView sheetView2 = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[1];
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo2.Tables[1];
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			string message = HTKKFile.GetMessage(winFpSpread.Sheets[winFpSpread.Sheets.Count - 1], GeneralFunctions.ArrayCellId("E_24"));
			string message2 = HTKKFile.GetMessage(winFpSpread.Sheets[winFpSpread.Sheets.Count - 1], GeneralFunctions.ArrayCellId("E_26"));
			int num = HTKKFile.ColumnLettersToIndex("X");
			int num2 = HTKKFile.ColumnLettersToIndex("AG");
			int num3 = HTKKFile.ColumnLettersToIndex("AF");
			int num4 = HTKKFile.ColumnLettersToIndex("W");
			int item = HTKKFile.ColumnLettersToIndex("AG");
			int item2 = HTKKFile.ColumnLettersToIndex("I");
			List<string> list = new List<string>();
			for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidateBackGround(sheetView2, dataLocationTemplateRowStart2 + i - 1, num - 1);
				string text = sheetView2.Cells[dataLocationTemplateRowStart2 + i - 1, num2 - 1].Text.Trim();
				double num5 = Convert.ToDouble(sheetView2.Cells[dataLocationTemplateRowStart2 + i - 1, num - 1].Value);
				if (text.Equals(DON_VI_TIEN_VND))
				{
					num5 = Math.Round(Convert.ToDouble(sheetView2.Cells[dataLocationTemplateRowStart2 + i - 1, num - 1].Value), 0);
				}
				string text2 = sheetView2.Cells[dataLocationTemplateRowStart2 + i - 1, num3 - 1].Text.Trim();
				List<int> list2 = new List<int>();
				list2.Add(item2);
				list2.Add(item);
				List<string> list3 = new List<string>();
				list3.Add(text2);
				list3.Add(text);
				double num6 = 0.0;
				if (string.IsNullOrEmpty(text))
				{
					continue;
				}
				List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, hTKKCMCTableInfo, list2, list3);
				if (rowIndexInSViewByCompares.Count > 0)
				{
					foreach (int item3 in rowIndexInSViewByCompares)
					{
						num6 = ((!text.Equals(DON_VI_TIEN_VND)) ? (num6 + Convert.ToDouble(sheetView.Cells[item3 - 1, num4 - 1].Value)) : (num6 + Math.Round(Convert.ToDouble(sheetView.Cells[item3 - 1, num4 - 1].Value), 0)));
					}
				}
				if (rowIndexInSViewByCompares.Count > 0)
				{
					if (num6 != num5)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView2, dataLocationTemplateRowStart2 + i - 1, num - 1, message);
						value = 2;
					}
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView2, dataLocationTemplateRowStart2 + i - 1, num - 1, message2 + text2 + " loại tiền " + text);
					value = 2;
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
		}
	}

	public bool UpdateFormula_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		TongTheoLoaiTienToKhai(winFpSpread, ToKhaiInfo);
		return true;
	}

	public bool UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		TongTheoLoaiPhiPhuLuc(winFpSpread, ToKhaiInfo);
		TongTheoLoaiTienPhuLuc(winFpSpread, ToKhaiInfo);
		return true;
	}

	public void ClearData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			TongTheoLoaiTienToKhai(winFpSpread, _ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_0(FpSpread winFpSpread, int actionType)
	{
		try
		{
			TongTheoLoaiTienToKhai(winFpSpread, _ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			TongTheoLoaiPhiPhuLuc(winFpSpread, ToKhaiInfo);
			TongTheoLoaiTienPhuLuc(winFpSpread, ToKhaiInfo);
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
			TongTheoLoaiPhiPhuLuc(winFpSpread, _ToKhaiInfo);
			TongTheoLoaiTienPhuLuc(winFpSpread, _ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		CapNhatLaiCellTypeTuSViewToKhai(winFpSpread.Sheets[0], ToKhaiInfo);
		CapNhatLaiCellTypeTuSViewPl(winFpSpread.Sheets[1], ToKhaiInfo);
	}

	public void TongTheoLoaiTienToKhai(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sView = winFpSpread.Sheets[0];
			Dictionary<int, int> dictionary = new Dictionary<int, int>();
			dictionary.Add(col3_Tb1, colTien_Tb2);
			dictionary.Add(col3Ma_Tb1, colMaTien_Tb2);
			dictionary.Add(col4_Tb1, col4_Tb1);
			dictionary.Add(col6_Tb1, col6_Tb1);
			dictionary.Add(col7_Tb1, col7_Tb1);
			Dictionary<string, int> checkExists = new Dictionary<string, int>();
			GeneralFunctions.CopyTableToTable2(sView, hTKKCMCSheetInfo, hTKKCMCSheetInfo.Tables[0], 1, dictionary, deleteDataOld: true, ref checkExists, usingCellType: true);
			GeneralFunctions.fpSpReindexingTableTemplate(sView, hTKKCMCSheetInfo.Tables[1]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TongTheoLoaiPhiPhuLuc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			SheetView sView = winFpSpread.Sheets[1];
			Dictionary<int, int> dictionary = new Dictionary<int, int>();
			dictionary.Add(col3a_Tb1_Pl, colTenLoaiPhi_Tb2_Pl);
			dictionary.Add(col3aMa_Tb1_Pl, colMaLoaiPhi_Tb2_Pl);
			dictionary.Add(col4_Tb1_Pl, colTien_Tb2_Pl);
			dictionary.Add(col4Ma_Tb1_Pl, colMaTien_Tb2_Pl);
			dictionary.Add(col5_Tb1_Pl, col5_Tb1_Pl);
			dictionary.Add(col7_Tb1_Pl, col7_Tb1_Pl);
			dictionary.Add(col8_Tb1_Pl, col8_Tb1_Pl);
			Dictionary<string, int> checkExists = new Dictionary<string, int>();
			GeneralFunctions.CopyTableToTable2(sView, hTKKCMCSheetInfo, hTKKCMCSheetInfo.Tables[0], 1, dictionary, deleteDataOld: true, ref checkExists, usingCellType: true);
			GeneralFunctions.fpSpReindexingTableTemplate(sView, hTKKCMCSheetInfo.Tables[1]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TongTheoLoaiTienPhuLuc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			SheetView sView = winFpSpread.Sheets[1];
			Dictionary<int, int> dictionary = new Dictionary<int, int>();
			dictionary.Add(col4_Tb1_Pl, colTien_Tb3_Pl);
			dictionary.Add(col4Ma_Tb1_Pl, colMaTien_Tb3_Pl);
			dictionary.Add(col5_Tb1_Pl, col5_Tb1_Pl);
			dictionary.Add(col7_Tb1_Pl, col7_Tb1_Pl);
			dictionary.Add(col8_Tb1_Pl, col8_Tb1_Pl);
			Dictionary<string, int> checkExists = new Dictionary<string, int>();
			GeneralFunctions.CopyTableToTable2(sView, hTKKCMCSheetInfo, hTKKCMCSheetInfo.Tables[0], 2, dictionary, deleteDataOld: true, ref checkExists, usingCellType: true);
			GeneralFunctions.fpSpReindexingTableTemplate(sView, hTKKCMCSheetInfo.Tables[2]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CapNhatLaiCellTypeTuSViewToKhai(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string value = sView.Cells[dataLocationTemplateRowStart + i - 1, col4_Tb1 - 1].Value.ToString().Trim();
				string text = sView.Cells[dataLocationTemplateRowStart + i - 1, col3Ma_Tb1 - 1].Text;
				HTKKCMCValidate validate = hTKKCMCTableInfo.RowsInfo.Cells[3].Validate;
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col4_Tb1, value, validate, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				ICellType cellType = sView.Cells[dataLocationTemplateRowStart + i - 1, col4_Tb1 - 1].CellType;
				sView.Cells[dataLocationTemplateRowStart + i - 1, col7_Tb1 - 1].CellType = cellType;
				sView.Cells[dataLocationTemplateRowStart + i - 1, col8_Tb1 - 1].CellType = cellType;
				sView.Cells[dataLocationTemplateRowStart + i - 1, col9_Tb1 - 1].CellType = cellType;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[0].Tables[1];
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
			{
				string value = sView.Cells[dataLocationTemplateRowStart2 + i - 1, col4_Tb1 - 1].Value.ToString().Trim();
				object value2 = sView.Cells[dataLocationTemplateRowStart2 + i - 1, col3Ma_Tb1 - 1].Value;
				if (value2 != null)
				{
					HTKKCMCValidate validate = hTKKCMCTableInfo2.RowsInfo.Cells[1].Validate;
					GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart2 + i, col4_Tb1, value, validate, (value2.ToString().Trim() == "" || value2.ToString().Trim() == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
					ICellType cellType = sView.Cells[dataLocationTemplateRowStart2 + i - 1, col4_Tb1 - 1].CellType;
					sView.Cells[dataLocationTemplateRowStart2 + i - 1, col5_Tb1 - 1].CellType = cellType;
					sView.Cells[dataLocationTemplateRowStart2 + i - 1, col7_Tb1 - 1].CellType = cellType;
					sView.Cells[dataLocationTemplateRowStart2 + i - 1, col8_Tb1 - 1].CellType = cellType;
					sView.Cells[dataLocationTemplateRowStart2 + i - 1, col9_Tb1 - 1].CellType = cellType;
				}
			}
			GeneralFunctions.SaveAllCellTypeOfSheet(sView, ToKhaiInfo.SheetInfos[0], 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CapNhatLaiCellTypeTuDataFileToKhai(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int sheetIndex)
	{
		SheetView sheetView = winFpSpread.Sheets[sheetIndex];
		HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[sheetIndex].Tables[0];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
		{
			string value = hTKKCMCTableInfo.RowsValues[i].RowValue[3];
			string text = hTKKCMCTableInfo.RowsValues[i].RowValue[10];
			HTKKCMCValidate validate = hTKKCMCTableInfo.RowsInfo.Cells[3].Validate;
			GeneralFunctions.InsertNumberCell(sheetView, dataLocationTemplateRowStart + i, col4_Tb1, value, validate, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			ICellType cellType = sheetView.Cells[dataLocationTemplateRowStart + i - 1, col4_Tb1 - 1].CellType;
			sheetView.Cells[dataLocationTemplateRowStart + i - 1, col5_Tb1 - 1].CellType = cellType;
			sheetView.Cells[dataLocationTemplateRowStart + i - 1, col7_Tb1 - 1].CellType = cellType;
			sheetView.Cells[dataLocationTemplateRowStart + i - 1, col8_Tb1 - 1].CellType = cellType;
			sheetView.Cells[dataLocationTemplateRowStart + i - 1, col9_Tb1 - 1].CellType = cellType;
		}
		HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[sheetIndex].Tables[1];
		int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
		{
			string value = hTKKCMCTableInfo2.RowsValues[i].RowValue[1];
			string text = hTKKCMCTableInfo2.RowsValues[i].RowValue[4];
			HTKKCMCValidate validate2 = hTKKCMCTableInfo2.RowsInfo.Cells[1].Validate;
			GeneralFunctions.InsertNumberCell(sheetView, dataLocationTemplateRowStart2 + i, col4_Tb1, value, validate2, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			ICellType cellType = sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, col4_Tb1 - 1].CellType;
			sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, col5_Tb1 - 1].CellType = cellType;
			sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, col7_Tb1 - 1].CellType = cellType;
			sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, col8_Tb1 - 1].CellType = cellType;
			sheetView.Cells[dataLocationTemplateRowStart2 + i - 1, col9_Tb1 - 1].CellType = cellType;
		}
	}

	public void CapNhatLaiCellTypeTuSViewPl(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[1].Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string value = sView.Cells[dataLocationTemplateRowStart + i - 1, col5_Tb1_Pl - 1].Value.ToString().Trim();
				string value2 = sView.Cells[dataLocationTemplateRowStart + i - 1, col7_Tb1_Pl - 1].Value.ToString().Trim();
				string value3 = sView.Cells[dataLocationTemplateRowStart + i - 1, col8_Tb1_Pl - 1].Value.ToString().Trim();
				string text = sView.Cells[dataLocationTemplateRowStart + i - 1, col4Ma_Tb1_Pl - 1].Text;
				HTKKCMCValidate validate = hTKKCMCTableInfo.RowsInfo.Cells[4].Validate;
				HTKKCMCValidate validate2 = hTKKCMCTableInfo.RowsInfo.Cells[6].Validate;
				HTKKCMCValidate validate3 = hTKKCMCTableInfo.RowsInfo.Cells[7].Validate;
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col5_Tb1_Pl, value, validate, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col7_Tb1_Pl, value2, validate2, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col8_Tb1_Pl, value3, validate3, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[1].Tables[1];
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
			{
				string value = sView.Cells[dataLocationTemplateRowStart2 + i - 1, col5_Tb1_Pl - 1].Value.ToString().Trim();
				string value2 = sView.Cells[dataLocationTemplateRowStart2 + i - 1, col7_Tb1_Pl - 1].Value.ToString().Trim();
				string value3 = sView.Cells[dataLocationTemplateRowStart2 + i - 1, col8_Tb1_Pl - 1].Value.ToString().Trim();
				object value4 = sView.Cells[dataLocationTemplateRowStart2 + i - 1, col4Ma_Tb1_Pl - 1].Value;
				if (value4 != null)
				{
					HTKKCMCValidate validate = hTKKCMCTableInfo2.RowsInfo.Cells[2].Validate;
					HTKKCMCValidate validate2 = hTKKCMCTableInfo2.RowsInfo.Cells[3].Validate;
					HTKKCMCValidate validate3 = hTKKCMCTableInfo2.RowsInfo.Cells[4].Validate;
					GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart2 + i, col5_Tb1_Pl, value, validate, (value4.ToString().Trim() == "" || value4.ToString().Trim() == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
					GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart2 + i, col7_Tb1_Pl, value2, validate2, (value4.ToString().Trim() == "" || value4.ToString().Trim() == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
					GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart2 + i, col8_Tb1_Pl, value3, validate3, (value4.ToString().Trim() == "" || value4.ToString().Trim() == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				}
			}
			HTKKCMCTableInfo hTKKCMCTableInfo3 = ToKhaiInfo.SheetInfos[1].Tables[2];
			int dataLocationTemplateRowStart3 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo3.RowsValues.Count; i++)
			{
				string value = sView.Cells[dataLocationTemplateRowStart3 + i - 1, col5_Tb1_Pl - 1].Value.ToString().Trim();
				string value2 = sView.Cells[dataLocationTemplateRowStart3 + i - 1, col7_Tb1_Pl - 1].Value.ToString().Trim();
				string value3 = sView.Cells[dataLocationTemplateRowStart3 + i - 1, col8_Tb1_Pl - 1].Value.ToString().Trim();
				object value4 = sView.Cells[dataLocationTemplateRowStart3 + i - 1, col4Ma_Tb1_Pl - 1].Value;
				if (value4 != null)
				{
					HTKKCMCValidate validate = hTKKCMCTableInfo3.RowsInfo.Cells[1].Validate;
					HTKKCMCValidate validate2 = hTKKCMCTableInfo3.RowsInfo.Cells[2].Validate;
					HTKKCMCValidate validate3 = hTKKCMCTableInfo3.RowsInfo.Cells[3].Validate;
					GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart3 + i, col5_Tb1_Pl, value, validate, (value4.ToString().Trim() == "" || value4.ToString().Trim() == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
					GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart3 + i, col7_Tb1_Pl, value2, validate2, (value4.ToString().Trim() == "" || value4.ToString().Trim() == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
					GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart3 + i, col8_Tb1_Pl, value3, validate3, (value4.ToString().Trim() == "" || value4.ToString().Trim() == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
				}
			}
			GeneralFunctions.SaveAllCellTypeOfSheet(sView, ToKhaiInfo.SheetInfos[1], 1);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CapNhatLaiCellTypeTuDataFilePL(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int sheetIndex)
	{
		SheetView sView = winFpSpread.Sheets[sheetIndex];
		HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[sheetIndex].Tables[0];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
		{
			string value = hTKKCMCTableInfo.RowsValues[i].RowValue[4];
			string value2 = hTKKCMCTableInfo.RowsValues[i].RowValue[6];
			string value3 = hTKKCMCTableInfo.RowsValues[i].RowValue[7];
			string text = hTKKCMCTableInfo.RowsValues[i].RowValue[10];
			HTKKCMCValidate validate = hTKKCMCTableInfo.RowsInfo.Cells[4].Validate;
			HTKKCMCValidate validate2 = hTKKCMCTableInfo.RowsInfo.Cells[6].Validate;
			HTKKCMCValidate validate3 = hTKKCMCTableInfo.RowsInfo.Cells[7].Validate;
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col5_Tb1_Pl, value, validate, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col7_Tb1_Pl, value2, validate2, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + i, col8_Tb1_Pl, value3, validate3, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
		}
		HTKKCMCTableInfo hTKKCMCTableInfo2 = ToKhaiInfo.SheetInfos[sheetIndex].Tables[1];
		int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
		{
			string value = hTKKCMCTableInfo2.RowsValues[i].RowValue[2];
			string value2 = hTKKCMCTableInfo2.RowsValues[i].RowValue[3];
			string value3 = hTKKCMCTableInfo2.RowsValues[i].RowValue[4];
			string text = hTKKCMCTableInfo2.RowsValues[i].RowValue[6];
			HTKKCMCValidate validate4 = hTKKCMCTableInfo2.RowsInfo.Cells[2].Validate;
			HTKKCMCValidate validate5 = hTKKCMCTableInfo2.RowsInfo.Cells[3].Validate;
			HTKKCMCValidate validate6 = hTKKCMCTableInfo2.RowsInfo.Cells[4].Validate;
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart2 + i, col5_Tb1_Pl, value, validate4, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart2 + i, col7_Tb1_Pl, value2, validate5, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart2 + i, col8_Tb1_Pl, value3, validate6, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
		}
		HTKKCMCTableInfo hTKKCMCTableInfo3 = ToKhaiInfo.SheetInfos[sheetIndex].Tables[2];
		int dataLocationTemplateRowStart3 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < hTKKCMCTableInfo3.RowsValues.Count; i++)
		{
			string value = hTKKCMCTableInfo3.RowsValues[i].RowValue[1];
			string value2 = hTKKCMCTableInfo3.RowsValues[i].RowValue[2];
			string value3 = hTKKCMCTableInfo3.RowsValues[i].RowValue[3];
			string text = hTKKCMCTableInfo3.RowsValues[i].RowValue[4];
			HTKKCMCValidate validate4 = hTKKCMCTableInfo3.RowsInfo.Cells[1].Validate;
			HTKKCMCValidate validate5 = hTKKCMCTableInfo3.RowsInfo.Cells[2].Validate;
			HTKKCMCValidate validate6 = hTKKCMCTableInfo3.RowsInfo.Cells[3].Validate;
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart3 + i, col5_Tb1_Pl, value, validate4, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart3 + i, col7_Tb1_Pl, value2, validate5, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
			GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart3 + i, col8_Tb1_Pl, value3, validate6, (text == "" || text == "VND") ? "0" : DEFAULT_LOAI_TIEN_KHAC);
		}
	}

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		//IL_002f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0035: Expected O, but got Unknown
		MethodInvoker val = null;
		try
		{
			FpSpread fpSpread = winFpSpread;
			if (val == null)
			{
				val = (MethodInvoker)delegate
				{
					winFpSpread.ActiveSheet.AutoCalculation = false;
					winFpSpread.SuspendLayout();
					HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
					SheetView sheetView = winFpSpread.Sheets[1];
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
					int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
					for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
					{
						sheetView.Cells[dataLocationTemplateRowStart + i - 1, col3b_Tb1_Pl - 1].Value = sheetView.Cells[dataLocationTemplateRowStart + i - 1, col3aMa_Tb1_Pl - 1].Value;
					}
					TongTheoLoaiPhiPhuLuc(winFpSpread, ToKhaiInfo);
					TongTheoLoaiTienPhuLuc(winFpSpread, ToKhaiInfo);
					CapNhatLaiCellTypeTuSViewPl(winFpSpread.Sheets[1], ToKhaiInfo);
					GeneralFunctions.SaveAllCellTypeOfSheet(sheetView, ToKhaiInfo.SheetInfos[1], 1);
					winFpSpread.ActiveSheet.AutoCalculation = true;
				};
			}
			((Control)fpSpread).Invoke((Delegate)(object)val);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
