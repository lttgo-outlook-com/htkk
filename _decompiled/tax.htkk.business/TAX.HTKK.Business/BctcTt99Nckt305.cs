using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.Model;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BctcTt99Nckt305
{
	private int rowStart_Tk = 50;

	private int rowEnd_HDLL_Tk = 259;

	private int rowStart_HDKLL_Tk = 260;

	private int rowEnd_Tk = 372;

	private int rowStart_PL01_1 = 52;

	private int rowEnd_PL01_1 = 84;

	private int rowStart_PL01_2 = 52;

	private int rowEnd_PL01_2 = 108;

	private int rowStart_PL01_3 = 53;

	private int rowEnd_PL01_3 = 128;

	private int colThuyetMinh = HTKKFile.ColumnLettersToIndex("U");

	private int colCuoiNam_TK = HTKKFile.ColumnLettersToIndex("V");

	private int colDauNam_TK = HTKKFile.ColumnLettersToIndex("W");

	private int colNamNay_Pl = HTKKFile.ColumnLettersToIndex("V");

	private int colNamTruoc_Pl = HTKKFile.ColumnLettersToIndex("W");

	private int[] cellTuNgay = GeneralFunctions.ArrayCellId("AB_9");

	private int[] cellDenNgay = GeneralFunctions.ArrayCellId("AC_9");

	private int[] cellLienTucKhongLienTuc = GeneralFunctions.ArrayCellId("AD_9");

	private int[] cellHDLienTuc = GeneralFunctions.ArrayCellId("B_26");

	private int[] cellHDKhongLienTuc = GeneralFunctions.ArrayCellId("B_28");

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			if (PeriodInfo.DoanhNghiepHoatDongLienTuc == "LT")
			{
				xmlNode.SelectSingleNode("HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/ngayKy").InnerText = xmlNode.SelectSingleNode("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/BCTHTC_HoatDongLienTuc/ngayLap").InnerText;
				xmlNode.SelectSingleNode("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/hoatDongLienTuc").InnerText = "true";
				xmlNode.SelectSingleNode("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/hoatDongKhongLienTuc").InnerText = "false";
			}
			else
			{
				xmlNode.SelectSingleNode("HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/ngayKy").InnerText = xmlNode.SelectSingleNode("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/BCTHTC_HoatDongLienTuc/ngayLap").InnerText;
				xmlNode.SelectSingleNode("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/hoatDongLienTuc").InnerText = "false";
				xmlNode.SelectSingleNode("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/hoatDongKhongLienTuc").InnerText = "true";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ExportXML_1(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ExportXML_2(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ExportXML_3(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SetFPSpreadExp(FpSpread winFpExp, List<HTKKCMCSheetInfo> sheetInfos, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKHeaderInfo HTKKHeaderInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			sheetView.Cells[cellTuNgay[1], cellTuNgay[0]].Text = PeriodInfo.TuNgay;
			sheetView.Cells[cellDenNgay[1], cellDenNgay[0]].Text = PeriodInfo.DenNgay;
			sheetView.Cells[cellLienTucKhongLienTuc[1], cellLienTucKhongLienTuc[0]].Text = ((PeriodInfo.DoanhNghiepHoatDongLienTuc == "LT") ? "1" : "0");
			sheetView.Cells[cellHDLienTuc[1], cellHDLienTuc[0]].Text = ((PeriodInfo.DoanhNghiepHoatDongLienTuc == "LT") ? "true" : "false");
			sheetView.Cells[cellHDKhongLienTuc[1], cellHDKhongLienTuc[0]].Text = ((PeriodInfo.DoanhNghiepHoatDongLienTuc == "LT") ? "false" : "true");
			if (PeriodInfo.DoanhNghiepHoatDongLienTuc == "LT")
			{
				foreach (HTKKCMCCellInfo staticCell in hTKKCMCSheetInfo.StaticCells)
				{
					if (staticCell.CellTemplate.RowIndex >= rowStart_Tk && staticCell.CellTemplate.RowIndex <= rowEnd_HDLL_Tk)
					{
						sheetView.Rows[staticCell.CellTemplate.RowIndex - 1].Visible = true;
						sheetView.Rows[rowStart_Tk - 1].Visible = true;
						sheetView.Rows[rowStart_Tk].Visible = true;
					}
					if (staticCell.CellTemplate.RowIndex >= rowStart_HDKLL_Tk && staticCell.CellTemplate.RowIndex <= rowEnd_Tk)
					{
						sheetView.Rows[staticCell.CellTemplate.RowIndex - 1].Visible = false;
						sheetView.Rows[rowStart_HDKLL_Tk - 1].Visible = false;
						sheetView.Rows[rowStart_HDKLL_Tk].Visible = false;
					}
				}
				return;
			}
			foreach (HTKKCMCCellInfo staticCell2 in hTKKCMCSheetInfo.StaticCells)
			{
				if (staticCell2.CellTemplate.RowIndex >= rowStart_Tk && staticCell2.CellTemplate.RowIndex <= rowEnd_HDLL_Tk)
				{
					sheetView.Rows[staticCell2.CellTemplate.RowIndex - 1].Visible = false;
					sheetView.Rows[rowStart_Tk - 1].Visible = false;
					sheetView.Rows[rowStart_Tk].Visible = false;
				}
				if (staticCell2.CellTemplate.RowIndex >= rowStart_HDKLL_Tk && staticCell2.CellTemplate.RowIndex <= rowEnd_Tk)
				{
					sheetView.Rows[staticCell2.CellTemplate.RowIndex - 1].Visible = true;
					sheetView.Rows[rowStart_HDKLL_Tk - 1].Visible = true;
					sheetView.Rows[rowStart_HDKLL_Tk].Visible = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
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
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_3(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			return true;
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
			int index = 0;
			SheetView sheetView = winFpSpread.Sheets[index];
			int value = 0;
			levelWarning.TryGetValue(0, out value);
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
			int num = 1;
			int value = 0;
			levelWarning.TryGetValue(num, out value);
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[num].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[num].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[num].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Validate_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int num = 2;
			int value = 0;
			levelWarning.TryGetValue(num, out value);
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[num].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[num].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[num].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Validate_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int num = 3;
			int value = 0;
			levelWarning.TryGetValue(num, out value);
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[num].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[num].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[num].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			int num = 0;
			SheetView sheetView = winFpSpread.Sheets[num];
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[num];
			int[] array = GeneralFunctions.ArrayCellId("B_24");
			if (idxRow == array[1] && idxCol == array[0])
			{
				bool flag = Convert.ToBoolean(sheetView.Cells[array[1], array[0]].Value);
				lockUnlockCotThuyetMinh(sheetView, sInfo, rowStart_Tk, rowEnd_Tk, colThuyetMinh, !flag, !flag);
			}
			int[] array2 = GeneralFunctions.ArrayCellId("B_22");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (idxCol == array2[0] && idxRow == array2[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, num, !Convert.ToBoolean(valueCustom), rowStart_Tk, rowEnd_Tk);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			int num = 1;
			SheetView sheetView = winFpSpread.Sheets[num];
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[num];
			int[] array = GeneralFunctions.ArrayCellId("B_24");
			if (idxRow == array[1] && idxCol == array[0])
			{
				bool flag = Convert.ToBoolean(sheetView.Cells[array[1], array[0]].Value);
				lockUnlockCotThuyetMinh(sheetView, sInfo, rowStart_PL01_1, rowEnd_PL01_1, colThuyetMinh, !flag, !flag);
			}
			int[] array2 = GeneralFunctions.ArrayCellId("B_22");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (idxCol == array2[0] && idxRow == array2[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, num, !Convert.ToBoolean(valueCustom), rowStart_PL01_1, rowEnd_PL01_1);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			int num = 2;
			SheetView sheetView = winFpSpread.Sheets[num];
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[num];
			int[] array = GeneralFunctions.ArrayCellId("B_24");
			if (idxRow == array[1] && idxCol == array[0])
			{
				bool flag = Convert.ToBoolean(sheetView.Cells[array[1], array[0]].Value);
				lockUnlockCotThuyetMinh(sheetView, sInfo, rowStart_PL01_1, rowEnd_PL01_1, colThuyetMinh, !flag, !flag);
			}
			int[] array2 = GeneralFunctions.ArrayCellId("B_22");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (idxCol == array2[0] && idxRow == array2[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, num, !Convert.ToBoolean(valueCustom), rowStart_PL01_2, rowEnd_PL01_2);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			int num = 3;
			SheetView sheetView = winFpSpread.Sheets[num];
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[num];
			int[] array = GeneralFunctions.ArrayCellId("B_24");
			if (idxRow == array[1] && idxCol == array[0])
			{
				bool flag = Convert.ToBoolean(sheetView.Cells[array[1], array[0]].Value);
				lockUnlockCotThuyetMinh(sheetView, sInfo, rowStart_PL01_1, rowEnd_PL01_1, colThuyetMinh, !flag, !flag);
			}
			int[] array2 = GeneralFunctions.ArrayCellId("B_22");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (idxCol == array2[0] && idxRow == array2[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, num, !Convert.ToBoolean(valueCustom), rowStart_PL01_3, rowEnd_PL01_3);
			}
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
			int num = 229;
			int num2 = 230;
			int num3 = 234;
			if ((iRow == num2 || iRow == num3) && (iCol == colCuoiNam_TK || iCol == colDauNam_TK))
			{
				sheetView.Cells[num - 1, (int)iCol - 1].Value = Convert.ToDouble(sheetView.Cells[num2 - 1, (int)iCol - 1].Value) + Convert.ToDouble(sheetView.Cells[num3 - 1, (int)iCol - 1].Value);
			}
			int num4 = 354;
			int num5 = 355;
			int num6 = 356;
			if ((iRow == num5 || iRow == num6) && (iCol == colCuoiNam_TK || iCol == colDauNam_TK))
			{
				sheetView.Cells[num4 - 1, (int)iCol - 1].Value = Convert.ToDouble(sheetView.Cells[num5 - 1, (int)iCol - 1].Value) + Convert.ToDouble(sheetView.Cells[num6 - 1, (int)iCol - 1].Value);
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellEditchange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void getSoKyTruoc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex, bool resetValue, int rowStart, int rowEnd)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("V");
			int num2 = HTKKFile.ColumnLettersToIndex("W");
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[sheetIndex];
			if (resetValue)
			{
				for (int i = 0; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells[i];
					if (hTKKCMCCellInfo.CellTemplate.RowIndex >= rowStart && hTKKCMCCellInfo.CellTemplate.RowIndex <= rowEnd && hTKKCMCCellInfo.CellTemplate.ColIndex == num)
					{
						sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, num2 - 1].Value = 0;
					}
				}
				return;
			}
			List<string> pathDataFileOld = getPathDataFileOld(ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, sheetIndex);
			string text = pathDataFileOld[0];
			if (!File.Exists(text))
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = new HTKKCMCSheetInfo();
			hTKKCMCSheetInfo2.GetDetailSheetInfo(text);
			for (int i = 0; i < hTKKCMCSheetInfo2.StaticCells.Count; i++)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo2.StaticCells[i];
				if (hTKKCMCCellInfo.CellTemplate.RowIndex >= rowStart && hTKKCMCCellInfo.CellTemplate.RowIndex <= rowEnd && hTKKCMCCellInfo.CellTemplate.ColIndex == num)
				{
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, num2 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo.Value);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private List<string> getPathDataFileOld(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex)
	{
		List<string> list = new List<string>();
		try
		{
			string text = HTKKConfig.DATAFILES + "\\" + TTinChung.mst + "\\" + ToKhaiInfo.SheetInfos[sheetIndex].DataFile;
			string text2 = "";
			string text3 = "";
			string text4 = "";
			text4 = ((!(PeriodInfo.DonViTinhMa == "NT")) ? "VND" : "NT");
			if (HTKKFormater.isNullString(PeriodInfo.TuNgay) || HTKKFormater.isNullString(PeriodInfo.DenNgay))
			{
				list.Add(text);
				return list;
			}
			text2 = PeriodInfo.TuNgay.Split(new char[1] { '/' })[0] + PeriodInfo.TuNgay.Split(new char[1] { '/' })[1] + (ConvertHelpers.ConvertToInt(PeriodInfo.TuNgay.Split(new char[1] { '/' })[2]) - 1);
			text3 = PeriodInfo.DenNgay.Split(new char[1] { '/' })[0] + PeriodInfo.DenNgay.Split(new char[1] { '/' })[1] + (ConvertHelpers.ConvertToInt(PeriodInfo.DenNgay.Split(new char[1] { '/' })[2]) - 1);
			text = text + "_" + (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) - 1) + "_" + text2 + "_" + text3 + "_" + text4 + "_" + ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
			list.Add(text);
		}
		catch (Exception ex)
		{
			list.Add("");
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void lockUnlockCotThuyetMinh(SheetView sView, HTKKCMCSheetInfo sInfo, int startRow, int endRow, int colLock, bool isLock, bool deleteData)
	{
		foreach (HTKKCMCCellInfo staticCell in sInfo.StaticCells)
		{
			if (staticCell.CellTemplate.RowIndex >= startRow && staticCell.CellTemplate.RowIndex <= endRow)
			{
				ISheetStyleModel style = sView.Models.Style;
				StyleInfo compositeInfo = style.GetCompositeInfo(staticCell.CellTemplate.RowIndex - 1, staticCell.CellTemplate.ColIndex - 1, -1, null);
				if (compositeInfo.Font == null || !compositeInfo.Font.Bold)
				{
					GeneralFunctions.LockUnLockCol(sView, staticCell.CellTemplate.RowIndex, colLock, isLock, deleteData);
				}
			}
		}
	}
}
