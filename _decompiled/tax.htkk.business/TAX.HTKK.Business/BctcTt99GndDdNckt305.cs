using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.Model;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BctcTt99GndDdNckt305
{
	private int rowStart_Tk = 52;

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

	private int[] cellTuNgay = GeneralFunctions.ArrayCellId("AB_11");

	private int[] cellDenNgay = GeneralFunctions.ArrayCellId("AC_11");

	private int[] cellKieuKy = GeneralFunctions.ArrayCellId("AD_11");

	private int[] cellKyKKhai = GeneralFunctions.ArrayCellId("AE_11");

	private int[] cellLayDataNamTruoc = GeneralFunctions.ArrayCellId("B_22");

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			sheetView.Cells[cellTuNgay[1], cellTuNgay[0]].Text = PeriodInfo.TuNgay;
			sheetView.Cells[cellDenNgay[1], cellDenNgay[0]].Text = PeriodInfo.DenNgay;
			sheetView.Cells[cellKieuKy[1], cellKieuKy[0]].Text = PeriodInfo.KieuKy;
			if (PeriodInfo.KieuKy == "Q")
			{
				sheetView.Cells[cellKyKKhai[1], cellKyKKhai[0]].Text = "Quý " + PeriodInfo.Quy;
			}
			else
			{
				sheetView.Cells[cellKyKKhai[1], cellKyKKhai[0]].Text = ((PeriodInfo.kyKKhai == "1") ? "6 tháng cuối kỳ" : "6 tháng đầu kỳ");
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
			int num2 = HTKKFile.ColumnLettersToIndex("V");
			int num3 = HTKKFile.ColumnLettersToIndex("W");
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
				getSoKyTruoc_Tk(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, num, !Convert.ToBoolean(valueCustom));
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
			int col = HTKKFile.ColumnLettersToIndex("V");
			int col2 = HTKKFile.ColumnLettersToIndex("W");
			int col3 = HTKKFile.ColumnLettersToIndex("X");
			int col4 = HTKKFile.ColumnLettersToIndex("Y");
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
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, num, !Convert.ToBoolean(valueCustom), rowStart_PL01_1, rowEnd_PL01_1, col, col2);
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, num, !Convert.ToBoolean(valueCustom), rowStart_PL01_1, rowEnd_PL01_1, col3, col4);
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
			int col = HTKKFile.ColumnLettersToIndex("V");
			int col2 = HTKKFile.ColumnLettersToIndex("W");
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
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, num, !Convert.ToBoolean(valueCustom), rowStart_PL01_2, rowEnd_PL01_2, col, col2);
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
			int col = HTKKFile.ColumnLettersToIndex("V");
			int col2 = HTKKFile.ColumnLettersToIndex("W");
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
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, num, !Convert.ToBoolean(valueCustom), rowStart_PL01_3, rowEnd_PL01_3, col, col2);
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

	private void getSoKyTruoc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex, bool resetValue, int rowStart, int rowEnd, int col4, int col5)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[sheetIndex];
			if (resetValue)
			{
				for (int i = 0; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells[i];
					if (hTKKCMCCellInfo.CellTemplate.RowIndex >= rowStart && hTKKCMCCellInfo.CellTemplate.RowIndex <= rowEnd && hTKKCMCCellInfo.CellTemplate.ColIndex == col4)
					{
						sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, col5 - 1].Value = 0;
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
				if (hTKKCMCCellInfo.CellTemplate.RowIndex >= rowStart && hTKKCMCCellInfo.CellTemplate.RowIndex <= rowEnd && hTKKCMCCellInfo.CellTemplate.ColIndex == col4)
				{
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, col5 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo.Value);
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
			string item = text;
			string item2 = text;
			string text2 = "";
			string text3 = "";
			if (HTKKFormater.isNullString(PeriodInfo.TuNgay) || HTKKFormater.isNullString(PeriodInfo.DenNgay))
			{
				list.Add(text);
				return list;
			}
			text2 = PeriodInfo.TuNgay.Split(new char[1] { '/' })[0] + PeriodInfo.TuNgay.Split(new char[1] { '/' })[1] + (ConvertHelpers.ConvertToInt(PeriodInfo.TuNgay.Split(new char[1] { '/' })[2]) - 1);
			text3 = PeriodInfo.DenNgay.Split(new char[1] { '/' })[0] + PeriodInfo.DenNgay.Split(new char[1] { '/' })[1] + (ConvertHelpers.ConvertToInt(PeriodInfo.DenNgay.Split(new char[1] { '/' })[2]) - 1);
			string text4 = HTKKConfig.DATAFILES + "\\" + HeaderInfo.mst;
			DirectoryInfo directoryInfo = new DirectoryInfo(text4);
			if (sheetIndex == 0)
			{
				text = text + "_Q0" + PeriodInfo.Quy + (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) - 1) + "_" + text2 + "_" + text3 + "_" + ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
				string value = ToKhaiInfo.SheetInfos[sheetIndex].DataFile.Replace("_DD", "").Replace("_TL", "") + "_" + (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) - 1) + "_";
				if (directoryInfo.Exists)
				{
					FileInfo[] files = directoryInfo.GetFiles();
					if (files.Length > 0)
					{
						FileInfo[] array = files;
						foreach (FileInfo fileInfo in array)
						{
							if (fileInfo.Name.Contains(value))
							{
								item2 = text4 + "\\" + fileInfo.Name;
								break;
							}
						}
					}
				}
			}
			if (sheetIndex == 1 || sheetIndex == 2 || sheetIndex == 3)
			{
				text = text + "_Q0" + PeriodInfo.Quy + (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) - 1) + "_" + text2 + "_" + text3 + "_" + ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
				int num = ConvertHelpers.ConvertToInt(PeriodInfo.Quy);
				int num2 = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
				int month = GeneralFunctions.GetNgayDauQuy(num, num2, 1, 1).Month;
				Dictionary<string, string> quyNamByThangNam = GeneralFunctions.GetQuyNamByThangNam(month, num2, HeaderInfo.NgayBatDauNamTaiChinh);
				if (num == 1)
				{
					num = 4;
					num2--;
				}
				else
				{
					num--;
				}
				string value2 = "";
				quyNamByThangNam.TryGetValue(num.ToString(), out value2);
				if (num2 == ConvertHelpers.ConvertToInt(value2))
				{
					string value3 = ToKhaiInfo.SheetInfos[sheetIndex].DataFile + "_Q0" + num + value2;
					string value4 = ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
					if (directoryInfo.Exists)
					{
						FileInfo[] files = directoryInfo.GetFiles();
						if (files.Length > 0)
						{
							FileInfo[] array = files;
							foreach (FileInfo fileInfo in array)
							{
								if (fileInfo.Name.Contains(value3) && fileInfo.Name.Contains(value4))
								{
									item = text4 + "\\" + fileInfo.Name;
									break;
								}
							}
						}
					}
				}
			}
			list.Add(text);
			list.Add(item);
			list.Add(item2);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}

	private void getSoKyTruoc_Tk(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex, bool resetValue)
	{
		try
		{
			int num = 50;
			int num2 = 259;
			int num3 = HTKKFile.ColumnLettersToIndex("V");
			int num4 = HTKKFile.ColumnLettersToIndex("W");
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[sheetIndex];
			if (resetValue)
			{
				for (int i = 0; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells[i];
					if (hTKKCMCCellInfo.CellTemplate.RowIndex >= rowStart_Tk && hTKKCMCCellInfo.CellTemplate.RowIndex <= rowEnd_Tk && hTKKCMCCellInfo.CellTemplate.ColIndex == num3)
					{
						sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, num4 - 1].Value = 0;
					}
				}
				return;
			}
			List<string> pathDataFileOldNam = getPathDataFileOldNam(ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, sheetIndex);
			string text = pathDataFileOldNam[0];
			text = text.Replace("_DD", "");
			if (!File.Exists(text))
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = new HTKKCMCSheetInfo();
			hTKKCMCSheetInfo2.GetDetailSheetInfo(text);
			for (int i = 0; i < hTKKCMCSheetInfo2.StaticCells.Count; i++)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo2.StaticCells[i];
				if (hTKKCMCCellInfo2.CellTemplate.RowIndex >= num && hTKKCMCCellInfo2.CellTemplate.RowIndex <= num2 && hTKKCMCCellInfo2.CellTemplate.ColIndex == num3)
				{
					string ctOld = GeneralFunctions.GetNumberInString(hTKKCMCCellInfo2.Path);
					HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cellNew) => cellNew.Path.Contains(ctOld) && cellNew.CellTemplate.RowIndex >= rowStart_Tk && cellNew.CellTemplate.RowIndex <= rowEnd_Tk);
					if (hTKKCMCCellInfo3 != null)
					{
						sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, num4 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo2.Value);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private List<string> getPathDataFileOldNam(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex)
	{
		List<string> list = new List<string>();
		try
		{
			string text = HTKKConfig.DATAFILES + "\\" + TTinChung.mst + "\\" + ToKhaiInfo.SheetInfos[sheetIndex].DataFile;
			string text2 = "";
			string text3 = "";
			string text4 = "";
			text4 = ((!(PeriodInfo.DonViTinhMa == "NT")) ? "VND" : "NT");
			if (HTKKFormater.isNullString(PeriodInfo.TuNgay))
			{
				list.Add(text);
				return list;
			}
			text2 = PeriodInfo.TuNgay.Split(new char[1] { '/' })[0] + PeriodInfo.TuNgay.Split(new char[1] { '/' })[1] + (ConvertHelpers.ConvertToInt(PeriodInfo.TuNgay.Split(new char[1] { '/' })[2]) - 1);
			string s = text2.Substring(0, 2) + "/" + text2.Substring(2, 2) + "/" + text2.Substring(4, 4);
			text3 = DateTime.ParseExact(s, "dd/MM/yyyy", CultureInfo.CurrentCulture).AddMonths(12).AddDays(-1.0)
				.ToString("ddMMyyyy");
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
