using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BctcTt99GndTlNckt305
{
	private int rowStart_Tk = 54;

	private int rowEnd_Tk = 77;

	private int[] cellKieuKy = GeneralFunctions.ArrayCellId("AK_8");

	private int[] cellKyKKhai = GeneralFunctions.ArrayCellId("AL_8");

	private HTKKCMCToKhaiInfo tKhaiInfoOld = new HTKKCMCToKhaiInfo();

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
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
			throw ex;
		}
	}

	public void AutoBindData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			SheetView sheetView = winFpSpread.Sheets[1];
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			SheetView sheetView = winFpSpread.Sheets[2];
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			if (ToKhaiInfo.SheetInfos[1] != null)
			{
				Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			}
			if (ToKhaiInfo.SheetInfos[2] != null)
			{
				Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			}
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				int value = item.Value;
				if (value == 2 && winFpSpread.TabStrip[key].BackColor != Color.Pink)
				{
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
				}
				else if (value == 1 && winFpSpread.TabStrip[key].BackColor == Color.White)
				{
					winFpSpread.TabStrip[key].BackColor = Color.Yellow;
				}
				else if (value == 0)
				{
					winFpSpread.TabStrip[key].BackColor = Color.White;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int index = 0;
			SheetView sheetView = winFpSpread.Sheets[index];
			int num = 0;
			int num2 = HTKKFile.ColumnLettersToIndex("AC");
			int num3 = HTKKFile.ColumnLettersToIndex("W");
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[44, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			if (num == 2 && winFpSpread.TabStrip[index].BackColor != Color.Pink)
			{
				winFpSpread.TabStrip[index].BackColor = Color.Pink;
			}
			else if (num == 1 && winFpSpread.TabStrip[index].BackColor == Color.White)
			{
				winFpSpread.TabStrip[index].BackColor = Color.Yellow;
			}
			else if (num == 0)
			{
				winFpSpread.TabStrip[index].BackColor = Color.White;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int index = 1;
			SheetView sheetView = winFpSpread.Sheets[index];
			int[] array = GeneralFunctions.ArrayCellId("B_17");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int index = 2;
			SheetView sheetView = winFpSpread.Sheets[index];
			int[] array = GeneralFunctions.ArrayCellId("C_38");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("D_52");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				getSoKyTruoc_Tk(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 0, !Convert.ToBoolean(valueCustom));
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[1];
			int[] array = GeneralFunctions.ArrayCellId("B_17");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 1, isReadDataFile: false, Convert.ToBoolean(valueCustom), callIsCheck: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[2];
			int[] array = GeneralFunctions.ArrayCellId("C_38");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 2, isReadDataFile: false, Convert.ToBoolean(valueCustom), callIsCheck: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("R") - 1;
			int num2 = HTKKFile.ColumnLettersToIndex("W") - 1;
			int[] array = new int[2];
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			string text = "I###+O###";
			if ((int)iCol == HTKKFile.ColumnLettersToIndex("I"))
			{
				winFpSpread.ActiveSheet.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("K") - 1].Formula = text.Replace("###", iRow.ToString());
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			string text = "R###+AN###";
			if ((int)iCol == HTKKFile.ColumnLettersToIndex("R"))
			{
				winFpSpread.ActiveSheet.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AB") - 1].Formula = text.Replace("###", iRow.ToString());
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
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
			throw ex;
		}
	}

	private void getSoKyTruoc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex, bool isReadDataFile, bool isCheckLayDl, bool callIsCheck)
	{
		try
		{
			List<string> pathDataFileOld = getPathDataFileOld(ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, sheetIndex);
			string text = pathDataFileOld[0];
			string text2 = pathDataFileOld[1];
			string text3 = pathDataFileOld[2];
			string text4 = text3;
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			int num = 0;
			int num2 = 0;
			int num3 = 0;
			int num4 = 0;
			int num5 = 0;
			int num6 = 0;
			int num7 = 0;
			int num8 = 0;
			if (sheetIndex == 0)
			{
				num = HTKKFile.ColumnLettersToIndex("R");
				num2 = HTKKFile.ColumnLettersToIndex("W");
				num3 = HTKKFile.ColumnLettersToIndex("AI");
			}
			if (sheetIndex == 1)
			{
				text4 = text;
				num = HTKKFile.ColumnLettersToIndex("I");
				num2 = HTKKFile.ColumnLettersToIndex("J");
				num3 = HTKKFile.ColumnLettersToIndex("N");
				num4 = HTKKFile.ColumnLettersToIndex("K");
				num5 = HTKKFile.ColumnLettersToIndex("O");
				num6 = HTKKFile.ColumnLettersToIndex("P");
				num7 = HTKKFile.ColumnLettersToIndex("L");
				num8 = HTKKFile.ColumnLettersToIndex("Q");
			}
			if (sheetIndex == 2)
			{
				text4 = text;
				num = HTKKFile.ColumnLettersToIndex("R");
				num2 = HTKKFile.ColumnLettersToIndex("W");
				num3 = HTKKFile.ColumnLettersToIndex("AM");
				num4 = HTKKFile.ColumnLettersToIndex("AB");
				num5 = HTKKFile.ColumnLettersToIndex("AN");
				num6 = HTKKFile.ColumnLettersToIndex("AO");
				num7 = HTKKFile.ColumnLettersToIndex("AG");
				num8 = HTKKFile.ColumnLettersToIndex("AP");
			}
			if (File.Exists(text4) && isCheckLayDl)
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
				hTKKCMCSheetInfo.GetDetailSheetInfo(text4);
				for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num)
					{
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
						if (!isReadDataFile && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex != 118)
						{
							sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num2 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
						}
					}
				}
			}
			else
			{
				for (int i = 1; i < ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Count; i++)
				{
					if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num && !isReadDataFile)
					{
						sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num2 - 1].Value = 0;
					}
				}
			}
			if (File.Exists(text))
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
				hTKKCMCSheetInfo.GetDetailSheetInfo(text);
				for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if ((sheetIndex == 1 || sheetIndex == 2) && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num4)
					{
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num8 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
						if (!isReadDataFile)
						{
							sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num7 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
						}
					}
				}
				return;
			}
			for (int i = 1; i < ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Count; i++)
			{
				if (sheetIndex == 1 && ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num4 && !isReadDataFile)
				{
					sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num7 - 1].Value = 0;
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
			string text4 = "";
			if (HTKKFormater.isNullString(PeriodInfo.TuNgay) || HTKKFormater.isNullString(PeriodInfo.DenNgay))
			{
				list.Add(text);
				return list;
			}
			text3 = PeriodInfo.TuNgay.Split(new char[1] { '/' })[0] + PeriodInfo.TuNgay.Split(new char[1] { '/' })[1] + (ConvertHelpers.ConvertToInt(PeriodInfo.TuNgay.Split(new char[1] { '/' })[2]) - 1);
			text4 = PeriodInfo.DenNgay.Split(new char[1] { '/' })[0] + PeriodInfo.DenNgay.Split(new char[1] { '/' })[1] + (ConvertHelpers.ConvertToInt(PeriodInfo.DenNgay.Split(new char[1] { '/' })[2]) - 1);
			string text5 = HTKKConfig.DATAFILES + "\\" + HeaderInfo.mst;
			DirectoryInfo directoryInfo = new DirectoryInfo(text5);
			if (sheetIndex == 0)
			{
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
								item2 = text5 + "\\" + fileInfo.Name;
								break;
							}
						}
					}
				}
			}
			if (sheetIndex == 1 || sheetIndex == 2)
			{
				text = text + "_Q0" + PeriodInfo.Quy + (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) - 1) + "_" + text3 + "_" + text4 + "_" + ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
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
									item = text5 + "\\" + fileInfo.Name;
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

	private void checkSoNamTruoc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex, bool isCheck)
	{
		try
		{
			int num = 0;
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			List<string> pathDataFileOld = getPathDataFileOld(ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, sheetIndex);
			string path = pathDataFileOld[2];
			int num2 = 0;
			int num3 = 0;
			int num4 = 0;
			int num5 = 0;
			int num6 = 0;
			int num7 = 0;
			int num8 = 0;
			int num9 = 0;
			if (sheetIndex == 0)
			{
				num2 = HTKKFile.ColumnLettersToIndex("R");
				num3 = HTKKFile.ColumnLettersToIndex("W");
				num4 = HTKKFile.ColumnLettersToIndex("AI");
			}
			if (sheetIndex == 1)
			{
				path = pathDataFileOld[0];
				num2 = HTKKFile.ColumnLettersToIndex("I");
				num3 = HTKKFile.ColumnLettersToIndex("J");
				num4 = HTKKFile.ColumnLettersToIndex("N");
				num5 = HTKKFile.ColumnLettersToIndex("K");
				num6 = HTKKFile.ColumnLettersToIndex("O");
				num7 = HTKKFile.ColumnLettersToIndex("P");
				num8 = HTKKFile.ColumnLettersToIndex("L");
				num9 = HTKKFile.ColumnLettersToIndex("Q");
			}
			if (sheetIndex == 2)
			{
				path = pathDataFileOld[0];
				num2 = HTKKFile.ColumnLettersToIndex("R");
				num3 = HTKKFile.ColumnLettersToIndex("W");
				num4 = HTKKFile.ColumnLettersToIndex("AM");
				num5 = HTKKFile.ColumnLettersToIndex("AB");
				num6 = HTKKFile.ColumnLettersToIndex("AN");
				num7 = HTKKFile.ColumnLettersToIndex("AO");
				num8 = HTKKFile.ColumnLettersToIndex("AG");
				num9 = HTKKFile.ColumnLettersToIndex("AP");
			}
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[26, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[27, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[sheetIndex];
			for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
			{
				string valueCustom;
				object value;
				if (File.Exists(path) && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num3 && !sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Locked)
				{
					valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num3, "NumberCellType");
					value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num4 - 1].Value;
					if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
					{
						if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor != Color.Pink)
						{
							num = 1;
							HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1, text);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1);
					}
				}
				if (File.Exists(pathDataFileOld[0]) && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num8 && !sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num8 - 1].Locked)
				{
					valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num8, "NumberCellType");
					value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num9 - 1].Value;
					if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
					{
						if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num8 - 1].BackColor != Color.Pink)
						{
							num = 1;
							HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num8 - 1, text);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num8 - 1);
					}
				}
				if (!File.Exists(pathDataFileOld[1]) || hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != num5 || sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Locked)
				{
					continue;
				}
				valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num5, "NumberCellType");
				value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num7 - 1].Value;
				if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
				{
					if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1].BackColor != Color.Pink)
					{
						num = 1;
						HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1, text2);
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1);
				}
			}
			if (num == 2 && winFpSpread.TabStrip[sheetIndex].BackColor != Color.Pink)
			{
				winFpSpread.TabStrip[sheetIndex].BackColor = Color.Pink;
			}
			else if (num == 1 && winFpSpread.TabStrip[sheetIndex].BackColor == Color.White)
			{
				winFpSpread.TabStrip[sheetIndex].BackColor = Color.Yellow;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void getSoKyTruoc_Tk(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex, bool resetValue)
	{
		try
		{
			int num = 50;
			int num2 = 259;
			int num3 = HTKKFile.ColumnLettersToIndex("V");
			int num4 = HTKKFile.ColumnLettersToIndex("W");
			int num5 = HTKKFile.ColumnLettersToIndex("R");
			int num6 = HTKKFile.ColumnLettersToIndex("W");
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
			text = text.Replace("_TL", "");
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
}
