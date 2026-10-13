using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BctcTt58Nckt305
{
	private HTKKCMCToKhaiInfo tKhaiInfoOld = new HTKKCMCToKhaiInfo();

	public bool ValidatePeriod_ALL(HTKKPeriodInfo PeriodInfo)
	{
		bool result = true;
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	public bool ValidatePeriod_axfpsPhuLuc(FpSpread winFpSpread)
	{
		bool result = true;
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
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

	public void AutoBindData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
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

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("D_17");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 0, isReadDataFile: false, Convert.ToBoolean(valueCustom), callIsCheck: true);
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
			int[] array = GeneralFunctions.ArrayCellId("D_17");
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

	public static void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
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

	public static void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
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

	public void isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			if (ToKhaiInfo.SheetInfos[0] != null)
			{
				Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			}
			if (ToKhaiInfo.SheetInfos[1] != null)
			{
				Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			}
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				int value = item.Value;
				if ((value != 2 || !(winFpSpread.TabStrip[key].BackColor == Color.Pink)) && (value != 1 || !(winFpSpread.TabStrip[key].BackColor == Color.Yellow)))
				{
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		//IL_0084: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			int index = 0;
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = 0;
			if (!winFpSpread.Sheets[1].Visible)
			{
				winFpSpread.ResumeLayout();
				Process[] processesByName = Process.GetProcessesByName("ProgressIndicatorTest");
				foreach (Process process in processesByName)
				{
					process.Kill();
				}
				frmMessage frmMessage = new frmMessage("0484", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				num = 2;
			}
			if ((num != 2 || !(winFpSpread.TabStrip[index].BackColor == Color.Pink)) && (num != 1 || !(winFpSpread.TabStrip[index].BackColor == Color.Yellow)))
			{
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
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			int num = 0;
			int num2 = 0;
			int num3 = 0;
			switch (sheetIndex)
			{
			case 0:
				num = HTKKFile.ColumnLettersToIndex("R");
				num2 = HTKKFile.ColumnLettersToIndex("W");
				num3 = HTKKFile.ColumnLettersToIndex("AC");
				break;
			case 1:
				num = HTKKFile.ColumnLettersToIndex("J");
				num2 = HTKKFile.ColumnLettersToIndex("K");
				num3 = HTKKFile.ColumnLettersToIndex("M");
				break;
			}
			if (File.Exists(text) && isCheckLayDl)
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
				hTKKCMCSheetInfo.GetDetailSheetInfo(text);
				for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num)
					{
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
						if (!isReadDataFile)
						{
							sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num2 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
						}
					}
				}
				return;
			}
			for (int i = 1; i < ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Count; i++)
			{
				if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num && !isReadDataFile)
				{
					sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num2 - 1].Value = 0;
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
			if (HTKKFormater.isNullString(PeriodInfo.TuNgay) || HTKKFormater.isNullString(PeriodInfo.DenNgay))
			{
				list.Add(text);
				return list;
			}
			text2 = PeriodInfo.TuNgay.Split(new char[1] { '/' })[0] + PeriodInfo.TuNgay.Split(new char[1] { '/' })[1] + (ConvertHelpers.ConvertToInt(PeriodInfo.TuNgay.Split(new char[1] { '/' })[2]) - 1);
			text3 = PeriodInfo.DenNgay.Split(new char[1] { '/' })[0] + PeriodInfo.DenNgay.Split(new char[1] { '/' })[1] + (ConvertHelpers.ConvertToInt(PeriodInfo.DenNgay.Split(new char[1] { '/' })[2]) - 1);
			text = text + "_" + (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) - 1) + "_" + text2 + "_" + text3 + "_" + ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
			list.Add(text);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}
}
