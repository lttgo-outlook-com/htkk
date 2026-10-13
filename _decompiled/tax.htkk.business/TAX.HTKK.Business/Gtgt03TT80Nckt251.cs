using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Gtgt03TT80Nckt251
{
	private string cellId_ct21 = "Q_36";

	private int[] CellId_ct21 => GeneralFunctions.ArrayCellId(cellId_ct21);

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int row = 35;
			int column = HTKKFile.ColumnLettersToIndex("U") - 1;
			int column2 = HTKKFile.ColumnLettersToIndex("Q") - 1;
			double preperiod_CT26Value = GetPreperiod_CT26Value(PeriodInfo, ToKhaiInfo, HeaderInfo);
			if (preperiod_CT26Value < 0.0)
			{
				sView.Cells[row, column].Value = preperiod_CT26Value * -1.0;
			}
			else
			{
				sView.Cells[row, column].Value = 0;
			}
			if (ToKhaiInfo.SheetInfos[0].IsReadDataFile || PeriodInfo.LoaiToKhai == "B")
			{
				sView.Cells[row, column2].ResetFormula();
			}
			else
			{
				sView.Cells[row, column2].Formula = "U36";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning, out object outPut)
	{
		outPut = 1;
		checkTKChinh(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
		return true;
	}

	private void checkTKChinh(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			switch (value)
			{
			case 2:
				sView.FpSpread.TabStrip[0].BackColor = Color.Pink;
				break;
			case 1:
				sView.FpSpread.TabStrip[0].BackColor = Color.Yellow;
				break;
			default:
				sView.FpSpread.TabStrip[0].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public int GetDefaultYear()
	{
		return DateTime.Now.Year;
	}

	private double GetPreperiod_CT26Value(HTKKPeriodInfo periodInfo, HTKKCMCToKhaiInfo toKhaiInfo, HTKKHeaderInfo headerInfo)
	{
		try
		{
			double result = 0.0;
			List<string> preperiodFileName_ThangQuy = GeneralFunctions.GetPreperiodFileName_ThangQuy(periodInfo, toKhaiInfo);
			string text = HTKKConfig.DATAFILES + "\\" + headerInfo.mst;
			DirectoryInfo directoryInfo = new DirectoryInfo(text);
			bool flag = periodInfo.LoaiToKhai.Equals("C");
			if (directoryInfo.Exists && flag && preperiodFileName_ThangQuy.Count > 0)
			{
				FileInfo[] files = directoryInfo.GetFiles();
				if (files.Length > 0)
				{
					bool flag2 = false;
					foreach (string item in preperiodFileName_ThangQuy)
					{
						FileInfo[] array = files;
						foreach (FileInfo fileInfo in array)
						{
							if (fileInfo.Name.Contains(item))
							{
								string valueInXml = GeneralFunctions.GetValueInXml(text + "\\" + fileInfo.Name, "Q_41");
								result = Convert.ToDouble(valueInXml);
								flag2 = true;
								break;
							}
						}
						if (flag2)
						{
							break;
						}
					}
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}
}
