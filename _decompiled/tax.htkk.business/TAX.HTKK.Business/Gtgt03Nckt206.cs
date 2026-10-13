using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Gtgt03Nckt206
{
	private string cellId_ct21 = "Q_36";

	private string cellId_ct22 = "L_24";

	private string cellId_ct23 = "J_27";

	private string cellId_ct24 = "L_27";

	private string cellId_ct25 = "L_28";

	private string cellId_ct26 = "J_30";

	private string cellId_ct27 = "J_31";

	private string cellId_ct28 = "L_31";

	private string cellId_ct29 = "J_32";

	private string cellId_ct30 = "J_33";

	private string cellId_ct31 = "L_33";

	private string cellId_ct32 = "J_34";

	private string cellId_ct33 = "L_34";

	private string cellId_ct34 = "J_35";

	private string cellId_ct39 = "L_40";

	private string cellId_ct40a = "L_42";

	private string cellId_ct40b = "L_43";

	private string cellId_ct40 = "L_44";

	private string cellId_ct41 = "L_45";

	private string cellId_ct42 = "L_46";

	private string cellId_ct43 = "L_47";

	private string cellId_K52 = "K_52";

	private string cellId_L52 = "L_52";

	private string cellId_O52 = "O_52";

	private string cellId_Q36 = "Q_36";

	private string cellId_Q41 = "Q_41";

	private string cellId_B18 = "B_18";

	private string cellId_N52 = "N_52";

	private int odrColC = HTKKFile.ColumnLettersToIndex("C");

	private int odrColD = HTKKFile.ColumnLettersToIndex("D");

	private int odrColE = HTKKFile.ColumnLettersToIndex("E");

	private int odrColF = HTKKFile.ColumnLettersToIndex("F");

	private int odrColL = HTKKFile.ColumnLettersToIndex("L");

	private int odrColG = HTKKFile.ColumnLettersToIndex("G");

	private int odrColH = HTKKFile.ColumnLettersToIndex("H");

	private int odrColI = HTKKFile.ColumnLettersToIndex("I");

	private int odrColJ = HTKKFile.ColumnLettersToIndex("J");

	private int odrColK = HTKKFile.ColumnLettersToIndex("K");

	private int odrColM = HTKKFile.ColumnLettersToIndex("M");

	private int odrColN = HTKKFile.ColumnLettersToIndex("N");

	private int[] CellId_ct21 => GeneralFunctions.ArrayCellId(cellId_ct21);

	private int[] CellId_ct22 => GeneralFunctions.ArrayCellId(cellId_ct22);

	private int[] CellId_ct23 => GeneralFunctions.ArrayCellId(cellId_ct23);

	private int[] CellId_ct24 => GeneralFunctions.ArrayCellId(cellId_ct24);

	private int[] CellId_ct25 => GeneralFunctions.ArrayCellId(cellId_ct25);

	private int[] CellId_ct26 => GeneralFunctions.ArrayCellId(cellId_ct26);

	private int[] CellId_ct27 => GeneralFunctions.ArrayCellId(cellId_ct27);

	private int[] CellId_ct28 => GeneralFunctions.ArrayCellId(cellId_ct28);

	private int[] CellId_ct29 => GeneralFunctions.ArrayCellId(cellId_ct29);

	private int[] CellId_ct30 => GeneralFunctions.ArrayCellId(cellId_ct30);

	private int[] CellId_ct31 => GeneralFunctions.ArrayCellId(cellId_ct31);

	private int[] CellId_ct32 => GeneralFunctions.ArrayCellId(cellId_ct32);

	private int[] CellId_ct33 => GeneralFunctions.ArrayCellId(cellId_ct33);

	private int[] CellId_ct34 => GeneralFunctions.ArrayCellId(cellId_ct34);

	private int[] CellId_ct39 => GeneralFunctions.ArrayCellId(cellId_ct39);

	private int[] CellId_ct40a => GeneralFunctions.ArrayCellId(cellId_ct40a);

	private int[] CellId_ct40b => GeneralFunctions.ArrayCellId(cellId_ct40b);

	private int[] CellId_ct40 => GeneralFunctions.ArrayCellId(cellId_ct40);

	private int[] CellId_ct41 => GeneralFunctions.ArrayCellId(cellId_ct41);

	private int[] CellId_ct42 => GeneralFunctions.ArrayCellId(cellId_ct42);

	private int[] CellId_ct43 => GeneralFunctions.ArrayCellId(cellId_ct43);

	private int[] CellId_K52 => GeneralFunctions.ArrayCellId(cellId_K52);

	private int[] CellId_L52 => GeneralFunctions.ArrayCellId(cellId_L52);

	private int[] CellId_O52 => GeneralFunctions.ArrayCellId(cellId_O52);

	private int[] CellId_Q36 => GeneralFunctions.ArrayCellId(cellId_Q36);

	private int[] CellId_Q41 => GeneralFunctions.ArrayCellId(cellId_Q41);

	public double CellValue_Q41 { get; set; }

	private int[] CellId_B18 => GeneralFunctions.ArrayCellId(cellId_B18);

	private int[] CellId_N52 => GeneralFunctions.ArrayCellId(cellId_N52);

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
			if (PeriodInfo.Thang.PadLeft(2, '0') == "01" || PeriodInfo.Quy.PadLeft(2, '0') == "01")
			{
				sView.Cells[row, column2].ResetFormula();
				sView.Cells[row, column].Value = 0;
				sView.Cells[row, column2].Value = 0;
				sView.Cells[row, column2].Locked = true;
			}
			else if (ToKhaiInfo.SheetInfos[0].IsReadDataFile || PeriodInfo.LoaiToKhai == "B")
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
