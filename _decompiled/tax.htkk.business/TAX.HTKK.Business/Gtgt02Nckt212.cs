using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Gtgt02Nckt212
{
	private string cellId_ct21 = "CT_40";

	private string cellId_ct21a = "CT_41";

	private string cellId_ct22 = "BW_43";

	private string cellId_ct23 = "CT_43";

	private string cellId_ct24 = "BW_45";

	private string cellId_ct25 = "CT_45";

	private string cellId_ct26 = "BW_46";

	private string cellId_ct27 = "CT_46";

	private string cellId_ct28 = "CT_47";

	private string cellId_ct28a = "CT_48";

	private string cellId_ct29 = "CT_49";

	private string cellId_ct30a = "CT_51";

	private string cellId_ct30 = "CT_52";

	private string cellId_ct31 = "CT_53";

	private string cellId_ct32 = "CT_54";

	private string cellId_K52 = "K_52";

	private string cellId_L52 = "L_52";

	private string cellId_O52 = "O_52";

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

	private int[] CellId_ct21a => GeneralFunctions.ArrayCellId(cellId_ct21a);

	private int[] CellId_ct22 => GeneralFunctions.ArrayCellId(cellId_ct22);

	private int[] CellId_ct23 => GeneralFunctions.ArrayCellId(cellId_ct23);

	private int[] CellId_ct24 => GeneralFunctions.ArrayCellId(cellId_ct24);

	private int[] CellId_ct25 => GeneralFunctions.ArrayCellId(cellId_ct25);

	private int[] CellId_ct26 => GeneralFunctions.ArrayCellId(cellId_ct26);

	private int[] CellId_ct27 => GeneralFunctions.ArrayCellId(cellId_ct27);

	private int[] CellId_ct28 => GeneralFunctions.ArrayCellId(cellId_ct28);

	private int[] CellId_ct28a => GeneralFunctions.ArrayCellId(cellId_ct28a);

	private int[] CellId_ct29 => GeneralFunctions.ArrayCellId(cellId_ct29);

	private int[] CellId_ct30a => GeneralFunctions.ArrayCellId(cellId_ct30a);

	private int[] CellId_ct30 => GeneralFunctions.ArrayCellId(cellId_ct30);

	private int[] CellId_ct31 => GeneralFunctions.ArrayCellId(cellId_ct31);

	private int[] CellId_ct32 => GeneralFunctions.ArrayCellId(cellId_ct32);

	public double CellValue_CT32 { get; set; }

	private int[] CellId_K52 => GeneralFunctions.ArrayCellId(cellId_K52);

	private int[] CellId_L52 => GeneralFunctions.ArrayCellId(cellId_L52);

	private int[] CellId_O52 => GeneralFunctions.ArrayCellId(cellId_O52);

	private int[] CellId_B18 => GeneralFunctions.ArrayCellId(cellId_B18);

	private int[] CellId_N52 => GeneralFunctions.ArrayCellId(cellId_N52);

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int row = 39;
			int column = HTKKFile.ColumnLettersToIndex("DO") - 1;
			double? preperiod_CT32Value = GetPreperiod_CT32Value(PeriodInfo, ToKhaiInfo, HeaderInfo);
			sView.Cells[row, column].Value = preperiod_CT32Value ?? 0.0;
			if (ToKhaiInfo.SheetInfos[0].IsReadDataFile || PeriodInfo.LoaiToKhai == "B")
			{
				sView.Cells[CellId_ct21[1], CellId_ct21[0]].ResetFormula();
			}
			else
			{
				sView.Cells[CellId_ct21[1], CellId_ct21[0]].Formula = "DO40";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		checkTKChinh(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
		return true;
	}

	private void checkTKChinh(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			double num = Convert.ToDouble(sView.Cells[CellId_ct30[1], CellId_ct30[0]].Value);
			double num2 = Convert.ToDouble(sView.Cells[CellId_ct31[1], CellId_ct31[0]].Value);
			double num3 = Convert.ToDouble(sView.Cells[CellId_ct32[1], CellId_ct32[0]].Value);
			double num4 = Convert.ToDouble(sView.Cells[CellId_ct29[1], CellId_ct29[0]].Value);
			double num5 = Convert.ToDouble(sView.Cells[CellId_ct30a[1], CellId_ct30a[0]].Value);
			double num6 = Convert.ToDouble(sView.Cells[CellId_ct28a[1], CellId_ct28a[0]].Value);
			double num7 = Convert.ToDouble(sView.Cells[CellId_ct28[1], CellId_ct28[0]].Value);
			double num8 = Convert.ToDouble(sView.Cells[CellId_ct21[1], CellId_ct21[0]].Value);
			double num9 = Convert.ToDouble(sView.Cells[CellId_ct21a[1], CellId_ct21a[0]].Value);
			if (num > num4 - num5)
			{
				HTKKFile.SetNodeBackGroundPink(sView, CellId_ct30[1], CellId_ct30[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_25")));
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sView, CellId_ct30[1], CellId_ct30[0]);
				if (num < 300000000.0 && num > 0.0)
				{
					HTKKFile.SetNodeBackGroundYellow(sView, CellId_ct30[1], CellId_ct30[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_39")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sView, CellId_ct30[1], CellId_ct30[0]);
				}
			}
			if (num2 > num4)
			{
				HTKKFile.SetNodeBackGroundPink(sView, CellId_ct31[1], CellId_ct31[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_30")));
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sView, CellId_ct31[1], CellId_ct31[0]);
			}
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

	public void SupportImportXML_1(HTKKCMCToKhaiInfo ToKhaiInfo, XmlNode xmlSheet)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[4];
			if (hTKKCMCTableInfo == null)
			{
				return;
			}
			hTKKCMCTableInfo.RowsInfo.RowDefaultDisplay = "0";
			hTKKCMCTableInfo.RowsValues[0].RowDisplay = "0";
			XmlNodeList val = xmlSheet.SelectNodes(hTKKCMCTableInfo.TablePath.Remove(0, hTKKCMCTableInfo.TablePath.IndexOf('/') + 1));
			for (int i = 0; i < val.Count; i++)
			{
				for (int j = 0; j < hTKKCMCTableInfo.RowsInfo.Cells.Count; j++)
				{
					XmlNode val2 = val[i].SelectSingleNode(hTKKCMCTableInfo.RowsInfo.Cells[j].Path);
					if (val2 != null)
					{
						val2.InnerText = hTKKCMCTableInfo.RowsInfo.Cells[j].DefaultValue;
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

	public int GetDefaultYear()
	{
		return DateTime.Now.Year;
	}

	private double? GetPreperiod_CT32Value(HTKKPeriodInfo periodInfo, HTKKCMCToKhaiInfo toKhaiInfo, HTKKHeaderInfo headerInfo)
	{
		try
		{
			double? result = null;
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
								string valueInXml = GeneralFunctions.GetValueInXml(text + "\\" + fileInfo.Name, cellId_ct32);
								if (string.IsNullOrEmpty(valueInXml))
								{
									return null;
								}
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
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
