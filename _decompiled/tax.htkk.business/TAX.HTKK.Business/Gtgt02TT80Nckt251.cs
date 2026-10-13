using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Gtgt02TT80Nckt251
{
	private string cellId_ct21 = "CT_70";

	private string cellId_ct21a = "CT_71";

	private string cellId_ct22 = "BW_73";

	private string cellId_ct23 = "CT_73";

	private string cellId_ct24 = "BW_76";

	private string cellId_ct25 = "CT_76";

	private string cellId_ct26 = "BW_77";

	private string cellId_ct27 = "CT_77";

	private string cellId_ct28 = "CT_78";

	private string cellId_ct28a = "CT_79";

	private string cellId_ct29 = "CT_81";

	private string cellId_ct30 = "CT_82";

	private string cellId_ct31 = "CT_83";

	private string cellId_ct32 = "CT_84";

	private string cellId_MaHSDADT = "AS_9";

	private string cellId_TenDADT = "W_20";

	private string cellId_DiaChiDADT = "AG_28";

	private string cellId_TinhDADT = "W_24";

	private string cellId_MaTinhDADT = "DO_24";

	private string cellId_HuyenDADT = "BP_24";

	private string cellId_MaHuyenDADT = "DP_24";

	private string cellId_XaDADT = "W_26";

	private string cellId_MaXaDADT = "DO_26";

	private string cellId_SoDADT = "AG_31";

	private string cellId_NgayDADT = "BL_31";

	private string cellId_CoQuanBHDADT = "CB_31";

	private string cellId_MSTDADT = "AG_35";

	private string cellId_TenChuDADT = "AG_33";

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

	private int[] CellId_ct30 => GeneralFunctions.ArrayCellId(cellId_ct30);

	private int[] CellId_ct31 => GeneralFunctions.ArrayCellId(cellId_ct31);

	private int[] CellId_ct32 => GeneralFunctions.ArrayCellId(cellId_ct32);

	private int[] CellId_MaHSDADT => GeneralFunctions.ArrayCellId(cellId_MaHSDADT);

	private int[] CellId_TenDADT => GeneralFunctions.ArrayCellId(cellId_TenDADT);

	private int[] CellId_DiaChiDADT => GeneralFunctions.ArrayCellId(cellId_DiaChiDADT);

	private int[] CellId_TinhDADT => GeneralFunctions.ArrayCellId(cellId_TinhDADT);

	private int[] CellId_MaTinhDADT => GeneralFunctions.ArrayCellId(cellId_MaTinhDADT);

	private int[] CellId_HuyenDADT => GeneralFunctions.ArrayCellId(cellId_HuyenDADT);

	private int[] CellId_MaHuyenDADT => GeneralFunctions.ArrayCellId(cellId_MaHuyenDADT);

	private int[] CellId_XaDADT => GeneralFunctions.ArrayCellId(cellId_XaDADT);

	private int[] CellId_MaXaDADT => GeneralFunctions.ArrayCellId(cellId_MaXaDADT);

	private int[] CellId_SoDADT => GeneralFunctions.ArrayCellId(cellId_SoDADT);

	private int[] CellId_NgayDADT => GeneralFunctions.ArrayCellId(cellId_NgayDADT);

	private int[] CellId_CoQuanBHDADT => GeneralFunctions.ArrayCellId(cellId_CoQuanBHDADT);

	private int[] CellId_MSTDADT => GeneralFunctions.ArrayCellId(cellId_MSTDADT);

	private int[] CellId_TenChuDADT => GeneralFunctions.ArrayCellId(cellId_TenChuDADT);

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int row = 69;
			int column = HTKKFile.ColumnLettersToIndex("DO") - 1;
			double? preperiod_CT32Value = GetPreperiod_CT32Value(PeriodInfo, ToKhaiInfo, HeaderInfo);
			sView.Cells[row, column].Value = preperiod_CT32Value ?? 0.0;
			if (ToKhaiInfo.SheetInfos[0].IsReadDataFile || PeriodInfo.LoaiToKhai == "B")
			{
				sView.Cells[CellId_ct21[1], CellId_ct21[0]].ResetFormula();
			}
			else
			{
				sView.Cells[CellId_ct21[1], CellId_ct21[0]].Formula = "DO70";
			}
			sView.Cells[CellId_MaHSDADT[1], CellId_MaHSDADT[0]].Text = PeriodInfo.MaHSDuAnDT;
			TextCellType textCellType = new TextCellType();
			textCellType.WordWrap = true;
			textCellType.Multiline = true;
			sView.Cells[CellId_TinhDADT[1], CellId_TinhDADT[0]].CellType = textCellType;
			sView.Cells[CellId_HuyenDADT[1], CellId_HuyenDADT[0]].CellType = textCellType;
			sView.Cells[CellId_XaDADT[1], CellId_XaDADT[0]].CellType = textCellType;
			string[] tTDuAnDT = getTTDuAnDT(PeriodInfo.MaHSDuAnDT, HeaderInfo);
			if (tTDuAnDT != null && tTDuAnDT.Length > 0)
			{
				sView.Cells[CellId_TenDADT[1], CellId_TenDADT[0]].Text = tTDuAnDT[3];
				sView.Cells[CellId_TinhDADT[1], CellId_TinhDADT[0]].Text = tTDuAnDT[5];
				sView.Cells[CellId_MaTinhDADT[1], CellId_MaTinhDADT[0]].Text = tTDuAnDT[6];
				sView.Cells[CellId_HuyenDADT[1], CellId_HuyenDADT[0]].Text = "";
				sView.Cells[CellId_MaHuyenDADT[1], CellId_MaHuyenDADT[0]].Text = tTDuAnDT[8];
				sView.Cells[CellId_XaDADT[1], CellId_XaDADT[0]].Text = tTDuAnDT[9];
				sView.Cells[CellId_MaXaDADT[1], CellId_MaXaDADT[0]].Text = tTDuAnDT[10];
				sView.Cells[CellId_DiaChiDADT[1], CellId_DiaChiDADT[0]].Text = tTDuAnDT[4];
				sView.Cells[CellId_TenChuDADT[1], CellId_TenChuDADT[0]].Text = tTDuAnDT[11];
				sView.Cells[CellId_MSTDADT[1], CellId_MSTDADT[0]].Text = tTDuAnDT[12];
				sView.Cells[CellId_SoDADT[1], CellId_SoDADT[0]].Text = tTDuAnDT[19];
				sView.Cells[CellId_NgayDADT[1], CellId_NgayDADT[0]].Text = HTKKFormater.FormatStringToDate(tTDuAnDT[20]);
				sView.Cells[CellId_CoQuanBHDADT[1], CellId_CoQuanBHDADT[0]].Text = tTDuAnDT[21];
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
			string text = sView.Cells[CellId_TenDADT[1], CellId_TenDADT[0]].Text;
			string text2 = sView.Cells[CellId_TinhDADT[1], CellId_TinhDADT[0]].Text;
			string text3 = sView.Cells[CellId_HuyenDADT[1], CellId_HuyenDADT[0]].Text;
			string text4 = sView.Cells[CellId_XaDADT[1], CellId_XaDADT[0]].Text;
			string text5 = sView.Cells[CellId_DiaChiDADT[1], CellId_DiaChiDADT[0]].Text;
			string text6 = sView.Cells[CellId_SoDADT[1], CellId_SoDADT[0]].Text;
			string text7 = sView.Cells[CellId_NgayDADT[1], CellId_NgayDADT[0]].Text;
			string text8 = sView.Cells[CellId_CoQuanBHDADT[1], CellId_CoQuanBHDADT[0]].Text;
			string text9 = sView.Cells[CellId_MSTDADT[1], CellId_MSTDADT[0]].Text;
			string text10 = sView.Cells[CellId_TenChuDADT[1], CellId_TenChuDADT[0]].Text;
			string mst = HeaderInfo.mst;
			HTKKFile.RemoveValidateBackGround1(sView, CellId_TenDADT[1], CellId_TenDADT[0]);
			if (string.IsNullOrEmpty(text))
			{
				HTKKFile.SetNodeBackGroundPink(sView, CellId_TenDADT[1], CellId_TenDADT[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_37")));
				value = 2;
			}
			HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(TTinChung.mst);
			if (GeneralFunctions.CheckTinhTheoCQTDbhc(headerInfo.maCQTCapCuc, headerInfo.maCQTNoiNop, sView, PeriodInfo, CellId_TinhDADT, CellId_MaTinhDADT, CellId_HuyenDADT, CellId_MaHuyenDADT, CellId_XaDADT, CellId_MaXaDADT, null, null, null, PhuThuocTruSoChinh: true, isBackGroudColor: true))
			{
				value = 2;
			}
			HTKKFile.RemoveValidateBackGround1(sView, CellId_DiaChiDADT[1], CellId_DiaChiDADT[0]);
			if (string.IsNullOrEmpty(text5))
			{
				HTKKFile.SetNodeBackGroundPink(sView, CellId_DiaChiDADT[1], CellId_DiaChiDADT[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_37")));
				value = 2;
			}
			HTKKFile.RemoveValidateBackGround1(sView, CellId_TinhDADT[1], CellId_TinhDADT[0]);
			HTKKFile.RemoveValidateBackGround1(sView, CellId_HuyenDADT[1], CellId_HuyenDADT[0]);
			HTKKFile.RemoveValidateBackGround1(sView, CellId_XaDADT[1], CellId_XaDADT[0]);
			if (string.IsNullOrEmpty(text2))
			{
				HTKKFile.SetNodeBackGroundPink(sView, CellId_TinhDADT[1], CellId_TinhDADT[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_37")));
				value = 2;
			}
			else
			{
				string text11 = sView.Cells[CellId_MaTinhDADT[1], CellId_MaTinhDADT[0]].Text;
				if (!(PeriodInfo.maCQTCapCuc == "82500"))
				{
				}
			}
			if (!string.IsNullOrEmpty(text3))
			{
				string text11 = sView.Cells[CellId_MaTinhDADT[1], CellId_MaTinhDADT[0]].Text;
				string text12 = sView.Cells[CellId_MaHuyenDADT[1], CellId_MaHuyenDADT[0]].Text;
				if (!(PeriodInfo.maCQTCapCuc == "82500"))
				{
				}
			}
			if (string.IsNullOrEmpty(text4))
			{
				HTKKFile.SetNodeBackGroundPink(sView, CellId_XaDADT[1], CellId_XaDADT[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_37")));
				value = 2;
			}
			HTKKFile.RemoveValidateBackGround1(sView, CellId_SoDADT[1], CellId_SoDADT[0]);
			if (string.IsNullOrEmpty(text6))
			{
				HTKKFile.SetNodeBackGroundPink(sView, CellId_SoDADT[1], CellId_SoDADT[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_37")));
				value = 2;
			}
			HTKKFile.RemoveValidateBackGround1(sView, CellId_NgayDADT[1], CellId_NgayDADT[0]);
			DateTime dtDate;
			if (string.IsNullOrEmpty(text7))
			{
				HTKKFile.SetNodeBackGroundPink(sView, CellId_NgayDADT[1], CellId_NgayDADT[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_37")));
				value = 2;
			}
			else if (!HTKKFormater.IsDateTime(text7, "dd/MM/yyyy", out dtDate))
			{
				HTKKFile.SetNodeBackGroundPink(sView, CellId_NgayDADT[1], CellId_NgayDADT[0], HTKKFile.GetMessage("0071"));
				value = 2;
			}
			else if (dtDate.CompareTo(DateTime.Today) > 0)
			{
				HTKKFile.SetNodeBackGroundPink(sView, CellId_NgayDADT[1], CellId_NgayDADT[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_45")));
				value = 2;
			}
			HTKKFile.RemoveValidateBackGround1(sView, CellId_CoQuanBHDADT[1], CellId_CoQuanBHDADT[0]);
			if (string.IsNullOrEmpty(text8))
			{
				HTKKFile.SetNodeBackGroundPink(sView, CellId_CoQuanBHDADT[1], CellId_CoQuanBHDADT[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_37")));
				value = 2;
			}
			HTKKFile.RemoveValidateBackGround1(sView, CellId_MSTDADT[1], CellId_MSTDADT[0]);
			if (string.IsNullOrEmpty(text9))
			{
				HTKKFile.SetNodeBackGroundPink(sView, CellId_MSTDADT[1], CellId_MSTDADT[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_37")));
				value = 2;
			}
			else if (!ValidateBase.CheckMaSoThue(text9.Trim()))
			{
				HTKKFile.SetNodeBackGroundPink(sView, CellId_MSTDADT[1], CellId_MSTDADT[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_21")));
				value = 2;
			}
			else if ((mst.Length == 13 || mst.Length == 14) && !mst.Substring(0, 10).Equals(text9.Substring(0, 10)))
			{
				HTKKFile.SetNodeBackGroundPink(sView, CellId_MSTDADT[1], CellId_MSTDADT[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_18")));
				value = 2;
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

	private string[] getTTDuAnDT(string maDADT, HTKKHeaderInfo headerInfo)
	{
		//IL_006e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0075: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.DATAFILES + "\\" + headerInfo.mst;
			string path = text + "\\DK_Ma_du_an_DT_100.xml";
			if (File.Exists(path))
			{
				XmlDataDocument xmlDataDocument = new XmlDataDocument();
				FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read);
				((XmlDocument)xmlDataDocument).Load((Stream)fileStream);
				XmlNodeList elementsByTagName = ((XmlDocument)xmlDataDocument).GetElementsByTagName("RowData");
				foreach (XmlNode item in elementsByTagName)
				{
					XmlNode val = item;
					string innerText = val.InnerText;
					string[] array = Regex.Split(innerText, "###");
					if (array[1] == maDADT)
					{
						return array;
					}
				}
			}
			return null;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
