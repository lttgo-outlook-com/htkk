using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TncnKk02TT80Nckt283
{
	private bool iflag;

	private string mst = "";

	private int namKK = 0;

	private static string cellId_ct23 = "R_78";

	private static string cellId_ct28 = "R_83";

	private static string cellId_ct29 = "R_84";

	private int[] cellTemSoThang = GeneralFunctions.ArrayCellId("AD_78");

	private int[] cellTinh = GeneralFunctions.ArrayCellId("I_18");

	private int[] cellMaTinh = GeneralFunctions.ArrayCellId("AK_18");

	private int[] cellXa = GeneralFunctions.ArrayCellId("I_22");

	private int[] cellMaXa = GeneralFunctions.ArrayCellId("AK_22");

	private static int[] CellId_ct23 => GeneralFunctions.ArrayCellId(cellId_ct23);

	private static int[] CellId_ct28 => GeneralFunctions.ArrayCellId(cellId_ct28);

	private static int[] CellId_ct29 => GeneralFunctions.ArrayCellId(cellId_ct29);

	public bool IFlag()
	{
		return iflag;
	}

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int? num = (GetFirstNumber(PeriodInfo.TuThang).HasValue ? GetFirstNumber(PeriodInfo.TuThang) : ((int?)null));
			int? num2 = (GetFirstNumber(PeriodInfo.DenThang).HasValue ? GetFirstNumber(PeriodInfo.DenThang) : ((int?)null));
			string quy = PeriodInfo.Quy;
			string nam = PeriodInfo.Nam;
			namKK = ConvertHelpers.ConvertToInt(nam);
			sView.Cells[cellTemSoThang[1], cellTemSoThang[0]].Value = ((!num2.HasValue || !num.HasValue) ? new int?(1) : (num2 - num + 1));
			if (!ToKhaiInfo.SheetInfos[0].IsReadDataFile)
			{
				if (PeriodInfo.KieuKy == "Q")
				{
					string text = sView.Cells[CellId_ct23[1], CellId_ct23[0]].Text;
					if (num.HasValue && num2.HasValue && (string.IsNullOrEmpty(text) || text.Equals("0")))
					{
						sView.Cells[cellTemSoThang[1], cellTemSoThang[0]].Value = num2 - num + 1;
						if (namKK >= 2026)
						{
							sView.Cells[CellId_ct23[1], CellId_ct23[0]].Value = (num2 - num + 1) * 15500000;
						}
						else
						{
							sView.Cells[CellId_ct23[1], CellId_ct23[0]].Value = (num2 - num + 1) * 11000000;
						}
					}
				}
				else
				{
					sView.Cells[cellTemSoThang[1], cellTemSoThang[0]].Value = 1;
					if (namKK >= 2026)
					{
						sView.Cells[CellId_ct23[1], CellId_ct23[0]].Value = 15500000;
					}
					else
					{
						sView.Cells[CellId_ct23[1], CellId_ct23[0]].Value = 11000000;
					}
				}
			}
			sView.Cells[cellTinh[1], cellTinh[0]].Text = TTinChung.tenTinhNNT;
			sView.Cells[cellMaTinh[1], cellMaTinh[0]].Text = TTinChung.maTinhNNT;
			sView.Cells[cellXa[1], cellXa[0]].Text = TTinChung.tenXaPhuongNNT;
			sView.Cells[cellMaXa[1], cellMaXa[0]].Text = TTinChung.maXaPhuongNNT;
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
			if (winFpSpread.ActiveSheetIndex != 0)
			{
				return;
			}
			int num = HTKKFile.ColumnLettersToIndex("R");
			if (num == iCol && iRow >= 75 && iRow <= 82)
			{
				winFpSpread.Sheets[0].RecalculateAll();
				int? num2 = (GetFirstNumber(PeriodInfo.TuThang).HasValue ? GetFirstNumber(PeriodInfo.TuThang) : ((int?)null));
				int? sothang = (GetFirstNumber(PeriodInfo.DenThang).HasValue ? GetFirstNumber(PeriodInfo.DenThang) : ((int?)null)) - num2 + 1;
				if (PeriodInfo.KieuKy == "M")
				{
					sothang = 1;
				}
				double valueCT = Convert.ToDouble(winFpSpread.Sheets[0].Cells[CellId_ct28[1], CellId_ct28[0]].Value);
				winFpSpread.Sheets[0].Cells[CellId_ct29[1], CellId_ct29[0]].Formula = GetFormula29(valueCT, sothang);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		bool result = true;
		int value = 0;
		levelWarning.TryGetValue(0, out value);
		SheetView sheetView = winFpSpread.Sheets[0];
		int[] array = GeneralFunctions.ArrayCellId("T_17");
		string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[24, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[28, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text3 = sheetView.Cells[array[1], array[0]].Text.Trim();
		HTKKFile.RemoveValidate(sheetView, array[1], array[0]);
		if (string.IsNullOrEmpty(text3))
		{
			HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], text);
			value = 2;
		}
		else if (!ValidateBase.CheckMaSoThueCCCD(text3))
		{
			HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], text2);
			value = 2;
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
		return result;
	}

	private int? GetFirstNumber(string strInput)
	{
		try
		{
			if (string.IsNullOrEmpty(strInput))
			{
				return null;
			}
			string s = strInput.Substring(0, strInput.IndexOf('/'));
			return int.Parse(s);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	private string GetFormula29(double valueCT28, int? sothang)
	{
		string text = "0";
		double num = 1000000.0;
		double num2 = valueCT28 / (double)sothang.Value;
		string newValue = valueCT28 + "/" + sothang;
		text = ((namKK >= 2026) ? ((num2 <= 10.0 * num) ? "R83*5%" : ((num2 > 10.0 * num && num2 <= 30.0 * num) ? "10000000*5% + (R83-10000000)*10%" : ((num2 > 30.0 * num && num2 <= 60.0 * num) ? "10000000*5% + 20000000*10% + (R83-30000000)*20%" : ((!(num2 > 60.0 * num) || !(num2 <= 100.0 * num)) ? "10000000*5% + 20000000*10% + 30000000*20% + 40000000*30% + (R83-100000000)*35%" : "10000000*5% + 20000000*10% + 30000000*20% + (R83-60000000)*30%")))) : ((num2 <= 5.0 * num) ? "R83*5%" : ((num2 > 5.0 * num && num2 <= 10.0 * num) ? "5000000*5% + (R83-5000000)*10%" : ((num2 > 10.0 * num && num2 <= 18.0 * num) ? "5000000*5% + 5000000*10% + (R83-10000000)*15%" : ((num2 > 18.0 * num && num2 <= 32.0 * num) ? "5000000*5% + 5000000*10% + 8000000*15% + (R83-18000000)*20%" : ((num2 > 32.0 * num && num2 <= 52.0 * num) ? "5000000*5% + 5000000*10% + 8000000*15% + 14000000*20% + (R83-32000000)*25%" : ((!(num2 > 52.0 * num) || !(num2 <= 80.0 * num)) ? "5000000*5% + 5000000*10% + 8000000*15% + 14000000*20% + 20000000*25% + 28000000*30% + (R83-80000000)*35%" : "5000000*5% + 5000000*10% + 8000000*15% + 14000000*20% + 20000000*25% + (R83-52000000)*30%")))))));
		return "ROUND((" + text.Replace("R83", newValue) + ")*" + sothang + ",0)";
	}

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			XmlNode val = xmlNode.SelectSingleNode("HSoKhaiThue/TTinChung/TTinTKhaiThue/NNT/mst");
			if (val != null && !string.IsNullOrEmpty(val.InnerText))
			{
				string innerText = val.InnerText.Trim();
				XmlNode val2 = xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/mst_cu");
				string innerText2 = (mst = val2.InnerText.Trim());
				if (val2 != null)
				{
					val2.InnerText = innerText;
				}
				if (val != null)
				{
					val.InnerText = innerText2;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ExportXML_1(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			XmlNode val = xmlNode.SelectSingleNode("Header/mst");
			if (val != null)
			{
				val.InnerText = mst;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ExportXML_2(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			XmlNode val = xmlNode.SelectSingleNode("Header/mst");
			if (val != null)
			{
				val.InnerText = mst;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void EnterCell_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
		int[] array = GeneralFunctions.ArrayCellId("T_17");
		winFpSpread.Sheets[0].Cells[array[1], array[0]].Note = "";
		if ((int)iRow == array[1] + 1 && (int)iCol == array[0] + 1)
		{
			winFpSpread.Sheets[0].Cells[array[1], array[0]].NoteStyle = NoteStyle.StickyNote;
			winFpSpread.Sheets[0].Cells[array[1], array[0]].Note = xMLMessage.GetMessage("0655");
		}
	}
}
