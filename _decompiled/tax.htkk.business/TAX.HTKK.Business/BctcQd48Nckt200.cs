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

public class BctcQd48Nckt200
{
	private bool isExitDataFile = false;

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			if (ToKhaiInfo.SheetInfos[0] != null)
			{
				Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
			}
			if (ToKhaiInfo.SheetInfos[4] != null)
			{
				Validate_4(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
		return true;
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[10, HTKKFile.ColumnLettersToIndex("D") - 1].Value))
			{
				if (string.IsNullOrEmpty(sheetView.Cells[10, HTKKFile.ColumnLettersToIndex("T") - 1].Text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 10, HTKKFile.ColumnLettersToIndex("T") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, 10, HTKKFile.ColumnLettersToIndex("T") - 1);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, 10, HTKKFile.ColumnLettersToIndex("T") - 1);
			}
			if (value == 2 && winFpSpread.TabStrip[0].BackColor != Color.Pink)
			{
				winFpSpread.TabStrip[0].BackColor = Color.Pink;
			}
			else if (value == 1 && winFpSpread.TabStrip[0].BackColor == Color.White)
			{
				winFpSpread.TabStrip[0].BackColor = Color.Yellow;
			}
			else if (value == 0)
			{
				winFpSpread.TabStrip[0].BackColor = Color.White;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int num = 0;
			SheetView sheetView = winFpSpread.Sheets[4];
			for (int i = 144; i < 164; i++)
			{
				if (i != 149 && i != 157)
				{
					float num2 = float.Parse(sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Text);
					float num3 = float.Parse(sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Text);
					float num4 = float.Parse(sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("AM") - 1].Text);
					float num5 = float.Parse(sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("AH") - 1].Text);
					if (num4 != 0f && num5 != 0f)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, i - 1, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_292")));
						num = 2;
						HTKKFile.SetNodeBackGroundPink(sheetView, i - 1, HTKKFile.ColumnLettersToIndex("AH") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_292")));
						num = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, i - 1, HTKKFile.ColumnLettersToIndex("AM") - 1);
						HTKKFile.RemoveValidate(sheetView, i - 1, HTKKFile.ColumnLettersToIndex("AH") - 1);
					}
					if (num2 != 0f && num3 != 0f)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, i - 1, HTKKFile.ColumnLettersToIndex("R") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_292")));
						num = 2;
						HTKKFile.SetNodeBackGroundPink(sheetView, i - 1, HTKKFile.ColumnLettersToIndex("V") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_292")));
						num = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, i - 1, HTKKFile.ColumnLettersToIndex("R") - 1);
						HTKKFile.RemoveValidate(sheetView, i - 1, HTKKFile.ColumnLettersToIndex("V") - 1);
					}
					if (num == 2 && winFpSpread.TabStrip[0].BackColor != Color.Pink)
					{
						winFpSpread.TabStrip[0].BackColor = Color.Pink;
					}
					else if (num == 1 && winFpSpread.TabStrip[0].BackColor == Color.White)
					{
						winFpSpread.TabStrip[0].BackColor = Color.Yellow;
					}
					else if (num == 0)
					{
						winFpSpread.TabStrip[0].BackColor = Color.White;
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

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("D_11");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				sheetView.Cells[10, HTKKFile.ColumnLettersToIndex("T") - 1].Locked = !Convert.ToBoolean(valueCustom);
				if (!Convert.ToBoolean(valueCustom))
				{
					sheetView.Cells[10, HTKKFile.ColumnLettersToIndex("T") - 1].Text = "";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellChange_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[4];
			int num = HTKKFile.ColumnLettersToIndex("AH");
			int num2 = HTKKFile.ColumnLettersToIndex("AM");
			int num3 = HTKKFile.ColumnLettersToIndex("R");
			int num4 = HTKKFile.ColumnLettersToIndex("V");
			int num5 = HTKKFile.ColumnLettersToIndex("Z");
			int num6 = HTKKFile.ColumnLettersToIndex("AD");
			string text = "IF(R###-V###+Z###-AD###>0,R###-V###+Z###-AD###,0)";
			string text2 = "IF(R###-V###+Z###-AD###>0,0,ABS(R###-V###+Z###-AD###))";
			if (iCol == num3 || iCol == num4 || iCol == num5 || iCol == num6)
			{
				sheetView.Cells[(int)iRow - 1, num - 1].Formula = text.Replace("###", iRow.ToString());
				sheetView.Cells[(int)iRow - 1, num2 - 1].Formula = text2.Replace("###", iRow.ToString());
			}
			if (iCol == num || iCol == num2)
			{
				sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Formula = "";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("D_11");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			sheetView.Cells[10, HTKKFile.ColumnLettersToIndex("T") - 1].Locked = !Convert.ToBoolean(valueCustom);
			string text = HTKKConfig.DATAFILES + "\\" + TTinChung.mst + "\\" + ToKhaiInfo.SheetInfos[0].DataFile;
			text = text + "_" + int.Parse(PeriodInfo.Nam);
			text = text + "_" + ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
			if (File.Exists(text))
			{
				isExitDataFile = true;
			}
			if (!isExitDataFile)
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, 0);
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
			if (!isExitDataFile)
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, 1);
			}
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
			if (!isExitDataFile)
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, 2);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			if (!isExitDataFile)
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, 3);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void getSoKyTruoc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int sheetIndex)
	{
		try
		{
			string text = HTKKConfig.DATAFILES + "\\" + TTinChung.mst + "\\" + ToKhaiInfo.SheetInfos[sheetIndex].DataFile;
			text = text + "_" + (int.Parse(PeriodInfo.Nam) - 1);
			text = text + "_" + ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			int num = 0;
			int num2 = 0;
			int num3 = HTKKFile.ColumnLettersToIndex("AC");
			if (sheetIndex == 4)
			{
				return;
			}
			if (sheetIndex == 0 || sheetIndex == 2 || sheetIndex == 3)
			{
				num = HTKKFile.ColumnLettersToIndex("W");
				num2 = HTKKFile.ColumnLettersToIndex("R");
			}
			if (sheetIndex == 1)
			{
				num = HTKKFile.ColumnLettersToIndex("J");
				num2 = HTKKFile.ColumnLettersToIndex("I");
			}
			if (!File.Exists(text))
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
			hTKKCMCSheetInfo.GetDetailSheetInfo(text);
			for (int i = 0; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
			{
				if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num2)
				{
					sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ExportHeaderXML_0(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			List<HTKKCMCCellInfo> staticCells = ToKhaiInfo.SheetInfos[0].StaticCells;
			string innerText = "";
			string innerText2 = "";
			foreach (HTKKCMCCellInfo item in staticCells)
			{
				if (item.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayLap")
				{
					string value = item.Value;
					value = value.Replace("/", "");
					if (value.Length == 8)
					{
						innerText2 = value.Substring(4, 4) + "-" + value.Substring(2, 2) + "-" + value.Substring(0, 2);
					}
				}
				if (item.Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/giamDoc")
				{
					innerText = item.Value;
				}
			}
			XmlNode val = xmlNode.SelectSingleNode("HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue");
			XmlNode val2 = val.SelectSingleNode("nguoiKy");
			XmlNode val3 = val.SelectSingleNode("ngayKy");
			if (val2 != null)
			{
				val2.InnerText = innerText;
			}
			if (val3 != null)
			{
				val3.InnerText = innerText2;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
