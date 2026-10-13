using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Xml;
using FarPoint.Win;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BctcTt200Nckt218
{
	private int rowBeginLT = 53;

	private int rowEndLT = 172;

	private int rowBeginKLT = 173;

	private int rowEndKLT = 249;

	private int rowBeginLTExp = 54;

	private int rowEndLTExp = 172;

	private int rowBeginKLTExp = 173;

	private int rowEndKLTExp = 249;

	private HTKKCMCToKhaiInfo tKhaiInfoOld = new HTKKCMCToKhaiInfo();

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			if (PeriodInfo.DoanhNghiepHoatDongLienTuc == "LT")
			{
				if (HTKKFormater.isNullString(xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/CDKT_HoatDongLienTuc/tyGiaQuyDoi").InnerText))
				{
					xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/CDKT_HoatDongLienTuc/tyGiaQuyDoi").InnerText = "1";
				}
				xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/hoatDongLienTuc").InnerText = "true";
				xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/hoatDongKhongLienTuc").InnerText = "false";
				XmlNode val = xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/CDKT_HoatDongKhongLienTuc");
				xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh").RemoveChild(val);
			}
			else
			{
				if (HTKKFormater.isNullString(xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/CDKT_HoatDongKhongLienTuc/tyGiaQuyDoi").InnerText))
				{
					xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/CDKT_HoatDongKhongLienTuc/tyGiaQuyDoi").InnerText = "1";
				}
				xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/hoatDongLienTuc").Value = "false";
				xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/hoatDongKhongLienTuc").Value = "true";
				XmlNode val2 = xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/CDKT_HoatDongLienTuc");
				xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh").RemoveChild(val2);
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
			XmlNode val = xmlNode.SelectSingleNode("PL_KQHDSXKD/PhanKy/ngayLap");
			if (val != null)
			{
				string innerText = val.InnerText;
				if (innerText.Length >= 8)
				{
					val.InnerText = innerText.Substring(4, 4) + "-" + innerText.Substring(2, 1).PadLeft(2, '0') + "-" + innerText.Substring(0, 1).PadLeft(2, '0');
				}
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
			XmlNode val = xmlNode.SelectSingleNode("PL_LCTTTT/PhanKy/ngayLap");
			if (val != null)
			{
				string innerText = val.InnerText;
				if (innerText.Length >= 8)
				{
					val.InnerText = innerText.Substring(4, 4) + "-" + innerText.Substring(2, 1).PadLeft(2, '0') + "-" + innerText.Substring(0, 1).PadLeft(2, '0');
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ExportXML_3(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			XmlNode val = xmlNode.SelectSingleNode("PL_LCTTGT/PhanKy/ngayLap");
			if (val != null)
			{
				string innerText = val.InnerText;
				if (innerText.Length >= 8)
				{
					val.InnerText = innerText.Substring(4, 4) + "-" + innerText.Substring(2, 1).PadLeft(2, '0') + "-" + innerText.Substring(0, 1).PadLeft(2, '0');
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void SetFPSpreadExp(FpSpread winFpExp, List<HTKKCMCSheetInfo> sheetInfos, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKHeaderInfo HTKKHeaderInfo)
	{
		try
		{
			int num = 0;
			for (num = 0; num < winFpExp.Sheets.Count - 1; num++)
			{
				SheetView sheetView = winFpExp.Sheets[num];
				HTKKCMCSheetInfo hTKKCMCSheetInfo = sheetInfos[num];
				if (num == 0)
				{
					if (PeriodInfo.DoanhNghiepHoatDongLienTuc == "LT")
					{
						GeneralFunctions.setHindenRow(sheetView, rowBeginKLTExp, rowEndKLTExp - 1);
						sheetView.Rows[22].Visible = false;
					}
					else if (PeriodInfo.DoanhNghiepHoatDongLienTuc == "KL")
					{
						GeneralFunctions.setHindenRow(sheetView, rowBeginLTExp, rowEndLTExp - 1);
						sheetView.Rows[21].Visible = false;
					}
				}
			}
			winFpExp.ActiveSheetIndex = 0;
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			if (PeriodInfo.DonViTinhMa == "NT")
			{
				for (int i = 0; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex > rowBeginLT && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex < rowEndKLT)
					{
						NumberCellType numberCellType = new NumberCellType();
						numberCellType.Separator = ".";
						numberCellType.DecimalSeparator = ",";
						numberCellType.DecimalPlaces = 2;
						numberCellType.ShowSeparator = true;
						numberCellType.NegativeFormat = NegativeFormat.Parentheses;
						numberCellType.MaximumValue = GeneralFunctions.getMaxNumber(hTKKCMCSheetInfo.StaticCells[i].Validate.MaxValue);
						numberCellType.MinimumValue = GeneralFunctions.getMinNumber(hTKKCMCSheetInfo.StaticCells[i].Validate.MinValue);
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex - 1].CellType = numberCellType;
					}
				}
				NumberCellType numberCellType2 = new NumberCellType();
				numberCellType2.Separator = ".";
				numberCellType2.DecimalSeparator = ",";
				numberCellType2.DecimalPlaces = 0;
				numberCellType2.MinimumValue = 0.0;
				numberCellType2.ShowSeparator = true;
				numberCellType2.MaximumValue = 999999999999999.0;
				sheetView.Cells[37, HTKKFile.ColumnLettersToIndex("O") - 1].CellType = numberCellType2;
			}
			else
			{
				sheetView.Cells[37, HTKKFile.ColumnLettersToIndex("L") - 1].Text = "";
				sheetView.Cells[37, HTKKFile.ColumnLettersToIndex("O") - 1].Value = 0;
				Color color = default(Color);
				color = Color.FromArgb(240, 240, 240);
				sheetView.Cells[37, HTKKFile.ColumnLettersToIndex("O") - 1].BackColor = color;
				sheetView.Cells[37, HTKKFile.ColumnLettersToIndex("O") - 1].ForeColor = color;
				sheetView.Cells[37, HTKKFile.ColumnLettersToIndex("O") - 1].Border = new BevelBorder(BevelBorderType.Raised, color, color);
				sheetView.Cells[37, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				for (int i = 0; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex > rowBeginLT && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex < rowEndKLT)
					{
						NumberCellType numberCellType = new NumberCellType();
						numberCellType.DecimalPlaces = 0;
						numberCellType.NegativeFormat = NegativeFormat.Parentheses;
						numberCellType.Separator = ".";
						numberCellType.DecimalSeparator = ",";
						numberCellType.MaximumValue = GeneralFunctions.getMaxNumber(hTKKCMCSheetInfo.StaticCells[i].Validate.MaxValue);
						numberCellType.MinimumValue = GeneralFunctions.getMinNumber(hTKKCMCSheetInfo.StaticCells[i].Validate.MinValue);
						numberCellType.ShowSeparator = true;
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex - 1].CellType = numberCellType;
					}
				}
			}
			winFpSpread.ActiveSheetIndex = 0;
			if (PeriodInfo.DoanhNghiepHoatDongLienTuc == "LT")
			{
				sheetView.Cells[44, HTKKFile.ColumnLettersToIndex("K") - 1].Text = "1";
				sheetView.Cells[44, HTKKFile.ColumnLettersToIndex("L") - 1].Text = "0";
				for (int j = rowBeginKLT; j < rowEndKLT; j++)
				{
					Row row = winFpSpread.ActiveSheet.Rows[j];
					row.Visible = false;
				}
				winFpSpread.ActiveSheet.Rows[2].Visible = false;
			}
			else if (PeriodInfo.DoanhNghiepHoatDongLienTuc == "KL")
			{
				sheetView.Cells[44, HTKKFile.ColumnLettersToIndex("K") - 1].Text = "";
				sheetView.Cells[44, HTKKFile.ColumnLettersToIndex("L") - 1].Text = "1";
				for (int j = rowBeginLT; j < rowEndLT; j++)
				{
					Row row = winFpSpread.ActiveSheet.Rows[j];
					row.Visible = false;
				}
				winFpSpread.ActiveSheet.Rows[1].Visible = false;
			}
			int column = HTKKFile.ColumnLettersToIndex("R") - 1;
			int column2 = HTKKFile.ColumnLettersToIndex("W") - 1;
			int[] array = new int[2];
			if (PeriodInfo.DoanhNghiepHoatDongLienTuc.ToString() == "LT")
			{
				array[0] = 152;
				array[1] = 153;
			}
			else
			{
				array[0] = 232;
				array[1] = 233;
			}
			double num = double.Parse(sheetView.Cells[array[0] - 1, column].Value.ToString());
			double num2 = double.Parse(sheetView.Cells[array[1] - 1, column].Value.ToString());
			if (num + num2 == 0.0)
			{
				sheetView.Cells[array[0] - 2, column].Locked = false;
			}
			double num3 = double.Parse(sheetView.Cells[array[0] - 1, column2].Value.ToString());
			double num4 = double.Parse(sheetView.Cells[array[1] - 1, column2].Value.ToString());
			if (num3 + num4 == 0.0)
			{
				sheetView.Cells[array[0] - 2, column2].Locked = false;
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
			if (PeriodInfo.DonViTinhMa == "NT")
			{
				for (int i = 0; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex >= 22 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex <= 41)
					{
						NumberCellType numberCellType = new NumberCellType();
						numberCellType.Separator = ".";
						numberCellType.DecimalSeparator = ",";
						numberCellType.DecimalPlaces = 2;
						numberCellType.ShowSeparator = true;
						numberCellType.NegativeFormat = NegativeFormat.Parentheses;
						numberCellType.MaximumValue = GeneralFunctions.getMaxNumber(hTKKCMCSheetInfo.StaticCells[i].Validate.MaxValue);
						numberCellType.MinimumValue = GeneralFunctions.getMinNumber(hTKKCMCSheetInfo.StaticCells[i].Validate.MinValue);
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex - 1].CellType = numberCellType;
					}
				}
				return;
			}
			for (int i = 0; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
			{
				if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex >= 22 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex <= 41)
				{
					NumberCellType numberCellType = new NumberCellType();
					numberCellType.DecimalPlaces = 0;
					numberCellType.NegativeFormat = NegativeFormat.Parentheses;
					numberCellType.Separator = ".";
					numberCellType.DecimalSeparator = ",";
					numberCellType.MaximumValue = GeneralFunctions.getMaxNumber(hTKKCMCSheetInfo.StaticCells[i].Validate.MaxValue);
					numberCellType.MinimumValue = GeneralFunctions.getMinNumber(hTKKCMCSheetInfo.StaticCells[i].Validate.MinValue);
					numberCellType.ShowSeparator = true;
					sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex - 1].CellType = numberCellType;
				}
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			SheetView sheetView = winFpSpread.Sheets[2];
			if (PeriodInfo.DonViTinhMa == "NT")
			{
				for (int i = 0; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex >= 44 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex <= 72)
					{
						NumberCellType numberCellType = new NumberCellType();
						numberCellType.Separator = ".";
						numberCellType.DecimalSeparator = ",";
						numberCellType.DecimalPlaces = 2;
						numberCellType.ShowSeparator = true;
						numberCellType.NegativeFormat = NegativeFormat.Parentheses;
						numberCellType.MaximumValue = GeneralFunctions.getMaxNumber(hTKKCMCSheetInfo.StaticCells[i].Validate.MaxValue);
						numberCellType.MinimumValue = GeneralFunctions.getMinNumber(hTKKCMCSheetInfo.StaticCells[i].Validate.MinValue);
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex - 1].CellType = numberCellType;
					}
				}
				return;
			}
			for (int i = 0; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
			{
				if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex >= 44 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex <= 72)
				{
					NumberCellType numberCellType = new NumberCellType();
					numberCellType.DecimalPlaces = 0;
					numberCellType.NegativeFormat = NegativeFormat.Parentheses;
					numberCellType.Separator = ".";
					numberCellType.DecimalSeparator = ",";
					numberCellType.MaximumValue = GeneralFunctions.getMaxNumber(hTKKCMCSheetInfo.StaticCells[i].Validate.MaxValue);
					numberCellType.MinimumValue = GeneralFunctions.getMinNumber(hTKKCMCSheetInfo.StaticCells[i].Validate.MinValue);
					numberCellType.ShowSeparator = true;
					sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex - 1].CellType = numberCellType;
				}
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[3];
			SheetView sheetView = winFpSpread.Sheets[3];
			if (PeriodInfo.DonViTinhMa == "NT")
			{
				for (int i = 0; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex >= 48 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex <= 87)
					{
						NumberCellType numberCellType = new NumberCellType();
						numberCellType.Separator = ".";
						numberCellType.DecimalSeparator = ",";
						numberCellType.DecimalPlaces = 2;
						numberCellType.ShowSeparator = true;
						numberCellType.NegativeFormat = NegativeFormat.Parentheses;
						numberCellType.MaximumValue = GeneralFunctions.getMaxNumber(hTKKCMCSheetInfo.StaticCells[i].Validate.MaxValue);
						numberCellType.MinimumValue = GeneralFunctions.getMinNumber(hTKKCMCSheetInfo.StaticCells[i].Validate.MinValue);
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex - 1].CellType = numberCellType;
					}
				}
				return;
			}
			for (int i = 0; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
			{
				if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex >= 48 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex <= 87)
				{
					NumberCellType numberCellType = new NumberCellType();
					numberCellType.DecimalPlaces = 0;
					numberCellType.NegativeFormat = NegativeFormat.Parentheses;
					numberCellType.Separator = ".";
					numberCellType.DecimalSeparator = ",";
					numberCellType.MaximumValue = GeneralFunctions.getMaxNumber(hTKKCMCSheetInfo.StaticCells[i].Validate.MaxValue);
					numberCellType.MinimumValue = GeneralFunctions.getMinNumber(hTKKCMCSheetInfo.StaticCells[i].Validate.MinValue);
					numberCellType.ShowSeparator = true;
					sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex - 1].CellType = numberCellType;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			bool result = true;
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			if (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) > 2015)
			{
				Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
				Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
				Validate_3(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			}
			if (winFpSpread.TabStrip[0].BackColor != Color.White && winFpSpread.TabStrip[1].BackColor != Color.White && winFpSpread.TabStrip[2].BackColor != Color.White && winFpSpread.TabStrip[3].BackColor != Color.White)
			{
				result = false;
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
					result = false;
					winFpSpread.TabStrip[key].BackColor = Color.Yellow;
				}
				else
				{
					result = false;
					winFpSpread.TabStrip[key].BackColor = Color.White;
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

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int index = 0;
			SheetView sheetView = winFpSpread.Sheets[index];
			int num = 0;
			int num2 = HTKKFile.ColumnLettersToIndex("AC");
			int num3 = HTKKFile.ColumnLettersToIndex("W");
			int num4 = HTKKFile.ColumnLettersToIndex("O");
			int num5 = 38;
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[44, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, num5, num4, "NumberCellType");
			if (PeriodInfo.DonViTinhMa.Equals("NT"))
			{
				if (ConvertHelpers.ConvertToDouble(valueCustom) == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num5 - 1, num4 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_39")));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num5 - 1, num4 - 1);
				}
			}
			if (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) > 2015 && tKhaiInfoOld.SheetInfos.Count > 0 && tKhaiInfoOld.SheetInfos[index] != null)
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = tKhaiInfoOld.SheetInfos[index];
				for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != num3)
					{
						continue;
					}
					Color backColor = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor;
					string formula = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Formula;
					string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num3, "NumberCellType");
					if (valueCustom2 != hTKKCMCSheetInfo.StaticCells[i].Value && !HTKKFormater.isNullString(valueCustom2))
					{
						if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor != Color.Pink)
						{
							if (string.IsNullOrEmpty(formula))
							{
								num = 1;
								HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1, text);
							}
							else
							{
								sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Note = text;
								sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor = backColor;
							}
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1);
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor = backColor;
					}
				}
			}
			if (num == 2 && winFpSpread.TabStrip[0].BackColor != Color.Pink)
			{
				winFpSpread.TabStrip[index].BackColor = Color.Pink;
			}
			else if (num == 1 && winFpSpread.TabStrip[0].BackColor == Color.White)
			{
				winFpSpread.TabStrip[index].BackColor = Color.Yellow;
			}
			else
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
			int num = 0;
			int num2 = HTKKFile.ColumnLettersToIndex("AC");
			int num3 = HTKKFile.ColumnLettersToIndex("W");
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[44, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			if (tKhaiInfoOld.SheetInfos.Count > 0 && tKhaiInfoOld.SheetInfos[index] != null)
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = tKhaiInfoOld.SheetInfos[index];
				for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != num3)
					{
						continue;
					}
					Color backColor = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor;
					string formula = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Formula;
					string valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num3, "NumberCellType");
					if (valueCustom != hTKKCMCSheetInfo.StaticCells[i].Value && !HTKKFormater.isNullString(valueCustom))
					{
						if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor != Color.Pink)
						{
							if (string.IsNullOrEmpty(formula))
							{
								num = 1;
								HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1, text);
							}
							else
							{
								sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Note = text;
								sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor = backColor;
							}
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1);
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor = backColor;
					}
				}
			}
			if (num == 2 && winFpSpread.TabStrip[0].BackColor != Color.Pink)
			{
				winFpSpread.TabStrip[index].BackColor = Color.Pink;
			}
			else if (num == 1 && winFpSpread.TabStrip[0].BackColor == Color.White)
			{
				winFpSpread.TabStrip[index].BackColor = Color.Yellow;
			}
			else
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

	public void Validate_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int index = 2;
			SheetView sheetView = winFpSpread.Sheets[index];
			int num = 0;
			int num2 = HTKKFile.ColumnLettersToIndex("AC");
			int num3 = HTKKFile.ColumnLettersToIndex("W");
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[44, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			if (tKhaiInfoOld.SheetInfos.Count > 0 && tKhaiInfoOld.SheetInfos[index] != null)
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = tKhaiInfoOld.SheetInfos[index];
				for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != num3)
					{
						continue;
					}
					Color backColor = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor;
					string formula = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Formula;
					string valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num3, "NumberCellType");
					if (valueCustom != hTKKCMCSheetInfo.StaticCells[i].Value && !HTKKFormater.isNullString(valueCustom))
					{
						if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor != Color.Pink)
						{
							if (string.IsNullOrEmpty(formula))
							{
								num = 1;
								HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1, text);
							}
							else
							{
								sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Note = text;
								sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor = backColor;
							}
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1);
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor = backColor;
					}
				}
			}
			if (num == 2 && winFpSpread.TabStrip[0].BackColor != Color.Pink)
			{
				winFpSpread.TabStrip[index].BackColor = Color.Pink;
			}
			else if (num == 1 && winFpSpread.TabStrip[0].BackColor == Color.White)
			{
				winFpSpread.TabStrip[index].BackColor = Color.Yellow;
			}
			else
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

	public void Validate_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int index = 3;
			SheetView sheetView = winFpSpread.Sheets[index];
			int num = 0;
			int num2 = HTKKFile.ColumnLettersToIndex("AC");
			int num3 = HTKKFile.ColumnLettersToIndex("W");
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[44, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			if (tKhaiInfoOld.SheetInfos.Count > 0 && tKhaiInfoOld.SheetInfos[index] != null)
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = tKhaiInfoOld.SheetInfos[index];
				for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != num3)
					{
						continue;
					}
					Color backColor = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor;
					string formula = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Formula;
					string valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num3, "NumberCellType");
					if (valueCustom != hTKKCMCSheetInfo.StaticCells[i].Value && !HTKKFormater.isNullString(valueCustom))
					{
						if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor != Color.Pink)
						{
							if (string.IsNullOrEmpty(formula))
							{
								num = 1;
								HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1, text);
							}
							else
							{
								sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Note = text;
								sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor = backColor;
							}
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1);
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor = backColor;
					}
				}
			}
			if (num == 2 && winFpSpread.TabStrip[0].BackColor != Color.Pink)
			{
				winFpSpread.TabStrip[index].BackColor = Color.Pink;
			}
			else if (num == 1 && winFpSpread.TabStrip[0].BackColor == Color.White)
			{
				winFpSpread.TabStrip[index].BackColor = Color.Yellow;
			}
			else
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

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("D") - 1;
			int num2 = 51;
			if (idxCol == num && idxRow == num2)
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, 0, sheetView.Cells[num2, num].Value.ToBoolOrDefault());
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			int num = HTKKFile.ColumnLettersToIndex("B") - 1;
			int num2 = 16;
			if (idxCol == num && idxRow == num2)
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, 1, sheetView.Cells[num2, num].Value.ToBoolOrDefault());
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[2];
			int num = HTKKFile.ColumnLettersToIndex("C") - 1;
			int num2 = 38;
			if (idxCol == num && idxRow == num2)
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, 2, sheetView.Cells[num2, num].Value.ToBoolOrDefault());
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[4];
			int num = HTKKFile.ColumnLettersToIndex("C") - 1;
			int num2 = 42;
			if (idxCol == num && idxRow == num2)
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, 3, sheetView.Cells[num2, num].Value.ToBoolOrDefault());
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
			int column = HTKKFile.ColumnLettersToIndex("R") - 1;
			int column2 = HTKKFile.ColumnLettersToIndex("W") - 1;
			int[] array = new int[2];
			if (PeriodInfo.DoanhNghiepHoatDongLienTuc.ToString() == "LT")
			{
				array[0] = 152;
				array[1] = 153;
			}
			else
			{
				array[0] = 232;
				array[1] = 233;
			}
			string text = HTKKFile.ColumnIndexToLetters(winFpSpread.ActiveSheet.ActiveCell.Column.Index + 1) + (winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1);
			if (text == "R" + array[0] || text == "R" + array[1])
			{
				double num = ConvertHelpers.ConvertToDouble1(sheetView.Cells[array[0] - 1, column].Text);
				double num2 = ConvertHelpers.ConvertToDouble1(sheetView.Cells[array[1] - 1, column].Text);
				if (num + num2 != 0.0)
				{
					if (array[0] == 152)
					{
						sheetView.Cells[array[0] - 2, column].Locked = true;
						sheetView.Cells[array[0] - 2, column].Formula = "R152+R153";
					}
					else
					{
						sheetView.Cells[array[0] - 2, column].Locked = true;
						sheetView.Cells[array[0] - 2, column].Formula = "R232+R233";
					}
				}
				else
				{
					sheetView.Cells[array[0] - 2, column].Locked = false;
				}
			}
			if (text == "W" + array[0] || text == "W" + array[1])
			{
				double num3 = ConvertHelpers.ConvertToDouble1(sheetView.Cells[array[0] - 1, column2].Text);
				double num4 = ConvertHelpers.ConvertToDouble1(sheetView.Cells[array[1] - 1, column2].Text);
				if (num3 + num4 != 0.0)
				{
					if (array[0] == 152)
					{
						sheetView.Cells[array[0] - 2, column2].Locked = true;
						sheetView.Cells[array[0] - 2, column2].Formula = "W152+W153";
						sheetView.Cells[array[0] - 2, column2].Locked = true;
					}
					else
					{
						sheetView.Cells[array[0] - 2, column2].Locked = true;
						sheetView.Cells[array[0] - 2, column2].Formula = "W232+W233";
						sheetView.Cells[array[0] - 2, column2].Locked = true;
					}
				}
				else
				{
					sheetView.Cells[array[0] - 2, column2].Locked = false;
				}
			}
			string text2 = "R_62;W_62;R_71;W_71;R_75;W_75;R_90;W_90;R_94;W_94;R_97;W_97;R_100;W_100;R_103;W_103;R_111;W_111;R_157;W_157;R_237;W_237";
			string text3 = "I_40;J_40;I_41;J_41";
			if (winFpSpread.ActiveSheetIndex == 0)
			{
				string value = HTKKFile.ColumnIndexToLetters((int)iCol) + "_" + iRow;
				if (text2.IndexOf(value) >= 0 && ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value) > 0.0)
				{
					winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value = ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value) * -1.0;
				}
			}
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
			string text = "R_62;W_62;R_71;W_71;R_75;W_75;R_90;W_90;R_94;W_94;R_97;W_97;R_100;W_100;R_103;W_103;R_111;W_111;R_157;W_157";
			string text2 = "I_40;J_40;I_41;J_41";
			if (winFpSpread.ActiveSheetIndex == 1)
			{
				string value = HTKKFile.ColumnIndexToLetters((int)iCol) + "_" + iRow;
				if (text2.IndexOf(value) >= 0 && ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value) > 0.0)
				{
					winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value = ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value) * -1.0;
				}
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
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("D");
			int num2 = 52;
			if (iCol == num && iRow == num2)
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, 0, sheetView.Cells[num2, num].Value.ToBoolOrDefault());
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private bool getSoKyTruoc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int sheetIndex, bool isChecked)
	{
		try
		{
			if (tKhaiInfoOld.SheetInfos.Count == 0)
			{
				tKhaiInfoOld = ToKhaiInfo;
			}
			string text = HTKKConfig.DATAFILES + "\\" + TTinChung.mst + "\\" + ToKhaiInfo.SheetInfos[sheetIndex].DataFile;
			int num = int.Parse(PeriodInfo.Nam) - 1;
			text = text + "_" + num + "_0101" + num + "_3112" + num;
			text = ((!(PeriodInfo.DonViTinhMa == "NT")) ? (text + "_VND") : (text + "_NT"));
			text = text + "_" + ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			int num2 = 0;
			int num3 = 0;
			int num4 = HTKKFile.ColumnLettersToIndex("AC");
			if (sheetIndex == 0 || sheetIndex == 2 || sheetIndex == 3)
			{
				num2 = HTKKFile.ColumnLettersToIndex("W");
				num3 = HTKKFile.ColumnLettersToIndex("R");
			}
			if (sheetIndex == 1)
			{
				num2 = HTKKFile.ColumnLettersToIndex("J");
				num3 = HTKKFile.ColumnLettersToIndex("I");
			}
			if (File.Exists(text))
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
				hTKKCMCSheetInfo.GetDetailSheetInfo(text);
				tKhaiInfoOld.SheetInfos[sheetIndex] = hTKKCMCSheetInfo;
				for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num3)
					{
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num2 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num4 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
					}
				}
				return true;
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
