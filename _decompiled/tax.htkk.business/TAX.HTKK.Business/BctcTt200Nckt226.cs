using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BctcTt200Nckt226
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
				xmlNode.SelectSingleNode("HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/ngayKy").InnerText = xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/CDKT_HoatDongLienTuc/ngayLap").InnerText;
				xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/hoatDongLienTuc").InnerText = "true";
				xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/hoatDongKhongLienTuc").InnerText = "false";
				XmlNode val = xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/CDKT_HoatDongKhongLienTuc");
				xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh").RemoveChild(val);
			}
			else
			{
				xmlNode.SelectSingleNode("HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/ngayKy").InnerText = xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/CDKT_HoatDongKhongLienTuc/ngayLap").InnerText;
				xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/hoatDongLienTuc").InnerText = "false";
				xmlNode.SelectSingleNode("HSoKhaiThue/CTieuTKhaiChinh/hoatDongKhongLienTuc").InnerText = "true";
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
				if (PeriodInfo.DoanhNghiepHoatDongLienTuc == "LT")
				{
					if (num == 0)
					{
						sheetView.Cells[4, HTKKFile.ColumnLettersToIndex("T") - 1].Text = "Mẫu số B 01 - DN";
					}
					if (num == 1)
					{
						sheetView.Cells[0, HTKKFile.ColumnLettersToIndex("AA") - 1].Text = "Mẫu số B 02 - DN";
					}
					if (num == 2)
					{
						sheetView.Cells[0, HTKKFile.ColumnLettersToIndex("T") - 1].Text = "Mẫu số B 03 - DN";
					}
					if (num == 3)
					{
						sheetView.Cells[0, HTKKFile.ColumnLettersToIndex("T") - 1].Text = "Mẫu số B 03 - DN";
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
			winFpSpread.ActiveSheetIndex = 0;
			if (PeriodInfo.DoanhNghiepHoatDongLienTuc == "LT")
			{
				sheetView.Cells[44, HTKKFile.ColumnLettersToIndex("K") - 1].Text = "1";
				sheetView.Cells[44, HTKKFile.ColumnLettersToIndex("L") - 1].Text = "0";
				for (int i = rowBeginKLT; i < rowEndKLT; i++)
				{
					Row row = winFpSpread.ActiveSheet.Rows[i];
					row.Visible = false;
				}
				winFpSpread.ActiveSheet.Rows[2].Visible = false;
			}
			else if (PeriodInfo.DoanhNghiepHoatDongLienTuc == "KL")
			{
				sheetView.Cells[44, HTKKFile.ColumnLettersToIndex("K") - 1].Text = "";
				sheetView.Cells[44, HTKKFile.ColumnLettersToIndex("L") - 1].Text = "1";
				for (int i = rowBeginLT; i < rowEndLT; i++)
				{
					Row row = winFpSpread.ActiveSheet.Rows[i];
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
			int[] array2 = GeneralFunctions.ArrayCellId("D_51");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			sheetView.Cells[50, HTKKFile.ColumnLettersToIndex("T") - 1].Locked = !Convert.ToBoolean(valueCustom);
			int[] array3 = GeneralFunctions.ArrayCellId("D_52");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array3[1] + 1, array3[0] + 1, "2");
			if (ConvertHelpers.ConvertToBool(valueCustom2))
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 0, ToKhaiInfo.SheetInfos[0].IsReadDataFile, resetValue: false);
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
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex >= 22 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex <= 41 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != HTKKFile.ColumnLettersToIndex("H"))
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
			}
			else
			{
				for (int i = 0; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex >= 22 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex <= 41 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != HTKKFile.ColumnLettersToIndex("H"))
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
			int[] array = GeneralFunctions.ArrayCellId("B_17");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (ConvertHelpers.ConvertToBool(valueCustom))
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 1, ToKhaiInfo.SheetInfos[1].IsReadDataFile, resetValue: false);
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
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex >= 44 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex <= 72 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != HTKKFile.ColumnLettersToIndex("P"))
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
			}
			else
			{
				for (int i = 0; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex >= 44 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex <= 72 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != HTKKFile.ColumnLettersToIndex("P"))
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
			int[] array = GeneralFunctions.ArrayCellId("C_38");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (ConvertHelpers.ConvertToBool(valueCustom))
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 2, ToKhaiInfo.SheetInfos[1].IsReadDataFile, resetValue: false);
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
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex >= 48 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex <= 87 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != HTKKFile.ColumnLettersToIndex("P"))
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
			}
			else
			{
				for (int i = 0; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex >= 48 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex <= 87 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != HTKKFile.ColumnLettersToIndex("P"))
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
			int[] array = GeneralFunctions.ArrayCellId("C_42");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (ConvertHelpers.ConvertToBool(valueCustom))
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 3, ToKhaiInfo.SheetInfos[1].IsReadDataFile, resetValue: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		//IL_0109: Unknown result type (might be due to invalid IL or missing references)
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
			if (ToKhaiInfo.SheetInfos[2] != null)
			{
				Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			}
			if (ToKhaiInfo.SheetInfos[3] != null)
			{
				Validate_3(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			}
			if (winFpSpread.Sheets[2].Visible && winFpSpread.Sheets[3].Visible)
			{
				winFpSpread.TabStrip[2].BackColor = Color.Pink;
				winFpSpread.TabStrip[3].BackColor = Color.Pink;
				frmMessage frmMessage = new frmMessage("0395", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
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
			int num6 = 51;
			int num7 = HTKKFile.ColumnLettersToIndex("T");
			int num8 = HTKKFile.ColumnLettersToIndex("D");
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[num6 - 1, num8 - 1].Value))
			{
				if (string.IsNullOrEmpty(sheetView.Cells[num6 - 1, num7 - 1].Text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num6 - 1, num7 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_40")));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num6 - 1, num7 - 1);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, num6 - 1, num7 - 1);
			}
			int[] array = GeneralFunctions.ArrayCellId("D_52");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 0, Convert.ToBoolean(valueCustom2));
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
			int index = 1;
			SheetView sView = winFpSpread.Sheets[index];
			int[] array = GeneralFunctions.ArrayCellId("B_17");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 1, Convert.ToBoolean(valueCustom));
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
			SheetView sView = winFpSpread.Sheets[index];
			int[] array = GeneralFunctions.ArrayCellId("C_39");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 2, Convert.ToBoolean(valueCustom));
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
			SheetView sView = winFpSpread.Sheets[index];
			int[] array = GeneralFunctions.ArrayCellId("C_43");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 3, Convert.ToBoolean(valueCustom));
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
			int[] array = GeneralFunctions.ArrayCellId("D_51");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				sheetView.Cells[50, HTKKFile.ColumnLettersToIndex("T") - 1].Locked = !Convert.ToBoolean(valueCustom);
				if (!Convert.ToBoolean(valueCustom))
				{
					sheetView.Cells[50, HTKKFile.ColumnLettersToIndex("T") - 1].Text = "";
				}
			}
			int[] array2 = GeneralFunctions.ArrayCellId("D_52");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (idxCol == array2[0] && idxRow == array2[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 0, isReadDataFile: false, !Convert.ToBoolean(valueCustom2));
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
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 1, isReadDataFile: false, !Convert.ToBoolean(valueCustom));
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
			int[] array = GeneralFunctions.ArrayCellId("C_39");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 2, isReadDataFile: false, !Convert.ToBoolean(valueCustom));
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[3];
			int[] array = GeneralFunctions.ArrayCellId("C_43");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 3, isReadDataFile: false, !Convert.ToBoolean(valueCustom));
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void CellChange_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			string text = "R_65;W_65";
			if (winFpSpread.ActiveSheetIndex == 3)
			{
				string value = HTKKFile.ColumnIndexToLetters((int)iCol) + "_" + iRow;
				if (text.IndexOf(value) >= 0 && ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value) > 0.0)
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void getSoKyTruoc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex, bool isReadDataFile, bool resetValue)
	{
		try
		{
			List<string> pathDataFileOld = getPathDataFileOld(ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, sheetIndex);
			string text = pathDataFileOld[0];
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			int num = 0;
			int num2 = 0;
			int num3 = 0;
			if (sheetIndex == 0 || sheetIndex == 2 || sheetIndex == 3)
			{
				num = HTKKFile.ColumnLettersToIndex("R");
				num2 = HTKKFile.ColumnLettersToIndex("W");
				num3 = HTKKFile.ColumnLettersToIndex("AI");
			}
			if (sheetIndex == 1)
			{
				num = HTKKFile.ColumnLettersToIndex("I");
				num2 = HTKKFile.ColumnLettersToIndex("J");
				num3 = HTKKFile.ColumnLettersToIndex("N");
			}
			if (File.Exists(text) && !resetValue)
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
			string text4 = "";
			text4 = ((!(PeriodInfo.DonViTinhMa == "NT")) ? "VND" : "NT");
			if (HTKKFormater.isNullString(PeriodInfo.TuNgay) || HTKKFormater.isNullString(PeriodInfo.DenNgay))
			{
				list.Add(text);
				return list;
			}
			text2 = PeriodInfo.TuNgay.Split(new char[1] { '/' })[0] + PeriodInfo.TuNgay.Split(new char[1] { '/' })[1] + (ConvertHelpers.ConvertToInt(PeriodInfo.TuNgay.Split(new char[1] { '/' })[2]) - 1);
			text3 = PeriodInfo.DenNgay.Split(new char[1] { '/' })[0] + PeriodInfo.DenNgay.Split(new char[1] { '/' })[1] + (ConvertHelpers.ConvertToInt(PeriodInfo.DenNgay.Split(new char[1] { '/' })[2]) - 1);
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

	private void checkSoNamTruoc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex, bool isCheck)
	{
		try
		{
			int num = 0;
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			List<string> pathDataFileOld = getPathDataFileOld(ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, sheetIndex);
			int num2 = 0;
			int num3 = 0;
			int num4 = 0;
			if (sheetIndex == 0 || sheetIndex == 2 || sheetIndex == 3)
			{
				num2 = HTKKFile.ColumnLettersToIndex("R");
				num3 = HTKKFile.ColumnLettersToIndex("W");
				num4 = HTKKFile.ColumnLettersToIndex("AI");
				if (pathDataFileOld == null || (pathDataFileOld != null && !File.Exists(pathDataFileOld[0])))
				{
					return;
				}
			}
			if (sheetIndex == 1)
			{
				num2 = HTKKFile.ColumnLettersToIndex("I");
				num3 = HTKKFile.ColumnLettersToIndex("J");
				num4 = HTKKFile.ColumnLettersToIndex("N");
			}
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[52, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[sheetIndex];
			for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
			{
				if (!File.Exists(pathDataFileOld[0]) || hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != num3 || sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Locked)
				{
					continue;
				}
				string valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num3, "NumberCellType");
				object value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num4 - 1].Value;
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
			if ((num != 2 || !(winFpSpread.TabStrip[sheetIndex].BackColor == Color.Pink)) && (num != 1 || !(winFpSpread.TabStrip[sheetIndex].BackColor == Color.Yellow)))
			{
				if (num == 2 && winFpSpread.TabStrip[sheetIndex].BackColor != Color.Pink)
				{
					winFpSpread.TabStrip[sheetIndex].BackColor = Color.Pink;
				}
				else if (num == 1 && winFpSpread.TabStrip[sheetIndex].BackColor == Color.White)
				{
					winFpSpread.TabStrip[sheetIndex].BackColor = Color.Yellow;
				}
				else if (num == 0)
				{
					winFpSpread.TabStrip[sheetIndex].BackColor = Color.White;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateThongTinChung(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			string text = "251;253;255;257";
			for (int i = 0; i < hTKKCMCSheetInfo.StaticCells.Count - 1; i++)
			{
				if (text.IndexOf(hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex.ToString()) <= -1)
				{
					continue;
				}
				if (PeriodInfo.DoanhNghiepHoatDongLienTuc == "LT")
				{
					if (hTKKCMCSheetInfo.StaticCells[i].Path.IndexOf("CDKT_HoatDongLienTuc") > -1)
					{
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex - 1].Text = hTKKCMCSheetInfo.StaticCells[i].Value;
					}
				}
				else if (hTKKCMCSheetInfo.StaticCells[i].Path.IndexOf("CDKT_HoatDongKhongLienTuc") > -1)
				{
					sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex - 1].Text = hTKKCMCSheetInfo.StaticCells[i].Value;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			UpdateThongTinChung(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
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
			UpdateThongTinChung(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
