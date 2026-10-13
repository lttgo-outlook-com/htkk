using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Sddnn01TT80Nckt251
{
	private int colTenTinh = HTKKFile.ColumnLettersToIndex("C");

	private int colMaTinh = HTKKFile.ColumnLettersToIndex("AO");

	private int colTenHuyen = HTKKFile.ColumnLettersToIndex("F");

	private int colMaHuyen = HTKKFile.ColumnLettersToIndex("AP");

	private int colTenXa = HTKKFile.ColumnLettersToIndex("I");

	private int colMaXa = HTKKFile.ColumnLettersToIndex("AQ");

	private int colTo = HTKKFile.ColumnLettersToIndex("L");

	private int colDuongPho = HTKKFile.ColumnLettersToIndex("N");

	private int colSoNha = HTKKFile.ColumnLettersToIndex("P");

	private int colAJ = HTKKFile.ColumnLettersToIndex("AJ");

	private string colAJStr = "AJ";

	private List<int> listDiaChi = new List<int>();

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("AK_8");
			sView.Cells[array[1], array[0]].Value = PeriodInfo.LanXuatBan;
			listDiaChi.Add(colTenTinh);
			listDiaChi.Add(colMaTinh);
			listDiaChi.Add(colTenXa);
			listDiaChi.Add(colMaXa);
			listDiaChi.Add(colTo);
			listDiaChi.Add(colDuongPho);
			listDiaChi.Add(colSoNha);
			GeneralFunctions.ThayDoiTinhTheoCQTDangBang2Cap(sView, ToKhaiInfo.SheetInfos[0], PeriodInfo, -1, colTenTinh, colMaTinh, colTenHuyen, colMaHuyen, colTenXa, colMaXa);
			GeneralFunctions.SetAllValueInRowSame_Table(sView, ToKhaiInfo.SheetInfos[0].Tables[0], listDiaChi, isLock: true, isDelete: false);
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
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = HTKKFile.ColumnLettersToIndex("V");
			int num2 = HTKKFile.ColumnLettersToIndex("X");
			int num3 = HTKKFile.ColumnLettersToIndex("Z");
			int num4 = HTKKFile.ColumnLettersToIndex("AB");
			int num5 = HTKKFile.ColumnLettersToIndex("AH");
			int num6 = HTKKFile.ColumnLettersToIndex("AR");
			int num7 = HTKKFile.ColumnLettersToIndex("AS");
			if (iCol == num2 || iCol == num3)
			{
				string text = sheetView.Cells[(int)iRow - 1, num6 - 1].Text.Trim();
				string text2 = sheetView.Cells[(int)iRow - 1, num7 - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2))
				{
					double dinhSuatThue = GetDinhSuatThue(text, text2);
					if (dinhSuatThue > 0.0)
					{
						sheetView.Cells[(int)iRow - 1, num4 - 1].Value = dinhSuatThue;
					}
					else
					{
						sheetView.Cells[(int)iRow - 1, num4 - 1].Value = 0;
					}
				}
			}
			if (iCol == colTenTinh || iCol == colTenHuyen || iCol == colTenXa || iCol == colTo || iCol == colDuongPho || iCol == colSoNha)
			{
				GeneralFunctions.SetAllValueInRowSame_Table(sheetView, ToKhaiInfo.SheetInfos[0].Tables[0], listDiaChi, isLock: true, isDelete: false);
			}
			string pathCellCt22 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TinhThue/ct22";
			string pathCellCt23 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TinhThue/ct21";
			string pathCellCt24 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TinhThue/ct20";
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCellCt22));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCellCt23));
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCellCt24));
			if (iCol == num || iCol == num4)
			{
				sheetView.Cells[(int)iRow - 1, num5 - 1].Formula = colAJStr + iRow;
			}
			if (iCol == hTKKCMCCellInfo.CellTemplate.ColIndex && iRow == hTKKCMCCellInfo.CellTemplate.RowIndex)
			{
				for (int num8 = 0; num8 < hTKKCMCTableInfo.RowsValues.Count; num8++)
				{
					sheetView.Cells[dataLocationTemplateRowStart + num8 - 1, num5 - 1].Formula = colAJStr + (dataLocationTemplateRowStart + num8);
				}
			}
			if (iCol == num || iCol == num4 || (iCol == hTKKCMCCellInfo3.CellTemplate.ColIndex && iRow == hTKKCMCCellInfo3.CellTemplate.RowIndex))
			{
				sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Formula = colAJStr + (hTKKCMCCellInfo3.CellTemplate.RowIndex + 2);
				sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, colAJ - 1].Value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("F");
			int num3 = HTKKFile.ColumnLettersToIndex("I");
			int num4 = HTKKFile.ColumnLettersToIndex("R");
			int num5 = HTKKFile.ColumnLettersToIndex("T");
			int num6 = HTKKFile.ColumnLettersToIndex("V");
			int num7 = HTKKFile.ColumnLettersToIndex("X");
			int num8 = HTKKFile.ColumnLettersToIndex("Z");
			int num9 = HTKKFile.ColumnLettersToIndex("AB");
			int[] array = GeneralFunctions.ArrayCellId("M_77");
			int[] array2 = GeneralFunctions.ArrayCellId("M_79");
			int[] array3 = GeneralFunctions.ArrayCellId("M_81");
			int[] array4 = GeneralFunctions.ArrayCellId("M_85");
			int[] array5 = GeneralFunctions.ArrayCellId("AE_56");
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_13"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_28"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29"));
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num10 = hTKKCMCTableInfo.RowsValues.Count - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string value2 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text.Trim();
				if (string.IsNullOrEmpty(value2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num - 1, message);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num - 1);
				}
				string value3 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Text.Trim();
				if (string.IsNullOrEmpty(value3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1, message);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num3 - 1);
				}
				string value4 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Text.Trim();
				if (string.IsNullOrEmpty(value4))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1, message);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num4 - 1);
				}
				string value5 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num5 - 1].Text.Trim();
				if (string.IsNullOrEmpty(value5))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num5 - 1, message);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num5 - 1);
				}
				double num11 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num6 - 1].Value);
				if (num11 == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1, message);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num6 - 1);
				}
				string value6 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num7 - 1].Text.Trim();
				if (string.IsNullOrEmpty(value6))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num7 - 1, message);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num7 - 1);
				}
				string value7 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num8 - 1].Text.Trim();
				if (string.IsNullOrEmpty(value7))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1, message);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num8 - 1);
				}
				double num12 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + i - 1, num9 - 1].Value);
				if (num12 == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + i - 1, num9 - 1, message);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + i - 1, num9 - 1);
				}
			}
			string text = sheetView.Cells[array[1] + num10, array[0]].Text.Trim();
			if (string.IsNullOrEmpty(text))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array[1] + num10, array[0], message);
				value = 2;
			}
			else
			{
				DateTime t = DateTime.ParseExact(text, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				if (DateTime.Compare(t, DateTime.Now) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array[1] + num10, array[0], message3);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, array[1] + num10, array[0]);
				}
			}
			string text2 = sheetView.Cells[array2[1] + num10, array2[0]].Text.Trim();
			if (!string.IsNullOrEmpty(text2))
			{
				DateTime t2 = DateTime.ParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				if (DateTime.Compare(t2, DateTime.Now) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array2[1] + num10, array2[0], message3);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, array2[1] + num10, array2[0]);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, array2[1] + num10, array2[0]);
			}
			double num13 = Convert.ToDouble(sheetView.Cells[array3[1] + num10, array3[0]].Value);
			double num14 = Convert.ToDouble(sheetView.Cells[array5[1], array5[0]].Value);
			if (num13 > num14)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array3[1] + num10, array3[0], message4);
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, array3[1] + num10, array3[0]);
			}
			HTKKFile.RemoveValidate(sheetView, array4[1] + num10, array4[0]);
			double num15 = Convert.ToDouble(sheetView.Cells[array4[1] + num10, array4[0]].Value);
			if (num15 == 0.0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array4[1] + num10, array4[0], message);
				value = 2;
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[0].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[0].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[0].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public double GetDinhSuatThue(string ma_loai_dat, string ma_hang_dat)
	{
		//IL_004a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0050: Expected O, but got Unknown
		string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_Map_ThueSuatDat.xml";
		if (!File.Exists(text))
		{
			return -1.0;
		}
		XmlDocument val = new XmlDocument();
		val.Load(text);
		XmlNodeList val2 = ((XmlNode)val.DocumentElement).SelectNodes("/Root/Map");
		if (val2 != null && val2.Count > 0)
		{
			XmlNode val3 = ((IEnumerable)val2).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["MaLoaiDat"] != null && ((XmlNode)x.Attributes["MaLoaiDat"]).Value == ma_loai_dat));
			if (val3 != null && val3.ChildNodes.Count > 0)
			{
				XmlNode val4 = ((IEnumerable)val3.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["MaHangDat"] != null && ((XmlNode)x.Attributes["MaHangDat"]).Value == ma_hang_dat));
				if (val4 != null && val4.Attributes["ThueSuat"] != null)
				{
					return Convert.ToDouble(((XmlNode)val4.Attributes["ThueSuat"]).Value);
				}
			}
		}
		return -1.0;
	}

	public bool UpdateFormula_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			GeneralFunctions.ThayDoiTinhTheoCQTDangBang(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0], PeriodInfo, -1, colTenTinh, colMaTinh, colTenHuyen, colMaHuyen, colTenXa, colMaXa);
			GeneralFunctions.SetAllValueInRowSame_Table(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0].Tables[0], listDiaChi, isLock: true, isDelete: false);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}
}
