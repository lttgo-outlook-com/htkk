using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BlpTbac03Nckt210
{
	private Dictionary<string, string> dicBLPTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicBLP = new Dictionary<string, string[]>();

	private string pathNguoiKy = "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/nguoiKy";

	private string pathNgayKy = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayKy";

	private string pathNguoiLapBieu = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/nguoiLapBieu";

	private int[] CellId_KinhGui => GeneralFunctions.ArrayCellId("J_2");

	private int[] CellId_PhuongPhapHuy => GeneralFunctions.ArrayCellId("K_6");

	private int[] CellId_ThoiGianTB_Gio => GeneralFunctions.ArrayCellId("L_8");

	private int[] CellId_ThoiGianTB_Phut => GeneralFunctions.ArrayCellId("N_8");

	private int[] CellId_ThoiGianTB_Ngay => GeneralFunctions.ArrayCellId("O_8");

	private int idxCol_MaBienLai => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_MaBienLai_hide => HTKKFile.ColumnLettersToIndex("AI") - 1;

	private int idxCol_TenBienLai => HTKKFile.ColumnLettersToIndex("E") - 1;

	private int idxCol_KyHieuMau => HTKKFile.ColumnLettersToIndex("O") - 1;

	private int idxCol_KyHieuBL => HTKKFile.ColumnLettersToIndex("S") - 1;

	private int idxCol_TuSo => HTKKFile.ColumnLettersToIndex("V") - 1;

	private int idxCol_DenSo => HTKKFile.ColumnLettersToIndex("Y") - 1;

	private int idxCol_SoLuong => HTKKFile.ColumnLettersToIndex("AA") - 1;

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("L");
			int num2 = HTKKFile.ColumnLettersToIndex("N");
			int num3 = HTKKFile.ColumnLettersToIndex("O");
			int num4 = 8;
			if ((HTKKFormater.isNullString(sView.Cells[num4 - 1, num - 1].Text) || sView.Cells[num4 - 1, num - 1].Text == "0") && (HTKKFormater.isNullString(sView.Cells[num4 - 1, num2 - 1].Text) || sView.Cells[num4 - 1, num2 - 1].Text == "0") && HTKKFormater.isNullString(sView.Cells[num4 - 1, num3 - 1].Text))
			{
				sView.Cells[num4 - 1, num - 1].Text = DateTime.Now.Hour.ToString();
				sView.Cells[num4 - 1, num2 - 1].Text = DateTime.Now.Minute.ToString();
				sView.Cells[num4 - 1, num3 - 1].Text = DateTime.Now.Date.ToString("dd/MM/yyyy");
				int num5 = HTKKFile.ColumnLettersToIndex("L");
				int num6 = HTKKFile.ColumnLettersToIndex("N");
				int num7 = HTKKFile.ColumnLettersToIndex("O");
				string text = sView.Cells[num4 - 1, num7 - 1].Text;
				string text2 = sView.Cells[num4 - 1, num5 - 1].Text;
				string text3 = sView.Cells[num4 - 1, num6 - 1].Text;
				string strInput = "00";
				if (text.IndexOf("/") > 0 && !HTKKFormater.isNullString(text2) && !HTKKFormater.isNullString(text3))
				{
					string[] array = text.Split(new char[1] { '/' });
					sView.Cells[GeneralFunctions.ArrayCellId("O_9")[1], GeneralFunctions.ArrayCellId("O_9")[0]].Text = array[2] + "-" + array[1] + "-" + array[0] + "T" + GeneralFunctions.GetStrByLengh(text2, 2) + ":" + GeneralFunctions.GetStrByLengh(text3, 2) + ":" + GeneralFunctions.GetStrByLengh(strInput, 2) + "Z";
				}
				else
				{
					sView.Cells[GeneralFunctions.ArrayCellId("O_9")[1], GeneralFunctions.ArrayCellId("O_9")[0]].Text = "";
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNgayKy));
			string newValue = sView.Cells[CellId_ThoiGianTB_Ngay[1], CellId_ThoiGianTB_Ngay[0]].Text.Trim();
			HTKKFile.AddOrUpdateDicStr(GeneralFunctions.dicVar, "NgayTB03BLP", newValue);
			if (!ToKhaiInfo.SheetInfos[0].IsReadDataFile)
			{
				sView.Cells[CellId_KinhGui[1], CellId_KinhGui[0]].Text = TTinChung.tenCQTNoiNop;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public bool KeyPress(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int idxSheet, int idxRow, int idxCol, KeyPressEventArgs e)
	{
		try
		{
			if (idxSheet == 0)
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
				if (idxRow >= num && idxRow <= num2 && (idxCol == idxCol_TuSo || idxCol == idxCol_DenSo))
				{
					GeneralFunctions.CompareKeycode(e, 1);
				}
			}
			return false;
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			if (iRow >= dataLocationTemplateRowStart && iRow <= num && iCol - 1 == idxCol_MaBienLai)
			{
				string text = sheetView.Cells[(int)iRow - 1, idxCol_MaBienLai].Text;
				if (HTKKFormater.isNullString(text))
				{
					sheetView.Cells[(int)iRow - 1, idxCol_MaBienLai_hide].ResetValue();
					sheetView.Cells[(int)iRow - 1, idxCol_TenBienLai].ResetValue();
					sheetView.Cells[(int)iRow - 1, idxCol_KyHieuMau].ResetValue();
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, idxCol_MaBienLai_hide].Text = text;
					sheetView.Cells[(int)iRow - 1, idxCol_TenBienLai].Text = dicBLP[text][2];
					sheetView.Cells[(int)iRow - 1, idxCol_KyHieuMau].Text = text;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void LeaveCell_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, long iRow, long iCol, long iNewRow, long iNewCol)
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

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
			if (iRow - 1 >= num && iRow - 1 <= num2)
			{
				if (iCol - 1 == idxCol_KyHieuBL)
				{
					string text = sheetView.Cells[(int)iRow - 1, idxCol_KyHieuBL].Text;
					sheetView.Cells[(int)iRow - 1, idxCol_KyHieuBL].Text = GeneralFunctions.FormatKyHieuBLP(text);
				}
				else if (iCol - 1 == idxCol_KyHieuMau)
				{
					string text = sheetView.Cells[(int)iRow - 1, idxCol_KyHieuMau].Text;
					sheetView.Cells[(int)iRow - 1, idxCol_KyHieuMau].Text = text.ToUpper();
				}
				else if (iCol - 1 == idxCol_TuSo || iCol - 1 == idxCol_DenSo)
				{
					string text2 = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
					if (!HTKKFormater.isNullString(text2))
					{
						sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatSoHD(text2);
					}
					string text3 = sheetView.Cells[(int)iRow - 1, idxCol_TuSo].Text;
					string text4 = sheetView.Cells[(int)iRow - 1, idxCol_DenSo].Text;
					if (!HTKKFormater.isNullString(text3) && !HTKKFormater.isNullString(text4))
					{
						double num3 = ConvertHelpers.ConvertToDouble(text3);
						double num4 = ConvertHelpers.ConvertToDouble(text4);
						sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].ResetFormula();
						if (num3 == 0.0 && num4 == 0.0)
						{
							sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].Value = 0;
						}
						else if (num3 == 0.0 && num4 != 0.0)
						{
							sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].Value = num4;
						}
						else
						{
							sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].Value = num4 - num3 + 1.0;
						}
					}
					else if (HTKKFormater.isNullString(text3) && HTKKFormater.isNullString(text4))
					{
						sheetView.Cells[(int)iRow - 1, idxCol_SoLuong].ResetValue();
					}
				}
			}
			SheetView sheetView2 = winFpSpread.Sheets[0];
			int num5 = HTKKFile.ColumnLettersToIndex("L");
			int num6 = HTKKFile.ColumnLettersToIndex("N");
			int num7 = HTKKFile.ColumnLettersToIndex("O");
			int num8 = 8;
			if ((int)iRow != num8 || ((int)iCol != num5 && (int)iCol != num6 && (int)iCol != num7))
			{
				return;
			}
			int num9 = HTKKFile.ColumnLettersToIndex("L");
			int num10 = HTKKFile.ColumnLettersToIndex("N");
			int num11 = HTKKFile.ColumnLettersToIndex("O");
			string text5 = sheetView2.Cells[num8 - 1, num11 - 1].Text;
			string text6 = sheetView2.Cells[num8 - 1, num9 - 1].Text;
			string text7 = sheetView2.Cells[num8 - 1, num10 - 1].Text;
			string strInput = "00";
			if (HTKKFormater.isNullString(text6))
			{
				text6 = "0";
			}
			if (HTKKFormater.isNullString(text7))
			{
				text7 = "0";
			}
			if (!HTKKFormater.isNullString(text5))
			{
				if (text5.IndexOf("/") > 0 && !HTKKFormater.isNullString(text6) && !HTKKFormater.isNullString(text7))
				{
					string[] array = text5.Split(new char[1] { '/' });
					sheetView2.Cells[GeneralFunctions.ArrayCellId("O_9")[1], GeneralFunctions.ArrayCellId("O_9")[0]].Text = array[2] + "-" + array[1] + "-" + array[0] + "T" + GeneralFunctions.GetStrByLengh(text6, 2) + ":" + GeneralFunctions.GetStrByLengh(text7, 2) + ":" + GeneralFunctions.GetStrByLengh(strInput, 2) + "Z";
				}
				else
				{
					sheetView2.Cells[GeneralFunctions.ArrayCellId("O_9")[1], GeneralFunctions.ArrayCellId("O_9")[0]].Text = "";
				}
				HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNgayKy));
				string newValue = sheetView2.Cells[CellId_ThoiGianTB_Ngay[1], CellId_ThoiGianTB_Ngay[0]].Text.Trim();
				HTKKFile.AddOrUpdateDicStr(GeneralFunctions.dicVar, "NgayTB03BLP", newValue);
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
		try
		{
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				int value = item.Value;
				if (value == 2 && winFpSpread.TabStrip[key].BackColor != Color.Pink)
				{
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = 0;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_14"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19"));
			string message7 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20"));
			string message8 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_21"));
			string message9 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22"));
			string message10 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23"));
			string message11 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27"));
			string message12 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_28"));
			string message13 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29"));
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			string[] array = GeneralFunctions.checkTrung(winFpSpread, num2, "V", "Y", "B", "E", "O", "S", "", isMST: false, 1);
			string[] array2 = new string[rowsInfo.Cells.Count];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array2[0] = sheetView.Cells[num2 + i, idxCol_MaBienLai].Text.Trim();
				array2[1] = sheetView.Cells[num2 + i, idxCol_TenBienLai].Text.Trim();
				array2[2] = sheetView.Cells[num2 + i, idxCol_KyHieuMau].Text.Trim();
				array2[3] = sheetView.Cells[num2 + i, idxCol_KyHieuBL].Text.Trim();
				array2[4] = sheetView.Cells[num2 + i, idxCol_TuSo].Text.Trim();
				array2[5] = sheetView.Cells[num2 + i, idxCol_DenSo].Text.Trim();
				if (!GeneralFunctions.HaveToValidateRow(rowsInfo, array2))
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_MaBienLai);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_KyHieuMau);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_TuSo);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_DenSo);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_TenBienLai);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_KyHieuBL);
					continue;
				}
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_MaBienLai);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_KyHieuMau);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_TuSo);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_DenSo);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_TenBienLai);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_KyHieuBL);
				if (HTKKFormater.isNullString(array2[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_MaBienLai, message7);
					num = 2;
				}
				if (HTKKFormater.isNullString(array2[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_TenBienLai, message7);
					num = 2;
				}
				if (HTKKFormater.isNullString(array2[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuMau, message7);
					num = 2;
				}
				else
				{
					string strLoai = (HTKKFormater.isNullString(array2[0]) ? "" : dicBLP[array2[0]][0]);
					string text = GeneralFunctions.CheckMauSoBLP(array2[2], strLoai, array2[0]);
					if (text.Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuMau, message);
						num = 2;
					}
					else if (text.Equals("2"))
					{
						if (array2[3].Length > 0)
						{
							if (array2[3].Substring(array2[3].Length - 1).Equals("E"))
							{
								if (GeneralFunctions.CheckSoLienBLPDT(array2[2], array2[0], array2[3]).Equals("3"))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuMau, message13);
									num = 2;
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuMau, message3);
								num = 2;
							}
						}
					}
					else if (array2[3].Length > 0 && array2[3].Substring(array2[3].Length - 1).Equals("E") && GeneralFunctions.CheckSoLienBLPDT(array2[2], array2[0], array2[3]).Equals("3"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuMau, message13);
						num = 2;
					}
				}
				if (HTKKFormater.isNullString(array2[3]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuBL, message7);
					num = 2;
				}
				else
				{
					string text2 = GeneralFunctions.CheckSoBLP(array2[3], array2[0], DateTime.ParseExact(PeriodInfo.kyKKhai, "dd/MM/yyyy", CultureInfo.InvariantCulture) >= new DateTime(2017, 1, 1));
					if (HTKKFormater.isNullString(text2) || text2.Equals("1") || text2.Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_KyHieuBL, message2);
						num = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_KyHieuBL);
					}
				}
				bool flag = false;
				if (HTKKFormater.isNullString(array2[4]) && HTKKFormater.isNullString(array2[5]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_TuSo, message7);
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_DenSo, message7);
					num = 2;
					flag = true;
				}
				else if (HTKKFormater.isNullString(array2[4]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_TuSo, message7);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_DenSo);
					num = 2;
					flag = true;
				}
				else if (HTKKFormater.isNullString(array2[5]))
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_TuSo);
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_DenSo, message7);
					num = 2;
					flag = true;
				}
				else
				{
					double num3 = ConvertHelpers.ConvertToDouble(array2[4]);
					double num4 = ConvertHelpers.ConvertToDouble(array2[5]);
					if (num3 > num4)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_TuSo, message6);
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_DenSo, message6);
						num = 2;
						flag = true;
					}
				}
				if (flag || array == null || array.Length <= 1)
				{
					continue;
				}
				for (int j = 0; j < array.Length; j++)
				{
					if (num2 + i == ConvertHelpers.ConvertToInt(array[j]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_TuSo, message4);
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_DenSo, message4);
						num = 2;
					}
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNgayKy));
			string text3 = sheetView.Cells[CellId_ThoiGianTB_Ngay[1], CellId_ThoiGianTB_Ngay[0]].Text.Trim();
			string text4 = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
			DateTime t = new DateTime(2011, 1, 1);
			bool flag2 = true;
			DateTime dtDate;
			if (HTKKFormater.isNullString(text3))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ThoiGianTB_Ngay[1], CellId_ThoiGianTB_Ngay[0], message7);
				num = 2;
			}
			else if (!HTKKFormater.IsDateTime(text3, "dd/MM/yyyy", out dtDate))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ThoiGianTB_Ngay[1], CellId_ThoiGianTB_Ngay[0], HTKKFile.GetMessage("0403"));
				num = 2;
			}
			else if (DateTime.Compare(dtDate, t) < 0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ThoiGianTB_Ngay[1], CellId_ThoiGianTB_Ngay[0], message11);
				num = 2;
			}
			else if (DateTime.Compare(dtDate, DateTime.Now) > 0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ThoiGianTB_Ngay[1], CellId_ThoiGianTB_Ngay[0], message8);
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ThoiGianTB_Ngay[1], CellId_ThoiGianTB_Ngay[0]);
				flag2 = false;
			}
			if (HTKKFormater.isNullString(text4))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message7);
				num = 2;
			}
			else
			{
				HTKKFormater.IsDateTime(text4, "dd/MM/yyyy", out var dtDate2);
				if (DateTime.Compare(dtDate2, DateTime.Now) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message9);
					num = 2;
				}
				else if (!flag2)
				{
					HTKKFormater.IsDateTime(text3, "dd/MM/yyyy", out dtDate);
					if (DateTime.Compare(dtDate2, dtDate) < 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message12);
						num = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNguoiKy));
			if (hTKKCMCCellInfo2 != null)
			{
				string str = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
				if (HTKKFormater.isNullString(str))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, message7);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNguoiLapBieu));
			if (hTKKCMCCellInfo3 != null)
			{
				string str2 = sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text.Trim();
				if (HTKKFormater.isNullString(str2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, message7);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1);
				}
			}
			string str3 = sheetView.Cells[CellId_PhuongPhapHuy[1], CellId_PhuongPhapHuy[0]].Text.Trim();
			if (HTKKFormater.isNullString(str3))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_PhuongPhapHuy[1], CellId_PhuongPhapHuy[0], message7);
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_PhuongPhapHuy[1], CellId_PhuongPhapHuy[0]);
			}
			HTKKFile.AddOrUpdateDicStr(GeneralFunctions.dicVar, "NgayTB03BLP", text3);
			switch (num)
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
			throw ex;
		}
	}

	public Dictionary<string, string> GetDMBienLaiPhi(string pathFile, HTKKPeriodInfo PeriodInfo)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0065: Expected O, but got Unknown
		try
		{
			if (dicBLPTmp.Count > 0)
			{
				return dicBLPTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[1];
				string value2 = array[1];
				if (!dictionary2.ContainsKey(key))
				{
					dictionary2.Add(key, value2);
				}
				if (!dictionary.ContainsKey(key))
				{
					dictionary.Add(key, array);
				}
			}
			dicBLP = dictionary;
			dicBLPTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("L");
			int num3 = HTKKFile.ColumnLettersToIndex("N");
			int num4 = HTKKFile.ColumnLettersToIndex("O");
			int[] array = GeneralFunctions.ArrayCellId("O_9");
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int dataLocationTemplateRowStart = hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables[0].RowsValues.Count; i++)
			{
				string[] array2 = hTKKCMCSheetInfo.Tables[0].RowsValues[i].RowValue.ToArray();
				if (array2 == null || HTKKFormater.isNullString(array2[1]))
				{
					continue;
				}
				string text = "";
				foreach (KeyValuePair<string, string[]> item in dicBLP)
				{
					if (item.Value.Length <= 0 || !(item.Value[2] == array2[1]))
					{
						continue;
					}
					text = item.Key;
					break;
				}
				winFpSpread.Sheets[0].Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text = text;
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayHuyBienLai"));
			string text2 = "";
			if (hTKKCMCCellInfo != null)
			{
				text2 = hTKKCMCCellInfo.Value;
			}
			if (!HTKKFormater.isNullString(text2))
			{
				winFpSpread.Sheets[0].Cells[array[1], array[0]].Text = text2;
				string text3 = "";
				string text4 = "";
				string text5 = "";
				string text6 = "";
				if (text2.IndexOf("T") > -1 && text2.IndexOf(":") > -1)
				{
					string[] array3 = text2.Split(new char[1] { 'T' });
					string text7 = array3[0].Replace("-", "");
					text6 = text7.Substring(6, 2) + "/" + text7.Substring(4, 2) + "/" + text7.Substring(0, 4);
					text3 = array3[1].Split(new char[1] { ':' })[0];
					text4 = array3[1].Split(new char[1] { ':' })[1];
					text5 = array3[1].Split(new char[1] { ':' })[2].Substring(0, 2);
					winFpSpread.Sheets[0].Cells[7, num4 - 1].Text = text6;
					winFpSpread.Sheets[0].Cells[7, num2 - 1].Text = text3;
					winFpSpread.Sheets[0].Cells[7, num3 - 1].Text = text4;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ExportXML_0(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals("path2"));
			xmlNode.SelectSingleNode("HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhai").InnerText = hTKKCMCCellInfo.Value;
			xmlNode.SelectSingleNode("HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhaiTuNgay").InnerText = hTKKCMCCellInfo.Value;
			xmlNode.SelectSingleNode("HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhaiDenNgay").InnerText = hTKKCMCCellInfo.Value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
