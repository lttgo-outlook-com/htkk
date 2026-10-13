using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using Newtonsoft.Json.Linq;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TkttIir01Nckt288
{
	private int colC = HTKKFile.ColumnLettersToIndex("C");

	private int colD = HTKKFile.ColumnLettersToIndex("D");

	private int[] cell_TenDonViTien = GeneralFunctions.ArrayCellId("AS_37");

	private int[] cell_MaDonViTien = GeneralFunctions.ArrayCellId("FS_37");

	private int[] cell_1_1_co = GeneralFunctions.ArrayCellId("E_55");

	private int[] cell_1_1_khong = GeneralFunctions.ArrayCellId("E_57");

	private int[] cell_1_2_1_ct2 = GeneralFunctions.ArrayCellId("AH_67");

	private int[] cell_1_2_1_ct3 = GeneralFunctions.ArrayCellId("AT_67");

	private string pathMuc_1_3_2_1_co = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_1/Muc_1.3.2.1/thayDoiBCTC_Co";

	private string pathMuc_1_3_2_1_khong = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_1/Muc_1.3.2.1/thayDoiBCTC_Khong";

	private string pathMuc_1_3_2_1_d_co = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_1/Muc_1.3.2.1/ThongTinQuyDinhChiuThue/ChiTietQuyDinhChiuThue/giamTruTrachNhiem_Co";

	private string pathMuc_1_3_2_1_d_khong = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_1/Muc_1.3.2.1/ThongTinQuyDinhChiuThue/ChiTietQuyDinhChiuThue/giamTruTrachNhiem_Khong";

	private string pathMuc_1_3_2_2_co = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_1/Muc_1.3.2.2/thayDoiBCTC_Co";

	private string pathMuc_1_3_2_2_khong = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_1/Muc_1.3.2.2/thayDoiBCTC_Khong";

	private string pathMuc_1_3_3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_1/Muc_1.3.3/thayDoi_AnhHuong";

	private string path_Muc_2_3_ct4 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.3/soQuocGia_ThueTTTC";

	private string path_Muc_2_3_ct6 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.3/soQuocGia";

	private string path_Muc_3_8_3_TongCt3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.8.3/tongSoLuongNhanVien";

	private string path_Muc_3_8_3_TongCt5 = "path_tong_ct5";

	private string tablePath_Muc_1_3_2_1_b = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_1/Muc_1.3.2.1/ThongTinDVHopThanh/ChiTietTTDVHopThanh";

	private string tablePath_Muc_1_3_2_1_d = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_1/Muc_1.3.2.1/ThongTinQuyDinhChiuThue/ChiTietQuyDinhChiuThue/ChiTietTTQuyDinhChiuThue";

	private string tablePath_Muc_1_3_3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_1/Muc_1.3.3/ChiTietThayDoi";

	private string tablePath_Muc_2_1_2_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.1.2.1/ChiTietQDGiamTru";

	private string tablePath_Muc_2_1_2_2 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.1.2.2/ChiTietGiamTruTN";

	private string tablePath_Muc_3_1_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.1.1/ChiTietXacDinhTN_Lo";

	private string tablePath_Muc_3_2_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.2.1/ChiTietTongSoThueDC";

	private string tablePath_Muc_3_8_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.8.1/Bang1/ChiTietBang1";

	private string tablePath_Muc_3_2_4_1_a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.2.4.1/TomTatTongQuat/ChiTietTomTatTongQuat";

	private string tablePath_Muc_3_2_4_1_b = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.2.4.1/PhanTichDieuChinh/ChiTietPhanTichDieuChinh";

	private string tablePath_Muc_3_2_4_2 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.2.4.2/ChiTietCoCheThuHoi";

	private string tablePath_Muc_3_2_4_3_a1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.2.4.3/Muc_a.1/ChiTietKTHoanLai";

	private string tablePath_Muc_3_2_4_3_a2 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.2.4.3/Muc_a.2/ChiTietTSHoanLai";

	private string tablePath_Muc_3_2_4_3_b = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.2.4.3/Muc_b/ChiTietPhaiTra";

	private string tablePath_Muc_3_3_3_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.3.3/Muc_3.3.3.1/ChiTietThuHoi";

	private string tablePath_Muc_3_4 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.4/ChiTietApDung";

	private string tablePath_Muc_3_5_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.5.1/ChiTietSoThueBS";

	private string tablePath_Muc_3_6 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.6/ChiTietSoThueBoSungTT";

	private string tablePath_Muc_3_7 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.7/ChiTietTinhToanThueBs";

	private string tablePath_Muc_3_8_3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.8.3/ChiTietKhauTruThue";

	private string tablePath_Muc_4_1_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_4/Muc_4.1.1/ChiTietDanhSach";

	private string tablePath_Muc_4_1_2 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_4/Muc_4.1.2/ChiTietPhanBoCSTT";

	private string tablePath_Muc_4_1_3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_4/Muc_4.1.3/ChiTietPhanBoDVTC";

	private string tablePath_Muc_4_1_5 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_4/Muc_4.1.5/ChiTietDieuChinhTNQD";

	private string tablePath_Muc_4_2_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_4/Muc_4.2.1/ChiTietDieuChinhCP";

	private string tablePath_Muc_4_2_3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_4/Muc_4.2.3/ChiTietDieuChinhHoanLai";

	private string tablePath_Muc_4_3_3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_4/Muc_4.3.3/ChiTietLuaChonDieuChinh";

	private string path_Muc_3_2_4_3_NamChuyenTiep = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.2.4.3/nam_ChuyenTiep";

	private int[] cellMaCt9 = GeneralFunctions.ArrayCellId("AA_26");

	private int[] cellCt10_KhacVnd_Pl_01_1 = GeneralFunctions.ArrayCellId("G_29");

	private int[] cellCt10_Vnd_Pl_01_1 = GeneralFunctions.ArrayCellId("M_29");

	private int[] cellCt11_TiGia_Pl_01_1 = GeneralFunctions.ArrayCellId("C_31");

	private string tablePath_Muc_II = "TKhai_01_TNDN_IIR/Muc_II/so_ThueTNDN/ChiTiet";

	private string sumCt10_Path_Muc_II = "TKhai_01_TNDN_IIR/Muc_II/so_ThueTNDN/tong_SoThuePhanBo";

	private string ct17PlPath = "TKhai_01_TNDN_IIR/Muc_III/ct17";

	private string ct18PlPath = "TKhai_01_TNDN_IIR/Muc_III/ct18";

	private string pathTongHopKhbs = "path_tong_hop_khbs";

	private string pathTenDonViTienKhbs = "path_ten_don_vi_tien_khbs";

	private string pathMaDonViTienKhbs = "path_ma_don_vi_tien_khbs";

	private JArray mapAnHienBang = null;

	private string tablePath_Pl_01_2 = "TKhai_01_TM/ChiTiet";

	private string pathKhbs_A_I_3_b = "PL01_KHBS/Muc_A/Muc_I/Muc_3/ct3b";

	private List<MapCongThuc> mapCongThuc = new List<MapCongThuc>();

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("Z");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			sheetView.Cells[cell_1_2_1_ct2[1], cell_1_2_1_ct2[0]].Text = PeriodInfo.TuNgay.Substring(0, 5);
			sheetView.Cells[cell_1_2_1_ct3[1], cell_1_2_1_ct3[0]].Text = PeriodInfo.DenNgay.Substring(0, 5);
			string json = File.ReadAllText("MapAnHienCotTrongBang.json");
			JObject jObject = JObject.Parse(json);
			mapAnHienBang = (JArray)jObject[ToKhaiInfo.MenuID];
			GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
			mapCongThuc = GeneralFunctions.MapCongThucChoSua(PeriodInfo.menuID);
			ThayDoiDonViTien_Tk(sheetView, hTKKCMCSheetInfo);
			if (PeriodInfo.LoaiToKhai == "B")
			{
				GeneralFunctions.LockUnLockCol(sheetView, cell_TenDonViTien[1] + 1, cell_TenDonViTien[0] + 1, locked: true, deleteData: false);
			}
			GeneralFunctions.lockUnlockColsInTableByFile(sheetView, hTKKCMCSheetInfo, -1, -1, mapAnHienBang);
			LoadCongThucCacBangTheoFileTk(sheetView, hTKKCMCSheetInfo);
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_Muc_3_2_4_3_NamChuyenTiep);
			string text = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text;
			if (string.IsNullOrEmpty(text) || text != PeriodInfo.Nam)
			{
				lockUnlockTableNamChuyenTiep_3_2_4_3(sheetView, hTKKCMCSheetInfo, PeriodInfo);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LoadCongThucCacBangTheoFileTk(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_3_8_1);
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		GeneralFunctions.sumInTableByFile(sView, sInfo, dataLocationTemplateRowStart, HTKKFile.ColumnLettersToIndex("D"), mapAnHienBang);
		HTKKCMCTableInfo hTKKCMCTableInfo2 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_4_1_1);
		int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
		GeneralFunctions.sumInTableByFile(sView, sInfo, dataLocationTemplateRowStart2, HTKKFile.ColumnLettersToIndex("D"), mapAnHienBang);
		HTKKCMCTableInfo hTKKCMCTableInfo3 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_4_2_1);
		int dataLocationTemplateRowStart3 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart;
		GeneralFunctions.sumInTableByFile(sView, sInfo, dataLocationTemplateRowStart3, HTKKFile.ColumnLettersToIndex("D"), mapAnHienBang);
		HTKKCMCTableInfo hTKKCMCTableInfo4 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_4_2_3);
		int dataLocationTemplateRowStart4 = hTKKCMCTableInfo4.Location.DataLocationTemplateRowStart;
		GeneralFunctions.sumInTableByFile(sView, sInfo, dataLocationTemplateRowStart4, HTKKFile.ColumnLettersToIndex("D"), mapAnHienBang);
	}

	public void AutoBindData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			ThayDoiDonViTien_Pl_01_1(winFpSpread, ToKhaiInfo);
			LockUnLock_Pl_01_1(winFpSpread, ToKhaiInfo, -1, -1);
			string text = sheetView.Cells[cellMaCt9[1], cellMaCt9[0]].Text;
			bool flag = Convert.ToBoolean(sheetView.Cells[cellCt10_Vnd_Pl_01_1[1], cellCt10_Vnd_Pl_01_1[0]].Value);
			if (PeriodInfo.LoaiToKhai == "B")
			{
				GeneralFunctions.LockUnLockCol(sheetView, cellCt10_KhacVnd_Pl_01_1[1] + 1, cellCt10_KhacVnd_Pl_01_1[0] + 1, locked: true, deleteData: false);
				GeneralFunctions.LockUnLockCol(sheetView, cellCt10_Vnd_Pl_01_1[1] + 1, cellCt10_Vnd_Pl_01_1[0] + 1, locked: true, deleteData: false);
			}
			if (PeriodInfo.LoaiToKhai == "B")
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathTongHopKhbs);
				if (text != "VND" && flag)
				{
					GeneralFunctions.UpdateFileMapMctKhbs(PeriodInfo.menuID, "MCTPath", "path_tong_hop_khbs", "MCT", "[18]");
					GeneralFunctions.UpdateFileMapMctKhbs(PeriodInfo.menuID, "MCTPath", "path_tong_hop_khbs", "Caption", "Tổng số thuế phân bổ cho công ty mẹ từ số thuế bổ sung của các đơn vị hợp thành chịu thuế suất thấp sau khi bù trừ nghĩa vụ thuế theo đồng Việt Nam trong trường hợp khai Tờ khai thông tin bằng đồng ngoại tệ và lựa chọn khai Tờ khai thuế TNDN bổ sung, nộp thuế TNDN bổ sung bằng đồng Việt Nam (nếu có)");
				}
				else
				{
					GeneralFunctions.UpdateFileMapMctKhbs(PeriodInfo.menuID, "MCTPath", "path_tong_hop_khbs", "MCT", "[17]");
					GeneralFunctions.UpdateFileMapMctKhbs(PeriodInfo.menuID, "MCTPath", "path_tong_hop_khbs", "Caption", "Tổng số thuế phân bổ cho công ty mẹ tại Việt Nam từ số thuế bổ sung của các đơn vị hợp thành chịu thuế suất thấp sau khi bù trừ nghĩa vụ thuế theo quy định IIR");
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			ThayDoiDonViTien_Pl_01_2(winFpSpread, ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int index = ToKhaiInfo.SheetInfos.Count - 1;
			SheetView sheetView2 = winFpSpread.Sheets[index];
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[index];
			int index2 = ToKhaiInfo.SheetInfos.Count - 2;
			SheetView sheetView3 = winFpSpread.Sheets[index2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo3 = ToKhaiInfo.SheetInfos[index2];
			if (!(PeriodInfo.LoaiToKhai == "B"))
			{
				return;
			}
			if (PeriodInfo.LoaiToKhai == "B")
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo3.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathKhbs_A_I_3_b);
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMaDonViTienKhbs);
				string text = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text;
				if (hTKKCMCCellInfo != null)
				{
					NumberCellType numberCellType = (NumberCellType)sheetView3.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].CellType;
					numberCellType.DecimalPlaces = ((!(text == "VND")) ? 2 : 0);
					sheetView3.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].CellType = numberCellType;
				}
			}
			int[] array = GeneralFunctions.ArrayCellId("BQ_39");
			int[] array2 = GeneralFunctions.ArrayCellId("AA_37");
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathTenDonViTienKhbs);
			string text2 = sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text;
			sheetView2.Cells[array[1], array[0]].Value = text2;
			sheetView3.Cells[array2[1], array2[0]].Value = text2;
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sView = winFpSpread.Sheets[0];
			if (iRow == cell_MaDonViTien[1] + 1)
			{
				ThayDoiDonViTien_Tk(sView, hTKKCMCSheetInfo);
				ThayDoiDonViTien_Pl_01_1(winFpSpread, ToKhaiInfo);
				ThayDoiDonViTien_Pl_01_2(winFpSpread, ToKhaiInfo);
				LockUnLock_Pl_01_1(winFpSpread, ToKhaiInfo, (int)iRow, (int)iCol);
			}
			if (iCol == colD)
			{
				GeneralFunctions.lockUnlockColsInTableByFile(sView, hTKKCMCSheetInfo, (int)iRow, (int)iCol, mapAnHienBang);
			}
			TinhCongThucChoSua_Tk(sView, hTKKCMCSheetInfo, (int)iRow, (int)iCol);
			DayBangTuToKhaiSangPl_01_1_Table3_8_1(winFpSpread, ToKhaiInfo, (int)iRow, (int)iCol);
			DayBangTu_4_1_1Sang_3_1_1(winFpSpread, ToKhaiInfo, (int)iRow, (int)iCol);
			DayBangTu_4_2_1Sang_3_2_1(winFpSpread, ToKhaiInfo, (int)iRow, (int)iCol);
			DayBangTu_4_2_3Sang_3_2_4_1_b(winFpSpread, ToKhaiInfo, (int)iRow, (int)iCol);
			List<int> list = new List<int>();
			list.Add(HTKKFile.ColumnLettersToIndex("BI"));
			list.Add(HTKKFile.ColumnLettersToIndex("BO"));
			list.Add(HTKKFile.ColumnLettersToIndex("BT"));
			list.Add(HTKKFile.ColumnLettersToIndex("BZ"));
			list.Add(HTKKFile.ColumnLettersToIndex("CF"));
			ThemNhieuDongNamTaiChinh_3_2_4_2(sView, hTKKCMCSheetInfo, (int)iRow, (int)iCol, tablePath_Muc_3_2_4_2, list);
			List<int> list2 = new List<int>();
			list2.Add(HTKKFile.ColumnLettersToIndex("BI"));
			list2.Add(HTKKFile.ColumnLettersToIndex("BO"));
			list2.Add(HTKKFile.ColumnLettersToIndex("BT"));
			list2.Add(HTKKFile.ColumnLettersToIndex("BZ"));
			list2.Add(HTKKFile.ColumnLettersToIndex("CF"));
			ThemNhieuDongNamTaiChinh_3_3_3_1(sView, hTKKCMCSheetInfo, (int)iRow, (int)iCol, tablePath_Muc_3_3_3_1, list2);
			sua1DongCacDongKhacCungLoaiAnTheo(sView, hTKKCMCSheetInfo, (int)iRow, (int)iCol, tablePath_Muc_3_2_4_2, "FS", "02", "W", "FU");
			sua1DongCacDongKhacCungLoaiAnTheo(sView, hTKKCMCSheetInfo, (int)iRow, (int)iCol, tablePath_Muc_3_2_4_2, "FS", "02", "AG", "");
			sua1DongCacDongKhacCungLoaiAnTheo(sView, hTKKCMCSheetInfo, (int)iRow, (int)iCol, tablePath_Muc_3_3_3_1, "FS", "02", "W", "FU");
			sua1DongCacDongKhacCungLoaiAnTheo(sView, hTKKCMCSheetInfo, (int)iRow, (int)iCol, tablePath_Muc_3_3_3_1, "FS", "02", "AG", "");
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_Muc_3_2_4_3_NamChuyenTiep);
			if (iRow == hTKKCMCCellInfo.CellTemplate.RowIndex)
			{
				lockUnlockTableNamChuyenTiep_3_2_4_3(sView, hTKKCMCSheetInfo, PeriodInfo);
			}
			if (iCol == colD)
			{
				GeneralFunctions.sumInTableByFile(sView, hTKKCMCSheetInfo, (int)iRow, (int)iCol, mapAnHienBang);
			}
			TinhCongThuc_Tk_4_1_2(sView, hTKKCMCSheetInfo, (int)iRow, (int)iCol);
			TinhCongThuc_Tk_4_1_3(sView, hTKKCMCSheetInfo, (int)iRow, (int)iCol);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[1];
			SheetView sView = winFpSpread.Sheets[1];
			if (iCol == colC)
			{
				GeneralFunctions.lockUnlockColsInTableByFile(sView, sInfo, (int)iRow, (int)iCol, mapAnHienBang);
			}
			TinhCongThucChoSua_Pl01_1(sView, sInfo, (int)iRow, (int)iCol);
			if (iCol == colC)
			{
				GeneralFunctions.sumInTableByFile(sView, sInfo, (int)iRow, (int)iCol, mapAnHienBang);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[2];
			SheetView sView = winFpSpread.Sheets[2];
			TinhCongThucChoSua_Pl01_2(sView, sInfo, (int)iRow, (int)iCol);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SheetTabChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int sheetIndex)
	{
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		checkTKChinh(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL_01_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		if (winFpSpread.TabStrip[0].BackColor == Color.Pink)
		{
			return false;
		}
		if (winFpSpread.TabStrip[1].BackColor == Color.Pink)
		{
			return false;
		}
		if (winFpSpread.TabStrip[2].BackColor == Color.Pink)
		{
			return false;
		}
		return true;
	}

	private void checkTKChinh(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
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

	private void checkPL_01_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			int num = HTKKFile.ColumnLettersToIndex("E");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[1];
			SheetView sheetView = winFpSpread.Sheets[0];
			SheetView sheetView2 = winFpSpread.Sheets[1];
			switch (value)
			{
			case 2:
				sheetView2.FpSpread.TabStrip[1].BackColor = Color.Pink;
				break;
			case 1:
				sheetView2.FpSpread.TabStrip[1].BackColor = Color.Yellow;
				break;
			default:
				sheetView2.FpSpread.TabStrip[1].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL_01_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(2, out value);
			int num = HTKKFile.ColumnLettersToIndex("E");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[2];
			SheetView sheetView = winFpSpread.Sheets[0];
			SheetView sheetView2 = winFpSpread.Sheets[2];
			string message = HTKKFile.GetMessage(sheetView2, GeneralFunctions.ArrayCellId("E_103"));
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo2.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Pl_01_2);
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == "ff");
			int dataLocationTemplateRowStart = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			List<string> list = new List<string>();
			for (int num2 = 0; num2 < hTKKCMCTableInfo2.RowsValues.Count; num2++)
			{
				string text = sheetView.Cells[dataLocationTemplateRowStart + num2 - 1, HTKKFile.ColumnLettersToIndex("Z") - 1].Text;
				if (!string.IsNullOrEmpty(text))
				{
					list.Add(text);
				}
			}
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int num2 = 0; num2 < hTKKCMCTableInfo.RowsValues.Count; num2++)
			{
				string text2 = sheetView2.Cells[dataLocationTemplateRowStart2 + num2 - 1, num - 1].Text;
				if (!string.IsNullOrEmpty(text2))
				{
					if (list.Count == 0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView2, dataLocationTemplateRowStart2 + num2 - 1, num - 1, message);
						value = 2;
					}
					else if (!list.Contains(text2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView2, dataLocationTemplateRowStart2 + num2 - 1, num - 1, message);
						value = 2;
					}
				}
			}
			switch (value)
			{
			case 2:
				sheetView2.FpSpread.TabStrip[2].BackColor = Color.Pink;
				break;
			case 1:
				sheetView2.FpSpread.TabStrip[2].BackColor = Color.Yellow;
				break;
			default:
				sheetView2.FpSpread.TabStrip[2].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public int checkValidateTable1ColOk(SheetView sView, HTKKCMCTableInfo table, List<int> colsCheck, string mess)
	{
		int result = 0;
		int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
		for (int i = 0; i < table.RowsValues.Count; i++)
		{
			bool flag = true;
			foreach (HTKKCMCTableCellInfo cell in table.RowsInfo.Cells)
			{
				try
				{
					if (!colsCheck.Contains(cell.CellTemplate.ColIndex))
					{
						continue;
					}
					HTKKFile.RemoveValidate(sView, dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1);
					if (cell.Controltype == "16")
					{
						double num = Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1].Value);
						if (num != Convert.ToDouble(cell.DefaultValue))
						{
							flag = false;
						}
					}
					else
					{
						string text = sView.Cells[dataLocationTemplateRowStart + i - 1, cell.CellTemplate.ColIndex - 1].Text;
						if (text != cell.DefaultValue)
						{
							flag = false;
						}
					}
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
				}
			}
			if (!flag)
			{
				continue;
			}
			foreach (HTKKCMCTableCellInfo cell2 in table.RowsInfo.Cells)
			{
				if (colsCheck.Contains(cell2.CellTemplate.ColIndex))
				{
					HTKKFile.SetNodeBackGroundPink(sView, dataLocationTemplateRowStart + i - 1, cell2.CellTemplate.ColIndex - 1, mess);
				}
			}
			result = 2;
		}
		return result;
	}

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sView = winFpSpread.Sheets[0];
			int iColClick = winFpSpread.ActiveSheet.ActiveColumnIndex + 1;
			int iRowClick = winFpSpread.ActiveSheet.ActiveRowIndex + 1;
			resetOtherCheckboxTK(sView, sInfo, iRowClick, iColClick);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int iCol = winFpSpread.ActiveSheet.ActiveColumnIndex + 1;
			int iRow = winFpSpread.ActiveSheet.ActiveRowIndex + 1;
			LockUnLock_Pl_01_1(winFpSpread, ToKhaiInfo, iRow, iCol);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_1(FpSpread winFpSpread, int actionType)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void resetOtherCheckboxTK(SheetView sView, HTKKCMCSheetInfo sInfo, int iRowClick, int iColClick)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc_1_3_2_1_co);
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc_1_3_2_1_khong);
		HTKKCMCCellInfo hTKKCMCCellInfo3 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc_1_3_2_1_d_co);
		HTKKCMCCellInfo hTKKCMCCellInfo4 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc_1_3_2_1_d_khong);
		HTKKCMCCellInfo hTKKCMCCellInfo5 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc_1_3_2_2_co);
		HTKKCMCCellInfo hTKKCMCCellInfo6 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc_1_3_2_2_khong);
		HTKKCMCCellInfo hTKKCMCCellInfo7 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc_1_3_3);
		if (iRowClick == cell_1_1_co[1] + 1 || iRowClick == cell_1_1_khong[1] + 1)
		{
			GeneralFunctions.resetOtherCheckboxStatic(sView, sInfo, cell_1_1_co[1] + 1, cell_1_1_khong[1] + 1, iRowClick, iColClick);
		}
		if (iRowClick == hTKKCMCCellInfo.CellTemplate.RowIndex || iRowClick == hTKKCMCCellInfo2.CellTemplate.RowIndex)
		{
			GeneralFunctions.resetOtherCheckboxStatic(sView, sInfo, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo2.CellTemplate.RowIndex, iRowClick, iColClick);
		}
		if (iRowClick == hTKKCMCCellInfo3.CellTemplate.RowIndex || iRowClick == hTKKCMCCellInfo4.CellTemplate.RowIndex)
		{
			GeneralFunctions.resetOtherCheckboxStatic(sView, sInfo, hTKKCMCCellInfo3.CellTemplate.RowIndex, hTKKCMCCellInfo4.CellTemplate.RowIndex, iRowClick, iColClick);
		}
		if (iRowClick == hTKKCMCCellInfo5.CellTemplate.RowIndex || iRowClick == hTKKCMCCellInfo6.CellTemplate.RowIndex)
		{
			GeneralFunctions.resetOtherCheckboxStatic(sView, sInfo, hTKKCMCCellInfo5.CellTemplate.RowIndex, hTKKCMCCellInfo6.CellTemplate.RowIndex, iRowClick, iColClick);
		}
		if (iRowClick == hTKKCMCCellInfo7.CellTemplate.RowIndex)
		{
			GeneralFunctions.resetOtherCheckboxStatic(sView, sInfo, hTKKCMCCellInfo7.CellTemplate.RowIndex, hTKKCMCCellInfo7.CellTemplate.RowIndex, iRowClick, iColClick);
		}
		HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_4_3_3);
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		if (iRowClick >= dataLocationTemplateRowStart && iRowClick <= dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count)
		{
			List<int> list = new List<int>();
			list.Add(HTKKFile.ColumnLettersToIndex("BC"));
			list.Add(HTKKFile.ColumnLettersToIndex("BH"));
			GeneralFunctions.resetOtherCheckboxDinamic(sView, sInfo, iRowClick, iColClick, list);
		}
	}

	public void changFormatCellDecimalPlacesStatic(SheetView sView, HTKKCMCSheetInfo sInfo, int decimalPlaces, List<int[]> cellsBoQua)
	{
		try
		{
			List<HTKKCMCCellInfo> list = sInfo.StaticCells.FindAll((HTKKCMCCellInfo x) => x.Controltype == "16");
			foreach (HTKKCMCCellInfo item in list)
			{
				int[] target = new int[2]
				{
					item.CellTemplate.RowIndex,
					item.CellTemplate.ColIndex
				};
				List<int[]> list2 = cellsBoQua.Where((int[] a) => a.SequenceEqual(target)).ToList();
				if (list2 == null || list2.Count == 0)
				{
					NumberCellType numberCellType = (NumberCellType)sView.Cells[item.CellTemplate.RowIndex - 1, item.CellTemplate.ColIndex - 1].CellType;
					numberCellType.DecimalPlaces = decimalPlaces;
					sView.Cells[item.CellTemplate.RowIndex - 1, item.CellTemplate.ColIndex - 1].CellType = numberCellType;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void changFormatCellDecimalPlacesDinamic(SheetView sView, HTKKCMCTableInfo table, int decimalPlaces, List<int> colsBoQua)
	{
		try
		{
			List<HTKKCMCTableCellInfo> list = table.RowsInfo.Cells.FindAll((HTKKCMCTableCellInfo x) => !colsBoQua.Contains(x.CellTemplate.ColIndex) && x.Controltype == "16");
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			for (int num = 0; num < table.RowsValues.Count; num++)
			{
				foreach (HTKKCMCTableCellInfo item in list)
				{
					NumberCellType numberCellType = (NumberCellType)sView.Cells[dataLocationTemplateRowStart + num - 1, item.CellTemplate.ColIndex - 1].CellType;
					numberCellType.DecimalPlaces = decimalPlaces;
					sView.Cells[dataLocationTemplateRowStart + num - 1, item.CellTemplate.ColIndex - 1].CellType = numberCellType;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ThayDoiDonViTien_Tk(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		string text = sView.Cells[cell_MaDonViTien[1], cell_MaDonViTien[0]].Text;
		int decimalPlaces = ((!string.IsNullOrEmpty(text) && !(text == "VND")) ? 4 : 0);
		List<int[]> list = new List<int[]>();
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_Muc_2_3_ct4);
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_Muc_2_3_ct6);
		HTKKCMCCellInfo hTKKCMCCellInfo3 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_Muc_3_8_3_TongCt3);
		HTKKCMCCellInfo hTKKCMCCellInfo4 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_Muc_3_8_3_TongCt5);
		list.Add(new int[2]
		{
			hTKKCMCCellInfo.CellTemplate.RowIndex,
			hTKKCMCCellInfo.CellTemplate.ColIndex
		});
		list.Add(new int[2]
		{
			hTKKCMCCellInfo2.CellTemplate.RowIndex,
			hTKKCMCCellInfo2.CellTemplate.ColIndex
		});
		list.Add(new int[2]
		{
			hTKKCMCCellInfo3.CellTemplate.RowIndex,
			hTKKCMCCellInfo3.CellTemplate.ColIndex
		});
		list.Add(new int[2]
		{
			hTKKCMCCellInfo4.CellTemplate.RowIndex,
			hTKKCMCCellInfo4.CellTemplate.ColIndex
		});
		changFormatCellDecimalPlacesStatic(sView, sInfo, decimalPlaces, list);
		List<int> list2 = new List<int>();
		foreach (HTKKCMCTableInfo table in sInfo.Tables)
		{
			list2.Clear();
			HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_1_3_2_1_b);
			HTKKCMCTableInfo hTKKCMCTableInfo2 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_1_3_2_1_d);
			HTKKCMCTableInfo hTKKCMCTableInfo3 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_1_3_3);
			HTKKCMCTableInfo hTKKCMCTableInfo4 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_2_1_2_1);
			HTKKCMCTableInfo hTKKCMCTableInfo5 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_2_1_2_2);
			HTKKCMCTableInfo hTKKCMCTableInfo6 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_3_4);
			HTKKCMCTableInfo hTKKCMCTableInfo7 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_3_5_1);
			HTKKCMCTableInfo hTKKCMCTableInfo8 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_3_6);
			HTKKCMCTableInfo hTKKCMCTableInfo9 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_3_7);
			if (table.TablePath == hTKKCMCTableInfo.TablePath)
			{
				list2.Add(HTKKFile.ColumnLettersToIndex("CB"));
			}
			if (table.TablePath == hTKKCMCTableInfo2.TablePath)
			{
				list2.Add(HTKKFile.ColumnLettersToIndex("AU"));
			}
			if (table.TablePath == hTKKCMCTableInfo3.TablePath)
			{
				list2.Add(HTKKFile.ColumnLettersToIndex("CI"));
				list2.Add(HTKKFile.ColumnLettersToIndex("CO"));
			}
			if (table.TablePath == hTKKCMCTableInfo4.TablePath)
			{
				list2.Add(HTKKFile.ColumnLettersToIndex("BT"));
			}
			if (table.TablePath == hTKKCMCTableInfo5.TablePath)
			{
				list2.Add(HTKKFile.ColumnLettersToIndex("W"));
			}
			if (table.TablePath == hTKKCMCTableInfo6.TablePath)
			{
				list2.Add(HTKKFile.ColumnLettersToIndex("BR"));
			}
			if (table.TablePath == hTKKCMCTableInfo7.TablePath)
			{
				list2.Add(HTKKFile.ColumnLettersToIndex("BT"));
				list2.Add(HTKKFile.ColumnLettersToIndex("BY"));
				list2.Add(HTKKFile.ColumnLettersToIndex("CO"));
				list2.Add(HTKKFile.ColumnLettersToIndex("CP"));
			}
			if (table.TablePath == hTKKCMCTableInfo8.TablePath)
			{
				list2.Add(HTKKFile.ColumnLettersToIndex("BE"));
			}
			if (table.TablePath == hTKKCMCTableInfo9.TablePath)
			{
				list2.Add(HTKKFile.ColumnLettersToIndex("AQ"));
			}
			if (table.TablePath == tablePath_Muc_3_8_1)
			{
				list2.Add(HTKKFile.ColumnLettersToIndex("CF"));
			}
			if (table.TablePath == tablePath_Muc_3_8_3)
			{
				list2.Add(HTKKFile.ColumnLettersToIndex("V"));
				list2.Add(HTKKFile.ColumnLettersToIndex("AJ"));
			}
			if (table.TablePath == tablePath_Muc_4_1_5)
			{
				list2.Add(HTKKFile.ColumnLettersToIndex("CO"));
				list2.Add(HTKKFile.ColumnLettersToIndex("CC"));
			}
			changFormatCellDecimalPlacesDinamic(sView, table, decimalPlaces, list2);
		}
	}

	public void ThayDoiDonViTien_Pl_01_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		SheetView sheetView2 = winFpSpread.Sheets[1];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
		string text = sheetView.Cells[cell_MaDonViTien[1], cell_MaDonViTien[0]].Text;
		int decimalPlaces = ((!string.IsNullOrEmpty(text) && !(text == "VND")) ? 4 : 0);
		HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_II);
		List<int[]> list = new List<int[]>();
		HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == ct17PlPath);
		HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == ct18PlPath);
		list.Add(new int[2]
		{
			hTKKCMCCellInfo.CellTemplate.RowIndex,
			hTKKCMCCellInfo.CellTemplate.ColIndex
		});
		list.Add(new int[2]
		{
			hTKKCMCCellInfo2.CellTemplate.RowIndex,
			hTKKCMCCellInfo2.CellTemplate.ColIndex
		});
		changFormatCellDecimalPlacesStatic(sheetView2, hTKKCMCSheetInfo, decimalPlaces, list);
		NumberCellType numberCellType = (NumberCellType)sheetView2.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].CellType;
		numberCellType.DecimalPlaces = ((!(text == "VND")) ? 2 : 0);
		sheetView2.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].CellType = numberCellType;
		HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathTongHopKhbs);
		HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMaDonViTienKhbs);
		string text2 = sheetView2.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text;
		NumberCellType numberCellType2 = (NumberCellType)sheetView2.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].CellType;
		numberCellType2.DecimalPlaces = ((!(text2 == "VND")) ? 2 : 0);
		sheetView2.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].CellType = numberCellType2;
		List<int> list2 = new List<int>();
		foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
		{
			list2.Clear();
			if (table.TablePath == tablePath_Muc_II)
			{
				list2.Add(HTKKFile.ColumnLettersToIndex("S"));
			}
			changFormatCellDecimalPlacesDinamic(sheetView2, table, decimalPlaces, list2);
		}
	}

	public void ThayDoiDonViTien_Pl_01_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		SheetView sView = winFpSpread.Sheets[2];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
		string text = sheetView.Cells[cell_MaDonViTien[1], cell_MaDonViTien[0]].Text;
		int decimalPlaces = ((!string.IsNullOrEmpty(text) && !(text == "VND")) ? 4 : 0);
		List<int> colsBoQua = new List<int>();
		HTKKCMCTableInfo table = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Pl_01_2);
		changFormatCellDecimalPlacesDinamic(sView, table, decimalPlaces, colsBoQua);
	}

	public void TinhCongThucChoSua_Tk(SheetView sView, HTKKCMCSheetInfo sInfo, int iRow, int iCol)
	{
		HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_3_2_4_1_a);
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		if (iRow >= dataLocationTemplateRowStart && iRow <= dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count && (iCol == HTKKFile.ColumnLettersToIndex("BI") || iCol == HTKKFile.ColumnLettersToIndex("BQ")))
		{
			sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("AZ") - 1].Value = Convert.ToDouble(sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("BI") - 1].Value) + Convert.ToDouble(sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("BQ") - 1].Value);
		}
		HTKKCMCTableInfo hTKKCMCTableInfo2 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_3_2_4_3_a2);
		int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
		if (iRow >= dataLocationTemplateRowStart2 && iRow <= dataLocationTemplateRowStart2 + hTKKCMCTableInfo2.RowsValues.Count && iCol >= HTKKFile.ColumnLettersToIndex("AQ"))
		{
			double num = Convert.ToDouble(sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("AQ") - 1].Value);
			double num2 = Convert.ToDouble(sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("AX") - 1].Value);
			double num3 = Convert.ToDouble(sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("BE") - 1].Value);
			if (num2 != 0.0)
			{
				sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("BL") - 1].Value = num2 - num3;
			}
			else
			{
				sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("BL") - 1].Value = num - num3;
			}
		}
		HTKKCMCTableInfo hTKKCMCTableInfo3 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_3_8_1);
		int dataLocationTemplateRowStart3 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart;
		if (iRow >= dataLocationTemplateRowStart3 && iRow <= dataLocationTemplateRowStart3 + hTKKCMCTableInfo3.RowsValues.Count && (iCol == HTKKFile.ColumnLettersToIndex("D") || (iCol >= HTKKFile.ColumnLettersToIndex("BR") && iCol <= HTKKFile.ColumnLettersToIndex("CF"))))
		{
			string text = sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("FS") - 1].Text;
			double num4 = Convert.ToDouble(sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("BR") - 1].Value);
			double num5 = Convert.ToDouble(sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("CF") - 1].Value);
			if (text == "04")
			{
				sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("CN") - 1].Value = Math.Round(num4 * (num5 / 100.0), 4);
			}
		}
		IEnumerable<HTKKCMCCellInfo> enumerable = sInfo.StaticCells.Where((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex == iRow);
		bool flag = false;
		foreach (HTKKCMCCellInfo staticCell in enumerable)
		{
			List<MapCongThuc> list = mapCongThuc;
			Predicate<MapCongThuc> match = (MapCongThuc x) => x.MCTPath == staticCell.Path;
			if (list.FindIndex(match) >= 0)
			{
				flag = true;
				break;
			}
		}
		if (flag)
		{
			GeneralFunctions.CapNhatCongThucLoaiChoSuaCell(sView, sInfo, iRow, iCol, mapCongThuc);
			return;
		}
		int num6 = GeneralFunctions.DetectIndexOfTable(sInfo, iRow);
		if (num6 >= 0)
		{
			GeneralFunctions.CapNhatCongThucLoaiChoSuaCell(sView, sInfo, iRow, iCol, mapCongThuc);
		}
	}

	public void TinhCongThuc_Tk_4_1_2(SheetView sView, HTKKCMCSheetInfo sInfo, int iRow, int iCol)
	{
		HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_4_1_2);
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		if (iRow < dataLocationTemplateRowStart || iRow > dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count || (iCol != HTKKFile.ColumnLettersToIndex("D") && iCol != HTKKFile.ColumnLettersToIndex("AX") && iCol != HTKKFile.ColumnLettersToIndex("CC")))
		{
			return;
		}
		string text = sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("FS") - 1].Text;
		double num = Convert.ToDouble(sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("AX") - 1].Value);
		double num2 = Convert.ToDouble(sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("CC") - 1].Value);
		if (text == "03")
		{
			sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("CL") - 1].Value = num - num2;
		}
		if (!(text == "04"))
		{
			return;
		}
		sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("CL") - 1].Value = 0;
		for (int num3 = iRow - 1; num3 >= dataLocationTemplateRowStart; num3--)
		{
			string text2 = sView.Cells[num3 - 1, HTKKFile.ColumnLettersToIndex("FS") - 1].Text;
			if (text2 == "03")
			{
				double num4 = Convert.ToDouble(sView.Cells[num3 - 1, HTKKFile.ColumnLettersToIndex("AX") - 1].Value);
				double num5 = Convert.ToDouble(sView.Cells[num3 - 1, HTKKFile.ColumnLettersToIndex("CC") - 1].Value);
				sView.Cells[num3 - 1, HTKKFile.ColumnLettersToIndex("CL") - 1].Value = num4 - num5;
				break;
			}
		}
	}

	public void TinhCongThuc_Tk_4_1_3(SheetView sView, HTKKCMCSheetInfo sInfo, int iRow, int iCol)
	{
		HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_4_1_3);
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		if (iRow < dataLocationTemplateRowStart || iRow > dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count || (iCol != HTKKFile.ColumnLettersToIndex("D") && iCol != HTKKFile.ColumnLettersToIndex("AX") && iCol != HTKKFile.ColumnLettersToIndex("CC")))
		{
			return;
		}
		string text = sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("FS") - 1].Text;
		double num = Convert.ToDouble(sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("AX") - 1].Value);
		double num2 = Convert.ToDouble(sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("CC") - 1].Value);
		if (text == "03")
		{
			sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("CL") - 1].Value = num - num2;
		}
		if (!(text == "04"))
		{
			return;
		}
		sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("CL") - 1].Value = 0;
		for (int num3 = iRow - 1; num3 >= dataLocationTemplateRowStart; num3--)
		{
			string text2 = sView.Cells[num3 - 1, HTKKFile.ColumnLettersToIndex("FS") - 1].Text;
			if (text2 == "03")
			{
				double num4 = Convert.ToDouble(sView.Cells[num3 - 1, HTKKFile.ColumnLettersToIndex("AX") - 1].Value);
				double num5 = Convert.ToDouble(sView.Cells[num3 - 1, HTKKFile.ColumnLettersToIndex("CC") - 1].Value);
				sView.Cells[num3 - 1, HTKKFile.ColumnLettersToIndex("CL") - 1].Value = num4 - num5;
				break;
			}
		}
	}

	public void TinhCongThucChoSua_Pl01_1(SheetView sView, HTKKCMCSheetInfo sInfo, int iRow, int iCol)
	{
	}

	public void TinhCongThucChoSua_Pl01_2(SheetView sView, HTKKCMCSheetInfo sInfo, int iRow, int iCol)
	{
		if (iCol == HTKKFile.ColumnLettersToIndex("T") || iCol == HTKKFile.ColumnLettersToIndex("X"))
		{
			GeneralFunctions.CapNhatCongThucLoaiChoSuaCell(sView, sInfo, iRow, iCol, mapCongThuc);
		}
	}

	public void LockUnLock_Pl_01_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int iRow, int iCol)
	{
		try
		{
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[1];
			SheetView sheetView = winFpSpread.Sheets[0];
			SheetView sheetView2 = winFpSpread.Sheets[1];
			string text = sheetView.Cells[cell_MaDonViTien[1], cell_MaDonViTien[0]].Text;
			bool flag = true;
			bool flag2 = false;
			if (iRow > 0 && iCol > 0)
			{
				if (iRow == cellCt10_KhacVnd_Pl_01_1[1] + 1 && iCol == cellCt10_KhacVnd_Pl_01_1[0] + 1 && Convert.ToBoolean(sheetView2.Cells[cellCt10_KhacVnd_Pl_01_1[1], cellCt10_KhacVnd_Pl_01_1[0]].Value))
				{
					sheetView2.Cells[cellCt10_Vnd_Pl_01_1[1], cellCt10_Vnd_Pl_01_1[0]].Value = false;
					flag2 = true;
					flag = false;
				}
				else if (iRow == cellCt10_Vnd_Pl_01_1[1] + 1 && iCol == cellCt10_Vnd_Pl_01_1[0] + 1 && Convert.ToBoolean(sheetView2.Cells[cellCt10_Vnd_Pl_01_1[1], cellCt10_Vnd_Pl_01_1[0]].Value))
				{
					sheetView2.Cells[cellCt10_KhacVnd_Pl_01_1[1], cellCt10_KhacVnd_Pl_01_1[0]].Value = false;
					flag = true;
					flag2 = false;
				}
			}
			else
			{
				flag = Convert.ToBoolean(sheetView2.Cells[cellCt10_Vnd_Pl_01_1[1], cellCt10_Vnd_Pl_01_1[0]].Value);
				flag2 = Convert.ToBoolean(sheetView2.Cells[cellCt10_KhacVnd_Pl_01_1[1], cellCt10_KhacVnd_Pl_01_1[0]].Value);
			}
			if (cellCt10_KhacVnd_Pl_01_1 != null)
			{
				if (!string.IsNullOrEmpty(text) && text == "VND")
				{
					sheetView2.Cells[cellCt10_Vnd_Pl_01_1[1], cellCt10_Vnd_Pl_01_1[0]].Value = true;
					sheetView2.Cells[cellCt10_KhacVnd_Pl_01_1[1], cellCt10_KhacVnd_Pl_01_1[0]].Value = false;
					GeneralFunctions.LockUnLockCol(sheetView2, cellCt10_KhacVnd_Pl_01_1[1] + 1, cellCt10_KhacVnd_Pl_01_1[0] + 1, locked: true, deleteData: false);
					GeneralFunctions.LockUnLockCol(sheetView2, cellCt10_Vnd_Pl_01_1[1] + 1, cellCt10_Vnd_Pl_01_1[0] + 1, locked: true, deleteData: false);
				}
				else
				{
					GeneralFunctions.LockUnLockCol(sheetView2, cellCt10_KhacVnd_Pl_01_1[1] + 1, cellCt10_KhacVnd_Pl_01_1[0] + 1, locked: false, deleteData: false);
					GeneralFunctions.LockUnLockCol(sheetView2, cellCt10_Vnd_Pl_01_1[1] + 1, cellCt10_Vnd_Pl_01_1[0] + 1, locked: false, deleteData: false);
				}
			}
			if (cellCt11_TiGia_Pl_01_1 != null)
			{
				if (!string.IsNullOrEmpty(text) && text != "VND" && flag)
				{
					GeneralFunctions.lockUnlockMultiCellsStatic(sheetView2, sInfo, cellCt11_TiGia_Pl_01_1[1] + 1, cellCt11_TiGia_Pl_01_1[1] + 1, isLock: false, isDeleteData: false);
				}
				else
				{
					GeneralFunctions.lockUnlockMultiCellsStatic(sheetView2, sInfo, cellCt11_TiGia_Pl_01_1[1] + 1, cellCt11_TiGia_Pl_01_1[1] + 1, isLock: true, isDeleteData: true);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void DayBangTuToKhaiSangPl_01_1_Table3_8_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int iRow, int iCol)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		SheetView sView = winFpSpread.Sheets[1];
		HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[1];
		HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_3_8_1);
		HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo2.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_II);
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
		if (iRow >= dataLocationTemplateRowStart && iRow <= dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count)
		{
			List<int> list = new List<int>();
			list.Add(HTKKFile.ColumnLettersToIndex("D"));
			list.Add(HTKKFile.ColumnLettersToIndex("FS"));
			list.Add(HTKKFile.ColumnLettersToIndex("O"));
			list.Add(HTKKFile.ColumnLettersToIndex("W"));
			list.Add(HTKKFile.ColumnLettersToIndex("AC"));
			list.Add(HTKKFile.ColumnLettersToIndex("FT"));
			list.Add(HTKKFile.ColumnLettersToIndex("BE"));
			list.Add(HTKKFile.ColumnLettersToIndex("AX"));
			list.Add(HTKKFile.ColumnLettersToIndex("BR"));
			list.Add(HTKKFile.ColumnLettersToIndex("BY"));
			list.Add(HTKKFile.ColumnLettersToIndex("BL"));
			list.Add(HTKKFile.ColumnLettersToIndex("CF"));
			list.Add(HTKKFile.ColumnLettersToIndex("CN"));
			List<int> list2 = new List<int>();
			list2.Add(HTKKFile.ColumnLettersToIndex("C"));
			list2.Add(HTKKFile.ColumnLettersToIndex("AA"));
			list2.Add(HTKKFile.ColumnLettersToIndex("E"));
			list2.Add(HTKKFile.ColumnLettersToIndex("D"));
			list2.Add(HTKKFile.ColumnLettersToIndex("I"));
			list2.Add(HTKKFile.ColumnLettersToIndex("AB"));
			list2.Add(HTKKFile.ColumnLettersToIndex("N"));
			list2.Add(HTKKFile.ColumnLettersToIndex("O"));
			list2.Add(HTKKFile.ColumnLettersToIndex("P"));
			list2.Add(HTKKFile.ColumnLettersToIndex("Q"));
			list2.Add(HTKKFile.ColumnLettersToIndex("R"));
			list2.Add(HTKKFile.ColumnLettersToIndex("S"));
			list2.Add(HTKKFile.ColumnLettersToIndex("T"));
			List<string> list3 = new List<string>();
			list3.Add("01");
			list3.Add("02");
			list3.Add("04");
			GeneralFunctions.TongHopTuBangSangBangTheoLoai(winFpSpread, ToKhaiInfo, 0, 1, tablePath_Muc_3_8_1, tablePath_Muc_II, list, list2, list3, HTKKFile.ColumnLettersToIndex("FS"));
			Pl_01_1_Table3_8_1_SumCt10(sView, hTKKCMCSheetInfo2);
		}
	}

	public void Pl_01_1_Table3_8_1_SumCt10(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_II);
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == sumCt10_Path_Muc_II);
		double num = 0.0;
		for (int num2 = 0; num2 < hTKKCMCTableInfo.RowsValues.Count; num2++)
		{
			string text = sView.Cells[dataLocationTemplateRowStart + num2 - 1, HTKKFile.ColumnLettersToIndex("AA") - 1].Text;
			if (text == "02")
			{
				num += Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + num2 - 1, HTKKFile.ColumnLettersToIndex("T") - 1].Value);
			}
		}
		sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = num;
	}

	public void DayBangTu_4_1_1Sang_3_1_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int iRow, int iCol)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_3_1_1);
		HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_4_1_1);
		int dataLocationTemplateRowStart = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
		int dataLocationTemplateRowStart2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		if (iRow < dataLocationTemplateRowStart || iRow > dataLocationTemplateRowStart + hTKKCMCTableInfo2.RowsValues.Count)
		{
			return;
		}
		List<int> list = new List<int>();
		foreach (HTKKCMCTableCellInfo cell in hTKKCMCTableInfo2.RowsInfo.Cells)
		{
			if (cell.CellTemplate.ColIndex != HTKKFile.ColumnLettersToIndex("AI") && cell.CellTemplate.ColIndex != HTKKFile.ColumnLettersToIndex("AO"))
			{
				list.Add(cell.CellTemplate.ColIndex);
			}
		}
		List<int> list2 = new List<int>();
		foreach (HTKKCMCTableCellInfo cell2 in hTKKCMCTableInfo.RowsInfo.Cells)
		{
			list2.Add(cell2.CellTemplate.ColIndex);
		}
		List<string> list3 = new List<string>();
		list3.Add("01");
		list3.Add("02");
		GeneralFunctions.TongHopTuBangSangBangTheoLoai(winFpSpread, ToKhaiInfo, 0, 0, tablePath_Muc_4_1_1, tablePath_Muc_3_1_1, list, list2, list3, HTKKFile.ColumnLettersToIndex("FS"));
		winFpSpread.Sheets[0].ActiveRowIndex = iRow;
	}

	public void DayBangTu_4_2_1Sang_3_2_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int iRow, int iCol)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_3_2_1);
		HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_4_2_1);
		int dataLocationTemplateRowStart = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
		int dataLocationTemplateRowStart2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		if (iRow < dataLocationTemplateRowStart || iRow > dataLocationTemplateRowStart + hTKKCMCTableInfo2.RowsValues.Count)
		{
			return;
		}
		List<int> list = new List<int>();
		foreach (HTKKCMCTableCellInfo cell in hTKKCMCTableInfo2.RowsInfo.Cells)
		{
			if (cell.CellTemplate.ColIndex != HTKKFile.ColumnLettersToIndex("AI") && cell.CellTemplate.ColIndex != HTKKFile.ColumnLettersToIndex("AR"))
			{
				list.Add(cell.CellTemplate.ColIndex);
			}
		}
		List<int> list2 = new List<int>();
		foreach (HTKKCMCTableCellInfo cell2 in hTKKCMCTableInfo.RowsInfo.Cells)
		{
			list2.Add(cell2.CellTemplate.ColIndex);
		}
		List<string> list3 = new List<string>();
		list3.Add("01");
		list3.Add("02");
		GeneralFunctions.TongHopTuBangSangBangTheoLoai(winFpSpread, ToKhaiInfo, 0, 0, tablePath_Muc_4_2_1, tablePath_Muc_3_2_1, list, list2, list3, HTKKFile.ColumnLettersToIndex("FS"));
		winFpSpread.Sheets[0].ActiveRowIndex = iRow;
	}

	public void DayBangTu_4_2_3Sang_3_2_4_1_b(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int iRow, int iCol)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_3_2_4_1_b);
		HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_4_2_3);
		int dataLocationTemplateRowStart = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
		int dataLocationTemplateRowStart2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		if (iRow < dataLocationTemplateRowStart || iRow > dataLocationTemplateRowStart + hTKKCMCTableInfo2.RowsValues.Count)
		{
			return;
		}
		List<int> list = new List<int>();
		foreach (HTKKCMCTableCellInfo cell in hTKKCMCTableInfo2.RowsInfo.Cells)
		{
			if (cell.CellTemplate.ColIndex != HTKKFile.ColumnLettersToIndex("AI") && cell.CellTemplate.ColIndex != HTKKFile.ColumnLettersToIndex("AR"))
			{
				list.Add(cell.CellTemplate.ColIndex);
			}
		}
		List<int> list2 = new List<int>();
		foreach (HTKKCMCTableCellInfo cell2 in hTKKCMCTableInfo.RowsInfo.Cells)
		{
			list2.Add(cell2.CellTemplate.ColIndex);
		}
		List<string> list3 = new List<string>();
		list3.Add("01");
		list3.Add("02");
		GeneralFunctions.TongHopTuBangSangBangTheoLoai(winFpSpread, ToKhaiInfo, 0, 0, tablePath_Muc_4_2_3, tablePath_Muc_3_2_4_1_b, list, list2, list3, HTKKFile.ColumnLettersToIndex("FS"));
		winFpSpread.Sheets[0].ActiveRowIndex = iRow;
	}

	public void ThemNhieuDongNamTaiChinh_3_2_4_2(SheetView sView, HTKKCMCSheetInfo sInfo, int iRow, int iCol, string tablePath, List<int> colsLock)
	{
		if (iCol != colD)
		{
			return;
		}
		int num = sInfo.Tables.FindIndex((HTKKCMCTableInfo x) => x.TablePath == tablePath);
		HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[num];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		List<HTKKCMCRowValue> list = new List<HTKKCMCRowValue>();
		if (iRow < dataLocationTemplateRowStart || iRow > dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count)
		{
			return;
		}
		string text = sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("FS") - 1].Text;
		if (!(text == "02"))
		{
			return;
		}
		for (int num2 = 0; num2 < 6; num2++)
		{
			HTKKCMCRowValue hTKKCMCRowValue = new HTKKCMCRowValue();
			foreach (HTKKCMCTableCellInfo cell in hTKKCMCTableInfo.RowsInfo.Cells)
			{
				hTKKCMCRowValue.RowValue.Add("");
			}
			int num3 = 0;
			foreach (HTKKCMCTableCellInfo cell2 in hTKKCMCTableInfo.RowsInfo.Cells)
			{
				hTKKCMCRowValue.RowValue[num3] = cell2.DefaultValue;
				num3++;
			}
			list.Add(hTKKCMCRowValue);
		}
		for (int num2 = 0; num2 < list.Count; num2++)
		{
			HTKKCMCRowValue hTKKCMCRowValue2 = list[num2];
			hTKKCMCRowValue2.RowValue[0] = "Nhóm đơn vị tính toán thuế";
			hTKKCMCRowValue2.RowValue[1] = "02";
			hTKKCMCRowValue2.RowValue[7] = num2 switch
			{
				4 => "Năm tài chính liền trước thứ nhất", 
				3 => "Năm tài chính liền trước thứ 2", 
				2 => "Năm tài chính liền trước thứ 3", 
				1 => "Năm tài chính liền trước thứ 4", 
				0 => "Năm tài chính liền trước thứ 5", 
				_ => "Năm tài chính báo cáo", 
			};
		}
		int num4 = 0;
		foreach (HTKKCMCRowValue item in list)
		{
			GeneralFunctions.InsertTableRow(sView, iRow + num4 - 1, -1, iRow + num4 - 1, -1, -1, sInfo, num);
			num4++;
			if (num4 == list.Count - 1)
			{
				break;
			}
		}
		for (int num2 = 0; num2 < list.Count; num2++)
		{
			int num5 = 0;
			foreach (HTKKCMCTableCellInfo cell3 in hTKKCMCTableInfo.RowsInfo.Cells)
			{
				sView.Cells[iRow + num2 - 1, cell3.CellTemplate.ColIndex - 1].Value = list[num2].RowValue[num5];
				num5++;
			}
			GeneralFunctions.lockUnlockColsInTableByFile(sView, sInfo, iRow + num2, colD, mapAnHienBang);
		}
		for (int num2 = 1; num2 < list.Count; num2++)
		{
			for (int num6 = 0; num6 < colsLock.Count; num6++)
			{
				GeneralFunctions.LockUnLockCol(sView, iRow + num2, colsLock[num6], locked: true, deleteData: true);
				if (num6 + 1 == num2)
				{
					break;
				}
			}
		}
	}

	public void ThemNhieuDongNamTaiChinh_3_3_3_1(SheetView sView, HTKKCMCSheetInfo sInfo, int iRow, int iCol, string tablePath, List<int> colsLock)
	{
		if (iCol != colD)
		{
			return;
		}
		int num = sInfo.Tables.FindIndex((HTKKCMCTableInfo x) => x.TablePath == tablePath);
		HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[num];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		List<HTKKCMCRowValue> list = new List<HTKKCMCRowValue>();
		if (iRow < dataLocationTemplateRowStart || iRow > dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count)
		{
			return;
		}
		string text = sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex("FS") - 1].Text;
		if (!(text == "02"))
		{
			return;
		}
		for (int num2 = 0; num2 < 5; num2++)
		{
			HTKKCMCRowValue hTKKCMCRowValue = new HTKKCMCRowValue();
			foreach (HTKKCMCTableCellInfo cell in hTKKCMCTableInfo.RowsInfo.Cells)
			{
				hTKKCMCRowValue.RowValue.Add("");
			}
			int num3 = 0;
			foreach (HTKKCMCTableCellInfo cell2 in hTKKCMCTableInfo.RowsInfo.Cells)
			{
				hTKKCMCRowValue.RowValue[num3] = cell2.DefaultValue;
				num3++;
			}
			list.Add(hTKKCMCRowValue);
		}
		for (int num2 = 0; num2 < list.Count; num2++)
		{
			HTKKCMCRowValue hTKKCMCRowValue2 = list[num2];
			hTKKCMCRowValue2.RowValue[0] = "Nhóm đơn vị tính toán thuế";
			hTKKCMCRowValue2.RowValue[1] = "02";
			hTKKCMCRowValue2.RowValue[7] = num2 switch
			{
				3 => "Năm tài chính liền trước thứ nhất", 
				2 => "Năm tài chính liền trước thứ 2", 
				1 => "Năm tài chính liền trước thứ 3", 
				0 => "Năm tài chính liền trước thứ 4", 
				_ => "Năm tài chính báo cáo", 
			};
		}
		int num4 = 0;
		foreach (HTKKCMCRowValue item in list)
		{
			GeneralFunctions.InsertTableRow(sView, iRow + num4 - 1, -1, iRow + num4 - 1, -1, -1, sInfo, num);
			num4++;
			if (num4 == list.Count - 1)
			{
				break;
			}
		}
		for (int num2 = 0; num2 < list.Count; num2++)
		{
			int num5 = 0;
			foreach (HTKKCMCTableCellInfo cell3 in hTKKCMCTableInfo.RowsInfo.Cells)
			{
				sView.Cells[iRow + num2 - 1, cell3.CellTemplate.ColIndex - 1].Value = list[num2].RowValue[num5];
				num5++;
			}
			GeneralFunctions.lockUnlockColsInTableByFile(sView, sInfo, iRow + num2, colD, mapAnHienBang);
		}
		for (int num2 = 1; num2 < list.Count; num2++)
		{
			for (int num6 = 0; num6 < colsLock.Count; num6++)
			{
				GeneralFunctions.LockUnLockCol(sView, iRow + num2, colsLock[num6], locked: true, deleteData: true);
				if (num6 + 1 == num2)
				{
					break;
				}
			}
			if (num2 == list.Count - 1)
			{
				GeneralFunctions.LockUnLockCol(sView, iRow + num2, HTKKFile.ColumnLettersToIndex("CF"), locked: true, deleteData: true);
			}
		}
	}

	public void lockUnlockTableNamChuyenTiep_3_2_4_3(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKPeriodInfo PeriodInfo)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_Muc_3_2_4_3_NamChuyenTiep);
		string text = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text;
		HTKKCMCTableInfo table = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_3_2_4_3_a1);
		HTKKCMCTableInfo table2 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_3_2_4_3_a2);
		HTKKCMCTableInfo table3 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Muc_3_2_4_3_b);
		if (!string.IsNullOrEmpty(text) && text == PeriodInfo.Nam)
		{
			GeneralFunctions.LockUnLockColsInTable(sView, table, GeneralFunctions.getAllColumnsInTable(table), locked: false, deleteData: false);
			GeneralFunctions.LockUnLockColsInTable(sView, table2, GeneralFunctions.getAllColumnsInTable(table2), locked: false, deleteData: false);
			GeneralFunctions.LockUnLockColsInTable(sView, table3, GeneralFunctions.getAllColumnsInTable(table3), locked: false, deleteData: false);
		}
		else
		{
			GeneralFunctions.LockUnLockColsInTable(sView, table, GeneralFunctions.getAllColumnsInTable(table), locked: true, deleteData: true);
			GeneralFunctions.LockUnLockColsInTable(sView, table2, GeneralFunctions.getAllColumnsInTable(table2), locked: true, deleteData: true);
			GeneralFunctions.LockUnLockColsInTable(sView, table3, GeneralFunctions.getAllColumnsInTable(table3), locked: true, deleteData: true);
		}
	}

	public void sua1DongCacDongKhacCungLoaiAnTheo(SheetView sView, HTKKCMCSheetInfo sInfo, int iRow, int iCol, string tablePath, string colMaLoai, string maLoaiSoSanh, string colTenEdit, string colMaEdit)
	{
		if (iCol != HTKKFile.ColumnLettersToIndex(colTenEdit))
		{
			return;
		}
		int index = sInfo.Tables.FindIndex((HTKKCMCTableInfo x) => x.TablePath == tablePath);
		HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[index];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		List<HTKKCMCRowValue> list = new List<HTKKCMCRowValue>();
		if (iRow < dataLocationTemplateRowStart || iRow > dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count)
		{
			return;
		}
		object value = sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex(colMaLoai) - 1].Value;
		if (value != maLoaiSoSanh)
		{
			return;
		}
		string text = sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex(colTenEdit) - 1].Text;
		string value2 = "";
		if (!string.IsNullOrEmpty(colMaEdit))
		{
			value2 = sView.Cells[iRow - 1, HTKKFile.ColumnLettersToIndex(colMaEdit) - 1].Text;
		}
		int num = iRow;
		while (num >= dataLocationTemplateRowStart)
		{
			object value3 = sView.Cells[num - 1, HTKKFile.ColumnLettersToIndex(colMaLoai) - 1].Value;
			if (value3 == maLoaiSoSanh)
			{
				sView.Cells[num - 1, HTKKFile.ColumnLettersToIndex(colTenEdit) - 1].Value = text;
				if (!string.IsNullOrEmpty(colMaEdit))
				{
					sView.Cells[num - 1, HTKKFile.ColumnLettersToIndex(colMaEdit) - 1].Value = value2;
				}
				num--;
				continue;
			}
			break;
		}
		for (num = iRow; num <= dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count; num++)
		{
			object value3 = sView.Cells[num - 1, HTKKFile.ColumnLettersToIndex(colMaLoai) - 1].Value;
			if (value3 == maLoaiSoSanh)
			{
				sView.Cells[num - 1, HTKKFile.ColumnLettersToIndex(colTenEdit) - 1].Value = text;
				if (!string.IsNullOrEmpty(colMaEdit))
				{
					sView.Cells[num - 1, HTKKFile.ColumnLettersToIndex(colMaEdit) - 1].Value = value2;
				}
				continue;
			}
			break;
		}
	}
}
