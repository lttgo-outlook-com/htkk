using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TkttQdmtt01Nckt288
{
	private int[] cell_TenDonViTien = GeneralFunctions.ArrayCellId("AS_37");

	private int[] cell_MaDonViTien = GeneralFunctions.ArrayCellId("DA_37");

	private int[] cell_1_1_co = GeneralFunctions.ArrayCellId("E_54");

	private int[] cell_1_1_khong = GeneralFunctions.ArrayCellId("E_56");

	private string path_1_3_2_co = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_1/Muc_1.3.2/thayDoiBCTC_Co";

	private string path_1_3_2_khong = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_1/Muc_1.3.2/thayDoiBCTC_Khong";

	private string path_1_3_3_co = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_1/Muc_1.3.3/thayDoiBCTC_Co";

	private string path_1_3_3_khong = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_1/Muc_1.3.3/thayDoiBCTC_Khong";

	private string path_1_3_4_a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_1/Muc_1.3.4/thayDoi_AnhHuong";

	private string path_2_1_MaQuyDinhGiamTru = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.1/ma_QuyDinhGiamTru";

	private string path_2_1_1_a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.1/TongDoanhThu/nam_TC";

	private string path_2_1_1_b = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.1/TongDoanhThu/nam_TCLienTruocThu1";

	private string path_2_1_1_c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.1/TongDoanhThu/nam_TCLienTruocThu2";

	private string path_2_1_1_d = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.1/TongDoanhThu/binhQuan_3NamTC";

	private string pathMuc_2_1_2_ct1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.1.2/tongDoanhThu";

	private string pathMuc_2_1_2_ct2 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.1.2/loiNhuan_TruocThue";

	private string pathMuc_2_1_2_ct3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.1.2/tongThue_DonGianHoa";

	private string pathMuc_2_1_2_ct4 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.1.2/thueSuatThucTe_DonGianHoa";

	private string pathMuc2_2_check = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.2/thueSuatThucTe";

	private string pathMuc2_2_Start = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.2/DoanhThu_BCTC/nam_TC";

	private string pathMuc2_2_End = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.2/ThuNhap_ThueTTTC/binhQuan_3NamTC";

	private string path_Muc_2_1_2_ct4 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.1.2/thueSuatThucTe_DonGianHoa";

	private string pathMuc_2_3_ct4 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.3/soQuocGia_ThueTTTC";

	private string pathMuc_2_3_ct6 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.3/soQuocGia";

	private string path_3_1_2_1_ct1_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.1.2.1/chiTienLuong";

	private string path_3_1_2_1_ct2_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.1.2.1/giaTriTaiSan";

	private string path_3_1_2_1_ct2_2_ten = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.1.2.1/ten_TyLeGiamTruTaiSan";

	private string path_3_1_2_1_ct2_2_ma = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.1.2.1/ma_TyLeGiamTruTaiSan";

	private string path_3_1_2_1_ct1_2_ten = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.1.2.1/ten_TyLeGiamTruTienLuong";

	private string path_3_1_2_1_ct1_2_ma = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.1.2.1/ma_TyLeGiamTruTienLuong";

	private string path_3_2_3_1_ctB = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.2.3.1/ct_B";

	private string path_3_2_3_3_NamChuyenTiep = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.2.3.3/nam_ChuyenTiep";

	private string path_3_2_3_3_Start = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.2.3.3/Muc_a/khoan_ChuyenTiep";

	private string path_3_2_3_3_a_2_ct1_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.2.3.3/Muc_a/khoan_TaiSanChuyenTiep";

	private string path_3_2_3_3_a_2_ct1_2 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.2.3.3/Muc_a/khoan_TaiSanTinhLai";

	private string path_3_2_3_3_a_2_ct1_3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.2.3.3/Muc_a/khoan_TaiSanPhatSinh";

	private string path_3_2_3_3_a_2_ct1_4 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.2.3.3/Muc_a/khoan_TaiSanGhiNhan";

	private string pathMuc3_4_ct1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.4/thue_ThuNhapHoacLo";

	private string pathMuc3_4_ct4 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.4/thue_DieuChinh";

	private string pathMuc_3_4_ct5 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.4/thue_ThucTe";

	private string pathMuc_3_6_ct1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.6/thueBoSung_TyLe";

	private string pathMuc_3_6_ct3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.6/thueBoSung_LoiNhuan";

	private string pathMuc_3_6_ct4 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.6/thueBoSung_DieuChinh";

	private string path_Muc_3_2_3_1_B = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.2.3.1/ct_B";

	private string path_Muc_3_2_3_1_C = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.2.3.1/ct_C";

	private string path_Muc_3_2_3_1_D = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.2.3.1/ct_D";

	private string tablePath_1_3_2 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_1/Muc_1.3.2/ChiTietDanhSach";

	private string tablePath_1_3_3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_1/Muc_1.3.3/ChiTietDanhSach";

	private string tablePath_1_3_4 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_1/Muc_1.3.4/ChiTietDanhSach";

	private string tablePath_2_3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_2/Muc_2.3/ChiTietDanhSach";

	private string tablePath_3_1_2_2 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.1.2.2/ChiTietDanhSach";

	private string tablePath_3_1_2_3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.1.2.3/ChiTietDanhSach";

	private string tablePath_3_2_3_3_b = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.2.3.3/Muc_b/ChiTietDanhSach";

	private string tablePath_3_5_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_3/Muc_3.5.1/ChiTietDanhSach";

	private string tablePath_4_1_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_4/Muc_4.1.1/ChiTietDanhSach";

	private string tablePath_4_1_2 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_4/Muc_4.1.2/ChiTietDanhSach";

	private string tablePath_4_1_3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_4/Muc_4.1.3/ChiTietDanhSach";

	private string tablePath_4_1_4 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_4/Muc_4.1.4/ChiTietDanhSach";

	private string tablePath_4_2_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_4/Muc_4.2.1/ChiTietDanhSach";

	private string tablePath_4_4_1 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_4/Muc_4.4.1/ChiTietDanhSach";

	private string tablePath_4_2_2 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_4/Muc_4.2.2/ChiTietDanhSach";

	private string tablePath_4_2_3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_4/Muc_4.2.3/ChiTietDanhSach";

	private string tablePath_4_3_3 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Muc_4/Muc_4.3.3/ChiTietDanhSach";

	private int colZ = HTKKFile.ColumnLettersToIndex("Z");

	private int colT = HTKKFile.ColumnLettersToIndex("T");

	private int[] cellMaCt9 = GeneralFunctions.ArrayCellId("AA_26");

	private string pathCt10_KhacVnd_Pl_01_1 = "TKhai_01_TNDN_QDMTT/LoaiTienTNDN/tien_KhacVND";

	private string pathCt10_Vnd_Pl_01_1 = "TKhai_01_TNDN_QDMTT/LoaiTienTNDN/tien_VND";

	private string pathCt11_TiGia_Pl_01_1 = "TKhai_01_TNDN_QDMTT/ty_GiaQuyDoi";

	private int[] cellKhacVND = GeneralFunctions.ArrayCellId("G_29");

	private int[] cellA6 = GeneralFunctions.ArrayCellId("O_47");

	private int[] cellA7 = GeneralFunctions.ArrayCellId("O_48");

	private int[] cellA10 = GeneralFunctions.ArrayCellId("O_51");

	private int[] cellB = GeneralFunctions.ArrayCellId("O_52");

	private int[] cellB1 = GeneralFunctions.ArrayCellId("O_53");

	private int[] cellB1Formular = GeneralFunctions.ArrayCellId("AA_53");

	private int[] cellB2 = GeneralFunctions.ArrayCellId("O_54");

	private string pathTongCot5_II = "TKhai_01_TNDN_QDMTT/Muc_II/so_ThueTNDN/tong_TyLePhanBo";

	private string pathCtA6_Pl_01_1 = "TKhai_01_TNDN_QDMTT/Muc_I/ct_A6";

	private string pathMaCtA7_Pl_01_1 = "TKhai_01_TNDN_QDMTT/Muc_I/ct_A7";

	private string pathTenCtA7_Pl_01_1 = "path_ct_A7";

	private string pathCtA9_Pl_01_1 = "TKhai_01_TNDN_QDMTT/Muc_I/ct_A9";

	private string pathTenCtA10_Pl_01_1 = "path_ct_A10";

	private string pathMaCtA10_Pl_01_1 = "TKhai_01_TNDN_QDMTT/Muc_I/ct_A10";

	private string pathCtB1_Pl_01_1 = "TKhai_01_TNDN_QDMTT/Muc_I/ct_B1";

	private string pathCtC_Pl_01_1 = "TKhai_01_TNDN_QDMTT/Muc_I/ct_C";

	private string pathCtD_Pl_01_1 = "TKhai_01_TNDN_QDMTT/Muc_I/ct_D";

	private string pathCtE_Pl_01_1 = "TKhai_01_TNDN_QDMTT/Muc_I/ct_E";

	private string tablePath_Pl_01_1 = "TKhai_01_TNDN_QDMTT/Muc_II/so_ThueTNDN/ChiTiet";

	private string tablePath_Pl_01_2 = "TKhai_01_TM/ChiTiet";

	private List<MapCongThuc> mapCongThuc = new List<MapCongThuc>();

	private string pathTongHopKhbs = "path_tong_hop_khbs";

	private string pathTenDonViTienKhbs = "path_ten_don_vi_tien_khbs";

	private string pathMaDonViTienKhbs = "path_ma_don_vi_tien_khbs";

	private string pathKhbs_A_I_3_b = "PL01_KHBS/Muc_A/Muc_I/Muc_3/ct3b";

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sView = winFpSpread.Sheets[0];
			GeneralFunctions.checkLockUnLockCells(ToKhaiInfo, winFpSpread);
			lockUnlockTable_1_3_4(sView, sInfo);
			lockUnlock2_1_1(sView, sInfo);
			lockUnlock2_1_2(sView, sInfo);
			lockUnlock3_2_3_3(sView, sInfo, PeriodInfo);
			lockUnlockTable_3_2_3_3_b(sView, sInfo, PeriodInfo);
			mapCongThuc = GeneralFunctions.MapCongThucChoSua(PeriodInfo.menuID);
			ThayDoiDonViTien_Tk(sView, sInfo);
			if (PeriodInfo.LoaiToKhai == "B")
			{
				GeneralFunctions.LockUnLockCol(sView, cell_TenDonViTien[1] + 1, cell_TenDonViTien[0] + 1, locked: true, deleteData: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			ThayDoiDonViTien_Pl_01_1(winFpSpread, ToKhaiInfo);
			LockUnLock_Pl_01_1(winFpSpread, ToKhaiInfo);
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCt10_KhacVnd_Pl_01_1);
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCt10_Vnd_Pl_01_1);
			string text = sheetView.Cells[cellMaCt9[1], cellMaCt9[0]].Text;
			bool flag = Convert.ToBoolean(sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value);
			if (PeriodInfo.LoaiToKhai == "B")
			{
				GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo.CellTemplate.ColIndex, locked: true, deleteData: false);
				GeneralFunctions.LockUnLockCol(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex, hTKKCMCCellInfo2.CellTemplate.ColIndex, locked: true, deleteData: false);
			}
			if (PeriodInfo.LoaiToKhai == "B")
			{
				if (text != "VND" && flag)
				{
					GeneralFunctions.UpdateFileMapMctKhbs(PeriodInfo.menuID, "MCTPath", pathTongHopKhbs, "MCT", "(E)");
					GeneralFunctions.UpdateFileMapMctKhbs(PeriodInfo.menuID, "MCTPath", pathTongHopKhbs, "Caption", "Tổng số thuế bổ sung của đơn vị hợp thành theo đồng Việt Nam trong trường hợp khai Tờ khai thông tin bằng đồng ngoại tệ và lựa chọn khai Tờ khai thuế TNDN bổ sung, nộp thuế TNDN bổ sung bằng đồng Việt Nam (nếu có)");
				}
				else
				{
					GeneralFunctions.UpdateFileMapMctKhbs(PeriodInfo.menuID, "MCTPath", pathTongHopKhbs, "MCT", "(D)");
					GeneralFunctions.UpdateFileMapMctKhbs(PeriodInfo.menuID, "MCTPath", pathTongHopKhbs, "Caption", "Tổng số thuế bổ sung của đơn vị hợp thành tại Việt Nam (D = A x B + C)");
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
				NumberCellType numberCellType = (NumberCellType)sheetView3.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].CellType;
				numberCellType.DecimalPlaces = ((!(text == "VND")) ? 2 : 0);
				sheetView3.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].CellType = numberCellType;
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
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_2_1_MaQuyDinhGiamTru);
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_3_2_3_3_NamChuyenTiep);
			if (iRow == hTKKCMCCellInfo.CellTemplate.RowIndex)
			{
				lockUnlock2_1_1(sView, hTKKCMCSheetInfo);
				lockUnlock2_1_2(sView, hTKKCMCSheetInfo);
			}
			if (iRow == hTKKCMCCellInfo2.CellTemplate.RowIndex)
			{
				lockUnlock3_2_3_3(sView, hTKKCMCSheetInfo, PeriodInfo);
				lockUnlockTable_3_2_3_3_b(sView, hTKKCMCSheetInfo, PeriodInfo);
			}
			if (iRow == cell_MaDonViTien[1] + 1)
			{
				ThayDoiDonViTien_Tk(sView, hTKKCMCSheetInfo);
				ThayDoiDonViTien_Pl_01_1(winFpSpread, ToKhaiInfo);
				ThayDoiDonViTien_Pl_01_2(winFpSpread, ToKhaiInfo);
				LockUnLock_Pl_01_1(winFpSpread, ToKhaiInfo);
			}
			TinhCongThucChoSua_Tk(sView, hTKKCMCSheetInfo, (int)iRow, (int)iCol);
			LayGiaTriTKSangPL_01_1(winFpSpread, ToKhaiInfo, (int)iRow, (int)iCol);
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
			TinhCongThucChoSua_Pl01_1(sView, sInfo, (int)iRow, (int)iCol);
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
		checkPL_01_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
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
			string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_61"));
			string message2 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_65"));
			string message3 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_66"));
			string message4 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_110"));
			string message5 = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_36"));
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_2_1_MaQuyDinhGiamTru);
			string text = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text;
			HTKKCMCTableInfo table = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_1_3_2);
			int num = GeneralFunctions.checkTrungGiaTri1CotTrongBang(sView, table, colZ, message3);
			if (num > 0)
			{
				value = num;
			}
			HTKKCMCTableInfo table2 = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_1_3_3);
			int num2 = GeneralFunctions.checkTrungGiaTri1CotTrongBang(sView, table2, colZ, message3);
			if (num2 > 0)
			{
				value = num2;
			}
			HTKKCMCTableInfo table3 = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_1_3_4);
			int num3 = GeneralFunctions.checkTrungGiaTri1CotTrongBang(sView, table3, colT, message3);
			if (num3 > 0)
			{
				value = num3;
			}
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc_2_1_2_ct1);
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc_2_1_2_ct2);
			HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc_2_1_2_ct3);
			HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc_2_1_2_ct4);
			double num4 = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value);
			double num5 = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value);
			double num6 = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value);
			double num7 = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Value);
			HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1);
			if ((text.IndexOf("01") >= 0 || text.IndexOf("02") >= 0 || text.IndexOf("03") >= 0) && num4 == 0.0 && num5 == 0.0 && num6 == 0.0 && num7 == 0.0)
			{
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, message4);
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, message4);
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1, message4);
				HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1, message4);
				value = 2;
			}
			HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc2_2_check);
			HTKKCMCCellInfo cellMuc2_2_Start = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc2_2_Start);
			HTKKCMCCellInfo cellMuc2_2_End = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc2_2_End);
			bool flag = Convert.ToBoolean(sView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Value);
			List<HTKKCMCCellInfo> list = hTKKCMCSheetInfo.StaticCells.FindAll((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex >= cellMuc2_2_Start.CellTemplate.RowIndex && x.CellTemplate.RowIndex <= cellMuc2_2_End.CellTemplate.RowIndex);
			foreach (HTKKCMCCellInfo item in list)
			{
				HTKKFile.RemoveValidate(sView, item.CellTemplate.RowIndex - 1, item.CellTemplate.ColIndex - 1);
			}
			if (flag)
			{
				bool flag2 = true;
				foreach (HTKKCMCCellInfo item2 in list)
				{
					double num8 = Convert.ToDouble(sView.Cells[item2.CellTemplate.RowIndex - 1, item2.CellTemplate.ColIndex - 1].Value);
					if (num8 != 0.0)
					{
						flag2 = false;
						break;
					}
				}
				if (flag2)
				{
					foreach (HTKKCMCCellInfo item3 in list)
					{
						HTKKFile.SetNodeBackGroundPink(sView, item3.CellTemplate.RowIndex - 1, item3.CellTemplate.ColIndex - 1, message5);
						value = 2;
					}
				}
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
			string message = HTKKFile.GetMessage(sheetView2, GeneralFunctions.ArrayCellId("E_84"));
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo2.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Pl_01_1);
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_1_3_2);
			int dataLocationTemplateRowStart = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			List<string> list = new List<string>();
			for (int num2 = 0; num2 < hTKKCMCTableInfo2.RowsValues.Count; num2++)
			{
				string text = sheetView.Cells[dataLocationTemplateRowStart + num2 - 1, colZ - 1].Text;
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
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_1_3_2);
			int dataLocationTemplateRowStart = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			List<string> list = new List<string>();
			for (int num2 = 0; num2 < hTKKCMCTableInfo2.RowsValues.Count; num2++)
			{
				string text = sheetView.Cells[dataLocationTemplateRowStart + num2 - 1, colZ - 1].Text;
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sView = winFpSpread.Sheets[0];
			int iColClick = winFpSpread.ActiveSheet.ActiveColumnIndex + 1;
			int num = winFpSpread.ActiveSheet.ActiveRowIndex + 1;
			if (num == 20)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_01_TKTT_QDMTT.xls");
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_1_3_4_a);
			resetOtherCheckboxTK(sView, hTKKCMCSheetInfo, num, iColClick);
			if (num == hTKKCMCCellInfo.CellTemplate.RowIndex)
			{
				lockUnlockTable_1_3_4(sView, hTKKCMCSheetInfo);
			}
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
			LockUnLock_Pl_01_1(winFpSpread, ToKhaiInfo);
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
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_1_3_2_co);
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_1_3_2_khong);
		HTKKCMCCellInfo hTKKCMCCellInfo3 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_1_3_3_co);
		HTKKCMCCellInfo hTKKCMCCellInfo4 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_1_3_3_khong);
		HTKKCMCCellInfo hTKKCMCCellInfo5 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_1_3_4_a);
		if (iRowClick == cell_1_1_co[1] + 1 || iRowClick == cell_1_1_khong[1] + 1 || iRowClick == hTKKCMCCellInfo.CellTemplate.RowIndex || iRowClick == hTKKCMCCellInfo2.CellTemplate.RowIndex || iRowClick == hTKKCMCCellInfo3.CellTemplate.RowIndex || iRowClick == hTKKCMCCellInfo4.CellTemplate.RowIndex || iRowClick == hTKKCMCCellInfo5.CellTemplate.RowIndex)
		{
			GeneralFunctions.resetOtherCheckboxStatic(sView, sInfo, cell_1_1_co[1] + 1, cell_1_1_khong[1] + 1, iRowClick, iColClick);
			GeneralFunctions.resetOtherCheckboxStatic(sView, sInfo, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo2.CellTemplate.RowIndex, iRowClick, iColClick);
			GeneralFunctions.resetOtherCheckboxStatic(sView, sInfo, hTKKCMCCellInfo3.CellTemplate.RowIndex, hTKKCMCCellInfo4.CellTemplate.RowIndex, iRowClick, iColClick);
			GeneralFunctions.resetOtherCheckboxStatic(sView, sInfo, hTKKCMCCellInfo5.CellTemplate.RowIndex, hTKKCMCCellInfo5.CellTemplate.RowIndex, iRowClick, iColClick);
		}
		HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_4_3_3);
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		if (iRowClick >= dataLocationTemplateRowStart && iRowClick <= dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count)
		{
			List<int> list = new List<int>();
			list.Add(HTKKFile.ColumnLettersToIndex("AO"));
			list.Add(HTKKFile.ColumnLettersToIndex("BC"));
			GeneralFunctions.resetOtherCheckboxDinamic(sView, sInfo, iRowClick, iColClick, list);
		}
	}

	public void lockUnlockTable_1_3_4(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_1_3_4_a);
		HTKKCMCTableInfo table = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_1_3_4);
		if (!Convert.ToBoolean(sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value))
		{
			GeneralFunctions.LockUnLockColsInTable(sView, table, GeneralFunctions.getAllColumnsInTable(table), locked: true, deleteData: true);
		}
		else
		{
			GeneralFunctions.LockUnLockColsInTable(sView, table, GeneralFunctions.getAllColumnsInTable(table), locked: false, deleteData: false);
		}
	}

	public void lockUnlock2_1_1(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_2_1_MaQuyDinhGiamTru);
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_2_1_1_a);
		HTKKCMCCellInfo hTKKCMCCellInfo3 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_2_1_1_b);
		HTKKCMCCellInfo hTKKCMCCellInfo4 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_2_1_1_c);
		HTKKCMCCellInfo hTKKCMCCellInfo5 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_2_1_1_d);
		string text = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text;
		text = text.Replace("GTTN_", "");
		string[] array = text.Split(new string[1] { ";" }, StringSplitOptions.None);
		bool flag = false;
		string[] array2 = array;
		foreach (string value in array2)
		{
			if (!string.IsNullOrEmpty(value) && Convert.ToInt32(value) >= 4)
			{
				flag = true;
				break;
			}
		}
		if (!string.IsNullOrEmpty(text) && flag)
		{
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, hTKKCMCCellInfo2.CellTemplate.RowIndex, hTKKCMCCellInfo2.CellTemplate.RowIndex, isLock: false, isDeleteData: false);
			GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo3.CellTemplate.RowIndex, hTKKCMCCellInfo3.CellTemplate.ColIndex, locked: false, deleteData: false);
			GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo4.CellTemplate.RowIndex, hTKKCMCCellInfo4.CellTemplate.ColIndex, locked: false, deleteData: false);
			GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex, hTKKCMCCellInfo5.CellTemplate.ColIndex, locked: false, deleteData: false);
		}
		else
		{
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, hTKKCMCCellInfo2.CellTemplate.RowIndex, hTKKCMCCellInfo2.CellTemplate.RowIndex, isLock: true, isDeleteData: true);
			GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo3.CellTemplate.RowIndex, hTKKCMCCellInfo3.CellTemplate.ColIndex, locked: true, deleteData: true);
			GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo4.CellTemplate.RowIndex, hTKKCMCCellInfo4.CellTemplate.ColIndex, locked: true, deleteData: true);
			GeneralFunctions.LockUnLockCol(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex, hTKKCMCCellInfo5.CellTemplate.ColIndex, locked: true, deleteData: true);
		}
	}

	public void lockUnlock2_1_2(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_2_1_MaQuyDinhGiamTru);
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc_2_1_2_ct1);
		string text = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text;
		text = text.Replace("GTTN_", "");
		string[] array = text.Split(new string[1] { ";" }, StringSplitOptions.None);
		bool flag = false;
		string[] array2 = array;
		foreach (string value in array2)
		{
			if (!string.IsNullOrEmpty(value) && Convert.ToInt32(value) < 4)
			{
				flag = true;
				break;
			}
		}
		if (!string.IsNullOrEmpty(text) && flag)
		{
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, hTKKCMCCellInfo2.CellTemplate.RowIndex, hTKKCMCCellInfo2.CellTemplate.RowIndex, isLock: false, isDeleteData: false);
		}
		else
		{
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, hTKKCMCCellInfo2.CellTemplate.RowIndex, hTKKCMCCellInfo2.CellTemplate.RowIndex, isLock: true, isDeleteData: true);
		}
	}

	public void lockUnlock3_2_3_3(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKPeriodInfo PeriodInfo)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_3_2_3_3_NamChuyenTiep);
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_3_2_3_3_Start);
		HTKKCMCCellInfo hTKKCMCCellInfo3 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_3_2_3_3_a_2_ct1_4);
		string text = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text;
		if (!string.IsNullOrEmpty(text) && text == PeriodInfo.Nam)
		{
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, hTKKCMCCellInfo2.CellTemplate.RowIndex, hTKKCMCCellInfo3.CellTemplate.RowIndex, isLock: false, isDeleteData: false);
		}
		else
		{
			GeneralFunctions.lockUnlockMultiCellsStatic(sView, sInfo, hTKKCMCCellInfo2.CellTemplate.RowIndex, hTKKCMCCellInfo3.CellTemplate.RowIndex, isLock: true, isDeleteData: true);
		}
	}

	public void lockUnlockTable_3_2_3_3_b(SheetView sView, HTKKCMCSheetInfo sInfo, HTKKPeriodInfo PeriodInfo)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_3_2_3_3_NamChuyenTiep);
		string text = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text;
		HTKKCMCTableInfo table = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_3_2_3_3_b);
		if (!string.IsNullOrEmpty(text) && text == PeriodInfo.Nam)
		{
			GeneralFunctions.LockUnLockColsInTable(sView, table, GeneralFunctions.getAllColumnsInTable(table), locked: false, deleteData: false);
		}
		else
		{
			GeneralFunctions.LockUnLockColsInTable(sView, table, GeneralFunctions.getAllColumnsInTable(table), locked: true, deleteData: true);
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
		List<int[]> list = new List<int[]>();
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_Muc_2_1_2_ct4);
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc_3_4_ct5);
		HTKKCMCCellInfo hTKKCMCCellInfo3 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc_3_6_ct1);
		HTKKCMCCellInfo hTKKCMCCellInfo4 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc_2_3_ct4);
		HTKKCMCCellInfo hTKKCMCCellInfo5 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc_2_3_ct6);
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
		list.Add(new int[2]
		{
			hTKKCMCCellInfo5.CellTemplate.RowIndex,
			hTKKCMCCellInfo5.CellTemplate.ColIndex
		});
		if (string.IsNullOrEmpty(text) || text == "VND")
		{
			changFormatCellDecimalPlacesStatic(sView, sInfo, 0, list);
		}
		else
		{
			changFormatCellDecimalPlacesStatic(sView, sInfo, 4, list);
		}
		List<int> list2 = new List<int>();
		HTKKCMCTableInfo table = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_2_3);
		HTKKCMCTableInfo table2 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_3_1_2_2);
		HTKKCMCTableInfo table3 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_3_1_2_3);
		HTKKCMCTableInfo table4 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_3_2_3_3_b);
		HTKKCMCTableInfo table5 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_3_5_1);
		HTKKCMCTableInfo table6 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_4_1_1);
		HTKKCMCTableInfo table7 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_4_1_2);
		HTKKCMCTableInfo table8 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_4_1_3);
		HTKKCMCTableInfo table9 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_4_1_4);
		HTKKCMCTableInfo table10 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_4_2_1);
		HTKKCMCTableInfo table11 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_4_2_2);
		HTKKCMCTableInfo table12 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_4_2_3);
		HTKKCMCTableInfo table13 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_4_4_1);
		if (string.IsNullOrEmpty(text) || text == "VND")
		{
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table, 0, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table2, 0, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table3, 0, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table4, 0, list2);
			list2.Clear();
			list2.Add(HTKKFile.ColumnLettersToIndex("AP"));
			list2.Add(HTKKFile.ColumnLettersToIndex("AS"));
			list2.Add(HTKKFile.ColumnLettersToIndex("BH"));
			list2.Add(HTKKFile.ColumnLettersToIndex("BK"));
			changFormatCellDecimalPlacesDinamic(sView, table5, 0, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table6, 0, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table7, 0, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table8, 0, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table9, 0, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table10, 0, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table11, 0, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table12, 0, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table13, 0, list2);
		}
		else
		{
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table, 4, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table2, 4, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table3, 4, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table4, 4, list2);
			list2.Clear();
			list2.Add(HTKKFile.ColumnLettersToIndex("AP"));
			list2.Add(HTKKFile.ColumnLettersToIndex("AS"));
			list2.Add(HTKKFile.ColumnLettersToIndex("BH"));
			list2.Add(HTKKFile.ColumnLettersToIndex("BK"));
			changFormatCellDecimalPlacesDinamic(sView, table5, 4, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table6, 4, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table7, 4, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table8, 4, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table9, 4, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table10, 4, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table11, 4, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table12, 4, list2);
			list2.Clear();
			changFormatCellDecimalPlacesDinamic(sView, table13, 4, list2);
		}
	}

	public void ThayDoiDonViTien_Pl_01_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		SheetView sheetView2 = winFpSpread.Sheets[1];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
		string text = sheetView.Cells[cell_MaDonViTien[1], cell_MaDonViTien[0]].Text;
		HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathTongCot5_II);
		HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCtE_Pl_01_1);
		HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCtD_Pl_01_1);
		List<int[]> list = new List<int[]>();
		list.Add(new int[2]
		{
			cellA7[1] + 1,
			cellA7[0] + 1
		});
		list.Add(new int[2]
		{
			cellA10[1] + 1,
			cellA10[0] + 1
		});
		list.Add(new int[2]
		{
			cellB[1] + 1,
			cellB[0] + 1
		});
		list.Add(new int[2]
		{
			cellB2[1] + 1,
			cellB2[0] + 1
		});
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
		if (string.IsNullOrEmpty(text) || text == "VND")
		{
			changFormatCellDecimalPlacesStatic(sheetView2, hTKKCMCSheetInfo, 0, list);
		}
		else
		{
			changFormatCellDecimalPlacesStatic(sheetView2, hTKKCMCSheetInfo, 4, list);
		}
		if (!string.IsNullOrEmpty(text) && !(text == "VND"))
		{
			NumberCellType numberCellType = (NumberCellType)sheetView2.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].CellType;
			numberCellType.DecimalPlaces = 2;
			sheetView2.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].CellType = numberCellType;
		}
		List<int> list2 = new List<int>();
		list2.Add(HTKKFile.ColumnLettersToIndex("Q"));
		HTKKCMCTableInfo table = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Pl_01_1);
		if (string.IsNullOrEmpty(text) || text == "VND")
		{
			changFormatCellDecimalPlacesDinamic(sheetView2, table, 0, list2);
		}
		else
		{
			changFormatCellDecimalPlacesDinamic(sheetView2, table, 4, list2);
		}
		HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathTongHopKhbs);
		HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMaDonViTienKhbs);
		string text2 = sheetView2.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Text;
		NumberCellType numberCellType2 = (NumberCellType)sheetView2.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].CellType;
		numberCellType2.DecimalPlaces = ((!(text2 == "VND")) ? 2 : 0);
		sheetView2.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].CellType = numberCellType2;
	}

	public void ThayDoiDonViTien_Pl_01_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		SheetView sView = winFpSpread.Sheets[2];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
		string text = sheetView.Cells[cell_MaDonViTien[1], cell_MaDonViTien[0]].Text;
		List<int> colsBoQua = new List<int>();
		HTKKCMCTableInfo table = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Pl_01_2);
		if (string.IsNullOrEmpty(text) || text == "VND")
		{
			changFormatCellDecimalPlacesDinamic(sView, table, 0, colsBoQua);
		}
		else
		{
			changFormatCellDecimalPlacesDinamic(sView, table, 4, colsBoQua);
		}
	}

	public void TinhCongThucChoSua_Tk(SheetView sView, HTKKCMCSheetInfo sInfo, int iRow, int iCol)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_Muc_3_2_3_1_B);
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_Muc_3_2_3_1_C);
		HTKKCMCCellInfo hTKKCMCCellInfo3 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_Muc_3_2_3_1_D);
		if ((iRow == hTKKCMCCellInfo2.CellTemplate.RowIndex && iCol == hTKKCMCCellInfo2.CellTemplate.ColIndex) || (iRow == hTKKCMCCellInfo3.CellTemplate.RowIndex && iCol == hTKKCMCCellInfo3.CellTemplate.ColIndex))
		{
			sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value) + Convert.ToDouble(sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value);
		}
		HTKKCMCCellInfo hTKKCMCCellInfo4 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_3_2_3_3_a_2_ct1_1);
		HTKKCMCCellInfo hTKKCMCCellInfo5 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_3_2_3_3_a_2_ct1_2);
		HTKKCMCCellInfo hTKKCMCCellInfo6 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_3_2_3_3_a_2_ct1_3);
		HTKKCMCCellInfo hTKKCMCCellInfo7 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_3_2_3_3_a_2_ct1_4);
		if (iRow == hTKKCMCCellInfo7.CellTemplate.RowIndex && iCol < hTKKCMCCellInfo7.CellTemplate.ColIndex)
		{
			double num = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value);
			double num2 = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Value);
			double num3 = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Value);
			if (num2 != 0.0)
			{
				sView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Value = num2 - num3;
			}
			else
			{
				sView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Value = num - num3;
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
		}
		else
		{
			int num4 = GeneralFunctions.DetectIndexOfTable(sInfo, iRow);
			if (num4 >= 0)
			{
				GeneralFunctions.CapNhatCongThucLoaiChoSuaCell(sView, sInfo, iRow, iCol, mapCongThuc);
			}
		}
		HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_4_1_1);
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		HTKKCMCTableInfo hTKKCMCTableInfo2 = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_4_2_1);
		int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
		HTKKCMCCellInfo hTKKCMCCellInfo8 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc_3_4_ct5);
		if ((iRow >= dataLocationTemplateRowStart && iRow <= dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count) || (iRow >= dataLocationTemplateRowStart2 && iRow <= dataLocationTemplateRowStart2 + hTKKCMCTableInfo2.RowsValues.Count) || (iRow == hTKKCMCCellInfo8.CellTemplate.RowIndex && iCol < hTKKCMCCellInfo8.CellTemplate.ColIndex))
		{
			GeneralFunctions.CapNhatCongThucLoaiChoSuaCell(sView, sInfo, hTKKCMCCellInfo8.CellTemplate.RowIndex, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1, mapCongThuc);
		}
		HTKKCMCCellInfo hTKKCMCCellInfo9 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_3_1_2_1_ct1_1);
		HTKKCMCCellInfo hTKKCMCCellInfo10 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc3_4_ct1);
		HTKKCMCCellInfo hTKKCMCCellInfo11 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc_3_6_ct3);
		if ((iRow >= dataLocationTemplateRowStart && iRow <= dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count) || iRow == hTKKCMCCellInfo9.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo10.CellTemplate.RowIndex)
		{
			GeneralFunctions.CapNhatCongThucLoaiChoSuaCell(sView, sInfo, hTKKCMCCellInfo11.CellTemplate.RowIndex, 0, mapCongThuc);
		}
	}

	public void TinhCongThucChoSua_Pl01_1(SheetView sView, HTKKCMCSheetInfo sInfo, int iRow, int iCol)
	{
		GeneralFunctions.CapNhatCongThucLoaiChoSuaCell(sView, sInfo, iRow, iCol, mapCongThuc);
		int num = GeneralFunctions.DetectIndexOfTable(sInfo, iRow);
		if (num >= 0)
		{
			GeneralFunctions.CapNhatCongThucLoaiChoSuaCell(sView, sInfo, iRow, iCol, mapCongThuc);
		}
		HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCt10_Vnd_Pl_01_1);
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCtD_Pl_01_1);
		HTKKCMCCellInfo hTKKCMCCellInfo3 = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCtE_Pl_01_1);
		HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_Pl_01_1);
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		int[] array = GeneralFunctions.ArrayCellId("C_31");
		int[] array2 = GeneralFunctions.ArrayCellId("O_41");
		int num2 = HTKKFile.ColumnLettersToIndex("Q");
		int num3 = HTKKFile.ColumnLettersToIndex("R");
		string text = sView.Cells[cellMaCt9[1], cellMaCt9[0]].Text;
		bool flag = Convert.ToBoolean(sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
		double num4 = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value);
		double num5 = Convert.ToDouble(sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value);
		if ((iRow == array[1] + 1 && iCol == array[0] + 1) || (iRow < dataLocationTemplateRowStart && iRow >= array2[1] + 1) || iRow == hTKKCMCCellInfo.CellTemplate.RowIndex)
		{
			for (int num6 = 0; num6 < hTKKCMCTableInfo.RowsValues.Count; num6++)
			{
				if (text != "VND" && flag)
				{
					sView.Cells[dataLocationTemplateRowStart + num6 - 1, num3 - 1].Value = Math.Round(Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + num6 - 1, num2 - 1].Value) / 100.0 * num5, 4);
				}
				else
				{
					sView.Cells[dataLocationTemplateRowStart + num6 - 1, num3 - 1].Value = Math.Round(Convert.ToDouble(sView.Cells[dataLocationTemplateRowStart + num6 - 1, num2 - 1].Value) / 100.0 * num4, 4);
				}
			}
		}
		else if (iRow > hTKKCMCCellInfo3.CellTemplate.RowIndex && iCol == num2)
		{
			if (text != "VND" && flag)
			{
				sView.Cells[iRow - 1, num3 - 1].Value = Math.Round(Convert.ToDouble(sView.Cells[iRow - 1, num2 - 1].Value) / 100.0 * num5, 4);
			}
			else
			{
				sView.Cells[iRow - 1, num3 - 1].Value = Math.Round(Convert.ToDouble(sView.Cells[iRow - 1, num2 - 1].Value) / 100.0 * num4, 4);
			}
		}
	}

	public void LayGiaTriTKSangPL_01_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int iRow, int iCol)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		SheetView sheetView2 = winFpSpread.Sheets[1];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[1];
		HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_3_1_2_1_ct2_1);
		HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCtA6_Pl_01_1);
		if (iRow == hTKKCMCCellInfo.CellTemplate.RowIndex && iCol == hTKKCMCCellInfo.CellTemplate.ColIndex)
		{
			sheetView2.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value;
		}
		HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_3_1_2_1_ct2_2_ten);
		HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_3_1_2_1_ct2_2_ma);
		HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathTenCtA7_Pl_01_1);
		HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMaCtA7_Pl_01_1);
		if (iRow == hTKKCMCCellInfo3.CellTemplate.RowIndex && iCol == hTKKCMCCellInfo3.CellTemplate.ColIndex)
		{
			sheetView2.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Value = sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value;
			sheetView2.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Value = sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value;
		}
		HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_3_1_2_1_ct1_1);
		HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCtA9_Pl_01_1);
		if (iRow == hTKKCMCCellInfo7.CellTemplate.RowIndex && iCol == hTKKCMCCellInfo7.CellTemplate.ColIndex)
		{
			sheetView2.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Value = sheetView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Value;
		}
		HTKKCMCCellInfo hTKKCMCCellInfo9 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_3_1_2_1_ct1_2_ten);
		HTKKCMCCellInfo hTKKCMCCellInfo10 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == path_3_1_2_1_ct1_2_ma);
		HTKKCMCCellInfo hTKKCMCCellInfo11 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathTenCtA10_Pl_01_1);
		HTKKCMCCellInfo hTKKCMCCellInfo12 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMaCtA10_Pl_01_1);
		if (iRow == hTKKCMCCellInfo9.CellTemplate.RowIndex && iCol == hTKKCMCCellInfo9.CellTemplate.ColIndex)
		{
			sheetView2.Cells[hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1].Value = sheetView.Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Value;
			sheetView2.Cells[hTKKCMCCellInfo12.CellTemplate.RowIndex - 1, hTKKCMCCellInfo12.CellTemplate.ColIndex - 1].Value = sheetView.Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Value;
		}
		HTKKCMCCellInfo hTKKCMCCellInfo13 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc3_4_ct4);
		HTKKCMCCellInfo hTKKCMCCellInfo14 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCtB1_Pl_01_1);
		if (iRow == hTKKCMCCellInfo13.CellTemplate.RowIndex && iCol == hTKKCMCCellInfo13.CellTemplate.ColIndex)
		{
			sheetView2.Cells[hTKKCMCCellInfo14.CellTemplate.RowIndex - 1, hTKKCMCCellInfo14.CellTemplate.ColIndex - 1].Value = sheetView.Cells[hTKKCMCCellInfo13.CellTemplate.RowIndex - 1, hTKKCMCCellInfo13.CellTemplate.ColIndex - 1].Value;
		}
		HTKKCMCCellInfo hTKKCMCCellInfo15 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathMuc_3_6_ct4);
		HTKKCMCCellInfo hTKKCMCCellInfo16 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCtC_Pl_01_1);
		if (iRow == hTKKCMCCellInfo15.CellTemplate.RowIndex && iCol == hTKKCMCCellInfo15.CellTemplate.ColIndex)
		{
			sheetView2.Cells[hTKKCMCCellInfo16.CellTemplate.RowIndex - 1, hTKKCMCCellInfo16.CellTemplate.ColIndex - 1].Value = sheetView.Cells[hTKKCMCCellInfo15.CellTemplate.RowIndex - 1, hTKKCMCCellInfo15.CellTemplate.ColIndex - 1].Value;
		}
		HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath_4_2_1);
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		if (iRow >= dataLocationTemplateRowStart && iRow <= dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count)
		{
			sheetView2.Cells[cellB1[1], cellB1[0]].Value = sheetView2.Cells[cellB1Formular[1], cellB1Formular[0]].Value;
		}
	}

	public void TinhCongThucChoSua_Pl01_2(SheetView sView, HTKKCMCSheetInfo sInfo, int iRow, int iCol)
	{
		if (iCol == HTKKFile.ColumnLettersToIndex("T") || iCol == HTKKFile.ColumnLettersToIndex("X"))
		{
			GeneralFunctions.CapNhatCongThucLoaiChoSuaCell(sView, sInfo, iRow, iCol, mapCongThuc);
		}
	}

	public void LockUnLock_Pl_01_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			SheetView sheetView = winFpSpread.Sheets[0];
			SheetView sheetView2 = winFpSpread.Sheets[1];
			int num = winFpSpread.ActiveSheet.ActiveColumnIndex + 1;
			int num2 = winFpSpread.ActiveSheet.ActiveRowIndex + 1;
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCt10_KhacVnd_Pl_01_1);
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCt10_Vnd_Pl_01_1);
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCt11_TiGia_Pl_01_1);
			string text = sheetView.Cells[cell_MaDonViTien[1], cell_MaDonViTien[0]].Text;
			bool flag = true;
			bool flag2 = false;
			if (num2 == hTKKCMCCellInfo.CellTemplate.RowIndex && num == hTKKCMCCellInfo.CellTemplate.ColIndex && Convert.ToBoolean(sheetView2.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value))
			{
				sheetView2.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = false;
				flag2 = true;
				flag = false;
			}
			else if (num2 == hTKKCMCCellInfo2.CellTemplate.RowIndex && num == hTKKCMCCellInfo2.CellTemplate.ColIndex && Convert.ToBoolean(sheetView2.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value))
			{
				sheetView2.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = false;
				flag = true;
				flag2 = false;
			}
			if (hTKKCMCCellInfo != null)
			{
				if (!string.IsNullOrEmpty(text) && text == "VND")
				{
					sheetView2.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = true;
					sheetView2.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = false;
					GeneralFunctions.LockUnLockCol(sheetView2, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo.CellTemplate.ColIndex, locked: true, deleteData: false);
					GeneralFunctions.LockUnLockCol(sheetView2, hTKKCMCCellInfo2.CellTemplate.RowIndex, hTKKCMCCellInfo2.CellTemplate.ColIndex, locked: true, deleteData: false);
				}
				else
				{
					GeneralFunctions.LockUnLockCol(sheetView2, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo.CellTemplate.ColIndex, locked: false, deleteData: false);
					GeneralFunctions.LockUnLockCol(sheetView2, hTKKCMCCellInfo2.CellTemplate.RowIndex, hTKKCMCCellInfo2.CellTemplate.ColIndex, locked: false, deleteData: false);
				}
			}
			if (hTKKCMCCellInfo3 != null)
			{
				if (!string.IsNullOrEmpty(text) && text != "VND" && flag)
				{
					GeneralFunctions.lockUnlockMultiCellsStatic(sheetView2, hTKKCMCSheetInfo, hTKKCMCCellInfo3.CellTemplate.RowIndex, hTKKCMCCellInfo3.CellTemplate.RowIndex, isLock: false, isDeleteData: false);
				}
				else
				{
					GeneralFunctions.lockUnlockMultiCellsStatic(sheetView2, hTKKCMCSheetInfo, hTKKCMCCellInfo3.CellTemplate.RowIndex, hTKKCMCCellInfo3.CellTemplate.RowIndex, isLock: true, isDeleteData: true);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == pathCtE_Pl_01_1);
			if (hTKKCMCCellInfo4 != null)
			{
				if (!string.IsNullOrEmpty(text) && text != "VND" && flag)
				{
					GeneralFunctions.LockUnLockCol(sheetView2, hTKKCMCCellInfo4.CellTemplate.RowIndex, hTKKCMCCellInfo4.CellTemplate.ColIndex, locked: false, deleteData: false);
				}
				else
				{
					GeneralFunctions.LockUnLockCol(sheetView2, hTKKCMCCellInfo4.CellTemplate.RowIndex, hTKKCMCCellInfo4.CellTemplate.ColIndex, locked: true, deleteData: true);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
