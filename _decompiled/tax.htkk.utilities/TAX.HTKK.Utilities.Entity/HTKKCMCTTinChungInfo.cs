using System;
using System.Reflection;
using TAX.HTKK.Utilities.Core;

namespace TAX.HTKK.Utilities.Entity;

public class HTKKCMCTTinChungInfo
{
	public string maDVu { get; set; }

	public string tenDVu { get; set; }

	public string pbanDVu { get; set; }

	public string ttinNhaCCapDVu { get; set; }

	public string maTKhai { get; set; }

	public string tenTKhai { get; set; }

	public string moTaBMau { get; set; }

	public string pbanTKhaiXML { get; set; }

	public string loaiTKhai { get; set; }

	public string soLan { get; set; }

	public string kieuKy { get; set; }

	public string kyKKhai { get; set; }

	public string kyKKhaiTuNgay { get; set; }

	public string kyKKhaiDenNgay { get; set; }

	public string kyKKhaiTuThang { get; set; }

	public string kyKKhaiDenThang { get; set; }

	public string maCQTNoiNop { get; set; }

	public string maCQTQuanLy => maCQTNoiNop;

	public string tenCQTNoiNop { get; set; }

	public string tenCQTQuanLy => tenCQTNoiNop;

	public string ten => tenNNT;

	public string ngayLapTKhai { get; set; }

	public int soluong => 0;

	public string SoTKNganHang { get; set; }

	public string TenDonViChuQuan { get; set; }

	public string MaSoThueDonViChuQuan { get; set; }

	public string nguoiKy { get; set; }

	public string ngayKy { get; set; }

	public string nganhNgheKD { get; set; }

	public string mst { get; set; }

	public string tenNNT { get; set; }

	public string dchiNNT { get; set; }

	public string phuongXa { get; set; }

	public string maHuyenNNT { get; set; }

	public string tenXaPhuongNNT { get; set; }

	public string maXaPhuongNNT { get; set; }

	public string tenHuyenNNT { get; set; }

	public string maTinhNNT { get; set; }

	public string tenTinhNNT { get; set; }

	public string dthoaiNNT { get; set; }

	public string faxNNT { get; set; }

	public string emailNNT { get; set; }

	public string mstDLyThue { get; set; }

	public string tenDLyThue { get; set; }

	public string dchiDLyThue { get; set; }

	public string maHuyenDLyThue { get; set; }

	public string tenHuyenDLyThue { get; set; }

	public string tenXaPhuongDlyThue { get; set; }

	public string maXaPhuongDlyThue { get; set; }

	public string tenXaPhuongDLyThue { get; set; }

	public string maXaPhuongDLyThue { get; set; }

	public string maTinhDLyThue { get; set; }

	public string tenTinhDLyThue { get; set; }

	public string dthoaiDLyThue { get; set; }

	public string faxDLyThue { get; set; }

	public string emailDLyThue { get; set; }

	public string soHDongDLyThue { get; set; }

	public string ngayKyHDDLyThue { get; set; }

	public string tenNVienDLyThue { get; set; }

	public string cchiHNghe { get; set; }

	public string tenCQTCapCuc { get; set; }

	public string maCQTCapCuc { get; set; }

	public string thang { get; set; }

	public string quy { get; set; }

	public string nam { get; set; }

	public string maDonViTinh { get; set; }

	public string tenDonViTinh { get; set; }

	public string ngay { get; set; }

	public string ngayLapTKhaiBS { get; set; }

	public HTKKCMCTTinChungInfo()
	{
		TenDonViChuQuan = string.Empty;
		SoTKNganHang = string.Empty;
		MaSoThueDonViChuQuan = string.Empty;
		maDVu = string.Empty;
		tenDVu = string.Empty;
		pbanDVu = string.Empty;
		ttinNhaCCapDVu = string.Empty;
		maTKhai = string.Empty;
		tenTKhai = string.Empty;
		moTaBMau = string.Empty;
		pbanTKhaiXML = string.Empty;
		loaiTKhai = string.Empty;
		soLan = string.Empty;
		kieuKy = string.Empty;
		kyKKhai = string.Empty;
		kyKKhaiTuNgay = string.Empty;
		kyKKhaiDenNgay = string.Empty;
		kyKKhaiTuThang = string.Empty;
		kyKKhaiDenThang = string.Empty;
		maCQTNoiNop = string.Empty;
		tenCQTNoiNop = string.Empty;
		ngayLapTKhai = string.Empty;
		nguoiKy = string.Empty;
		ngayKy = string.Empty;
		nganhNgheKD = string.Empty;
		mst = string.Empty;
		tenNNT = string.Empty;
		dchiNNT = string.Empty;
		phuongXa = string.Empty;
		maHuyenNNT = string.Empty;
		maXaPhuongNNT = string.Empty;
		tenXaPhuongNNT = string.Empty;
		tenHuyenNNT = string.Empty;
		maTinhNNT = string.Empty;
		tenTinhNNT = string.Empty;
		dthoaiNNT = string.Empty;
		faxNNT = string.Empty;
		emailNNT = string.Empty;
		mstDLyThue = string.Empty;
		tenDLyThue = string.Empty;
		dchiDLyThue = string.Empty;
		maHuyenDLyThue = string.Empty;
		maXaPhuongDlyThue = string.Empty;
		tenXaPhuongDlyThue = string.Empty;
		tenXaPhuongDLyThue = string.Empty;
		maXaPhuongDLyThue = string.Empty;
		tenHuyenDLyThue = string.Empty;
		maTinhDLyThue = string.Empty;
		tenTinhDLyThue = string.Empty;
		dthoaiDLyThue = string.Empty;
		faxDLyThue = string.Empty;
		emailDLyThue = string.Empty;
		soHDongDLyThue = string.Empty;
		ngayKyHDDLyThue = string.Empty;
		tenNVienDLyThue = string.Empty;
		cchiHNghe = string.Empty;
		tenCQTCapCuc = string.Empty;
		maCQTCapCuc = string.Empty;
		thang = string.Empty;
		quy = string.Empty;
		nam = string.Empty;
		maDonViTinh = string.Empty;
		tenDonViTinh = string.Empty;
		ngay = string.Empty;
		ngayLapTKhaiBS = string.Empty;
	}

	public HTKKCMCCellBasicInfo GetCellLocation(string location)
	{
		HTKKCMCCellBasicInfo hTKKCMCCellBasicInfo = new HTKKCMCCellBasicInfo();
		try
		{
			if (location.Length > 0)
			{
				string[] array = location.Split(new char[1] { '_' });
				if (array.Length == 2)
				{
					hTKKCMCCellBasicInfo.RowIndex = int.Parse(array[1]);
					hTKKCMCCellBasicInfo.ColIndex = int.Parse(array[0]);
				}
				else
				{
					hTKKCMCCellBasicInfo = null;
				}
			}
			else
			{
				hTKKCMCCellBasicInfo = null;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			hTKKCMCCellBasicInfo = null;
		}
		return hTKKCMCCellBasicInfo;
	}
}
