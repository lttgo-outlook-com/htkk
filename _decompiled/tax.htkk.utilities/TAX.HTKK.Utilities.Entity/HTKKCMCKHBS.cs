namespace TAX.HTKK.Utilities.Entity;

public class HTKKCMCKHBS
{
	public HTKKCMCCellBasicInfo Cell { get; set; }

	public string MaChiTieu { get; set; }

	public string TenChiTieu { get; set; }

	public string DieuChinhTang { get; set; }

	public HTKKCMCKHBS()
	{
		Cell = new HTKKCMCCellBasicInfo();
		MaChiTieu = string.Empty;
		TenChiTieu = string.Empty;
		DieuChinhTang = string.Empty;
	}

	public HTKKCMCKHBS(HTKKCMCKHBS _HTKKCMCKHBS)
	{
		Cell = new HTKKCMCCellBasicInfo(_HTKKCMCKHBS.Cell);
		MaChiTieu = _HTKKCMCKHBS.MaChiTieu;
		TenChiTieu = _HTKKCMCKHBS.TenChiTieu;
		DieuChinhTang = _HTKKCMCKHBS.TenChiTieu;
	}
}
