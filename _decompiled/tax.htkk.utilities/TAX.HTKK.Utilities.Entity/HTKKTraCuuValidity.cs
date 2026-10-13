namespace TAX.HTKK.Utilities.Entity;

public class HTKKTraCuuValidity
{
	public string StartDate { get; set; }

	public string EndDate { get; set; }

	public string DataFile { get; set; }

	public string ThueKhauTru { get; set; }

	public string ThuePhaiNop { get; set; }

	public HTKKTraCuuValidity()
	{
		StartDate = string.Empty;
		EndDate = string.Empty;
		DataFile = string.Empty;
		ThueKhauTru = string.Empty;
		ThuePhaiNop = string.Empty;
	}
}
