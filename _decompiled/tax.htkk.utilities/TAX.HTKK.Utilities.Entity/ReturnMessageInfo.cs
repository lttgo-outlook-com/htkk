namespace TAX.HTKK.Utilities.Entity;

public class ReturnMessageInfo
{
	public string CellIndex { get; set; }

	public string ValueCellIndex { get; set; }

	public string MessageError { get; set; }

	public string ValueMessageError { get; set; }

	public ReturnMessageInfo()
	{
		CellIndex = string.Empty;
		ValueCellIndex = string.Empty;
		MessageError = string.Empty;
		ValueMessageError = string.Empty;
	}
}
