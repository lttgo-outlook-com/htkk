namespace TAX.HTKK.Utilities.Entity;

public class HTKKCMCDataLocation
{
	public int DataLocationTemplateRowStart { get; set; }

	public int DataLocationReportRowStart { get; set; }

	public int DataLocationTaiBangKeRowStart { get; set; }

	public int TableTemplateRowStart { get; set; }

	public int TableReportRowStart { get; set; }

	public HTKKCMCDataLocation()
	{
		DataLocationTemplateRowStart = 0;
		DataLocationReportRowStart = 0;
		DataLocationTaiBangKeRowStart = 0;
		TableTemplateRowStart = 0;
		TableReportRowStart = 0;
	}

	public HTKKCMCDataLocation(HTKKCMCDataLocation dataLocation)
	{
		DataLocationTemplateRowStart = dataLocation.DataLocationTemplateRowStart;
		DataLocationReportRowStart = dataLocation.DataLocationReportRowStart;
		DataLocationTaiBangKeRowStart = dataLocation.DataLocationTaiBangKeRowStart;
		TableTemplateRowStart = dataLocation.TableTemplateRowStart;
		TableReportRowStart = dataLocation.TableReportRowStart;
	}
}
