namespace ExcelTaxReport.Models
{
    /// <summary>
    /// client specific values
    /// </summary>
    public class ClientConfig
    {
        public int id { get; set; }
        public int client_id { get; set; }
        public string report_name { get; set; }
        public string base_path { get; set; }
        public string logo { get; set; }
    }
}
