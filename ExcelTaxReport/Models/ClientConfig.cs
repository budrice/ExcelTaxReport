namespace ExcelTaxReport.Models
{
    public class ClientConfig
    {
        public int id { get; set; }
        public int client_id { get; set; }
        public string report_name { get; set; }
        public string base_path { get; set; }
        public string logo { get; set; }
        //public string date_due_field { get; set; }
        //public byte detailed_additional_info { get; set; }
    }
}
