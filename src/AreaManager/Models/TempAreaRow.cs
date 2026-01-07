using System;

namespace AreaManager.Models
{
    public class TempAreaRow
    {
        public string Description { get; set; }
        public string Identifier { get; set; }
        public string Width { get; set; }
        public string Length { get; set; }
        public string AreaHa { get; set; }
        public string WithinExistingDisposition { get; set; }
        public string ExistingCutDisturbance { get; set; }
        public string NewCutDisturbance { get; set; }
    }
}
