using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace ReportLib.ReportData
{
    [System.Serializable]
    [XmlType]
    public partial class Reports
    {
        [XmlElement("Report")]
        public List<Report> Items { get; set; }
    }

    [System.Serializable]
    public partial class Report
    {
        [XmlElement]
        public string Name { get; set; }

        [XmlElement("ReportVal")]
        public List<ReportVal> ReportVal { get; set; }
    }

    [System.Serializable]
    public partial class ReportVal
    {
        [XmlElement]
        public double ReportRow { get; set; }

        [XmlElement]
        public double ReportCol { get; set; }

        [XmlElement]
        public double Val { get; set; }
    }
}
