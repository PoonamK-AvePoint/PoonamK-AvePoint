
using System.Collections.Generic;

namespace ADB.CopyDocument.Service.Models
{
    public class Parameters
    {
        public string SourceSiteUrl { get; set; }
        public string SourceFileUrl { get; set; }
        public string DestinationSiteUrl { get; set; }
        public string DestinationLibrary { get; set; }
        public string DestinationFolder { get; set; }
        public string DestinationFileName { get; set; }
        public Dictionary<string, string> MetadataForDestinationFile { get; set; }
        public bool IsMove { get; set; }
    }
}