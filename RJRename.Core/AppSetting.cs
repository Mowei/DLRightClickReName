using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Text;

namespace RJRename.Core
{
    public class AppSetting
    {
        public string NameFormatTemplate { get; set; }
        public string WorkNameXPath { get; set; }
        public string MakerNameXPath { get; set; }
        public string SaleDateXPath { get; set; }
        public string WorkGenreXPath { get; set; }

        public string ProductSampleImagesXPath { get; set; }

        public Dictionary<string, string> GenreTypes { get; set; }

        public Dictionary<string, string> CustomXPath { get; set; }
    }
}
