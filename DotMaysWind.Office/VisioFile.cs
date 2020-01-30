using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

using DotMaysWind.Office.Entity;
using DotMaysWind.Office.Helper;

namespace DotMaysWind.Office
{
    public class VisioFile : CompoundBinaryFile, IVisioFile
    {
        // really, this should be a collection of cells which contains a location and content
        private string _allText = "no text yet";
        public string AllText
        {
            get { return _allText; }
        }

        public VisioFile(String filePath) : base(filePath) { 
        }
    }
}