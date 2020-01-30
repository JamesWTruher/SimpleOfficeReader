using System;
using DotMaysWind.Office;
using System.Management.Automation;

namespace DotMaysWind.Office.Cmdlets
{
    public class OfficeFileInfo
    {
        public string Path {get; set; }
        public string Author { get; set; }
        public string Title { get; set; }
        public string LastAuthor {get; set; }
        public string FileType { get; set; }
        public DateTime Created {get; set;}
        public DateTime LastSaved { get; set; }
        public IOfficeFile originalObject { get; set; }
        /*
CodePage         1252
Title            020702 - Virtual Team Meeting
Subject          
Author           James W. Truher
Keyword          
Commenct         
Template         Normal.dot
LastAuthor       James W. Truher
Reversion        2
ApplicationName  Microsoft Word 10.0
EditTime         12/31/1600 4:15:00 PM
CreateDateTime   7/9/2002 8:40:00 AM
LastSaveDateTime 7/9/2002 8:40:00 AM
PageCount        1
WordCount        259
CharCount        1477
Security         0
    */
    }
    [Cmdlet(VerbsCommon.Get, "OfficeFileInfo")]
    public class GetOfficeFileInfoCommand : PSCmdlet
    {
        [Parameter(Mandatory=true,Position=0,ValueFromPipeline=true,ValueFromPipelineByPropertyName=true)]
        public string [] FileName { get; set; }

        protected override void ProcessRecord()
        {
            foreach (string file in FileName)
            {
                ProviderInfo pi;
                foreach (string resolvedFile in GetResolvedProviderPathFromPSPath(file, out pi))
                {
                    IOfficeFile info;
                    try {
                        info = OfficeFileFactory.CreateOfficeFile(resolvedFile);
                        var fileInfo = new OfficeFileInfo()
                        {
                            Path = resolvedFile,
                            originalObject = info
                        };
                        string author;
                        if (info.SummaryInformation.TryGetValue("Author", out author)) {
                            fileInfo.Author = author;
                        }
                        else {
                            fileInfo.Author = "unknown";
                        }
                        string title;
                        if (info.SummaryInformation.TryGetValue("Title", out title)) {
                            fileInfo.Title = title;
                        }
                        string lastAuthor;
                        if (info.SummaryInformation.TryGetValue("LastAuthor", out lastAuthor)) {
                            fileInfo.LastAuthor = lastAuthor;
                        }
                        string applicationName;
                        if (info.SummaryInformation.TryGetValue("ApplicationName", out applicationName)) {
                            fileInfo.FileType = applicationName;
                        }
                        string createdDate;
                        if (info.SummaryInformation.TryGetValue("CreateDateTime", out createdDate)) {
                            fileInfo.Created = DateTime.Parse(createdDate);
                        }
                        string lastChange;
                        if (info.SummaryInformation.TryGetValue("LastSaveDateTime", out lastChange)) {
                            fileInfo.LastSaved = DateTime.Parse(lastChange);
                        }
                        WriteObject(fileInfo);
                    }
                    catch (Exception e) {
                        ErrorRecord er = new ErrorRecord(e, "GetOfficeFileInfoCommand", ErrorCategory.InvalidData, resolvedFile);
                        WriteError(er);
                        continue;
                    }
                }
            }
        }

    }
}