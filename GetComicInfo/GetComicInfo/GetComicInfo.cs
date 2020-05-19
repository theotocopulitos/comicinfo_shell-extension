using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml;

/// <summary>
/// The GetComicInfo is an example shell context menu extension,
/// implemented with SharpShell. It extracts comic info from cbz files with 
/// embeded comicinfo.xml file.
/// </summary>

[ComVisible(true)]
[COMServerAssociation(AssociationType.ClassOfExtension, ".cbz")]
public class GetComicInfoExtension : SharpContextMenu
{
    /// <summary>
    /// Determines whether this instance can a shell
    /// context show menu, given the specified selected file list.
    /// </summary>
    /// <returns>
    /// <c>true</c> if this instance should show a shell context
    /// menu for the specified file list; otherwise, <c>false</c>.
    /// </returns>
    protected override bool CanShowMenu()
    {
        //  We always show the menu.
        return true;
    }

    /// <summary>
    /// Creates the context menu. This can be a single menu item or a tree of them.
    /// </summary>
    /// <returns>
    /// The context menu for the shell context menu.
    /// </returns>
    protected override ContextMenuStrip CreateMenu()
    {
        //  Create the menu strip.
        var menu = new ContextMenuStrip();

        //  Create a 'getInfo' item.
        var itemGetInfo = new ToolStripMenuItem
        {
            Text = "Comic Info...",
            //Image = Properties.Resources.CountLines
        };

        //  When we click, we'll get the comics info.
        itemGetInfo.Click += (sender, args) => GetInfo();

        //  Add the item to the context menu.
        menu.Items.Add(itemGetInfo);

        //  Return the menu.
        return menu;
    }

    /// <summary>
    /// Extracts info from comicinfo.xml embedded file.
    /// </summary>
    private void GetInfo()
    {
        //  Builder for the output.
        var builder = new StringBuilder();

        //  Go through each file.
        foreach (var filePath in SelectedItemPaths)
        {
            // loads the zip file
            ZipArchive archivo = ZipFile.OpenRead(filePath);

            // gets the comicinfo.xml file - add error handling
            ZipArchiveEntry entry = archivo.GetEntry("ComicInfo.xml");
            
            if (entry != null) {
            // opens the comicinfo.xml file              
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(entry.Open());
                
            XmlNodeList SeriesList = xmlDoc.GetElementsByTagName("Series");
            XmlNodeList VolumeList = xmlDoc.GetElementsByTagName("Volume");
            XmlNodeList IssueList = xmlDoc.GetElementsByTagName("Number");
            XmlNodeList TitleList = xmlDoc.GetElementsByTagName("Title");
            //XmlNodeList DayList = xmlDoc.GetElementsByTagName("Day");
            XmlNodeList MonthList = xmlDoc.GetElementsByTagName("Month");
            XmlNodeList YearList = xmlDoc.GetElementsByTagName("Year");
            XmlNodeList WriterList = xmlDoc.GetElementsByTagName("Writer");
            XmlNodeList PencillerList = xmlDoc.GetElementsByTagName("Penciller");
            XmlNodeList SummaryList = xmlDoc.GetElementsByTagName("Summary");


            // info
            builder.AppendLine(string.Format("{0} ({1}) #{2} ({3}-{4})\n\n\"{5}\"\n\n\n{6} | {7}\n\n\nSummary:\n\n {8}",
                        SeriesList[0].InnerXml, 
                        VolumeList[0].InnerXml,
                        IssueList[0].InnerXml,
                        YearList[0].InnerXml,
                        MonthList[0].InnerXml,
                        //DayList[0].InnerXml,
                        TitleList[0].InnerXml,
                        WriterList[0].InnerXml,
                        PencillerList[0].InnerXml,
                        SummaryList[0].InnerXml));
                        
           //  Show the ouput.                   
                    MessageBox.Show(builder.ToString());
                }

        }
    }
}
