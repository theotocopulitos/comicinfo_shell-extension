using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System;
using System.IO;
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
            try
            {
                // loads the zip file
                using (ZipArchive archivo = ZipFile.OpenRead(filePath))
                {
                    // gets the comicinfo.xml file - add error handling
                    ZipArchiveEntry entry = archivo.GetEntry("ComicInfo.xml");
                    
                    if (entry != null) 
                    {
                        // opens the comicinfo.xml file              
                        XmlDocument xmlDoc = new XmlDocument();
                        using (var stream = entry.Open())
                        {
                            xmlDoc.Load(stream);
                        }
                            
                        // Helper function to safely get XML element text
                        string GetElementText(string tagName)
                        {
                            var nodes = xmlDoc.GetElementsByTagName(tagName);
                            return nodes.Count > 0 ? nodes[0].InnerText : "N/A";
                        }

                        string series = GetElementText("Series");
                        string volume = GetElementText("Volume");
                        string issue = GetElementText("Number");
                        string title = GetElementText("Title");
                        string month = GetElementText("Month");
                        string year = GetElementText("Year");
                        string writer = GetElementText("Writer");
                        string penciller = GetElementText("Penciller");
                        string summary = GetElementText("Summary");

                        // info
                        builder.AppendLine(string.Format("{0} ({1}) #{2} ({3}-{4})\n\n\"{5}\"\n\n\n{6} | {7}\n\n\nSummary:\n\n {8}",
                                    series, 
                                    volume,
                                    issue,
                                    year,
                                    month,
                                    title,
                                    writer,
                                    penciller,
                                    summary));
                                    
                        //  Show the output.                   
                        MessageBox.Show(builder.ToString());
                    }
                    else
                    {
                        MessageBox.Show("No ComicInfo.xml found in " + Path.GetFileName(filePath));
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading comic file: " + ex.Message);
            }
        }
    }
}
