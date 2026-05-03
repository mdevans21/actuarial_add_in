using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace ActuarialAddIn;

[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    private static Microsoft.Office.Interop.Excel.Application? _excelApp;

    public void OnRibbonLoad(IRibbonUI ribbon)
    {
        _excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
    }

    public void OnInsertVersionsSheet(IRibbonControl control)
    {
        try
        {
            CreateOrUpdateVersionsSheet(insertAtLeft: true);
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show(
                $"Error creating Versions sheet: {ex.Message}",
                "Actuarial Add-In",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error);
        }
    }

    public void OnRefreshVersionsSheet(IRibbonControl control)
    {
        try
        {
            CreateOrUpdateVersionsSheet(insertAtLeft: false);
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show(
                $"Error refreshing Versions sheet: {ex.Message}",
                "Actuarial Add-In",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error);
        }
    }

    public void OnShowAbout(IRibbonControl control)
    {
        var version = VersionInfo.CurrentVersion;
        var message = $"Actuarial Add-In\n\n" +
                      $"Version: {version}\n" +
                      $"Build Date: {VersionInfo.BuildDate}\n\n" +
                      $"A comprehensive Excel add-in for actuarial calculations\n" +
                      $"including distributions, reserving, copulas, and more.\n\n" +
                      $"GitHub: {VersionInfo.GitHubUrl}";

        System.Windows.Forms.MessageBox.Show(
            message,
            "About Actuarial Add-In",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Information);
    }

    public void OnOpenGitHub(IRibbonControl control)
    {
        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = VersionInfo.GitHubUrl,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show(
                $"Error opening GitHub: {ex.Message}",
                "Actuarial Add-In",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error);
        }
    }

    private void CreateOrUpdateVersionsSheet(bool insertAtLeft)
    {
        if (_excelApp == null)
            _excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;

        var workbook = _excelApp.ActiveWorkbook;
        if (workbook == null)
        {
            System.Windows.Forms.MessageBox.Show(
                "Please open a workbook first.",
                "Actuarial Add-In",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Warning);
            return;
        }

        Microsoft.Office.Interop.Excel.Worksheet? versionsSheet = null;

        // Check if Versions sheet already exists
        foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.Name == "Versions")
            {
                versionsSheet = sheet;
                break;
            }
        }

        // Create new sheet if it doesn't exist
        if (versionsSheet == null)
        {
            if (insertAtLeft)
            {
                versionsSheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.Add(Before: workbook.Worksheets[1]);
            }
            else
            {
                versionsSheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.Add();
            }
            versionsSheet.Name = "Versions";
        }
        else if (insertAtLeft)
        {
            // Move to leftmost position if requested
            versionsSheet.Move(Before: workbook.Worksheets[1]);
        }

        // Clear existing content
        versionsSheet.Cells.Clear();

        // Populate the sheet
        PopulateVersionsSheet(versionsSheet);

        // Activate the sheet
        versionsSheet.Activate();
    }

    private void PopulateVersionsSheet(Microsoft.Office.Interop.Excel.Worksheet sheet)
    {
        int row = 1;

        // Title
        sheet.Cells[row, 1] = "Actuarial Add-In Version History";
        var titleRange = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, 4]];
        titleRange.Merge();
        titleRange.Font.Bold = true;
        titleRange.Font.Size = 16;
        row += 2;

        // Current version info
        sheet.Cells[row, 1] = "Current Version:";
        sheet.Cells[row, 2] = VersionInfo.CurrentVersion;
        sheet.Cells[row, 1].Font.Bold = true;
        row++;

        sheet.Cells[row, 1] = "Build Date:";
        sheet.Cells[row, 2] = VersionInfo.BuildDate;
        sheet.Cells[row, 1].Font.Bold = true;
        row++;

        sheet.Cells[row, 1] = "GitHub:";
        sheet.Cells[row, 2] = VersionInfo.GitHubUrl;
        sheet.Cells[row, 1].Font.Bold = true;

        // Make GitHub URL a hyperlink
        var linkCell = sheet.Cells[row, 2];
        sheet.Hyperlinks.Add(linkCell, VersionInfo.GitHubUrl);
        row += 2;

        // Commit history header
        sheet.Cells[row, 1] = "Commit History";
        var historyTitleRange = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, 4]];
        historyTitleRange.Merge();
        historyTitleRange.Font.Bold = true;
        historyTitleRange.Font.Size = 14;
        row += 1;

        // Column headers
        sheet.Cells[row, 1] = "Commit";
        sheet.Cells[row, 2] = "Date";
        sheet.Cells[row, 3] = "Message";
        var headerRange = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, 3]];
        headerRange.Font.Bold = true;
        headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
        row++;

        // Commit data
        foreach (var commit in VersionInfo.GetCommitHistory())
        {
            sheet.Cells[row, 1] = commit.ShortHash;
            sheet.Cells[row, 2] = commit.Date;
            sheet.Cells[row, 3] = commit.Message;
            row++;
        }

        // Auto-fit columns
        sheet.Columns["A:C"].AutoFit();

        // Set minimum widths
        if (((Microsoft.Office.Interop.Excel.Range)sheet.Columns["C"]).ColumnWidth < 50)
        {
            ((Microsoft.Office.Interop.Excel.Range)sheet.Columns["C"]).ColumnWidth = 50;
        }
    }
}
