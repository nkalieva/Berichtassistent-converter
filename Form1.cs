using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using Action = System.Action;
using Application = System.Windows.Forms.Application;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;
using OfficeOpenXml;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace Resave
{
    public partial class Form1 : Form

    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        private int smallHeight = 140; // Initial height
        private int largeHeight = 300; // Height with progress bars
        private DateTime startTime;  // Processing start time
        private int processedComments = 0;  // Number of processed comments
        private double averageTimePerComment = 0;  // Average processing time per comment (in seconds)
        public Form1()
        {
            InitializeComponent();
            this.Text = "Excel Datei";
            this.Height = smallHeight; 
        }
        private void UpdateFormSize()
        {
            int baseHeight = 150; 
            int progressBarHeight = 40; 
            int labelHeight = 20; 

            int extraHeight = 0;

            if (progressBar1.Visible) extraHeight += progressBarHeight;
            if (progressBar2.Visible) extraHeight += progressBarHeight;
            if (label1.Visible) extraHeight += labelHeight;

            this.Height = baseHeight + extraHeight;
        }

        private async void btnShow_Click(object sender, EventArgs e)
        {

            OpenFileDialog dlg = new OpenFileDialog
            {
                Title = "Datei auswählen",
                Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
            };

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                if (dlg.FileName.EndsWith(".xlsx"))
                {
                  
                    btnShow.Text = dlg.FileName;
                    progressBar1.Style = ProgressBarStyle.Continuous;
                    progressBar1.Visible = true;
                    progressBar1.Value = 0;
                    lblProgress1.Visible = true;
                    lblProgress1.Text = "0%";
                    btnShow.Enabled = false;
                    lblText.Visible = true;
                    lblText.Text = "Suche nach Kommentaren...";
                    label1.Visible = false;

                    var progress = new Progress<int>(value =>
                    {
                        progressBar1.Value = value;
                        lblProgress1.Text = $"{value}%";
                    });

                    bool commentsFound = false;
                    try
                    {
                        UpdateFormSize();
                        commentsFound = await ReadExcelComments(dlg.FileName, progress);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Fehler: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        progressBar1.Visible = false;
                        lblProgress1.Visible = false;
                        lblText.Visible = false;
                        UpdateFormSize();

                    }

                    if (!commentsFound)
                    {
                        label1.Visible = true;
                        //MessageBox.Show("Keine Kommentare gefunden.", "Ergebnis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                       
                        btnShow.Enabled = true;
                        btnSave.Enabled = false;
                        UpdateFormSize();
                        return;
                    }
                    label1.Visible = false;
                    btnSave.Enabled = true;
                    btnShow.Enabled = true;
                }
                else
                {
                    MessageBox.Show("Bitte wählen Sie eine Excel-Datei aus.", "Ungültige Datei", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private async Task<bool> ReadExcelComments(string filePath, IProgress<int> progress)
        {
            ExcelApp excelApp = new ExcelApp();
            Workbook workbook = null;
            bool foundComments = false;

            try
            {
                workbook = excelApp.Workbooks.Open(filePath);
                int totalSheets = workbook.Sheets.Count;
                int processedSheets = 0;

                foreach (Worksheet worksheet in workbook.Sheets)
                {
                    UpdateFormSize();
                    processedSheets++;
                    int percent = (int)((processedSheets / (double)totalSheets) * 100);
                    progress.Report(Math.Max(1, Math.Min(percent, 100)));

                    Invoke(new Action(() =>
                    {
                        // lblText1.Text = $"Prüfe: {worksheet.Name} ({processedSheets}/{totalSheets})...";
                        Application.DoEvents();
                    }));


                    try
                    {
                        Comments comments = worksheet.Comments;
                        if (comments.Count > 0)
                        {
                            foreach (Comment comment in comments)
                            {
                                Range cell = comment.Parent as Range;
                                string commentText = comment.Text();
                                if (!string.IsNullOrEmpty(commentText))
                                {
                                    foundComments = true;
                                    progress.Report(100);
                                    //MessageBox.Show("Kommentare gefunden!", "Gefundene Mandant Kommentare", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    break;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Fehler beim Verarbeiten von {worksheet.Name}: {ex.Message}");
                        continue;
                    }

                    Marshal.ReleaseComObject(worksheet);
                    if (foundComments) break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler beim Auslesen der Excel-Datei: " + ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                workbook?.Close(false);
                excelApp.Quit();
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                UpdateFormSize();
            }

            return foundComments;
        }

        private Task RunInSTAThread(Action action)
        {
            var tcs = new TaskCompletionSource<object>();

            Thread thread = new Thread(() =>
            {
                try
                {
                    action();
                    tcs.SetResult(null);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

            return tcs.Task;
        }
        private async void btnSave_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                Title = "Notizen speichern"
            };

            if (saveFileDialog1.ShowDialog() != DialogResult.OK) return;

            string savePath = saveFileDialog1.FileName;

            // UI preparation

            progressBar1.Value = 0;
            progressBar1.Style = ProgressBarStyle.Blocks;
            progressBar1.Visible = true;
            lblProgress1.Visible = true;
            lblProgress1.Text = "0%";

            progressBar2.Value = 0;
            progressBar2.Style = ProgressBarStyle.Continuous;
            progressBar2.Visible = false;
            lblProgress2.Visible = false;
            lblProgress2.Text = "0%";

            lblText.Visible = true;
            btnShow.Enabled = false;
            btnSave.Enabled = false;
            lblText.Text = "Tabellenblätter werden verarbeitet:";

            try
            {
                UpdateFormSize();
                await RunInSTAThread(() =>
                {
                  
                    ExcelApp excelApp = new ExcelApp { DisplayAlerts = false };
                    Workbook sourceWorkbook = null;
                    Workbook newWorkbook = null;
  
                    try
                    {
                        sourceWorkbook = excelApp.Workbooks.Open(btnShow.Text);
                        newWorkbook = excelApp.Workbooks.Add();

                        int sheetIndex = 0;

                        foreach (Worksheet sheet in sourceWorkbook.Sheets)
                        {
                            
                            sheet.Copy(After: newWorkbook.Sheets[newWorkbook.Sheets.Count]);
                            sheetIndex++;
                            int progress = (int)((sheetIndex / (double)sourceWorkbook.Sheets.Count) * 100);
                            Invoke(new Action(() =>
                            {
                                lblProgress1.Text = $"0%";
                                lblProgress1.Text = $"{progress}%";
                                progressBar1.Value = progress;
                                lblText.Text = $"Tabellenblätter werden verarbeitet: {sheetIndex}/{sourceWorkbook.Sheets.Count}";

                                // Forced UI update before resetting progressBar2
                                Application.DoEvents();

                                // Resetting progressBar2
                                progressBar2.Value = 0;
                                UpdateFormSize();
                            }));
                        }

                            if (newWorkbook.Sheets.Count > 1)
                        {
                            Worksheet firstSheet = (Worksheet)newWorkbook.Sheets[1];
                            firstSheet.Delete();
                        }

                        int totalSheets = newWorkbook.Sheets.Count; 
                        int currentSheetIndex = 0;
                        int totalComments = 0;
                        int processedComments = 0;
                        bool valuesFound = false;
                        startTime = DateTime.Now;

                        int estimatedTimePerComment = 1; // 1 second per comment
                                                         //totalEstimatedMinutes = (totalComments / 60); // Estimated total time in minutes

                        // Iterate through all sheets and count the number of comments
                        foreach (Worksheet worksheet in newWorkbook.Sheets)
                        {
                            totalComments += worksheet.Comments.Count;
                        }

                        if (totalComments == 0)
                        {
                            MessageBox.Show("Keine Kommentare gefunden.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }

                        // Set the progress step
                        int lastReportedProgress1 = 0;

                        // Process each sheet
                        foreach (Worksheet worksheet in newWorkbook.Sheets)
                        {
                            currentSheetIndex++;

                            // Update UI for all sheets
                            Invoke(new Action(() =>
                            {
                                progressBar2.Visible = true;
                                lblProgress2.Visible = true;
                                lblText1.Visible = true;
                                lblProgress2.Text = $"0%";
                                lblText.Text = $"Datei wird verarbeitet. Tabellenblatt: {currentSheetIndex}/{totalSheets} ";
                                lblText1.Text = $"Aktuelles Tabellenblatt: {worksheet.Name}";
                                progressBar2.Value = 0;
                                UpdateFormSize();
                            }));
                           

                            try
                            {
                                // If there are no comments on the sheet, just update the progress
                                Comments comments = worksheet.Comments;
                                int sheetTotalComments = comments.Count;
                                int sheetProcessedComments = 0;

                                // Process comments (if any)
                                foreach (Comment comment in comments)
                                {
                                    processedComments++;
                                    sheetProcessedComments++;

                                    // Update progress for the current sheet
                                    int progressSheet = (int)((sheetProcessedComments / (double)sheetTotalComments) * 100);
                                    Invoke(new Action(() =>
                                    {
                                        progressBar2.Value = progressSheet;
                                        lblProgress2.Text = $"{progressSheet}%";
                                    }));

                                    Range cell = comment.Parent as Range;
                                    string commentText = comment.Text();

                                    if (!string.IsNullOrEmpty(commentText))
                                    {
                                        var data = new Dictionary<string, string>
                        {
                            { "Mandant", ExtractValue(commentText, "Mandant:") },
                            { "BerichtsNr", ExtractValue(commentText, "BerichtsNr:") },
                            { "ZeilenNr", ExtractValue(commentText, "ZeilenNr:") },
                            { "UStrukt", ExtractValue(commentText, "UStrukt:") },
                            { "GkvUkv", ExtractValue(commentText, "GKV/UKV:") },
                            { "Datenbasis", ExtractValue(commentText, "Datenbasis:") },
                            { "IndexNr", ExtractValue(commentText, "IndexNr:") },
                            { "Summe", ExtractValue(commentText, "Summe:") },
                            { "Per", ExtractValue(commentText, "Per:") },
                            { "Jahr", ExtractValue(commentText, "Jahr:") },
                            { "Faktor", ExtractValue(commentText, "Faktor:") }
                        };
                                        var formula = GenerateFormula(data);
                                        if (!string.IsNullOrEmpty(formula))
                                        {
                                            SetExcelFormula(cell, formula);
                                            cell.Comment?.Delete();
                                            valuesFound = true;
                                        }
                                    }
                                }

                                Invoke(new Action(() =>
                                {
                                    progressBar2.Value = 100;
                                    lblProgress2.Text = $"100%";
                                }));
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Fehler beim Verarbeiten von {worksheet.Name}: {ex.Message}");
                            }
                            Invoke(new Action(() =>
                            {
                                UpdateRemainingTotalTime(processedComments, totalComments);
                            }));

                            Marshal.ReleaseComObject(worksheet);
                            // Update progress across all sheets
                            int progressSheetGlobal = (int)((currentSheetIndex / (double)totalSheets) * 100);
                            Invoke(new Action(() =>
                            {
                                progressBar1.Value = progressSheetGlobal;
                                lblProgress1.Text = $"{progressSheetGlobal}%";
                            }));

                            Marshal.ReleaseComObject(worksheet);
                        }

                        if (!valuesFound)
                        {
                            Invoke(new Action(() =>
                            {
                                MessageBox.Show("Keine relevanten Werte gefunden.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }));
                        }
                        ConfirmAllBA2Formulas(newWorkbook);
                        newWorkbook.SaveAs(savePath);
                       
                    }
                    finally
                    {
                        sourceWorkbook?.Close(false);
                        newWorkbook?.Close(false);
                        Marshal.ReleaseComObject(sourceWorkbook);
                        Marshal.ReleaseComObject(newWorkbook);
                        Marshal.ReleaseComObject(excelApp);
                      
                    }
                  
                });

                MessageBox.Show("Datei erfolgreich gespeichert!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Speichern der Datei: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                progressBar1.Visible = false;
                progressBar2.Visible = false;
                lblProgress1.Visible = false;
                lblProgress2.Visible = false;
                lblText.Visible = false;
                lblText1.Visible = false;
                lblText.Text = string.Empty;
                btnShow.Text = "Zu konvertierende Excel-Datei öffnen";
                btnShow.Enabled = true;
                btnSave.Enabled = false;
                UpdateFormSize();
            }
        }

        private void UpdateRemainingTotalTime(int processedComments, int totalComments)
        {
            if (totalComments == 0) return; 

        
            double elapsedTime = (DateTime.Now - startTime).TotalSeconds;
            processedComments++;

            if (processedComments == 1)
            {
                averageTimePerComment = elapsedTime;
            }
            else
            {
                averageTimePerComment = elapsedTime / processedComments;
            }

            double remainingTimeInSeconds = (totalComments - processedComments) * averageTimePerComment;

            int remainingMinutes = (int)(remainingTimeInSeconds / 60);
            int remainingSeconds = (int)(remainingTimeInSeconds % 60);

            this.Text = $"Excel Datei - Noch {remainingMinutes} Minuten {remainingSeconds} Sekunden";
        }

        private string GenerateFormula(Dictionary<string, string> data)
        {
            string separator = ";"; // в DE Excel

            var nonEmptyValues = data.Values
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Select(v => v.Any(char.IsLetter) || v.Contains('.') || v.Contains(" ") ? $"\"{v}\"" : v)
                .ToList();

            return nonEmptyValues.Count > 0
                ? $"=BA2.DATEN_BERICHTZEILE({string.Join(separator, nonEmptyValues)})"
                : string.Empty;
        }


        private void ConfirmFormula(Range cell)
        {
            if (cell == null) return;

            try
            {
                var app = cell.Application;
                var hwnd = (IntPtr)app.Hwnd;

                // Select the sheet and the cell
                var sheet = cell.Worksheet;
                sheet.Activate();
                cell.Select();

                SetForegroundWindow(hwnd);
                Thread.Sleep(100); 
             
                SendKeys.SendWait("{F2}");
                Thread.Sleep(100);
                SendKeys.SendWait("{ENTER}");

                // Let Excel finish the calculation
                Application.DoEvents();
                Thread.Sleep(100);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Fehler beim Bestätigen der Formel: {ex.Message}");
            }
        }


        private void ConfirmAllBA2Formulas(Workbook workbook)
        {
            Excel.Application excelApp = workbook.Application;

            foreach (Worksheet sheet in workbook.Sheets)
            {
                sheet.Activate();
                Range usedRange = sheet.UsedRange;

                foreach (Range cell in usedRange)
                {
                    if (cell == null) continue;

                    string formula = "";
                    try
                    {
                        formula = cell.FormulaLocal?.ToString() ?? "";
                    }
                    catch
                    {
                        continue;
                    }

                    string cleaned = formula.TrimStart('@', '\'');

                    if (!string.IsNullOrWhiteSpace(cleaned) &&
                        cleaned.StartsWith("=BA2.DATEN_BERICHTZEILE", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"Formel in {cell.Address}: {cleaned}");
                        ConfirmFormula(cell);
                        Thread.Sleep(100); 
                    }
                }

                Marshal.ReleaseComObject(sheet);
            }
        }


        private void SetExcelFormula(Range cell, string formula)
        {
            if (cell == null || string.IsNullOrWhiteSpace(formula))
                return;

            try
            {
                if (formula.StartsWith("'"))
                    formula = formula.Substring(1);

                if (!formula.TrimStart().StartsWith("="))
                    formula = "=" + formula;

                cell.FormulaLocal = formula;

                Console.WriteLine($"Formel gespeichert: {formula}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fehler beim Einfügen der Formel: {ex.Message}");
            }
        }



        private string ExtractValue(string commentText, string keyword)
        {
            if (string.IsNullOrEmpty(commentText))
            {
                return string.Empty; // If there is no comment, return an empty string

            }

            int startIndex = commentText.IndexOf(keyword);
            if (startIndex >= 0)
            {
                startIndex += keyword.Length;
                int endIndex = commentText.IndexOf("\n", startIndex);
                if (endIndex < 0) endIndex = commentText.Length;
                string value = commentText.Substring(startIndex, endIndex - startIndex).Trim();

                Console.WriteLine($"Extracting '{keyword}': {value}");

                return value;
            }
            return string.Empty; // If the keyword is not found, return an empty string
        }

        private void umwandeln_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.ShowDialog();
            btnShow.Text = dlg.SelectedPath;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load_1(object sender, EventArgs e)
        {

        }

        private void progressBar2_Click(object sender, EventArgs e)
        {

        }

        private void lblText1_Click(object sender, EventArgs e)
        {

        }

        private void lblProgress2_Click(object sender, EventArgs e)
        {

        }
    }
}

