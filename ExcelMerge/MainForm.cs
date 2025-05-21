using System;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace ExcelMerge;

public class MainForm : Form
{
    private Button selectFilesButton;
    private Button combineButton;
    private Label statusLabel;
    private OpenFileDialog openFileDialog;
    private SaveFileDialog saveFileDialog;
    private string[] excelFiles;
    
    
    public MainForm()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        this.Text = "Unificador de Excel";
        this.Size = new System.Drawing.Size(400, 200);
        this.FormBorderStyle = FormBorderStyle.FixedSingle;
        this.MaximizeBox = false;

        selectFilesButton = new Button
        {
            Text = "Seleccionar",
            Location = new System.Drawing.Point(20, 20),
            Size = new Size(150, 30),
        };
        selectFilesButton.Click += selectFilesButton_Click;

        combineButton = new Button
        {
            Text = "Combinar excel",
            Location = new System.Drawing.Point(20, 60),
            Size = new Size(150, 30),
            Enabled = false
        };
        combineButton.Click += CombineButton_Click;
        
        statusLabel = new Label
        {
            Text = "Selecciona los archivos Excel para combinar.",
            Location = new System.Drawing.Point(20, 100),
            Size = new System.Drawing.Size(350, 30)
        };
        
        openFileDialog = new OpenFileDialog
        {
            Multiselect = true,
            Filter = "Archivos Excel|*.xlsx",
            Title = "Seleccionar Archivos Excel"
        };

        saveFileDialog = new SaveFileDialog
        {
            Filter = "Archivo Excel|*.xlsx",
            Title = "Guardar Archivo Combinado",
            FileName = "unificado.xlsx"
        };

        this.Controls.Add(selectFilesButton);
        this.Controls.Add(combineButton);
        this.Controls.Add(statusLabel);
    }

    private void CombineButton_Click(object? sender, EventArgs e)
    {
        if (saveFileDialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(saveFileDialog.FileName)))
                {
                    foreach (var file in excelFiles)
                    {
                        using (var sourcePackage = new ExcelPackage(new FileInfo(file)))
                        {
                            foreach (var sheet in sourcePackage.Workbook.Worksheets)
                            {
                                string sheetName = sheet.Name;
                                int suffix = 1;
                                while (package.Workbook.Worksheets[sheetName] !=  null)
                                {
                                    sheetName = $"{sheetName}_{suffix}";
                                }
                                package.Workbook.Worksheets.Add(sheetName, sheet);
                            }
                        }
                    }
                    package.Save();
                }
                statusLabel.Text = $"Archivos combinados en [{saveFileDialog.FileName}]";
                MessageBox.Show("Hojas unificadas exitosamente!", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception exception)
            {
                statusLabel.Text = "Error al combinar archivos.";
                MessageBox.Show($"Error: {exception.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private void selectFilesButton_Click(object? sender, EventArgs e)
    {
        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            excelFiles = openFileDialog.FileNames;
            statusLabel.Text = $"Seleccionados: {excelFiles.Length} archivo(s).";
            combineButton.Enabled = excelFiles.Length > 0;
        }
    }
    
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }
}