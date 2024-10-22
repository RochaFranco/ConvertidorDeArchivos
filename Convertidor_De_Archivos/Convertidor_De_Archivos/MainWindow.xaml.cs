using OfficeOpenXml;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Convertidor_De_Archivos;
using Convertidor_De_Archivos.Clases;
using OfficeOpenXml; // EPPlus
using iTextSharp.text; // iTextSharp
using iTextSharp.text.pdf; // iTextSharp
using System.IO;
using Paragraph = iTextSharp.text.Paragraph;
using Microsoft.Win32;

namespace Convertidor_De_Archivos
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        public void ConvertExcelToPdf(string excelFilePath, string pdfFilePath)
        {
            // Establecer el contexto de la licencia antes de usar EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Leer el archivo Excel usando EPPlus
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets[0]; // Usamos la primera hoja de cálculo

                // Crear un documento PDF con iTextSharp
                using (FileStream fs = new FileStream(pdfFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    Document pdfDoc = new Document(PageSize.A4);
                    PdfWriter writer = PdfWriter.GetInstance(pdfDoc, fs);
                    pdfDoc.Open();

                    // Crear una tabla en el PDF con el número correcto de columnas
                    PdfPTable table = new PdfPTable(worksheet.Dimension.End.Column);
                    table.WidthPercentage = 100;

                    // Iterar sobre cada fila y columna en la hoja de Excel
                    for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
                    {
                        for (int j = 1; j <= worksheet.Dimension.End.Column; j++)
                        {
                            var cell = worksheet.Cells[i, j];
                            string cellValue = cell.Text;

                            // Crear una celda en el PDF
                            PdfPCell pdfCell = new PdfPCell(new Phrase(cellValue));

                            // Aplicar color de fondo si está definido en Excel
                            if (cell.Style.Fill.PatternType != OfficeOpenXml.Style.ExcelFillStyle.None)
                            {
                                var bgColor = cell.Style.Fill.BackgroundColor.Rgb;
                                if (!string.IsNullOrEmpty(bgColor))
                                {
                                    var color = new BaseColor(
                                        int.Parse(bgColor.Substring(0, 2), System.Globalization.NumberStyles.HexNumber),
                                        int.Parse(bgColor.Substring(2, 2), System.Globalization.NumberStyles.HexNumber),
                                        int.Parse(bgColor.Substring(4, 2), System.Globalization.NumberStyles.HexNumber)
                                    );
                                    pdfCell.BackgroundColor = color;
                                }
                            }

                            // Aplicar bordes a la celda
                            pdfCell.BorderWidth = 1;
                            table.AddCell(pdfCell);
                        }
                    }

                    // Añadir la tabla al documento PDF
                    pdfDoc.Add(table);
                    pdfDoc.Close();
                }
            }
        }


        private void Button_Subir(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            bool? succes = fileDialog.ShowDialog();
            if (succes == true)
            {
                string path = fileDialog.FileName;

                string fileExtension = System.IO.Path.GetExtension(path).ToLower();

                Text_Preview.Text = path;
                Extension_Preview.Text = fileExtension;


            }
            else
            {

            }
            
        }

        private void Button_Convertir(object sender, RoutedEventArgs e)
        {
            if (Extension_Preview.Text == ".xlsx" && Opciones.SelectedIndex == 0)
            {
                ConvertExcelToPdf(Text_Preview.Text, @"D:\YoutubeVideos\ArchivoConvertido.pdf");
            }
        }
    }
}