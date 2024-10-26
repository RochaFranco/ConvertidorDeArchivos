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
using Paragraph = iTextSharp.text.Paragraph;
using Microsoft.Win32;
using iTextSharp.text.pdf.parser;
using System.Drawing.Imaging;
using System.Drawing;
using OfficeOpenXml.Drawing;



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

        public void ConvertPdfWithImagesToExcel(string pdfFilePath, string excelFilePath)
        {

            // Establecer el contexto de la licencia antes de usar EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Inicializar EPPlus y crear el archivo Excel
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Datos Extraídos");

                // Leer el archivo PDF usando iTextSharp
                using (PdfReader reader = new PdfReader(pdfFilePath))
                {
                    int totalRows = 1; // Empezar en la primera fila de Excel

                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        // Extraer el texto de cada página
                        string pageText = PdfTextExtractor.GetTextFromPage(reader, i);

                        // Dividir el texto en líneas
                        var lines = pageText.Split('\n');

                        foreach (var line in lines)
                        {
                            // Dividir cada línea en palabras o columnas usando un delimitador
                            var columns = line.Split(' '); // Cambia el delimitador según sea necesario

                            // Añadir el texto a las celdas de Excel
                            for (int col = 0; col < columns.Length; col++)
                            {
                                worksheet.Cells[totalRows, col + 1].Value = columns[col].Trim();
                            }

                            totalRows++; // Mover a la siguiente fila en Excel
                        }

                        // Extraer las imágenes de la página actual y añadirlas al Excel
                        var images = GetImagesFromPdfPage(reader, i);

                        foreach (var image in images)
                        {
                            using (var memoryStream = new MemoryStream())
                            {
                                // Guardar la imagen en el MemoryStream como PNG
                                image.Save(memoryStream, ImageFormat.Png);
                                memoryStream.Position = 0; // Reiniciar la posición del stream

                                // Añadir la imagen al archivo Excel desde el MemoryStream
                                var excelImage = worksheet.Drawings.AddPicture($"Image_{i}_{totalRows}", memoryStream);

                                // Colocar la imagen en la celda correspondiente
                                excelImage.SetPosition(totalRows - 1, 0, 0, 0); // Ajusta según sea necesario
                            }

                            totalRows++; // Aumentar fila después de insertar la imagen
                        }
                    }
                }

                // Guardar el archivo Excel
                FileInfo excelFile = new FileInfo(excelFilePath);
                package.SaveAs(excelFile);
            }
        }

        // Método auxiliar para extraer imágenes de una página de PDF
        private List<System.Drawing.Image> GetImagesFromPdfPage(PdfReader pdfReader, int pageNumber)
        {
            List<System.Drawing.Image> images = new List<System.Drawing.Image>();

            var pdfResources = pdfReader.GetPageN(pageNumber).GetAsDict(PdfName.RESOURCES);
            var xObject = pdfResources.GetAsDict(PdfName.XOBJECT);

            if (xObject != null)
            {
                foreach (var name in xObject.Keys)
                {
                    var obj = xObject.GetAsIndirectObject(name);
                    var pdfObj = PdfReader.GetPdfObject(obj);
                    if (pdfObj is PdfStream stream && PdfName.IMAGE.Equals(stream.GetAsName(PdfName.SUBTYPE)))
                    {
                        var pdfImage = new PdfImageObject((PRStream)stream);
                        var image = pdfImage.GetDrawingImage();
                        if (image != null)
                        {
                            images.Add(image);
                        }
                    }
                }
            }

            return images;
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
            else if (Extension_Preview.Text == ".pdf" && Opciones.SelectedIndex == 2)
            {
                ConvertPdfWithImagesToExcel(Text_Preview.Text, @"D:\YoutubeVideos\ArchivoConvertido.xlsx");
            }
        }
    }
}