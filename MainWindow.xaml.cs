using Microsoft.Win32;
using System;
using System.Threading.Tasks;
using System.Windows;
using InvoiceOCR_MultiFormat.Services;
using InvoiceOCR_MultiFormat.Extractors;

namespace InvoiceOCR_MultiFormat
{
    public partial class MainWindow : Window
    {
        private readonly OcrSpaceService _ocrService;
        private readonly UniversalInvoiceExtractor _extractor;

        public MainWindow()
        {
            InitializeComponent();
            _ocrService = new OcrSpaceService();
            _extractor = new UniversalInvoiceExtractor();
        }

        private async void Upload_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Filter = "Invoice Files|*.pdf;*.png;*.jpg;*.jpeg;*.bmp;*.tiff|PDF Files|*.pdf|Image Files|*.png;*.jpg;*.jpeg;*.bmp;*.tiff",
                Title = "Select Invoice (PDF or Image)"
            };

            if (fileDialog.ShowDialog() != true)
                return;

            await ProcessInvoiceAsync(fileDialog.FileName);
        }

        private async Task ProcessInvoiceAsync(string filePath)
        {
            try
            {
                UploadButton.IsEnabled = false;
                ShowLoadingState();
                StatusText.Text = "⏳ Processing invoice with OCR.space API...";

                string text = await _ocrService.ExtractTextAsync(filePath);

                if (string.IsNullOrWhiteSpace(text))
                {
                    MessageBox.Show("No text could be extracted from the file.", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    ResetAllFields();
                    StatusText.Text = "❌ Extraction failed";
                    UploadButton.IsEnabled = true;
                    return;
                }

                string preview = text.Length > 2000 ? text.Substring(0, 2000) + "\n\n... (truncated)" : text;
                MessageBox.Show($"Raw OCR Output ({text.Length} characters):\n\n{preview}",
                    "OCR Debug Data",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                CompanyNameBox.Text = _extractor.ExtractCompanyName(text);
                InvoiceNoBox.Text = _extractor.ExtractInvoiceNumber(text);
                DateBox.Text = _extractor.ExtractDate(text);
                TRNBox.Text = _extractor.ExtractTRN(text);
                SalesPersonBox.Text = _extractor.ExtractSalesPerson(text);
                PaymentTermsBox.Text = _extractor.ExtractPaymentTerms(text);
                ShipDateBox.Text = _extractor.ExtractShipDate(text);
                DONumberBox.Text = _extractor.ExtractDONumber(text);
                SONumberBox.Text = _extractor.ExtractSONumber(text);

                var items = _extractor.ExtractLineItems(text);
                InvoiceItemsGrid.ItemsSource = items;

                StatusText.Text = $"✅ Complete! Extracted {items.Count} line item(s)";
                UploadButton.IsEnabled = true;

                MessageBox.Show($"✅ Successfully extracted {items.Count} line item(s)!",
                    "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}\n\nStack: {ex.StackTrace}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                StatusText.Text = "❌ Error occurred";
                ResetAllFields();
                UploadButton.IsEnabled = true;
            }
        }

        private void ShowLoadingState()
        {
            CompanyNameBox.Text = "⏳ Processing...";
            InvoiceNoBox.Text = "⏳ Processing...";
            DateBox.Text = "⏳ Processing...";
            TRNBox.Text = "⏳ Processing...";
            SalesPersonBox.Text = "⏳ Processing...";
            PaymentTermsBox.Text = "⏳ Processing...";
            ShipDateBox.Text = "⏳ Processing...";
            DONumberBox.Text = "⏳ Processing...";
            SONumberBox.Text = "⏳ Processing...";
            InvoiceItemsGrid.ItemsSource = null;
        }

        private void ResetAllFields()
        {
            CompanyNameBox.Text = "N/A";
            InvoiceNoBox.Text = "N/A";
            DateBox.Text = "N/A";
            TRNBox.Text = "N/A";
            SalesPersonBox.Text = "N/A";
            PaymentTermsBox.Text = "N/A";
            ShipDateBox.Text = "N/A";
            DONumberBox.Text = "N/A";
            SONumberBox.Text = "N/A";
            InvoiceItemsGrid.ItemsSource = null;
        }
    }
}