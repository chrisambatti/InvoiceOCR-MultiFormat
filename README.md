# 📄 InvoiceOCR-MultiFormat

A powerful **multi-format invoice data extraction application** that uses Optical Character Recognition (OCR) to automatically extract key information from invoices in PDF and image formats.

Designed for automation, accounting workflows, and document digitization, this tool converts unstructured invoice data into structured, usable text.

---

## 🚀 Features

- 📄 Supports multiple input formats  
  - PDF files  
  - Scanned invoices  
  - Images (JPG, PNG, etc.)

- 🔍 Accurate OCR extraction using Tesseract

- 🧾 Extracts key invoice details such as:
  - Company / Vendor Name  
  - Invoice Number  
  - Invoice Date  
  - TRN / Tax ID (if available)  
  - Total Amount  
  - Line item details (if detected)

- 🖥️ Clean desktop interface (WPF)

- ⚡ Fast processing with local execution (no cloud required)

- 🔐 Privacy-friendly — data stays on your machine

---

## 🧠 How It Works

1. Upload an invoice (PDF or image)
2. PDF files are converted into images (if required)
3. OCR engine processes the document
4. Text is parsed using pattern matching
5. Structured invoice data is displayed

---

## 🛠️ Technology Stack

- 💻 C# (.NET / WPF)
- 🔎 Tesseract OCR Engine
- 📑 PDF processing libraries
- 🧩 Regex-based text parsing

---



---

## 📦 Installation
2️⃣ Install prerequisites

.NET SDK (recommended latest version)

Tesseract OCR installed on your system

## 👉 Download Tesseract:
https://github.com/tesseract-ocr/tesseract

Make sure the executable path is configured correctly.
### 1️⃣ Clone the repository
```bash
git clone https://github.com/chrisambatti/InvoiceOCR-MultiFormat.git
cd InvoiceOCR-MultiFormat
```
## ▶️ Usage

Run the application

Click Upload

Select an invoice file (PDF/Image)

View extracted data instantly

## 📂 Supported File Types

PDF (.pdf)

JPEG (.jpg / .jpeg)

PNG (.png)

Scanned documents

## 📊 Example Use Cases

Accounting automation

Expense tracking systems

Accounts payable workflows

Data digitization projects

OCR research and experimentation

## ⚠️ Limitations

Accuracy depends on image quality

Handwritten invoices may not be recognized

Complex layouts may require additional parsing logic

Very low-resolution scans may produce incorrect results

## 🔮 Future Enhancements

Batch processing of multiple invoices

Export to Excel / CSV

Database integration

AI-assisted field detection

Support for more document types

## 🤝 Contributing

Contributions are welcome!

Fork the repository

Create a feature branch

Commit your changes

Submit a pull request
