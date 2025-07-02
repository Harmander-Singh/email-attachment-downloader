# 📥 Automated Email Attachment Downloader (WPF + Outlook Interop)

A professional Windows desktop application that connects to Microsoft Outlook and downloads email attachments automatically — filtered by sender, date, or file type — and saves them in neatly organized folders with real-time progress updates.

---

## 🔑 Key Features

### 🎨 Professional UI Design
- Modern WPF interface with rounded corners and soft styling
- Real-time progress tracking and status messages
- Clear layout with sections for input, progress, and results

### 📧 Smart Email Processing
- Filter emails by sender (email or display name)
- Optional date range filtering
- File type filter: download only `.pdf`, `.docx`, `.xlsx`, etc.
- Processes Inbox or custom folders

### 📁 Intelligent File Organization
- Automatically creates folders like `2025-07-01_Invoice Amazon`
- Handles duplicate filenames by appending numbers
- Cleans filenames by removing invalid characters
- Displays file sizes in a human-readable format (e.g., `2.3 MB`)

### ⚡ Advanced Functionality
- Real-time progress and log display
- Cancel operation mid-download
- Error handling for COM and Outlook-specific issues
- Proper cleanup of Outlook COM objects for stability

---

## 🛠️ Prerequisites

- ✅ Microsoft Outlook (desktop) installed and configured
- ✅ .NET 6.0 SDK or later
- ✅ Visual Studio 2022
- ✅ Windows 10/11

---

## 🚀 How It Works

1. Launch the app  
2. Enter the **sender's email** or name  
3. Choose a **download location**  
4. Optionally select a **date range** and **file types**  
5. Click **Start Download**  
6. Watch real-time progress and logs  
7. Files will be saved in folders like:
   📁 D:\EmailAttachments\2025-07-01_Invoice from Amazon
   📄 invoice_123.pdf (321 KB)

---

## 📸 Preview

![Preview](https://raw.githubusercontent.com/harmander-singh/email-attachment-downloader/main/preview.png)

---

## 🔧 Setup Instructions

1. Clone the repository or download the ZIP  
2. Open the solution in Visual Studio 2022  
3. Restore NuGet packages  
   ```bash
   Install-Package Microsoft.Office.Interop.Outlook
4. Build and run the WPF application

---

### 📩 Want Custom Features?

This application can be customized to meet your specific needs.

Available customizations:
- ✅ Gmail or IMAP support  
- ✅ Office 365 OAuth integration  
- ✅ Cloud sync (OneDrive, Dropbox, Google Drive)  
- ✅ Branded installer (.exe)  
- ✅ Logging to a database or via email notifications  

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).  
You are free to use, modify, and distribute it for personal or commercial use.

---

## 🙌 Author

Made with 💙 by [@harmander-singh](https://github.com/harmander-singh)  
Contributions, stars, and feedback are welcome!


