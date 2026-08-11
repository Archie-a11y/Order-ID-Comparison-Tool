# 📊 Order ID Comparison Tool

A high-performance, memory-efficient PHP web utility designed to seamlessly compare master order lists against settlement or income sheets to identify discrepancies, unmatched transactions, and missing records.

This tool is optimized for low-resource environments (such as Apache/cPanel hosting) and utilizes zero-memory-leak streaming algorithms to safely process large datasets without causing server execution timeouts (500/503 errors).

---

## ✨ Features

- ⚡ **Optimized Memory Footprint**: Leverages **OpenSpout v4** for stream-reading and stream-writing to complete comparison tasks with minimal memory usage, preventing traditional PHP memory exhaustion on large files.
- 🌐 **Native Multilingual UI**: Localized translation dictionaries for English, Chinese (简体中文), and Bahasa Melayu dynamically saved via persistent client cookies.
- 🌓 **Dynamic Theme Engine**: Smooth transition between Light and Dark visual modes using Bootstrap 5.3 components and styling.
- 🔍 **Dynamic Worksheet Extraction**: Powered by client-side **SheetJS**, the interface instantly reads and populates the available Excel sheets/tabs as soon as a file is selected.
- 🔒 **Zero-Disk Clutter & Privacy**: Temporary Excel outputs are automatically purged upon successful user download, or automatically cleaned up after 1 hour if left unclaimed.
- 📊 **Interactive Results Dashboard**: Displays visual statistics with a dual-colored stacked progress bar showing the matched/missing distribution.

---

## 🏗️ Tech Stack

- **Backend**: PHP 8.2 / 8.3 / 8.4
- **Streaming Core**: OpenSpout v4 (Excel & CSV high-speed processor)
- **Frontend CSS Framework**: Bootstrap v5.3.3 & Bootstrap Icons
- **Client-Side Document Parser**: SheetJS (xlsx.full.min.js)
- **Dependency Manager**: Composer

---

## 🖼️ Project Screenshots & User Guide

### 📍 Step 1: Interface Preparation & Localization
Select your preferred user interface language (English, Chinese, or Malay) and toggle the Light/Dark mode contrast settings to match your working environment.
<img width="1366" height="605" alt="image" src="https://github.com/user-attachments/assets/e2c42dcc-4e8e-4bc2-8955-f5a0001070c8" />


### 📍 Step 2: Upload Files & Select Target Worksheets
Load your master order file (Orders File) and the income workbook. Once the income file is uploaded, the sheet names will populate in the dropdown menu.
<img width="437" height="107" alt="image" src="https://github.com/user-attachments/assets/8eb4aebd-5017-40a9-83d4-bdf21bb00254" />


### 📍 Step 3: Match Configuration
Specify the single-letter or multi-letter columns containing the unique Order IDs and Status identifiers on both spreadsheets.
<img width="1366" height="603" alt="image" src="https://github.com/user-attachments/assets/c48fc141-bd2d-4051-a929-48c72694eebf" />


### 📍 Step 4: Analyze Summary Dashboard & Download Results
View the live performance statistics displaying total orders processed, match percentage, and total discrepancies before generating the final file.
<img width="1366" height="605" alt="image" src="https://github.com/user-attachments/assets/507c99ab-7f4c-4a89-aa3f-82901ff40019" />


---

## 🛠️ Project Setup

### Prerequisites
- **PHP**: Version 8.2 or higher
- **Composer**: Dependency manager installed on your machine or server environment
- **PHP Extensions**: Ensure `ext-zip`, `ext-xmlreader`, and `ext-dom` are enabled in your `php.ini` setup

### Installation Steps

1. **Clone the Repository**
   ```bash
   git clone https://github.com/Archie-a11y/Order-ID-Comparison-Tool.git
   cd Order-ID-Comparison-Tool
   ```

2. **Install Vendor Dependencies**
   Run the following composer sequence to securely fetch OpenSpout and initialize autoload features:
   ```bash
   composer install --no-dev --optimize-autoloader
   ```

3. **Verify Folder Permissions**
   Ensure that the target directory has permission to create and delete directories for local caching:
   ```bash
   chmod -R 755 uploads/
   ```

4. **Launch Local Server**
   Start a local PHP test environment to run the script:
   ```bash
   php -S localhost:8000
   ```
   Open `http://localhost:8000` in your web browser.
