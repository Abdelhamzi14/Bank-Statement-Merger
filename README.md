# Bank Statement Merger

This application is a **100% client-side** tool built with React, Vite, and Tailwind CSS. It allows you to merge bank statement CSVs, auto-categorize transactions, and export consolidated reports without ever sending your sensitive financial data to a server.

## Data Privacy & Security

- **100% Local**: All processing, parsing, and categorization happens entirely in your browser.
- **Zero Server Storage**: No data is uploaded to, stored on, or processed by any remote server.
- **Temporary State**: Data is held in volatile memory (React State) and is completely wiped upon page refresh or when the "Reset" button is pressed.
- **Secure by Design**: Designed for privacy-conscious users who need bank reconciliation without third-party exposure.

## Features

- **CSV/Excel Upload**: Local ingestion of bank statements with flexible column mapping.
- **Auto-Categorizing**: Rule-based logic for Merchant Fees, Chargebacks, and more.
- **Bulk Editing**: Batch categorize transactions efficiently.
- **Excel Export**: Generate and download professional reports locally.

## Tech Stack

- **Framework**: React 19 + Vite
- **Styling**: Tailwind CSS 4
- **Animations**: Motion (formerly Framer Motion)
- **Icons**: Lucide React
- **Data Handling**: PapaParse (CSV), ExcelJS (XLSX)
