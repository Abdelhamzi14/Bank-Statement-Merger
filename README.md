# Bank Statement Merger

This application is built with React, Vite, and Tailwind CSS. It allows you to merge bank statement CSVs, auto-categorize transactions, and export consolidated reports.

## Vercel Deployment

To deploy this application on Vercel:

1. **Push to GitHub**: Push your code to a GitHub repository.
2. **Import to Vercel**: Connect your GitHub repository to Vercel.
3. **Configure Environment Variables**:
   - In the Vercel dashboard, go to your project's **Settings** > **Environment Variables**.
   - Add a key named `GEMINI_API_KEY` with your Google Gemini API key as the value.
4. **Deploy**: Vercel will automatically detect the Vite configuration and deploy the app.

## Features

- **CSV Upload**: Support for multiple bank statements with flexible column mapping.
- **Auto-Categorizing**: Powered by categorized logic to separate bank transfers, fees, and chargebacks.
- **Bulk Editing**: Multi-select transactions to categorize them in batches.
- **Journal Entry Review**: Refine NetSuite account numbers before export.
- **Excel Export**: Download professional reports using `exceljs`.

## Tech Stack

- **Framework**: React 19 + Vite
- **Styling**: Tailwind CSS 4
- **Animations**: Motion (formerly Framer Motion)
- **Icons**: Lucide React
- **Data Handling**: PapaParse (CSV), ExcelJS (XLSX)
