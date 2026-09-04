import { useState, useMemo } from 'react';
import { type BankTransaction, type Category, type StatementFile } from '../types';
import { buildExternalIdMap, getExternalId } from '../lib/externalId';
import { 
  Download, 
  Table, 
  FileText, 
  Layers, 
  Grid, 
  CheckCircle2, 
  ChevronDown, 
  ChevronUp, 
  ArrowDownRight, 
  ArrowUpRight, 
  DollarSign,
  PieChart
} from 'lucide-react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { motion, AnimatePresence } from 'motion/react';
import { cn } from '../lib/utils';

interface Props {
  files: StatementFile[];
  transactions: BankTransaction[];
  externalIdStartSequence?: number;
}

const CATEGORIES: Category[] = [
  'Chargebacks',
  'Merchant Fees',
  'Settlement Deposit',
  'Bank Transfer',
  'Miscellaneous'
];

const CATEGORY_STYLES: Record<Category, string> = {
  'Chargebacks': 'border-l-red-500',
  'Merchant Fees': 'border-l-orange-400',
  'Settlement Deposit': 'border-l-emerald-500',
  'Bank Transfer': 'border-l-blue-400',
  'Miscellaneous': 'border-l-slate-400'
};

const CATEGORY_BADGE_BG: Record<Category, string> = {
  'Chargebacks': 'bg-red-50 text-red-700 border-red-200',
  'Merchant Fees': 'bg-orange-50 text-orange-700 border-orange-200',
  'Settlement Deposit': 'bg-emerald-50 text-emerald-700 border-emerald-200',
  'Bank Transfer': 'bg-blue-50 text-blue-700 border-blue-200',
  'Miscellaneous': 'bg-slate-100 text-slate-700 border-slate-200'
};

const CATEGORY_STYLE_TEXT: Record<Category, string> = {
  'Chargebacks': 'text-red-700',
  'Merchant Fees': 'text-orange-600',
  'Settlement Deposit': 'text-emerald-700',
  'Bank Transfer': 'text-blue-600',
  'Miscellaneous': 'text-slate-600'
};

const getDebit = (t: BankTransaction) => {
  if (t.debit != null) return t.debit;
  return t.amount < 0 ? Math.abs(t.amount) : 0;
};

const getCredit = (t: BankTransaction) => {
  if (t.credit != null) return t.credit;
  return t.amount > 0 ? t.amount : 0;
};

const formatCurrency = (val: number) => {
  return val.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};

export default function SummaryAndExport({ files, transactions, externalIdStartSequence = 1 }: Props) {
  const [activeTab, setActiveTab] = useState<'all' | 'grouped' | 'matrix' | 'separate'>('all');
  const [expandedFiles, setExpandedFiles] = useState<Record<string, boolean>>({});

  // 1. Compute overall Category Consolidation Stats (Separate)
  const categoryStats = useMemo(() => {
    return CATEGORIES.map(category => {
      const items = transactions.filter(t => t.category === category);
      const total = items.reduce((sum, item) => sum + item.amount, 0);
      const debit = items.reduce((sum, item) => sum + getDebit(item), 0);
      const credit = items.reduce((sum, item) => sum + getCredit(item), 0);
      return { 
        name: category, 
        count: items.length, 
        total, 
        debit, 
        credit,
        percentOfCount: transactions.length > 0 ? (items.length / transactions.length) * 100 : 0
      };
    });
  }, [transactions]);

  // 2. Derive unique file names
  const uniqueFileNames = useMemo(() => {
    const names = new Set<string>();
    files.forEach(f => {
      if (f.file?.name) names.add(f.file.name);
    });
    transactions.forEach(t => {
      if (t.sourceFile) names.add(t.sourceFile);
    });
    return Array.from(names);
  }, [files, transactions]);

  // 3. Compute detailed File Breakdown Stats (Grouped by File & Categories)
  const fileStats = useMemo(() => {
    return uniqueFileNames.map((fileName, idx) => {
      const fileTxns = transactions.filter(t => t.sourceFile === fileName);
      const total = fileTxns.reduce((sum, item) => sum + item.amount, 0);
      const debit = fileTxns.reduce((sum, item) => sum + getDebit(item), 0);
      const credit = fileTxns.reduce((sum, item) => sum + getCredit(item), 0);

      // Breakdown by category within this specific file
      const categoriesBreakdown = CATEGORIES.map(cat => {
        const catTxns = fileTxns.filter(t => t.category === cat);
        const catTotal = catTxns.reduce((sum, item) => sum + item.amount, 0);
        const catDebit = catTxns.reduce((sum, item) => sum + getDebit(item), 0);
        const catCredit = catTxns.reduce((sum, item) => sum + getCredit(item), 0);
        return {
          category: cat,
          count: catTxns.length,
          total: catTotal,
          debit: catDebit,
          credit: catCredit,
          percentOfFileCount: fileTxns.length > 0 ? (catTxns.length / fileTxns.length) * 100 : 0,
        };
      });

      return {
        id: `file-${idx}`,
        name: fileName,
        count: fileTxns.length,
        total,
        debit,
        credit,
        percentOfTotalCount: transactions.length > 0 ? (fileTxns.length / transactions.length) * 100 : 0,
        categories: categoriesBreakdown
      };
    });
  }, [uniqueFileNames, transactions]);

  // Grand totals
  const totalVolume = useMemo(() => transactions.reduce((s, t) => s + t.amount, 0), [transactions]);
  const totalDebits = useMemo(() => transactions.reduce((s, t) => s + getDebit(t), 0), [transactions]);
  const totalCredits = useMemo(() => transactions.reduce((s, t) => s + getCredit(t), 0), [transactions]);

  const toggleFileExpand = (fileName: string) => {
    setExpandedFiles(prev => ({
      ...prev,
      [fileName]: prev[fileName] === undefined ? false : !prev[fileName]
    }));
  };

  const isFileExpanded = (fileName: string) => {
    // Default to open
    return expandedFiles[fileName] !== false;
  };

  const exportToExcel = async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'BankStatementMerger';
      workbook.lastModifiedBy = 'BankStatementMerger';
      workbook.created = new Date();

      const headerFill: ExcelJS.Fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1E3A8A' }
      };

      const subHeaderFill: ExcelJS.Fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF334155' } // slate-700
      };

      const fileGroupFill: ExcelJS.Fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE2E8F0' } // slate-200
      };

      // ==========================================
      // 1. DASHBOARD SUMMARY WORKSHEET
      // ==========================================
      const wsSummary = workbook.addWorksheet('Dashboard Summary');

      // Professional Color Palette
      const COLOR_GREEN = 'FF16A34A'; // Credits / Inflows / Positive Net
      const COLOR_RED = 'FFDC2626';   // Debits / Outflows / Negative Net
      const COLOR_BLACK = 'FF0F172A'; // Zeros / Neutral / Labels
      const CURRENCY_FMT = '$#,##0.00;($#,##0.00);"$0.00"';

      // Helpers for amount formatting
      const formatNetCell = (cell: ExcelJS.Cell, val: number, isBold: boolean = false) => {
        cell.numFmt = CURRENCY_FMT;
        let color = COLOR_BLACK;
        if (val > 0.0001) color = COLOR_GREEN;
        else if (val < -0.0001) color = COLOR_RED;
        cell.font = { bold: isBold, color: { argb: color }, size: isBold ? 11 : 10 };
      };

      const formatDebitCell = (cell: ExcelJS.Cell, val: number | null, isBold: boolean = false) => {
        cell.numFmt = CURRENCY_FMT;
        const color = (val && val > 0.0001) ? COLOR_RED : COLOR_BLACK;
        cell.font = { bold: isBold, color: { argb: color }, size: isBold ? 11 : 10 };
      };

      const formatCreditCell = (cell: ExcelJS.Cell, val: number | null, isBold: boolean = false) => {
        cell.numFmt = CURRENCY_FMT;
        const color = (val && val > 0.0001) ? COLOR_GREEN : COLOR_BLACK;
        cell.font = { bold: isBold, color: { argb: color }, size: isBold ? 11 : 10 };
      };

      // Table Header Theme (Executive Navy)
      const matrixHeaderFill: ExcelJS.Fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1E3A8A' } // Royal Navy
      };

      // Title Banner
      const titleRow = wsSummary.addRow(['BANK RECONCILIATION CONSOLIDATED EXECUTIVE REPORT']);
      titleRow.height = 30;
      titleRow.getCell(1).font = { bold: true, size: 15, color: { argb: 'FF0F172A' } };

      const subTitleRow = wsSummary.addRow([
        `Generated: ${new Date().toLocaleString()}  |  Source Files: ${uniqueFileNames.length}  |  Processed Records: ${transactions.length.toLocaleString()}  |  Consolidated Net Position: $${formatCurrency(totalVolume)}`
      ]);
      subTitleRow.height = 18;
      subTitleRow.getCell(1).font = { italic: true, size: 10, color: { argb: 'FF475569' } };
      wsSummary.addRow([]); // Spacer

      // ----------------------------------------------------
      // FILE & CATEGORY CROSS-TABULATION MATRIX
      // ----------------------------------------------------
      const sec1Header = wsSummary.addRow(['FILE & CATEGORY CROSS-TABULATION MATRIX']);
      sec1Header.height = 25;
      sec1Header.getCell(1).font = { bold: true, size: 12, color: { argb: 'FF1E3A8A' } };
      sec1Header.getCell(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFEFF6FF' } // soft blue-50 tint
      };

      const matrixHeaders = [
        'Statement File Name',
        'Chargebacks',
        'Merchant Fees',
        'Settlement Deposit',
        'Bank Transfer',
        'Miscellaneous',
        'File Net Total',
        'Total Records'
      ];
      const matrixHeaderRow = wsSummary.addRow(matrixHeaders);
      matrixHeaderRow.height = 26;
      matrixHeaderRow.eachCell((cell, colNumber) => {
        cell.fill = matrixHeaderFill;
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10.5 };
        cell.alignment = { 
          horizontal: colNumber === 1 ? 'left' : colNumber === 8 ? 'center' : 'right', 
          vertical: 'middle' 
        };
      });

      // Data rows for Matrix
      fileStats.forEach((f, idx) => {
        const catMap = new Map(f.categories.map(c => [c.category, c.total]));
        const isEven = idx % 2 === 0;
        const row = wsSummary.addRow([
          f.name,
          catMap.get('Chargebacks') || 0,
          catMap.get('Merchant Fees') || 0,
          catMap.get('Settlement Deposit') || 0,
          catMap.get('Bank Transfer') || 0,
          catMap.get('Miscellaneous') || 0,
          f.total,
          f.count
        ]);
        row.height = 22;

        const rowFill: ExcelJS.Fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: isEven ? 'FFFFFFFF' : 'FFF8FAFC' }
        };

        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          cell.fill = rowFill;
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
          };

          if (colNumber === 1) {
            cell.font = { bold: true, color: { argb: 'FF1E293B' }, size: 10 };
            cell.alignment = { horizontal: 'left', vertical: 'middle' };
          } else if (colNumber >= 2 && colNumber <= 6) {
            const val = Number(cell.value) || 0;
            formatNetCell(cell, val, false);
            cell.alignment = { horizontal: 'right', vertical: 'middle' };
          } else if (colNumber === 7) {
            formatNetCell(cell, f.total, true);
            cell.alignment = { horizontal: 'right', vertical: 'middle' };
          } else if (colNumber === 8) {
            cell.numFmt = '#,##0';
            cell.font = { bold: false, color: { argb: COLOR_BLACK }, size: 10 };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
          }
        });
      });

      // Matrix Grand Consolidated Total Row
      const catTotalsMap = new Map(categoryStats.map(c => [c.name, c.total]));
      const matrixTotalRow = wsSummary.addRow([
        'CONSOLIDATED MATRIX TOTAL',
        catTotalsMap.get('Chargebacks') || 0,
        catTotalsMap.get('Merchant Fees') || 0,
        catTotalsMap.get('Settlement Deposit') || 0,
        catTotalsMap.get('Bank Transfer') || 0,
        catTotalsMap.get('Miscellaneous') || 0,
        totalVolume,
        transactions.length
      ]);
      matrixTotalRow.height = 26;

      const matrixTotalFill: ExcelJS.Fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE2E8F0' } // slate-200 highlight
      };

      matrixTotalRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        cell.fill = matrixTotalFill;
        cell.border = {
          top: { style: 'thin', color: { argb: 'FF64748B' } },
          bottom: { style: 'double', color: { argb: 'FF0F172A' } }
        };

        if (colNumber === 1) {
          cell.font = { bold: true, color: { argb: 'FF0F172A' }, size: 11 };
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
        } else if (colNumber >= 2 && colNumber <= 6) {
          const val = Number(cell.value) || 0;
          formatNetCell(cell, val, true);
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
        } else if (colNumber === 7) {
          formatNetCell(cell, totalVolume, true);
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
        } else if (colNumber === 8) {
          cell.numFmt = '#,##0';
          cell.font = { bold: true, color: { argb: 'FF0F172A' }, size: 11 };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }
      });

      // Adjust column widths on Summary worksheet for optimal readability
      wsSummary.getColumn(1).width = 38; // Statement File Name
      wsSummary.getColumn(2).width = 20; // Chargebacks
      wsSummary.getColumn(3).width = 20; // Merchant Fees
      wsSummary.getColumn(4).width = 22; // Settlement Deposit
      wsSummary.getColumn(5).width = 20; // Bank Transfer
      wsSummary.getColumn(6).width = 20; // Miscellaneous
      wsSummary.getColumn(7).width = 22; // Net Total
      wsSummary.getColumn(8).width = 16; // Total Records

      // ==========================================
      // 2. INDIVIDUAL CATEGORY WORKSHEETS
      // ==========================================
      CATEGORIES.forEach(category => {
        const categoryItems = transactions.filter(t => t.category === category);
        if (categoryItems.length > 0) {
          const sheetName = category.replace(/[\\/*?[\]]/g, '').substring(0, 31);
          const ws = workbook.addWorksheet(sheetName);

          const columns = [
            { header: 'Date', key: 'date', width: 15 },
            { header: 'Description', key: 'description', width: 50 },
            { header: 'Account#', key: 'accountNumber', width: 20 },
            { header: 'Amount', key: 'amount', width: 18 },
            { header: 'NS Debit Account', key: 'nsDebit', width: 25 },
            { header: 'Raw Debit', key: 'debit', width: 18 },
            { header: 'Raw Credit', key: 'credit', width: 18 },
            { header: 'NS Credit Account', key: 'nsCredit', width: 25 },
            { header: 'Adj Debit', key: 'adjDebit', width: 18 },
            { header: 'Adj Credit', key: 'adjCredit', width: 18 }
          ];

          if (category === 'Merchant Fees') {
            columns.push({ header: 'Department', key: 'department', width: 15 });
          }

          columns.push({ header: 'Source File', key: 'sourceFile', width: 30 });

          ws.columns = columns;

          // Header row styling
          const headerRow = ws.getRow(1);
          headerRow.height = 25;
          headerRow.eachCell((cell) => {
            cell.fill = matrixHeaderFill;
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
          });

          // Data rows
          categoryItems.forEach(item => {
            const blankIfZero = (val: number | null) => (val === 0 || val === null) ? null : val;
            const itemDebit = getDebit(item);
            const itemCredit = getCredit(item);

            const rowData: any = {
              date: item.date,
              description: item.description,
              accountNumber: item.accountNumber || 'N/A',
              amount: item.amount,
              nsDebit: item.netsuiteDebitAccount || '',
              debit: blankIfZero(itemDebit),
              credit: blankIfZero(itemCredit),
              nsCredit: item.netsuiteCreditAccount || '',
              adjDebit: blankIfZero(itemCredit),
              adjCredit: blankIfZero(itemDebit),
              sourceFile: item.sourceFile
            };

            if (category === 'Merchant Fees') {
              rowData.department = 'Finance';
            }

            const r = ws.addRow(rowData);
            
            // Apply Red / Green / Black amount styling
            formatNetCell(r.getCell('amount'), item.amount, false);
            formatDebitCell(r.getCell('debit'), itemDebit, false);
            formatCreditCell(r.getCell('credit'), itemCredit, false);
            formatCreditCell(r.getCell('adjDebit'), itemCredit, false);
            formatDebitCell(r.getCell('adjCredit'), itemDebit, false);

            try {
              const debitCell = r.getCell('nsDebit');
              if (debitCell && item.netsuiteDebitAccount) debitCell.font = { bold: true };
              const creditCell = r.getCell('nsCredit');
              if (creditCell && item.netsuiteCreditAccount) creditCell.font = { color: { argb: 'FFFF0000' }, bold: true };
            } catch (e) {
              // Ignore
            }
          });

          // Calculate category totals
          const catTotalAmount = categoryItems.reduce((sum, item) => sum + item.amount, 0);
          const catTotalRawDebit = categoryItems.reduce((sum, item) => sum + getDebit(item), 0);
          const catTotalRawCredit = categoryItems.reduce((sum, item) => sum + getCredit(item), 0);
          const catTotalAdjDebit = catTotalRawCredit;
          const catTotalAdjCredit = catTotalRawDebit;

          // Append Category Total Row
          const totalRowData: any = {
            date: 'TOTAL',
            description: `Total for ${category}`,
            accountNumber: '',
            amount: catTotalAmount,
            nsDebit: '',
            debit: catTotalRawDebit,
            credit: catTotalRawCredit,
            nsCredit: '',
            adjDebit: catTotalAdjDebit,
            adjCredit: catTotalAdjCredit,
            sourceFile: `${categoryItems.length} records`
          };

          if (category === 'Merchant Fees') {
            totalRowData.department = '';
          }

          const totalRow = ws.addRow(totalRowData);
          totalRow.height = 24;

          totalRow.eachCell({ includeEmpty: true }, (cell) => {
            cell.font = { bold: true, size: 11 };
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFF1F5F9' } // slate-100 fill
            };
            cell.border = {
              top: { style: 'thin', color: { argb: 'FF94A3B8' } },
              bottom: { style: 'double', color: { argb: 'FF0F172A' } }
            };
          });

          // Apply Red / Green / Black bold formatting to the totals
          formatNetCell(totalRow.getCell('amount'), catTotalAmount, true);
          formatDebitCell(totalRow.getCell('debit'), catTotalRawDebit, true);
          formatCreditCell(totalRow.getCell('credit'), catTotalRawCredit, true);
          formatCreditCell(totalRow.getCell('adjDebit'), catTotalAdjDebit, true);
          formatDebitCell(totalRow.getCell('adjCredit'), catTotalAdjCredit, true);
          
          ws.views = [{ state: 'frozen', ySplit: 1 }];
        }
      });

      // ==========================================
      // 3. JE UPLOAD WORKSHEET (Yellow Tab Color)
      // ==========================================
      const wsJE = workbook.addWorksheet('JE Upload');
      // Highlight the tab in bright yellow
      wsJE.properties.tabColor = { argb: 'FFFFFF00' };

      // NetSuite JE Columns
      wsJE.columns = [
        { header: 'External ID', key: 'externalId', width: 24 },
        { header: 'Date', key: 'date', width: 14 },
        { header: 'Subsidiary', key: 'subsidiary', width: 14 },
        { header: 'Memo', key: 'memo', width: 38 },
        { header: 'Currency', key: 'currency', width: 16 },
        { header: 'Account', key: 'account', width: 18 },
        { header: 'Debit', key: 'debit', width: 16 },
        { header: 'Credit', key: 'credit', width: 16 },
        { header: 'Line Description', key: 'lineDescription', width: 38 },
        { header: 'Department', key: 'department', width: 16 },
        { header: 'Posting Period', key: 'postingPeriod', width: 16 },
        { header: 'Start Date', key: 'startDate', width: 14 },
        { header: 'End Date', key: 'endDate', width: 14 },
        { header: 'Entity Name', key: 'entityName', width: 20 },
        { header: 'Reversal', key: 'reversal', width: 14 },
        { header: 'Reversal Date', key: 'reversalDate', width: 14 },
        { header: 'Attachment', key: 'attachments', width: 18 }
      ];

      // Set entire tab column defaults to Text format '@'
      wsJE.columns.forEach((col, idx) => {
        if (idx !== 6 && idx !== 7) {
          col.numFmt = '@';
        }
      });

      // Header row styling
      const jeHeaderRow = wsJE.getRow(1);
      jeHeaderRow.height = 26;
      jeHeaderRow.eachCell((cell, colNumber) => {
        if (colNumber === 7) {
          // Debit Header: Crimson Red
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFB91C1C' }
          };
          cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10.5 };
        } else if (colNumber === 8) {
          // Credit Header: Emerald Green
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF15803D' }
          };
          cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10.5 };
        } else if (colNumber === 17) {
          // Column Q "Attachment" Header: Highlighted in Red
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFDC2626' }
          };
          cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10.5 };
        } else {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF1E293B' } // Executive Slate Header
          };
          cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10.5 };
        }
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });

      // Categories to include in JE Upload: Chargebacks, Merchant Fees, Bank Transfer
      const JE_CATEGORIES: Category[] = ['Chargebacks', 'Merchant Fees', 'Bank Transfer'];

      // Filter transactions from the specified category tabs
      const jeItems = transactions.filter(t => JE_CATEGORIES.includes(t.category));

      // Sort consistently: category order, then date
      jeItems.sort((a, b) => {
        const catDiff = JE_CATEGORIES.indexOf(a.category) - JE_CATEGORIES.indexOf(b.category);
        if (catDiff !== 0) return catDiff;
        return (a.date || '').localeCompare(b.date || '');
      });

      // External ID numbering: similar category sharing the same date (e.g. Chargebacks 0001...)
      const extIdMap = buildExternalIdMap(transactions, externalIdStartSequence);

      // Helper to format date as MMMM-YYYY (e.g. "January-2024")
      const getMMMMYYYYFromDate = (dateVal: any): string => {
        if (!dateVal || dateVal === 'N/A') return '';
        const MONTH_NAMES = [
          'January', 'February', 'March', 'April', 'May', 'June',
          'July', 'August', 'September', 'October', 'November', 'December'
        ];

        // If number or numeric string (Excel serial date)
        if (typeof dateVal === 'number' || (!isNaN(Number(dateVal)) && !String(dateVal).includes('-') && !String(dateVal).includes('/'))) {
          const num = Number(dateVal);
          if (num > 20000 && num < 70000) {
            const d = new Date(Math.round((num - 25569) * 86400 * 1000));
            if (!isNaN(d.getTime())) {
              return `${MONTH_NAMES[d.getUTCMonth()]}-${d.getUTCFullYear()}`;
            }
          }
        }

        const str = String(dateVal).trim();

        // Pattern: MM/DD/YYYY or M/D/YYYY or MM-DD-YYYY or M-D-YYYY
        const mdyMatch = str.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})/);
        if (mdyMatch) {
          const m = parseInt(mdyMatch[1], 10);
          let y = parseInt(mdyMatch[3], 10);
          if (y < 100) y += 2000;
          if (m >= 1 && m <= 12) {
            return `${MONTH_NAMES[m - 1]}-${y}`;
          }
        }

        // Pattern: YYYY-MM-DD or YYYY/MM/DD
        const ymdMatch = str.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);
        if (ymdMatch) {
          const y = parseInt(ymdMatch[1], 10);
          const m = parseInt(ymdMatch[2], 10);
          if (m >= 1 && m <= 12) {
            return `${MONTH_NAMES[m - 1]}-${y}`;
          }
        }

        // Fallback standard Date parsing
        const parsed = new Date(str);
        if (!isNaN(parsed.getTime())) {
          return `${MONTH_NAMES[parsed.getMonth()]}-${parsed.getFullYear()}`;
        }

        return '';
      };

      // Populate JE Upload rows
      const blankIfZero = (val: number | null) => (val === 0 || val === null) ? null : val;

      jeItems.forEach(item => {
        const extId = getExternalId(item, extIdMap, externalIdStartSequence);
        const rawDebit = getDebit(item);
        const rawCredit = getCredit(item);
        const adjDebit = rawCredit;
        const adjCredit = rawDebit;

        const nsDebitAccount = item.netsuiteDebitAccount || '';
        const nsCreditAccount = item.netsuiteCreditAccount || '';

        // Derive Memo with MMMM-YYYY prefix based on the Date column
        const monthYear = getMMMMYYYYFromDate(item.date);
        const origDesc = String(item.description || '').trim();
        const memoVal = monthYear
          ? (origDesc.startsWith(monthYear) ? origDesc : (origDesc ? `${monthYear} ${origDesc}` : monthYear))
          : origDesc;

        // Step 1: NS Debit Account with Raw Debit / Raw Credit
        const row1Data = {
          externalId: String(extId),
          date: String(item.date || ''),
          subsidiary: '3',
          memo: memoVal,
          currency: 'US Dollar',
          account: String(nsDebitAccount),
          debit: blankIfZero(rawDebit),
          credit: blankIfZero(rawCredit),
          lineDescription: String(item.description || ''),
          department: nsDebitAccount === '51069' ? 'Finance' : '',
          postingPeriod: '',
          startDate: '',
          endDate: '',
          entityName: '',
          reversal: '',
          reversalDate: '',
          attachments: ''
        };

        const r1 = wsJE.addRow(row1Data);
        r1.height = 20;

        // Ensure text formatting across all standard text columns
        r1.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          if (colNumber !== 7 && colNumber !== 8) {
            cell.numFmt = '@';
          }
        });

        // Highlight Column Q "Attachment" in Red
        const r1Attach = r1.getCell('attachments');
        r1Attach.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFEE2E2' } // Soft Red Highlight
        };
        r1Attach.border = {
          top: { style: 'thin', color: { argb: 'FFFCA5A5' } },
          bottom: { style: 'thin', color: { argb: 'FFFCA5A5' } },
          left: { style: 'thin', color: { argb: 'FFFCA5A5' } },
          right: { style: 'thin', color: { argb: 'FFFCA5A5' } }
        };
        r1Attach.numFmt = '@';

        // Highlight Debit and Credit based on Category Tab details (Raw Debit: Red, Raw Credit: Green)
        const r1Debit = r1.getCell('debit');
        const r1Credit = r1.getCell('credit');
        if (rawDebit && rawDebit > 0.0001) {
          r1Debit.numFmt = '#,##0.00';
          r1Debit.font = { bold: true, color: { argb: COLOR_RED } };
          r1Debit.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
          r1Debit.border = { top: { style: 'thin', color: { argb: 'FFFCA5A5' } }, bottom: { style: 'thin', color: { argb: 'FFFCA5A5' } } };
          r1Debit.alignment = { horizontal: 'right', vertical: 'middle' };
        } else {
          r1Debit.numFmt = '@';
        }

        if (rawCredit && rawCredit > 0.0001) {
          r1Credit.numFmt = '#,##0.00';
          r1Credit.font = { bold: true, color: { argb: COLOR_GREEN } };
          r1Credit.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
          r1Credit.border = { top: { style: 'thin', color: { argb: 'FF86EFAC' } }, bottom: { style: 'thin', color: { argb: 'FF86EFAC' } } };
          r1Credit.alignment = { horizontal: 'right', vertical: 'middle' };
        } else {
          r1Credit.numFmt = '@';
        }

        // Account cell styling matching category tab (NS Debit is bold)
        const r1Acc = r1.getCell('account');
        if (nsDebitAccount) {
          r1Acc.font = { bold: true, color: { argb: 'FF0F172A' } };
          r1Acc.numFmt = '@';
        }

        // Step 2 (Duplicate the process): NS Credit Account with Adj Debit / Adj Credit
        const row2Data = {
          externalId: String(extId),
          date: String(item.date || ''),
          subsidiary: '3',
          memo: memoVal,
          currency: 'US Dollar',
          account: String(nsCreditAccount),
          debit: blankIfZero(adjDebit),
          credit: blankIfZero(adjCredit),
          lineDescription: String(item.description || ''),
          department: nsCreditAccount === '51069' ? 'Finance' : '',
          postingPeriod: '',
          startDate: '',
          endDate: '',
          entityName: '',
          reversal: '',
          reversalDate: '',
          attachments: ''
        };

        const r2 = wsJE.addRow(row2Data);
        r2.height = 20;

        // Ensure text formatting across all standard text columns
        r2.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          if (colNumber !== 7 && colNumber !== 8) {
            cell.numFmt = '@';
          }
        });

        // Highlight Column Q "Attachment" in Red
        const r2Attach = r2.getCell('attachments');
        r2Attach.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFEE2E2' } // Soft Red Highlight
        };
        r2Attach.border = {
          top: { style: 'thin', color: { argb: 'FFFCA5A5' } },
          bottom: { style: 'thin', color: { argb: 'FFFCA5A5' } },
          left: { style: 'thin', color: { argb: 'FFFCA5A5' } },
          right: { style: 'thin', color: { argb: 'FFFCA5A5' } }
        };
        r2Attach.numFmt = '@';

        // Highlight Debit and Credit based on Category Tab details (Adj Debit: Green, Adj Credit: Red)
        const r2Debit = r2.getCell('debit');
        const r2Credit = r2.getCell('credit');
        if (adjDebit && adjDebit > 0.0001) {
          r2Debit.numFmt = '#,##0.00';
          r2Debit.font = { bold: true, color: { argb: COLOR_GREEN } };
          r2Debit.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
          r2Debit.border = { top: { style: 'thin', color: { argb: 'FF86EFAC' } }, bottom: { style: 'thin', color: { argb: 'FF86EFAC' } } };
          r2Debit.alignment = { horizontal: 'right', vertical: 'middle' };
        } else {
          r2Debit.numFmt = '@';
        }

        if (adjCredit && adjCredit > 0.0001) {
          r2Credit.numFmt = '#,##0.00';
          r2Credit.font = { bold: true, color: { argb: COLOR_RED } };
          r2Credit.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
          r2Credit.border = { top: { style: 'thin', color: { argb: 'FFFCA5A5' } }, bottom: { style: 'thin', color: { argb: 'FFFCA5A5' } } };
          r2Credit.alignment = { horizontal: 'right', vertical: 'middle' };
        } else {
          r2Credit.numFmt = '@';
        }

        // Account cell styling matching category tab (NS Credit is red bold)
        const r2Acc = r2.getCell('account');
        if (nsCreditAccount) {
          r2Acc.font = { bold: true, color: { argb: 'FFFF0000' } };
          r2Acc.numFmt = '@';
        }
      });

      // Freeze header row for comfortable scrolling
      wsJE.views = [{ state: 'frozen', ySplit: 1 }];

      const buffer = await workbook.xlsx.writeBuffer();
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
      saveAs(
        new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 
        `Bank_Consolidated_Report_${timestamp}.xlsx`
      );
    } catch (error) {
      console.error('Excel Export Error:', error);
      alert('Failed to generate Excel report. Please check if your browser allows large downloads.');
    }
  };

  return (
    <div className="space-y-8 pb-12">
      {/* Top Banner KPI Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-4 border-b border-slate-200">
        <div>
          <div className="flex items-center gap-2">
            <span className="text-[10px] font-bold uppercase tracking-widest text-indigo-600 bg-indigo-50 px-2.5 py-1 rounded-full border border-indigo-100">
              Stage 4: Executive Review
            </span>
            <span className="text-[10px] font-bold uppercase tracking-widest text-slate-500 bg-slate-100 px-2.5 py-1 rounded-full">
              {uniqueFileNames.length} {uniqueFileNames.length === 1 ? 'File' : 'Files'} Consolidated
            </span>
          </div>
          <h2 className="text-2xl font-bold text-slate-900 tracking-tight mt-2">
            Dashboard Summary & Consolidated Breakdown
          </h2>
          <p className="text-xs text-slate-500 mt-1">
            Review file cross-tabulation, category schedules, and NetSuite JE Upload before exporting.
          </p>
        </div>

        {/* Global Action: Excel Download */}
        <div className="flex items-center gap-3">
          <button
            onClick={exportToExcel}
            className="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-bold text-xs flex items-center gap-2 shadow-sm transition-all active:scale-95"
          >
            <Download size={16} />
            Export Complete Excel (with JE Upload)
          </button>
        </div>
      </div>

      {/* Top Overview KPI Cards */}
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-4">
        <div className="p-5 rounded-2xl bg-white border border-slate-200/80 shadow-xs">
          <div className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-2">Total Files</div>
          <div className="text-2xl font-bold font-mono text-slate-900">{uniqueFileNames.length}</div>
          <div className="text-[10px] text-slate-400 mt-1">Loaded Statements</div>
        </div>

        <div className="p-5 rounded-2xl bg-white border border-slate-200/80 shadow-xs">
          <div className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-2">Total Records</div>
          <div className="text-2xl font-bold font-mono text-slate-900">{transactions.length.toLocaleString()}</div>
          <div className="text-[10px] text-slate-400 mt-1">Reconciled Items</div>
        </div>

        <div className="p-5 rounded-2xl bg-white border border-slate-200/80 shadow-xs">
          <div className="text-[10px] font-bold text-emerald-600 uppercase tracking-wider mb-2 flex items-center gap-1">
            <ArrowDownRight size={12} /> Total Inflow (Credits)
          </div>
          <div className="text-2xl font-bold font-mono text-emerald-700">
            ${formatCurrency(totalCredits)}
          </div>
          <div className="text-[10px] text-slate-400 mt-1">Gross Deposits</div>
        </div>

        <div className="p-5 rounded-2xl bg-white border border-slate-200/80 shadow-xs">
          <div className="text-[10px] font-bold text-rose-600 uppercase tracking-wider mb-2 flex items-center gap-1">
            <ArrowUpRight size={12} /> Total Outflow (Debits)
          </div>
          <div className="text-2xl font-bold font-mono text-rose-700">
            ${formatCurrency(totalDebits)}
          </div>
          <div className="text-[10px] text-slate-400 mt-1">Gross Expenses/Debits</div>
        </div>

        <div className="p-5 rounded-2xl bg-indigo-50/70 border border-indigo-200 shadow-xs col-span-2 sm:col-span-1">
          <div className="text-[10px] font-bold text-indigo-700 uppercase tracking-wider mb-2 flex items-center gap-1">
            <DollarSign size={12} /> Net Consolidated
          </div>
          <div className={cn("text-2xl font-bold font-mono tracking-tight", totalVolume >= 0 ? "text-indigo-900" : "text-rose-700")}>
            ${formatCurrency(totalVolume)}
          </div>
          <div className="text-[10px] text-indigo-600/80 font-medium mt-1">Net Balance Change</div>
        </div>
      </div>

      {/* Category Summary Cards (Category Consolidation Separate) */}
      <div>
        <div className="flex items-center justify-between mb-3">
          <h3 className="text-xs font-bold text-slate-900 uppercase tracking-wider flex items-center gap-2">
            <PieChart size={14} className="text-indigo-500" />
            Category Consolidation (Overview)
          </h3>
          <span className="text-[10px] text-slate-400 uppercase font-semibold">5 Standard Buckets</span>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
          {categoryStats.map((s, i) => (
            <motion.div
              key={s.name}
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ delay: i * 0.04 }}
              className={cn(
                "elevated-card border-none border-t-4 p-5 bg-white transition-all hover:shadow-md",
                CATEGORY_STYLES[s.name as Category]
              )}
            >
              <div className="flex items-center justify-between mb-2">
                <span className="text-[11px] font-bold text-slate-700 truncate" title={s.name}>{s.name}</span>
                <span className="text-[10px] font-mono text-slate-400 font-semibold">{s.percentOfCount.toFixed(0)}%</span>
              </div>
              <div className={cn("text-xl font-bold font-mono tracking-tight", CATEGORY_STYLE_TEXT[s.name as Category])}>
                ${formatCurrency(s.total)}
              </div>
              <div className="flex items-center justify-between mt-3 pt-2 border-t border-slate-100 text-[10px] text-slate-500">
                <span>{s.count} Items</span>
                <span className="font-mono">In: ${formatCurrency(s.credit)}</span>
              </div>
            </motion.div>
          ))}
        </div>
      </div>

      {/* Navigation View Switcher */}
      <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-200 pb-3">
        <div className="flex items-center gap-1.5 p-1 bg-slate-100 rounded-xl">
          <button
            onClick={() => setActiveTab('all')}
            className={cn(
              "px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5",
              activeTab === 'all' 
                ? "bg-white text-slate-900 shadow-xs" 
                : "text-slate-500 hover:text-slate-900"
            )}
          >
            <Layers size={13} />
            Complete Breakdown
          </button>
          <button
            onClick={() => setActiveTab('grouped')}
            className={cn(
              "px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5",
              activeTab === 'grouped' 
                ? "bg-white text-slate-900 shadow-xs" 
                : "text-slate-500 hover:text-slate-900"
            )}
          >
            <FileText size={13} />
            Grouped by File ({fileStats.length})
          </button>
          <button
            onClick={() => setActiveTab('matrix')}
            className={cn(
              "px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5",
              activeTab === 'matrix' 
                ? "bg-white text-slate-900 shadow-xs" 
                : "text-slate-500 hover:text-slate-900"
            )}
          >
            <Grid size={13} />
            File vs Category Matrix
          </button>
          <button
            onClick={() => setActiveTab('separate')}
            className={cn(
              "px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5",
              activeTab === 'separate' 
                ? "bg-white text-slate-900 shadow-xs" 
                : "text-slate-500 hover:text-slate-900"
            )}
          >
            <Table size={13} />
            Separate Tables
          </button>
        </div>

        <div className="text-[11px] text-slate-500 font-medium">
          Showing <span className="font-bold text-slate-900">{fileStats.length} files</span> with total breakdown
        </div>
      </div>

      {/* VIEW SECTION 1: GROUPED BREAKDOWN BY FILE */}
      {(activeTab === 'all' || activeTab === 'grouped') && (
        <div className="space-y-4">
          <div className="flex items-center justify-between">
            <div>
              <h3 className="text-sm font-bold text-slate-900 uppercase tracking-wide flex items-center gap-2">
                <FileText size={16} className="text-indigo-600" />
                Grouped Breakdown By File
              </h3>
              <p className="text-xs text-slate-500">
                Detailed category breakdown and subtotal metrics for each statement file.
              </p>
            </div>
            <div className="flex items-center gap-2">
              <button
                onClick={() => {
                  const allOpen: Record<string, boolean> = {};
                  fileStats.forEach(f => { allOpen[f.name] = true; });
                  setExpandedFiles(allOpen);
                }}
                className="text-[11px] font-semibold text-indigo-600 hover:underline px-2 py-1"
              >
                Expand All
              </button>
              <span className="text-slate-300">|</span>
              <button
                onClick={() => {
                  const allClosed: Record<string, boolean> = {};
                  fileStats.forEach(f => { allClosed[f.name] = false; });
                  setExpandedFiles(allClosed);
                }}
                className="text-[11px] font-semibold text-slate-500 hover:underline px-2 py-1"
              >
                Collapse All
              </button>
            </div>
          </div>

          <div className="space-y-4">
            {fileStats.map((file) => {
              const isOpen = isFileExpanded(file.name);
              return (
                <div 
                  key={file.id} 
                  className="rounded-2xl border border-slate-200 bg-white overflow-hidden shadow-xs transition-all"
                >
                  {/* File Header Bar */}
                  <div 
                    onClick={() => toggleFileExpand(file.name)}
                    className="px-6 py-4 bg-slate-50/80 hover:bg-slate-100/70 cursor-pointer flex flex-col md:flex-row md:items-center justify-between gap-3 border-b border-slate-200/70 transition-colors"
                  >
                    <div className="flex items-center gap-3 min-w-0">
                      <div className="w-8 h-8 rounded-lg bg-indigo-100 text-indigo-700 flex items-center justify-center shrink-0">
                        <FileText size={16} />
                      </div>
                      <div className="min-w-0">
                        <div className="flex items-center gap-2 flex-wrap">
                          <h4 className="text-sm font-bold text-slate-900 truncate" title={file.name}>
                            {file.name}
                          </h4>
                          <span className="text-[10px] font-bold px-2 py-0.5 rounded-md bg-slate-200/80 text-slate-700">
                            {file.count} {file.count === 1 ? 'Record' : 'Records'} ({file.percentOfTotalCount.toFixed(1)}%)
                          </span>
                        </div>
                        <div className="text-[11px] text-slate-500 mt-0.5 flex items-center gap-3">
                          <span>Inflow: <strong className="text-emerald-600 font-mono">${formatCurrency(file.credit)}</strong></span>
                          <span>•</span>
                          <span>Outflow: <strong className="text-rose-600 font-mono">${formatCurrency(file.debit)}</strong></span>
                        </div>
                      </div>
                    </div>

                    <div className="flex items-center justify-between md:justify-end gap-4 shrink-0">
                      <div className="text-right">
                        <div className="text-[10px] font-bold uppercase text-slate-400">File Net Total</div>
                        <div className={cn("text-base font-bold font-mono", file.total >= 0 ? "text-slate-900" : "text-rose-700")}>
                          ${formatCurrency(file.total)}
                        </div>
                      </div>
                      <div className="w-7 h-7 rounded-full bg-white border border-slate-200 flex items-center justify-center text-slate-500">
                        {isOpen ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
                      </div>
                    </div>
                  </div>

                  {/* Collapsible File Category Breakdown Table */}
                  <AnimatePresence>
                    {isOpen && (
                      <motion.div
                        initial={{ opacity: 0, height: 0 }}
                        animate={{ opacity: 1, height: 'auto' }}
                        exit={{ opacity: 0, height: 0 }}
                        transition={{ duration: 0.2 }}
                      >
                        <div className="p-4 overflow-x-auto">
                          <table className="w-full text-left border-collapse">
                            <thead>
                              <tr className="border-b border-slate-200 text-[10px] font-bold text-slate-500 uppercase tracking-wider">
                                <th className="py-2.5 px-4">Category</th>
                                <th className="py-2.5 px-4 text-center">Transactions</th>
                                <th className="py-2.5 px-4 text-right">Debits (Outflow)</th>
                                <th className="py-2.5 px-4 text-right">Credits (Inflow)</th>
                                <th className="py-2.5 px-4 text-right">Category Net Total</th>
                                <th className="py-2.5 px-4 text-right">% of File</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100 text-xs">
                              {file.categories.map((cat) => (
                                <tr 
                                  key={cat.category}
                                  className={cn(
                                    "hover:bg-slate-50/60 transition-colors",
                                    cat.count === 0 && "opacity-40"
                                  )}
                                >
                                  <td className="py-2.5 px-4 font-semibold text-slate-800 flex items-center gap-2">
                                    <span className={cn(
                                      "inline-block w-2 h-2 rounded-full",
                                      cat.category === 'Chargebacks' && "bg-red-500",
                                      cat.category === 'Merchant Fees' && "bg-orange-500",
                                      cat.category === 'Settlement Deposit' && "bg-emerald-500",
                                      cat.category === 'Bank Transfer' && "bg-blue-500",
                                      cat.category === 'Miscellaneous' && "bg-slate-400"
                                    )} />
                                    {cat.category}
                                  </td>
                                  <td className="py-2.5 px-4 text-center font-mono">
                                    {cat.count}
                                  </td>
                                  <td className="py-2.5 px-4 text-right font-mono text-slate-600">
                                    {cat.debit > 0 ? `$${formatCurrency(cat.debit)}` : '—'}
                                  </td>
                                  <td className="py-2.5 px-4 text-right font-mono text-slate-600">
                                    {cat.credit > 0 ? `$${formatCurrency(cat.credit)}` : '—'}
                                  </td>
                                  <td className={cn(
                                    "py-2.5 px-4 text-right font-mono font-bold",
                                    CATEGORY_STYLE_TEXT[cat.category as Category]
                                  )}>
                                    ${formatCurrency(cat.total)}
                                  </td>
                                  <td className="py-2.5 px-4 text-right font-mono text-slate-400">
                                    {cat.percentOfFileCount.toFixed(0)}%
                                  </td>
                                </tr>
                              ))}

                              {/* File Subtotal Row */}
                              <tr className="bg-slate-50 font-bold border-t border-slate-200 text-slate-900">
                                <td className="py-3 px-4 uppercase text-[11px] tracking-wider">
                                  File Subtotal ({file.name})
                                </td>
                                <td className="py-3 px-4 text-center font-mono">
                                  {file.count}
                                </td>
                                <td className="py-3 px-4 text-right font-mono text-rose-700">
                                  ${formatCurrency(file.debit)}
                                </td>
                                <td className="py-3 px-4 text-right font-mono text-emerald-700">
                                  ${formatCurrency(file.credit)}
                                </td>
                                <td className="py-3 px-4 text-right font-mono text-sm">
                                  ${formatCurrency(file.total)}
                                </td>
                                <td className="py-3 px-4 text-right font-mono">
                                  100%
                                </td>
                              </tr>
                            </tbody>
                          </table>
                        </div>
                      </motion.div>
                    )}
                  </AnimatePresence>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* VIEW SECTION 2: CROSS-TABULATION MATRIX TABLE (FILE x CATEGORY) */}
      {(activeTab === 'all' || activeTab === 'matrix') && (
        <div className="space-y-3">
          <div className="flex items-center justify-between">
            <div>
              <h3 className="text-sm font-bold text-slate-900 uppercase tracking-wide flex items-center gap-2">
                <Grid size={16} className="text-indigo-600" />
                Cross-Tabulation: File by Category Matrix
              </h3>
              <p className="text-xs text-slate-500">
                At-a-glance comparison showing net volume for each category across all loaded statements.
              </p>
            </div>
          </div>

          <div className="rounded-2xl border border-slate-200 bg-white overflow-hidden shadow-xs">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-100/80 border-b border-slate-200 text-[10px] font-bold text-slate-600 uppercase tracking-wider">
                    <th className="py-3.5 px-4 min-w-[200px]">File Name</th>
                    <th className="py-3.5 px-3 text-right text-red-700">Chargebacks</th>
                    <th className="py-3.5 px-3 text-right text-orange-700">Merchant Fees</th>
                    <th className="py-3.5 px-3 text-right text-emerald-700">Settlement Dep.</th>
                    <th className="py-3.5 px-3 text-right text-blue-700">Bank Transfer</th>
                    <th className="py-3.5 px-3 text-right text-slate-600">Miscellaneous</th>
                    <th className="py-3.5 px-4 text-right bg-slate-200/50 font-bold text-slate-900">File Net Total</th>
                    <th className="py-3.5 px-3 text-center">Items</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {fileStats.map((file) => {
                    const catMap = new Map(file.categories.map(c => [c.category, c.total]));
                    return (
                      <tr key={file.id} className="hover:bg-slate-50/70 transition-colors">
                        <td className="py-3 px-4 font-semibold text-slate-900 truncate max-w-[240px]" title={file.name}>
                          <div className="flex items-center gap-2">
                            <FileText size={14} className="text-slate-400 shrink-0" />
                            <span className="truncate">{file.name}</span>
                          </div>
                        </td>
                        <td className="py-3 px-3 text-right font-mono">
                          {catMap.get('Chargebacks') ? `$${formatCurrency(catMap.get('Chargebacks')!)}` : '—'}
                        </td>
                        <td className="py-3 px-3 text-right font-mono">
                          {catMap.get('Merchant Fees') ? `$${formatCurrency(catMap.get('Merchant Fees')!)}` : '—'}
                        </td>
                        <td className="py-3 px-3 text-right font-mono">
                          {catMap.get('Settlement Deposit') ? `$${formatCurrency(catMap.get('Settlement Deposit')!)}` : '—'}
                        </td>
                        <td className="py-3 px-3 text-right font-mono">
                          {catMap.get('Bank Transfer') ? `$${formatCurrency(catMap.get('Bank Transfer')!)}` : '—'}
                        </td>
                        <td className="py-3 px-3 text-right font-mono">
                          {catMap.get('Miscellaneous') ? `$${formatCurrency(catMap.get('Miscellaneous')!)}` : '—'}
                        </td>
                        <td className="py-3 px-4 text-right font-mono font-bold bg-slate-50/60">
                          ${formatCurrency(file.total)}
                        </td>
                        <td className="py-3 px-3 text-center font-mono text-slate-500">
                          {file.count}
                        </td>
                      </tr>
                    );
                  })}

                  {/* Matrix Grand Consolidated Row */}
                  <tr className="bg-slate-100 font-bold border-t-2 border-slate-300 text-slate-900">
                    <td className="py-4 px-4 uppercase text-[11px] tracking-wider">
                      Consolidated Total
                    </td>
                    {CATEGORIES.map(cat => {
                      const stat = categoryStats.find(s => s.name === cat);
                      return (
                        <td key={cat} className="py-4 px-3 text-right font-mono">
                          ${formatCurrency(stat ? stat.total : 0)}
                        </td>
                      );
                    })}
                    <td className="py-4 px-4 text-right font-mono text-sm bg-slate-200/80 text-indigo-950 font-bold">
                      ${formatCurrency(totalVolume)}
                    </td>
                    <td className="py-4 px-3 text-center font-mono">
                      {transactions.length}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* VIEW SECTION 3: SEPARATE OVERVIEWS (FILE SUMMARY + CATEGORY CONSOLIDATION) */}
      {(activeTab === 'all' || activeTab === 'separate') && (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* Separate File Breakdown Table */}
          <div className="rounded-2xl border border-slate-200 bg-white overflow-hidden shadow-xs flex flex-col">
            <div className="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
              <h3 className="text-xs font-bold text-slate-900 uppercase tracking-wider flex items-center gap-2">
                <FileText size={15} className="text-indigo-600" />
                Total Breakdown By File Name (Separate)
              </h3>
              <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-indigo-50 text-indigo-700 border border-indigo-100">
                {fileStats.length} Files
              </span>
            </div>

            <div className="p-0 overflow-x-auto flex-1">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="border-b border-slate-200 text-[10px] font-bold text-slate-500 uppercase tracking-wider bg-slate-50/50">
                    <th className="py-2.5 px-4">File Name</th>
                    <th className="py-2.5 px-3 text-center">Items</th>
                    <th className="py-2.5 px-3 text-right">Inflow ($)</th>
                    <th className="py-2.5 px-3 text-right">Outflow ($)</th>
                    <th className="py-2.5 px-4 text-right">Net Amount</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {fileStats.map(f => (
                    <tr key={f.id} className="hover:bg-slate-50/50 transition-colors">
                      <td className="py-3 px-4 font-semibold text-slate-800 truncate max-w-[180px]" title={f.name}>
                        {f.name}
                      </td>
                      <td className="py-3 px-3 text-center font-mono text-slate-600">
                        {f.count}
                      </td>
                      <td className="py-3 px-3 text-right font-mono text-emerald-600">
                        ${formatCurrency(f.credit)}
                      </td>
                      <td className="py-3 px-3 text-right font-mono text-rose-600">
                        ${formatCurrency(f.debit)}
                      </td>
                      <td className="py-3 px-4 text-right font-mono font-bold text-slate-900">
                        ${formatCurrency(f.total)}
                      </td>
                    </tr>
                  ))}
                  <tr className="bg-slate-50 font-bold border-t border-slate-200 text-slate-900">
                    <td className="py-3 px-4 uppercase text-[10px]">Total All Files</td>
                    <td className="py-3 px-3 text-center font-mono">{transactions.length}</td>
                    <td className="py-3 px-3 text-right font-mono text-emerald-700">${formatCurrency(totalCredits)}</td>
                    <td className="py-3 px-3 text-right font-mono text-rose-700">${formatCurrency(totalDebits)}</td>
                    <td className="py-3 px-4 text-right font-mono text-indigo-900">${formatCurrency(totalVolume)}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          {/* Separate Category Consolidation Table */}
          <div className="rounded-2xl border border-slate-200 bg-white overflow-hidden shadow-xs flex flex-col">
            <div className="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
              <h3 className="text-xs font-bold text-slate-900 uppercase tracking-wider flex items-center gap-2">
                <PieChart size={15} className="text-indigo-600" />
                Category Consolidation (Separate)
              </h3>
              <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-indigo-50 text-indigo-700 border border-indigo-100">
                5 Categories
              </span>
            </div>

            <div className="p-0 overflow-x-auto flex-1">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="border-b border-slate-200 text-[10px] font-bold text-slate-500 uppercase tracking-wider bg-slate-50/50">
                    <th className="py-2.5 px-4">Category</th>
                    <th className="py-2.5 px-3 text-center">Items</th>
                    <th className="py-2.5 px-3 text-right">Debits ($)</th>
                    <th className="py-2.5 px-3 text-right">Credits ($)</th>
                    <th className="py-2.5 px-4 text-right">Net Value</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {categoryStats.map(s => (
                    <tr key={s.name} className="hover:bg-slate-50/50 transition-colors">
                      <td className="py-3 px-4 font-semibold text-slate-800 flex items-center gap-2">
                        <span className={cn(
                          "px-2 py-0.5 rounded text-[10px] font-bold border",
                          CATEGORY_BADGE_BG[s.name as Category]
                        )}>
                          {s.name}
                        </span>
                      </td>
                      <td className="py-3 px-3 text-center font-mono text-slate-600">
                        {s.count}
                      </td>
                      <td className="py-3 px-3 text-right font-mono text-rose-600">
                        ${formatCurrency(s.debit)}
                      </td>
                      <td className="py-3 px-3 text-right font-mono text-emerald-600">
                        ${formatCurrency(s.credit)}
                      </td>
                      <td className={cn("py-3 px-4 text-right font-mono font-bold", CATEGORY_STYLE_TEXT[s.name as Category])}>
                        ${formatCurrency(s.total)}
                      </td>
                    </tr>
                  ))}
                  <tr className="bg-slate-50 font-bold border-t border-slate-200 text-slate-900">
                    <td className="py-3 px-4 uppercase text-[10px]">Total Consolidated</td>
                    <td className="py-3 px-3 text-center font-mono">{transactions.length}</td>
                    <td className="py-3 px-3 text-right font-mono text-rose-700">${formatCurrency(totalDebits)}</td>
                    <td className="py-3 px-3 text-right font-mono text-emerald-700">${formatCurrency(totalCredits)}</td>
                    <td className="py-3 px-4 text-right font-mono text-indigo-900">${formatCurrency(totalVolume)}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* Bottom Final Export Banner */}
      <div className="rounded-2xl bg-gradient-to-r from-slate-900 to-indigo-950 p-6 sm:p-8 text-white flex flex-col md:flex-row md:items-center justify-between gap-6 shadow-md">
        <div className="space-y-1">
          <div className="flex items-center gap-2 text-indigo-300 text-xs font-bold uppercase tracking-wider">
            <CheckCircle2 size={15} className="text-emerald-400" />
            Audit Complete · Reconciled
          </div>
          <h4 className="text-xl font-bold tracking-tight">Ready to Export Your Full Financial Package?</h4>
          <p className="text-xs text-slate-300 max-w-2xl leading-relaxed">
            The generated Excel workbook will include this complete Dashboard Summary (with File Summary, Category Consolidation, Cross-Tab Matrix, and Grouped File Breakdown), plus individual dedicated worksheets for Chargebacks, Merchant Fees, Settlement Deposit, Bank Transfer, and Miscellaneous.
          </p>
        </div>

        <button
          onClick={exportToExcel}
          className="px-6 py-3.5 bg-indigo-500 hover:bg-indigo-400 text-white font-bold rounded-xl text-xs flex items-center justify-center gap-2 shrink-0 shadow-sm transition-all active:scale-95"
        >
          <Download size={16} />
          Download Consolidated Excel
        </button>
      </div>
    </div>
  );
}


