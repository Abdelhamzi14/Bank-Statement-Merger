import { type BankTransaction, type Category, type StatementFile } from '../types';
import { Download, Table, AlertCircle } from 'lucide-react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { motion } from 'motion/react';
import { cn } from '../lib/utils';

interface Props {
  files: StatementFile[];
  transactions: BankTransaction[];
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

export default function SummaryAndExport({ files, transactions }: Props) {
  const stats = CATEGORIES.map(category => {
    const items = transactions.filter(t => t.category === category);
    const total = items.reduce((sum, item) => sum + item.amount, 0);
    return { name: category, count: items.length, total };
  });

  const accountStats = files.map(file => {
    const items = transactions.filter(t => t.sourceFile === file.file.name);
    const total = items.reduce((sum, item) => sum + item.amount, 0);
    return { name: file.file.name, count: items.length, total };
  });

  const exportToExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'MergeXpress';
    workbook.lastModifiedBy = 'MergeXpress';
    workbook.created = new Date();

    const headerStyle: Partial<ExcelJS.Style> = {
      font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } }, // Dark Blue (Indigo-900ish)
      alignment: { horizontal: 'center', vertical: 'middle' },
      border: {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      }
    };

    // 1. Dashboard Summary
    const wsSummary = workbook.addWorksheet('Dashboard Summary', {
      views: [{ state: 'frozen', ySplit: 0, xSplit: 0 }]
    });

    wsSummary.addRow(['Bank Reconciliation Consolidated Report']).font = { bold: true, size: 14, italic: true };
    wsSummary.addRow([`Report Generated: ${new Date().toLocaleString()}`]).font = { italic: true, size: 10 };
    wsSummary.addRow([]);

    const fileHeader = wsSummary.addRow(['FILE ACCOUNT SUMMARY']);
    fileHeader.font = { bold: true, size: 12 };
    const fileSubHeader = wsSummary.addRow(['File Name', 'Transactions Count', 'Net Amount']);
    fileSubHeader.eachCell((cell) => cell.style = headerStyle as ExcelJS.Style);

    accountStats.forEach(s => {
      const row = wsSummary.addRow([s.name, s.count, s.total]);
      row.getCell(3).numFmt = '#,##0.00';
    });

    wsSummary.addRow([]);
    const catHeader = wsSummary.addRow(['CATEGORY CONSOLIDATION']);
    catHeader.font = { bold: true, size: 12 };
    const catSubHeader = wsSummary.addRow(['Category Name', 'Line Items', 'Total Vol.']);
    catSubHeader.eachCell((cell) => cell.style = headerStyle as ExcelJS.Style);

    stats.forEach(s => {
      const row = wsSummary.addRow([s.name, s.count, s.total]);
      row.getCell(3).numFmt = '#,##0.00';
    });

    wsSummary.addRow([]);
    const totalRow = wsSummary.addRow(['TOTAL CONSOLIDATED POSITION', '', transactions.reduce((s, t) => s + t.amount, 0)]);
    totalRow.font = { bold: true };
    totalRow.getCell(3).numFmt = '#,##0.00';

    wsSummary.getColumn(1).width = 45;
    wsSummary.getColumn(2).width = 20;
    wsSummary.getColumn(3).width = 20;

    // 2. Category Tabs
    CATEGORIES.forEach(category => {
      const categoryItems = transactions.filter(t => t.category === category);
      if (categoryItems.length > 0) {
        const ws = workbook.addWorksheet(category.substring(0, 31), {
          views: [{ state: 'frozen', ySplit: 1 }]
        });

        const isSpecialTab = ['Chargebacks', 'Merchant Fees', 'Bank Transfer', 'Miscellaneous'].includes(category);

        const columns = [
          { header: 'Date', key: 'date', width: 15 },
          { header: 'Description', key: 'description', width: 50 },
          { header: 'Account#', key: 'accountNumber', width: 20 },
          { header: 'Amount', key: 'amount', width: 15 },
        ];

        if (isSpecialTab) {
          columns.push(
            { header: 'NetSuite Account Number - Debit', key: 'nsDebit', width: 30 },
            { header: 'Debit', key: 'debit', width: 15 },
            { header: 'Credit', key: 'credit', width: 15 },
            { header: 'NetSuite Account Number - Credit', key: 'nsCredit', width: 30 },
            { header: 'Debit', key: 'swappedDebit', width: 15 },
            { header: 'Credit', key: 'swappedCredit', width: 15 },
            { header: 'Source File', key: 'sourceFile', width: 30 }
          );
        } else {
          columns.push(
            { header: 'Debit', key: 'debit', width: 15 },
            { header: 'Credit', key: 'credit', width: 15 },
            { header: 'Source File', key: 'sourceFile', width: 30 }
          );
        }

        ws.columns = columns;

        // Style header row
        const rowHeader = ws.getRow(1);
        rowHeader.height = 25;
        rowHeader.eachCell((cell) => {
          cell.style = headerStyle as ExcelJS.Style;
        });

        categoryItems.forEach(item => {
          const blankIfZero = (val: number | null) => (val === 0 || val === null) ? null : val;

          const rowData: any = {
            date: item.date,
            description: item.description,
            accountNumber: item.accountNumber || 'N/A',
            amount: item.amount,
            debit: blankIfZero(item.debit),
            credit: blankIfZero(item.credit),
            sourceFile: item.sourceFile
          };

          if (isSpecialTab) {
            rowData.nsDebit = item.netsuiteDebitAccount || '';
            rowData.nsCredit = item.netsuiteCreditAccount || '';
            rowData.swappedDebit = blankIfZero(item.credit);
            rowData.swappedCredit = blankIfZero(item.debit);
          }

          const r = ws.addRow(rowData);
          r.getCell('amount').numFmt = '#,##0.00';
          r.getCell('debit').numFmt = '#,##0.00';
          r.getCell('credit').numFmt = '#,##0.00';
          
          if (isSpecialTab) {
            const creditCell = r.getCell('nsCredit');
            creditCell.font = { color: { argb: 'FFFF0000' }, bold: true };
            r.getCell('swappedDebit').numFmt = '#,##0.00';
            r.getCell('swappedCredit').numFmt = '#,##0.00';
          }
        });
      }
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
    saveAs(new Blob([buffer]), `Bank_Consolidated_Report_${timestamp}.xlsx`);
  };

  const CATEGORY_STYLE_TEXT: Record<Category, string> = {
    'Chargebacks': 'text-red-700',
    'Merchant Fees': 'text-orange-600',
    'Settlement Deposit': 'text-emerald-700',
    'Bank Transfer': 'text-blue-600',
    'Miscellaneous': 'text-slate-600'
  };

  return (
    <div className="space-y-8 pb-12">
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-6">
        {stats.map((s, i) => (
          <motion.div
            key={s.name}
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: i * 0.05 }}
            className={cn(
              "elevated-card border-none border-t-4 p-6",
              CATEGORY_STYLES[s.name as Category]
            )}
          >
            <div className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-3">{s.name}</div>
            <div className={cn("text-2xl font-bold font-mono tracking-tighter", CATEGORY_STYLE_TEXT[s.name as Category])}>
              {s.total.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
            </div>
            <div className={cn(
              "text-[10px] font-semibold uppercase mt-2 px-2 py-0.5 rounded-full inline-block",
              s.total >= 0 ? "bg-emerald-50 text-emerald-700" : "bg-rose-50 text-rose-700"
            )}>
              {s.count} Line Items
            </div>
          </motion.div>
        ))}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
        <div className="lg:col-span-2 space-y-6">
          <div className="modern-panel">
            <div className="border-b border-slate-100 px-8 py-5 flex items-center justify-between bg-slate-50/50">
              <h3 className="text-xs font-bold text-slate-900 uppercase tracking-widest flex items-center gap-2">
                <Table size={14} className="text-indigo-500" />
                Account Reconciliation Status
              </h3>
            </div>
            <div className="divide-y divide-slate-100">
              {accountStats.map(s => (
                <div key={s.name} className="px-8 py-5 flex items-center justify-between hover:bg-slate-50/50 transition-colors">
                  <div className="overflow-hidden pr-4">
                    <p className="text-sm font-bold text-slate-900 tracking-tight truncate mb-1" title={s.name}>{s.name}</p>
                    <p className="text-[10px] text-slate-400 font-medium uppercase tracking-wider">Mapping verified · {s.count} records</p>
                  </div>
                  <p className="text-sm font-bold text-slate-900 font-mono">
                    {s.total.toLocaleString(undefined, { minimumFractionDigits: 2 })}
                  </p>
                </div>
              ))}
            </div>
          </div>
        </div>

        <div className="space-y-6">
          <div className="bg-indigo-600 rounded-2xl p-8 text-white relative overflow-hidden shadow-lg shadow-indigo-100">
            <div className="relative z-10">
              <div className="text-[10px] font-bold uppercase tracking-[0.2em] mb-4 text-indigo-200">Reconciliation Ready</div>
              <h3 className="text-2xl font-bold mb-3 tracking-tight">Generate Report</h3>
              <p className="text-indigo-100/70 text-sm mb-8 leading-relaxed">
                All transaction data has been verified and categorized. Click below to download the consolidated Excel workbook.
              </p>
              <button
                onClick={exportToExcel}
                className="w-full bg-white hover:bg-slate-50 text-indigo-600 font-bold py-4 px-6 rounded-xl flex items-center justify-center gap-3 transition-all active:scale-95 shadow-md shadow-indigo-900/20"
              >
                <Download size={20} />
                <span className="text-sm">Download Excel</span>
              </button>
            </div>
            <div className="absolute top-0 right-0 -translate-y-1/4 translate-x-1/4 w-32 h-32 bg-white/10 rounded-full blur-3xl opacity-50" />
          </div>

          <div className="modern-panel p-8">
            <div className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-4">Total Consolidated Value</div>
            <p className="text-3xl font-bold font-mono text-slate-900 tracking-tighter">
              {transactions.reduce((s, t) => s + t.amount, 0).toLocaleString(undefined, { minimumFractionDigits: 2 })}
            </p>
            <div className="flex items-center gap-2 mt-3">
              <div className="w-2 h-2 rounded-full bg-emerald-500" />
              <span className="text-[10px] text-slate-500 font-bold uppercase tracking-wider">Verified USD Liquidity</span>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

