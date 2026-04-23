import { useState } from 'react';
import { type StatementFile, type BankTransaction, type Category } from './types';
import { categorizeTransaction } from './lib/categorizer';
import FileUploadZone from './components/FileUploadZone';
import GroupingStep from './components/GroupingStep';
import JournalEntryStep from './components/JournalEntryStep';
import SummaryAndExport from './components/SummaryAndExport';
import { motion, AnimatePresence } from 'motion/react';
import { ChevronRight, ChevronLeft, CreditCard, CheckCircle, Info } from 'lucide-react';
import { cn } from './lib/utils';

export default function App() {
  const [step, setStep] = useState<1 | 2 | 3 | 4>(1);
  const [files, setFiles] = useState<StatementFile[]>([]);
  const [transactions, setTransactions] = useState<BankTransaction[]>([]);

  const getNetSuiteDebitAccount = (category: Category) => {
    switch (category) {
      case 'Chargebacks': return '42050';
      case 'Merchant Fees': return '51069';
      default: return '';
    }
  };

  const getNetSuiteCreditAccount = (accNum: string) => {
    const cleanAcc = accNum.trim().replace(/^0+/, ''); // Remove leading zeros for 0258 case
    const rawAcc = accNum.trim();
    
    if (rawAcc === '9616') return '11421';
    if (rawAcc === '9733') return '11003';
    if (rawAcc === '9725') return '11002';
    if (rawAcc === '9577') return '11001';
    if (rawAcc === '0231') return '11423';
    if (rawAcc === '1831') return '11424';
    if (rawAcc === '0258' || rawAcc === '258') return '11422';
    if (rawAcc === '8703') return '11420';
    
    // Also check clean version just in case
    if (cleanAcc === '231') return '11423';
    if (cleanAcc === '258') return '11422';
    
    return '';
  };

  const handleProcessFiles = () => {
    const allTransactions: BankTransaction[] = [];
    
    files.forEach(file => {
      file.data.forEach((row, index) => {
        const descField = file.mappings.description;
        const dateField = file.mappings.date;
        const amtField = file.mappings.amount;
        const debitField = file.mappings.debit;
        const creditField = file.mappings.credit;
        const accField = file.mappings.accountNumber;

        // More lenient check - we just need a description to show something
        if (row[descField]) {
          const description = String(row[descField] || '').trim();
          if (!description) return; // Skip truly empty lines

          const category = categorizeTransaction(description);
          
          let amount = 0;
          let debit: number | null = null;
          let credit: number | null = null;

          const parseAmt = (val: any) => {
            if (val === undefined || val === null || val === '') return 0;
            let s = String(val).trim();
            
            // Handle accounting format: (1,234.56) -> -1234.56
            if (s.startsWith('(') && s.endsWith(')')) {
              s = '-' + s.substring(1, s.length - 1);
            }
            
            // Remove currency symbols, commas, and whitespace
            // We keep digits, decimal points, and leading minus
            const clean = s.replace(/[^\d.-]/g, '');
            
            // Handle edge case where multiple dots might exist (common in some locales, though usually we expect one)
            // Or if the string was just "$" it results in ""
            if (!clean || clean === '-') return 0;
            
            const parsed = parseFloat(clean);
            return isNaN(parsed) ? 0 : parsed;
          };

          if (amtField && row[amtField] !== undefined && row[amtField] !== '') {
            amount = parseAmt(row[amtField]);
            if (amount < 0) {
              debit = Math.abs(amount);
              credit = 0;
            } else {
              credit = amount;
              debit = 0;
            }
          } else {
            debit = parseAmt(row[debitField]);
            credit = parseAmt(row[creditField]);
            
            // If we have separate debit/credit, amount is net
            // But if we accidentally mapped the SAME column to both, 
            // and it's a signed amount column, this logic fails.
            // Check if debitField and creditField are distinct
            if (debitField === creditField) {
              amount = debit; // Fallback to raw value
              if (amount < 0) {
                debit = Math.abs(amount);
                credit = 0;
              } else {
                credit = amount;
                debit = 0;
              }
            } else {
              amount = (credit || 0) - (debit || 0);
            }
          }

          const date = dateField ? String(row[dateField] || '') : 'N/A';
          const accountNumber = accField ? String(row[accField] || '') : '';

          allTransactions.push({
            id: `${file.id}-${index}`,
            sourceFile: file.file.name,
            date,
            description,
            amount,
            debit,
            credit,
            accountNumber,
            category,
            originalCategory: category,
            netsuiteDebitAccount: getNetSuiteDebitAccount(category),
            netsuiteCreditAccount: getNetSuiteCreditAccount(accountNumber),
            raw: row
          });
        }
      });
    });

    setTransactions(allTransactions);
    setStep(2);
  };

  const updateTransactionCategory = (id: string, category: Category) => {
    setTransactions(prev => prev.map(t => t.id === id ? { 
      ...t, 
      category,
      netsuiteDebitAccount: getNetSuiteDebitAccount(category)
    } : t));
  };

  const bulkUpdateTransactionCategories = (ids: string[], category: Category) => {
    const debitAcc = getNetSuiteDebitAccount(category);
    setTransactions(prev => prev.map(t => ids.includes(t.id) ? {
      ...t,
      category,
      netsuiteDebitAccount: debitAcc
    } : t));
  };

  const updateTransactionRecord = (id: string, updates: Partial<BankTransaction>) => {
    setTransactions(prev => prev.map(t => t.id === id ? { ...t, ...updates } : t));
  };

  const handleNext = () => {
    if (step === 1) handleProcessFiles();
    else if (step === 2) setStep(3);
    else if (step === 3) setStep(4);
  };

  const handleReset = () => {
    setFiles([]);
    setTransactions([]);
    setStep(1);
  };

  const handleBack = () => {
    if (step === 2) setStep(1);
    else if (step === 3) setStep(2);
    else if (step === 4) setStep(3);
  };

  const stepInfo = {
    1: { title: 'Upload Statements', desc: 'Securely add your bank CSV files to begin processing.' },
    2: { title: 'Review & Group', desc: 'Transactions are auto-categorized. Review and adjust as needed.' },
    3: { title: 'Journal Entry', desc: 'Verify and adjust NetSuite account numbers for accuracy.' },
    4: { title: 'Summary & Export', desc: 'View consolidated results and download your final report.' }
  };

  return (
    <div className="min-h-screen bg-slate-50/50 flex flex-col font-sans">
      <header className="bg-white/80 backdrop-blur-md border-b border-slate-200/60 sticky top-0 z-30 h-16 flex items-center shadow-sm">
        <div className="container mx-auto px-8 flex items-center justify-between">
          <div className="flex items-center space-x-3 text-left">
            <div className="w-9 h-9 bg-indigo-600 flex items-center justify-center rounded-lg shadow-md shadow-indigo-200">
              <span className="font-bold text-white text-xs">SM</span>
            </div>
            <div>
              <h1 className="text-lg font-bold tracking-tight text-slate-900 leading-none">
                Bank <span className="font-medium text-slate-500">Statement Merger</span>
              </h1>
              <div className="flex items-center gap-1.5 mt-1">
                <div className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Local Processing Only</span>
              </div>
            </div>
          </div>

          <div className="hidden md:flex items-center space-x-8">
            <StepBadge num={1} label="Upload" active={step === 1} completed={step > 1} />
            <StepBadge num={2} label="Organize" active={step === 2} completed={step > 2} />
            <StepBadge num={3} label="Journal" active={step === 3} completed={step > 3} />
            <StepBadge num={4} label="Review" active={step === 4} completed={step > 4} />
          </div>

          <div className="flex items-center space-x-3">
            {step > 1 && (
              <button
                onClick={handleReset}
                className="px-4 py-2 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 transition-all cursor-pointer text-xs font-semibold"
              >
                Reset
              </button>
            )}
            {step > 1 && (
              <button
                onClick={handleBack}
                className="px-3 py-1.5 rounded-sm border border-slate-200 hover:bg-slate-50 text-slate-600 transition-all cursor-pointer text-[10px] font-bold uppercase tracking-wider"
              >
                Back
              </button>
            )}
            <button
              onClick={handleNext}
              disabled={step === 1 && files.length === 0}
              className={cn(
                "px-4 py-1.5 rounded-sm font-bold uppercase tracking-wider text-[10px] flex items-center space-x-2 transition-all active:scale-95 disabled:opacity-50 disabled:pointer-events-none cursor-pointer",
                step === 4 
                  ? "hidden" 
                  : "bg-indigo-600 text-white shadow-sm hover:bg-indigo-700"
              )}
            >
              <span>
                {step === 1 ? 'Start Processing' : 
                 step === 2 ? 'Review Accounts' : 
                 step === 3 ? 'Finalize Report' : 'Export'}
              </span>
            </button>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 container mx-auto px-8 py-10 max-w-[1600px]">
        <div className="mb-8 text-left">
          <motion.div
            key={step}
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.4 }}
          >
            <h2 className="text-sm font-bold text-slate-700 uppercase tracking-widest mb-1 flex items-center gap-2">
              {stepInfo[step].title}
              <span className="h-0.5 w-12 bg-indigo-600/20" />
            </h2>
            <p className="text-[11px] text-slate-500 max-w-2xl italic leading-relaxed">
              {stepInfo[step].desc}
            </p>
          </motion.div>
        </div>

        <section className="relative">
          <AnimatePresence mode="wait">
            {step === 1 && (
              <motion.div
                key="step1"
                initial={{ opacity: 0, scale: 0.98, y: 10 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 1.02, y: -10 }}
                transition={{ duration: 0.3 }}
                className="elevated-card p-8 min-h-[400px]"
              >
                <FileUploadZone files={files} onFilesChange={setFiles} />
              </motion.div>
            )}

            {step === 2 && (
              <motion.div
                key="step2"
                initial={{ opacity: 0, scale: 0.98, y: 10 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 1.02, y: -10 }}
                transition={{ duration: 0.3 }}
                className="elevated-card p-8 min-h-[500px]"
              >
                <GroupingStep 
                  transactions={transactions} 
                  onUpdateCategory={updateTransactionCategory} 
                  onBulkUpdateCategory={bulkUpdateTransactionCategories}
                />
              </motion.div>
            )}

            {step === 3 && (
              <motion.div
                key="step3"
                initial={{ opacity: 0, scale: 0.98, y: 10 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 1.02, y: -10 }}
                transition={{ duration: 0.3 }}
                className="elevated-card p-8 min-h-[500px]"
              >
                <JournalEntryStep 
                  transactions={transactions} 
                  onUpdateTransaction={updateTransactionRecord} 
                />
              </motion.div>
            )}

            {step === 4 && (
              <motion.div
                key="step4"
                initial={{ opacity: 0, scale: 0.98, y: 10 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 1.02, y: -10 }}
                transition={{ duration: 0.3 }}
                className="elevated-card p-8 min-h-[500px]"
              >
                <SummaryAndExport files={files} transactions={transactions} />
              </motion.div>
            )}
          </AnimatePresence>
        </section>

        {step === 1 && files.length === 0 && (
          <div className="mt-12 p-6 bg-blue-50 border border-blue-100 rounded-3xl flex items-start space-x-4 max-w-3xl mx-auto text-left">
            <div className="bg-blue-100 p-2 rounded-xl text-blue-600 shrink-0">
              <Info size={20} />
            </div>
            <div>
              <h4 className="text-sm font-bold text-blue-900 mb-1">Getting Started</h4>
              <p className="text-xs text-blue-700 leading-relaxed">
                Upload your CSV bank statements. We'll automatically find column headers like Date, Description, and Amount. After processing, you can refine categories like <strong>Merchant Fees</strong> or <strong>Chargebacks</strong>.
              </p>
            </div>
          </div>
        )}
      </main>

      <footer className="h-12 bg-slate-800 text-slate-400 px-8 flex items-center justify-between shrink-0">
        <div className="flex items-center gap-6">
          <span className="text-[9px] uppercase tracking-tighter">&copy; {new Date().getFullYear()} MergerXpress</span>
          <div className="flex items-center gap-2 px-2 py-0.5 bg-slate-700/50 rounded border border-slate-700">
            <span className="w-1.5 h-1.5 rounded-full bg-emerald-400" />
            <span className="text-[9px] uppercase tracking-tighter text-emerald-400 font-bold">100% Client-Side Private</span>
          </div>
        </div>
        <div className="text-[9px] italic opacity-60">
          No data is ever sent to or stored on our servers. Processing occurs entirely in your browser.
        </div>
      </footer>
    </div>
  );
}

function StepBadge({ num, label, active, completed }: { num: number; label: string; active: boolean; completed: boolean }) {
  return (
    <div className={cn(
      "flex items-center gap-2 transition-all duration-300 py-1.5 px-3 rounded-full",
      active ? "bg-indigo-50 border border-indigo-100 shadow-sm" : "opacity-50"
    )}>
      <div className={cn(
        "w-6 h-6 rounded-full flex items-center justify-center text-[11px] font-bold transition-all",
        completed ? "bg-emerald-500 text-white" : active ? "bg-indigo-600 text-white" : "bg-white border border-slate-200 text-slate-500"
      )}>
        {num}
      </div>
      <span className={cn(
        "text-[11px] font-semibold tracking-wide",
        active ? "text-indigo-900" : (completed ? "text-emerald-700" : "text-slate-500")
      )}>
        {label}
      </span>
    </div>
  );
}

