import { useState, useMemo } from 'react';
import { type BankTransaction, type Category } from '../types';
import { buildExternalIdMap, getExternalId } from '../lib/externalId';
import { motion } from 'motion/react';
import { CreditCard, AlertCircle, Save, Filter, ChevronDown, Hash } from 'lucide-react';
import { cn } from '../lib/utils';

interface JournalEntryStepProps {
  transactions: BankTransaction[];
  onUpdateTransaction: (id: string, updates: Partial<BankTransaction>) => void;
  onBulkUpdateTransactions: (ids: string[], updates: Partial<BankTransaction>) => void;
  externalIdStartSequence: number;
  onUpdateExternalIdStartSequence: (seq: number) => void;
}

const JOURNAL_CATEGORIES: Category[] = ['Chargebacks', 'Merchant Fees', 'Bank Transfer', 'Miscellaneous'];

export default function JournalEntryStep({ 
  transactions, 
  onUpdateTransaction, 
  onBulkUpdateTransactions,
  externalIdStartSequence = 1,
  onUpdateExternalIdStartSequence
}: JournalEntryStepProps) {
  const [search, setSearch] = useState('');
  const [filter, setFilter] = useState<Category | 'All'>('All');
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [bulkDebitValue, setBulkDebitValue] = useState('');

  // Map of external IDs based on category, date, and selected starting sequence
  const extIdMap = useMemo(() => {
    return buildExternalIdMap(transactions, externalIdStartSequence);
  }, [transactions, externalIdStartSequence]);

  const samplePadded = useMemo(() => {
    const num = Math.max(1, Math.floor(Number(externalIdStartSequence) || 1));
    const padLength = Math.max(4, String(num).length);
    return String(num).padStart(padLength, '0');
  }, [externalIdStartSequence]);

  // Filter for categories that need NetSuite mapping review
  const reviewableTransactions = useMemo(() => {
    return transactions.filter(t => {
      const matchesSearch = t.description.toLowerCase().includes(search.toLowerCase()) || 
                           t.accountNumber.toLowerCase().includes(search.toLowerCase());
      
      const isRestrictedCategory = JOURNAL_CATEGORIES.includes(t.category);
      const matchesFilter = filter === 'All' ? isRestrictedCategory : t.category === filter;
      
      return matchesSearch && matchesFilter;
    });
  }, [transactions, search, filter]);

  const toggleSelectAll = () => {
    if (selectedIds.size === reviewableTransactions.length) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(reviewableTransactions.map(t => t.id)));
    }
  };

  const toggleSelectOne = (id: string) => {
    const next = new Set(selectedIds);
    if (next.has(id)) next.delete(id);
    else next.add(id);
    setSelectedIds(next);
  };

  const handleBulkApply = () => {
    if (selectedIds.size === 0 || !bulkDebitValue.trim()) return;
    onBulkUpdateTransactions(Array.from(selectedIds), { netsuiteDebitAccount: bulkDebitValue });
    setBulkDebitValue('');
    setSelectedIds(new Set());
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row md:items-end justify-between gap-4">
        <div>
          <h3 className="text-xl font-bold text-slate-800">Journal Entry Review</h3>
          <p className="text-sm text-slate-500">Verify NetSuite account mappings for your transactions.</p>
          
          {selectedIds.size > 0 && (
            <motion.div 
              initial={{ opacity: 0, y: 5 }}
              animate={{ opacity: 1, y: 0 }}
              className="mt-4 flex items-center gap-2 p-2 bg-indigo-50 border border-indigo-100 rounded-xl"
            >
              <div className="flex items-center gap-2 px-3 py-1.5 border-r border-indigo-200">
                <span className="text-[10px] font-bold text-indigo-700 uppercase">{selectedIds.size} Selected</span>
              </div>
              <div className="flex items-center gap-2 px-2">
                <input
                  type="text"
                  placeholder="Bulk Debit Acc#"
                  value={bulkDebitValue}
                  onChange={(e) => setBulkDebitValue(e.target.value)}
                  className="px-3 py-1.5 rounded-lg border border-indigo-200 bg-white text-xs font-mono w-40 focus:outline-none focus:ring-2 focus:ring-indigo-500/20"
                />
                <button
                  onClick={handleBulkApply}
                  className="px-3 py-1.5 bg-indigo-600 text-white rounded-lg text-xs font-bold hover:bg-indigo-700 transition-colors shadow-sm shadow-indigo-100"
                >
                  Apply to Selected
                </button>
              </div>
            </motion.div>
          )}
        </div>
        
        <div className="flex flex-wrap items-center gap-3">
          <input
            type="text"
            placeholder="Search..."
            className="px-4 py-2 rounded-xl border border-slate-200 bg-white text-slate-900 placeholder:text-slate-400 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-all text-xs shadow-sm min-w-[200px]"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />

          <div className="relative">
            <Filter className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
            <select
              className="pl-9 pr-10 py-2 rounded-xl border border-slate-200 bg-white text-slate-700 appearance-none focus:outline-none focus:ring-2 focus:ring-indigo-500/20 text-xs font-semibold transition-all cursor-pointer shadow-sm min-w-[180px]"
              value={filter}
              onChange={(e) => setFilter(e.target.value as Category | 'All')}
            >
              <option value="All">All Restricted Categories</option>
              {JOURNAL_CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
            <div className="absolute right-3 top-1/2 -translate-y-1/2 pointer-events-none text-slate-400">
              <ChevronDown size={14} />
            </div>
          </div>
        </div>
      </div>

      {/* External ID Sequence Configuration Card */}
      <div className="bg-white border border-slate-200 rounded-2xl p-4 shadow-xs">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="flex items-start sm:items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-indigo-50 border border-indigo-100 flex items-center justify-center shrink-0 text-indigo-600">
              <Hash size={20} />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-sm font-bold text-slate-900">External ID Sequence Start</span>
                <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-indigo-50 text-indigo-700 border border-indigo-100">
                  JE Upload Tab
                </span>
              </div>
              <p className="text-xs text-slate-500 mt-0.5">
                Select or type the starting number for the NetSuite External ID sequence in the exported Excel file.
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-3">
            <div className="flex items-center gap-2 bg-slate-50 border border-slate-200 px-3 py-1.5 rounded-xl">
              <label htmlFor="extIdStartInput" className="text-xs font-bold text-slate-700 whitespace-nowrap">
                Start Sequence #:
              </label>
              <input
                id="extIdStartInput"
                type="number"
                min={1}
                max={999999}
                value={externalIdStartSequence || ''}
                onChange={(e) => {
                  const val = parseInt(e.target.value, 10);
                  onUpdateExternalIdStartSequence(isNaN(val) || val < 1 ? 1 : val);
                }}
                placeholder="1"
                className="w-24 px-2.5 py-1 bg-white border border-slate-300 rounded-lg text-xs font-mono font-bold text-indigo-700 text-center focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500"
              />
            </div>

            {/* Quick sequence presets */}
            <div className="flex items-center gap-1">
              <span className="text-[10px] font-bold text-slate-400 uppercase mr-1">Presets:</span>
              {[1, 1001, 5001, 10001].map((seq) => (
                <button
                  key={seq}
                  type="button"
                  onClick={() => onUpdateExternalIdStartSequence(seq)}
                  className={cn(
                    "px-2.5 py-1.5 rounded-lg text-xs font-mono font-bold transition-all",
                    externalIdStartSequence === seq
                      ? "bg-indigo-600 text-white shadow-xs"
                      : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                  )}
                >
                  {seq === 1 ? '0001' : seq}
                </button>
              ))}
            </div>

            {/* Live Preview badge */}
            <div className="flex items-center gap-2 px-3 py-1.5 bg-indigo-50/70 border border-indigo-100 rounded-xl text-xs">
              <span className="text-[10px] font-bold uppercase text-indigo-600">Sample:</span>
              <span className="font-mono font-bold text-indigo-800">
                Chargebacks {samplePadded}
              </span>
            </div>
          </div>
        </div>
      </div>

      <div className="modern-panel">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-200">
                <th className="px-6 py-4 w-12">
                  <input 
                    type="checkbox" 
                    className="w-4 h-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer"
                    checked={reviewableTransactions.length > 0 && selectedIds.size === reviewableTransactions.length}
                    onChange={toggleSelectAll}
                  />
                </th>
                <th className="px-6 py-4 label-micro w-32">Date</th>
                <th className="px-6 py-4 label-micro flex-1">Description</th>
                <th className="px-6 py-4 label-micro w-28 text-right">Amount</th>
                <th className="px-6 py-4 label-micro w-40 text-center">Category</th>
                <th className="px-4 py-4 label-micro w-44 text-center text-indigo-700 font-bold">External ID</th>
                <th className="px-6 py-4 label-micro w-48">NetSuite Debit</th>
                <th className="px-6 py-4 label-micro w-48 text-rose-600 font-bold">NetSuite Credit</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {reviewableTransactions.length === 0 ? (
                <tr>
                  <td colSpan={8} className="px-6 py-12 text-center text-slate-400">
                    <div className="flex flex-col items-center gap-3">
                      <CreditCard size={32} className="opacity-20" />
                      <p className="text-sm font-medium">No transactions requiring manual review found.</p>
                    </div>
                  </td>
                </tr>
              ) : (
                reviewableTransactions.map((t) => (
                  <tr key={t.id} className={cn(
                    "hover:bg-slate-50/50 transition-colors",
                    selectedIds.has(t.id) && "bg-indigo-50/30 hover:bg-indigo-50/50"
                  )}>
                    <td className="px-6 py-4">
                      <input 
                        type="checkbox" 
                        className="w-4 h-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer"
                        checked={selectedIds.has(t.id)}
                        onChange={() => toggleSelectOne(t.id)}
                      />
                    </td>
                    <td className="px-6 py-4 text-xs font-medium text-slate-600">{t.date}</td>
                    <td className="px-6 py-4">
                      <div className="text-xs font-bold text-slate-800 line-clamp-1">{t.description}</div>
                      <div className="text-[10px] text-slate-400 mt-0.5">Account: {t.accountNumber || 'N/A'}</div>
                    </td>
                    <td className={cn(
                      "px-6 py-4 text-xs font-bold text-right tabular-nums",
                      t.amount < 0 ? "text-rose-600" : "text-emerald-600"
                    )}>
                      {t.amount.toLocaleString('en-US', { style: 'currency', currency: 'USD' })}
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex justify-center">
                        <span className={cn(
                          "px-2 py-0.5 rounded-full text-[10px] font-bold uppercase tracking-wider",
                          t.category === 'Chargebacks' ? "bg-rose-50 text-rose-700" :
                          t.category === 'Merchant Fees' ? "bg-amber-50 text-amber-700" :
                          t.category === 'Bank Transfer' ? "bg-blue-50 text-blue-700" :
                          "bg-slate-50 text-slate-600"
                        )}>
                          {t.category}
                        </span>
                      </div>
                    </td>
                    <td className="px-4 py-4 text-center">
                      <span className="inline-flex items-center px-2.5 py-1 rounded-md text-[11px] font-mono font-bold bg-indigo-50 text-indigo-700 border border-indigo-200/80 shadow-2xs">
                        {getExternalId(t, extIdMap, externalIdStartSequence)}
                      </span>
                    </td>
                    <td className="px-6 py-4">
                      <input 
                        type="text"
                        value={t.netsuiteDebitAccount || ''}
                        onChange={(e) => onUpdateTransaction(t.id, { netsuiteDebitAccount: e.target.value })}
                        className="w-full px-3 py-1.5 bg-white border border-slate-200 rounded-md text-xs font-mono focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-all font-bold"
                        placeholder="Debit Acc#"
                      />
                    </td>
                    <td className="px-6 py-4">
                      <input 
                        type="text"
                        value={t.netsuiteCreditAccount || ''}
                        onChange={(e) => onUpdateTransaction(t.id, { netsuiteCreditAccount: e.target.value })}
                        className="w-full px-3 py-1.5 bg-rose-50/50 border border-rose-200 rounded-md text-xs font-mono text-rose-700 font-bold focus:outline-none focus:ring-2 focus:ring-rose-500/20 focus:border-rose-500 transition-all placeholder:text-rose-300"
                        placeholder="Credit Acc#"
                      />
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
