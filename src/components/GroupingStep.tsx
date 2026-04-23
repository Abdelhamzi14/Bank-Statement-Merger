import { useState, useMemo } from 'react';
import { type BankTransaction, type Category } from '../types';
import { cn } from '../lib/utils';
import { motion, AnimatePresence } from 'motion/react';
import { Search, Filter, AlertCircle, LayoutGrid, List, ChevronDown, CheckSquare, Square, Trash2, CheckCircle2, X } from 'lucide-react';
import { TableVirtuoso } from 'react-virtuoso';

interface Props {
  transactions: BankTransaction[];
  onUpdateCategory: (id: string, category: Category) => void;
  onBulkUpdateCategory: (ids: string[], category: Category) => void;
}

const CATEGORIES: Category[] = [
  'Chargebacks',
  'Merchant Fees',
  'Settlement Deposit',
  'Bank Transfer',
  'Miscellaneous'
];

export default function GroupingStep({ transactions, onUpdateCategory, onBulkUpdateCategory }: Props) {
  const [search, setSearch] = useState('');
  const [filter, setFilter] = useState<Category | 'All'>('All');
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [bulkCategory, setBulkCategory] = useState<Category>(CATEGORIES[0]);

  const filtered = useMemo(() => {
    return transactions.filter(t => {
      const matchesSearch = t.description.toLowerCase().includes(search.toLowerCase()) || 
                           t.sourceFile.toLowerCase().includes(search.toLowerCase());
      const matchesFilter = filter === 'All' || t.category === filter;
      return matchesSearch && matchesFilter;
    });
  }, [transactions, search, filter]);

  const toggleSelection = (id: string) => {
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const toggleSelectAll = () => {
    if (selectedIds.size === filtered.length && filtered.length > 0) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(filtered.map(t => t.id)));
    }
  };

  const handleBulkApply = () => {
    onBulkUpdateCategory(Array.from(selectedIds), bulkCategory);
    setSelectedIds(new Set());
  };

  const isAllSelected = filtered.length > 0 && selectedIds.size === filtered.length;
  const isSomeSelected = selectedIds.size > 0 && selectedIds.size < filtered.length;

  return (
    <div className="space-y-6 flex flex-col min-h-[900px]">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 shrink-0 px-2 pt-2">
        <div className="flex items-center gap-4 flex-1">
          <div className="relative flex-1 max-w-md">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
            <input
              type="text"
              placeholder="Search descriptions..."
              className="w-full pl-10 pr-4 py-2.5 rounded-xl border border-slate-200 bg-white text-slate-900 placeholder:text-slate-400 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-all text-sm shadow-sm"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
          </div>
          
          {selectedIds.size > 0 && (
            <motion.div 
              initial={{ opacity: 0, x: -10 }}
              animate={{ opacity: 1, x: 0 }}
              className="flex items-center gap-2 bg-slate-900 text-white px-4 py-2 rounded-xl shadow-lg"
            >
              <span className="text-[10px] font-bold uppercase tracking-widest text-slate-400">
                {selectedIds.size} Selected
              </span>
              <div className="h-4 w-px bg-slate-700 mx-2" />
              <select
                value={bulkCategory}
                onChange={(e) => setBulkCategory(e.target.value as Category)}
                className="bg-transparent border-none text-[10px] font-bold uppercase tracking-wider text-white focus:outline-none cursor-pointer"
              >
                {CATEGORIES.map(c => <option key={c} value={c} className="bg-slate-900">{c}</option>)}
              </select>
              <button 
                onClick={handleBulkApply}
                className="ml-2 bg-indigo-600 hover:bg-indigo-500 text-white text-[10px] font-bold uppercase tracking-widest px-3 py-1 rounded-lg transition-colors flex items-center gap-1.5"
              >
                <CheckCircle2 size={12} />
                Apply
              </button>
              <button 
                onClick={() => setSelectedIds(new Set())}
                className="ml-1 hover:text-rose-400 transition-colors"
                title="Cancel selection"
              >
                <X size={14} />
              </button>
            </motion.div>
          )}
        </div>

        <div className="flex items-center gap-3">
          <div className="relative">
            <Filter className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
            <select
              className="pl-9 pr-10 py-2.5 rounded-xl border border-slate-200 bg-white text-slate-700 appearance-none focus:outline-none focus:ring-2 focus:ring-indigo-500/20 text-xs font-semibold transition-all cursor-pointer shadow-sm"
              value={filter}
              onChange={(e) => setFilter(e.target.value as Category | 'All')}
            >
              <option value="All">All Categories</option>
              {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
            <div className="absolute right-3 top-1/2 -translate-y-1/2 pointer-events-none text-slate-400">
              <ChevronDown size={14} />
            </div>
          </div>
        </div>
      </div>

      <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm flex-1 flex flex-col min-h-[750px]">
        {filtered.length > 0 ? (
          <div className="flex-1 overflow-hidden relative">
            <TableVirtuoso
              data={filtered}
              useWindowScroll={false}
              style={{ height: '800px', width: '100%' }}
              fixedHeaderContent={() => (
                <tr className="bg-slate-50 border-b border-slate-200 text-slate-500 text-[11px] font-bold uppercase tracking-wider text-left">
                  <th className="px-6 py-4 w-12 text-center">
                    <button 
                      onClick={toggleSelectAll}
                      className={cn(
                        "transition-colors",
                        isAllSelected || isSomeSelected ? "text-indigo-600" : "text-slate-300 hover:text-slate-400"
                      )}
                    >
                      {isAllSelected ? <CheckSquare size={18} /> : isSomeSelected ? <CheckSquare size={18} className="opacity-60" /> : <Square size={18} />}
                    </button>
                  </th>
                  <th className="px-6 py-4 w-32 whitespace-nowrap">Date</th>
                  <th className="px-6 py-4 w-24">Account</th>
                  <th className="px-6 py-4">Transaction Details</th>
                  <th className="px-6 py-4 text-right w-24">Debit</th>
                  <th className="px-6 py-4 text-right w-24">Credit</th>
                  <th className="px-6 py-4 text-right w-32">Amount</th>
                  <th className="px-6 py-4 w-44">Categorization</th>
                  <th className="px-6 py-4 whitespace-nowrap w-40 text-right pr-10">File</th>
                </tr>
              )}
              itemContent={(index, t) => (
                <>
                  <td className="px-6 py-4 text-center border-b border-slate-100">
                    <button 
                      onClick={() => toggleSelection(t.id)}
                      className={cn(
                        "transition-colors",
                        selectedIds.has(t.id) ? "text-indigo-600" : "text-slate-300 hover:text-slate-400"
                      )}
                    >
                      {selectedIds.has(t.id) ? <CheckSquare size={18} /> : <Square size={18} />}
                    </button>
                  </td>
                  <td className="px-6 py-4 text-slate-500 font-medium text-xs border-b border-slate-100">{t.date}</td>
                  <td className="px-6 py-4 text-slate-400 text-[11px] font-mono border-b border-slate-100">{t.accountNumber || '—'}</td>
                  <td className="px-6 py-4 border-b border-slate-100">
                    <p className="text-slate-900 font-semibold text-sm max-w-xs md:max-w-md truncate" title={t.description}>
                      {t.description}
                    </p>
                  </td>
                  <td className="px-6 py-4 font-mono text-xs font-semibold text-right text-rose-500 border-b border-slate-100">
                    {t.debit ? t.debit.toLocaleString(undefined, { minimumFractionDigits: 2 }) : '—'}
                  </td>
                  <td className="px-6 py-4 font-mono text-xs font-semibold text-right text-emerald-600 border-b border-slate-100">
                    {t.credit ? t.credit.toLocaleString(undefined, { minimumFractionDigits: 2 }) : '—'}
                  </td>
                  <td className={cn(
                    "px-6 py-4 font-mono text-xs font-bold text-right whitespace-nowrap border-b border-slate-100",
                    t.amount >= 0 ? "text-emerald-700" : "text-rose-700"
                  )}>
                    {t.amount >= 0 ? '' : ''}{t.amount.toLocaleString(undefined, { minimumFractionDigits: 2 })}
                  </td>
                  <td className="px-6 py-4 border-b border-slate-100">
                    <select
                      value={t.category}
                      onChange={(e) => onUpdateCategory(t.id, e.target.value as Category)}
                      className={cn(
                        "px-3 py-1.5 rounded-lg font-bold text-[10px] uppercase border transition-all cursor-pointer focus:outline-none focus:ring-2 focus:ring-opacity-20",
                        t.category === 'Chargebacks' ? 'bg-rose-50 border-rose-100 text-rose-700 focus:ring-rose-500' :
                        t.category === 'Merchant Fees' ? 'bg-amber-50 border-amber-100 text-amber-700 focus:ring-amber-500' :
                        t.category === 'Settlement Deposit' ? 'bg-emerald-50 border-emerald-100 text-emerald-700 focus:ring-emerald-500' :
                        t.category === 'Bank Transfer' ? 'bg-blue-50 border-blue-100 text-blue-700 focus:ring-blue-500' : 
                        'bg-slate-50 border-slate-100 text-slate-600 focus:ring-slate-500'
                      )}
                    >
                      {CATEGORIES.map(c => (
                        <option key={c} value={c} className="text-slate-900 bg-white font-sans normal-case text-xs">{c}</option>
                      ))}
                    </select>
                  </td>
                  <td className="px-6 py-4 text-right pr-10 border-b border-slate-100">
                    <span className="text-[10px] font-semibold text-slate-400 bg-slate-50 px-2 py-1 rounded-md inline-block max-w-[120px] ml-auto truncate uppercase tracking-tight" title={t.sourceFile}>
                      {fileDisplayName(t.sourceFile)}
                    </span>
                  </td>
                </>
              )}
            />
          </div>
        ) : (
          <div className="p-12 text-center h-full flex flex-col items-center justify-center">
            <AlertCircle className="mx-auto text-slate-300 mb-4" size={48} />
            <h3 className="text-lg font-medium text-slate-900 mb-1">No transactions found</h3>
            <p className="text-slate-500 text-sm">Try adjusting your filters or search terms.</p>
          </div>
        )}
      </div>
    </div>
  );
}

function fileDisplayName(name: string) {
  if (name.length > 15) return '...' + name.slice(-12);
  return name;
}
