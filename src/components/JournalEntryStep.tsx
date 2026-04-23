import { useState, useMemo } from 'react';
import { type BankTransaction, type Category } from '../types';
import { motion } from 'motion/react';
import { CreditCard, AlertCircle, Save, Filter, ChevronDown } from 'lucide-react';
import { cn } from '../lib/utils';

interface JournalEntryStepProps {
  transactions: BankTransaction[];
  onUpdateTransaction: (id: string, updates: Partial<BankTransaction>) => void;
}

const JOURNAL_CATEGORIES: Category[] = ['Chargebacks', 'Merchant Fees', 'Bank Transfer', 'Miscellaneous'];

export default function JournalEntryStep({ transactions, onUpdateTransaction }: JournalEntryStepProps) {
  const [search, setSearch] = useState('');
  const [filter, setFilter] = useState<Category | 'All'>('All');

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

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h3 className="text-xl font-bold text-slate-800">Journal Entry Review</h3>
          <p className="text-sm text-slate-500">Verify NetSuite account mappings for your transactions.</p>
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
          <div className="px-3 py-1 bg-amber-50 border border-amber-100 rounded-md flex items-center gap-2">
            <AlertCircle size={14} className="text-amber-500" />
            <span className="text-[10px] font-bold text-amber-700 uppercase tracking-wider">Review Required</span>
          </div>
        </div>
      </div>

      <div className="modern-panel">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-200">
                <th className="px-6 py-4 label-micro w-32">Date</th>
                <th className="px-6 py-4 label-micro flex-1">Description</th>
                <th className="px-6 py-4 label-micro w-28 text-right">Amount</th>
                <th className="px-6 py-4 label-micro w-40 text-center">Category</th>
                <th className="px-6 py-4 label-micro w-48">NetSuite Debit</th>
                <th className="px-6 py-4 label-micro w-48 text-rose-600 font-bold">NetSuite Credit</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {reviewableTransactions.length === 0 ? (
                <tr>
                  <td colSpan={6} className="px-6 py-12 text-center text-slate-400">
                    <div className="flex flex-col items-center gap-3">
                      <CreditCard size={32} className="opacity-20" />
                      <p className="text-sm font-medium">No transactions requiring manual review found.</p>
                    </div>
                  </td>
                </tr>
              ) : (
                reviewableTransactions.map((t) => (
                  <tr key={t.id} className="hover:bg-slate-50/50 transition-colors">
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
                    <td className="px-6 py-4">
                      <input 
                        type="text"
                        value={t.netsuiteDebitAccount || ''}
                        onChange={(e) => onUpdateTransaction(t.id, { netsuiteDebitAccount: e.target.value })}
                        className="w-full px-3 py-1.5 bg-white border border-slate-200 rounded-md text-xs font-mono focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-all"
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
