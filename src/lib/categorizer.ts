import { type Category, type BankTransaction } from '../types';

const CHARGEBACK_EXACT = [
  'AMERICAN EXPRESS CHGBCK/ADJ TSTBLUEPAY',
  'BANKCARD-5642 BTOT ADJ 476579550413025',
  'AMERICAN EXPRESS CHGBCK/ADJ TSTNU',
  'CHARGEBACK MERCHANT BANKCD 496118245885',
  'CHGBCK/ADJ AMERICAN EXPRESS 1758778971',
  'CHGBK REV  MERCHANT BANKCD 496118245885'
];

const MERCHANT_FEES_EXACT = [
  'BANKCARD-5642 MTOT DISC 476579550413025',
  'AMERICAN EXPRESS AXP DISCNT TSTBLUEPAY',
  'BANKCARD-5642 DISCOUNT 476579550413025',
  'AMERICAN EXPRESS AXP DISCNT TSTNU',
  'Account Analysis Charge',
  'AXP DISCNT AMERICAN EXPRESS 1758778971',
  'Service Charge',
  'BANKCARD DEPOSIT MERCH FEES 000998300041244'
];

const SETTLEMENT_DEPOSITS_EXACT = [
  'BANKCARD-5642 MTOT DEP 476579550413025',
  'AMERICAN EXPRESS SETTLEMENT TSTBLUEPAY',
  'AMERICAN EXPRESS SETTLEMENT TSTNU',
  'BANKCARD DEPOSIT SETTLEMENT 000998300041244',
  'DEPOSIT    MERCHANT BANKCD 496118245885',
  'SETTLEMENT AMERICAN EXPRESS 1758778971',
  'MTOT DEP   BANKCARD',
  'MTOT DEP BANKCARD',
  'AMERICAN EXPRESS COLLECTION'
];

export function categorizeTransaction(description: string): Category {
  const desc = description.toUpperCase().trim();
  const normalizedDesc = desc.replace(/\s+/g, ' ');

  // 1. Chargebacks
  if (CHARGEBACK_EXACT.some(pattern => desc === pattern.toUpperCase() || normalizedDesc === pattern.toUpperCase().replace(/\s+/g, ' ')) || 
      desc.includes('CHGBCK') || 
      desc.includes('CHARGEBACK') ||
      desc.includes('CHGBK') ||
      desc.includes('CHGBK REV') ||
      desc.includes('BANKCARD DEPOSIT MERCH CHBK')) {
    return 'Chargebacks';
  }

  // 2. Merchant Fees
  if (MERCHANT_FEES_EXACT.some(pattern => desc === pattern.toUpperCase() || normalizedDesc === pattern.toUpperCase().replace(/\s+/g, ' ')) ||
      desc.includes('AXP') ||
      desc.includes('AXP DISCNT')) {
    return 'Merchant Fees';
  }

  // 3. Settlement Deposits
  if (SETTLEMENT_DEPOSITS_EXACT.some(pattern => desc === pattern.toUpperCase() || normalizedDesc === pattern.toUpperCase().replace(/\s+/g, ' ')) ||
      desc.includes('SETTLEMENT') ||
      desc.includes('DEPOSIT') ||
      desc.includes('AMERICAN EXPRESS COLLECTION') ||
      desc.includes('MTOT DEP') ||
      normalizedDesc.includes('MTOT DEP BANKCARD')) {
    return 'Settlement Deposit';
  }

  // 4. Bank Transfers
  if (desc.includes('TRANSFER FROM') || desc.includes('TRANSF FROM')) {
    return 'Bank Transfer';
  }

  // 5. Miscellaneous
  return 'Miscellaneous';
}

export function detectHeaders(headers: string[]): { 
  date: string; 
  description: string; 
  amount: string;
  debit: string;
  credit: string;
  accountNumber: string;
} {
  const res = { date: '', description: '', amount: '', debit: '', credit: '', accountNumber: '' };
  
  const hTrim = headers.map(h => h.toLowerCase().replace(/["']/g, '').trim());
  
  const dateAliases = ['date', 'posted date', 'transaction date', 'posted_date', 'trans date', 'effective date'];
  const descAliases = ['description', 'memo', 'details', 'transaction', 'name', 'payee', 'narrative', 'remarks'];
  const amtAliases = ['amount', 'value', 'transaction amount', 'net amount', 'sum'];
  const debitAliases = ['debit', 'withdraw', 'charge', 'amount debit', 'debit amount', 'withdrawal'];
  const creditAliases = ['credit', 'deposit', 'amount credit', 'credit amount', 'payment'];
  const accAliases = ['account', 'account number', 'account#', 'acc#', 'account_number', 'masked account'];

  const findHeader = (aliases: string[]) => {
    // 1. Exact match
    let found = headers.find(h => aliases.includes(h.toLowerCase().trim()));
    if (found) return found;

    // 2. Contains match
    found = headers.find(h => aliases.some(a => h.toLowerCase().includes(a)));
    return found || '';
  };

  res.date = findHeader(dateAliases);
  res.description = findHeader(descAliases);
  // Specifically for amount, we want to be careful if debit/credit exist
  res.amount = findHeader(amtAliases);
  res.debit = findHeader(debitAliases);
  res.credit = findHeader(creditAliases);
  res.accountNumber = findHeader(accAliases);

  // Fallbacks for date if still not found
  if (!res.date) res.date = headers.find(h => h.toLowerCase().includes('date')) || '';
  if (!res.description) res.description = headers.find(h => h.toLowerCase().includes('desc') || h.toLowerCase().includes('memo')) || '';
  
  // If amount not found but we have debit/credit, that's fine. 
  // If no debit/credit/amount, look for anything with "amt"
  if (!res.amount && !res.debit && !res.credit) {
    res.amount = headers.find(h => h.toLowerCase().includes('amt')) || '';
  }

  return res;
}
