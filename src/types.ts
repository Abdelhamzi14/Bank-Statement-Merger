export type Category = 
  | 'Chargebacks' 
  | 'Merchant Fees' 
  | 'Settlement Deposit' 
  | 'Bank Transfer' 
  | 'Miscellaneous';

export interface BankTransaction {
  id: string;
  sourceFile: string;
  date: string;
  description: string;
  amount: number;
  debit: number | null;
  credit: number | null;
  accountNumber: string;
  category: Category;
  originalCategory: Category;
  netsuiteDebitAccount?: string;
  netsuiteCreditAccount?: string;
  raw: any;
}

export interface StatementFile {
  id: string;
  file: File;
  headers: string[];
  data: any[];
  mappings: {
    date: string;
    description: string;
    amount: string;
    debit: string;
    credit: string;
    accountNumber: string;
  };
  status: 'pending' | 'parsing' | 'mapped' | 'error';
}

export interface ProcessingState {
  step: 1 | 2 | 3 | 4;
  files: StatementFile[];
  transactions: BankTransaction[];
}
