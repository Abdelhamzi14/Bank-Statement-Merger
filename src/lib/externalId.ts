import { type BankTransaction, type Category } from '../types';

export const JE_CATEGORIES: Category[] = ['Chargebacks', 'Merchant Fees', 'Bank Transfer'];

/**
 * Computes the External ID mapping for transactions grouped by Category and Date.
 * Numbering sequence begins at startSequence (default 1).
 */
export function buildExternalIdMap(
  transactions: BankTransaction[],
  startSequence: number = 1
): Map<string, string> {
  const jeItems = transactions.filter(t => JE_CATEGORIES.includes(t.category));

  // Sort consistently: category order, then date
  const sorted = [...jeItems].sort((a, b) => {
    const catDiff = JE_CATEGORIES.indexOf(a.category) - JE_CATEGORIES.indexOf(b.category);
    if (catDiff !== 0) return catDiff;
    return (a.date || '').localeCompare(b.date || '');
  });

  const extIdMap = new Map<string, string>();
  const categoryCounters = new Map<string, number>();

  const startNum = !isNaN(Number(startSequence))
    ? Math.max(1, Math.floor(Number(startSequence)))
    : 1;
  const padLength = Math.max(4, String(startNum).length);

  sorted.forEach(item => {
    const key = `${item.category}_${item.date}`;
    if (!extIdMap.has(key)) {
      const currentCount = categoryCounters.has(item.category)
        ? categoryCounters.get(item.category)! + 1
        : startNum;
      categoryCounters.set(item.category, currentCount);
      const padded = String(currentCount).padStart(padLength, '0');
      extIdMap.set(key, `${item.category} ${padded}`);
    }
  });

  return extIdMap;
}

/**
 * Retrieves the formatted External ID for a specific transaction.
 */
export function getExternalId(
  item: BankTransaction,
  extIdMap: Map<string, string>,
  startSequence: number = 1
): string {
  const key = `${item.category}_${item.date}`;
  if (extIdMap.has(key)) {
    return extIdMap.get(key)!;
  }
  const startNum = !isNaN(Number(startSequence))
    ? Math.max(1, Math.floor(Number(startSequence)))
    : 1;
  const padLength = Math.max(4, String(startNum).length);
  return `${item.category} ${String(startNum).padStart(padLength, '0')}`;
}
