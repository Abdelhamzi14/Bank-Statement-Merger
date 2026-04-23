import { useState, useRef, useCallback } from 'react';
import { Upload, X, FileText, CheckCircle2, AlertCircle } from 'lucide-react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { type StatementFile } from '../types';
import { detectHeaders } from '../lib/categorizer';
import { cn } from '../lib/utils';
import { motion, AnimatePresence } from 'motion/react';

interface Props {
  files: StatementFile[];
  onFilesChange: (files: StatementFile[]) => void;
}

export default function FileUploadZone({ files, onFilesChange }: Props) {
  const [isDragging, setIsDragging] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const processFile = useCallback((file: File) => {
    return new Promise<StatementFile>((resolve, reject) => {
      const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls');

      if (isExcel) {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const data = new Uint8Array(e.target?.result as ArrayBuffer);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
            
            const headers = jsonData.length > 0 ? Object.keys(jsonData[0] as object) : [];
            const mapping = detectHeaders(headers);

            resolve({
              id: crypto.randomUUID(),
              file,
              headers,
              data: jsonData,
              mappings: mapping,
              status: mapping.date && mapping.description && (mapping.amount || (mapping.debit && mapping.credit)) ? 'mapped' : 'pending'
            });
          } catch (err) {
            reject(err);
          }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
      } else {
        // CSV Parsing
        Papa.parse(file, {
          header: true,
          skipEmptyLines: 'greedy',
          transformHeader: (h) => h.trim(),
          complete: (results) => {
            const headers = results.meta.fields || [];
            const mapping = detectHeaders(headers);
            
            resolve({
              id: crypto.randomUUID(),
              file,
              headers,
              data: results.data,
              mappings: mapping,
              status: mapping.date && mapping.description && (mapping.amount || (mapping.debit && mapping.credit)) ? 'mapped' : 'pending'
            });
          },
          error: (err) => reject(err)
        });
      }
    });
  }, []);

  const handleFiles = async (newFiles: FileList | null) => {
    if (!newFiles) return;
    setError(null);
    try {
      const processed = await Promise.all(Array.from(newFiles).map(processFile)) as StatementFile[];
      onFilesChange([...files, ...processed]);
    } catch (err) {
      console.error(err);
      setError('Failed to process one or more files. Please ensure they are valid CSV or Excel statements.');
    }
  };

  const removeFile = (id: string) => {
    onFilesChange(files.filter(f => f.id !== id));
  };

  return (
    <div className="space-y-8 max-w-5xl mx-auto px-4">
      <div
        onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
        onDragLeave={() => setIsDragging(false)}
        onDrop={(e) => {
          e.preventDefault();
          setIsDragging(false);
          handleFiles(e.dataTransfer.files);
        }}
        onClick={() => fileInputRef.current?.click()}
        className={cn(
          "relative group border-2 border-dashed rounded-3xl p-24 transition-all duration-300 cursor-pointer text-center",
          isDragging 
            ? "border-indigo-500 bg-indigo-50/50" 
            : "border-slate-200 bg-white hover:border-indigo-200 hover:bg-slate-50/50 shadow-sm"
        )}
      >
        <input
          type="file"
          ref={fileInputRef}
          className="hidden"
          multiple
          accept=".csv,.xlsx,.xls"
          onChange={(e) => handleFiles(e.target.files)}
        />
        
        <div className="flex flex-col items-center">
          <div className={cn(
            "w-16 h-16 rounded-2xl flex items-center justify-center mb-6 transition-all duration-300",
            isDragging 
              ? "bg-indigo-600 text-white shadow-lg shadow-indigo-200" 
              : "bg-slate-100 text-slate-400 group-hover:bg-indigo-50 group-hover:text-indigo-600 group-hover:scale-105"
          )}>
            <Upload size={28} />
          </div>
          <h3 className="text-xl font-bold text-slate-900 mb-2">Import Statement Data</h3>
          <p className="text-sm text-slate-500 max-w-sm mx-auto leading-relaxed">
            Drag and drop CSV or Excel files here, or click to browse your computer for local statement ingestion.
          </p>
        </div>
      </div>

      <AnimatePresence>
        {error && (
          <motion.div
            initial={{ opacity: 0, height: 0 }}
            animate={{ opacity: 1, height: 'auto' }}
            exit={{ opacity: 0, height: 0 }}
            className="bg-red-50 border border-red-100 text-red-600 p-5 rounded-2xl flex items-center gap-4 text-sm font-medium shadow-sm"
          >
            <div className="w-10 h-10 rounded-full bg-red-100 flex items-center justify-center shrink-0">
              <AlertCircle size={20} />
            </div>
            {error}
          </motion.div>
        )}

        {files.length > 0 && (
          <motion.div
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6 pt-4"
          >
            {files.map((file) => (
              <motion.div
                key={file.id}
                layout
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                exit={{ opacity: 0, scale: 0.95 }}
                className="bg-white border border-slate-200/60 rounded-2xl p-5 flex items-center justify-between group transition-all hover:border-indigo-200 hover:shadow-md"
              >
                <div className="flex items-center space-x-4 overflow-hidden">
                  <div className="w-10 h-10 rounded-xl bg-slate-50 flex items-center justify-center text-slate-600 flex-shrink-0 group-hover:bg-indigo-50 group-hover:text-indigo-600 transition-colors">
                    <FileText size={20} />
                  </div>
                  <div className="overflow-hidden text-left">
                    <p className="text-sm font-bold text-slate-900 truncate tracking-tight">{file.file.name}</p>
                    <p className="text-[11px] text-slate-500 font-medium">{file.data.length} transactions · {file.status}</p>
                  </div>
                </div>
                
                <div className="flex items-center space-x-3">
                  {file.status === 'mapped' && (
                    <CheckCircle2 size={18} className="text-emerald-500" />
                  )}
                  <button
                    onClick={(e) => { e.stopPropagation(); removeFile(file.id); }}
                    className="w-8 h-8 flex items-center justify-center rounded-lg text-slate-300 hover:text-red-500 hover:bg-red-50 transition-all opacity-0 group-hover:opacity-100"
                  >
                    <X size={16} />
                  </button>
                </div>
              </motion.div>
            ))}
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}
