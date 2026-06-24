"use client";
import { useCallback, useState } from "react";
import { Upload, FileSpreadsheet } from "lucide-react";

interface FileUploadProps {
  onFile: (file: File) => void;
}

export default function FileUpload({ onFile }: FileUploadProps) {
  const [dragging, setDragging] = useState(false);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragging(false);
      const file = e.dataTransfer.files[0];
      if (file && (file.name.endsWith(".xlsx") || file.name.endsWith(".xls"))) {
        onFile(file);
      }
    },
    [onFile]
  );

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) onFile(file);
  };

  return (
    <label
      className={`flex flex-col items-center justify-center gap-4 w-full h-52 rounded-xl border-2 border-dashed cursor-pointer transition-all duration-200
        ${dragging
          ? "border-violet-400 bg-violet-950/30"
          : "border-slate-700 bg-slate-900/50 hover:border-violet-500 hover:bg-slate-900"
        }`}
      onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={handleDrop}
    >
      <input
        type="file"
        accept=".xlsx,.xls"
        className="hidden"
        onChange={handleChange}
      />
      <div className={`p-4 rounded-full transition-colors ${dragging ? "bg-violet-500/20" : "bg-slate-800"}`}>
        {dragging
          ? <FileSpreadsheet className="w-8 h-8 text-violet-400" />
          : <Upload className="w-8 h-8 text-slate-400" />
        }
      </div>
      <div className="text-center">
        <p className="text-slate-300 font-medium">Drop your Excel file here</p>
        <p className="text-slate-500 text-sm mt-1">or click to browse · .xlsx and .xls supported</p>
      </div>
    </label>
  );
}
