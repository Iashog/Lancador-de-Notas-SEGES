import React, { useState, useCallback } from 'react';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { FileSpreadsheet, ClipboardCopy, CheckCircle2, AlertCircle, Info } from "lucide-react";
import * as XLSX from 'xlsx';
import { showSuccess, showError } from "@/utils/toast";
import { MadeWithDyad } from "@/components/made-with-dyad";

interface StudentData {
  [key: string]: any;
}

interface ActivityMapping {
  label: string;
  column: string;
}

const Index = () => {
  const [data, setData] = useState<StudentData[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  const [studentNameCol, setStudentNameCol] = useState<string>('');
  const [activities, setActivities] = useState<ActivityMapping[]>([
    { label: 'Atividade 1 (Av)', column: '' },
    { label: 'Atividade 1 (Rec)', column: '' },
    { label: 'Atividade 2 (Av)', column: '' },
    { label: 'Atividade 2 (Rec)', column: '' },
    { label: 'Atividade 3 (Av)', column: '' },
    { label: 'Atividade 3 (Rec)', column: '' },
  ]);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const jsonData = XLSX.utils.sheet_to_json(ws) as any[];
        
        if (jsonData.length > 0) {
          const cols = Object.keys(jsonData[0]);
          setColumns(cols);
          setData(jsonData);
          showSuccess("Planilha carregada com sucesso!");
        }
      } catch (err) {
        showError("Erro ao ler o arquivo Excel.");
      }
    };
    reader.readAsBinaryString(file);
  };

  const updateActivityMapping = (index: number, column: string) => {
    setActivities(prev => {
      const next = [...prev];
      next[index] = { ...next[index], column };
      return next;
    });
  };

  const generateScript = useCallback(() => {
    if (!studentNameCol) {
      showError("Selecione a coluna com o nome dos alunos.");
      return;
    }

    const scriptData = data.map(row => {
      const name = String(row[studentNameCol] || '').trim().toUpperCase();
      const grades = activities.map(act => row[act.column] !== undefined ? row[act.column] : null);
      return { name, grades };
    });

    const script = `
(function() {
  const studentsData = ${JSON.stringify(scriptData)};
  console.log('Iniciando lançamento de notas...');
  let count = 0;

  studentsData.forEach(student => {
    const row = Array.from(document.querySelectorAll('tr')).find(tr => 
      tr.innerText.toUpperCase().includes(student.name)
    );

    if (row) {
      const inputs = Array.from(row.querySelectorAll('input[type="text"]'));
      student.grades.forEach((grade, idx) => {
        if (grade !== null && inputs[idx]) {
          inputs[idx].value = grade.toString().replace('.', ',');
          ['input', 'change', 'blur'].forEach(type => 
            inputs[idx].dispatchEvent(new Event(type, { bubbles: true }))
          );
        }
      });
      count++;
      row.style.backgroundColor = '#e6fffa';
    }
  });

  alert('Processamento concluído! ' + count + ' alunos atualizados.');
})();
    `;

    navigator.clipboard.writeText(script);
    showSuccess("Script copiado!");
  }, [data, studentNameCol, activities]);

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8">
      <div className="max-w-6xl mx-auto space-y-8">
        <header className="text-center space-y-4">
          <div className="inline-flex items-center justify-center p-4 bg-indigo-600 rounded-3xl shadow-xl shadow-indigo-100">
            <FileSpreadsheet className="text-white w-10 h-10" />
          </div>
          <div className="space-y-2">
            <h1 className="text-4xl font-extrabold text-slate-900 tracking-tight">Lançador de Notas SEGES</h1>
            <p className="text-slate-500 text-lg">Automatize o preenchimento de diários eletrônicos</p>
          </div>
        </header>

        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
          <div className="lg:col-span-4 space-y-6">
            <Card className="border-none shadow-sm">
              <CardHeader>
                <CardTitle className="text-xl flex items-center gap-2">
                  <Info className="w-5 h-5 text-indigo-500" />
                  1. Importar Excel
                </CardTitle>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="space-y-2">
                  <Label htmlFor="file">Arquivo da Planilha</Label>
                  <Input id="file" type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
                </div>

                {columns.length > 0 && (
                  <div className="space-y-2 pt-2">
                    <Label>Coluna do Nome</Label>
                    <Select value={studentNameCol} onValueChange={setStudentNameCol}>
                      <SelectTrigger>
                        <SelectValue placeholder="Selecione..." />
                      </SelectTrigger>
                      <SelectContent>
                        {columns.map(c => <SelectItem key={c} value={c}>{c}</SelectItem>)}
                      </SelectContent>
                    </Select>
                  </div>
                )}
              </CardContent>
            </Card>

            {columns.length > 0 && (
              <Card className="border-none shadow-sm">
                <CardHeader>
                  <CardTitle className="text-xl">2. Mapeamento</CardTitle>
                  <CardDescription>Vincule as colunas às atividades</CardDescription>
                </CardHeader>
                <CardContent className="space-y-4">
                  {activities.map((act, i) => (
                    <div key={i} className="space-y-1">
                      <Label className="text-xs text-slate-500 uppercase font-bold">{act.label}</Label>
                      <Select value={act.column} onValueChange={(v) => updateActivityMapping(i, v)}>
                        <SelectTrigger className="h-9">
                          <SelectValue placeholder="Ignorar" />
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value="">(Ignorar)</SelectItem>
                          {columns.map(c => <SelectItem key={c} value={c}>{c}</SelectItem>)}
                        </SelectContent>
                      </Select>
                    </div>
                  ))}
                  <Button onClick={generateScript} className="w-full bg-indigo-600 hover:bg-indigo-700 mt-4 py-6 text-lg rounded-2xl shadow-lg shadow-indigo-100">
                    <ClipboardCopy className="w-5 h-5 mr-2" />
                    Copiar Script
                  </Button>
                </CardContent>
              </Card>
            )}
          </div>

          <div className="lg:col-span-8">
            <Card className="border-none shadow-sm min-h-[600px]">
              <CardHeader className="border-b border-slate-100">
                <CardTitle>Visualização dos Dados</CardTitle>
              </CardHeader>
              <CardContent className="p-0">
                {data.length > 0 ? (
                  <div className="overflow-auto max-h-[700px]">
                    <Table>
                      <TableHeader className="bg-slate-50 sticky top-0">
                        <TableRow>
                          <TableHead className="font-bold">Aluno</TableHead>
                          {activities.filter(a => a.column).map((a, i) => (
                            <TableHead key={i} className="text-center font-bold">{a.label}</TableHead>
                          ))}
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {data.slice(0, 20).map((row, i) => (
                          <TableRow key={i}>
                            <TableCell className="font-medium">{row[studentNameCol] || '-'}</TableCell>
                            {activities.filter(a => a.column).map((a, j) => (
                              <TableCell key={j} className="text-center">{row[a.column] ?? '-'}</TableCell>
                            ))}
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </div>
                ) : (
                  <div className="flex flex-col items-center justify-center h-[500px] text-slate-400">
                    <FileSpreadsheet className="w-16 h-16 mb-4 opacity-10" />
                    <p>Aguardando upload da planilha...</p>
                  </div>
                )}
              </CardContent>
            </Card>
          </div>
        </div>
        <MadeWithDyad />
      </div>
    </div>
  );
};

export default Index;