"use client";

import React, { useState } from 'react';
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
  name: string;
  [key: string]: any;
}

const Index = () => {
  const [data, setData] = useState<StudentData[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  const [mapping, setMapping] = useState({
    studentName: '',
    activities: [
      { label: 'Atividade 1 (Av)', column: '' },
      { label: 'Atividade 1 (Rec)', column: '' },
      { label: 'Atividade 2 (Av)', column: '' },
      { label: 'Atividade 2 (Rec)', column: '' },
      { label: 'Atividade 3 (Av)', column: '' },
      { label: 'Atividade 3 (Rec)', column: '' },
    ]
  });

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
    const newActivities = [...mapping.activities];
    newActivities[index].column = column;
    setMapping({ ...mapping, activities: newActivities });
  };

  const generateScript = () => {
    if (!mapping.studentName) {
      showError("Selecione a coluna com o nome dos alunos.");
      return;
    }

    const scriptData = data.map(row => {
      const studentName = String(row[mapping.studentName] || '').trim().toUpperCase();
      const grades = mapping.activities.map(act => row[act.column] !== undefined ? row[act.column] : null);
      return { name: studentName, grades };
    });

    const script = `
(function() {
  const studentsData = ${JSON.stringify(scriptData)};
  const rows = Array.from(document.querySelectorAll('tr')).filter(tr => tr.querySelector('td.aluno') || tr.innerText.includes('Nº'));
  
  console.log('Iniciando lançamento de notas...');
  let count = 0;

  studentsData.forEach(student => {
    // Tenta encontrar a linha do aluno pelo nome
    const row = Array.from(document.querySelectorAll('tr')).find(tr => 
      tr.innerText.toUpperCase().includes(student.name)
    );

    if (row) {
      const inputs = Array.from(row.querySelectorAll('input[type="text"]'));
      // O SEGES costuma ter inputs em ordem: Av1, Rec1, Av2, Rec2...
      student.grades.forEach((grade, idx) => {
        if (grade !== null && inputs[idx]) {
          inputs[idx].value = grade.toString().replace('.', ',');
          // Dispara eventos para o sistema reconhecer a mudança e salvar
          inputs[idx].dispatchEvent(new Event('input', { bubbles: true }));
          inputs[idx].dispatchEvent(new Event('change', { bubbles: true }));
          inputs[idx].dispatchEvent(new Event('blur', { bubbles: true }));
        }
      });
      count++;
      row.style.backgroundColor = '#e6fffa'; // Destaca a linha processada
    } else {
      console.warn('Aluno não encontrado:', student.name);
    }
  });

  alert('Processamento concluído! ' + count + ' alunos atualizados. Verifique se as notas foram salvas (destaque amarelo do SEGES).');
})();
    `;

    navigator.clipboard.writeText(script);
    showSuccess("Script copiado para a área de transferência!");
  };

  return (
    <div className="min-h-screen bg-[#f8fafc] p-4 md:p-8">
      <div className="max-w-5xl mx-auto space-y-8">
        <header className="text-center space-y-2">
          <div className="inline-flex items-center justify-center p-3 bg-blue-600 rounded-2xl mb-4 shadow-lg shadow-blue-200">
            <FileSpreadsheet className="text-white w-8 h-8" />
          </div>
          <h1 className="text-3xl font-bold text-slate-900 tracking-tight">Lançador de Notas SEGES</h1>
          <p className="text-slate-500 max-w-2xl mx-auto">
            Automatize o lançamento de notas do 1º Trimestre de forma segura e rápida.
          </p>
        </header>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          {/* Coluna de Configuração */}
          <div className="lg:col-span-1 space-y-6">
            <Card className="border-none shadow-sm overflow-hidden">
              <CardHeader className="bg-white border-b border-slate-100">
                <CardTitle className="text-lg flex items-center gap-2">
                  <Info className="w-5 h-5 text-blue-500" />
                  1. Carregar Dados
                </CardTitle>
              </CardHeader>
              <CardContent className="p-6 space-y-4">
                <div className="space-y-2">
                  <Label htmlFor="excel-upload">Planilha Excel (.xlsx)</Label>
                  <Input 
                    id="excel-upload" 
                    type="file" 
                    accept=".xlsx, .xls" 
                    onChange={handleFileUpload}
                    className="cursor-pointer"
                  />
                </div>

                {columns.length > 0 && (
                  <div className="space-y-4 pt-4 animate-in fade-in slide-in-from-top-2">
                    <div className="space-y-2">
                      <Label>Coluna do Nome do Aluno</Label>
                      <Select onValueChange={(val) => setMapping({...mapping, studentName: val})}>
                        <SelectTrigger>
                          <SelectValue placeholder="Selecione a coluna" />
                        </SelectTrigger>
                        <SelectContent>
                          {columns.map(col => (
                            <SelectItem key={col} value={col}>{col}</SelectItem>
                          ))}
                        </SelectContent>
                      </Select>
                    </div>
                  </div>
                )}
              </CardContent>
            </Card>

            {columns.length > 0 && (
              <Card className="border-none shadow-sm overflow-hidden animate-in fade-in slide-in-from-bottom-2">
                <CardHeader className="bg-white border-b border-slate-100">
                  <CardTitle className="text-lg">2. Mapear Atividades</CardTitle>
                  <CardDescription>Relacione as colunas do Excel com os campos do SEGES</CardDescription>
                </CardHeader>
                <CardContent className="p-6 space-y-4">
                  {mapping.activities.map((act, idx) => (
                    <div key={idx} className="space-y-1.5">
                      <Label className="text-xs font-semibold text-slate-500 uppercase tracking-wider">
                        {act.label}
                      </Label>
                      <Select onValueChange={(val) => updateActivityMapping(idx, val)}>
                        <SelectTrigger className="h-9">
                          <SelectValue placeholder="Ignorar campo" />
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value="">(Ignorar)</SelectItem>
                          {columns.map(col => (
                            <SelectItem key={col} value={col}>{col}</SelectItem>
                          ))}
                        </SelectContent>
                      </Select>
                    </div>
                  ))}
                  
                  <Button 
                    onClick={generateScript}
                    className="w-full bg-blue-600 hover:bg-blue-700 text-white mt-4 h-11 rounded-xl shadow-md shadow-blue-100 transition-all active:scale-95"
                  >
                    <ClipboardCopy className="w-4 h-4 mr-2" />
                    Gerar e Copiar Script
                  </Button>
                </CardContent>
              </Card>
            )}
          </div>

          {/* Coluna de Visualização */}
          <div className="lg:col-span-2">
            <Card className="border-none shadow-sm h-full min-h-[500px] overflow-hidden">
              <CardHeader className="bg-white border-b border-slate-100 flex flex-row items-center justify-between">
                <div>
                  <CardTitle className="text-lg">Pré-visualização dos Dados</CardTitle>
                  <CardDescription>Confira se os dados estão corretos antes de gerar o script</CardDescription>
                </div>
                {data.length > 0 && (
                  <div className="px-3 py-1 bg-green-50 text-green-600 rounded-full text-xs font-medium flex items-center gap-1">
                    <CheckCircle2 className="w-3 h-3" />
                    {data.length} alunos encontrados
                  </div>
                )}
              </CardHeader>
              <CardContent className="p-0">
                {data.length > 0 ? (
                  <div className="overflow-auto max-h-[600px]">
                    <Table>
                      <TableHeader className="bg-slate-50 sticky top-0 z-10">
                        <TableRow>
                          <TableHead className="w-[250px]">Aluno</TableHead>
                          {mapping.activities.filter(a => a.column).map((act, i) => (
                            <TableHead key={i} className="text-center">{act.label}</TableHead>
                          ))}
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {data.slice(0, 15).map((row, i) => (
                          <TableRow key={i} className="hover:bg-slate-50/50 transition-colors">
                            <TableCell className="font-medium text-slate-700">
                              {row[mapping.studentName] || '---'}
                            </TableCell>
                            {mapping.activities.filter(a => a.column).map((act, j) => (
                              <TableCell key={j} className="text-center text-slate-600">
                                {row[act.column] ?? '-'}
                              </TableCell>
                            ))}
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                    {data.length > 15 && (
                      <div className="p-4 text-center text-slate-400 text-sm border-t border-slate-100">
                        Exibindo os primeiros 15 de {data.length} alunos...
                      </div>
                    )}
                  </div>
                ) : (
                  <div className="flex flex-col items-center justify-center h-[400px] text-slate-400 space-y-4">
                    <div className="p-4 bg-slate-50 rounded-full">
                      <FileSpreadsheet className="w-12 h-12 opacity-20" />
                    </div>
                    <p>Faça o upload de uma planilha para começar</p>
                  </div>
                )}
              </CardContent>
            </Card>

            <div className="mt-6 p-4 bg-amber-50 border border-amber-100 rounded-xl flex gap-3">
              <AlertCircle className="w-5 h-5 text-amber-500 shrink-0" />
              <div className="text-sm text-amber-800 space-y-1">
                <p className="font-semibold">Como usar no SEGES:</p>
                <ol className="list-decimal ml-4 space-y-1 opacity-90">
                  <li>Abra a tela de lançamento de notas no SEGES.</li>
                  <li>Pressione <b>F12</b> ou clique com o botão direito e vá em <b>Inspecionar</b>.</li>
                  <li>Clique na aba <b>Console</b>.</li>
                  <li>Cole o script gerado e pressione <b>Enter</b>.</li>
                  <li>Aguarde o preenchimento automático e verifique se as notas ficaram amarelas (salvas).</li>
                </ol>
              </div>
            </div>
          </div>
        </div>
        
        <MadeWithDyad />
      </div>
    </div>
  );
};

export default Index;