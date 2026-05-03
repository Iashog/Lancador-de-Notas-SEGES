import React, { useState, useMemo } from 'react';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { FileSpreadsheet, ClipboardCopy, Filter, GraduationCap, BookOpen, CheckCircle2 } from "lucide-react";
import * as XLSX from 'xlsx';
import { showSuccess, showError } from "@/utils/toast";
import { MadeWithDyad } from "@/components/made-with-dyad";

interface StudentData {
  [key: string]: any;
}

const AREAS = [
  { id: 'CH', name: 'Ciências Humanas', av: 'CH', rec: 'RCH' },
  { id: 'CN', name: 'Ciências da Natureza', av: 'CN', rec: 'RCN' },
  { id: 'LI', name: 'Linguagens', av: 'LI', rec: 'RLI' },
  { id: 'MT', name: 'Matemática', av: 'MT', rec: 'RMT' },
];

const Index = () => {
  const [data, setData] = useState<StudentData[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  
  // Configurações de Colunas
  const [studentNameCol, setStudentNameCol] = useState<string>('');
  const [turmaCol, setTurmaCol] = useState<string>('');
  
  // Filtros Ativos
  const [selectedTurma, setSelectedTurma] = useState<string>('all');
  const [selectedArea, setSelectedArea] = useState<string>('');

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
          
          // Tenta pré-selecionar colunas comuns
          const nameCol = cols.find(c => c.toLowerCase().includes('nome') || c.toLowerCase().includes('aluno'));
          const tCol = cols.find(c => c.toLowerCase().includes('turma'));
          if (nameCol) setStudentNameCol(nameCol);
          if (tCol) setTurmaCol(tCol);
          
          showSuccess("Planilha carregada com sucesso!");
        }
      } catch (err) {
        showError("Erro ao ler o arquivo Excel.");
      }
    };
    reader.readAsBinaryString(file);
  };

  const uniqueTurmas = useMemo(() => {
    if (!turmaCol || data.length === 0) return [];
    const turmas = data.map(item => String(item[turmaCol] || 'Sem Turma')).filter(Boolean);
    return Array.from(new Set(turmas)).sort();
  }, [data, turmaCol]);

  const filteredData = useMemo(() => {
    let result = data;
    if (selectedTurma !== 'all' && turmaCol) {
      result = result.filter(item => String(item[turmaCol]) === selectedTurma);
    }
    return result;
  }, [data, selectedTurma, turmaCol]);

  const generateScript = () => {
    if (!studentNameCol) return showError("Selecione a coluna do Nome.");
    if (!selectedArea) return showError("Selecione a Área de conhecimento.");

    const areaInfo = AREAS.find(a => a.id === selectedArea);
    if (!areaInfo) return;

    const scriptData = filteredData.map(row => {
      const name = String(row[studentNameCol] || '').trim().toUpperCase();
      const avGrade = row[areaInfo.av];
      const recGrade = row[areaInfo.rec];
      
      // No SEGES são 3 categorias, cada uma com Av e Rec (total 6 campos por aluno)
      // Repetimos a nota da área para as 3 categorias conforme solicitado
      const grades = [
        avGrade, recGrade, // Atividade Discursiva
        avGrade, recGrade, // Prova Interdisciplinar
        avGrade, recGrade  // Produção Escrita
      ];
      
      return { name, grades };
    });

    const script = `
(function() {
  const studentsData = ${JSON.stringify(scriptData)};
  console.log('Iniciando lançamento para área: ${areaInfo.name}');
  let count = 0;

  studentsData.forEach(student => {
    const row = Array.from(document.querySelectorAll('tr')).find(tr => 
      tr.innerText.toUpperCase().includes(student.name)
    );

    if (row) {
      const inputs = Array.from(row.querySelectorAll('input[type="text"]'));
      student.grades.forEach((grade, idx) => {
        if (grade !== undefined && grade !== null && inputs[idx]) {
          inputs[idx].value = grade.toString().replace('.', ',');
          ['input', 'change', 'blur'].forEach(type => 
            inputs[idx].dispatchEvent(new Event(type, { bubbles: true }))
          );
        }
      });
      count++;
      row.style.backgroundColor = '#f0fdf4';
      row.style.borderLeft = '4px solid #22c55e';
    }
  });

  alert('Sucesso! ' + count + ' alunos da turma "${selectedTurma}" foram atualizados.');
})();
    `;

    navigator.clipboard.writeText(script);
    showSuccess("Script de lançamento copiado!");
  };

  return (
    <div className="min-h-screen bg-[#f8fafc] p-4 md:p-8">
      <div className="max-w-7xl mx-auto space-y-6">
        
        {/* Header Minimalista */}
        <header className="flex flex-col md:flex-row md:items-center justify-between gap-4 bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
          <div className="flex items-center gap-4">
            <div className="p-3 bg-indigo-600 rounded-2xl shadow-lg shadow-indigo-100">
              <GraduationCap className="text-white w-8 h-8" />
            </div>
            <div>
              <h1 className="text-2xl font-bold text-slate-900">Lançador SEGES</h1>
              <p className="text-slate-500 text-sm">Filtre por turma e área para automatizar o diário</p>
            </div>
          </div>
          
          <div className="flex items-center gap-2">
            <Input 
              type="file" 
              accept=".xlsx,.xls" 
              onChange={handleFileUpload}
              className="max-w-[250px] cursor-pointer file:mr-4 file:py-1 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-indigo-50 file:text-indigo-700 hover:file:bg-indigo-100"
            />
          </div>
        </header>

        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          
          {/* Painel de Controle */}
          <div className="lg:col-span-4 space-y-6">
            
            {/* Configuração de Colunas */}
            <Card className="border-none shadow-sm overflow-hidden">
              <CardHeader className="bg-slate-50/50 border-b border-slate-100">
                <CardTitle className="text-sm font-bold flex items-center gap-2">
                  <Filter className="w-4 h-4 text-indigo-500" />
                  1. Configurar Planilha
                </CardTitle>
              </CardHeader>
              <CardContent className="p-5 space-y-4">
                <div className="space-y-2">
                  <Label className="text-xs font-semibold text-slate-500 uppercase">Coluna da Turma</Label>
                  <Select value={turmaCol} onValueChange={setTurmaCol}>
                    <SelectTrigger className="bg-slate-50 border-none">
                      <SelectValue placeholder="Selecione a coluna..." />
                    </SelectTrigger>
                    <SelectContent>
                      {columns.map(c => <SelectItem key={c} value={c}>{c}</SelectItem>)}
                    </SelectContent>
                  </Select>
                </div>
                <div className="space-y-2">
                  <Label className="text-xs font-semibold text-slate-500 uppercase">Coluna do Nome</Label>
                  <Select value={studentNameCol} onValueChange={setStudentNameCol}>
                    <SelectTrigger className="bg-slate-50 border-none">
                      <SelectValue placeholder="Selecione a coluna..." />
                    </SelectTrigger>
                    <SelectContent>
                      {columns.map(c => <SelectItem key={c} value={c}>{c}</SelectItem>)}
                    </SelectContent>
                  </Select>
                </div>
              </CardContent>
            </Card>

            {/* Filtros de Lançamento */}
            <Card className="border-none shadow-sm overflow-hidden">
              <CardHeader className="bg-slate-50/50 border-b border-slate-100">
                <CardTitle className="text-sm font-bold flex items-center gap-2">
                  <CheckCircle2 className="w-4 h-4 text-green-500" />
                  2. Selecionar Alvo
                </CardTitle>
              </CardHeader>
              <CardContent className="p-5 space-y-4">
                <div className="space-y-2">
                  <Label className="text-xs font-semibold text-slate-500 uppercase">Filtrar Turma</Label>
                  <Select value={selectedTurma} onValueChange={setSelectedTurma} disabled={!turmaCol}>
                    <SelectTrigger className="bg-slate-50 border-none">
                      <SelectValue placeholder="Todas as turmas" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="all">Todas as turmas</SelectItem>
                      {uniqueTurmas.map(t => <SelectItem key={t} value={t}>{t}</SelectItem>)}
                    </SelectContent>
                  </Select>
                </div>

                <div className="space-y-2">
                  <Label className="text-xs font-semibold text-slate-500 uppercase">Área de Conhecimento</Label>
                  <div className="grid grid-cols-2 gap-2">
                    {AREAS.map(area => (
                      <button
                        key={area.id}
                        onClick={() => setSelectedArea(area.id)}
                        className={`p-3 rounded-xl text-xs font-bold transition-all border-2 ${
                          selectedArea === area.id 
                          ? 'bg-indigo-600 border-indigo-600 text-white shadow-md' 
                          : 'bg-white border-slate-100 text-slate-600 hover:border-indigo-200'
                        }`}
                      >
                        {area.id}
                      </button>
                    ))}
                  </div>
                </div>

                <Button 
                  onClick={generateScript} 
                  disabled={!selectedArea || filteredData.length === 0}
                  className="w-full bg-indigo-600 hover:bg-indigo-700 py-6 text-md font-bold rounded-2xl shadow-lg shadow-indigo-100 mt-2"
                >
                  <ClipboardCopy className="w-5 h-5 mr-2" />
                  Copiar Script de Lançamento
                </Button>
              </CardContent>
            </Card>
          </div>

          {/* Visualização */}
          <div className="lg:col-span-8">
            <Card className="border-none shadow-sm h-full min-h-[500px] overflow-hidden">
              <CardHeader className="border-b border-slate-100 flex flex-row items-center justify-between">
                <div>
                  <CardTitle className="text-lg">Visualização da Turma</CardTitle>
                  <CardDescription>
                    {selectedTurma === 'all' ? 'Mostrando todos os alunos' : `Turma: ${selectedTurma}`}
                  </CardDescription>
                </div>
                {selectedArea && (
                  <div className="px-3 py-1 bg-green-50 text-green-700 rounded-full text-[10px] font-bold uppercase tracking-wider border border-green-100">
                    Mapeado: {selectedArea} & {AREAS.find(a => a.id === selectedArea)?.rec}
                  </div>
                )}
              </CardHeader>
              <CardContent className="p-0">
                {filteredData.length > 0 ? (
                  <div className="overflow-auto max-h-[600px]">
                    <Table>
                      <TableHeader className="bg-slate-50/50 sticky top-0 z-10">
                        <TableRow>
                          <TableHead className="font-bold text-slate-700">Aluno</TableHead>
                          {selectedArea && (
                            <>
                              <TableHead className="text-center font-bold text-indigo-600">Av ({selectedArea})</TableHead>
                              <TableHead className="text-center font-bold text-orange-600">Rec ({AREAS.find(a => a.id === selectedArea)?.rec})</TableHead>
                            </>
                          )}
                          {turmaCol && <TableHead className="text-right font-bold text-slate-400">Turma</TableHead>}
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {filteredData.map((row, i) => (
                          <TableRow key={i} className="hover:bg-slate-50/50 transition-colors">
                            <TableCell className="font-medium text-slate-900">{row[studentNameCol] || '-'}</TableCell>
                            {selectedArea && (
                              <>
                                <TableCell className="text-center font-bold text-indigo-600">{row[AREAS.find(a => a.id === selectedArea)!.av] ?? '-'}</TableCell>
                                <TableCell className="text-center font-bold text-orange-600">{row[AREAS.find(a => a.id === selectedArea)!.rec] ?? '-'}</TableCell>
                              </>
                            )}
                            {turmaCol && <TableCell className="text-right text-slate-400 text-xs">{row[turmaCol]}</TableCell>}
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </div>
                ) : (
                  <div className="flex flex-col items-center justify-center h-[500px] text-slate-400">
                    <FileSpreadsheet className="w-16 h-16 mb-4 opacity-10" />
                    <p className="text-sm">Carregue uma planilha para começar</p>
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