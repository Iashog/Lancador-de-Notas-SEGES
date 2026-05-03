import React, { useState, useMemo } from 'react';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { FileSpreadsheet, ClipboardCopy, Filter, GraduationCap, Settings2, CheckCircle2, Layers } from "lucide-react";
import * as XLSX from 'xlsx';
import { showSuccess, showError } from "@/utils/toast";
import { MadeWithDyad } from "@/components/made-with-dyad";

interface StudentData {
  [key: string]: any;
}

const AREAS = [
  { id: 'CH', name: 'Ciências Humanas' },
  { id: 'CN', name: 'Ciências da Natureza' },
  { id: 'LI', name: 'Linguagens' },
  { id: 'MT', name: 'Matemática' },
];

const CATEGORIES = [
  { id: 'all', name: 'Todas as Categorias' },
  { id: '0', name: 'Atividade Discursiva' },
  { id: '1', name: 'Prova Interdisciplinar' },
  { id: '2', name: 'Produção Escrita' },
];

const FIELDS = [
  { id: 'both', name: 'Avaliação e Recuperação' },
  { id: 'av', name: 'Apenas Avaliação (Av)' },
  { id: 'rec', name: 'Apenas Recuperação (Rec)' },
];

const Index = () => {
  const [data, setData] = useState<StudentData[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  
  // Configuração Base da Planilha
  const [studentNameCol, setStudentNameCol] = useState<string>('');
  const [turmaCol, setTurmaCol] = useState<string>('');
  
  // Filtros e Mapeamento (Ordem solicitada)
  const [selectedTurma, setSelectedTurma] = useState<string>('');
  const [selectedArea, setSelectedArea] = useState<string>('');
  const [selectedCategory, setSelectedCategory] = useState<string>('all');
  const [selectedField, setSelectedField] = useState<string>('both');

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
          
          const nameCol = cols.find(c => c.toLowerCase().includes('nome') || c.toLowerCase().includes('aluno'));
          const tCol = cols.find(c => c.toLowerCase().includes('turma'));
          if (nameCol) setStudentNameCol(nameCol);
          if (tCol) setTurmaCol(tCol);
          
          showSuccess("Planilha carregada!");
        }
      } catch (err) {
        showError("Erro ao ler Excel.");
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
    if (!selectedTurma || selectedTurma === 'all') return data;
    return data.filter(item => String(item[turmaCol]) === selectedTurma);
  }, [data, selectedTurma, turmaCol]);

  const generateScript = () => {
    if (!selectedTurma) return showError("Selecione a Turma.");
    if (!selectedArea) return showError("Selecione a Área.");

    const avCol = selectedArea;
    const recCol = `R${selectedArea}`;

    const scriptData = filteredData.map(row => ({
      name: String(row[studentNameCol] || '').trim().toUpperCase(),
      av: row[avCol],
      rec: row[recCol]
    }));

    const script = `
(function() {
  const studentsData = ${JSON.stringify(scriptData)};
  const category = "${selectedCategory}";
  const field = "${selectedField}";
  
  console.log('Iniciando lançamento...');
  let count = 0;

  studentsData.forEach(student => {
    const row = Array.from(document.querySelectorAll('tr')).find(tr => 
      tr.innerText.toUpperCase().includes(student.name)
    );

    if (row) {
      const inputs = Array.from(row.querySelectorAll('input[type="text"]'));
      
      // Mapeamento de índices no SEGES:
      // Cat 0 (Discursiva): 0 (Av), 1 (Rec)
      // Cat 1 (Interdisciplinar): 2 (Av), 3 (Rec)
      // Cat 2 (Produção): 4 (Av), 5 (Rec)
      
      const targetIndices = [];
      if (category === 'all') {
        targetIndices.push(0, 1, 2, 3, 4, 5);
      } else {
        const base = parseInt(category) * 2;
        targetIndices.push(base, base + 1);
      }

      targetIndices.forEach(idx => {
        const isAv = idx % 2 === 0;
        const val = isAv ? student.av : student.rec;
        
        // Filtra pelo campo selecionado (Av, Rec ou Ambos)
        if (field === 'av' && !isAv) return;
        if (field === 'rec' && isAv) return;

        if (val !== undefined && val !== null && inputs[idx]) {
          inputs[idx].value = val.toString().replace('.', ',');
          ['input', 'change', 'blur'].forEach(type => 
            inputs[idx].dispatchEvent(new Event(type, { bubbles: true }))
          );
        }
      });

      count++;
      row.style.backgroundColor = '#f0fdf4';
    }
  });

  alert('Sucesso! ' + count + ' alunos atualizados.');
})();
    `;

    navigator.clipboard.writeText(script);
    showSuccess("Script copiado!");
  };

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8">
      <div className="max-w-6xl mx-auto space-y-6">
        
        <header className="flex items-center justify-between bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
          <div className="flex items-center gap-4">
            <div className="p-3 bg-indigo-600 rounded-2xl shadow-lg shadow-indigo-100">
              <GraduationCap className="text-white w-8 h-8" />
            </div>
            <div>
              <h1 className="text-2xl font-bold text-slate-900 tracking-tight">Lançador SEGES</h1>
              <p className="text-slate-500 text-sm">Configuração inteligente de diário</p>
            </div>
          </div>
          <Input 
            type="file" 
            accept=".xlsx,.xls" 
            onChange={handleFileUpload}
            className="max-w-[220px] text-xs"
          />
        </header>

        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          
          {/* Configuração da Planilha (Passo 0) */}
          <div className="lg:col-span-12">
            <Card className="border-none shadow-sm">
              <CardContent className="p-4 flex flex-wrap gap-6 items-end">
                <div className="flex-1 min-w-[200px] space-y-2">
                  <Label className="text-[10px] font-bold uppercase text-slate-400">Coluna do Nome</Label>
                  <Select value={studentNameCol} onValueChange={setStudentNameCol}>
                    <SelectTrigger className="bg-slate-50 border-none h-10">
                      <SelectValue placeholder="Selecione..." />
                    </SelectTrigger>
                    <SelectContent>
                      {columns.map(c => <SelectItem key={c} value={c}>{c}</SelectItem>)}
                    </SelectContent>
                  </Select>
                </div>
                <div className="flex-1 min-w-[200px] space-y-2">
                  <Label className="text-[10px] font-bold uppercase text-slate-400">Coluna da Turma</Label>
                  <Select value={turmaCol} onValueChange={setTurmaCol}>
                    <SelectTrigger className="bg-slate-50 border-none h-10">
                      <SelectValue placeholder="Selecione..." />
                    </SelectTrigger>
                    <SelectContent>
                      {columns.map(c => <SelectItem key={c} value={c}>{c}</SelectItem>)}
                    </SelectContent>
                  </Select>
                </div>
              </CardContent>
            </Card>
          </div>

          {/* Filtros de Lançamento (Ordem Solicitada) */}
          <div className="lg:col-span-4 space-y-6">
            <Card className="border-none shadow-sm overflow-hidden">
              <CardHeader className="bg-indigo-600 text-white">
                <CardTitle className="text-sm font-bold flex items-center gap-2">
                  <Settings2 className="w-4 h-4" />
                  Configurar Lançamento
                </CardTitle>
              </CardHeader>
              <CardContent className="p-5 space-y-5">
                
                {/* 1. Turma (Obrigatório) */}
                <div className="space-y-2">
                  <Label className="text-xs font-bold text-slate-600 flex items-center gap-2">
                    <span className="w-5 h-5 rounded-full bg-indigo-100 text-indigo-600 flex items-center justify-center text-[10px]">1</span>
                    Turma (Obrigatório)
                  </Label>
                  <Select value={selectedTurma} onValueChange={setSelectedTurma}>
                    <SelectTrigger className="bg-slate-50 border-slate-100">
                      <SelectValue placeholder="Selecione a turma..." />
                    </SelectTrigger>
                    <SelectContent>
                      {uniqueTurmas.map(t => <SelectItem key={t} value={t}>{t}</SelectItem>)}
                    </SelectContent>
                  </Select>
                </div>

                {/* 2. Área (Obrigatório) */}
                <div className="space-y-2">
                  <Label className="text-xs font-bold text-slate-600 flex items-center gap-2">
                    <span className="w-5 h-5 rounded-full bg-indigo-100 text-indigo-600 flex items-center justify-center text-[10px]">2</span>
                    Área (Obrigatório)
                  </Label>
                  <div className="grid grid-cols-2 gap-2">
                    {AREAS.map(area => (
                      <button
                        key={area.id}
                        onClick={() => setSelectedArea(area.id)}
                        className={`p-2 rounded-lg text-[10px] font-bold transition-all border-2 ${
                          selectedArea === area.id 
                          ? 'bg-indigo-600 border-indigo-600 text-white' 
                          : 'bg-white border-slate-100 text-slate-500 hover:border-indigo-200'
                        }`}
                      >
                        {area.id}
                      </button>
                    ))}
                  </div>
                </div>

                {/* 3. Tipo de Avaliação (Opcional) */}
                <div className="space-y-2">
                  <Label className="text-xs font-bold text-slate-600 flex items-center gap-2">
                    <span className="w-5 h-5 rounded-full bg-slate-100 text-slate-400 flex items-center justify-center text-[10px]">3</span>
                    Tipo de Avaliação (Opcional)
                  </Label>
                  <Select value={selectedCategory} onValueChange={setSelectedCategory}>
                    <SelectTrigger className="bg-slate-50 border-slate-100">
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      {CATEGORIES.map(cat => <SelectItem key={cat.id} value={cat.id}>{cat.name}</SelectItem>)}
                    </SelectContent>
                  </Select>
                </div>

                {/* 4. Campo (Opcional) */}
                <div className="space-y-2">
                  <Label className="text-xs font-bold text-slate-600 flex items-center gap-2">
                    <span className="w-5 h-5 rounded-full bg-slate-100 text-slate-400 flex items-center justify-center text-[10px]">4</span>
                    Campo de Nota (Opcional)
                  </Label>
                  <Select value={selectedField} onValueChange={setSelectedField}>
                    <SelectTrigger className="bg-slate-50 border-slate-100">
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      {FIELDS.map(f => <SelectItem key={f.id} value={f.id}>{f.name}</SelectItem>)}
                    </SelectContent>
                  </Select>
                </div>

                <Button 
                  onClick={generateScript} 
                  disabled={!selectedTurma || !selectedArea}
                  className="w-full bg-indigo-600 hover:bg-indigo-700 py-6 text-md font-bold rounded-2xl shadow-lg shadow-indigo-100 mt-4"
                >
                  <ClipboardCopy className="w-5 h-5 mr-2" />
                  Copiar Script
                </Button>
              </CardContent>
            </Card>
          </div>

          {/* Visualização */}
          <div className="lg:col-span-8">
            <Card className="border-none shadow-sm h-full min-h-[500px] overflow-hidden">
              <CardHeader className="border-b border-slate-100 flex flex-row items-center justify-between">
                <div>
                  <CardTitle className="text-lg">Visualização do Lançamento</CardTitle>
                  <CardDescription>
                    {selectedTurma ? `Turma: ${selectedTurma}` : 'Selecione uma turma para visualizar'}
                  </CardDescription>
                </div>
                {selectedArea && (
                  <div className="flex gap-2">
                    <div className="px-3 py-1 bg-indigo-50 text-indigo-700 rounded-full text-[10px] font-bold border border-indigo-100">
                      Av: {selectedArea}
                    </div>
                    <div className="px-3 py-1 bg-orange-50 text-orange-700 rounded-full text-[10px] font-bold border border-orange-100">
                      Rec: R{selectedArea}
                    </div>
                  </div>
                )}
              </CardHeader>
              <CardContent className="p-0">
                {filteredData.length > 0 && selectedTurma ? (
                  <div className="overflow-auto max-h-[600px]">
                    <Table>
                      <TableHeader className="bg-slate-50/50 sticky top-0 z-10">
                        <TableRow>
                          <TableHead className="font-bold text-slate-700">Aluno</TableHead>
                          {selectedArea && (
                            <>
                              <TableHead className="text-center font-bold text-indigo-600">Nota ({selectedArea})</TableHead>
                              <TableHead className="text-center font-bold text-orange-600">Rec (R{selectedArea})</TableHead>
                            </>
                          )}
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {filteredData.map((row, i) => (
                          <TableRow key={i} className="hover:bg-slate-50/50 transition-colors">
                            <TableCell className="font-medium text-slate-900 text-xs">{row[studentNameCol] || '-'}</TableCell>
                            {selectedArea && (
                              <>
                                <TableCell className="text-center font-bold text-indigo-600 text-xs">{row[selectedArea] ?? '-'}</TableCell>
                                <TableCell className="text-center font-bold text-orange-600 text-xs">{row[`R${selectedArea}`] ?? '-'}</TableCell>
                              </>
                            )}
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </div>
                ) : (
                  <div className="flex flex-col items-center justify-center h-[500px] text-slate-400">
                    <Layers className="w-16 h-16 mb-4 opacity-10" />
                    <p className="text-sm">Selecione a turma e a área para visualizar os dados</p>
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