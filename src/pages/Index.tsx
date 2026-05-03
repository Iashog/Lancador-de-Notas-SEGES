import React, { useState, useMemo } from 'react';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { FileSpreadsheet, ClipboardCopy, GraduationCap, Settings2, Layers, AlertCircle, ChevronDown } from "lucide-react";
import * as XLSX from 'xlsx';
import { showSuccess, showError } from "@/utils/toast";
import { MadeWithDyad } from "@/components/made-with-dyad";
import { Collapsible, CollapsibleContent, CollapsibleTrigger } from "@/components/ui/collapsible";

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
  const [allColumns, setAllColumns] = useState<string[]>([]);
  const [studentNameCol, setStudentNameCol] = useState<string>('');
  const [turmaCol, setTurmaCol] = useState<string>('');
  const [showConfig, setShowConfig] = useState(false);
  
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
        const bData = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(bData, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet) as any[];
        
        if (jsonData.length > 0) {
          const cols = Object.keys(jsonData[0]);
          setAllColumns(cols);
          setData(jsonData);
          
          // Lógica de detecção inteligente baseada no conteúdo
          let detectedTurma = '';
          let detectedName = '';

          cols.forEach(col => {
            const sampleValue = String(jsonData[0][col] || '');
            // Padrão SEGES: Começa com número, tem ª e hífen (ex: 1ªV01-LCH)
            if (/\dª.*-/.test(sampleValue)) {
              detectedTurma = col;
            } 
            // Se não for turma e tiver espaços, provavelmente é nome
            else if (sampleValue.includes(' ') && !detectedName) {
              detectedName = col;
            }
          });

          // Fallback para nomes de colunas se a análise de dados falhar
          if (!detectedTurma) detectedTurma = cols.find(c => /turma|classe/i.test(c)) || '';
          if (!detectedName) detectedName = cols.find(c => /nome|aluno/i.test(c)) || '';

          setTurmaCol(detectedTurma);
          setStudentNameCol(detectedName);
          
          if (!detectedTurma || !detectedName) {
            setShowConfig(true);
            showError("Não identificamos as colunas automaticamente. Verifique os ajustes.");
          } else {
            showSuccess("Planilha carregada e colunas identificadas!");
          }
        }
      } catch (err) {
        showError("Erro ao processar Excel.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const uniqueTurmas = useMemo(() => {
    if (!turmaCol || data.length === 0) return [];
    const turmas = data.map(item => String(item[turmaCol] || '')).filter(Boolean);
    return Array.from(new Set(turmas)).sort();
  }, [data, turmaCol]);

  const filteredData = useMemo(() => {
    if (!selectedTurma || !turmaCol) return [];
    return data.filter(item => String(item[turmaCol]) === selectedTurma);
  }, [data, selectedTurma, turmaCol]);

  const generateScript = () => {
    if (!selectedTurma) return showError("Selecione a Turma.");
    if (!selectedArea) return showError("Selecione a Área.");

    const scriptData = filteredData.map(row => ({
      name: String(row[studentNameCol] || '').trim().toUpperCase(),
      av: row[selectedArea],
      rec: row[`R${selectedArea}`]
    }));

    const script = `
(function() {
  const studentsData = ${JSON.stringify(scriptData)};
  const category = "${selectedCategory}";
  const field = "${selectedField}";
  
  let count = 0;
  studentsData.forEach(student => {
    const row = Array.from(document.querySelectorAll('tr')).find(tr => 
      tr.innerText.toUpperCase().includes(student.name)
    );

    if (row) {
      const inputs = Array.from(row.querySelectorAll('input[type="text"]'));
      const targetIndices = category === 'all' ? [0, 1, 2, 3, 4, 5] : [parseInt(category)*2, parseInt(category)*2 + 1];

      targetIndices.forEach(idx => {
        const isAv = idx % 2 === 0;
        const val = isAv ? student.av : student.rec;
        if ((field === 'av' && !isAv) || (field === 'rec' && isAv)) return;

        if (val !== undefined && val !== null && inputs[idx]) {
          inputs[idx].value = val.toString().replace('.', ',');
          ['input', 'change', 'blur'].forEach(t => inputs[idx].dispatchEvent(new Event(t, { bubbles: true })));
        }
      });
      count++;
      row.style.backgroundColor = '#f0fdf4';
    }
  });
  alert('Sucesso! ' + count + ' alunos atualizados.');
})();`;

    navigator.clipboard.writeText(script);
    showSuccess("Script copiado!");
  };

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8">
      <div className="max-w-6xl mx-auto space-y-6">
        
        <header className="flex flex-col md:flex-row items-center justify-between bg-white p-6 rounded-3xl shadow-sm border border-slate-100 gap-4">
          <div className="flex items-center gap-4">
            <div className="p-3 bg-indigo-600 rounded-2xl shadow-lg shadow-indigo-100">
              <GraduationCap className="text-white w-8 h-8" />
            </div>
            <div>
              <h1 className="text-2xl font-bold text-slate-900 tracking-tight">Lançador SEGES</h1>
              <p className="text-slate-500 text-sm">Automatização inteligente de diários</p>
            </div>
          </div>
          <div className="flex items-center gap-3 bg-slate-50 p-2 rounded-2xl border border-slate-100">
            <FileSpreadsheet className="w-5 h-5 text-indigo-600 ml-2" />
            <Input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} className="max-w-[200px] border-none bg-transparent text-xs cursor-pointer" />
          </div>
        </header>

        <Collapsible open={showConfig} onOpenChange={setShowConfig} className="w-full">
          <div className="flex justify-end mb-2">
            <CollapsibleTrigger asChild>
              <Button variant="ghost" size="sm" className="text-[10px] text-slate-400 hover:text-indigo-600 uppercase font-bold tracking-widest">
                {showConfig ? 'Ocultar Ajustes' : 'Ajustar Colunas da Planilha'}
                <ChevronDown className={`ml-1 w-3 h-3 transition-transform ${showConfig ? 'rotate-180' : ''}`} />
              </Button>
            </CollapsibleTrigger>
          </div>
          <CollapsibleContent>
            <Card className="border-dashed border-2 border-slate-200 bg-slate-50/50 mb-6">
              <CardContent className="p-4 grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-1">
                  <Label className="text-[10px] font-bold text-slate-500 uppercase">Coluna do Nome do Aluno</Label>
                  <Select value={studentNameCol} onValueChange={setStudentNameCol}>
                    <SelectTrigger className="bg-white h-9 text-xs"><SelectValue placeholder="Selecione..." /></SelectTrigger>
                    <SelectContent>{allColumns.map(c => <SelectItem key={c} value={c}>{c}</SelectItem>)}</SelectContent>
                  </Select>
                </div>
                <div className="space-y-1">
                  <Label className="text-[10px] font-bold text-slate-500 uppercase">Coluna da Turma</Label>
                  <Select value={turmaCol} onValueChange={setTurmaCol}>
                    <SelectTrigger className="bg-white h-9 text-xs"><SelectValue placeholder="Selecione..." /></SelectTrigger>
                    <SelectContent>{allColumns.map(c => <SelectItem key={c} value={c}>{c}</SelectItem>)}</SelectContent>
                  </Select>
                </div>
              </CardContent>
            </Card>
          </CollapsibleContent>
        </Collapsible>

        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          <div className="lg:col-span-4 space-y-6">
            <Card className="border-none shadow-sm overflow-hidden">
              <CardHeader className="bg-indigo-600 text-white">
                <CardTitle className="text-sm font-bold flex items-center gap-2"><Settings2 className="w-4 h-4" /> Configurar Lançamento</CardTitle>
              </CardHeader>
              <CardContent className="p-5 space-y-5">
                <div className="space-y-2">
                  <Label className="text-xs font-bold text-slate-600 flex items-center gap-2">
                    <span className="w-5 h-5 rounded-full bg-indigo-100 text-indigo-600 flex items-center justify-center text-[10px]">1</span>
                    Turma (Obrigatório)
                  </Label>
                  <Select value={selectedTurma} onValueChange={setSelectedTurma} disabled={data.length === 0}>
                    <SelectTrigger className="bg-slate-50 border-slate-100 h-11">
                      <SelectValue placeholder={data.length === 0 ? "Carregue uma planilha..." : "Selecione a turma..."} />
                    </SelectTrigger>
                    <SelectContent>{uniqueTurmas.map(t => <SelectItem key={t} value={t}>{t}</SelectItem>)}</SelectContent>
                  </Select>
                </div>

                <div className="space-y-2">
                  <Label className="text-xs font-bold text-slate-600 flex items-center gap-2">
                    <span className="w-5 h-5 rounded-full bg-indigo-100 text-indigo-600 flex items-center justify-center text-[10px]">2</span>
                    Área (Obrigatório)
                  </Label>
                  <div className="grid grid-cols-2 gap-2">
                    {AREAS.map(area => (
                      <button key={area.id} onClick={() => setSelectedArea(area.id)} className={`p-3 rounded-xl text-xs font-bold transition-all border-2 ${selectedArea === area.id ? 'bg-indigo-600 border-indigo-600 text-white shadow-md' : 'bg-white border-slate-100 text-slate-500 hover:border-indigo-200'}`}>
                        {area.id}
                      </button>
                    ))}
                  </div>
                </div>

                <div className="space-y-2">
                  <Label className="text-xs font-bold text-slate-600 flex items-center gap-2">
                    <span className="w-5 h-5 rounded-full bg-slate-100 text-slate-400 flex items-center justify-center text-[10px]">3</span>
                    Tipo de Avaliação (Opcional)
                  </Label>
                  <Select value={selectedCategory} onValueChange={setSelectedCategory}>
                    <SelectTrigger className="bg-slate-50 border-slate-100 h-11"><SelectValue /></SelectTrigger>
                    <SelectContent>{CATEGORIES.map(cat => <SelectItem key={cat.id} value={cat.id}>{cat.name}</SelectItem>)}</SelectContent>
                  </Select>
                </div>

                <div className="space-y-2">
                  <Label className="text-xs font-bold text-slate-600 flex items-center gap-2">
                    <span className="w-5 h-5 rounded-full bg-slate-100 text-slate-400 flex items-center justify-center text-[10px]">4</span>
                    Campo de Nota (Opcional)
                  </Label>
                  <Select value={selectedField} onValueChange={setSelectedField}>
                    <SelectTrigger className="bg-slate-50 border-slate-100 h-11"><SelectValue /></SelectTrigger>
                    <SelectContent>{FIELDS.map(f => <SelectItem key={f.id} value={f.id}>{f.name}</SelectItem>)}</SelectContent>
                  </Select>
                </div>

                <Button onClick={generateScript} disabled={!selectedTurma || !selectedArea} className="w-full bg-indigo-600 hover:bg-indigo-700 py-7 text-md font-bold rounded-2xl shadow-lg mt-4 transition-all active:scale-95">
                  <ClipboardCopy className="w-5 h-5 mr-2" /> Copiar Script
                </Button>
              </CardContent>
            </Card>
          </div>

          <div className="lg:col-span-8">
            <Card className="border-none shadow-sm h-full min-h-[600px] overflow-hidden flex flex-col">
              <CardHeader className="border-b border-slate-100 flex flex-row items-center justify-between bg-white">
                <div>
                  <CardTitle className="text-lg">Visualização do Lançamento</CardTitle>
                  <CardDescription>{selectedTurma ? `Turma: ${selectedTurma}` : 'Selecione uma turma para visualizar'}</CardDescription>
                </div>
              </CardHeader>
              <CardContent className="p-0 flex-1 bg-white">
                {filteredData.length > 0 ? (
                  <div className="overflow-auto max-h-[650px]">
                    <Table>
                      <TableHeader className="bg-slate-50/50 sticky top-0 z-10">
                        <TableRow>
                          <TableHead className="font-bold text-slate-700 pl-6">Aluno</TableHead>
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
                            <TableCell className="font-medium text-slate-900 text-xs pl-6">{row[studentNameCol] || '-'}</TableCell>
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
                  <div className="flex flex-col items-center justify-center h-full min-h-[500px] text-slate-400">
                    <Layers className="w-16 h-16 mb-4 opacity-10" />
                    <p className="text-sm font-medium">Aguardando seleção de turma e área</p>
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