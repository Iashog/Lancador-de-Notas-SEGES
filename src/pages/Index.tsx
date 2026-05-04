import React, { useState, useMemo } from 'react';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { FileSpreadsheet, ClipboardCopy, Settings2, Layers, GraduationCap, Plus, Trash2, Info, CheckCircle2 } from "lucide-react";
import * as XLSX from 'xlsx';
import { showSuccess, showError } from "@/utils/toast";
import { MadeWithDyad } from "@/components/made-with-dyad";

interface SegesCategory {
  id: string;
  name: string;
  avCol: string;
  recCol: string;
}

const Index = () => {
  const [workbookData, setWorkbookData] = useState<Record<string, any[]>>({});
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [allColumns, setAllColumns] = useState<string[]>([]);
  
  const [selectedTurma, setSelectedTurma] = useState<string>('');
  const [studentNameCol, setStudentNameCol] = useState<string>('');
  
  // Categorias dinâmicas que o usuário define
  const [segesCategories, setSegesCategories] = useState<SegesCategory[]>([
    { id: '1', name: 'DISCURSIVA', avCol: '', recCol: '' },
    { id: '2', name: 'INTERDISCIPLINAR', avCol: '', recCol: '' },
    { id: '3', name: 'PRODUCAO ESCRITA', avCol: '', recCol: '' },
  ]);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bData = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(bData, { type: 'array' });
        
        const newWorkbookData: Record<string, any[]> = {};
        let detectedCols: string[] = [];

        workbook.SheetNames.forEach(name => {
          const worksheet = workbook.Sheets[name];
          const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" }) as any[][];
          
          // Tenta encontrar a linha de cabeçalho (primeira com conteúdo significativo)
          const headerIdx = rows.findIndex(r => r.filter(c => String(c).trim() !== "").length > 3);
          if (headerIdx === -1) return;

          const headers = rows[headerIdx].map((h, i) => String(h || `Coluna ${i+1}`).trim());
          detectedCols = Array.from(new Set([...detectedCols, ...headers]));

          const dataRows = rows.slice(headerIdx + 1);
          newWorkbookData[name] = dataRows.map(row => {
            const obj: any = {};
            headers.forEach((h, i) => obj[h] = row[i]);
            return obj;
          }).filter(row => String(Object.values(row)[0]).length > 2);
        });

        setWorkbookData(newWorkbookData);
        setSheetNames(workbook.SheetNames);
        setAllColumns(detectedCols);
        
        if (workbook.SheetNames.length > 0) {
          setSelectedTurma(workbook.SheetNames[0]);
          const nameCol = detectedCols.find(c => /nome|aluno/i.test(c)) || detectedCols[0];
          setStudentNameCol(nameCol);
          showSuccess("Planilha carregada com sucesso!");
        }
      } catch (err) {
        showError("Erro ao ler o arquivo Excel.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const addCategory = () => {
    setSegesCategories([...segesCategories, { id: Date.now().toString(), name: '', avCol: '', recCol: '' }]);
  };

  const removeCategory = (id: string) => {
    setSegesCategories(segesCategories.filter(c => c.id !== id));
  };

  const updateCategory = (id: string, field: keyof SegesCategory, value: string) => {
    setSegesCategories(segesCategories.map(c => c.id === id ? { ...c, [field]: value } : c));
  };

  const filteredData = useMemo(() => workbookData[selectedTurma] || [], [workbookData, selectedTurma]);

  const generateScript = () => {
    if (!selectedTurma || !studentNameCol) return showError("Selecione a turma e a coluna de nomes.");
    
    const activeMappings = segesCategories.filter(c => c.name && (c.avCol || c.recCol));
    if (activeMappings.length === 0) return showError("Configure ao menos uma categoria com colunas do Excel.");

    const scriptData = filteredData.map(row => {
      const mapping = activeMappings.map(cat => ({
        segesName: cat.name.toUpperCase(),
        avValue: row[cat.avCol] !== undefined ? row[cat.avCol] : "",
        recValue: row[cat.recCol] !== undefined ? row[cat.recCol] : ""
      }));
      return { 
        name: String(row[studentNameCol]).split(/[-—(]/)[0].trim().toUpperCase(),
        fullName: String(row[studentNameCol]).trim().toUpperCase(),
        mapping 
      };
    });

    const script = `
(function() {
  const studentsData = ${JSON.stringify(scriptData)};
  
  // Função para encontrar os índices das colunas no SEGES baseado no nome da categoria
  const getSegesColumnIndices = () => {
    const allThs = Array.from(document.querySelectorAll('thead tr:first-child th'));
    // Filtra apenas as colunas que possuem inputs (geralmente as de avaliação)
    const evalThs = allThs.filter(th => {
      const text = th.innerText.toUpperCase();
      return text.length > 2 && !text.includes('TOTAL') && !text.includes('FALTAS');
    });
    
    const map = {};
    evalThs.forEach((th, index) => {
      map[th.innerText.toUpperCase()] = {
        avIdx: index * 2,
        recIdx: index * 2 + 1
      };
    });
    return map;
  };

  const segesMap = getSegesColumnIndices();
  let count = 0;
  let notFound = [];

  studentsData.forEach(student => {
    const row = Array.from(document.querySelectorAll('tr')).find(tr => {
      const text = tr.innerText.toUpperCase();
      return text.includes(student.name);
    });

    if (row) {
      const inputs = Array.from(row.querySelectorAll('input[type="text"]'));
      
      student.mapping.forEach(m => {
        // Tenta encontrar a coluna no SEGES que contenha o nome configurado
        const segesKey = Object.keys(segesMap).find(key => key.includes(m.segesName));
        
        if (segesKey) {
          const { avIdx, recIdx } = segesMap[segesKey];

          if (m.avValue !== "" && inputs[avIdx]) {
            inputs[avIdx].value = m.avValue.toString().replace('.', ',');
            ['input', 'change', 'blur'].forEach(t => inputs[avIdx].dispatchEvent(new Event(t, { bubbles: true })));
          }
          if (m.recValue !== "" && inputs[recIdx]) {
            inputs[recIdx].value = m.recValue.toString().replace('.', ',');
            ['input', 'change', 'blur'].forEach(t => inputs[recIdx].dispatchEvent(new Event(t, { bubbles: true })));
          }
        }
      });
      
      count++;
      row.style.backgroundColor = '#f0fdf4';
      row.style.borderLeft = '4px solid #22c55e';
    } else {
      notFound.push(student.fullName);
    }
  });

  const msg = 'Sucesso! ' + count + ' alunos processados.';
  const warn = notFound.length > 0 ? '\\n\\n' + notFound.length + ' alunos não encontrados na página:\\n' + notFound.join('\\n') : '';
  alert(msg + warn);
})();`;

    navigator.clipboard.writeText(script);
    showSuccess("Script Universal copiado!");
  };

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8">
      <div className="max-w-7xl mx-auto space-y-6">
        
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          
          {/* Painel de Configuração Esquerdo */}
          <div className="lg:col-span-5 space-y-6">
            <Card className="border-none shadow-xl overflow-hidden rounded-3xl">
              <CardHeader className="bg-gradient-to-br from-indigo-600 to-violet-700 text-white pb-8">
                <div className="flex items-center gap-3 mb-2">
                  <div className="p-2 bg-white/20 rounded-xl">
                    <GraduationCap className="w-6 h-6" />
                  </div>
                  <div>
                    <CardTitle className="text-xl font-bold">Lançador Universal</CardTitle>
                    <CardDescription className="text-indigo-100 text-xs">Configure qualquer planilha para o SEGES</CardDescription>
                  </div>
                </div>
              </CardHeader>
              
              <CardContent className="p-6 space-y-6 -mt-6">
                <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-100 space-y-4">
                  <div className="space-y-2">
                    <Label className="text-[10px] font-bold text-slate-400 uppercase tracking-widest flex items-center gap-2">
                      <FileSpreadsheet className="w-3 h-3" /> 1. Importar Dados
                    </Label>
                    <Input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} className="h-11 text-xs rounded-xl border-slate-200" />
                  </div>

                  <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-2">
                      <Label className="text-[10px] font-bold text-slate-400 uppercase">2. Turma (Aba)</Label>
                      <Select value={selectedTurma} onValueChange={setSelectedTurma}>
                        <SelectTrigger className="h-10 text-xs rounded-xl"><SelectValue placeholder="Aba..." /></SelectTrigger>
                        <SelectContent>{sheetNames.map(n => <SelectItem key={n} value={n}>{n}</SelectItem>)}</SelectContent>
                      </Select>
                    </div>
                    <div className="space-y-2">
                      <Label className="text-[10px] font-bold text-slate-400 uppercase">3. Coluna Nomes</Label>
                      <Select value={studentNameCol} onValueChange={setStudentNameCol}>
                        <SelectTrigger className="h-10 text-xs rounded-xl"><SelectValue placeholder="Coluna..." /></SelectTrigger>
                        <SelectContent>{allColumns.map(c => <SelectItem key={c} value={c}>{c}</SelectItem>)}</SelectContent>
                      </Select>
                    </div>
                  </div>
                </div>

                <div className="space-y-4">
                  <div className="flex items-center justify-between px-1">
                    <Label className="text-xs font-bold text-indigo-600 uppercase tracking-wider flex items-center gap-2">
                      <Settings2 className="w-4 h-4" /> 4. Mapeamento de Notas
                    </Label>
                    <Button onClick={addCategory} variant="ghost" size="sm" className="h-8 text-[10px] text-indigo-600 hover:bg-indigo-50 rounded-lg gap-1">
                      <Plus className="w-3 h-3" /> Adicionar Campo
                    </Button>
                  </div>

                  <div className="space-y-3 max-h-[400px] overflow-y-auto pr-2 custom-scrollbar">
                    {segesCategories.map((cat) => (
                      <div key={cat.id} className="p-4 bg-slate-50 rounded-2xl border border-slate-100 space-y-3 relative group transition-all hover:border-indigo-200 hover:shadow-md hover:shadow-indigo-50">
                        <Button 
                          variant="ghost" 
                          size="icon" 
                          onClick={() => removeCategory(cat.id)}
                          className="absolute top-2 right-2 h-7 w-7 text-slate-300 hover:text-red-500 hover:bg-red-50 rounded-full opacity-0 group-hover:opacity-100 transition-opacity"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </Button>
                        
                        <div className="space-y-1">
                          <Label className="text-[9px] text-slate-400 font-bold uppercase">Nome no SEGES (ex: PROVA)</Label>
                          <Input 
                            placeholder="Ex: DISCURSIVA ou PROVA"
                            value={cat.name} 
                            onChange={(e) => updateCategory(cat.id, 'name', e.target.value)}
                            className="h-9 text-xs font-bold bg-white rounded-lg border-slate-200 focus:ring-indigo-500"
                          />
                        </div>

                        <div className="grid grid-cols-2 gap-3">
                          <div className="space-y-1">
                            <Label className="text-[9px] text-slate-400 font-bold uppercase">Coluna Av. (Excel)</Label>
                            <Select value={cat.avCol} onValueChange={(v) => updateCategory(cat.id, 'avCol', v)}>
                              <SelectTrigger className="h-8 text-[10px] bg-white rounded-lg"><SelectValue placeholder="Escolher..." /></SelectTrigger>
                              <SelectContent>{allColumns.map(c => <SelectItem key={c} value={c}>{c}</SelectItem>)}</SelectContent>
                            </Select>
                          </div>
                          <div className="space-y-1">
                            <Label className="text-[9px] text-slate-400 font-bold uppercase">Coluna Rec. (Excel)</Label>
                            <Select value={cat.recCol} onValueChange={(v) => updateCategory(cat.id, 'recCol', v)}>
                              <SelectTrigger className="h-8 text-[10px] bg-white rounded-lg"><SelectValue placeholder="Escolher..." /></SelectTrigger>
                              <SelectContent>{allColumns.map(c => <SelectItem key={c} value={c}>{c}</SelectItem>)}</SelectContent>
                            </Select>
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>

                <Button 
                  onClick={generateScript} 
                  disabled={!selectedTurma || segesCategories.every(c => !c.avCol && !c.recCol)} 
                  className="w-full bg-indigo-600 hover:bg-indigo-700 py-8 text-lg font-bold rounded-2xl shadow-xl shadow-indigo-100 transition-all active:scale-95"
                >
                  <ClipboardCopy className="w-6 h-6 mr-2" /> Copiar Script Completo
                </Button>

                <div className="bg-amber-50 p-4 rounded-2xl border border-amber-100 flex gap-3">
                  <Info className="w-5 h-5 text-amber-500 shrink-0 mt-0.5" />
                  <p className="text-[10px] text-amber-800 leading-relaxed">
                    <strong>Como usar:</strong> Após copiar o script, vá ao SEGES na tela de lançar notas, aperte <strong>F12</strong>, clique em <strong>Console</strong>, cole o código e aperte <strong>Enter</strong>.
                  </p>
                </div>
              </CardContent>
            </Card>
          </div>

          {/* Painel de Visualização Direito */}
          <div className="lg:col-span-7">
            <Card className="border-none shadow-xl h-full min-h-[600px] overflow-hidden flex flex-col rounded-3xl">
              <CardHeader className="border-b border-slate-100 bg-white p-6">
                <div className="flex items-center justify-between">
                  <div>
                    <CardTitle className="text-xl font-bold text-slate-800">Conferência de Dados</CardTitle>
                    <CardDescription className="text-xs">Verifique se as notas estão nas colunas certas</CardDescription>
                  </div>
                  {filteredData.length > 0 && (
                    <div className="flex items-center gap-2 text-[10px] font-bold text-emerald-600 bg-emerald-50 px-4 py-2 rounded-full border border-emerald-100">
                      <CheckCircle2 className="w-3 h-3" />
                      {filteredData.length} Alunos Identificados
                    </div>
                  )}
                </div>
              </CardHeader>
              <CardContent className="p-0 flex-1 bg-white">
                {filteredData.length > 0 ? (
                  <div className="overflow-auto max-h-[800px]">
                    <Table>
                      <TableHeader className="bg-slate-50/80 backdrop-blur-sm sticky top-0 z-10">
                        <TableRow>
                          <TableHead className="font-bold text-slate-700 pl-8 py-4">Aluno</TableHead>
                          {segesCategories.filter(c => c.name).map(cat => (
                            <TableHead key={cat.id} className="text-center font-bold text-indigo-600 text-[10px] uppercase leading-tight px-4">
                              {cat.name}<br/><span className="text-slate-400 font-normal">Av | Rec</span>
                            </TableHead>
                          ))}
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {filteredData.map((row, i) => (
                          <TableRow key={i} className="hover:bg-indigo-50/30 transition-colors border-b border-slate-50">
                            <TableCell className="font-medium text-slate-900 text-[10px] pl-8 py-3">{row[studentNameCol] || '-'}</TableCell>
                            {segesCategories.filter(c => c.name).map(cat => (
                              <TableCell key={cat.id} className="text-center text-[10px] font-bold px-4">
                                <span className="text-indigo-600">{row[cat.avCol] !== undefined && row[cat.avCol] !== "" ? row[cat.avCol] : "-"}</span>
                                <span className="mx-1.5 text-slate-200">|</span>
                                <span className="text-orange-600">{row[cat.recCol] !== undefined && row[cat.recCol] !== "" ? row[cat.recCol] : "-"}</span>
                              </TableCell>
                            ))}
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </div>
                ) : (
                  <div className="flex flex-col items-center justify-center h-full min-h-[500px] text-slate-400 p-12 text-center">
                    <div className="w-24 h-24 bg-slate-50 rounded-full flex items-center justify-center mb-6">
                      <Layers className="w-10 h-10 opacity-20" />
                    </div>
                    <p className="text-lg font-bold text-slate-600">Nenhum dado para exibir</p>
                    <p className="text-sm opacity-60 mt-2 max-w-xs mx-auto">
                      Importe sua planilha e configure o mapeamento das colunas para ver a prévia do lançamento aqui.
                    </p>
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