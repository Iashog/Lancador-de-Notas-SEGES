import React, { useState, useMemo } from 'react';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { FileSpreadsheet, ClipboardCopy, Settings2, Layers, GraduationCap, CheckCircle2, AlertCircle } from "lucide-react";
import * as XLSX from 'xlsx';
import { showSuccess, showError } from "@/utils/toast";
import { MadeWithDyad } from "@/components/made-with-dyad";

interface WorkbookData {
  [sheetName: string]: any[];
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
  const [workbookData, setWorkbookData] = useState<WorkbookData>({});
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  
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
        
        const newWorkbookData: WorkbookData = {};
        workbook.SheetNames.forEach(name => {
          const worksheet = workbook.Sheets[name];
          const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" }) as any[][];
          
          let areaRowIndex = rows.findIndex(row => 
            row.some(cell => ['CH', 'CN', 'LI', 'MT'].includes(String(cell).trim().toUpperCase()))
          );

          if (areaRowIndex === -1) return;

          const rawAreas = rows[areaRowIndex].map(h => String(h || '').trim().toUpperCase());
          const processedHeaders: string[] = [];
          let lastArea = "";

          rawAreas.forEach((area, i) => {
            if (['CH', 'CN', 'LI', 'MT'].includes(area)) {
              lastArea = area;
            }
            if (area.startsWith('R') && ['RCH', 'RCN', 'RLI', 'RMT'].includes(area)) {
                processedHeaders.push(area);
            } else {
                processedHeaders.push(lastArea || `COL_${i}`);
            }
          });

          const dataRows = rows.slice(areaRowIndex + 1);
          const formattedData = dataRows.map(row => {
            const studentObj: any = { _raw: row };
            AREAS.forEach(area => {
                studentObj[area.id] = row.filter((_, i) => processedHeaders[i] === area.id);
                studentObj[`R${area.id}`] = row.filter((_, i) => processedHeaders[i] === `R${area.id}`);
            });
            studentObj._name = String(row.find(cell => String(cell).length > 5 && !/\d/.test(String(cell))) || "");
            return studentObj;
          }).filter(row => row._name.length > 3 && !row._name.includes("TOTAL") && !row._name.includes("MÉDIA"));

          newWorkbookData[name] = formattedData;
        });

        setWorkbookData(newWorkbookData);
        setSheetNames(workbook.SheetNames);
        
        if (workbook.SheetNames.length > 0) {
          setSelectedTurma(workbook.SheetNames[0]);
          showSuccess("Planilha processada!");
        }
      } catch (err) {
        showError("Erro ao processar Excel.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const filteredData = useMemo(() => {
    return workbookData[selectedTurma] || [];
  }, [workbookData, selectedTurma]);

  const generateScript = () => {
    if (!selectedTurma) return showError("Selecione a Turma.");
    if (!selectedArea) return showError("Selecione a Área.");

    const scriptData = filteredData.map(row => {
      const notes = row[selectedArea] || [];
      const recs = row[`R${selectedArea}`] || [];
      
      // Se uma categoria específica for selecionada, enviamos apenas ela para o script
      let finalValue = "";
      let targetInputIndex = -1;

      if (selectedCategory !== 'all') {
        const catIdx = parseInt(selectedCategory);
        if (selectedField === 'av' || selectedField === 'both') {
            finalValue = String(notes[catIdx] || "");
            targetInputIndex = catIdx * 2;
        } else if (selectedField === 'rec') {
            finalValue = String(recs[catIdx] || "");
            targetInputIndex = (catIdx * 2) + 1;
        }
      }

      return {
        name: row._name.trim().toUpperCase(),
        value: finalValue,
        index: targetInputIndex,
        // Fallback para 'all'
        allNotes: notes,
        allRecs: recs
      };
    });

    const isSingle = selectedCategory !== 'all' && selectedField !== 'both';

    const script = `
(function() {
  const studentsData = ${JSON.stringify(scriptData)};
  const isSingle = ${isSingle};
  const category = "${selectedCategory}";
  const field = "${selectedField}";
  
  let count = 0;
  studentsData.forEach(student => {
    const row = Array.from(document.querySelectorAll('tr')).find(tr => 
      tr.innerText.toUpperCase().includes(student.name)
    );

    if (row) {
      const inputs = Array.from(row.querySelectorAll('input[type="text"]'));
      
      if (isSingle) {
        // Modo Ultra-Preciso: Apenas 1 coluna
        const input = inputs[student.index];
        if (input && student.value !== "") {
          input.value = student.value.toString().replace('.', ',');
          ['input', 'change', 'blur'].forEach(t => input.dispatchEvent(new Event(t, { bubbles: true })));
        }
      } else {
        // Modo Múltiplo (Todas as categorias ou Av+Rec)
        const targetIndices = category === 'all' ? [0, 1, 2, 3, 4, 5] : [parseInt(category)*2, parseInt(category)*2 + 1];
        targetIndices.forEach(idx => {
          const catIdx = Math.floor(idx / 2);
          const isAv = idx % 2 === 0;
          if ((field === 'av' && !isAv) || (field === 'rec' && isAv)) return;
          const val = isAv ? student.allNotes[catIdx] : student.allRecs[catIdx];
          if (val !== undefined && val !== null && val !== "" && inputs[idx]) {
            inputs[idx].value = val.toString().replace('.', ',');
            ['input', 'change', 'blur'].forEach(t => inputs[idx].dispatchEvent(new Event(t, { bubbles: true })));
          }
        });
      }
      count++;
      row.style.backgroundColor = '#f0fdf4';
      row.style.borderLeft = '4px solid #22c55e';
    }
  });
  alert('Sucesso! ' + count + ' alunos atualizados na turma "${selectedTurma}".');
})();`;

    navigator.clipboard.writeText(script);
    showSuccess("Script exclusivo copiado!");
  };

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8">
      <div className="max-w-6xl mx-auto space-y-6">
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          <div className="lg:col-span-4 space-y-6">
            <Card className="border-none shadow-sm overflow-hidden">
              <CardHeader className="bg-indigo-600 text-white pb-8">
                <div className="flex items-center gap-3 mb-2">
                  <GraduationCap className="w-6 h-6" />
                  <CardTitle className="text-lg font-bold">Lançador SEGES</CardTitle>
                </div>
                <CardDescription className="text-indigo-100 text-xs">
                  Gere scripts precisos para o lançamento de notas.
                </CardDescription>
              </CardHeader>
              
              <CardContent className="p-5 space-y-6 -mt-6">
                <div className="bg-white p-4 rounded-2xl shadow-sm border border-slate-100 space-y-3">
                  <Label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider flex items-center gap-2">
                    <FileSpreadsheet className="w-3 h-3" /> 1. Planilha
                  </Label>
                  <Input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} className="h-12 text-xs" />
                </div>

                <div className="space-y-2">
                  <Label className="text-xs font-bold text-slate-600">2. Turma</Label>
                  <Select value={selectedTurma} onValueChange={setSelectedTurma} disabled={sheetNames.length === 0}>
                    <SelectTrigger className="bg-slate-50 h-11"><SelectValue placeholder="Selecione..." /></SelectTrigger>
                    <SelectContent>{sheetNames.map(name => <SelectItem key={name} value={name}>{name}</SelectItem>)}</SelectContent>
                  </Select>
                </div>

                <div className="space-y-2">
                  <Label className="text-xs font-bold text-slate-600">3. Área</Label>
                  <div className="grid grid-cols-2 gap-2">
                    {AREAS.map(area => (
                      <button key={area.id} onClick={() => setSelectedArea(area.id)} className={`p-2 rounded-lg text-xs font-bold border-2 transition-all ${selectedArea === area.id ? 'bg-indigo-600 border-indigo-600 text-white' : 'bg-white border-slate-100 text-slate-500'}`}>{area.id}</button>
                    ))}
                  </div>
                </div>

                <div className="grid grid-cols-1 gap-4">
                  <div className="space-y-2">
                    <Label className="text-[10px] font-bold text-slate-400 uppercase">Categoria</Label>
                    <Select value={selectedCategory} onValueChange={setSelectedCategory}>
                      <SelectTrigger className="bg-slate-50 h-10 text-xs"><SelectValue /></SelectTrigger>
                      <SelectContent>{CATEGORIES.map(cat => <SelectItem key={cat.id} value={cat.id}>{cat.name}</SelectItem>)}</SelectContent>
                    </Select>
                  </div>
                  <div className="space-y-2">
                    <Label className="text-[10px] font-bold text-slate-400 uppercase">Campo</Label>
                    <Select value={selectedField} onValueChange={setSelectedField}>
                      <SelectTrigger className="bg-slate-50 h-10 text-xs"><SelectValue /></SelectTrigger>
                      <SelectContent>{FIELDS.map(f => <SelectItem key={f.id} value={f.id}>{f.name}</SelectItem>)}</SelectContent>
                    </Select>
                  </div>
                </div>

                <Button onClick={generateScript} disabled={!selectedTurma || !selectedArea} className="w-full bg-indigo-600 hover:bg-indigo-700 py-6 text-md font-bold rounded-xl shadow-lg mt-2">
                  <ClipboardCopy className="w-5 h-5 mr-2" /> Copiar Script
                </Button>
              </CardContent>
            </Card>
          </div>

          <div className="lg:col-span-8">
            <Card className="border-none shadow-sm h-full min-h-[600px] overflow-hidden flex flex-col">
              <CardHeader className="border-b border-slate-100 bg-white">
                <div className="flex items-center justify-between">
                  <CardTitle className="text-lg">Conferência de Notas</CardTitle>
                  {selectedArea && selectedCategory !== 'all' && (
                    <div className="flex items-center gap-2 text-[10px] font-bold text-amber-600 bg-amber-50 px-3 py-1 rounded-full border border-amber-100">
                      <AlertCircle className="w-3 h-3" />
                      Lançando apenas: {CATEGORIES.find(c => c.id === selectedCategory)?.name} ({selectedField.toUpperCase()})
                    </div>
                  )}
                </div>
              </CardHeader>
              <CardContent className="p-0 flex-1 bg-white">
                {filteredData.length > 0 ? (
                  <div className="overflow-auto max-h-[700px]">
                    <Table>
                      <TableHeader className="bg-slate-50/50 sticky top-0 z-10">
                        <TableRow>
                          <TableHead className="font-bold text-slate-700 pl-6">Aluno</TableHead>
                          <TableHead className="text-center font-bold text-indigo-600">
                            {selectedCategory === 'all' ? `Notas ${selectedArea}` : `Nota a Lançar`}
                          </TableHead>
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {filteredData.map((row, i) => {
                          const notes = row[selectedArea] || [];
                          const recs = row[`R${selectedArea}`] || [];
                          let displayValue = "";
                          
                          if (selectedCategory === 'all') {
                            displayValue = notes.join(' | ');
                          } else {
                            const idx = parseInt(selectedCategory);
                            if (selectedField === 'av') displayValue = notes[idx] || "-";
                            else if (selectedField === 'rec') displayValue = recs[idx] || "-";
                            else displayValue = `Av: ${notes[idx] || "-"} | Rec: ${recs[idx] || "-"}`;
                          }

                          return (
                            <TableRow key={i} className="hover:bg-slate-50/50 transition-colors">
                              <TableCell className="font-medium text-slate-900 text-xs pl-6">{row._name || '-'}</TableCell>
                              <TableCell className="text-center font-bold text-indigo-600 text-xs">
                                {displayValue}
                              </TableCell>
                            </TableRow>
                          );
                        })}
                      </TableBody>
                    </Table>
                  </div>
                ) : (
                  <div className="flex flex-col items-center justify-center h-full min-h-[500px] text-slate-400">
                    <Layers className="w-16 h-16 mb-4 opacity-10" />
                    <p className="text-sm font-medium">Aguardando dados da planilha</p>
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