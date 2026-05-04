import React, { useState, useMemo, useEffect } from 'react';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Settings2, Layers, Navigation, Plus, Trash2, Search } from "lucide-react";
import * as XLSX from 'xlsx';
import { showSuccess, showError } from "@/utils/toast";
import { MadeWithDyad } from "@/components/made-with-dyad";

// Helper para converter letra (A, B, C) em índice (0, 1, 2)
const columnLetterToIndex = (letter: string) => {
  try {
    return XLSX.utils.decode_col(letter.toUpperCase().trim());
  } catch (e) {
    return -1;
  }
};

// Helper para converter índice em letra
const indexToColumnLetter = (index: number) => {
  return XLSX.utils.encode_col(index);
};

interface SegesCategory {
  id: string;
  name: string;
  avCol: string; // Letra da coluna (ex: "D")
  recCol: string; // Letra da coluna (ex: "E")
}

const Index = () => {
  const [rawSheets, setRawSheets] = useState<{[key: string]: any[][]}>({});
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedTurma, setSelectedTurma] = useState<string>('');
  const [studentNameCol, setStudentNameCol] = useState<string>(''); // Letra da coluna de nomes
  
  const [segesCategories, setSegesCategories] = useState<SegesCategory[]>([
    { id: '1', name: 'Atividade Discursiva', avCol: '', recCol: '' },
    { id: '2', name: 'Prova Interdisciplinar', avCol: '', recCol: '' },
    { id: '3', name: 'Produção Escrita', avCol: '', recCol: '' },
  ]);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bData = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(bData, { type: 'array' });
        
        const sheets: {[key: string]: any[][]} = {};
        workbook.SheetNames.forEach(name => {
          const worksheet = workbook.Sheets[name];
          sheets[name] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" }) as any[][];
        });

        setRawSheets(sheets);
        setSheetNames(workbook.SheetNames);
        
        if (workbook.SheetNames.length > 0) {
          setSelectedTurma(workbook.SheetNames[0]);
          showSuccess("Planilha carregada!");
        }
      } catch (err) {
        showError("Erro ao processar Excel.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Auto-detectar coluna de nomes quando a turma muda
  useEffect(() => {
    if (!selectedTurma || !rawSheets[selectedTurma]) return;
    
    const rows = rawSheets[selectedTurma];
    // Procura nas primeiras 15 linhas por palavras-chave
    for (let i = 0; i < Math.min(rows.length, 15); i++) {
      const row = rows[i];
      const nameIdx = row.findIndex(cell => /nome|aluno|estudante/i.test(String(cell)));
      if (nameIdx !== -1) {
        setStudentNameCol(indexToColumnLetter(nameIdx));
        break;
      }
    }
  }, [selectedTurma, rawSheets]);

  const addCategory = () => {
    setSegesCategories([...segesCategories, { id: Date.now().toString(), name: 'Nova Avaliação', avCol: '', recCol: '' }]);
  };

  const removeCategory = (id: string) => {
    setSegesCategories(segesCategories.filter(c => c.id !== id));
  };

  const updateCategory = (id: string, field: keyof SegesCategory, value: string) => {
    setSegesCategories(segesCategories.map(c => c.id === id ? { ...c, [field]: value.toUpperCase() } : c));
  };

  // Processa os dados para exibição e script
  const processedData = useMemo(() => {
    if (!selectedTurma || !rawSheets[selectedTurma] || !studentNameCol) return [];
    
    const rows = rawSheets[selectedTurma];
    const nameIdx = columnLetterToIndex(studentNameCol);
    
    // Encontra onde começam os dados (primeira linha após o cabeçalho que tem um nome válido)
    const headerIdx = rows.findIndex(r => /nome|aluno|estudante/i.test(String(r[nameIdx])));
    const startIdx = headerIdx !== -1 ? headerIdx + 1 : 0;

    return rows.slice(startIdx)
      .filter(row => String(row[nameIdx] || "").trim().length > 3) // Filtra linhas sem nome real
      .map(row => {
        const studentData: any = { name: String(row[nameIdx]).trim() };
        segesCategories.forEach(cat => {
          const avIdx = columnLetterToIndex(cat.avCol);
          const recIdx = columnLetterToIndex(cat.recCol);
          studentData[cat.id + '_av'] = avIdx !== -1 ? row[avIdx] : "";
          studentData[cat.id + '_rec'] = recIdx !== -1 ? row[recIdx] : "";
        });
        return studentData;
      });
  }, [rawSheets, selectedTurma, studentNameCol, segesCategories]);

  const generateScript = () => {
    if (processedData.length === 0) return showError("Nenhum dado para processar.");

    const scriptData = processedData.map(student => ({
      name: student.name.toUpperCase(),
      mapping: segesCategories.map(cat => ({
        search: cat.name.toUpperCase(),
        av: student[cat.id + '_av'] || "",
        rec: student[cat.id + '_rec'] || ""
      }))
    }));

    const script = `
(function() {
  const studentsData = ${JSON.stringify(scriptData)};
  let count = 0;
  studentsData.forEach(student => {
    const row = Array.from(document.querySelectorAll('tr')).find(tr => tr.innerText.toUpperCase().includes(student.name));
    if (row) {
      const inputs = Array.from(row.querySelectorAll('input[type="text"]'));
      student.mapping.forEach((map) => {
        const categoriesWithInputs = Array.from(document.querySelectorAll('thead th')).filter(th => 
          ['DISCURSIVA', 'INTERDISCIPLINAR', 'PRODUCAO ESCRITA', 'PROVA', 'PROJETO', 'ATIVIDADE'].some(term => th.innerText.toUpperCase().includes(term))
        );
        const targetTh = categoriesWithInputs.find(th => th.innerText.toUpperCase().includes(map.search));
        if (targetTh) {
          const catPos = categoriesWithInputs.indexOf(targetTh);
          const idxAv = catPos * 2;
          const idxRec = catPos * 2 + 1;
          if (map.av !== "" && inputs[idxAv]) {
            inputs[idxAv].value = map.av.toString().replace('.', ',');
            ['input', 'change', 'blur'].forEach(t => inputs[idxAv].dispatchEvent(new Event(t, { bubbles: true })));
          }
          if (map.rec !== "" && inputs[idxRec]) {
            inputs[idxRec].value = map.rec.toString().replace('.', ',');
            ['input', 'change', 'blur'].forEach(t => inputs[idxRec].dispatchEvent(new Event(t, { bubbles: true })));
          }
        }
      });
      count++;
      row.style.backgroundColor = '#f0fdf4';
    }
  });
  alert('Sucesso! ' + count + ' alunos atualizados.');
})();`;

    navigator.clipboard.writeText(script);
    showSuccess("Script copiado! Use no SEGES.");
  };

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8">
      <div className="max-w-7xl mx-auto space-y-6">
        
        {/* Cabeçalho Estilizado */}
        <Card className="border-none shadow-lg overflow-hidden mb-8">
          <CardHeader className="bg-[#4338ca] text-white p-8 md:p-10">
            <div className="flex flex-col md:flex-row items-center gap-6 md:gap-8">
              <div className="bg-white/15 p-5 rounded-[2rem] border border-white/10 shadow-inner">
                <Navigation className="w-12 h-12 text-white fill-white/10 rotate-45" />
              </div>
              <div className="text-center md:text-left">
                <CardTitle className="text-4xl md:text-5xl font-black tracking-tight mb-2">
                  Lançador de Notas SEGES
                </CardTitle>
                <CardDescription className="text-indigo-100 text-lg md:text-xl font-medium opacity-90">
                  Mapeamento inteligente por coordenadas de planilha.
                </CardDescription>
              </div>
            </div>
          </CardHeader>
        </Card>

        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          
          {/* Painel de Configuração */}
          <div className="lg:col-span-5 space-y-6">
            <Card className="border-none shadow-sm overflow-hidden">
              <CardHeader className="bg-slate-50 border-b border-slate-100">
                <div className="flex items-center gap-3">
                  <Settings2 className="w-5 h-5 text-indigo-600" />
                  <CardTitle className="text-sm font-bold text-slate-700 uppercase tracking-wider">Configuração</CardTitle>
                </div>
              </CardHeader>
              
              <CardContent className="p-5 space-y-6">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <Label className="text-xs font-bold">1. Planilha Excel</Label>
                    <Input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} className="text-[10px] h-10" />
                  </div>
                  <div className="space-y-2">
                    <Label className="text-xs font-bold">2. Turma (Aba)</Label>
                    <Select value={selectedTurma} onValueChange={setSelectedTurma}>
                      <SelectTrigger className="h-10 text-xs"><SelectValue placeholder="Selecione..." /></SelectTrigger>
                      <SelectContent>{sheetNames.map(n => <SelectItem key={n} value={n}>{n}</SelectItem>)}</SelectContent>
                    </Select>
                  </div>
                </div>

                <div className="space-y-2">
                  <Label className="text-xs font-bold flex items-center gap-2">
                    3. Coluna de Nomes (Letra) 
                    {studentNameCol && <span className="text-green-600 flex items-center gap-1"><Search className="w-3 h-3"/> Detectado</span>}
                  </Label>
                  <Input 
                    placeholder="Ex: B" 
                    value={studentNameCol} 
                    onChange={(e) => setStudentNameCol(e.target.value.toUpperCase())}
                    className="h-10 text-sm font-bold uppercase"
                  />
                </div>

                <div className="space-y-4 pt-4 border-t border-slate-100">
                  <div className="flex items-center justify-between">
                    <Label className="text-xs font-bold text-indigo-600 uppercase tracking-wider">4. Mapear Notas (Letras)</Label>
                    <Button onClick={addCategory} variant="outline" size="sm" className="h-7 text-[10px] gap-1">
                      <Plus className="w-3 h-3" /> Add Categoria
                    </Button>
                  </div>

                  <div className="space-y-4">
                    {segesCategories.map((cat) => (
                      <div key={cat.id} className="p-4 bg-slate-50 rounded-2xl border border-slate-100 space-y-3 relative group">
                        <Button 
                          variant="ghost" 
                          size="icon" 
                          onClick={() => removeCategory(cat.id)}
                          className="absolute top-2 right-2 h-6 w-6 text-slate-300 hover:text-red-500 opacity-0 group-hover:opacity-100 transition-opacity"
                        >
                          <Trash2 className="w-3 h-3" />
                        </Button>
                        
                        <div className="space-y-1">
                          <Label className="text-[10px] text-slate-400 font-bold uppercase">Nome no SEGES</Label>
                          <Input 
                            value={cat.name} 
                            onChange={(e) => updateCategory(cat.id, 'name', e.target.value)}
                            className="h-8 text-xs font-bold bg-white"
                          />
                        </div>

                        <div className="grid grid-cols-2 gap-3">
                          <div className="space-y-1">
                            <Label className="text-[10px] text-slate-400 font-bold uppercase">Coluna Av. (Letra)</Label>
                            <Input 
                              placeholder="Ex: D"
                              value={cat.avCol}
                              onChange={(e) => updateCategory(cat.id, 'avCol', e.target.value)}
                              className="h-8 text-xs font-bold bg-white uppercase"
                            />
                          </div>
                          <div className="space-y-1">
                            <Label className="text-[10px] text-slate-400 font-bold uppercase">Coluna Rec. (Letra)</Label>
                            <Input 
                              placeholder="Ex: E"
                              value={cat.recCol}
                              onChange={(e) => updateCategory(cat.id, 'recCol', e.target.value)}
                              className="h-8 text-xs font-bold bg-white uppercase"
                            />
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>

                <Button onClick={generateScript} disabled={!selectedTurma} className="w-full bg-[#4338ca] hover:bg-[#3730a3] py-7 text-md font-bold rounded-2xl shadow-lg shadow-indigo-100 mt-4">
                  <ClipboardCopy className="w-5 h-5 mr-2" /> Copiar Script de Lançamento
                </Button>
              </CardContent>
            </Card>
          </div>

          {/* Visualização */}
          <div className="lg:col-span-7">
            <Card className="border-none shadow-sm h-full min-h-[600px] overflow-hidden flex flex-col">
              <CardHeader className="border-b border-slate-100 bg-white">
                <CardTitle className="text-lg">Pré-visualização</CardTitle>
                <CardDescription>Dados extraídos das coordenadas informadas</CardDescription>
              </CardHeader>
              <CardContent className="p-0 flex-1 bg-white">
                {processedData.length > 0 ? (
                  <div className="overflow-auto max-h-[750px]">
                    <Table>
                      <TableHeader className="bg-slate-50/50 sticky top-0 z-10">
                        <TableRow>
                          <TableHead className="font-bold text-slate-700 pl-6">Aluno</TableHead>
                          {segesCategories.map(cat => (
                            <TableHead key={cat.id} className="text-center font-bold text-indigo-600 text-[10px] uppercase leading-tight">
                              {cat.name}<br/><span className="text-slate-400">Col {cat.avCol || '?'}/{cat.recCol || '?'}</span>
                            </TableHead>
                          ))}
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {processedData.map((student, i) => (
                          <TableRow key={i} className="hover:bg-slate-50/50 transition-colors">
                            <TableCell className="font-medium text-slate-900 text-[10px] pl-6">{student.name}</TableCell>
                            {segesCategories.map(cat => (
                              <TableCell key={cat.id} className="text-center text-[10px] font-bold">
                                <span className="text-indigo-600">{student[cat.id + '_av'] || '-'}</span>
                                <span className="mx-1 text-slate-300">|</span>
                                <span className="text-orange-600">{student[cat.id + '_rec'] || '-'}</span>
                              </TableCell>
                            ))}
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </div>
                ) : (
                  <div className="flex flex-col items-center justify-center h-full min-h-[500px] text-slate-400 p-12 text-center">
                    <Layers className="w-16 h-16 mb-4 opacity-10" />
                    <p className="text-sm font-medium">Aguardando coordenadas</p>
                    <p className="text-xs opacity-60 mt-2 max-w-xs">Informe as letras das colunas (A, B, C...) para ver os dados aqui.</p>
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