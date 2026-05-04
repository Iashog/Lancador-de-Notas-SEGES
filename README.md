# Lançador de Notas SEGES

Este aplicativo foi desenvolvido para automatizar o lançamento de notas no sistema SEGES (Sedu-ES) a partir de planilhas Excel personalizadas.

## Como usar

1. **Carregue sua Planilha**: Selecione o arquivo `.xlsx` ou `.xls`.
2. **Selecione a Turma**: Escolha a aba correspondente à turma que deseja lançar.
3. **Configure as Colunas**:
   - Informe a letra da coluna que contém os nomes dos alunos (ex: B).
   - Para cada avaliação, informe a letra da coluna de Nota (Av) e Recuperação (Rec).
4. **Gere o Script**: Clique em "Copiar Script de Lançamento".
5. **Execute no SEGES**:
   - Abra a página de lançamento de notas da turma no SEGES.
   - Pressione `F12` para abrir o Console do Desenvolvedor.
   - Se for sua primeira vez usando o console, o navegador pode bloquear a colagem. Digite `allow pasting` e aperte Enter para liberar.
   - Cole o script copiado e aperte `Enter`.

## Segurança e Privacidade

- **Processamento Local**: Seus dados não são enviados para nenhum servidor. Todo o processamento do Excel e a geração do script ocorrem localmente no seu navegador.
- **Código Aberto**: O script gerado é transparente e pode ser lido antes da execução.

---
*Desenvolvido para facilitar a vida do professor capixaba.*