# 🏭 Automação Industrial: Contador de Produção com Compatibilidade Universal & Trava (Time-Bomb)

[![C#](https://shields.io)](https://microsoft.com)
[![VBA](https://shields.io)](https://microsoft.com)

## 📝 Descrição do Projeto
Este projeto apresenta uma solução estável de *Indústria 4.0* desenvolvida para substituir processos manuais de inventário e auditoria no chão de fábrica. Trata-se de uma *Biblioteca de Vínculo Dinâmico (DLL)* escrita em C# (.NET Framework) exposta via arquitetura COM Interop, integrada a planilhas corporativas Microsoft Excel através de gatilhos acionados em VBA.

O sistema intercepta pulsos de produção de uma cortadeira de blocos cerâmicos, realiza a agregação cronológica dos dados por janelas horárias e registra as métricas em tabelas estruturadas nominais em tempo real.

---

## 🛠️ Engenharia de Software e Diferenciais Técnicos

*   *Abstração Dinâmica (Universal Compatibility):* Emprego da palavra-chave dynamic no loop de iteração das células do Excel. Esta abordagem elimina dependências rígidas de compilação (early binding) vinculadas a versões específicas de assemblies do Office, tornando a DLL universal e compatível com qualquer versão do Excel instalada nas máquinas de produção.
*   *Mecanismo de Lock (Time-Bomb):* Inclusão de rotina de governança e segurança de código com validação cronológica que bloqueia a execução da biblioteca a partir de *2027*, exibindo uma interface visual (MessageBox) personalizada com dados de contato do desenvolvedor para manutenção preventiva.
*   *Gerenciamento de Memória COM:* Mitigação total de vazamentos de memória (Memory Leaks) no processo EXCEL.EXE através da liberação explícita de ponteiros não-gerenciados com Marshal.ReleaseComObject() combinada com a coleta forçada de lixo (GC.Collect()).
*   *Resiliência Baseada em ListObjects:* Acesso e escrita indexados via propriedade .Cells em tabelas estruturadas nominais, garantindo integridade dos dados mesmo se o operador mover a tabela de posição na planilha.

---

## 💻 Estrutura dos Códigos no Repositório

*   *Jinx.csproj*: Arquivo de configuração do projeto estruturado com pacotes NuGet de interoperabilidade e suporte ao Microsoft.CSharp.
*   *Calculador.cs (ou Class1.cs)*: Código-fonte em C# contendo a lógica de agregação por hora, o tratamento de memória e a trava temporal.
*   *BotaoClique.vba*: Gatilho em código VBA embutido no botão da planilha para instanciar a DLL com tratamento de exceções (On Error GoTo).

---

## 📈 Impacto de Negócio
*   *Confiabilidade Estatística:* Eliminação de 100% de erros manuais de digitação de produção pelos operadores de turno.
*   *Custo Zero de Infraestrutura:* Aproveitamento do ecossistema de software que a indústria já possuía (Excel), sem necessidade de aquisição de licenças caras de sistemas ERP proprietários.
