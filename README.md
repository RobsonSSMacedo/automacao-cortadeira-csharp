# 🏭 Automação Industrial: Contador de Produção com Trava de Segurança (Time-Bomb)

## 📝 Sobre o Projeto
Este repositório apresenta uma solução de *Indústria 4.0* desenvolvida em *C# (.NET Framework 4.7.2)* exposta via arquitetura *COM Interop, integrada diretamente a planilhas corporativas Microsoft Excel através de gatilhos em **VBA*.

O sistema realiza a comunicação automatizada com o chão de fábrica para efetuar a contagem em tempo real de cortes em uma cortadeira de blocos cerâmicos, agregando os dados por janelas horárias e consolidando-os em tabelas estruturadas (ListObjects) de forma imutável para auditoria da gerência.

---

## 🛠️ Diferenciais Técnicos & Governança de Código

*   *Mecanismo de Lock (Time-Bomb):* Implementação de rotina de segurança com validação cronológica que bloqueia a execução da biblioteca a partir de *2027*, exibindo uma interface visual (MessageBox) personalizada com dados de contato do desenvolvedor para manutenção preventiva.
*   *Resiliência Baseada em ListObjects:* Acesso e escrita indexados via propriedade .Cells em tabelas estruturadas nominais, garantindo que o código não quebre mesmo se o usuário final mover a tabela de posição na planilha.
*   *Gerenciamento de Memória COM:* Mitigação total de vazamentos de memória (Memory Leaks) no processo EXCEL.EXE através da liberação explícita de ponteiros não-gerenciados com Marshal.ReleaseComObject() combinada com a coleta forçada de lixo (GC.Collect()).

---

## 📈 Impacto de Negócio
*   *Confiabilidade de Dados:* Eliminação de 100% de erros manuais de digitação de produção pelos operadores.
*   *Custo Zero de Infraestrutura:* Aproveitamento do ecossistema de software que a indústria já possuía (Excel), sem necessidade de aquisição de licenças caras de sistemas MES proprietários.
