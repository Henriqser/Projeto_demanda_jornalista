# Projeto RPA de Automação de Consultas de Dados

## Descrição

Este projeto foi desenvolvido como uma demonstração de conhecimento técnico em automação de processos utilizando Python. Ele simula um sistema de consulta de dados que pode ser utilizado por equipes de jornalistas para atender demandas simples de forma autônoma, especialmente em situações em que a área de dados está ocupada ou indisponível.

Embora o código não esteja funcional devido à censura de informações sensíveis, ele serve como uma vitrine para demonstrar habilidades em desenvolvimento de software, automação, integração com bancos de dados e interfaces gráficas.

---

## Funcionalidades

- **Interface Gráfica (GUI):** Desenvolvida com `Tkinter`, a interface permite que o usuário selecione o tipo de relatório desejado, insira parâmetros como datas e IDs, e execute consultas de forma intuitiva.
- **Consultas ao Banco de Dados:** Simula a execução de queries em um banco de dados PostgreSQL para gerar relatórios personalizados.
- **Exportação para Excel:** Os resultados das consultas podem ser salvos em arquivos Excel, facilitando o compartilhamento e análise dos dados.
- **Validação de Entradas:** O sistema valida os dados inseridos pelo usuário, garantindo que estejam no formato correto antes de executar as consultas.

---

## Estrutura do Projeto

- **`script.py`:** Contém o código principal do sistema, incluindo a interface gráfica e a lógica de execução das consultas.
- **`projeto jornalistas.bat`:** Um script em batch para facilitar a execução do programa Python diretamente no Windows.

---

## Tecnologias Utilizadas

- **Python:** Linguagem principal do projeto.
    - **Bibliotecas:** 
        - `Tkinter` para a interface gráfica.
        - `psycopg2` para conexão com o banco de dados PostgreSQL.
        - `pandas` para manipulação de dados.

        ---

        ## Layout do Projeto

        Veja abaixo o layout do sistema de consulta de dados:

        ![Interface do sistema](Layout%20sistema.png)
      
        - `dotenv` para gerenciamento de variáveis de ambiente.
- **PostgreSQL:** Banco de dados utilizado para simular as consultas.
- **Batch Script:** Para automatizar a execução do programa no Windows.

---

## Observações

- Este projeto é apenas uma demonstração e não está funcional devido à censura de informações sensíveis.
- O foco principal é demonstrar habilidades técnicas, incluindo:
    - Desenvolvimento de interfaces gráficas.
    - Integração com bancos de dados.
    - Manipulação de dados e exportação para Excel.
    - Automação de processos.

---

## Autor

Desenvolvido por Henrique Silva.  
Este projeto foi criado com o objetivo de demonstrar habilidades técnicas e contribuir para a automação de processos em equipes de comunicação.

---  
