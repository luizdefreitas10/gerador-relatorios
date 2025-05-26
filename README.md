📄 Gerador de Relatórios de Fiscalização
Projeto desenvolvido para automatizar a geração de relatórios de fiscalizações realizadas pelo setor de transportes. A partir de uma planilha contendo informações das fiscalizações e fotos associadas, o script gera automaticamente arquivos Word (.docx) e PDF para cada fiscalização.

✅ Funcionalidades
📥 Leitura automática de planilha Excel com os dados das fiscalizações.

🖼️ Inclusão de múltiplas fotos associadas a cada fiscalização.

📄 Geração automática de relatórios em formato Word (.docx).

🖨️ Conversão automática dos relatórios para PDF.

🔄 Marcação na planilha das fiscalizações que já tiveram relatório gerado, evitando duplicidade.

🛡️ Executa somente fiscalizações ainda não processadas.

🗂️ Estrutura de Pastas
bash
Copiar código
gerador-de-relatorios/
├── src/
│   ├── assets/           # Pasta com as fotos das fiscalizações
│   ├── reports/          # Relatórios gerados (Word e PDF)
│   ├── planilha_fiscalizacao.xlsx   # Planilha com dados das fiscalizações
│   └── main.py           # Script principal
├── venv/                 # Ambiente virtual Python
├── requirements.txt      # Arquivo de dependências
└── README.md             # Este arquivo
🛠️ Tecnologias e Bibliotecas Utilizadas
Python 3.10+

pandas - Para manipulação da planilha Excel.

python-docx - Para criação de arquivos Word.

docx2pdf - Para conversão de Word para PDF.

openpyxl - Para leitura de arquivos Excel.

▶️ Como Executar o Projeto
1. Clone o repositório
bash
Copiar código
git clone https://github.com/seu-usuario/gerador-de-relatorios.git
cd gerador-de-relatorios
2. Crie e ative o ambiente virtual
Windows:

bash
Copiar código
python -m venv venv
venv\Scripts\activate
Linux/MacOS:

bash
Copiar código
python3 -m venv venv
source venv/bin/activate
3. Instale as dependências
bash
Copiar código
pip install -r requirements.txt
4. Organize os arquivos
Coloque as fotos na pasta src/assets/.

Coloque a planilha planilha_fiscalizacao.xlsx na pasta src/.

5. Execute o script
bash
Copiar código
cd src
python main.py
✅ Formato esperado da Planilha (planilha_fiscalizacao.xlsx)
Data	Local	Pessoal Responsável	Não conformidade	Fotos	Relatório Gerado
2025-05-20	Terminal Central	João Silva	Piso danificado	foto1.jpg;foto2.jpg	(Deixe vazio)

Fotos: nomes separados por ponto e vírgula ;.

Relatório Gerado: o script automaticamente preencherá com "Sim" após gerar o relatório.

💡 Funcionalidade de Controle de Processamento
Antes de gerar o relatório, o script verifica a coluna Relatório Gerado.

Se vazio, gera o relatório e marca como "Sim".

Se "Sim", ignora e segue para a próxima fiscalização.

📑 Saída
Arquivos .docx e .pdf gerados automaticamente na pasta src/reports/.

❗ Requisitos
Python instalado

Permissão de execução de scripts (no Windows: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser)

📌 Observações importantes
As fotos devem ter nomes exatamente iguais aos referenciados na planilha.

O script não sobrescreve relatórios já gerados.

Após execução, a planilha será atualizada com a marcação "Sim" na coluna Relatório Gerado.

🤝 Contribuições
Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests.

📝 Licença
Este projeto está licenciado sob a MIT License.

📞 Contato
Desenvolvido por Luiz de Freitas
📧 Email: luiz.defreitas@arpe.pe.gov.br
