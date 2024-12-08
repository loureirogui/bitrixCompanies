Automatização de Cadastro e Atualização de Empresas no Bitrix24
Descrição
Este projeto fornece um script em Python para automatizar a consulta, cadastro e atualização de empresas no CRM Bitrix24. Utilizando CPF/CNPJ como identificador principal, o script interage com a API do Bitrix24 e uma planilha Excel contendo os dados das empresas.

Funcionalidades Principais
Consulta de Empresas: Verifica se uma empresa já está cadastrada no Bitrix24 com base no CNPJ.
Atualização de Dados: Atualiza as informações de empresas existentes com base em dados da planilha.
Cadastro de Novas Empresas: Registra automaticamente empresas que não estão no Bitrix24.
Flexibilidade: Permite personalização dos campos do Bitrix24 e do formato da planilha.

Pré-requisitos
Antes de usar o script, certifique-se de ter o seguinte configurado em seu ambiente:
Python 3.x (preferencialmente a versão mais recente).
Conta ativa no Bitrix24 com acesso à API.

Bibliotecas Necessárias
Instale as bibliotecas Python utilizadas no projeto:

pip install requests openpyxl art

Como Usar
1. Configurar a Planilha de Entrada
Salve sua planilha Excel no mesmo diretório do script.
O script espera que os dados estejam organizados da seguinte forma:
Coluna 1: CNPJ
Coluna 2: Receita Bruta
Coluna 3: Email
Coluna 4: Telefone
Coluna 5: Nome da Empresa
Coluna 6: Responsável
Coluna 7: Status do Cliente (Ativo ou Inativo).

3. Configurar os Campos do Bitrix24
No início do código, ajuste os mapeamentos de campos para que correspondam às suas configurações no Bitrix24:
FIELD_COMPANY_TYPE = 'COMPANY_TYPE'
FIELD_TITLE = 'TITLE'
FIELD_CNPJ = 'UF_CRM_1701275490640'
FIELD_REVENUE = 'UF_CRM_1727441546022'
FIELD_EMAIL = 'EMAIL'
FIELD_PHONE = 'PHONE'
FIELD_RESPONSIBLE = 'UF_CRM_1727358267819'

3. Executar o Script
Inicie o script:
python bitrixCompanies.py

Quando solicitado, insira:
O nome da sua planilha (sem a extensão .xlsx).

O token da API do Bitrix24.


5. Verificar Logs
O script exibirá mensagens no console informando se as empresas foram cadastradas ou atualizadas com sucesso.
Qualquer erro será detalhado para facilitar a depuração.

Estrutura do Código
Funções Principais
processSpreadsheet(): Lê os dados da planilha e processa cada linha.
queryCompany(): Consulta uma empresa no Bitrix24 pelo CNPJ.
updateCompany(): Atualiza os dados de uma empresa já existente.
registerCompany(): Registra uma nova empresa no Bitrix24.

Manipulação de Dados
Emails e telefones são separados por ponto-e-vírgula (;) na planilha e tratados pelo script para a API do Bitrix24.
Tecnologias Utilizadas
Python: Linguagem principal do projeto.
Bitrix24 API: Para consulta, registro e atualização de empresas no CRM.
OpenPyXL: Para leitura e manipulação de planilhas Excel.
Art: Para um toque de estilo na interface CLI.

Contato
Desenvolvido por Guilherme Loureiro.
Em caso de dúvidas, entre em contato pelo e-mail: guilherme.c.loureiro@gmail.com
