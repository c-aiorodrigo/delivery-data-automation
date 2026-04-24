# Automação de Análise de Delivery (ETL & Relatórios)

## A História por Trás do Projeto

Este projeto nasceu da observação de uma dor real de negócio. Durante minha atuação no setor de recepção de uma empresa, notei que a consolidação e análise dos dados de delivery eram feitas de forma 100% manual no Excel. Era um processo operacional exaustivo que consumia cerca de **2 horas** de trabalho contínuo.

Mesmo possuindo uma base iniciante em programação na época, decidi assumir a responsabilidade de resolver esse gargalo. Desenvolvi este script em Python que automatizou todo o fluxo, transformando um processo demorado e sujeito a falhas humanas em uma tarefa de **apenas 1 clique**.

Decidi publicar este projeto no GitHub com dois propósitos:

1. **Portfólio:** Demonstrar minha capacidade de identificar problemas reais e criar soluções tecnológicas para eles.
2. **Evolução Contínua:** Este é o código na sua versão original. Meu objetivo agora é utilizar este repositório como um ambiente de estudos, melhorando a arquitetura do código, aplicando novos conceitos e documentando minha evolução profissional através dos próximos commits.

---

## O que o projeto faz?

O script funciona como um pipeline de ETL (Extração, Transformação e Carga) simplificado:

- **Extração e Limpeza:** Lê dados brutos de planilhas de forma segura, padroniza nomes e corrige tipagens de dados automaticamente.
- **Inteligência de Negócio:** Calcula os principais KPIs de vendas (Faturamento, Volume de Pedidos, Crescimento YoY e MTD).
- **Geração de Relatórios:** Cria automaticamente um novo arquivo Excel segmentado por abas (Top Performers, Lojas Zeradas, Resumo Geral) e salva em pastas dinâmicas.
- **Comunicação:** Integra-se com o WhatsApp para disparar um resumo executivo com os resultados na hora.

## Tecnologias Utilizadas

- **Python** (Linguagem principal)
- **Pandas** (Manipulação e análise de dados)
- **Pywhatkit** (Automação de envio de mensagens no WhatsApp)
- **Python-dotenv** (Segurança e gerenciamento de variáveis de ambiente)

## Segurança

O projeto utiliza variáveis de ambiente (`.env`) para garantir que dados sensíveis, como números de telefone pessoal, não fiquem expostos no código público.

## Como rodar o projeto localmente

1. Clone este repositório:

   ```bash
   git clone [https://github.com/SEU_USUARIO/delivery-data-automation.git](https://github.com/SEU_USUARIO/delivery-data-automation.git)

   ```

2. Instale as dependências necessárias:

   ```bash
   pip install pandas pywhatkit python-dotenv openpyxl

   ```

3. Crie um arquivo .env na raiz do projeto e adicione seu número:

   ```Plaintext
   WHATSAPP_NUMERO=+55SEUNUMERO

   ```

4. Execute o script principal:
   ```bash
   python seu_script.py
   ```

## Próximos Passos (Roadmap de Melhorias)

Como mencionei, este projeto está em constante evolução. Nos próximos commits, pretendo focar em:

[ ] Refatorar o código monolítico para Clean Architecture (separar ETL em módulos).

[ ] Implementar validação estrita dos dados de entrada da planilha.

[ ] Criar testes automatizados para as lógicas de cálculo de datas.
