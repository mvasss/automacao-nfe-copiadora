# automacao-nfe-copiadora

Script em Python para emitir notas fiscais no portal da prefeitura de Osasco automaticamente via Selenium.

## 💡 O que o projeto faz

Este script automatiza a emissão de Notas Fiscais de Serviço Eletrônicas (NFS-e) no site da prefeitura de Osasco, poupando tempo de empresas que emitem notas com frequência.  
Ele realiza login no sistema, preenche os dados da nota e envia automaticamente.

## 🧰 Tecnologias utilizadas

- Python 3.x  
- Selenium  
- Pandas  
- ChromeDriver

## 📝 Pré-requisitos

- Python instalado  
- Google Chrome instalado  
- ChromeDriver compatível com sua versão do navegador  
- Instalar dependências com:
bash
pip install -r requirements.txt

## 📂 Como usar
1-Preencha o arquivo de entrada (Excel) com os dados das notas.

2-Atualize o script com seu login/senha no portal da prefeitura.

3-Execute com:
python emissor_nfe_selenium.py

4-O navegador será aberto automaticamente e as notas serão emitidas.

## 📄 Exemplo de estrutura do Excel
- PAGADOR   CPF		       Valor
- João      00000000000	 150,00

## 📌 Observações
- O código está em fase inicial/testes.
- O login e senha são armazenados diretamente no script (não recomendado para produção).

